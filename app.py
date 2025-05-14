import streamlit as st
import pandas as pd
import io
import uuid
from difflib import get_close_matches

# Настройка страницы Streamlit
st.set_page_config(page_title="Сопоставление 1C и Moz", layout="wide")

# Создаем табы: Инструкция и Приложение
tabs = st.tabs(["📄 Инструкция", "🚀 Приложение"])

with tabs[0]:
    st.header("Как пользоваться программой")
    st.markdown("""
1. Загрузите **два** файла: _1C.xlsx_ и _Moz.xlsx_ через форму во вкладке «Приложение».
2. Нажмите кнопку **«Обработать»**.
3. Дождитесь выполнения — после этого появится кнопка **«Скачать результат»**.
4. Файлы не сохраняются на диск: всё происходит в памяти!

---

**Обязательные поля в входных файлах**

- **Файл 1C**:
    - `номер рег`
    - `код 1с`
    - `наименование полное`
    - `форма випуску`
    - `дозування`
    - `кількість одиниць лікарського засобу у споживчій упаковці`

- **Файл Moz**:
    - `міжнародна непатентована або загальноприйнята назва лікарського засобу`
    - `торговельна назва лікарського засобу`
    - `форма випуску`
    - `дозування`
    - `кількість одиниць лікарського засобу у споживчій упаковці`
    - `найменування виробника, країна`
    - `код атх`
    - `номер реєстраційного посвідчення на лікарський засіб`
    - `дата закінчення строку дії реєстраційного посвідчення на лікарський засіб`
    - `задекларована зміна оптово-відпускної ціни`
    - `офіційний курс та вид іноземної валюти`
    - `дата та номер наказу МОЗ про декларування змін оптово-відпускної ціни на лікарські засоби`

> Если в ваших файлах названия столбцов чуть отличаются, пожалуйста замените их как в примере выше 😊
    """)

with tabs[1]:
    st.header("Сопоставление файлов")
    uploaded_1c = st.file_uploader("Выберите файл 1C (Excel)", type=['xls', 'xlsx'], key="file1c")
    uploaded_moz = st.file_uploader("Выберите файл Moz (Excel)", type=['xls', 'xlsx'], key="filemoz")

    if uploaded_1c and uploaded_moz:
        if st.button("Обработать"):
            try:
                # Чтение файлов
                df_1c = pd.read_excel(uploaded_1c, engine='openpyxl')
                df_moz = pd.read_excel(uploaded_moz, engine='openpyxl')

                # Нормализация названий
                def normalize_column_names(df, required_columns, column_mapping):
                    df.columns = df.columns.str.strip().str.lower()
                    mapping_lower = {k.lower(): v for k, v in column_mapping.items()}
                    df.rename(columns=mapping_lower, inplace=True)
                    for col in required_columns:
                        if col not in df.columns:
                            matches = get_close_matches(col, df.columns, n=1, cutoff=0.8)
                            if matches:
                                df.rename(columns={matches[0]: col}, inplace=True)
                    return df

                # Основная логика обработки
                def process_data(df_1c, df_moz):
                    required_1c = [
                        'номер рег', 'код 1с', 'наименование полное',
                        'форма випуску', 'дозування', 'кількість одиниць лікарського засобу у споживчій упаковці'
                    ]
                    required_moz = [
                        'міжнародна непатентована або загальноприйнята назва лікарського засобу',
                        'торговельна назва лікарського засобу',
                        'форма випуску', 'дозування',
                        'кількість одиниць лікарського засобу у споживчій упаковці',
                        'найменування виробника, країна',
                        'код атх',
                        'номер реєстраційного посвідчення на лікарський засіб',
                        'дата закінчення строку дії реєстраційного посвідчення на лікарський засіб',
                        'задекларована зміна оптово-відпускної ціни',
                        'офіційний курс та вид іноземної валюти',
                        'дата та номер наказу моз про декларування змін оптово-відпускної ціни на лікарські засоби'
                    ]

                    df_1c = normalize_column_names(df_1c, required_1c, {})
                    df_moz = normalize_column_names(df_moz, required_moz, {})

                    df_1c = df_1c[df_1c['номер рег'].astype(str).str.contains('UA', na=False)]

                    matched_indices = set()
                    output = []

                    for code in df_moz['номер реєстраційного посвідчення на лікарський засіб'].unique():
                        df_moz_code = df_moz[df_moz['номер реєстраційного посвідчення на лікарський засіб'] == code]
                        df_1c_code = df_1c[df_1c['номер рег'] == code]
                        if len(df_moz_code) == 1 and len(df_1c_code) == 1:
                            idx = df_moz_code.index[0]
                            matched_indices.add(idx)
                            row_1c = df_1c_code.iloc[0]
                            row_moz = df_moz_code.iloc[0]
                            output.append({
                                'Код 1С': row_1c['код 1с'],
                                'Цена': row_moz['задекларована зміна оптово-відпускної ціни'],
                                'Номер рег': row_1c['номер рег'],
                                'Наименование полное': row_1c['наименование полное']
                            })
                        else:
                            for idx, row_moz in df_moz_code.iterrows():
                                matches = df_1c_code[
                                    (df_1c_code['форма випуску'] == row_moz['форма випуску']) &
                                    (df_1c_code['дозування'] == row_moz['дозування']) &
                                    (df_1c_code['кількість одиниць лікарського засобу у споживчій упаковці'] == row_moz['кількість одиниць лікарського засобу у споживчій упаковці'])
                                ]
                                if not matches.empty:
                                    matched_indices.add(idx)
                                    row_1c2 = matches.iloc[0]
                                    output.append({
                                        'Код 1С': row_1c2['код 1с'],
                                        'Цена': row_moz['задекларована зміна оптово-відпускної ціни'],
                                        'Номер рег': row_1c2['номер рег'],
                                        'Наименование полное': row_1c2['наименование полное']
                                    })
                                    break

                    matched_df = pd.DataFrame(output)
                    unmatched_df = df_moz.drop(index=matched_indices)
                    return matched_df, unmatched_df

                # Запуск обработки
                matched_df, unmatched_df = process_data(df_1c, df_moz)

                # Подготовка Excel в памяти
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    matched_df.to_excel(writer, sheet_name='Matched', index=False)
                    unmatched_df.to_excel(writer, sheet_name='Unmatched_Moz', index=False)
                buffer.seek(0)

                # Кнопка скачивания
                filename = f"output_{uuid.uuid4().hex}.xlsx"
                st.download_button(
                    "Скачать результат",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Ошибка в процессе обработки: {e}")
