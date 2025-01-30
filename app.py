from flask import Flask, request, redirect, url_for, send_file, render_template, flash
import pandas as pd
import os
import sys
import logging
from datetime import datetime
import uuid
from difflib import get_close_matches

app = Flask(__name__)
app.secret_key = 'supersecretkey'

UPLOAD_FOLDER = 'uploads'
FILE_1C_PATH = os.path.join(UPLOAD_FOLDER, '1c.xlsx')

def resource_path(relative_path):
    """Получение абсолютного пути к ресурсу (актуально при использовании PyInstaller)."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Создаем папку для загрузки, если её нет
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- Настройка логирования ---
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Лог-файл
file_handler = logging.FileHandler('app.log', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# Вывод в консоль
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

@app.before_request
def log_request_info():
    """
    Логируем информацию о каждом входящем запросе:
    метод, путь, IP-адрес клиента и его User-Agent.
    """
    logging.info(
        f"Получен запрос: метод={request.method}, "
        f"путь={request.path}, IP={request.remote_addr}, "
        f"User-Agent={request.user_agent}"
    )

@app.route('/')
def index():
    logging.info("Отображение главной страницы.")
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    logging.info("Начало обработки запроса на загрузку файлов (upload_file).")
    try:
        # ---- Обработка файла 1C ----
        if 'file1c' in request.files:
            file1c = request.files['file1c']
            if file1c.filename != '':
                logging.info(f"Получен файл 1C: {file1c.filename}")
                if os.path.exists(FILE_1C_PATH):
                    os.remove(FILE_1C_PATH)
                    logging.debug(f"Старый файл 1C {FILE_1C_PATH} удалён.")
                file1c.save(FILE_1C_PATH)
                logging.info(f"Файл 1C сохранён в {FILE_1C_PATH}")
            else:
                logging.warning("Файл 1C был загружен, но имя файла пустое.")
        else:
            logging.warning("Файл 1C не был найден в запросе.")

        # ---- Обработка файла Moz ----
        if 'filemoz' not in request.files:
            flash('Файл Moz отсутствует.')
            logging.warning("Файл Moz отсутствует в запросе.")
            return redirect(request.url)

        filemoz = request.files['filemoz']
        if filemoz.filename == '':
            flash('Файл Moz не выбран.')
            logging.warning("Файл Moz не был выбран пользователем.")
            return redirect(request.url)

        filemoz_path = os.path.join(UPLOAD_FOLDER, filemoz.filename)
        filemoz.save(filemoz_path)
        logging.info(f"Файл Moz сохранён в {filemoz_path}")

        # ---- Проверка сохранения ----
        if not os.path.exists(FILE_1C_PATH) or not os.path.exists(filemoz_path):
            flash('Ошибка при сохранении файлов.')
            logging.error("Ошибка при сохранении файлов 1C или Moz.")
            return redirect(url_for('index'))

        # ---- Генерация уникального имени ----
        unique_id = uuid.uuid4().hex
        logging.info(f"Сгенерирован уникальный ID для выходного файла: {unique_id}")

        # ---- Основная обработка ----
        output_file = process_files(FILE_1C_PATH, filemoz_path, unique_id)
        logging.info(f"Отправка файла пользователю: {output_file}")
        return send_file(output_file, as_attachment=True)

    except Exception as e:
        logging.error(f'Ошибка при обработке запроса /upload: {str(e)}', exc_info=True)
        flash(f'Ошибка при обработке файла: {str(e)}')
        return redirect(url_for('index'))

def normalize_column_names(df, required_columns, column_mapping):
    """
    Нормализация названий столбцов:
    1) Приведение к нижнему регистру,
    2) Удаление лишних пробелов,
    3) Применение принудительного переименования (column_mapping),
    4) Поиск близких совпадений, если столбец не найден напрямую.
    """
    logging.debug("Начинаем нормализацию названий столбцов DataFrame.")
    original_columns = list(df.columns)
    df.columns = df.columns.str.strip().str.lower()
    logging.debug(f"Изначальные названия столбцов: {original_columns}")
    logging.debug(f"После .lower() и .strip() столбцы: {list(df.columns)}")

    # Преобразуем ключи mapping тоже к нижнему регистру
    column_mapping_lower = {k.lower(): v for k, v in column_mapping.items()}
    # Переименовываем те колонки, которые явно указаны в column_mapping
    df.rename(columns=column_mapping_lower, inplace=True)

    # Проверяем, есть ли ещё "пропавшие" колонки из списка required_columns
    missing_columns = [col for col in required_columns if col not in df.columns]
    for col in missing_columns:
        # Ищем наиболее похожее название
        matches = get_close_matches(col, df.columns, n=1, cutoff=0.8)
        if matches:
            df.rename(columns={matches[0]: col}, inplace=True)
            logging.info(f"Переименован столбец '{matches[0]}' -> '{col}' (близкое совпадение).")
        else:
            logging.warning(f"Не удалось найти столбец для '{col}' в DataFrame.")

    return df

def process_files(file1c, filemoz, unique_id):
    """
    Основная логика:
    1. Читаем Excel-файлы,
    2. Нормализуем столбцы (Moz и 1C),
    3. Проверяем наличие требуемых столбцов,
    4. Сопоставляем,
    5. Разделяем на Matched / Unmatched,
    6. Сохраняем результат в Excel.
    """
    try:
        logging.info("Начало обработки файлов (process_files).")

        # ---- Чтение файла 1C ----
        file1c_extension = os.path.splitext(file1c)[1].lower()
        if file1c_extension in ['.xls', '.xlsx']:
            df_1c = pd.read_excel(file1c, engine='openpyxl')
            logging.debug(f"Файл 1C прочитан, размер: {df_1c.shape}")
        else:
            raise ValueError("Неподдерживаемое расширение файла 1C")

        # ---- Чтение файла Moz ----
        filemoz_extension = os.path.splitext(filemoz)[1].lower()
        if filemoz_extension in ['.xls', '.xlsx']:
            df_moz = pd.read_excel(filemoz, engine='openpyxl')
            logging.debug(f"Файл Moz прочитан, размер: {df_moz.shape}")
        else:
            raise ValueError("Неподдерживаемое расширение файла Moz")

        logging.info("Файлы успешно прочитаны (1C и Moz).")
        logging.debug(f"Столбцы Moz до нормализации: {list(df_moz.columns)}")

        # ---- Настройка обязательных столбцов Moz ----
        required_columns_moz = [
            'міжнародна непатентована або загальноприйнята назва лікарського засобу',
            'торговельна назва лікарського засобу',
            'форма випуску',
            'дозування',
            'кількість одиниць лікарського засобу у споживчій упаковці',
            'найменування виробника, країна',
            'код атх',
            'номер реєстраційного посвідчення на лікарський засіб',
            'дата закінчення строку дії реєстраційного посвідчення на лікарський засіб',
            'задекларована зміна оптово-відпускної ціни',
            'офіційний курс та вид іноземної валюти',
            'дата та номер наказу моз про декларування змін оптово-відпускної ціни на лікарські засоби'
        ]

        # ---- Маппинг столбцов Moz ----
        # Здесь надо вписать точное название колонки из Excel (уже в .lower()) и нужное краткое имя.
        column_mapping_moz = {
            # Пример для колонки с "Офіційний курс та вид ..." — мы режем длинное название до короткого:
            "офіційний курс та вид іноземної валюти, встановлений національним банком україни на дату подання декларації зміни оптово-відпускної ціни на лікарський засіб*":
                "офіційний курс та вид іноземної валюти",

            # Если нужно, добавляйте дополнительные правила:
            # "дата та номер наказу моз про декларування змін ...": "дата та номер наказу моз про декларування змін оптово-відпускної ціни на лікарські засоби"
        }

        # ---- Нормализация Moz ----
        logging.info("Нормализация столбцов файла Moz.")
        df_moz = normalize_column_names(df_moz, required_columns_moz, column_mapping_moz)

        # Проверяем, что теперь всё в порядке
        missing_columns_moz = [col for col in required_columns_moz if col not in df_moz.columns]
        if missing_columns_moz:
            logging.error(f"Отсутствуют необходимые столбцы в файле Moz: {', '.join(missing_columns_moz)}")
            raise ValueError(f"Отсутствуют необходимые столбцы в файле Moz: {', '.join(missing_columns_moz)}")

        # ---- Настройка обязательных столбцов 1C ----
        required_columns_1c = [
            'номер рег',
            'код 1с',
            'наименование полное',
            'форма випуску',
            'дозування',
            'кількість одиниць лікарського засобу у споживчій упаковці'
        ]
        column_mapping_1c = {
            # Если в вашем файле 1C есть отличия — впишите их сюда
        }

        # ---- Нормализация 1C ----
        logging.info("Нормализация столбцов файла 1C.")
        df_1c = normalize_column_names(df_1c, required_columns_1c, column_mapping_1c)

        # Проверяем, что всё в порядке
        missing_columns_1c = [col for col in required_columns_1c if col not in df_1c.columns]
        if missing_columns_1c:
            logging.error(f"Отсутствуют необходимые столбцы в файле 1C: {', '.join(missing_columns_1c)}")
            raise ValueError(f"Отсутствуют необходимые столбцы в файле 1C: {', '.join(missing_columns_1c)}")

        # ---- Фильтрация 1C ----
        df_1c_filtered = df_1c[df_1c['номер рег'].str.contains('UA', na=False)]
        logging.debug(f"Фильтрация 1C по 'UA': {df_1c_filtered.shape} записей.")

        output_data = []
        matched_moz_indices = set()

        # ---- Проверяем "номер реєстраційного посвідчення..." ----
        if 'номер реєстраційного посвідчення на лікарський засіб' not in df_moz.columns:
            raise ValueError("Нет столбца 'номер реєстраційного посвідчення на лікарський засіб' в Moz.")

        # ---- Сопоставление ----
        unique_codes = df_moz['номер реєстраційного посвідчення на лікарський засіб'].unique()
        logging.debug(f"Уникальных рег. номеров в Moz: {len(unique_codes)}")

        for code in unique_codes:
            df_moz_code = df_moz[df_moz['номер реєстраційного посвідчення на лікарський засіб'] == code]
            df_1c_code = df_1c_filtered[df_1c_filtered['номер рег'] == code]

            # 1) Если по коду всего одна запись у 1C и у Moz
            if len(df_moz_code) == 1 and len(df_1c_code) == 1:
                matched_moz_indices.add(df_moz_code.index[0])
                код_1с = df_1c_code['код 1с'].values[0]
                цена = df_moz_code['задекларована зміна оптово-відпускної ціни'].values[0]
                номер_рег = df_1c_code['номер рег'].values[0]
                наименование_полное = df_1c_code['наименование полное'].values[0]

                output_data.append({
                    'Код 1С': код_1с,
                    'Цена': цена,
                    'Номер рег': номер_рег,
                    'Наименование полное': наименование_полное
                })

            # 2) Иначе пытаемся сопоставить по форме, дозировке, количеству
            else:
                for index_moz, row_moz in df_moz_code.iterrows():
                    df_1c_matches = df_1c_code[
                        (df_1c_code['форма випуску'] == row_moz['форма випуску']) &
                        (df_1c_code['дозування'] == row_moz['дозування']) &
                        (df_1c_code['кількість одиниць лікарського засобу у споживчій упаковці']
                         == row_moz['кількість одиниць лікарського засобу у споживчій упаковці'])
                    ]
                    if not df_1c_matches.empty:
                        matched_moz_indices.add(index_moz)
                        код_1с = df_1c_matches['код 1с'].values[0]
                        цена = row_moz['задекларована зміна оптово-відпускної ціни']
                        номер_рег = df_1c_matches['номер рег'].values[0]
                        наименование_полное = df_1c_matches['наименование полное'].values[0]
                        output_data.append({
                            'Код 1С': код_1с,
                            'Цена': цена,
                            'Номер рег': номер_рег,
                            'Наименование полное': наименование_полное
                        })
                        break

        output_df = pd.DataFrame(output_data)
        logging.debug(f"Сопоставленных записей: {len(output_df)}")

        # ---- Несопоставленные строки Moz ----
        unmatched_moz_df = df_moz.drop(index=matched_moz_indices)
        logging.debug(f"Несопоставленных записей в Moz: {len(unmatched_moz_df)}")

        # ---- Столбцы, которые хотим включить в "Unmatched_Moz" ----
        columns_to_include_output = [
            'міжнародна непатентована або загальноприйнята назва лікарського засобу',
            'торговельна назва лікарського засобу',
            'форма випуску',
            'дозування',
            'кількість одиниць лікарського засобу у споживчій упаковці',
            'найменування виробника, країна',
            'код атх',
            'номер реєстраційного посвідчення на лікарський засіб',
            'дата закінчення строку дії реєстраційного посвідчення на лікарський засіб',
            'задекларована зміна оптово-відпускної ціни',
            'офіційний курс та вид іноземної валюти',
            'дата та номер наказу моз про декларування змін оптово-відпускної ціни на лікарські засоби'
        ]

        # Проверяем, что все есть в unmatched_moz_df
        missing_columns_output = [col for col in columns_to_include_output if col not in unmatched_moz_df.columns]
        if missing_columns_output:
            logging.error(
                f"Отсутствуют необходимые столбцы в нераспознанных записях Moz: {', '.join(missing_columns_output)}"
            )
            raise ValueError(
                f"Отсутствуют необходимые столбцы в нераспознанных записях Moz: {', '.join(missing_columns_output)}"
            )

        unmatched_moz_df = unmatched_moz_df[columns_to_include_output]

        # ---- Сохранение результата ----
        output_file = os.path.join(UPLOAD_FOLDER, f'output_{unique_id}.xlsx')
        logging.info(f"Сохранение итоговых данных в файл: {output_file}")

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name='Matched', index=False)
            unmatched_moz_df.to_excel(writer, sheet_name='Unmatched_Moz', index=False)

        logging.info(f"Выходной файл успешно сохранён: {output_file}")
        return output_file

    except Exception as e:
        logging.error(f'Ошибка при обработке файлов в process_files: {str(e)}', exc_info=True)
        raise

if __name__ == '__main__':
    # Запуск приложения (debug=True для вывода отладочных сообщений Flask)
    app.run(host='0.0.0.0', port=9100, debug=True)
