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
    """Получение абсолютного пути к ресурсу, работает для разработки и для PyInstaller"""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Настройка логирования
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Создаем файловый обработчик с указанием кодировки utf-8
file_handler = logging.FileHandler('app.log', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# Создаем форматтер и добавляем его в обработчик
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# Добавляем обработчик к логгеру
logger.addHandler(file_handler)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Обработка файла 1C
        if 'file1c' in request.files:
            file1c = request.files['file1c']
            if file1c.filename != '':
                if os.path.exists(FILE_1C_PATH):
                    os.remove(FILE_1C_PATH)
                file1c.save(FILE_1C_PATH)
                logging.info(f"Сохранен файл 1C в {FILE_1C_PATH}")
        else:
            logging.warning("Файл 1C не был загружен.")

        # Обработка файла Moz
        if 'filemoz' not in request.files:
            flash('Файл Moz отсутствует.')
            logging.warning("Файл Moz отсутствует в запросе.")
            return redirect(request.url)

        filemoz = request.files['filemoz']

        if filemoz.filename == '':
            flash('Файл Moz не выбран.')
            logging.warning("Файл Moz не был выбран.")
            return redirect(request.url)

        filemoz_path = os.path.join(UPLOAD_FOLDER, filemoz.filename)
        filemoz.save(filemoz_path)
        logging.info(f"Сохранен файл Moz в {filemoz_path}")

        # Проверка, что файлы сохранены корректно
        if not os.path.exists(FILE_1C_PATH) or not os.path.exists(filemoz_path):
            flash('Ошибка при сохранении файлов.')
            logging.error("Ошибка при сохранении файлов 1C или Moz.")
            return redirect(url_for('index'))

        # Генерация уникального имени для выходного файла
        unique_id = uuid.uuid4().hex
        output_file = process_files(FILE_1C_PATH, filemoz_path, unique_id)
        return send_file(output_file, as_attachment=True)

    except Exception as e:
        logging.error(f'Ошибка при обработке файла: {str(e)}', exc_info=True)
        flash(f'Ошибка при обработке файла: {str(e)}')
        return redirect(url_for('index'))

def normalize_column_names(df, required_columns, column_mapping):
    """
    Нормализация названий столбцов: приведение к нижнему регистру и удаление лишних пробелов.
    Переименование столбцов на ожидаемые имена с использованием отображения (mapping).
    Если некоторые столбцы не найдены, пытаемся найти близкие совпадения.
    """
    # Приведение названий столбцов к нижнему регистру и удаление лишних пробелов
    df.columns = df.columns.str.strip().str.lower()

    # Создаем копию отображения с ключами в нижнем регистре
    column_mapping_lower = {k.lower(): v for k, v in column_mapping.items()}

    # Переименование столбцов с использованием mapping
    df.rename(columns=column_mapping_lower, inplace=True)

    # Проверяем наличие необходимых столбцов и пытаемся найти близкие совпадения
    missing_columns = [col for col in required_columns if col not in df.columns]
    for col in missing_columns:
        matches = get_close_matches(col, df.columns, n=1, cutoff=0.8)
        if matches:
            df.rename(columns={matches[0]: col}, inplace=True)
            logging.info(f"Переименован столбец '{matches[0]}' в '{col}'")
        else:
            logging.warning(f"Не удалось найти столбец для '{col}'")

    return df

def process_files(file1c, filemoz, unique_id):
    try:
        # Определяем, какой модуль использовать для чтения файлов
        file1c_extension = os.path.splitext(file1c)[1].lower()
        filemoz_extension = os.path.splitext(filemoz)[1].lower()

        if file1c_extension in ['.xls', '.xlsx']:
            df_1c = pd.read_excel(file1c, engine='openpyxl')
        else:
            raise ValueError("Неподдерживаемое расширение файла 1C")

        if filemoz_extension in ['.xls', '.xlsx']:
            df_moz = pd.read_excel(filemoz, engine='openpyxl')
        else:
            raise ValueError("Неподдерживаемое расширение файла Moz")

        logging.info("Файлы успешно прочитаны")

        # Логирование столбцов файла Moz
        logging.debug(f"Столбцы в файле Moz: {list(df_moz.columns)}")

        # Список необходимых столбцов для Moz
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

        # Отображение: фактическое название -> ожидаемое название
        column_mapping_moz = {
            'офіційний курс та вид іноземної валюти, встановлений національним банком україни на дату подання декларації зміни оптово-відпускної ціни на лікарський засіб*': 'офіційний курс та вид іноземної валюти'
            # Добавьте другие необходимые отображения здесь
        }

        # Нормализация названий столбцов
        df_moz = normalize_column_names(df_moz, required_columns_moz, column_mapping_moz)

        # Обновленный список необходимых столбцов после нормализации
        required_columns_moz_normalized = [
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

        # Проверяем наличие необходимых столбцов
        missing_columns_moz = [col for col in required_columns_moz_normalized if col not in df_moz.columns]
        if missing_columns_moz:
            logging.error(f"Отсутствуют необходимые столбцы в файле Moz: {', '.join(missing_columns_moz)}")
            raise ValueError(f"Отсутствуют необходимые столбцы в файле Moz: {', '.join(missing_columns_moz)}")

        # Нормализация названий столбцов для 1C
        required_columns_1c = ['номер рег', 'код 1с', 'наименование полное', 'форма випуску', 'дозування', 'кількість одиниць лікарського засобу у споживчій упаковці']
        column_mapping_1c = {
            # Добавьте отображения для столбцов 1C, если необходимо
            # Например:
            # 'номер регистрации': 'номер рег'
        }

        df_1c = normalize_column_names(df_1c, required_columns_1c, column_mapping_1c)

        # Проверяем наличие необходимых столбцов в 1C
        missing_columns_1c = [col for col in required_columns_1c if col not in df_1c.columns]
        if missing_columns_1c:
            logging.error(f"Отсутствуют необходимые столбцы в файле 1C: {', '.join(missing_columns_1c)}")
            raise ValueError(f"Отсутствуют необходимые столбцы в файле 1C: {', '.join(missing_columns_1c)}")

        # Фильтрация данных из 1C
        df_1c_filtered = df_1c[df_1c['номер рег'].str.contains('UA', na=False)]
        output_data = []
        matched_moz_indices = set()

        if 'номер реєстраційного посвідчення на лікарський засіб' not in df_moz.columns:
            raise ValueError("Отсутствует столбец 'номер реєстраційного посвідчення на лікарський засіб' в файле Moz")

        unique_codes = df_moz['номер реєстраційного посвідчення на лікарський засіб'].unique()

        for code in unique_codes:
            df_moz_code = df_moz[df_moz['номер реєстраційного посвідчення на лікарський засіб'] == code]
            df_1c_code = df_1c_filtered[df_1c_filtered['номер рег'] == code]

            if len(df_moz_code) == 1 and len(df_1c_code) == 1:
                matched_moz_indices.add(df_moz_code.index[0])
                код_1с = df_1c_code['код 1с'].values[0]
                цена = df_moz_code['задекларована зміна оптово-відпускної ціни'].values[0]
                номер_рег = df_1c_code['номер рег'].values[0]
                наименование_полное = df_1c_code['наименование полное'].values[0]
                output_data.append(
                    {'Код 1С': код_1с, 'Цена': цена, 'Номер рег': номер_рег, 'Наименование полное': наименование_полное})
            else:
                for index_moz, row_moz in df_moz_code.iterrows():
                    df_1c_matches = df_1c_code[
                        (df_1c_code['форма випуску'] == row_moz['форма випуску']) &
                        (df_1c_code['дозування'] == row_moz['дозування']) &
                        (df_1c_code['кількість одиниць лікарського засобу у споживчій упаковці'] == row_moz[
                            'кількість одиниць лікарського засобу у споживчій упаковці'])
                        ]
                    if not df_1c_matches.empty:
                        matched_moz_indices.add(index_moz)
                        код_1с = df_1c_matches['код 1с'].values[0]
                        цена = row_moz['задекларована зміна оптово-відпускної ціни']
                        номер_рег = df_1c_matches['номер рег'].values[0]
                        наименование_полное = df_1c_matches['наименование полное'].values[0]
                        output_data.append({'Код 1С': код_1с, 'Цена': цена, 'Номер рег': номер_рег,
                                            'Наименование полное': наименование_полное})
                        break

        output_df = pd.DataFrame(output_data)

        # Определяем строки Moz, которые не были найдены в 1C
        unmatched_moz_df = df_moz.drop(index=matched_moz_indices)

        # Выбираем только необходимые столбцы для unmatched_moz_df
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

        # Проверяем наличие необходимых столбцов
        missing_columns_output = [col for col in columns_to_include_output if col not in unmatched_moz_df.columns]
        if missing_columns_output:
            logging.error(f"Отсутствуют необходимые столбцы в файле Moz: {', '.join(missing_columns_output)}")
            raise ValueError(f"Отсутствуют необходимые столбцы в файле Moz: {', '.join(missing_columns_output)}")

        unmatched_moz_df = unmatched_moz_df[columns_to_include_output]

        # Генерация уникального имени для выходного файла
        output_file = os.path.join(UPLOAD_FOLDER, f'output_{unique_id}.xlsx')
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name='Matched', index=False)
            unmatched_moz_df.to_excel(writer, sheet_name='Unmatched_Moz', index=False)
        logging.info(f"Выходной файл сохранен в {output_file}")
        return output_file

    except Exception as e:
        logging.error(f'Ошибка при обработке файлов: {str(e)}', exc_info=True)
        raise

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=9100, debug=True)
