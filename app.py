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

# Убедимся, что папка uploads существует
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# ------------------------------------------------------------------------------
# Настройка логирования (В файл и в консоль)
# ------------------------------------------------------------------------------
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

file_handler = logging.FileHandler('app.log', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)


@app.route('/')
def index():
    """Главная страница."""
    logging.info("Переход на главную страницу (index).")
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Обрабатывает загрузку файлов и запускает обработку данных."""
    try:
        logging.info("Начало обработки запроса /upload")

        # Проверка и сохранение файла 1C
        if 'file1c' in request.files:
            file1c = request.files['file1c']
            if file1c.filename != '':
                if os.path.exists(FILE_1C_PATH):
                    os.remove(FILE_1C_PATH)
                    logging.debug("Старый файл 1C удален.")
                file1c.save(FILE_1C_PATH)
                logging.info(f"Файл 1C сохранен в {FILE_1C_PATH}")
            else:
                logging.warning("Файл 1C выбран, но имя файла пустое.")

        # Проверка файла Moz
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
        logging.info(f"Файл Moz сохранен в {filemoz_path}")

        unique_id = uuid.uuid4().hex
        logging.debug(f"Сгенерирован уникальный ID: {unique_id}")

        output_file = process_files(FILE_1C_PATH, filemoz_path, unique_id)
        logging.info(f"Обработка завершена. Отправляем результат: {output_file}")

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        logging.error(f'Ошибка при обработке файла: {str(e)}', exc_info=True)
        flash(f'Ошибка при обработке файла: {str(e)}')
        return redirect(url_for('index'))


def process_files(file1c, filemoz, unique_id):
    """Основная логика обработки файлов."""
    try:
        logging.info("Начало обработки файлов в process_files.")

        df_1c = pd.read_excel(file1c, engine='openpyxl')
        df_moz = pd.read_excel(filemoz, engine='openpyxl')

        logging.debug(f"Столбцы в файле Moz перед обработкой: {list(df_moz.columns)}")

        # Переименование 11-й колонки в "Курс"
        if len(df_moz.columns) >= 11:
            old_name = df_moz.columns[10]
            df_moz.rename(columns={old_name: 'Курс'}, inplace=True)
            logging.info(f"Столбец '{old_name}' переименован в 'Курс'")

        # Проверка, что колонка успешно переименована
        if 'Курс' not in df_moz.columns:
            msg = "Ошибка! Не удалось переименовать 11-й столбец в 'Курс'."
            logging.error(msg)
            raise ValueError(msg)

        logging.debug(f"Столбцы в файле Moz после обработки: {list(df_moz.columns)}")

        output_file = os.path.join(UPLOAD_FOLDER, f'output_{unique_id}.xlsx')

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_moz.to_excel(writer, sheet_name='Moz_Processed', index=False)

        logging.info(f"Выходной файл успешно сохранен в {output_file}")
        return output_file

    except Exception as e:
        logging.error(f'Ошибка при обработке файлов: {str(e)}', exc_info=True)
        raise


if __name__ == '__main__':
    logging.info("Запуск приложения Flask на http://0.0.0.0:9100")
    app.run(host='0.0.0.0', port=9100, debug=True)
