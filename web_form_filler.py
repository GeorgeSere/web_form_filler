# ИМПОРТ БИБИЛИОТЕК
import os
import json
import requests
import zipfile
import time
import re
import pandas as pd
from threading import Thread
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import WebDriverException
import tkinter as tk
from tkinter import scrolledtext


# ЗАГРУЗКА КОНСТАНТ ИЗ ФАЙЛА config.json
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

FORM_URL = config['form_url']
CHROMEDRIVER_URL = config['chromedriver_url']
DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), config['download_dir'])
RESULT_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), config['result_file_path'])
WAIT_TIMEOUT = config['wait_timeout']


def download_and_extract_chromedriver(url, extract_to='.'):
    local_zip_path = os.path.join(extract_to, 'chromedriver.zip')
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_zip_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    with zipfile.ZipFile(local_zip_path, 'r') as zip_ref:
        for member in zip_ref.namelist():
            filename = os.path.basename(member)
            # Skip directories
            if not filename:
                continue
            # Copy file (taken from zipfile's extract)
            source = zip_ref.open(member)
            target = open(os.path.join(extract_to, filename), "wb")
            with source, target:
                target.write(source.read())
    os.remove(local_zip_path)

# ПОЛУЧЕНИЕ ПУТИ К ПАПКЕ ЗАГРУЗКИ
def get_download_path():
    if os.name == 'nt':  # Windows
        return os.path.join(os.environ['USERPROFILE'], 'Downloads')
    else:  # MacOS, Linux
        return os.path.join(os.path.expanduser('~'), 'Downloads')

# УДАЛЕНИЕ СТАРОГО ФАЙЛА-ИСХОДНИКА, ЕСЛИ ЕСТЬ
def delete_existing_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)
        # Если файл с таким именем в папке загрузок есть ,то удаляем его перед загрузкой нового

# ПОДГОТОВКА ДАННЫХ
def prepare_data(file_path):
    # Чтение таблицы-исходника
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()  # Убираем лишние пробелы в названиях столбцов
    df['Phone Number'] = df['Phone Number'].astype(str) # приводим телефон к типу str
    # Подготовка данных для заполнения
    data_to_fill = []
    error_messages = []
    for index, row in df.iterrows():
        first_name = row['First Name']
        last_name = row['Last Name']
        email = row['Email']
        address = row['Address']

        # Проверка корректности данных
        if not first_name.isalpha():
            error_messages.append(f"Некорректное имя: {first_name}")
            continue
        if not last_name.isalpha():
            error_messages.append(f"Некорректная фамилия: {last_name}")
            continue

        pattern = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') # проверка корректности email (регулярка)
        if not pattern.match(email):
            error_messages.append(f"Некорректный email: {email}")
            continue
            
        if not (any(char.isdigit() for char in address) and any(char.isalpha() for char in address)):
            error_messages.append(f"Некорректный адрес: {address}")
            continue

        data_to_fill.append({
            'first_name': first_name,
            'last_name': last_name,
            'phone': row['Phone Number'],
            'email': email,
            'address': address,
            'company_name': row['Company Name'],
            'role_in_company': row['Role in Company']
        })

    return data_to_fill, error_messages

# ПАРКСИНГ СТРАНИЦЫ С РЕЗУЛЬТАТАМИ
def parse_results(text):
    pattern = r'\d+' #создаём паттерн - поиск в строке одной и более цифр
    results = re.findall(pattern, text) # находим все совпадения pattern в строке и записываем в список results
    return results

# ЗАПИСЬ В РЕЗУЛЬТИРУЮЩИЙ ФАЙЛ
def write_results(results, start_time, error_messages=None):
    end_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

    with open(RESULT_FILE_PATH, 'a') as file: # путь константа из config
        file.write(f"Время начала выполнения: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start_time))}\n")
        file.write(f"Процент успешного заполнения полей: {results[0]} %\n")
        file.write(f"Заполнено полей: {results[1]} / {results[2]}\n")
        file.write(f"Время выполнения: {results[3]} миллисекунд\n")
        if error_messages:
            file.write(f"Ошибка: {error_messages}\n")
        else:
            file.write("Запуск программы прошел успешно\n")
        file.write(f"Время окончания выполнения: {end_time}\n")
        file.write("-" * 40 + "\n")
# ОСНОВНАЯ ФУНКЦИЯ ЗАПОЛНЕНИЯ ВЕБ-ФОРМЫ
def fill_web_form(form_url, output_box, driver):
    start_time = time.time()
    output_box.insert(tk.END, "Заполнение форм началось...\n")
    driver.get(form_url)

    wait = WebDriverWait(driver, WAIT_TIMEOUT) # время ожидания - константа из config файла

    try:
        start_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-large') and contains(text(), 'Start')]"))
        )
        start_button.click()

        download_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn waves-effect waves-light uiColorPrimary') and contains(text(), 'Download Excel')]"))
        )
        download_button.click()

        downloaded_file = os.path.join(DOWNLOAD_DIR, 'challenge.xlsx')

        # Удаление существующего файла перед загрузкой нового
        delete_existing_file(downloaded_file)

        while not os.path.exists(downloaded_file):
            time.sleep(1)

        data_to_fill, error_messages = prepare_data(downloaded_file)
        output_box.insert(tk.END, "Данные подготовлены...\n")

        forms_cnt = 0

        for entry in data_to_fill:
            try:
                last_name_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@ng-reflect-name='labelLastName']"))
                )
                last_name_input.send_keys(entry['last_name'])
                # Проверка введенного значения
                if last_name_input.get_attribute('value') != entry['last_name']:
                    output_box.insert(tk.END, f"Ошибка введенное значение Last Name не совпадает для {entry['last_name']}")

                first_name_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@ng-reflect-name='labelFirstName']"))
                )
                first_name_input.send_keys(entry['first_name'])
                # Проверка введенного значения
                if first_name_input.get_attribute('value') != entry['first_name']:
                    output_box.insert(tk.END, f"Ошибка введенное значение First Name не совпадает для {entry['first_name']}")

                company_name_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@ng-reflect-name='labelCompanyName']"))
                )
                company_name_input.send_keys(entry['company_name'])
                # Проверка введенного значения
                if company_name_input.get_attribute('value') != entry['company_name']:
                    output_box.insert(tk.END, f"Ошибка введенное значение Company Name не совпадает для {entry['company_name']}")

                role_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@ng-reflect-name='labelRole']"))
                )
                role_input.send_keys(entry['role_in_company'])
                # Проверка введенного значения
                if role_input.get_attribute('value') != entry['role_in_company']:
                    output_box.insert(tk.END, f"Ошибка введенное значение Role in Company не совпадает для {entry['role_in_company']}")

                address_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@ng-reflect-name='labelAddress']"))
                )
                address_input.send_keys(entry['address'])
                # Проверка введенного значения
                if address_input.get_attribute('value') != entry['address']:
                    output_box.insert(tk.END, f"Ошибка введенное значение Address не совпадает для {entry['address']}")

                email_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@ng-reflect-name='labelEmail']"))
                )
                email_input.send_keys(entry['email'])
                # Проверка введенного значения
                if email_input.get_attribute('value') != entry['email']:
                    output_box.insert(tk.END, f"Ошибка введенное значение Email не совпадает для {entry['email']}")

                phone_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@ng-reflect-name='labelPhone']"))
                )
                phone_input.send_keys(entry['phone'])
                # Проверка введенного значения
                if phone_input.get_attribute('value') != entry['phone']:
                    output_box.insert(tk.END, f"Ошибка введенное значение Phone не совпадает для {entry['phone']}")

                submit_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Submit']"))
                )
                submit_button.click()

                forms_cnt += 1

            except TimeoutException as e:
                error_message = str(e)
                output_box.insert(tk.END, f"Ошибка: {error_message}\n")
                continue

    except TimeoutException as e:
        error_message = str(e)
        output_box.insert(tk.END, f"Ошибка: {error_message}\n")

    execution_time = time.time() - start_time  # фиксируем время выполнения

    # Парсинг текста с результатами
    result_element = driver.find_element(By.XPATH, "//div[contains(@class, 'message2')]")
    result_text = result_element.text
    results = parse_results(result_text)

    # вывод инфы в окно gui
    output_box.insert(tk.END, f"Процент успещного заполнения полей: {results[0]} % \n")
    output_box.insert(tk.END, f"Заполнено полей: {results[1]} / {results[2]} \n")
    output_box.insert(tk.END, f"Время выполнения: {results[3]} миллисекунд\n")

    # Запись результатов в файл
    write_results(results, start_time, error_messages)

# МНОГОПОТОЧНАЯ ФУНКЦИЯ ЗАПОЛНЕНИЯ ВЕБ-ФОРМЫ (чтобы избежать подвисаний окна gui)
def fill_web_form_threaded(form_url, output_box, driver):
    def run():
        fill_web_form(form_url, output_box, driver)
    thread = Thread(target=run)
    thread.start()

# ФУНКЦИЯ ГРАФИЧЕСКОГО ИНТЕРФЕЙСА
def run_gui():
    try:
        if not os.path.exists(DOWNLOAD_DIR): # путь для загрузки исходника из config
            os.makedirs(DOWNLOAD_DIR)

        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        base_dir = os.path.dirname(os.path.abspath(__file__))
        chromedriver_path = os.path.join(base_dir, 'chromedriver.exe')
        if not os.path.exists(chromedriver_path):
            download_and_extract_chromedriver(CHROMEDRIVER_URL, base_dir)

        driver = webdriver.Chrome(executable_path=chromedriver_path, options=chrome_options)
    except WebDriverException:
        print("Ошибка при запуске WebDriver")
        return

    def on_exit():
        driver.quit()
        window.quit()


    window = tk.Tk()
    window.title('web_form_filler')
    window.configure(background='#8FBC8F')

    # Кнопка заполнения формы
    button_fill_form = tk.Button(window, text='ЗАПОЛНИТЬ ФОРМУ', command=lambda: fill_web_form_threaded(FORM_URL, output_box, driver), bg='#ffebcd') # url формы константа из config
    button_fill_form.grid(row=3, column=0, columnspan=5, padx=5, pady=5)

    # Окно вывода
    output_box = scrolledtext.ScrolledText(window, width=60, height=20, bg='white')
    output_box.grid(row=4, column=0, columnspan=5, padx=5, pady=5)

    # Кнопка выход
    button_exit = tk.Button(window, text='ВЫХОД', command=on_exit, bg='#ffdab9')
    button_exit.grid(row=5, column=0, columnspan=5, padx=5, pady=5)

    window.mainloop()

run_gui()
