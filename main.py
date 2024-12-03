import os
import pdfplumber
import pytesseract
from PIL import Image
import pandas as pd

# Укажите путь к Tesseract на вашей машине (для Windows)
pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'  # Укажите путь к Tesseract


def load_identifiers_from_excel(excel_file):
    """
    Загрузка идентификаторов из Excel файла.
    Столбцы файла должны быть: 'id' и 'string'.
    """
    df = pd.read_excel(excel_file)
    identifiers = dict(zip(df['string'], df['id']))  # Создаем словарь {ключевое_слово: идентификатор}
    return identifiers


def extract_text_from_image(image):
    """
    Используем Tesseract для извлечения текста из изображения.
    """
    text = pytesseract.image_to_string(image, lang='rus')  # Указываем 'rus' для русского языка
    return text


def extract_text_from_pdf(pdf_file):
    """
    Извлечение текста из сканированного PDF файла с использованием OCR.
    """
    text = ''
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Извлекаем изображение страницы
            image = page.to_image()
            pil_image = image.original  # Получаем изображение как объект PIL
            # Извлекаем текст с помощью OCR
            text += extract_text_from_image(pil_image)
    return text


def rename_files_in_directory(directory, identifiers):
    """
    Поочередно открывает PDF файлы в папке, извлекает текст через OCR и ищет ключевые слова.
    Если ключевое слово найдено, файл переименовывается на соответствующий идентификатор.
    Добавляет порядковый номер к файлам с одинаковыми названиями.
    """
    renamed_files = {}  # Словарь для отслеживания количества файлов с одинаковыми идентификаторами

    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        if os.path.isfile(filepath) and filename.lower().endswith(".pdf"):  # Проверяем, что это PDF файл
            # Извлекаем текст из PDF файла
            content = extract_text_from_pdf(filepath)

            # Проверяем, есть ли ключевое слово в содержимом файла
            for keyword, identifier in identifiers.items():
                if keyword in content:
                    # Если файл с таким именем уже существует, добавляем порядковый номер
                    if identifier in renamed_files:
                        renamed_files[identifier] += 1
                    else:
                        renamed_files[identifier] = 1

                    # Добавляем порядковый номер к имени файла, если он не первый
                    if renamed_files[identifier] > 1:
                        new_filename = f"{identifier}({renamed_files[identifier]}).pdf"
                    else:
                        new_filename = f"{identifier}.pdf"

                    new_filepath = os.path.join(directory, new_filename)

                    # Переименовываем файл
                    os.rename(filepath, new_filepath)
                    print(f"Файл {filename} переименован в {new_filename}")
                    break


if __name__ == "__main__":
    directory = "D:\workaem\pdf"  # Замените на путь к папке с PDF файлами
    excel_file = "D:\workaem\Excel.xlsx"  # Замените на путь к вашему Excel файлу

    # Загрузка идентификаторов из Excel файла
    identifiers = load_identifiers_from_excel(excel_file)

    # Поиск и переименование PDF файлов
    rename_files_in_directory(directory, identifiers)
