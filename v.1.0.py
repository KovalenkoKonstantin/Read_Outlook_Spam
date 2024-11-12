import os
import zipfile
import pandas as pd
import win32com.client
from io import BytesIO

def extract_excel_from_zip(zip_path):
    """Извлекает все Excel файлы из архива .zip."""
    excel_files = []
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
                with zip_ref.open(file_name) as file:
                    excel_files.append(pd.read_excel(file))
    return excel_files


def get_spam_folder():
    """Получает пользовательскую папку 'спам' в Outlook (включая вложенные папки)."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # Получаем все папки в почтовом ящике
        folders = outlook.Folders
        spam_folder = None

        # Проходим по всем папкам и подкаталогам
        for folder in folders:
            # print(f"Обнаружена папка: {folder.Name}")  # Выводим название каждой папки для диагностики
            if folder.Name == 'Спам':
                spam_folder = folder
                break
            # Если папка не найдена, проверяем ее подкаталоги
            if not spam_folder:
                for subfolder in folder.Folders:
                    # print(f"Обнаружен подкаталог: {subfolder.Name}")  # Выводим подкаталоги
                    if subfolder.Name == 'Спам':
                        spam_folder = subfolder
                        break
        if not spam_folder:
            print("Папка 'Спам' не найдена.")
        return spam_folder
    except Exception as e:
        print(f"Ошибка при получении папки 'Спам': {e}")
        return None


def get_last_email_from_spam_folder():
    """Получает последнее письмо из пользовательской папки 'спам'."""
    try:
        spam_folder = get_spam_folder()
        if not spam_folder:
            return None

        messages = spam_folder.Items
        # Сортируем по времени получения, последние сверху
        # messages.Sort("[ReceivedTime]", False)

        # Получаем последнее письмо
        last_message = messages.GetLast()
        return last_message
    except Exception as e:
        print(f"Ошибка при получении письма из папки 'Спам': {e}")
        return None

def process_email_attachments(email):
    """Обрабатывает вложения в письме, извлекая Excel файлы из архивов .zip."""
    try:
        attachments = email.Attachments
        for attachment in attachments:
            if attachment.FileName.endswith('.zip'):
                print(f'Архив: {attachment.FileName}')

                # Сохраняем файл .zip во временную папку
                temp_zip_path = os.path.join(os.getenv('TEMP'), attachment.FileName)
                attachment.SaveAsFile(temp_zip_path)

                # Извлекаем Excel файлы из архива
                excel_files = extract_excel_from_zip(temp_zip_path)

                for i, excel_df in enumerate(excel_files):

                    # Выводим первую строку и извлекаем нужную часть текста
                    first_row_value = excel_df.iloc[0, 0]

                    # Применяем метод, чтобы получить первое слово из строки
                    first_row_text = str(first_row_value).split(':')[0].strip()
                    print(f'Файл: {first_row_text}')

                    # Очищаем от пустых строк, если таковые есть
                    clean_df = excel_df.dropna(how='all')

                    # Получаем последнюю строку с данными
                    last_row = clean_df.iloc[-1]  # Последняя строка
                    # print(f'Последняя строка с данными: {last_row}')
                    print(clean_df)
                    print()

                # Удаляем временный zip файл после обработки
                os.remove(temp_zip_path)
    except Exception as e:
        print(f'Ошибка при обработке вложений: {e}')

def main():
    email = get_last_email_from_spam_folder()
    if email:
        print(f"Обрабатываем письмо от {email.SenderName}, "
              f"получено: {email.ReceivedTime}")
        process_email_attachments(email)
    else:
        print("Нет доступных писем в папке 'Спам'.")

if __name__ == "__main__":
    main()
