import win32com.client as win32
import os
from PyPDF2 import PdfMerger

# Функция для чтения путей и диапазона из текстового файла
def read_paths_and_range_from_file(txt_file_path):
    try:
        with open(txt_file_path, 'r') as file:
            paths = file.readlines()
            paths = [path.strip() for path in paths]
            if len(paths) < 4:
                print("Ошибка: текстовый файл должен содержать минимум 4 строки.")
                return None, None

            # Читаем пути и диапазон
            folder_path = paths[0]
            output_folder = paths[1]
            final_pdf_folder = paths[2]
            range_string = paths[3]

            # Разбираем диапазон
            range_parts = range_string.split('-')
            if len(range_parts) != 2:
                print("Ошибка: диапазон должен быть в формате '2991-2995'.")
                return None, None

            try:
                start_range = int(range_parts[0])
                end_range = int(range_parts[1])
                return (folder_path, output_folder, final_pdf_folder), (start_range, end_range)
            except ValueError:
                print("Ошибка: диапазон должен содержать только числа.")
                return None, None
    except Exception as e:
        print(f"Произошла ошибка при чтении файла путей: {e}")
        return None, None

def export_first_two_visible_sheets_to_pdf(excel_file_path, output_pdf_path):
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    excel_app.Visible = False

    try:
        workbook = excel_app.Workbooks.Open(excel_file_path)
        sheets = workbook.Sheets
        visible_sheets = [sheet for sheet in sheets if sheet.Visible == -1]

        if len(visible_sheets) < 2:
            print(f"Недостаточно видимых листов для экспорта в файле {excel_file_path}.")
            return

        sheets_to_export = visible_sheets[:2]
        for sheet in sheets:
            if sheet not in sheets_to_export:
                sheet.Visible = False

        first_sheet = sheets_to_export[0]
        first_sheet.PageSetup.PrintArea = "$A$1:$Q$200"

        second_sheet = sheets_to_export[1]
        if second_sheet.PageSetup.PrintArea == "":
            second_sheet.PageSetup.PrintArea = second_sheet.UsedRange.Address

        workbook.ExportAsFixedFormat(
            Type=0,
            Filename=output_pdf_path,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )

        print(f"PDF успешно сохранен для файла {excel_file_path}: {output_pdf_path}.")
    except Exception as e:
        print(f"Произошла ошибка с файлом {excel_file_path}: {e}")
    finally:
        for sheet in sheets:
            sheet.Visible = True
        workbook.Close(SaveChanges=False)
        excel_app.Quit()

def process_folder_recursive(folder_path, output_folder, number_range):
    os.makedirs(output_folder, exist_ok=True)
    pdf_files = []
    start_range, end_range = number_range

    for root, _, files in os.walk(folder_path):
        for file_name in files:
            if file_name.endswith(('.xlsx', '.xlsm')):
                # Извлекаем номер из имени папки
                folder_name = os.path.basename(root)
                try:
                    file_number = int(folder_name.split()[0])  # Берем первое число в названии
                except ValueError:
                    continue

                # Проверяем, входит ли номер в диапазон
                if start_range <= file_number <= end_range:
                    excel_file_path = os.path.join(root, file_name)
                    output_pdf_path = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}_invoice+specification.pdf")
                    export_first_two_visible_sheets_to_pdf(excel_file_path, output_pdf_path)
                    if os.path.exists(output_pdf_path):
                        pdf_files.append(output_pdf_path)

    return pdf_files

def merge_pdfs(pdf_files, final_pdf_folder):
    if not pdf_files:
        print("Нет PDF-файлов для объединения.")
        return

    numbers = []
    for pdf in pdf_files:
        base_name = os.path.basename(pdf)
        number = ''.join(filter(str.isdigit, base_name))
        if number:
            numbers.append(int(number))

    numbers.sort()
    ranges = []
    start = numbers[0]
    end = numbers[0]

    for i in range(1, len(numbers)):
        if numbers[i] == end + 1:
            end = numbers[i]
        else:
            if start == end:
                ranges.append(f"{start}")
            else:
                ranges.append(f"{start}-{end}")
            start = numbers[i]
            end = numbers[i]

    if start == end:
        ranges.append(f"{start}")
    else:
        ranges.append(f"{start}-{end}")

    final_pdf_name = f"Invoice+Specification {';'.join(ranges)}.pdf"
    final_pdf_path = os.path.join(final_pdf_folder, final_pdf_name)

    merger = PdfMerger()
    try:
        for pdf in pdf_files:
            merger.append(pdf)
        merger.write(final_pdf_path)
        merger.close()
        print(f"Объединенный PDF успешно создан: {final_pdf_path}")
    except Exception as e:
        print(f"Ошибка при объединении PDF: {e}")

if __name__ == "__main__":
    paths, number_range = read_paths_and_range_from_file('paths.txt')
    if paths and number_range:
        folder_path, output_folder, final_pdf_folder = paths
        os.makedirs(final_pdf_folder, exist_ok=True)
        pdf_files = process_folder_recursive(folder_path, output_folder, number_range)
        merge_pdfs(pdf_files, final_pdf_folder)
