import csv
import json
import os
import platform
import subprocess
import sys
from pathlib import Path

try:
    from weasyprint import HTML
except ImportError:
    print("Ошибка: библиотека WeasyPrint не найдена.")
    print("Пожалуйста, установите ее, выполнив команду:")
    print("pip install WeasyPrint")
    sys.exit(1)


def setup_directories():
    """Создает необходимые директории, если они не существуют."""
    for dir_name in ["data", "templates", "output", "temp"]:
        Path(dir_name).mkdir(exist_ok=True)


def find_files(directory, extensions):
    """Рекурсивно ищет файлы с заданными расширениями в директории."""
    found_files = []
    for ext in extensions:
        found_files.extend(Path(directory).rglob(f"*.{ext}"))
    return found_files


def select_item(items, title):
    """
    Выводит нумерованный список элементов и просит пользователя выбрать один.
    """
    print(f"\n--- {title} ---")
    if not items:
        print(f"Файлы не найдены. Проверьте, что они существуют.")
        return None

    for i, item in enumerate(items, 1):
        print(f"{i}. {item}")

    while True:
        try:
            choice = int(input(f"Выберите номер (1-{len(items)}): "))
            if 1 <= choice <= len(items):
                return items[choice - 1]
            else:
                print("Неверный номер. Попробуйте еще раз.")
        except ValueError:
            print("Пожалуйста, введите число.")


def load_data(file_path):
    """Загружает данные из CSV или JSON файла и приводит их к единой структуре."""
    try:
        with file_path.open("r", encoding="utf-8") as f:
            if file_path.suffix == ".csv":
                reader = csv.DictReader(f)
                invoices = {}
                for row in reader:
                    invoice_id = row["invoice_id"]
                    if invoice_id not in invoices:
                        invoices[invoice_id] = {
                            "invoice_id": invoice_id,
                            "customer_name": row.get("customer_name"),
                            "date": row.get("date"),
                            "items": [],
                        }
                    invoices[invoice_id]["items"].append({
                        "item_name": row.get("item_name"),
                        "quantity": row.get("quantity", 1),
                        "price": row.get("price", 0),
                        "amount": row.get("amount", 0),
                    })
                return list(invoices.values())
            elif file_path.suffix == ".json":
                data = json.load(f)
                invoices = {}
                for record in data:
                    invoice_id = record["invoice_id"]
                    if invoice_id not in invoices:
                        invoices[invoice_id] = {
                            "invoice_id": invoice_id,
                            "customer_name": record.get("customer_name"),
                            "date": record.get("date"),
                            "items": [],
                        }
                    # Обрабатываем и вложенные списки items, и плоскую структуру
                    if 'items' in record and isinstance(record['items'], list):
                        invoices[invoice_id]['items'].extend(record['items'])
                    else:
                        invoices[invoice_id]['items'].append({
                            "item_name": record.get("item_name"),
                            "quantity": record.get("quantity", 1),
                            "price": record.get("price", 0),
                            "amount": record.get("amount", 0),
                        })
                return list(invoices.values())
    except (IOError, json.JSONDecodeError, csv.Error, KeyError) as e:
        print(f"Ошибка при чтении или обработке файла {file_path}: {e}")
        return None


def select_records(records):
    """Позволяет пользователю выбрать записи для обработки по номеру в списке."""
    print("\n--- Выбор записей ---")
    if not records:
        print("Нет доступных записей.")
        return []

    # Отображаем записи с номерами для выбора
    for i, record in enumerate(records, 1):
        print(f"{i}. ID: {record.get('invoice_id', 'N/A')}, Покупатель: {record.get('customer_name', 'N/A')}")

    print("\nВведите номер записи, несколько номеров через запятую, или 'all' для всех.")
    
    while True:
        choice = input("Ваш выбор: ").strip().lower()
        if choice == 'all':
            return records

        if not choice:
            print("Вы ничего не ввели. Попробуйте еще раз.")
            continue

        try:
            # Преобразуем введенные номера в индексы
            selected_indices = [int(item.strip()) - 1 for item in choice.split(',')]
            
            valid_indices = []
            invalid_numbers = []
            for index in selected_indices:
                if 0 <= index < len(records):
                    valid_indices.append(index)
                else:
                    invalid_numbers.append(str(index + 1))
            
            if invalid_numbers:
                print(f"Неверные номера в списке: {', '.join(invalid_numbers)}. Доступно номеров от 1 до {len(records)}.")
                continue

            # Собираем выбранные записи, избегая дубликатов
            selected_records = [records[i] for i in sorted(list(set(valid_indices)))]
            
            if selected_records:
                return selected_records
            else:
                print("Не удалось выбрать записи. Попробуйте еще раз.")

        except ValueError:
            print("Ошибка. Пожалуйста, введите номера (числа) через запятую или 'all'.")


def generate_html(template_content, record):
    """Заполняет HTML-шаблон данными из записи."""
    content = template_content
    
    # Заполняем основные поля
    content = content.replace("{{ invoice_id }}", str(record.get("invoice_id", "")))
    content = content.replace("{{ customer_name }}", str(record.get("customer_name", "")))
    content = content.replace("{{ date }}", str(record.get("date", "")))

    # Генерируем строки таблицы для товаров
    item_rows = ""
    total_amount = 0
    for item in record.get("items", []):
        item_rows += "<tr>"
        item_rows += f"<td>{item.get('item_name', '')}</td>"
        item_rows += f"<td>{item.get('quantity', 0)}</td>"
        item_rows += f"<td>{item.get('price', 0)}</td>"
        item_rows += f"<td>{item.get('amount', 0)}</td>"
        item_rows += "</tr>\n"
        try:
            total_amount += float(item.get("amount", 0))
        except (ValueError, TypeError):
            pass # Игнорируем ошибки преобразования, если сумма некорректна

    content = content.replace("{{ item_rows }}", item_rows)
    content = content.replace("{{ total_amount }}", f"{total_amount:,.2f}".replace(",", " "))
    
    return content


def open_file(path):
    """Кросс-платформенно открывает файл."""
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(path)
        elif system == "Darwin":  # macOS
            subprocess.run(["open", path], check=True)
        else:  # Linux
            subprocess.run(["xdg-open", path], check=True)
    except (OSError, subprocess.CalledProcessError) as e:
        print(f"Не удалось открыть файл {path}: {e}")


def main():
    """Основная функция скрипта."""
    setup_directories()

    # 1. Загрузка данных
    data_files = find_files("data", ["csv", "json"])
    data_file_path = select_item(data_files, "Выберите файл с данными")
    if not data_file_path:
        return
    
    records = load_data(data_file_path)
    if not records:
        print("Не удалось загрузить данные.")
        return

    # 2. Загрузка шаблона
    template_files = find_files("templates", ["html"])
    template_file_path = select_item(template_files, "Выберите HTML-шаблон")
    if not template_file_path:
        return
        
    try:
        template_content = template_file_path.read_text(encoding="utf-8")
    except IOError as e:
        print(f"Ошибка чтения шаблона {template_file_path}: {e}")
        return

    # 3. Выбор записей
    selected_records = select_records(records)
    if not selected_records:
        print("Записи не выбраны. Завершение работы.")
        return

    # 4 & 5. Генерация HTML и PDF
    print("\nНачинаю генерацию PDF...")
    generated_files = []
    template_name = template_file_path.stem

    for record in selected_records:
        invoice_id = record.get("invoice_id", "unknown")
        
        # Генерация HTML
        html_content = generate_html(template_content, record)
        temp_html_path = Path("temp") / f"{invoice_id}_{template_name}.html"
        try:
            temp_html_path.write_text(html_content, encoding="utf-8")
        except IOError as e:
            print(f"Не удалось сохранить временный HTML: {e}")
            continue

        # Генерация PDF
        pdf_filename = f"({invoice_id})_{template_name}.pdf"
        output_pdf_path = Path("output") / pdf_filename
        
        try:
            HTML(string=html_content).write_pdf(output_pdf_path)
            print(f"  \u2713 Создан: {output_pdf_path}")
            generated_files.append(output_pdf_path)
        except Exception as e:
            print(f"  \u2717 Ошибка при создании PDF для ID {invoice_id}: {e}")

    # 6. Открытие PDF
    if generated_files:
        print("\nГенерация завершена.")
        open_choice = input("Открыть первый сгенерированный PDF? (y/n): ").lower()
        if open_choice == 'y':
            open_file(generated_files[0])
    else:
        print("\nНе было сгенерировано ни одного файла.")


if __name__ == "__main__":
    main()

