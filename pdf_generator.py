#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CLI-скрипт для генерации PDF из CSV данных и HTML шаблонов.
"""

import os
import sys
import csv
import datetime
import subprocess
from pathlib import Path
from typing import List, Tuple, Optional

# Исправление кодировки для Windows консоли
if sys.platform == 'win32':
    try:
        # Пытаемся установить UTF-8 кодировку для консоли
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
    except:
        # Если не получилось, пробуем через chcp
        try:
            os.system('chcp 65001 >nul 2>&1')
        except:
            pass

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

try:
    import pdfkit
    PDFKIT_AVAILABLE = True
except ImportError:
    PDFKIT_AVAILABLE = False

WEASYPRINT_AVAILABLE = False
# Проверяем weasyprint независимо от pdfkit
try:
    import weasyprint
    # Пробуем создать объект HTML, чтобы проверить доступность библиотек
    try:
        test_html = weasyprint.HTML(string='<html><body>test</body></html>')
        # Если дошли сюда без ошибок, значит библиотеки загрузились
        WEASYPRINT_AVAILABLE = True
    except (OSError, Exception) as e:
        # Библиотеки weasyprint установлены, но не могут загрузиться (часто на Windows)
        # Это может быть временная проблема, попробуем еще раз при генерации
        WEASYPRINT_AVAILABLE = False
except ImportError:
    WEASYPRINT_AVAILABLE = False


def find_wkhtmltopdf() -> Optional[str]:
    """Ищет путь к wkhtmltopdf.exe на Windows."""
    if sys.platform != 'win32':
        return None
    
    # Сначала проверяем PATH
    import shutil
    path_in_env = shutil.which('wkhtmltopdf')
    if path_in_env and os.path.exists(path_in_env):
        return path_in_env
    
    # Стандартные пути установки на Windows
    possible_paths = [
        r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe',
        r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe',
        r'C:\wkhtmltopdf\bin\wkhtmltopdf.exe',
        r'C:\ProgramData\wkhtmltopdf\bin\wkhtmltopdf.exe',
        os.path.expanduser(r'~\AppData\Local\Programs\wkhtmltopdf\bin\wkhtmltopdf.exe'),
    ]
    
    # Проверяем стандартные пути
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # Пробуем найти через where (Windows команда)
    try:
        result = subprocess.run(['where', 'wkhtmltopdf'], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0 and result.stdout.strip():
            found_path = result.stdout.strip().split('\n')[0].strip()
            if os.path.exists(found_path):
                return found_path
    except Exception:
        pass
    
    return None


def get_files_in_directory(directory: str, extension: str) -> List[str]:
    """Получает список файлов с указанным расширением в директории."""
    if not os.path.exists(directory):
        return []
    
    files = []
    for file in os.listdir(directory):
        if file.lower().endswith(extension.lower()):
            files.append(file)
    return sorted(files)


def display_file_list(files: List[str], file_type: str) -> None:
    """Выводит пронумерованный список файлов."""
    if not files:
        print(f"\n{file_type} не найдены.")
        return
    
    print(f"\n{file_type}:")
    for idx, file in enumerate(files, 1):
        print(f"  {idx}. {file}")


def get_user_choice(files: List[str], file_type: str) -> Optional[str]:
    """Получает выбор пользователя по номеру файла."""
    if not files:
        return None
    
    while True:
        try:
            choice = input(f"\nВыберите {file_type} (номер): ").strip()
            if not choice:
                print("Введите номер файла.")
                continue
            
            idx = int(choice) - 1
            if 0 <= idx < len(files):
                return files[idx]
            else:
                print(f"Неверный номер. Введите число от 1 до {len(files)}.")
        except ValueError:
            print("Введите корректное число.")
        except (KeyboardInterrupt, EOFError):
            print("\n\nОперация отменена пользователем.")
            sys.exit(0)


def read_csv_data(csv_path: str) -> Tuple[List[dict], Optional[str]]:
    """Читает CSV файл и возвращает данные и имя поставщика."""
    try:
        if PANDAS_AVAILABLE:
            df = pd.read_csv(csv_path, encoding='utf-8')
            # Проверяем наличие необходимых колонок
            required_columns = ['поставщик', 'товар', 'цена', 'количество']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Отсутствуют необходимые колонки: {', '.join(missing_columns)}")
            
            # Получаем имя поставщика (должно быть одинаковым для всех строк)
            contractor = df['поставщик'].iloc[0] if len(df) > 0 else ""
            
            # Преобразуем в список словарей
            data = df.to_dict('records')
            return data, contractor
        else:
            # Используем стандартный csv модуль
            data = []
            contractor = None
            with open(csv_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                required_columns = ['поставщик', 'товар', 'цена', 'количество']
                for row in reader:
                    # Проверяем наличие всех необходимых колонок
                    missing = [col for col in required_columns if col not in row]
                    if missing:
                        raise ValueError(f"Отсутствуют необходимые колонки: {', '.join(missing)}")
                    
                    if contractor is None:
                        contractor = row['поставщик']
                    elif row['поставщик'] != contractor:
                        print(f"Предупреждение: Найдены разные поставщики. Используется первый: {contractor}")
                    
                    data.append({
                        'поставщик': row['поставщик'],
                        'товар': row['товар'],
                        'цена': row['цена'],
                        'количество': row['количество']
                    })
            
            return data, contractor
    except Exception as e:
        print(f"Ошибка при чтении CSV файла: {e}")
        raise


def read_template(template_path: str) -> str:
    """Читает HTML шаблон."""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Ошибка при чтении HTML шаблона: {e}")
        raise


def generate_html(template: str, data: List[dict], contractor: str) -> str:
    """Генерирует HTML из шаблона и данных."""
    import re
    
    # Заменяем плейсхолдер поставщика
    html = template.replace('{{contractor}}', str(contractor))
    
    # Ищем tbody и строку таблицы с плейсхолдерами внутри tbody
    # Это гарантирует, что мы не затронем thead
    tbody_pattern = r'(<tbody[^>]*>)(.*?)(</tbody>)'
    tbody_match = re.search(tbody_pattern, html, re.DOTALL | re.IGNORECASE)
    
    if tbody_match:
        tbody_start = tbody_match.group(1)
        tbody_content = tbody_match.group(2)
        tbody_end = tbody_match.group(3)
        
        # Ищем строку с плейсхолдерами внутри tbody
        row_pattern = r'<tr[^>]*>.*?(?:\{\{product\}\}|\{\{price\}\}|\{\{qty\}\}|\{\{index\}\}|\{\{total\}\}).*?</tr>'
        row_match = re.search(row_pattern, tbody_content, re.DOTALL | re.IGNORECASE)
        
        if row_match:
            row_template = row_match.group(0)
            rows_html = []
            
            for idx, row in enumerate(data, 1):
                row_html = row_template
                # Заменяем плейсхолдеры в строке
                row_html = row_html.replace('{{product}}', str(row['товар']))
                row_html = row_html.replace('{{price}}', str(row['цена']))
                row_html = row_html.replace('{{qty}}', str(row['количество']))
                # Если в шаблоне есть номер строки, заменяем его (поддерживаем разные форматы)
                row_html = re.sub(r'\{\{index\}\}', str(idx), row_html, flags=re.IGNORECASE)
                # Вычисляем сумму, если нужно
                if '{{total}}' in row_html:
                    try:
                        price = float(str(row['цена']).replace(',', '.'))
                        qty = float(str(row['количество']).replace(',', '.'))
                        total = price * qty
                        row_html = row_html.replace('{{total}}', f'{total:.2f}')
                    except (ValueError, TypeError):
                        row_html = row_html.replace('{{total}}', '')
                
                rows_html.append(row_html)
            
            # Заменяем содержимое tbody: убираем шаблон строки и вставляем все строки
            new_tbody_content = tbody_content.replace(row_template, '\n            '.join(rows_html))
            # Заменяем весь tbody в HTML
            html = html.replace(tbody_match.group(0), tbody_start + new_tbody_content + tbody_end)
        else:
            # tbody найден, но строка с плейсхолдерами не найдена
            print("Предупреждение: В tbody не найдена строка таблицы с плейсхолдерами. Используется простая замена.")
            # Пробуем простую замену только внутри tbody
            new_tbody_content = tbody_content
            for row in data:
                if '{{product}}' in new_tbody_content:
                    new_tbody_content = new_tbody_content.replace('{{product}}', str(row['товар']), 1)
                if '{{price}}' in new_tbody_content:
                    new_tbody_content = new_tbody_content.replace('{{price}}', str(row['цена']), 1)
                if '{{qty}}' in new_tbody_content:
                    new_tbody_content = new_tbody_content.replace('{{qty}}', str(row['количество']), 1)
            html = html.replace(tbody_match.group(0), tbody_start + new_tbody_content + tbody_end)
    else:
        # Если tbody не найден, пытаемся найти строку в любом месте (старый метод)
        row_pattern = r'<tr[^>]*>.*?(?:\{\{product\}\}|\{\{price\}\}|\{\{qty\}\}).*?</tr>'
        match = re.search(row_pattern, html, re.DOTALL | re.IGNORECASE)
        
        if match:
            row_template = match.group(0)
            rows_html = []
            
            for idx, row in enumerate(data, 1):
                row_html = row_template
                row_html = row_html.replace('{{product}}', str(row['товар']))
                row_html = row_html.replace('{{price}}', str(row['цена']))
                row_html = row_html.replace('{{qty}}', str(row['количество']))
                row_html = re.sub(r'\{\{index\}\}', str(idx), row_html, flags=re.IGNORECASE)
                if '{{total}}' in row_html:
                    try:
                        price = float(str(row['цена']).replace(',', '.'))
                        qty = float(str(row['количество']).replace(',', '.'))
                        total = price * qty
                        row_html = row_html.replace('{{total}}', f'{total:.2f}')
                    except (ValueError, TypeError):
                        row_html = row_html.replace('{{total}}', '')
                rows_html.append(row_html)
            
            html = html.replace(row_template, '\n            '.join(rows_html))
        else:
            # Если паттерн не найден, пытаемся найти и заменить плейсхолдеры по одному
            # Это менее надежный метод, но может работать для простых шаблонов
            print("Предупреждение: Не найдена строка таблицы с плейсхолдерами. Используется простая замена.")
            for row in data:
                if '{{product}}' in html:
                    html = html.replace('{{product}}', str(row['товар']), 1)
                if '{{price}}' in html:
                    html = html.replace('{{price}}', str(row['цена']), 1)
                if '{{qty}}' in html:
                    html = html.replace('{{qty}}', str(row['количество']), 1)
    
    # Убираем оставшиеся плейсхолдеры (если они есть)
    html = html.replace('{{contractor}}', str(contractor))
    # Очищаем любые оставшиеся плейсхолдеры (на случай, если они не были использованы)
    html = re.sub(r'\{\{[^}]+\}\}', '', html)
    
    return html


def generate_pdf(html_content: str, output_path: str) -> None:
    """Генерирует PDF из HTML."""
    if PDFKIT_AVAILABLE:
        config = None
        wkhtmltopdf_path = find_wkhtmltopdf()
        
        if wkhtmltopdf_path:
            print(f"Найден wkhtmltopdf: {wkhtmltopdf_path}")
            config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
        else:
            # Пробуем использовать конфигурацию по умолчанию
            try:
                config = pdfkit.configuration()
            except Exception:
                pass
        
        if config:
            try:
                pdfkit.from_string(html_content, output_path, configuration=config)
                return
            except Exception as e:
                print(f"Ошибка при генерации PDF с pdfkit: {e}")
        
        # Если pdfkit не сработал, пробуем weasyprint
        print("Попытка использовать weasyprint...")
        try:
            import weasyprint
            weasyprint.HTML(string=html_content).write_pdf(output_path)
            print("PDF успешно создан с помощью weasyprint!")
            return
        except ImportError:
            pass  # weasyprint не установлен
        except Exception as e2:
            # weasyprint установлен, но не работает
            error_msg = f"pdfkit не смог сгенерировать PDF, и weasyprint также не работает.\n"
            error_msg += f"Ошибка weasyprint: {e2}\n\n"
            error_msg += "Решения:\n"
            error_msg += "1. Установите wkhtmltopdf с https://wkhtmltopdf.org/downloads.html\n"
            error_msg += "   И добавьте его в PATH (или переустановите с опцией 'Add to PATH')\n"
            error_msg += "2. Для weasyprint убедитесь, что установлены все зависимости GTK\n"
            error_msg += "   и перезапустите скрипт"
            raise Exception(error_msg)
        
        if not WEASYPRINT_AVAILABLE:
            error_msg = "pdfkit не смог найти wkhtmltopdf, а weasyprint недоступен.\n\n"
            error_msg += "Решения:\n"
            error_msg += "1. Установите wkhtmltopdf с https://wkhtmltopdf.org/downloads.html\n"
            error_msg += "   И добавьте его в PATH (или переустановите с опцией 'Add to PATH')\n"
            error_msg += "2. Или установите weasyprint: pip install weasyprint\n"
            error_msg += "   (требует GTK для Windows)"
            raise Exception(error_msg)
    elif WEASYPRINT_AVAILABLE:
        try:
            import weasyprint
            weasyprint.HTML(string=html_content).write_pdf(output_path)
        except Exception as e:
            raise Exception(f"Ошибка при генерации PDF с weasyprint: {e}")
    else:
        raise Exception("Не установлены ни pdfkit, ни weasyprint. Установите один из них: pip install pdfkit (требует wkhtmltopdf) или pip install weasyprint")


def open_pdf(file_path: str) -> None:
    """Открывает PDF файл средствами ОС Windows."""
    try:
        if sys.platform == 'win32':
            os.startfile(file_path)
        elif sys.platform == 'darwin':
            subprocess.run(['open', file_path])
        else:
            subprocess.run(['xdg-open', file_path])
    except Exception as e:
        print(f"Не удалось открыть PDF файл: {e}")


def main():
    """Основная функция."""
    print("=" * 60)
    print("Генератор PDF из CSV и HTML шаблонов")
    print("=" * 60)
    
    # Определяем пути
    base_dir = Path(__file__).parent
    data_dir = base_dir / 'data'
    templates_dir = base_dir / 'templates'
    output_dir = base_dir / 'output'
    
    # Создаем директорию output, если её нет
    output_dir.mkdir(exist_ok=True)
    
    # Получаем списки файлов
    csv_files = get_files_in_directory(str(data_dir), '.csv')
    html_templates = get_files_in_directory(str(templates_dir), '.html')
    
    # Выводим списки
    display_file_list(csv_files, "Доступные CSV файлы")
    display_file_list(html_templates, "Доступные HTML шаблоны")
    
    # Проверяем наличие файлов
    if not csv_files:
        print("\nОшибка: Не найдено CSV файлов в директории /data")
        sys.exit(1)
    
    if not html_templates:
        print("\nОшибка: Не найдено HTML шаблонов в директории /templates")
        sys.exit(1)
    
    # Получаем выбор пользователя
    selected_csv = get_user_choice(csv_files, "CSV файл")
    selected_template = get_user_choice(html_templates, "HTML шаблон")
    
    if not selected_csv or not selected_template:
        print("\nОшибка: Не выбран файл или шаблон.")
        sys.exit(1)
    
    # Формируем полные пути
    csv_path = data_dir / selected_csv
    template_path = templates_dir / selected_template
    
    print(f"\nВыбрано:")
    print(f"  CSV файл: {selected_csv}")
    print(f"  Шаблон: {selected_template}")
    
    try:
        # Читаем данные
        print("\nЧтение данных из CSV...")
        data, contractor = read_csv_data(str(csv_path))
        print(f"Прочитано {len(data)} записей. Поставщик: {contractor}")
        
        # Читаем шаблон
        print("Чтение HTML шаблона...")
        template = read_template(str(template_path))
        
        # Генерируем HTML
        print("Генерация HTML...")
        html_content = generate_html(template, data, contractor)
        
        # Генерируем имя файла
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_filename = f"nakladnaya_{timestamp}.pdf"
        pdf_path = output_dir / pdf_filename
        
        # Генерируем PDF
        print("Генерация PDF...")
        if not PDFKIT_AVAILABLE and not WEASYPRINT_AVAILABLE:
            print("\nОШИБКА: Не установлены библиотеки для генерации PDF.")
            print("Установите одну из них:")
            print("  pip install pdfkit")
            print("  или")
            print("  pip install weasyprint")
            sys.exit(1)
        
        generate_pdf(html_content, str(pdf_path))
        print(f"PDF успешно создан: {pdf_path}")
        
        # Открываем PDF
        print("Открытие PDF файла...")
        open_pdf(str(pdf_path))
        
        print("\nГотово!")
        
    except Exception as e:
        print(f"\nОшибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
