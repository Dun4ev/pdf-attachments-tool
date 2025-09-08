from pypdf import PdfReader
import os
import sys

# --- Исправление для вывода кириллицы в консоль Windows ---
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# Абсолютный путь к файлу
pdf_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "CT-DR-B-CS-AIR-II.24.1-00-1M-20250827-00_All_6.pdf"))
pdf_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "att.1_P 192-22.pdf"))

try:
    reader = PdfReader(pdf_path)
    print(f"Анализ файла: {os.path.basename(pdf_path)}")
    print("-" * 40)
    for i, page in enumerate(reader.pages):
        width = page.mediabox.width
        height = page.mediabox.height
        rotation = page.get('/Rotate', 0) or 0
        
        if width > height:
            orientation = "Альбомная (Landscape)"
        elif height > width:
            orientation = "Книжная (Portrait)"
        else:
            orientation = "Квадратная (Square)"

        print(f"Страница {i+1}:")
        print(f"  - Ширина: {width:.2f} пунктов")
        print(f"  - Высота: {height:.2f} пунктов")
        print(f"  - Ориентация: {orientation}")
        print(f"  - Поворот (в метаданных): {rotation}°")
        print("-" * 40)

except FileNotFoundError:
    print(f"Ошибка: Файл не найден по пути {pdf_path}")
except Exception as e:
    print(f"Произошла ошибка при чтении PDF: {e}")
