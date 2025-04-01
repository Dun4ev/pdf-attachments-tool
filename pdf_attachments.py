import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from io import BytesIO
from pathlib import Path
import platform
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def register_font():
    font_name = "Arial"
    font_path = None

    if platform.system() == "Windows":
        win_font = Path("C:/Windows/Fonts/arial.ttf")
        if win_font.exists():
            font_path = str(win_font)

    if not font_path:
        local_font = Path("DejaVuSans.ttf")
        if local_font.exists():
            font_name = "DejaVuSans"
            font_path = str(local_font)
        else:
            raise FileNotFoundError(
                "Не найден ни системный Arial, ни локальный DejaVuSans.ttf"
            )

    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name, font_path))

    return font_name

FONT_USED = register_font()

# Генерация страницы с надписью
def create_overlay(text, width, height):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(width, height))

    # Портретная ориентация — ничего не меняем
    if width <= height:
        can.setFont(FONT_USED, 12)
        can.drawRightString(width - 40, height - 30, text)
    else:
        # Альбомная — поворачиваем холст на 90° и сдвигаем его
        can.translate(width, 0)
        can.rotate(90)
        can.setFont(FONT_USED, 12)
        # width и height теперь поменялись местами, поэтому drawRightString по height
        can.drawRightString(height - 40, -30, text)

    can.save()
    packet.seek(0)
    return PdfReader(packet).pages[0]


# Вставка текста в PDF
def insert_text_to_pdf(pdf_path, text, save_as_new, prefix):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for page in reader.pages:
        overlay = create_overlay(
            text,
            float(page.mediabox.width),
            float(page.mediabox.height)
        )
        page.merge_page(overlay)
        writer.add_page(page)

    output_path = (
        pdf_path if not save_as_new else os.path.join(
            os.path.dirname(pdf_path),
            f"{prefix}_{os.path.basename(pdf_path)}"
        )
    )

    with open(output_path, "wb") as f:
        writer.write(f)

        
# Обработчик кнопки
def process_pdfs(save_as_new):
    if not any(file_paths):
        status_var.set("⚠ Не выбрано ни одного PDF-файла.")
        return

    any_error = False
    for i in range(4):
        path = file_paths[i]
        if path:
            try:
                text = entries[i].get().strip()
                prefix = f"att.{i+1}"
                insert_text_to_pdf(path, text, save_as_new, prefix)
            except Exception as e:
                any_error = True
                messagebox.showerror("Ошибка", f"Ошибка при обработке {path}:\n{e}")
                status_var.set(f"❌ Ошибка при обработке: {os.path.basename(path)}")
    if not any_error:
        status_var.set("✅ PDF-файлы успешно обработаны.")

# Выбор файла
def select_file(index):
    path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if path:
        file_paths[index] = path
        file_labels[index].config(text=os.path.basename(path))

def reset_fields():
    for i in range(4):
        default_text = f"Prilog 6.0{i+1} / Приложение 6.0{i+1}"
        entries[i].delete(0, tk.END)
        entries[i].insert(0, default_text)

        file_paths[i] = None
        file_labels[i].config(text="Файл не выбран")

    status_var.set("🔄 Поля сброшены по умолчанию.")
    
    
# GUI
root = tk.Tk()
root.title("PDF Приложения")
entries = []
file_labels = []
file_paths = [None]*4

for i in range(4):
    frame = tk.Frame(root)
    frame.pack(padx=10, pady=5, fill='x')
    
    entry = tk.Entry(frame, width=35)
    entry.insert(0, f"Prilog 6.0{i+1} / Приложение 6.0{i+1}")
    entry.pack(side='left')
    entries.append(entry)

    btn = tk.Button(frame, text="Выбрать PDF", command=lambda idx=i: select_file(idx))
    btn.pack(side='left', padx=5)

    label = tk.Label(frame, text="Файл не выбран", width=30, anchor='w')
    label.pack(side='left')
    file_labels.append(label)

# Контейнер для кнопок и пояснений
actions_frame = tk.Frame(root)
actions_frame.pack(padx=10, pady=10, fill='x')

# Левая колонка: кнопки
btn_frame = tk.Frame(actions_frame)
btn_frame.pack(side='left', padx=(0, 20))

tk.Button(btn_frame, text="🔄 Вернуть по умолчанию", width=30,
          command=reset_fields).pack(pady=3)
tk.Button(btn_frame, text="💾 Сохранить в тот же файл", width=30,
          command=lambda: process_pdfs(save_as_new=False)).pack(pady=3)
tk.Button(btn_frame, text="📝 Сохранить с переименованием", width=30,
          command=lambda: process_pdfs(save_as_new=True)).pack(pady=3)

# Правая колонка: пояснение
info_text = (
    "📌 Пояснения:\n"
    "🔄 Вернуть по умолчанию – сбрасывает названия и очищает файлы.\n"
    "💾 Сохранить в тот же файл – заменяет оригинал PDF.\n"
    "📝 Сохранить с переименованием – сохраняет копию с префиксом 'att.X_...'."
)

info_label = tk.Label(actions_frame, text=info_text, justify='left', anchor='nw')
info_label.pack(side='left')

# Статус-бар
status_var = tk.StringVar()
status_label = tk.Label(root, textvariable=status_var, fg="green", anchor='w', relief="sunken", bd=1)
status_label.pack(fill='x', padx=10, pady=(5, 10))
status_var.set("Готов к работе")

root.mainloop()