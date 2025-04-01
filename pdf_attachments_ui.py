import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from io import BytesIO
from pathlib import Path
import platform
import webbrowser
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# === СТИЛИ ===
BG_COLOR = "#f8f8fa"
BTN_COLOR = "#e0e0e0"
ENTRY_BG = "#ffffff"
ENTRY_FG = "#333333"
FONT = ("Segoe UI", 10)

root = tk.Tk()
root.title("PDF Приложения")
root.configure(bg=BG_COLOR)
root.option_add("*Font", FONT)

entries = []
file_labels = []
file_paths = [None]*4

# === РЕГИСТРАЦИЯ ШРИФТА ===
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
            raise FileNotFoundError("Не найден ни системный Arial, ни локальный DejaVuSans.ttf")
    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name, font_path))
    return font_name

FONT_USED = register_font()

# === PDF ОБРАБОТКА ===
def create_overlay(text, width, height):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(width, height))
    if width <= height:
        can.setFont(FONT_USED, 12)
        can.drawRightString(width - 40, height - 30, text)
    else:
        can.translate(width, 0)
        can.rotate(90)
        can.setFont(FONT_USED, 12)
        can.drawRightString(height - 40, -30, text)
    can.save()
    packet.seek(0)
    return PdfReader(packet).pages[0]

def insert_text_to_pdf(pdf_path, text, save_as_new, prefix):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        overlay = create_overlay(text, float(page.mediabox.width), float(page.mediabox.height))
        page.merge_page(overlay)
        writer.add_page(page)
    output_path = pdf_path if not save_as_new else os.path.join(os.path.dirname(pdf_path), f"{prefix}_{os.path.basename(pdf_path)}")
    with open(output_path, "wb") as f:
        writer.write(f)

# === ЛОГИКА ===
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

# === UI ===
for i in range(4):
    frame = tk.Frame(root, bg=BG_COLOR)
    frame.pack(padx=20, pady=6, fill='x')
    entry = tk.Entry(frame, width=35, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
    entry.insert(0, f"Prilog 6.0{i+1} / Приложение 6.0{i+1}")
    entry.pack(side='left', padx=(0, 10))
    entries.append(entry)
    btn = tk.Button(frame, text="📂 Выбрать PDF", command=lambda idx=i: select_file(idx), bg=BTN_COLOR, relief="flat")
    btn.pack(side='left', padx=(0, 10))
    label = tk.Label(frame, text="Файл не выбран", width=30, anchor='w', bg=BG_COLOR, fg="#555")
    label.pack(side='left')
    file_labels.append(label)

actions_frame = tk.Frame(root, bg=BG_COLOR)
actions_frame.pack(padx=20, pady=10, fill='x')

btn_frame = tk.Frame(actions_frame, bg=BG_COLOR)
btn_frame.pack(side='left', padx=(0, 20))

# Кнопки
btn_style = {"width": 30, "bg": BTN_COLOR, "activebackground": "#d5d5d5", "relief": "flat"}
tk.Button(btn_frame, text="🔄 Вернуть по умолчанию", command=reset_fields, **btn_style).pack(pady=3)
tk.Button(btn_frame, text="💾 Сохранить в тот же файл", command=lambda: process_pdfs(False), **btn_style).pack(pady=3)
tk.Button(btn_frame, text="📝 Сохранить с переименованием", command=lambda: process_pdfs(True), **btn_style).pack(pady=3)

info_text = (
    "📌 Пояснения:\n"
    "🔄 Вернуть по умолчанию – сбрасывает названия и очищает файлы.\n"
    "💾 Сохранить в тот же файл – заменяет оригинал PDF.\n"
    "📝 Сохранить с переименованием – создаёт копию pdf с 'att.X_...'.\n"
    "Текст будет добавлен в правом верхнем углу по\n" 
    "короткой стороне.\n"
)
info_label = tk.Label(actions_frame, text=info_text, justify='left', anchor='nw', bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
info_label.pack(side='left', anchor='n')

status_var = tk.StringVar()
status_label = tk.Label(root, textvariable=status_var, fg="green", anchor='w', relief="sunken", bd=1, bg="#f1f1f1", padx=5)
status_label.pack(fill='x', padx=20, pady=(5, 15))
status_var.set("Готов к работе")

def open_github(event=None):
    webbrowser.open_new("https://github.com/ТВОЙ_АДРЕС")

link_label = tk.Label(
    root,
    text="🔗GitHub",
    fg="blue",
    cursor="hand2",
    bg=BG_COLOR,
    font=("Segoe UI", 7, "underline")
)
link_label.bind("<Button-1>", open_github)

# 👇 размещаем в правом нижнем углу
link_label.place(relx=1.0, rely=1.0, anchor="se", x=-20, y=-10)

root.mainloop()
