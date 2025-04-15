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
from docx2pdf import convert  # <--- добавьте эту строку

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

root.geometry("1000x450")  # Установка фиксированного размера окна (ширина x высота)
root.resizable(False, False)  # Запрет изменения размера окна (по ширине, по высоте)

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
        status_var.set("✅ PDF-файлы Приложения успешно обработаны.")

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
    word_entry.delete(0, tk.END)  # Добавлено
    word_entry.insert(0, "Izvetaj_Отчет")  # Добавлено
    word_file_path[0] = None  # Добавлено
    word_file_label.config(text="Файл не выбран")  # Добавлено
    status_var.set("🔄 Поля сброшены по умолчанию.")

# === Word → PDF ===
word_file_path = [None]

def select_word_file():
    path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if path:
        word_file_path[0] = path
        word_file_label.config(text=os.path.basename(path))
        # Подставляем имя файла (без расширения) в поле
        base = os.path.splitext(os.path.basename(path))[0]
        word_entry.delete(0, tk.END)
        word_entry.insert(0, base)

def convert_word_to_pdf():
    if not word_file_path[0]:
        status_var.set("⚠ Сначала выберите Word-файл.")
        return
    # Получаем имя из поля, если оно не пустое, иначе из файла
    base_name = word_entry.get().strip() or os.path.splitext(os.path.basename(word_file_path[0]))[0]
    out_dir = os.path.dirname(word_file_path[0])
    out_pdf = os.path.join(out_dir, f"{base_name}.pdf")
    try:
        convert(word_file_path[0], out_pdf)
        if os.path.exists(out_pdf):
            status_var.set(f"✅ PDF создан: {os.path.basename(out_pdf)}")
        else:
            status_var.set("❌ Ошибка: PDF не был создан")
    except Exception as e:
        if "Word.Application.Quit" in str(e) and os.path.exists(out_pdf):
            status_var.set(f"✅ PDF создан: {os.path.basename(out_pdf)}")
        else:
            status_var.set(f"❌ Ошибка при конвертации: {str(e)}")

def create_merged_pdf():
    temp_files = []
    merged_writer = PdfWriter()
    # 1. Word → PDF (если выбран)
    word_pdf_path = None
    if word_file_path[0]:
        word_pdf_path = os.path.splitext(word_file_path[0])[0] + ".pdf"
        try:
            convert(word_file_path[0], word_pdf_path)
            if os.path.exists(word_pdf_path):
                temp_files.append(word_pdf_path)
            else:
                status_var.set("❌ Ошибка: PDF из Word не был создан")
                return
        except Exception as e:
            if "Word.Application.Quit" in str(e) and os.path.exists(word_pdf_path):
                temp_files.append(word_pdf_path)
            else:
                status_var.set(f"❌ Ошибка при конвертации Word: {str(e)}")
                return

    # 2. PDF-файлы с текстом
    pdf_temp_paths = []
    for i, path in enumerate(file_paths):
        if path:
            text = entries[i].get().strip()
            # Создаём временный файл с добавленным текстом
            temp_pdf = os.path.join(os.path.dirname(path), f"__temp_att_{i+1}.pdf")
            try:
                reader = PdfReader(path)
                writer = PdfWriter()
                for page in reader.pages:
                    overlay = create_overlay(text, float(page.mediabox.width), float(page.mediabox.height))
                    page.merge_page(overlay)
                    writer.add_page(page)
                with open(temp_pdf, "wb") as f:
                    writer.write(f)
                pdf_temp_paths.append(temp_pdf)
            except Exception as e:
                status_var.set(f"❌ Ошибка при обработке PDF: {os.path.basename(path)}\n{e}")
                # Удаляем временные файлы
                for f in pdf_temp_paths:
                    if os.path.exists(f):
                        os.remove(f)
                return
    temp_files.extend(pdf_temp_paths)

    if not temp_files:
        status_var.set("⚠ Не выбран ни один файл для объединения.")
        return

    # 3. Объединяем все PDF
    try:
        for pdf_path in temp_files:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                merged_writer.add_page(page)
        # Получаем имя из поля word_entry, если оно не пустое, иначе "merged"
        base_name = word_entry.get().strip() or "merged"
        merged_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"{base_name}_All.pdf",
            title="Сохранить объединённый PDF"
        )
        if merged_path:
            with open(merged_path, "wb") as f:
                merged_writer.write(f)
            status_var.set(f"✅ Общий PDF создан: {os.path.basename(merged_path)}")
        else:
            status_var.set("Операция отменена.")
    except Exception as e:
        status_var.set(f"❌ Ошибка при объединении: {str(e)}")
    finally:
        # Удаляем временные файлы
        for f in pdf_temp_paths:
            if os.path.exists(f):
                os.remove(f)

# === UI ===
# --- Блок для Word-файла ---
word_frame = tk.Frame(root, bg=BG_COLOR)
word_frame.pack(padx=20, pady=(15, 0), fill='x')

# Создаем вложенный фрейм для верхней строки
top_row = tk.Frame(word_frame, bg=BG_COLOR)
top_row.pack(fill='x')

word_entry = tk.Entry(top_row, width=40, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
word_entry.insert(0, "Izvetaj_Отчет")  # значение по умолчанию
word_entry.pack(side='left', padx=(0, 10))

word_btn = tk.Button(top_row, text="📄 Выбрать Word (.docx)", command=select_word_file, bg=BTN_COLOR, relief="flat")
word_btn.pack(side='left', padx=(0, 10))

word_convert_btn = tk.Button(top_row, text="➡️ Создать PDF из word", 
                          command=convert_word_to_pdf, bg=BTN_COLOR, relief="flat",
                          width=30)
word_convert_btn.pack(side='right', padx=20)

# Добавляем текст-подсказку под кнопкой
word_convert_note = tk.Label(word_frame, text="Создает PDF только из docx файла без приложений", 
                           anchor='e', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
word_convert_note.pack(side='right', padx=20, pady=(1, 0))

# Создаем отдельную строку для метки файла
word_file_label = tk.Label(word_frame, text="Файл не выбран", anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
word_file_label.pack(fill='x', padx=(0, 10), pady=(1, 0))

# После блока word_convert_btn добавьте:

# Добавить определение actions_frame перед использованием
actions_frame = tk.Frame(root, bg=BG_COLOR)
actions_frame.pack(padx=20, pady=5, fill='x')

# Разделительная линия
separator = tk.Frame(root, height=2, bg="#e0e0e0")
separator.pack(fill='x', padx=20, pady=(5, 10))

# --- Блок для PDF-файлов ---
apps_frame = tk.LabelFrame(root, text="Приложения", bg=BG_COLOR, fg="#222", font=("Segoe UI", 10, "bold"))
apps_frame.pack(padx=20, pady=(0, 6), fill='x')

# Создаем фрейм для кнопок сохранения справа
save_btn_frame = tk.Frame(apps_frame, bg=BG_COLOR)
save_btn_frame.pack(side='right', padx=10, pady=6)

# Кнопки сохранения
btn_style = {"width": 30, "bg": BTN_COLOR, "activebackground": "#d5d5d5", "relief": "flat"}
tk.Button(save_btn_frame, text="💾 Сохранить в тот же файл PDF", 
         command=lambda: process_pdfs(False), **btn_style).pack(pady=3)
tk.Button(save_btn_frame, text="📝 Сохранить с переименованием", 
         command=lambda: process_pdfs(True), **btn_style).pack(pady=3)

# Добавляем примечание под кнопками
note_text = (
    "💾 Сохранить в тот же файл PDF – заменяет оригинал PDF.\n"
    "📝 Сохранить с переименованием – создаёт копию pdf с 'att.X_...'\n"
    "\n"
    "Текст в приложении будет добавлен в правом верхнем углу по\n" 
    "короткой стороне страницы.\n"
    "Каждое прилжение сохранится отдельно.\n" 
)
note_label = tk.Label(save_btn_frame, text=note_text, justify='left', 
                     bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
note_label.pack(pady=(5, 0))

# Фрейм для полей ввода и выбора файлов
for i in range(4):
    frame = tk.Frame(apps_frame, bg=BG_COLOR)
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

# Разделительная линия
separator = tk.Frame(root, height=2, bg="#e0e0e0")
separator.pack(fill='x', padx=20, pady=(10, 5))

# Нижние кнопки
bottom_btn_frame = tk.Frame(root, bg=BG_COLOR)
bottom_btn_frame.pack(padx=20, pady=10, fill='x')

bottom_btn_style = {"bg": BTN_COLOR, "activebackground": "#d5d5d5", "relief": "flat", "width": 30}

# Создаем вертикальный фрейм для кнопок
buttons_frame = tk.Frame(bottom_btn_frame, bg=BG_COLOR)
buttons_frame.pack(side='left', padx=5)

# Размещаем кнопки вертикально
tk.Button(buttons_frame, text="🔄 Сброс/Вернуть по умолчанию", 
         command=reset_fields, **bottom_btn_style).pack(pady=(0,5))
tk.Button(buttons_frame, text="📚 Создать общий PDF", 
         command=create_merged_pdf, **bottom_btn_style).pack()

# Добавляем info_text после кнопок
info_text = (
    "📌 Пояснения:\n"
    "🔄 Вернуть по умолчанию – сбрасывает названия и очищает отмеченные файлы.\n"
    "📚 Создать общий PDF – создает общий PDF из word файла и Приложений.\n"
)
info_label = tk.Label(bottom_btn_frame, text=info_text, justify='left', anchor='nw', 
                     bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
info_label.pack(side='left', anchor='n', padx=(20, 0))

status_var = tk.StringVar()
status_label = tk.Label(root, textvariable=status_var, fg="green", anchor='w', relief="sunken", bd=1, bg="#f1f1f1", padx=5)
status_label.pack(fill='x', padx=20, pady=(5, 15))
status_var.set("Готов к работе")

def open_github(event=None):
    webbrowser.open_new("https://github.com/jackal100500/pdf-attachments-tool")

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