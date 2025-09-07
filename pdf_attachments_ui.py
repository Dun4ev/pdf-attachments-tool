import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter
from pypdf import Transformation
from reportlab.pdfgen import canvas
from io import BytesIO
from pathlib import Path
import platform
import webbrowser
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from docx2pdf import convert  # <--- добавьте эту строку
import shutil

# --- ИСТОРИЯ ПРОБЛЕМЫ И РЕШЕНИЕ ---
# Изначально проект использовал библиотеку PyPDF2. В ходе работы была обнаружена
# критическая проблема: при добавлении штампа на PDF-страницу, содержащийся
# в штампе текст становился невидимым в большинстве PDF-просмотрщиков.
#
# Диагностика с помощью библиотеки PyMuPDF показала, что в результирующем файле
# текст переставал существовать как текстовый объект и превращался в набор
# векторных кривых без заливки. Это происходило в момент слияния страниц
# (page.merge_transformed_page). Проблема не была связана с цветом, режимом
# рендеринга или используемым шрифтом.
#
# Коренной причиной оказалась устаревшая библиотека PyPDF2. В процессе отладки
# было замечено предупреждение о том, что PyPDF2 является deprecated и
# рекомендуется переход на ее официального преемника - pypdf.
#
# РЕШЕНИЕ:
# Проект был мигрирован с PyPDF2 на pypdf. Это включало:
# 1. Замену PyPDF2 на pypdf в requirements.txt.
# 2. Замену импортов "from PyPDF2" на "from pypdf".
#
# Это полностью решило проблему невидимости текста.
# ---

import sys
import subprocess # Added this line
import logging

# --- БЛОК ДЛЯ ОБРАБОТКИ ВЫВОДА В EXE ---
# Этот блок перенаправляет stdout/stderr в лог-файл, когда приложение
# скомпилировано с флагом --windowed (без консоли).
# Это предотвращает ошибки, когда вложенные библиотеки (например, docx2pdf)
# пытаются что-то напечатать.
# Код с ctypes для скрытия окна был удален, так как это делается
# с помощью флага --windowed при сборке PyInstaller.
if getattr(sys, 'frozen', False) and (sys.stdout is None or sys.stderr is None):
    exe_dir = os.path.dirname(sys.executable)
    log_path = os.path.join(exe_dir, 'pdf_attachments_ui.log')
    log_file = open(log_path, 'a', encoding='utf-8', buffering=1)
    sys.stdout = log_file
    sys.stderr = log_file

# Also initialize logging for both frozen and non-frozen runs
try:
    _base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.abspath('.')
    LOG_PATH = os.path.join(_base_dir, 'pdf_attachments_ui.log')
    logging.basicConfig(filename=LOG_PATH, level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', encoding='utf-8')
    logging.info('Logger initialized')
except Exception:
    LOG_PATH = None
# --- КОНЕЦ БЛОКА ---
        
# Helper function to find resources in PyInstaller bundle
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, "assets", relative_path)

# === СТИЛИ ===
BG_COLOR = "#f8f8fa"
BTN_COLOR = "#e0e0e0"
ENTRY_BG = "#ffffff"
ENTRY_FG = "#333333"
FONT = ("Segoe UI", 10)

root = tk.Tk()
root.title("PDF Приложения")
# Установка иконки
try:
    # Иконка теперь ищется в папке assets
    root.iconbitmap(resource_path("icon.ico"))
except tk.TclError:
    print("Не удалось загрузить иконку.") # Сообщение для отладки
root.configure(bg=BG_COLOR)
root.option_add("*Font", FONT)

# Авто-высота окна; ширину зафиксируем позже после построения UI
root.resizable(False, False)  # Запрет изменения размера окна (по ширине, по высоте)

entries = []
file_labels = []
file_paths = [None]*6

# Добавить в начало файла после других глобальных переменных
last_merged_pdf_path = [None]

# === РЕГИСТРАЦИЯ ШРИФТА ===
def register_font():
    font_name = "Arial"
    font_path = None
    if platform.system() == "Windows":
        win_font = Path("C:/Windows/Fonts/arial.ttf")
        if win_font.exists():
            font_path = str(win_font)
    if not font_path:
        # Используем resource_path для поиска шрифта в бандле
        local_font_path = resource_path("DejaVuSans.ttf")
        local_font = Path(local_font_path)
        if local_font.exists():
            font_name = "DejaVuSans"
            font_path = str(local_font)
        else:
            # Добавим вывод пути для отладки
            print(f"Не найден файл шрифта по пути: {local_font_path}")
            raise FileNotFoundError("Не найден ни системный Arial, ни локальный DejaVuSans.ttf")
    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name, font_path))
    return font_name

FONT_USED = register_font()

# === PDF ОБРАБОТКА ===
def create_overlay(text, width, height, rotation=0):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(width, height))
    can.setFont(FONT_USED, 12)
    
    margin = 40  # Отступ от краев
    
    can.saveState()

    # Логика для размещения текста в правом верхнем углу ВИДИМОЙ страницы.
    # 1. Перемещаем начало координат (translate) в целевой угол.
    # 2. Поворачиваем систему координат (rotate).
    # 3. Рисуем текст с выравниванием по правому краю в новой точке (0,0).

    if rotation == 90:
        # Видимый ВП угол -> сохраненный ВЛ (верхний левый) угол
        can.translate(margin, height - margin)
        can.rotate(90)
    elif rotation == 180:
        # Видимый ВП угол -> сохраненный НЛ (нижний левый) угол
        can.translate(margin, margin)
        can.rotate(180)
    elif rotation == 270:
        # Видимый ВП угол -> сохраненный НП (нижний правый) угол
        can.translate(width - margin, margin)
        can.rotate(270)
    elif width > height: # Альбомная страница без флага поворота (из Word)
        # Видимый ВП угол -> сохраненный ВП угол, но текст должен быть вертикальным
        can.translate(width - margin, height - margin)
        can.rotate(90)
    else: # Портретная страница без флага поворота
        # Видимый ВП угол -> сохраненный ВП угол
        can.translate(width - margin, height - margin)
        # Поворот не нужен
    
    # Рисуем текст с выравниванием по правому краю в точке (0,0) новой системы координат
    can.drawRightString(0, 0, text)
            
    can.restoreState()
    can.save()
    packet.seek(0)
    return PdfReader(packet).pages[0]

# --- Новая логика: штамп без прозрачности, учёт CropBox/Rotate ---
def _create_stamp_page(text: str, stamp_width: float = 200, stamp_height: float = 40):
    """
    Небольшая страница-штамп с текстом (Helvetica), без прозрачности.
    Текст выравнивается по правому краю в точке (stamp_width, 0) — это якорь.
    """
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(stamp_width, stamp_height))
    can.setFont("Helvetica", 12)
    can.drawRightString(stamp_width, 0, text)
    can.save()
    packet.seek(0)
    return PdfReader(packet).pages[0], stamp_width, stamp_height

# Redefine with advanced options (will override previous one at import time)
def _create_stamp_page(
    text: str,
    stamp_width: float = 240,
    stamp_height: float = 24,
    font_name: str = None,
    font_size: int = 12,
    draw_bg: bool = True,
    bg_padding: int = 3,
):
    from reportlab.pdfbase.pdfmetrics import stringWidth
    # Выбираем шрифт: используем зарегистрированный (Arial/DejaVuSans) при наличии
    if font_name is None:
        try:
            font_name = FONT_USED
        except Exception:
            font_name = "Helvetica"
    size = font_size
    while size >= 8:
        w = stringWidth(text, font_name, size)
        if w + 2 * bg_padding <= stamp_width:
            break
        size -= 1

    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(stamp_width, stamp_height))
    can.setFont(font_name, size)

    # Явно задаём «чёрный» в стандартном RGB и режим рисования "fill"
    try:
        # Используем RGB - это более универсально для экранного отображения
        can.setFillColorRGB(0, 0, 0)
    except Exception:
        # Fallback на CMYK, если что-то пойдет не так (маловероятно)
        can.setFillColorCMYK(0, 0, 0, 1)
    try:
        # Используем официальный API для установки режима рендеринга
        # 0 = Fill, 1 = Stroke, 2 = Fill then Stroke
        can.setTextRenderMode(0)
    except Exception:
        pass

    # Привязка к нижнему правому углу штампа -> якорь (stamp_width, 0)
    baseline_y = bg_padding
    can.drawRightString(stamp_width - bg_padding, baseline_y, text)

    can.save()
    packet.seek(0)
    return PdfReader(packet).pages[0], stamp_width, stamp_height


def _visible_box(page):
    box = getattr(page, "cropbox", None) or page.mediabox
    llx = float(box.left)
    lly = float(box.bottom)
    urx = float(box.right)
    ury = float(box.top)
    return llx, lly, urx, ury


def _anchor_and_angle(page, margin: float = 12.0):
    rotation = int(page.get("/Rotate", 0) or 0) % 360
    llx, lly, urx, ury = _visible_box(page)
    width = urx - llx
    height = ury - lly

    is_displayed_landscape = (rotation in (0, 180) and width > height) or \
                             (rotation in (90, 270) and height > width)

    deg = rotation
    if is_displayed_landscape:
        deg = (deg + 90) % 360

    if is_displayed_landscape:
        # --- ВЕРХНИЙ ЛЕВЫЙ угол для альбомной ориентации, со сдвигом ВНИЗ ---
        alignment = 'top-left'
        
        # Смещение на 10% от ВИЗУАЛЬНОЙ высоты страницы
        if rotation in (0, 180):
            visual_height = height
        else: # 90, 270
            visual_height = width
        offset = visual_height * 0.39

        if rotation == 0:
            ax = llx + margin
            ay = ury - margin - offset # Сдвиг вниз
        elif rotation == 90:
            # Визуально "вниз" - это сдвиг вправо в системе координат оригинала
            ax = llx + margin + offset
            ay = lly + margin
        elif rotation == 180:
            # Визуально "вниз" - это сдвиг вверх в системе координат оригинала
            ax = urx - margin
            ay = lly + margin + offset
        else:  # 270
            # Визуально "вниз" - это сдвиг влево в системе координат оригинала
            ax = urx - margin - offset
            ay = ury - margin
    else:
        # --- ВЕРХНИЙ ПРАВЫЙ угол для книжной ориентации (без изменений) ---
        alignment = 'top-right'
        if rotation == 0:
            ax = urx - margin
            ay = ury - margin
        elif rotation == 90:
            ax = llx + margin
            ay = ury - margin
        elif rotation == 180:
            ax = llx + margin
            ay = lly + margin
        else:  # 270
            ax = urx - margin
            ay = lly + margin
            
    return ax, ay, deg, alignment

def _merge_stamp(page, text: str, margin: float = 12.0):
    stamp, sw, sh = _create_stamp_page(text)
    ax, ay, deg, alignment = _anchor_and_angle(page, margin)
    
    import math
    rad = math.radians(deg)
    cos_d = math.cos(rad)
    sin_d = math.sin(rad)

    # Определяем, какой угол штампа будем выравнивать
    if alignment == 'top-left':
        # Верхний левый угол штампа в его координатах: (0, sh)
        x_stamp, y_stamp = 0, sh
    else: # 'top-right'
        # Верхний правый угол штампа: (sw, sh)
        x_stamp, y_stamp = sw, sh

    # Рассчитываем положение угла штампа после поворота
    x_rot = x_stamp * cos_d - y_stamp * sin_d
    y_rot = x_stamp * sin_d + y_stamp * cos_d
    
    # Вычисляем смещение, чтобы совместить повернутый угол с точкой (ax, ay)
    tx = ax - x_rot
    ty = ay - y_rot
    
    transform = Transformation().rotate(deg).translate(tx, ty)

    # Предпочитаем современный snake_case (верхний слой)
    if hasattr(page, "merge_transformed_page"):
        page.merge_transformed_page(stamp, transform)
        return

    # Fallback 1: применить трансформацию к штампу и слить
    try:
        stamp.add_transformation(transform)
        page.merge_page(stamp)
        return
    except Exception:
        pass

    # Fallback 2: старый camelCase (на крайний случай)
    if hasattr(page, "mergeTransformedPage"):
        m = transform.matrix
        ctm = (m[0][0], m[0][1], m[1][0], m[1][1], m[2][0], m[2][1])
        page.mergeTransformedPage(stamp, ctm)
        return

    # Последний резерв: просто слить (лучше так, чем упасть)
    page.merge_page(stamp)

def insert_text_to_pdf(pdf_path, text, save_as_new, prefix):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    # Новая реализация: небольшой штамп без прозрачности, точное позиционирование
    for page in reader.pages:
        # Ленивая загрузка вспомогательных функций, если добавлены ниже
        try:
            _merge_stamp_top_right
        except NameError:
            pass
        else:
            _merge_stamp(page, text, margin=12.0)
            writer.add_page(page)
            continue

        # Fallback на старый метод (если вспомогательные функции не определены)
        rotation = page.get('/Rotate', 0)
        overlay = create_overlay(
            text,
            float(page.mediabox.width),
            float(page.mediabox.height),
            rotation
        )
        page.merge_page(overlay)
        writer.add_page(page)
    output_path = pdf_path if not save_as_new else os.path.join(os.path.dirname(pdf_path), f"{prefix}_{os.path.basename(pdf_path)}")
    with open(output_path, "wb") as f:
        writer.write(f)

def insert_text_to_pdf_safe(pdf_path, text, save_as_new, prefix):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        _merge_stamp(page, text, margin=12.0)
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
    for i in range(6):
        path = file_paths[i]
        if path:
            try:
                text = entries[i].get().strip()
                prefix = f"att.{i+1}"
                insert_text_to_pdf_safe(path, text, save_as_new, prefix)
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
    for i in range(6):
        default_text = f"Prilog / Приложение 7.0{i+1}"
        entries[i].delete(0, tk.END)
        entries[i].insert(0, default_text)
        file_paths[i] = None
        file_labels[i].config(text="Файл не выбран")
    word_entry.delete(0, tk.END)
    word_entry.insert(0, "Izveštaj_Отчет")
    word_file_path[0] = None
    word_file_label.config(text="Файл не выбран")
    
    # Добавлено для сброса поля PDF-отчета
    pdf_report_entry.delete(0, tk.END)
    pdf_report_entry.insert(0, "Izveštaj_Отчет")
    pdf_report_file_path[0] = None
    pdf_report_label.config(text="Файл не выбран")

    if hasattr(root, 'pdf_link_label'):
        root.pdf_link_label.destroy()
    
    status_var.set("🔄 Поля сброшены по умолчанию.")

# === Word & PDF Отчеты ===
word_file_path = [None]
pdf_report_file_path = [None]

def select_word_file():
    path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if path:
        word_file_path[0] = path
        word_file_label.config(text=os.path.basename(path))
        base = os.path.splitext(os.path.basename(path))[0]
        word_entry.delete(0, tk.END)
        word_entry.insert(0, base)
        
        # Сброс поля PDF-отчета
        pdf_report_file_path[0] = None
        pdf_report_label.config(text="Файл не выбран")
        pdf_report_entry.delete(0, tk.END)
        pdf_report_entry.insert(0, "Izveštaj_Отчет")


def select_pdf_report_file():
    path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if path:
        pdf_report_file_path[0] = path
        pdf_report_label.config(text=os.path.basename(path))
        base = os.path.splitext(os.path.basename(path))[0]
        pdf_report_entry.delete(0, tk.END)
        pdf_report_entry.insert(0, base)

        # Сброс поля Word-отчета
        word_file_path[0] = None
        word_file_label.config(text="Файл не выбран")
        word_entry.delete(0, tk.END)
        word_entry.insert(0, "Izveštaj_Отчет")


def convert_word_to_pdf():
    if not word_file_path[0]:
        status_var.set("⚠ Сначала выберите Word-файл.")
        return
    
    base_name = word_entry.get().strip() or os.path.splitext(os.path.basename(word_file_path[0]))[0]
    out_dir = os.path.dirname(word_file_path[0])
    out_pdf = os.path.join(out_dir, f"{base_name}.pdf")
    
    try:
        convert(word_file_path[0], out_pdf)
        if os.path.exists(out_pdf):
            status_var.set(f"✅ PDF создан: {os.path.basename(out_pdf)}")
            create_pdf_link(out_pdf)
        else:
            status_var.set("❌ Ошибка: PDF не был создан")
    except Exception as e:
        if "Word.Application.Quit" in str(e) and os.path.exists(out_pdf):
            status_var.set(f"✅ PDF создан: {os.path.basename(out_pdf)}")
            create_pdf_link(out_pdf)
        else:
            status_var.set(f"❌ Ошибка при конвертации: {str(e)}")

def create_merged_pdf():
    temp_files = []
    merged_writer = PdfWriter()
    
    main_report_path = None
    base_name_for_save = "merged"

    # 1. Определяем основной отчет (PDF или Word)
    if pdf_report_file_path[0]:
        # Используем PDF-отчет
        main_report_path = pdf_report_file_path[0]
        text = pdf_report_entry.get().strip()
        base_name_for_save = text or os.path.splitext(os.path.basename(main_report_path))[0]
        
        # Копируем PDF-отчет во временный файл без добавления текста
        temp_main_pdf = os.path.join(os.path.dirname(main_report_path), f"__temp_main_{os.path.basename(main_report_path)}")
        try:
            shutil.copy(main_report_path, temp_main_pdf)
            temp_files.append(temp_main_pdf)
        except Exception as e:
            status_var.set(f"❌ Ошибка при копировании PDF-отчета: {e}")
            return

    elif word_file_path[0]:
        # Используем Word-отчет
        main_report_path = word_file_path[0]
        base_name_for_save = word_entry.get().strip() or os.path.splitext(os.path.basename(main_report_path))[0]
        word_pdf_path = os.path.join(os.path.dirname(main_report_path), f"{base_name_for_save}.pdf")
        
        try:
            convert(main_report_path, word_pdf_path)
            if os.path.exists(word_pdf_path):
                temp_files.append(word_pdf_path) # Этот файл временный только для этой операции
            else:
                status_var.set("❌ Ошибка: PDF из Word не был создан")
                return
        except Exception as e:
            if "Word.Application.Quit" in str(e) and os.path.exists(word_pdf_path):
                temp_files.append(word_pdf_path)
            else:
                status_var.set(f"❌ Ошибка при конвертации Word: {e}")
                return

    # 2. PDF-файлы приложений с текстом
    pdf_temp_paths = []
    for i, path in enumerate(file_paths):
        if path:
            text = entries[i].get().strip()
            temp_pdf = os.path.join(os.path.dirname(path), f"__temp_att_{i+1}.pdf")
            try:
                reader = PdfReader(path)
                writer = PdfWriter()
                for page in reader.pages:
                    # Используем штамп без прозрачности с учётом CropBox/Rotate
                    _merge_stamp(page, text, margin=12.0)
                    writer.add_page(page)
                with open(temp_pdf, "wb") as f:
                    writer.write(f)
                pdf_temp_paths.append(temp_pdf)
            except Exception as e:
                status_var.set(f"❌ Ошибка при обработке PDF: {os.path.basename(path)}\n{e}")
                for f in temp_files + pdf_temp_paths: # Очистка всех временных файлов
                    if os.path.exists(f): os.remove(f)
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
        
        merged_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"{base_name_for_save}_All.pdf",
            title="Сохранить объединённый PDF"
        )
        if merged_path:
            with open(merged_path, "wb") as f:
                merged_writer.write(f)
            last_merged_pdf_path[0] = merged_path
            status_var.set(f"✅ Общий PDF создан: {os.path.basename(merged_path)}")
            create_pdf_link(merged_path)
        else:
            status_var.set("Операция отменена.")
    except Exception as e:
        status_var.set(f"❌ Ошибка при объединении: {e}")
    finally:
        # 4. Удаляем все временные файлы
        for f in temp_files:
            if os.path.exists(f):
                os.remove(f)


def create_pdf_link(pdf_path):
    """Создает кликабельную ссылку на созданный PDF и ссылку на папку"""
    # Удаляем старые ссылки если они существуют
    if hasattr(root, 'links_frame'): # Check for the new frame
        root.links_frame.destroy()

    # Create a new frame to hold both links
    root.links_frame = tk.Frame(root, bg=BG_COLOR)
    root.links_frame.pack(before=status_label, pady=(0, 5)) # Pack the frame above status_label

    def open_pdf():
        os.startfile(pdf_path)

    # Create PDF file link
    filename = os.path.basename(pdf_path)
    pdf_link_label = tk.Label( # No longer root.pdf_link_label
        root.links_frame, # Pack into the new frame
        text=f"📎 Открыть {filename}",
        fg="#0066cc",
        cursor="hand2",
        bg=BG_COLOR,
        font=("Segoe UI", 9, "underline")
    )
    pdf_link_label.bind("<Button-1>", lambda e: open_pdf())
    pdf_link_label.pack(side='left', padx=(0, 10)) # Pack left in the new frame

    # --- New: Folder link ---
    folder_path = os.path.dirname(pdf_path)

    def open_folder(event=None):
        if folder_path and os.path.isdir(folder_path):
            try:
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                elif platform.system() == "Darwin": # macOS
                    subprocess.run(['open', folder_path])
                else: # Linux
                    subprocess.run(['xdg-open', folder_path])
            except Exception as e:
                messagebox.showwarning("Ошибка", f"Не удалось открыть папку: {e}")

    folder_link_label = tk.Label(
        root.links_frame, # Pack into the new frame
        text="📁 Открыть папку",
        fg="#0066cc",
        cursor="hand2",
        bg=BG_COLOR,
        font=("Segoe UI", 9, "underline")
    )
    folder_link_label.bind("<Button-1>", open_folder)
    folder_link_label.pack(side='right', padx=(10, 0)) # Pack right in the new frame, add padding to the left

# === UI ===
# --- Блок для Word-файла ---
word_frame = tk.Frame(root, bg=BG_COLOR)
word_frame.pack(padx=20, pady=(15, 0), fill='x')

# Создаем вложенный фрейм для верхней строки
top_row = tk.Frame(word_frame, bg=BG_COLOR)
top_row.pack(fill='x')

word_entry = tk.Entry(top_row, width=40, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
word_entry.insert(0, "Izveštaj_Отчет")  # значение по умолчанию
word_entry.pack(side='left', padx=(0, 10))

word_btn = tk.Button(top_row, text="📄 Выбрать Word (.docx)", command=select_word_file, bg=BTN_COLOR, relief="flat")
word_btn.pack(side='left', padx=(0, 10))

word_convert_btn = tk.Button(top_row, text="➡️ Создать PDF из word", 
                          command=convert_word_to_pdf, bg=BTN_COLOR, relief="flat",
                          width=30)
word_convert_btn.pack(side='right', padx=20)

# Добавляем текст-подсказку под кнопкой
word_convert_note = tk.Label(word_frame, text="Создает PDF из docx файла без приложений", 
                           anchor='e', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
word_convert_note.pack(side='right', padx=20, pady=(1, 0))

# Создаем отдельную строку для метки файла
word_file_label = tk.Label(word_frame, text="Файл не выбран", anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
word_file_label.pack(fill='x', padx=(0, 10), pady=(1, 0))

# --- Блок для PDF-отчета ---
pdf_report_frame = tk.Frame(root, bg=BG_COLOR)
pdf_report_frame.pack(padx=20, pady=(10, 0), fill='x')

pdf_report_top_row = tk.Frame(pdf_report_frame, bg=BG_COLOR)
pdf_report_top_row.pack(fill='x')

pdf_report_entry = tk.Entry(pdf_report_top_row, width=40, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
pdf_report_entry.insert(0, "Izveštaj_Отчет")
pdf_report_entry.pack(side='left', padx=(0, 10))

pdf_report_btn = tk.Button(pdf_report_top_row, text="📄 Выбрать PDF (.pdf)", command=select_pdf_report_file, bg=BTN_COLOR, relief="flat")
pdf_report_btn.pack(side='left', padx=(0, 10))

pdf_report_label = tk.Label(pdf_report_frame, text="Файл не выбран", anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
pdf_report_label.pack(fill='x', padx=(0, 10), pady=(1, 0))


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
    "Текст ""(Prilog / Приложение 7.0i)"" в PDF файле будет добавлен в правом верхнем углу покороткой стороне страницы.\n"
    "Каждое прилжение сохранится отдельно.\n"
)
note_label = tk.Label(save_btn_frame, text=note_text, justify='left', wraplength=220, 
                     bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
note_label.pack(pady=(5, 0))

# Фрейм для полей ввода и выбора файлов
for i in range(6):
    frame = tk.Frame(apps_frame, bg=BG_COLOR)
    frame.pack(padx=20, pady=6, fill='x')
    entry = tk.Entry(frame, width=35, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
    entry.insert(0, f"Prilog / Приложение 7.0{i+1}")
    entry.pack(side='left', padx=(0, 10))
    entries.append(entry)
    btn = tk.Button(frame, text="📂 Выбрать PDF", command=lambda idx=i: select_file(idx), bg=BTN_COLOR, relief="flat")
    btn.pack(side='left', padx=(0, 10))
    label = tk.Label(frame, text="Файл не выбран", width=45, anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
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
merge_btn = tk.Button(buttons_frame, 
                     text="📚 Создать общий PDF",
                     command=create_merged_pdf,
                     width=30,           # из bottom_btn_style
                     relief="flat",      # из bottom_btn_style
                     bg="#4CAF50",       # новый цвет фона
                     fg="white",         # цвет текста
                     activebackground="#45a049")
merge_btn.pack()

# Добавляем info_text после кнопок
info_text = (
    "🔄 Вернуть по умолчанию – сбрасывает названия и очищает отмеченные файлы.\n"
    "📚 Создать общий PDF – создает общий PDF из word файла и Приложений.\n"
)
info_label = tk.Label(bottom_btn_frame, text=info_text, justify='left', anchor='nw', 
                     bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
info_label.pack(side='left', anchor='n', padx=(20, 0))

status_var = tk.StringVar()
# Многострочный статус с переносом слов
status_label = tk.Label(
    root,
    textvariable=status_var,
    fg="green",
    anchor='w',
    justify='left',
    wraplength=920,
    relief="sunken",
    bd=1,
    bg="#f1f1f1",
    padx=5
)
status_label.pack(fill='x', padx=20, pady=(5, 15))
def show_status_details(event=None):
    top = tk.Toplevel(root)
    top.title("Детали сообщения")
    top.geometry("900x400")
    txt = tk.Text(top, wrap='word')
    txt.insert('1.0', status_var.get())
    txt.configure(state='disabled')
    txt.pack(fill='both', expand=True)
status_label.bind('<Double-Button-1>', show_status_details)
status_var.set("Готов к работе")

def open_github(event=None):
    webbrowser.open_new("https://github.com/Dun4ev/pdf-attachments-tool")

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

# После построения всего UI выставим авто-высоту и фиксированную ширину
root.update_idletasks()
desired_width = 1000
current_req_height = root.winfo_reqheight()
root.geometry(f"{desired_width}x{current_req_height}")

root.mainloop()
