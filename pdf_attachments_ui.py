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
import logging

# Закладки/оглавление: извлечение и запись через PyMuPDF, если доступен
try:
    from pdf_attachments.bookmarks import (
        extract_toc,
        compose_two_level_toc,
        compose_multi_attachment_toc,
        apply_toc,
        SourceToc,
    )
except Exception as _e:  # безопасный фолбэк
    extract_toc = None  # type: ignore
    compose_two_level_toc = None  # type: ignore
    compose_multi_attachment_toc = None  # type: ignore
    apply_toc = None  # type: ignore
    SourceToc = None  # type: ignore
    logging.getLogger(__name__).warning("Модуль pdf_attachments.bookmarks недоступен: %s", _e)

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
if getattr(sys, "frozen", False) and (sys.stdout is None or sys.stderr is None):
    exe_dir = os.path.dirname(sys.executable)
    log_path = os.path.join(exe_dir, "pdf_attachments_ui.log")
    log_file = open(log_path, "a", encoding="utf-8", buffering=1)
    sys.stdout = log_file
    sys.stderr = log_file

# Also initialize logging for both frozen and non-frozen runs
try:
    _base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.abspath(".")
    LOG_PATH = os.path.join(_base_dir, "pdf_attachments_ui.log")
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
folder_for_merge_path = [None] # <-- НОВАЯ ПЕРЕМЕННАЯ

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

# === НОВЫЕ ФУНКЦИИ ДЛЯ РЕЖИМА "ИЗ ПАПКИ" ===

def clear_individual_attachments():
    """Очищает все 6 слотов для ручного выбора приложений."""
    for i in range(6):
        default_text = f"Prilog / Приложение 7.0{i+1}"
        entries[i].delete(0, tk.END)
        entries[i].insert(0, default_text)
        file_paths[i] = None
        file_labels[i].config(text="Файл не выбран")

def clear_folder_for_merge():
    """Очищает выбор папки для слияния."""
    folder_for_merge_path[0] = None
    if 'folder_for_merge_label' in globals():
        folder_for_merge_label.config(text="Папка не выбрана")

def select_folder_for_merge():
    """Выбирает папку для слияния, очищая при этом ручные слоты."""
    path = filedialog.askdirectory(title="Выберите папку с PDF-приложениями")
    if path:
        folder_for_merge_path[0] = path
        if 'folder_for_merge_label' in globals():
            folder_for_merge_label.config(text=os.path.basename(path))
        clear_individual_attachments()
        status_var.set(f"🗂️ Выбрана папка для слияния: {os.path.basename(path)}")


# === PDF ОБРАБОТКА ===

# --- Новая логика: штамп без прозрачности, учёт CropBox/Rotate ---
# Redefine with advanced options (will override previous one at import time)
def _create_stamp_page(
    text: str,
    stamp_width: float = 240,
    stamp_height: float = 24,
    font_name: str = None,
    font_size: int = 11,
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
        # --- ВЕРХНИЙ ЛЕВЫЙ угол для альбомной ориентации (простая логика) ---
        alignment = 'top-left'
        # Целимся в визуальный верхний левый угол с отступом
        if rotation == 0:    # Визуальный ВЛ -> оригинальный ВЛ
            ax = llx + margin
            ay = ury - margin
        elif rotation == 90:   # Визуальный ВЛ -> оригинальный НЛ
            ax = llx + margin
            ay = lly + margin
        elif rotation == 180:  # Визуальный ВЛ -> оригинальный НП
            ax = urx - margin
            ay = lly + margin
        else:  # 270         # Визуальный ВЛ для 270°: берём левый край и видимую высоту
            # Для страниц типа 612x792 pt @270° (Letter, ландшафт от поворота)
            # используем левую границу по X, чтобы не уезжать вправо за предел
            visible_h = min(width, height)
            ax = lly + margin
            ay = llx + visible_h - margin
    else:
        # --- Книжная ориентация ---
        # Для повернутой на 270 градусов страницы, видимый верхний правый угол
        # на самом деле является физическим нижним правым углом.
        # Чтобы штамп не улетел, мы должны выравнивать его по его нижнему краю.
        if rotation == 270:
            alignment = 'bottom-right'
            ax, ay = urx - margin, lly + margin
        else:
            # Стандартные случаи для книжной ориентации
            alignment = 'top-right'
            if rotation == 0:
                ax, ay = urx - margin, ury - margin
            elif rotation == 90:
                ax, ay = llx + margin, ury - margin
            elif rotation == 180:
                ax, ay = llx + margin, lly + margin
            
    # Спец-фикс: для страниц с поворотом 270° и отображением в альбомной ориентации
    # вычисляем якорь по визуальной системе координат (верхний левый) и маппим в оригинальные координаты.
    try:
        if is_displayed_landscape and rotation == 270:
            visual_w = height  # при 270 визуальная ширина = высота бокса
            visual_h = width   # при 270 визуальная высота = ширина бокса
            dx = margin
            dy = visual_h - margin
            ax, ay = urx - dy, lly + dx
    except Exception:
        pass

    # Generalized override for pages with Rotate=270 and portrait base geometry (w < h):
    # Place stamp at visual top-right to avoid it going out on the right.
    try:
        llx2, lly2, urx2, ury2 = _visible_box(page)
        w2 = urx2 - llx2
        h2 = ury2 - lly2
        rot2 = int(page.get('/Rotate', 0) or 0) % 360
        if rot2 == 270 and (w2 < h2):
            alignment = 'top-right'
            visible_h = min(w2, h2)
            visible_w270=max(w2,h2)
            ax = urx2 - margin
            ay = lly2 + visible_w270 - margin

    except Exception:
        pass

    return ax, ay, deg, alignment

def _merge_stamp(page, text: str, margin: float = 12.0):
    # --- Новая логика: Адаптивный размер штампа для маленьких страниц ---
    
    # 1. Исходные параметры для штампа и шрифта
    initial_sw = 240.0
    initial_sh = 24.0
    initial_font_size = 11.0 # Желаемый размер шрифта
    
    # 2. Получаем размеры страницы
    box = getattr(page, "cropbox", None) or page.mediabox
    # Use visual width for scaling: if page rotated 90/270, swap width/height
    rotation = int(page.get('/Rotate', 0) or 0) % 360
    page_width = float(box.height) if rotation in (90, 270) else float(box.width)
    
    # 3. Проверка, не слишком ли велик штамп для этой страницы
    # Порог: если штамп занимает > 60% ширины, он считается слишком большим
    threshold = 0.60
    if initial_sw > page_width * threshold:
        # 4. Да, штамп слишком большой. Вычисляем коэффициент уменьшения.
        scale = (page_width * threshold) / initial_sw
        final_sw = initial_sw * scale
        final_sh = initial_sh * scale
        final_font_size = initial_font_size * scale
    else:
        # 5. Нет, страница достаточно большая. Используем стандартные размеры.
        final_sw = initial_sw
        final_sh = initial_sh
        final_font_size = initial_font_size

    # 6. Создаем штамп с финальными (возможно, уменьшенными) размерами
    stamp, sw, sh = _create_stamp_page(
        text,
        stamp_width=final_sw,
        stamp_height=final_sh,
        font_size=int(round(final_font_size)) # Округляем до целого
    )

    # 7. Вычисляем положение и поворот
    ax, ay, deg, alignment = _anchor_and_angle(page, margin)
    
    import math
    rad = math.radians(deg)
    cos_d = math.cos(rad)
    sin_d = math.sin(rad)

    # --- Новая, корректная логика вычисления трансформации ---
    # 1. Вычисляем координаты углов штампа после поворота
    c0 = (0, 0)
    c1 = (sw * cos_d, sw * sin_d)
    c2 = (-sh * sin_d, sh * cos_d)
    c3 = (sw * cos_d - sh * sin_d, sw * sin_d + sh * cos_d)

    # 2. Находим границы описанной рамки (bounding box) для повернутого штампа
    x_coords = [c0[0], c1[0], c2[0], c3[0]]
    y_coords = [c0[1], c1[1], c2[1], c3[1]]
    bbox_x_min, bbox_x_max = min(x_coords), max(x_coords)
    bbox_y_min, bbox_y_max = min(y_coords), max(y_coords)

    # 3. Определяем "ручку" на этой рамке в зависимости от выравнивания
    # Map desired page-corner alignment to the corresponding rotated-stamp corner
    if alignment == 'top-left':
        handle_x, handle_y = bbox_x_min, bbox_y_max
    elif alignment == 'bottom-right':
        handle_x, handle_y = bbox_x_max, bbox_y_min
    else:  # 'top-right'
        handle_x, handle_y = bbox_x_max, bbox_y_max

    # 4. Вычисляем смещение, чтобы переместить "ручку" в целевую точку (ax, ay)
    tx = ax - handle_x
    ty = ay - handle_y
    
    # 5. Применяем трансформацию: сначала поворот, потом смещение
    transform = Transformation().rotate(deg).translate(tx=tx, ty=ty)

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

def insert_text_to_pdf_safe(pdf_path, text, save_as_new, prefix):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        _merge_stamp(page, text, margin=12.0)
        writer.add_page(page)
    output_path = pdf_path if not save_as_new else os.path.join(os.path.dirname(pdf_path), f"{prefix}_{os.path.basename(pdf_path)}")
    with open(output_path, "wb") as f:
        writer.write(f)

# --- Override legacy API to ensure safe stamping only ---
def insert_text_to_pdf(pdf_path, text, save_as_new, prefix):
    return insert_text_to_pdf_safe(pdf_path, text, save_as_new, prefix)

# === Объединение с закладками ===
def merge_pdfs_with_bookmarks(parts: list[tuple[str, str]], output_path: str) -> None:
    """Объединяет PDF и создаёт верхнеуровневые закладки.

    Args:
        parts: Список (path, title). Для каждого PDF добавляется
            верхнеуровневая закладка `title`; существующие закладки
            исходников переносятся (если есть).
        output_path: Путь для сохранения результирующего PDF.

    Raises:
        FileNotFoundError: Если входного PDF нет на диске.
        Exception: Прочие ошибки чтения/записи PDF.
    """
    writer = PdfWriter()
    for pdf_path, title in parts:
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        # outline_item — верхнеуровневая закладка; import_outline — перенос исходных
        writer.append(pdf_path, outline_item=(title or None), import_outline=True)
    with open(output_path, "wb") as out:
        writer.write(out)

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
                messagebox.showerror("Ошибка", f"Ошибка при обработке {path}:\\n{e}")
                status_var.set(f"❌ Ошибка при обработке: {os.path.basename(path)}")
    if not any_error:
        status_var.set("✅ PDF-файлы Приложения успешно обработаны.")

def select_file(index):
    path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if path:
        file_paths[index] = path
        file_labels[index].config(text=os.path.basename(path))
        clear_folder_for_merge() # <-- ВЗАИМНОЕ ИСКЛЮЧЕНИЕ

def reset_fields():
    clear_individual_attachments()
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
    
    clear_folder_for_merge() # <-- СБРОС ВЫБОРА ПАПКИ
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
    # Данные для сборки TOC (если доступен модуль bookmarks/PyMuPDF)
    report_toc = None
    attachments_data = []  # <-- ИЗМЕНЕНИЕ: Единый список для данных приложений
    
    main_report_path = None
    base_name_for_save = "merged"
    has_report = False

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
            has_report = True
            # Извлечём TOC отчёта до последующей сборки
            if extract_toc is not None:
                try:
                    report_toc = extract_toc(main_report_path)
                except Exception as e:
                    logging.warning("Не удалось извлечь TOC отчёта: %s", e)
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
    # Извлечение TOC из PDF, созданного из Word (если ещё не извлечён)
    try:
        if report_toc is None and 'word_pdf_path' in locals() and os.path.exists(word_pdf_path) and extract_toc is not None:
            report_toc = extract_toc(word_pdf_path)
    except Exception as e:
        logging.warning("Не удалось извлечь TOC из PDF (Word-конверсия, post): %s", e)

    pdf_temp_paths = []
    for i, path in enumerate(file_paths):
        if path:
            text = entries[i].get().strip()
            temp_pdf = os.path.join(os.path.dirname(path), f"__temp_att_{i+1}.pdf")
            try:
                # Имя файла для заголовка узла Prilog N - <name>
                try:
                    base = os.path.splitext(os.path.basename(path))[0]
                except Exception:
                    base = os.path.basename(path)
                
                # Извлечём TOC исходного файла до штамповки (если доступно)
                toc_entry = None
                if extract_toc is not None:
                    try:
                        toc_entry = extract_toc(path)
                    except Exception as e:
                        logging.warning("Не удалось извлечь TOC приложения (%s): %s", os.path.basename(path), e)
                
                # <-- ИЗМЕНЕНИЕ: Сохраняем индекс, TOC и имя вместе
                if toc_entry:
                    attachments_data.append({"index": i + 1, "toc": toc_entry, "name": base})

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
                status_var.set(f"❌ Ошибка при обработке PDF: {os.path.basename(path)}\\n{e}")
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
            # Применяем TOC
            try:
                if apply_toc is not None and compose_multi_attachment_toc is not None and SourceToc is not None:
                    rep = report_toc if report_toc is not None else SourceToc(entries=[], pages=0)
                    
                    # <-- ИЗМЕНЕНИЕ: Передаем новый структурированный список
                    final_toc = compose_multi_attachment_toc(rep, attachments_data, report_title="Izvestaj", attachment_prefix="Prilog")
                    
                    apply_toc(merged_path, final_toc)
                    logging.info("TOC применён: Izvestaj/Prilog, записей: %d", len(final_toc))
                else:
                    logging.info("Пропуск применения TOC: модуль bookmarks/PyMuPDF недоступен")
            except Exception as e:
                logging.warning("Не удалось применить TOC к итоговому PDF: %s", e)
        else:
            status_var.set("Операция отменена.")
    except Exception as e:
        status_var.set(f"❌ Ошибка при объединении: {e}")
    finally:
        # 4. Удаляем все временные файлы
        for f in temp_files:
            if os.path.exists(f): os.remove(f)


def create_merged_pdf_from_folder():
    # --- ИЗМЕНЕНИЕ: Логика выбора папки вынесена наружу ---
    if not folder_for_merge_path[0]:
        status_var.set("❌ Сначала выберите папку с приложениями!")
        messagebox.showerror("Ошибка", "Папка с PDF-приложениями не выбрана.")
        return

    folder_path = folder_for_merge_path[0]
    # --------------------------------------------------------

    # --- Новая логика: получаем префикс и стартовый номер ---
    try:
        prefix_str = prefix_number_entry.get().strip()
    except (ValueError, NameError):
        prefix_str = ""
    try:
        start_num = int(start_number_entry.get())
    except (ValueError, NameError):
        start_num = 1
    # ----------------------------------------------------

    temp_files = []
    merged_writer = PdfWriter()
    report_toc = None
    attachments_data = []
    
    main_report_path = None
    base_name_for_save = "merged"

    # 1. Определяем основной отчет (PDF или Word)
    if pdf_report_file_path[0]:
        main_report_path = pdf_report_file_path[0]
        text = pdf_report_entry.get().strip()
        base_name_for_save = text or os.path.splitext(os.path.basename(main_report_path))[0]
        temp_main_pdf = os.path.join(os.path.dirname(main_report_path), f"__temp_main_{os.path.basename(main_report_path)}")
        try:
            shutil.copy(main_report_path, temp_main_pdf)
            temp_files.append(temp_main_pdf)
            if extract_toc is not None:
                try:
                    report_toc = extract_toc(main_report_path)
                except Exception as e:
                    logging.warning("Не удалось извлечь TOC отчёта: %s", e)
        except Exception as e:
            status_var.set(f"❌ Ошибка при копировании PDF-отчета: {e}")
            return
    elif word_file_path[0]:
        main_report_path = word_file_path[0]
        base_name_for_save = word_entry.get().strip() or os.path.splitext(os.path.basename(main_report_path))[0]
        word_pdf_path = os.path.join(os.path.dirname(main_report_path), f"{base_name_for_save}.pdf")
        try:
            convert(main_report_path, word_pdf_path)
            if os.path.exists(word_pdf_path):
                temp_files.append(word_pdf_path)
                if extract_toc is not None:
                    report_toc = extract_toc(word_pdf_path)
            else:
                status_var.set("❌ Ошибка: PDF из Word не был создан")
                return
        except Exception as e:
            if "Word.Application.Quit" in str(e) and os.path.exists(word_pdf_path):
                temp_files.append(word_pdf_path)
                if extract_toc is not None:
                    report_toc = extract_toc(word_pdf_path)
            else:
                status_var.set(f"❌ Ошибка при конвертации Word: {e}")
                return

    # 2. Обработка приложений из папки
    try:
        pdf_files = sorted([f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')])
    except Exception as e:
        status_var.set(f"❌ Не удалось прочитать файлы из папки: {e}")
        return

    if not pdf_files:
        status_var.set("⚠ В выбранной папке не найдено PDF-файлов.")
        return
    
    status_var.set(f"Найдено {len(pdf_files)} PDF-файлов. Идет обработка...")
    root.update_idletasks()

    for i, filename in enumerate(pdf_files):
        path = os.path.join(folder_path, filename)
        
        seq_num = start_num + i
        formatted_seq_num = f"{seq_num:02d}"
        new_number_str = f"{prefix_str}{formatted_seq_num}"
        bookmark_name = os.path.splitext(filename)[0]
        stamp_text = f"{new_number_str} - {bookmark_name}"

        temp_pdf = os.path.join(folder_path, f"__temp_att_{i+1}_{filename}")
        try:
            toc_entry = extract_toc(path) if extract_toc is not None else None
            if toc_entry:
                attachments_data.append({"title": stamp_text, "toc": toc_entry})

            reader = PdfReader(path)
            writer = PdfWriter()
            for page in reader.pages:
                _merge_stamp(page, stamp_text, margin=12.0)
                writer.add_page(page)
            with open(temp_pdf, "wb") as f:
                writer.write(f)
            temp_files.append(temp_pdf)
        except Exception as e:
            status_var.set(f"❌ Ошибка при обработке файла: {filename}\\n{e}")
            for f in temp_files:
                if os.path.exists(f): os.remove(f)
            return

    # 3. Объединение PDF
    try:
        for pdf_path in temp_files:
            merged_writer.append(pdf_path)
        
        merged_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"{base_name_for_save}_All_from_folder.pdf",
            title="Сохранить объединённый PDF"
        )
        if merged_path:
            with open(merged_path, "wb") as f:
                merged_writer.write(f)
            last_merged_pdf_path[0] = merged_path
            status_var.set(f"✅ Общий PDF создан: {os.path.basename(merged_path)}")
            create_pdf_link(merged_path)

            # 4. РУЧНАЯ СБОРКА ОГЛАВЛЕНИЯ (TOC)
            try:
                if apply_toc is not None and SourceToc is not None:
                    final_toc = []
                    offset = 0
                    # Отчет
                    if report_toc and report_toc.pages > 0:
                        final_toc.append((1, "Izvestaj", 1))
                        for lvl, title, page in report_toc.entries:
                            final_toc.append((max(2, int(lvl) + 1), str(title), max(1, int(page))))
                        offset = report_toc.pages
                    
                    # Приложения
                    for data in attachments_data:
                        att_title = data["title"]
                        att_toc = data["toc"]
                        final_toc.append((1, att_title, max(1, offset + 1)))
                        if att_toc:
                            for lvl, title, page in att_toc.entries:
                                final_toc.append((max(2, int(lvl) + 1), str(title), max(1, int(page) + offset)))
                            offset += att_toc.pages

                    apply_toc(merged_path, final_toc)
                    logging.info("TOC применён (ручная сборка): Izvestaj/Prilog, записей: %d", len(final_toc))
            except Exception as e:
                logging.warning("Не удалось применить TOC к итоговому PDF: %s", e)
        else:
            status_var.set("Операция отменена.")
    except Exception as e:
        status_var.set(f"❌ Ошибка при объединении: {e}")
    finally:
        # 5. Удаляем все временные файлы
        for f in temp_files:
            if os.path.exists(f): os.remove(f)


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

# --- Блок 1: Основной отчет ---
report_frame = tk.LabelFrame(root, text="Блок 1: Основной отчет", bg=BG_COLOR, fg="#222", font=("Segoe UI", 10, "bold"))
report_frame.pack(padx=20, pady=10, fill='x')

# --- Виджеты для Word-файла ---
word_frame = tk.Frame(report_frame, bg=BG_COLOR)
word_frame.pack(padx=10, pady=(5,10), fill='x')

top_row = tk.Frame(word_frame, bg=BG_COLOR)
top_row.pack(fill='x')

word_entry = tk.Entry(top_row, width=40, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
word_entry.insert(0, "Izveštaj_Отчет")
word_entry.pack(side='left', padx=(0, 10))

word_btn = tk.Button(top_row, text="📄 Выбрать Word (.docx)", command=select_word_file, bg=BTN_COLOR, relief="flat")
word_btn.pack(side='left', padx=(0, 10))

word_convert_btn = tk.Button(top_row, text="➡️ Создать PDF из word", 
                          command=convert_word_to_pdf, bg=BTN_COLOR, relief="flat",
                          width=30)
word_convert_btn.pack(side='right', padx=20)

word_convert_note = tk.Label(word_frame, text="Создает PDF из docx файла без приложений", 
                           anchor='e', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
word_convert_note.pack(side='right', padx=20, pady=(1, 0))

word_file_label = tk.Label(word_frame, text="Файл не выбран", anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
word_file_label.pack(fill='x', pady=(1, 0))

# --- Виджеты для PDF-отчета ---
pdf_report_frame = tk.Frame(report_frame, bg=BG_COLOR)
pdf_report_frame.pack(padx=10, pady=(5,10), fill='x')

pdf_report_top_row = tk.Frame(pdf_report_frame, bg=BG_COLOR)
pdf_report_top_row.pack(fill='x')

pdf_report_entry = tk.Entry(pdf_report_top_row, width=40, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
pdf_report_entry.insert(0, "Izveštaj_Отчет")
pdf_report_entry.pack(side='left', padx=(0, 10))

pdf_report_btn = tk.Button(pdf_report_top_row, text="📄 Выбрать PDF (.pdf)", command=select_pdf_report_file, bg=BTN_COLOR, relief="flat")
pdf_report_btn.pack(side='left', padx=(0, 10))

pdf_report_label = tk.Label(pdf_report_frame, text="Файл не выбран", anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
pdf_report_label.pack(fill='x', pady=(1, 0))


# --- Блок 2: Приложения (Ручной режим) ---
manual_apps_frame = tk.LabelFrame(root, text="Блок 2: Приложения", bg=BG_COLOR, fg="#222", font=("Segoe UI", 10, "bold"))
manual_apps_frame.pack(padx=20, pady=5, fill='x')

# --- Создаем главный фрейм, который разделим на две колонки ---
main_manual_frame = tk.Frame(manual_apps_frame, bg=BG_COLOR)
main_manual_frame.pack(padx=10, pady=5, fill='x')

left_col_frame = tk.Frame(main_manual_frame, bg=BG_COLOR)
left_col_frame.pack(side='left', fill='y', expand=True)

right_col_frame = tk.Frame(main_manual_frame, bg=BG_COLOR)
right_col_frame.pack(side='right', fill='y', padx=(20, 5))

# Фрейм для полей ввода и выбора файлов (в левой колонке)
for i in range(6):
    frame = tk.Frame(left_col_frame, bg=BG_COLOR)
    frame.pack(padx=10, pady=6, fill='x')
    entry = tk.Entry(frame, width=35, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
    entry.insert(0, f"Prilog / Приложение 7.0{i+1}")
    entry.pack(side='left', padx=(0, 10))
    entries.append(entry)
    btn = tk.Button(frame, text="📂 Выбрать PDF", command=lambda idx=i: select_file(idx), bg=BTN_COLOR, relief="flat")
    btn.pack(side='left', padx=(0, 10))
    label = tk.Label(frame, text="Файл не выбран", width=45, anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 8))
    label.pack(side='left')
    file_labels.append(label)

# --- ВОССТАНОВЛЕННЫЙ БЛОК --- 
# Кнопки для индивидуальной обработки (в правой колонке)
btn_style = {"width": 30, "bg": BTN_COLOR, "activebackground": "#d5d5d5", "relief": "flat"}
tk.Button(right_col_frame, text="💾 Сохранить в тот же файл PDF", 
         command=lambda: process_pdfs(False), **btn_style).pack(pady=3)
tk.Button(right_col_frame, text="📝 Сохранить с переименованием", 
         command=lambda: process_pdfs(True), **btn_style).pack(pady=3)

note_text = (
    "💾 Сохранить в тот же файл – заменяет оригинал PDF.\n"
    "📝 Сохранить с переименованием – создаёт копию pdf с префиксом 'att.X_...'\n\n"
    "Эти кнопки обрабатывают каждый файл индивидуально, \n"
    "но НЕ объединяют их в общий PDF."
)
note_label = tk.Label(right_col_frame, text=note_text, justify='left', wraplength=220, 
                     bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
note_label.pack(pady=(10, 0))
# --- КОНЕЦ ВОССТАНОВЛЕННОГО БЛОКА ---

# Кнопка для ручного режима (под левой колонкой)
manual_merge_action_frame = tk.Frame(manual_apps_frame, bg=BG_COLOR)
manual_merge_action_frame.pack(pady=10, anchor='w', padx=10)

merge_btn = tk.Button(manual_merge_action_frame, 
                     text="📚 Создать общий PDF из Блок 1 (Отчет) + Блок 2",
                     command=create_merged_pdf,
                     relief="flat",
                     bg="#4CAF50",
                     fg="white",
                     activebackground="#45a049")
merge_btn.pack(side='left')

note_merge_btn = tk.Label(manual_merge_action_frame, text="Создает общий PDF из отчета (Блок 1) и 6 слотов Приложений выше (Блок 2).", 
                     bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
note_merge_btn.pack(side='left', padx=(10, 0))


# --- Блок 3: Приложения из папки (Автоматический режим) ---
folder_apps_frame = tk.LabelFrame(root, text="Блок 3: Приложения из папки", bg=BG_COLOR, fg="#222", font=("Segoe UI", 10, "bold"))
folder_apps_frame.pack(padx=20, pady=10, fill='x')

# --- Верхняя часть блока 3 с выбором папки и стартовым номером ---
top_folder_frame = tk.Frame(folder_apps_frame, bg=BG_COLOR)
top_folder_frame.pack(padx=10, pady=10, fill='x')

folder_select_frame = tk.Frame(top_folder_frame, bg=BG_COLOR)
folder_select_frame.pack(side='left', fill='x', expand=True)

folder_btn = tk.Button(folder_select_frame, text="📂 Выбрать папку с Приложениями", command=select_folder_for_merge, bg=BTN_COLOR, relief="flat")
folder_btn.pack(side='left', padx=(0, 10))

folder_for_merge_label = tk.Label(folder_select_frame, text="Папка не выбрана", anchor='w', bg=BG_COLOR, fg="#555", font=("Segoe UI", 9))
folder_for_merge_label.pack(side='left')

# --- Новый виджет для префикса ---
prefix_num_frame = tk.Frame(top_folder_frame, bg=BG_COLOR)
prefix_num_frame.pack(side='left', padx=(20, 0))

tk.Label(prefix_num_frame, text="Префикс нумерации:", bg=BG_COLOR).pack(side='left')
prefix_number_entry = tk.Entry(prefix_num_frame, width=9, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
prefix_number_entry.insert(0, "Prilog 7.")
prefix_number_entry.pack(side='left', padx=5)
# --- Конец нового виджета ---

# --- Виджет для стартового номера ---
start_num_frame = tk.Frame(top_folder_frame, bg=BG_COLOR)
start_num_frame.pack(side='right', padx=(20, 0))

tk.Label(start_num_frame, text="Начать нумерацию с:", bg=BG_COLOR).pack(side='left')
start_number_entry = tk.Spinbox(start_num_frame, from_=1, to=999, width=4, bg=ENTRY_BG, fg=ENTRY_FG, relief="solid", bd=1)
start_number_entry.insert(0, "1")
start_number_entry.pack(side='left', padx=5)
# --- Конец нового виджета ---

# Кнопка для режима из папки
folder_merge_action_frame = tk.Frame(folder_apps_frame, bg=BG_COLOR)
folder_merge_action_frame.pack(pady=10, anchor='w', padx=10)

folder_merge_btn = tk.Button(folder_merge_action_frame,
                             text="🗂️ Создать общий PDF из Блок 1 (Отчет) + Блок 3",
                             command=create_merged_pdf_from_folder,
                             relief="flat",
                             bg="#FF9800",
                             fg="white",
                             activebackground="#FB8C00")
folder_merge_btn.pack(side='left')

note_folder_merge_btn = tk.Label(folder_merge_action_frame, text="Создает общий PDF из отчета (Блок 1) и Папки (Блок 3).",
                                bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
note_folder_merge_btn.pack(side='left', padx=(10, 0))


# --- Блок 4: Сброс и Статус ---
# Разделительная линия
separator = tk.Frame(root, height=2, bg="#e0e0e0")
separator.pack(fill='x', padx=20, pady=(10, 5))

# Кнопка сброса
reset_action_frame = tk.Frame(root, bg=BG_COLOR)
reset_action_frame.pack(pady=10, anchor='w', padx=30)

reset_btn = tk.Button(reset_action_frame, text="🔄 Сброс/Вернуть по умолчанию", 
         command=reset_fields, bg=BTN_COLOR, relief="flat")
reset_btn.pack(side='left')

note_reset_btn = tk.Label(reset_action_frame, text="Сбрасывает все поля и очищает все выбранные файлы.",
                        bg=BG_COLOR, fg="#444", font=("Segoe UI", 8))
note_reset_btn.pack(side='left', padx=(10, 0))


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