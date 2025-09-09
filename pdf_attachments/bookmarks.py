"""Утилиты для работы с оглавлением (закладками) PDF.

Используется PyMuPDF (fitz) для извлечения и записи TOC. Если PyMuPDF
недоступен, функции выполняются в «пустом» режиме без ошибок, но без эффекта.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List, Sequence, Tuple
import logging


try:
    import fitz  # type: ignore
except Exception as exc:  # pragma: no cover
    fitz = None  # type: ignore
    logging.getLogger(__name__).warning("PyMuPDF (fitz) не установлен: %s", exc)


TocEntry = Tuple[int, str, int]


@dataclass(frozen=True)
class SourceToc:
    """Контейнер TOC источника.

    Attributes:
        entries: Список записей TOC формата [level, title, page] (1-based).
        pages: Количество страниц исходного документа.
    """

    entries: List[TocEntry]
    pages: int


def extract_toc(path: str) -> SourceToc:
    """Извлечь TOC и число страниц из PDF.

    Если PyMuPDF недоступен или TOC отсутствует, возвращает пустые записи,
    при корректном pages.
    """
    if fitz is None:  # pragma: no cover
        return SourceToc(entries=[], pages=0)

    with fitz.open(path) as doc:  # type: ignore[attr-defined]
        try:
            toc: List[TocEntry] = [(int(l), str(t), int(p)) for l, t, p in doc.get_toc(simple=True)]
        except Exception:
            toc = []
        pages = int(doc.page_count)
    return SourceToc(entries=toc, pages=pages)


def compose_two_level_toc(
    report: SourceToc,
    attachments: Sequence[SourceToc],
    report_title: str = "Izvestaj",
    attachments_title: str = "Prilog",
) -> List[TocEntry]:
    """Собрать итоговый TOC с двумя верхними уровнями.

    - Уровень 1: report_title (стр. 1), далее все закладки отчёта со смещением +0 и level+1.
    - Уровень 1: attachments_title (стр. offset+1), далее все закладки приложений
      с level+1 и корректным смещением по сумме страниц.
    """
    result: List[TocEntry] = []

    # Блок отчёта
    if report.pages > 0:
        result.append((1, report_title, 1))
        for lvl, title, page in report.entries:
            lvl_adj = max(2, int(lvl) + 1)
            page_adj = max(1, int(page))
            result.append((lvl_adj, title, page_adj))

    # Смещение до начала приложений: это количество страниц отчёта
    offset = report.pages

    # Блок приложений
    if attachments:
        result.append((1, attachments_title, max(1, offset + 1)))

    cur_offset = offset
    for att in attachments:
        if att.pages <= 0:
            continue
        for lvl, title, page in att.entries:
            lvl_adj = max(2, int(lvl) + 1)
            page_adj = max(1, int(page) + cur_offset)
            result.append((lvl_adj, title, page_adj))
        cur_offset += att.pages

    return result


def compose_multi_attachment_toc(
    report: SourceToc,
    attachments: Sequence[SourceToc],
    report_title: str = "Izvestaj",
    attachment_prefix: str = "Prilog",
) -> List[TocEntry]:
    """Собрать TOC с верхним уровнем для отчёта и отдельным верхним уровнем для каждого приложения.

    Пример результата:
    - [1, "Izvestaj", 1]
      + дочерние отчёта (lvl+1)
    - [1, "Prilog 1", offset_of_att1+1]
      + дочерние att1 (lvl+1, page+offset_of_att1)
    - [1, "Prilog 2", offset_of_att2+1]
      + дочерние att2 (lvl+1, page+offset_of_att2)
    """
    result: List[TocEntry] = []

    # Отчёт
    if report.pages > 0:
        result.append((1, report_title, 1))
        for lvl, title, page in report.entries:
            result.append((max(2, int(lvl) + 1), str(title), max(1, int(page))))

    # Смещение для приложений
    offset = report.pages

    for idx, att in enumerate(attachments, start=1):
        if att.pages <= 0:
            continue
        top = f"{attachment_prefix} {idx}"
        result.append((1, top, max(1, offset + 1)))
        for lvl, title, page in att.entries:
            result.append((max(2, int(lvl) + 1), str(title), max(1, int(page) + offset)))
        offset += att.pages

    return result


def apply_toc(target_pdf_path: str, toc: Iterable[TocEntry]) -> None:
    """Применить TOC к PDF. Перезаписывает документ с новым TOC.

    Безопасно для вызова с пустым `toc` (в таком случае изменений нет).
    """
    entries = list(toc)
    if not entries or fitz is None:  # pragma: no cover
        return

    # Открываем и устанавливаем TOC
    import os
    import tempfile

    dir_name = os.path.dirname(target_pdf_path) or "."
    fd, tmp_path = tempfile.mkstemp(prefix="_toc_", suffix=".pdf", dir=dir_name)
    os.close(fd)
    try:
        with fitz.open(target_pdf_path) as doc:  # type: ignore[attr-defined]
            doc.set_toc(entries)
            doc.save(tmp_path)
        # Перенос после закрытия исходного документа (важно для Windows)
        os.replace(tmp_path, target_pdf_path)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
