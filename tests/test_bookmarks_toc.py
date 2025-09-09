"""Тесты логики сборки/применения TOC (закладок).

Требует наличия PyMuPDF (fitz). Если отсутствует — тесты автоматически скипаются.
"""

from pathlib import Path
import pytest


pytest.importorskip("fitz")

import fitz  # type: ignore

from pdf_attachments.bookmarks import extract_toc, compose_two_level_toc, apply_toc


def _make_pdf(path: Path, pages: int, toc: list[list[int | str]]) -> None:
    """Создать PDF с заданным числом страниц и TOC.

    toc: список троек [level, title, page]. Страницы — 1-based.
    """
    doc = fitz.open()  # type: ignore[attr-defined]
    for _ in range(pages):
        doc.new_page()
    if toc:
        doc.set_toc([[int(l), str(t), int(p)] for l, t, p in toc])
    doc.save(path.as_posix())
    doc.close()


def test_compose_and_apply_toc(tmp_path: Path) -> None:
    # Подготовка исходных документов
    report_pdf = tmp_path / "report.pdf"
    attach_pdf = tmp_path / "attach.pdf"

    # Исходные TOC
    report_src_toc = [
        [1, "R-Top", 1],
        [2, "R-Sub", 2],
    ]
    attach_src_toc = [
        [1, "A-Top", 1],
        [2, "A-Sub", 3],
    ]

    _make_pdf(report_pdf, pages=2, toc=report_src_toc)
    _make_pdf(attach_pdf, pages=3, toc=attach_src_toc)

    rep = extract_toc(report_pdf.as_posix())
    att = extract_toc(attach_pdf.as_posix())

    # Собираем итоговый TOC
    # Вариант B: отдельные верхние уровни для каждого приложения
    from pdf_attachments.bookmarks import compose_multi_attachment_toc
    final_toc = compose_multi_attachment_toc(
        rep,
        [att],
        report_title="Izvestaj",
        attachment_prefix="Prilog",
        attachment_names=["attach"],
    )

    # Объединяем файлы в один (через PyMuPDF для простоты теста)
    merged = tmp_path / "merged.pdf"
    doc = fitz.open()  # type: ignore[attr-defined]
    doc.insert_pdf(fitz.open(report_pdf.as_posix()))
    doc.insert_pdf(fitz.open(attach_pdf.as_posix()))
    doc.save(merged.as_posix())
    doc.close()

    # Применяем TOC и сверяем
    apply_toc(merged.as_posix(), final_toc)

    with fitz.open(merged.as_posix()) as out_doc:  # type: ignore[attr-defined]
        toc = out_doc.get_toc(simple=True)

    # Верхний уровень и смещения
    assert [1, "Izvestaj", 1] in toc
    # Имя файла добавляется к верхнему уровню
    assert [1, "Prilog 1 - attach", 3] in toc  # отчёт 2 стр., приложения начинаются с 3

    # Дочерние элементы отчёта повышены на уровень и без смещения страниц
    assert [2, "R-Top", 1] in toc
    assert [3, "R-Sub", 2] in toc

    # Дочерние элементы приложений повышены на уровень и со смещением +2
    assert [2, "A-Top", 3] in toc
    assert [3, "A-Sub", 5] in toc
