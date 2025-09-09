from pathlib import Path
from pypdf import PdfWriter, PdfReader

# Импортируем функцию из UI-модуля, чтобы проверить перенос закладок
from pdf_attachments_ui import merge_pdfs_with_bookmarks


def _make_pdf_with_bookmark(path: Path, inner_title: str) -> None:
    """Создать простой PDF с одной страницей и одной внутренней закладкой."""
    w = PdfWriter()
    w.add_blank_page(width=200, height=200)
    # Закладка на первую страницу
    w.add_outline_item(inner_title, page_number=0)
    with open(path, "wb") as f:
        w.write(f)


def _collect_outline_titles(reader: PdfReader) -> list[str]:
    """Собрать все заголовки из дерева закладок (по возможности)."""
    titles: list[str] = []
    outline = getattr(reader, "outline", None)
    if outline is None:
        outline = getattr(reader, "outlines", None)

    def walk(node):
        if node is None:
            return
        if isinstance(node, list):
            for x in node:
                walk(x)
            return
        # Пытаемся прочитать заголовок из возможных атрибутов
        title = getattr(node, "title", None)
        if title:
            titles.append(str(title))
        # Рекурсивно обходим потомков, если свойство существует
        children = getattr(node, "children", None)
        if children:
            for c in children:
                walk(c)

    walk(outline)
    return titles


def test_merge_preserves_and_adds_top_level_bookmarks(tmp_path: Path):
    a = tmp_path / "a.pdf"
    b = tmp_path / "b.pdf"
    out = tmp_path / "merged.pdf"

    _make_pdf_with_bookmark(a, "Inner A")
    _make_pdf_with_bookmark(b, "Inner B")

    parts = [
        (str(a), "Report"),
        (str(b), "Prilog 1"),
    ]

    merge_pdfs_with_bookmarks(parts, str(out))

    r = PdfReader(str(out))
    titles = _collect_outline_titles(r)

    # Верхнеуровневые закладки присутствуют
    assert "Report" in titles
    assert "Prilog 1" in titles

    # Исходные закладки также перенесены
    assert "Inner A" in titles
    assert "Inner B" in titles

