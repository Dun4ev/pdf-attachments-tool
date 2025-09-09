from pathlib import Path


def test_agents_md_exists_and_has_required_headings() -> None:
    """Проверяет наличие AGENTS.md и ключевых заголовков.

    - Файл лежит в корне репозитория.
    - Первый заголовок — "# Repository Guidelines".
    - Присутствуют обязательные секции.
    """

    repo_root = Path(__file__).resolve().parents[1]
    md_path = repo_root / "AGENTS.md"
    assert md_path.exists(), "Ожидался файл AGENTS.md в корне репозитория"

    content = md_path.read_text(encoding="utf-8")
    # Первая строка — заголовок документа
    first_line = content.splitlines()[0].strip()
    assert (
        first_line == "# Repository Guidelines"
    ), "Первый заголовок должен быть '# Repository Guidelines'"

    required = [
        "## Project Structure & Module Organization",
        "## Build, Test, and Development Commands",
        "## Coding Style & Naming Conventions",
        "## Testing Guidelines",
        "## Commit & Pull Request Guidelines",
        "## Security & Configuration Tips",
    ]
    for heading in required:
        assert heading in content, f"Отсутствует раздел: {heading}"

