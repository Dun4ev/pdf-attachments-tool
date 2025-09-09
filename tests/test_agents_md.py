"""Тесты для файла AGENTS.md.

Проверяются базовые инварианты: наличие файла, заголовок и ключевые разделы.
"""

from pathlib import Path


def test_agents_md_exists() -> None:
    path = Path("AGENTS.md")
    assert path.exists(), "Ожидался файл AGENTS.md в корне репозитория"


def test_agents_md_structure() -> None:
    content = Path("AGENTS.md").read_text(encoding="utf-8")
    # Заголовок документа
    assert content.splitlines()[0].strip() == "# Repository Guidelines"
    # Ключевые разделы
    required_sections = [
        "## Структура проекта",
        "## Сборка, запуск и тесты",
        "## Стиль кода и нейминг",
        "## Рекомендации по тестированию",
        "## Коммиты и Pull Request’ы",
        "## Безопасность и конфигурация",
    ]
    for section in required_sections:
        assert section in content, f"Отсутствует раздел: {section}"

