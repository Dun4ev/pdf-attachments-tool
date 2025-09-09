# Repository Guidelines

## Project Structure & Module Organization

- Source: `src/pdf_attachments/` — основная логика и CLI (`cli.py`).
- Tests: `tests/` — зеркалирует структуру пакета, фикстуры в `tests/conftest.py`.
- Scripts: `scripts/` — одноразовые утилиты и примеры запуска.
- Data: `data/` — тестовые файлы PDF/вложения (не коммитим чувствительное).
- Docs: `docs/` — заметки, схемы, ADR.
- Именование: модули/пакеты — `snake_case`, классы — `PascalCase`, функции/переменные — `snake_case`.

## Build, Test, and Development Commands

- Создать окружение (Windows): `py -3.10 -m venv .venv && .\.venv\Scripts\activate`.
- Установка: `pip install -U pip && pip install -e ".[dev]"`.
- Тесты: `pytest -q` (см. покрытие ниже).
- Стиль: `ruff check .` и автоформат `ruff format .` (или `black .`).
- Запуск CLI: `python -m pdf_attachments.cli --help`.

## Coding Style & Naming Conventions

- Python ≥3.10, PEP8, отступы 4 пробела, типы везде.
- Докстринги — стиль Google; избегать длинных функций, одна ответственность.
- Логирование через `logging` (без `print`); структурированные сообщения.
- Конфигурация через переменные окружения/`.env` (см. ниже).

## Testing Guidelines

- Framework: `pytest`; имена тестов `test_*.py`, функции — `test_*`.
- Покрытие: цель ≥85% (`pytest --cov=pdf_attachments --cov-report=term-missing`).
- Тесты зеркалируют дерево `src/`; edge-cases для файлов/путей/кодировок.
- При багфиксе: сначала падающий тест, затем исправление.

## Commit & Pull Request Guidelines

- Коммиты: Conventional Commits (`feat:`, `fix:`, `docs:`, `test:`, `chore:`); заголовок ≤72 символов.
- Сообщение вкла для чючает «почему/что/как»; ссылки на задачи.
- PR: краткое описание, связанные issues, инструкции для ревью, скриншоты/логи для UX/CLI-изменений.

## Security & Configuration Tips

- Секреты: не коммитить; используйте `.env` и образец `.env.example`.
- Пути/файлы: санитайзинг, запрет на произвольные пути извне.
- Режимы: по умолчанию dry‑run для потенциально разрушающих действий; явные флаги `--yes/--force`.
