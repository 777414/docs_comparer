#!/usr/bin/env python3
"""Сравнение двух версий DOCX-документа с учетом порядка пунктов."""

import json
import re
import sys
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

from docx import Document


MAIN_TITLE_PATTERN = re.compile(
    r"перечень\s+работ\s+на\s+объекте\s+долевого\s+строительства",
    re.IGNORECASE,
)
SIGNATURE_PATTERN = re.compile(
    r"^\s*(?:застройщик|участник\s+долевого\s+строительства)\s*:",
    re.IGNORECASE,
)
LIST_MARKER_PATTERN = re.compile(
    r"^\s*(?:[-–—•▪◦●○]\s*|(?:\d+|[а-яёa-z])[.)]\s*|[IVXLCDM]+[.)]\s*)",
    re.IGNORECASE,
)

HEADING_SIMILARITY_THRESHOLD = 0.78
MAX_HEADING_LENGTH = 100


def normalize_text(text: str) -> str:
    """Нормализует пробелы и технические переносы, сохраняя содержание."""
    text = text.replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "")
    text = text.replace("\r", "\n")
    return re.sub(r"\s+", " ", text).strip()


def normalize_content(text: str) -> str:
    """Ключ сравнения: регистр, пробелы и начальные маркеры списка не важны."""
    text = normalize_text(text)
    text = LIST_MARKER_PATTERN.sub("", text).strip()
    return text.casefold()


def normalize_heading(text: str) -> str:
    return normalize_content(text).strip(" \t\n.:;")


def clean_paragraph(text: str) -> str:
    return LIST_MARKER_PATTERN.sub("", normalize_text(text)).strip()


def read_docx(file_path: str) -> list[dict[str, str]]:
    """Читает только абзацы DOCX; таблицы намеренно игнорируются."""
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"Файл не найден: {path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Ожидается файл .docx: {path}")

    document = Document(str(path))
    blocks = []
    for paragraph in document.paragraphs:
        text = normalize_text(paragraph.text)
        if text:
            blocks.append({
                "text": text,
                "style": paragraph.style.name or "",
            })

    if not blocks:
        raise ValueError(f"В документе нет текста: {path}")
    return blocks


def find_content_start(blocks: list[dict[str, str]]) -> int:
    """Игнорирует текст до заголовка перечня работ, если он найден."""
    for index, block in enumerate(blocks):
        if MAIN_TITLE_PATTERN.search(block["text"]):
            return index + 1
    return 0


def is_signature_start(text: str) -> bool:
    return bool(SIGNATURE_PATTERN.match(text))


def is_bullet(text: str) -> bool:
    return bool(LIST_MARKER_PATTERN.match(text))


def is_heading(blocks: list[dict[str, str]], index: int) -> bool:
    """Эвристически определяет заголовок без фиксированного списка разделов."""
    block = blocks[index]
    text = normalize_text(block["text"])

    if not text or MAIN_TITLE_PATTERN.search(text):
        return False

    style = block["style"].casefold()
    if style.startswith(("heading", "заголовок")):
        return True

    if len(text) > MAX_HEADING_LENGTH or is_bullet(text):
        return False
    if index + 1 >= len(blocks):
        return False

    next_text = blocks[index + 1]["text"]
    next_style = blocks[index + 1]["style"].casefold()

    if is_bullet(next_text) or text.endswith(":"):
        return True

    if text.endswith(".") and (
        next_style.startswith(("heading", "заголовок")) or is_bullet(next_text)
    ):
        return True

    # Не считаем произвольную короткую строку заголовком только по длине:
    # это снижает риск разбиения обычного текста на ложные разделы.
    return False


def extract_sections(
    blocks: list[dict[str, str]],
) -> dict[str, dict[str, object]]:
    """Разбивает документ на разделы и сохраняет пункты независимо."""
    start = find_content_start(blocks)
    sections: dict[str, dict[str, object]] = {}
    current_heading = "вводная часть"
    current_items: list[str] = []

    def save_section() -> None:
        key = normalize_heading(current_heading)
        if key in sections:
            suffix = 2
            while f"{key} [{suffix}]" in sections:
                suffix += 1
            key = f"{key} [{suffix}]"
        sections[key] = {
            "display_name": current_heading,
            "items": current_items.copy(),
        }

    for index in range(start, len(blocks)):
        text = blocks[index]["text"]
        if is_signature_start(text):
            break

        if is_heading(blocks, index):
            save_section()
            current_heading = normalize_text(text).rstrip(".:")
            current_items = []
        else:
            item = clean_paragraph(text)
            if item:
                current_items.append(item)

    save_section()
    return sections


def section_counter(section: dict[str, object]) -> Counter:
    """Сравнивает пункты без учета порядка, сохраняя число дубликатов."""
    return Counter(
        normalized
        for item in section["items"]
        if (normalized := normalize_content(item))
    )


def heading_similarity(first: str, second: str) -> float:
    return SequenceMatcher(
        None, normalize_heading(first), normalize_heading(second)
    ).ratio()


def match_sections(
    old_sections: dict[str, dict[str, object]],
    new_sections: dict[str, dict[str, object]],
) -> tuple[list[tuple[str, str]], list[str], list[str]]:
    """Сначала сопоставляет точные заголовки, затем похожие."""
    old_keys = list(old_sections)
    new_keys = list(new_sections)
    unmatched_old = set(old_keys)
    unmatched_new = set(new_keys)
    matches: list[tuple[str, str]] = []

    for old_key in old_keys:
        if old_key in unmatched_new:
            matches.append((old_key, old_key))
            unmatched_old.remove(old_key)
            unmatched_new.remove(old_key)

    candidates = []
    for old_key in unmatched_old:
        for new_key in unmatched_new:
            score = heading_similarity(old_key, new_key)
            if score >= HEADING_SIMILARITY_THRESHOLD:
                candidates.append((score, old_key, new_key))

    candidates.sort(reverse=True)
    for _, old_key, new_key in candidates:
        if old_key in unmatched_old and new_key in unmatched_new:
            matches.append((old_key, new_key))
            unmatched_old.remove(old_key)
            unmatched_new.remove(new_key)

    return (
        matches,
        [key for key in old_keys if key in unmatched_old],
        [key for key in new_keys if key in unmatched_new],
    )


def compare_documents(
    old_sections: dict[str, dict[str, object]],
    new_sections: dict[str, dict[str, object]],
) -> list[str]:
    differences = []
    matches, removed, added = match_sections(old_sections, new_sections)

    for key in removed:
        if key != "вводная часть":
            differences.append(
                f"Удалён: раздел «{old_sections[key]['display_name']}»"
            )

    for key in added:
        if key != "вводная часть":
            differences.append(
                f"Добавлен: раздел «{new_sections[key]['display_name']}»"
            )

    for old_key, new_key in matches:
        if section_counter(old_sections[old_key]) == section_counter(new_sections[new_key]):
            continue

        if old_key == "вводная часть":
            differences.append("Изменён: вводная часть")
        else:
            differences.append(
                f"Изменён: раздел «{new_sections[new_key]['display_name']}»"
            )

    return differences


def build_result(differences: list[str]) -> dict:
    return {"total": len(differences), "diffs": differences}


def main() -> int:
    if len(sys.argv) != 3:
        print(
            f"Использование: {Path(sys.argv[0]).name} old.docx new.docx",
            file=sys.stderr,
        )
        return 2

    try:
        old_sections = extract_sections(read_docx(sys.argv[1]))
        new_sections = extract_sections(read_docx(sys.argv[2]))
        result = build_result(compare_documents(old_sections, new_sections))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError, KeyError) as exc:
        print(
            json.dumps({"error": str(exc)}, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
