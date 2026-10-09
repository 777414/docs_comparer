#!/usr/bin/env python3
"""Сравнение двух версий DOCX-документа без учета порядка пунктов."""

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
# Осторожная эвристика: не используем неоднозначные окончания вроде «-ены»,
# которые встречаются в существительных («стены»).
VERB_FORM_PATTERN = re.compile(
    r"\b[а-яёa-z]+(?:ется|ются|ится|атся|ятся|ывается|ивается|"
    r"ает|яет|ует|юет|уют|ают|яют|ишь|ешь|ете|им|ите|"
    r"ать|ять|ить|ыть|ться|чь|ено|ена|ал|ала|али|ил|ила|или)\b",
    re.IGNORECASE,
)
HEADING_SIMILARITY_THRESHOLD = 0.78
MAX_HEADING_LENGTH = 100


def normalize_text(text: str) -> str:
    """Устраняет технические различия пробелов и переносов."""
    text = text.replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "")
    text = text.replace("\r", "\n")
    return re.sub(r"\s+", " ", text).strip()


def normalize_content(text: str) -> str:
    """Нормализует текст, сохраняя пунктуацию, числа и отрицания."""
    text = normalize_text(text)
    text = LIST_MARKER_PATTERN.sub("", text).strip()
    return text.casefold()


def normalize_heading(text: str) -> str:
    return normalize_content(text).strip(" \t\n.:;")


def read_docx(file_path: str) -> list[dict[str, str]]:
    """Читает абзацы; таблицы игнорирует, учитывает заголовки внутри форматированных runs."""
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"Файл не найден: {path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Ожидается файл .docx: {path}")

    document = Document(str(path))
    blocks: list[dict[str, str]] = []

    def looks_like_inline_heading(run) -> bool:
        text = normalize_text(run.text)
        return bool(
            run.bold
            and text
            and len(text) <= MAX_HEADING_LENGTH
            and len(text.split()) <= 7
            and text.endswith((".", ":"))
            and not VERB_FORM_PATTERN.search(text)
        )

    for paragraph in document.paragraphs:
        # Заголовок может быть отдельным жирным run в том же абзаце,
        # что и предыдущий раздел (например, «Прочее.»).
        segments: list[str] = []
        current = ""
        for run in paragraph.runs:
            if looks_like_inline_heading(run) and len(normalize_text(current)) > 80:
                if normalize_text(current):
                    segments.append(current)
                segments.append(run.text)
                current = ""
            else:
                current += run.text
        if normalize_text(current):
            segments.append(current)

        for segment in segments:
            if normalize_text(segment):
                blocks.append({
                    "text": segment,
                    "style": paragraph.style.name or "",
                })

    if not blocks:
        raise ValueError(f"В документе нет текста: {path}")
    return blocks


def find_content_start(blocks: list[dict[str, str]]) -> int:
    """Игнорирует сведения до заголовка перечня работ, если он найден."""
    for index, block in enumerate(blocks):
        if MAIN_TITLE_PATTERN.search(normalize_text(block["text"])):
            return index + 1
    return 0


def is_signature_start(text: str) -> bool:
    return bool(SIGNATURE_PATTERN.match(text))


def is_bullet(text: str) -> bool:
    return bool(LIST_MARKER_PATTERN.match(text))


def is_heading(blocks: list[dict[str, str]], index: int) -> bool:
    """Эвристически определяет заголовки без фиксированного списка названий."""
    text = normalize_text(blocks[index]["text"])
    if not text or MAIN_TITLE_PATTERN.search(text) or is_bullet(text):
        return False

    style = blocks[index]["style"].casefold()
    if style.startswith(("heading", "заголовок")):
        return True
    if len(text) > MAX_HEADING_LENGTH or index + 1 >= len(blocks):
        return False
    if VERB_FORM_PATTERN.search(text):
        return False

    words = text.split()
    if len(words) <= 7 and (text.endswith((".", ":")) or len(words) <= 3):
        return True

    next_text = normalize_text(blocks[index + 1]["text"])
    if text.endswith(":") and len(words) <= 12:
        return True
    if is_bullet(next_text) and len(words) <= 10:
        return True
    return False


def extract_sections(
    blocks: list[dict[str, str]],
) -> dict[str, dict[str, object]]:
    """Разбивает документ на разделы, сохраняя текст пунктов и границы абзацев."""
    start = find_content_start(blocks)
    sections: dict[str, dict[str, object]] = {}
    current_heading = "вводная часть"
    current_paragraphs: list[str] = []

    def save_section() -> None:
        key = normalize_heading(current_heading)
        if key in sections:
            suffix = 2
            while f"{key} [{suffix}]" in sections:
                suffix += 1
            key = f"{key} [{suffix}]"
        sections[key] = {
            "display_name": current_heading,
            "text": "\n".join(current_paragraphs),
        }

    for index in range(start, len(blocks)):
        text = normalize_text(blocks[index]["text"])
        if is_signature_start(text):
            break
        if is_heading(blocks, index):
            save_section()
            current_heading = text.rstrip(".:")
            current_paragraphs = []
        else:
            current_paragraphs.append(text)

    save_section()
    return sections


def split_content_units(text: str) -> Counter:
    """
    Сравнивает пункты/предложения как мультимножество:
    порядок и разбиение по абзацам не важны, дубликаты сохраняются.
    """
    text = text.replace("\r", "\n")
    # Разделяем маркеры списков в начале абзаца и после пунктуации,
    # но не разделяем дефис в конструкции вроде «Полы - выполняется...».
    text = re.sub(
        r"(?:^|\n|(?<=[.;]))\s*[-–—•▪◦●○]\s*",
        "\n",
        text,
    )
    chunks = re.split(r"(?<=[.!?;])\s+", text)
    units = []
    for chunk in chunks:
        normalized = normalize_content(chunk)
        if normalized:
            units.append(normalized)
    return Counter(units)


def heading_similarity(first: str, second: str) -> float:
    return SequenceMatcher(
        None, normalize_heading(first), normalize_heading(second)
    ).ratio()


def match_sections(old_sections, new_sections):
    """Сопоставляет разделы по точным, затем по похожим заголовкам."""
    old_keys, new_keys = list(old_sections), list(new_sections)
    unmatched_old, unmatched_new = set(old_keys), set(new_keys)
    matches = []

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


def compare_documents(old_sections, new_sections) -> list[str]:
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
        old_text = split_content_units(old_sections[old_key]["text"])
        new_text = split_content_units(new_sections[new_key]["text"])
        if old_text == new_text:
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
        differences = compare_documents(old_sections, new_sections)
        print(json.dumps(build_result(differences), ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError, KeyError) as exc:
        print(
            json.dumps({"error": str(exc)}, ensure_ascii=False, indent=2),
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
