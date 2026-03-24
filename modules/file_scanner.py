import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set


ROOT_DIR = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")


def _canonical(value: str) -> str:
    return "".join(ch for ch in value.lower() if ch.isalnum())


def _clean_name(raw: str) -> str:
    text = re.sub(r"\s+", " ", raw.replace("_", " ").replace("-", " ")).strip()
    return text


def _extract_product_from_filename(path: Path) -> Optional[str]:
    stem = path.stem.strip()
    stem_lower = stem.lower()

    if not stem_lower.startswith("sd"):
        return None

    # Remove leading SD markers like "SD -", "SD_", "SD ".
    candidate = re.sub(r"^sd\s*[-_]?\s*", "", stem, flags=re.IGNORECASE)
    candidate = candidate.split("[")[0].strip()
    candidate = _clean_name(candidate)

    return candidate if candidate else None


def _extract_product_candidates(path: Path) -> Set[str]:
    candidates: Set[str] = set()

    from_filename = _extract_product_from_filename(path)
    if from_filename:
        candidates.add(from_filename)

    if path.parent != ROOT_DIR:
        folder_name = _clean_name(path.parent.name)
        if folder_name:
            candidates.add(folder_name)

    return {name for name in candidates if name}


def _iter_sd_docx_files(root: Path) -> Iterable[Path]:
    for docx_path in root.rglob("SD*.docx"):
        if docx_path.is_file():
            yield docx_path


def find_all_products() -> List[str]:
    """Scan the root folder and return unique detected product names."""
    products_by_key: Dict[str, str] = {}

    for docx_path in _iter_sd_docx_files(ROOT_DIR):
        for candidate in _extract_product_candidates(docx_path):
            key = _canonical(candidate)
            if key and key not in products_by_key:
                products_by_key[key] = candidate

    return sorted(products_by_key.values(), key=lambda x: x.lower())


def find_sd_files_for_product(product_name: str) -> List[Path]:
    """Return SD*.docx files that belong to the selected product."""
    selected_key = _canonical(product_name)
    matched: List[Path] = []

    if not selected_key:
        return matched

    for docx_path in _iter_sd_docx_files(ROOT_DIR):
        candidate_keys = {_canonical(name) for name in _extract_product_candidates(docx_path)}
        if selected_key in candidate_keys:
            matched.append(docx_path)

    return sorted(matched)
