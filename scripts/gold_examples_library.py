import argparse
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile
import hashlib

from lxml import etree

BASE = Path(__file__).resolve().parents[1]
LIB_PATH = BASE / "rules" / "gold_examples.json"
LOG_DIR = BASE / "log"
DOCX_OUT_DIR = BASE / "output" / "docx"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

HITL_PREFIX = "AI generated, teverifieren door HITL"
LOW_INFO_TEXT = "AI agent heeft te weinig info om dit zelf op te stellen"

SUPPORTED_TAGS = [
    "PRODUCT_SUMMARY",
    "CLIENT_NEEDS",
    "PRODUCT_DESCRIPTION",
    "ARCHITECTURAL_DESCRIPTION",
    "KEY_FEATURES",
    "SCOPE",
    "REQUIREMENTS",
    "VALUE_PROPOSITION",
    "DIFFERENTIATORS",
    "TRANSITION_TRANSFORMATION",
    "CLIENT_RESPONSIBILITIES",
    "OPERATIONAL_SUPPORT",
    "TERMS_CONDITIONS",
    "SLA_KPI",
    "PRICING_ELEMENTS",
]


def _load_lib():
    if not LIB_PATH.exists():
        payload = {
            "version": 1,
            "policy": "Style anchors only. Never use these examples as factual source evidence.",
            "examples": {tag: [] for tag in SUPPORTED_TAGS},
        }
        _save_lib(payload)
        return payload

    return json.loads(LIB_PATH.read_text(encoding="utf-8", errors="ignore"))


def _save_lib(payload):
    LIB_PATH.parent.mkdir(parents=True, exist_ok=True)
    LIB_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def _looks_generated(text):
    t = str(text or "")
    return HITL_PREFIX in t or LOW_INFO_TEXT in t


def _extract_sdt_text(docx_path, tag_name):
    with ZipFile(docx_path, "r") as z:
        xml = etree.fromstring(z.read("word/document.xml"))

    sdts = xml.xpath(f".//w:sdt[w:sdtPr/w:tag[@w:val='{tag_name}']]", namespaces=NS)
    if not sdts:
        return ""

    content = sdts[0].find("w:sdtContent", NS)
    if content is None:
        return ""

    texts = content.xpath(".//w:t/text()", namespaces=NS)
    return "\n".join(t.strip() for t in texts if t and t.strip()).strip()


def _normalize_id(raw):
    value = re.sub(r"[^a-zA-Z0-9_-]+", "_", str(raw or "").strip())
    return value.strip("_") or f"example_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def _text_hash(text):
    return hashlib.sha256(str(text or "").encode("utf-8", errors="ignore")).hexdigest()[:16]


def _find_latest_output_docx(sd_stem):
    stem = str(sd_stem or "").strip()
    if not stem:
        return None

    candidates = []
    for p in DOCX_OUT_DIR.glob("*_FINAL*.docx"):
        if p.stem.startswith(stem):
            candidates.append(p)

    if not candidates:
        fallback = f"{stem}_FINAL.docx"
        fallback_path = DOCX_OUT_DIR / fallback
        if fallback_path.exists():
            return fallback_path
        return None

    return max(candidates, key=lambda x: x.stat().st_mtime)


def _parse_quality_rows_from_mapped_log(log_path):
    rows = []
    if not log_path.exists():
        return rows

    lines = log_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    starts = [i for i, line in enumerate(lines) if "START xml_to_docx" in line]
    scoped = lines[starts[-1]:] if starts else lines

    pat = (
        r"Quality SDT '([^']+)': overall=(\d+); coverage=(\d+); specificity=(\d+); "
        r"evidence_count=(\d+); policy_compliance=(\d+); fill_type=([^;]+); low_score=(yes|no)"
    )

    for line in scoped:
        m = re.search(pat, line)
        if not m:
            continue
        rows.append(
            {
                "tag": m.group(1).strip(),
                "overall": int(m.group(2)),
                "coverage": int(m.group(3)),
                "specificity": int(m.group(4)),
                "evidence_count": int(m.group(5)),
                "policy_compliance": int(m.group(6)),
                "fill_type": m.group(7).strip(),
                "low_score": m.group(8).strip().lower() == "yes",
            }
        )

    return rows


def cmd_list(_args):
    lib = _load_lib()
    examples = lib.get("examples", {})
    print(f"Library: {LIB_PATH}")
    for tag in SUPPORTED_TAGS:
        count = len(examples.get(tag, [])) if isinstance(examples.get(tag), list) else 0
        print(f"- {tag}: {count}")


def cmd_add_from_docx(args):
    tag = args.tag.strip().upper()
    if tag not in SUPPORTED_TAGS:
        raise ValueError(f"Unsupported tag: {tag}")

    docx_path = Path(args.docx)
    if not docx_path.exists():
        raise FileNotFoundError(f"Missing DOCX: {docx_path}")

    text = _extract_sdt_text(docx_path, args.source_tag or tag)
    if not text:
        raise ValueError(f"No SDT content found for tag '{args.source_tag or tag}' in {docx_path}")

    if _looks_generated(text):
        raise ValueError("Refusing generated fallback text. Only approved factual chapter samples are allowed.")

    lib = _load_lib()
    examples = lib.setdefault("examples", {})
    bucket = examples.setdefault(tag, [])
    if not isinstance(bucket, list):
        bucket = []
        examples[tag] = bucket

    entry = {
        "id": _normalize_id(args.example_id or f"{docx_path.stem}_{tag}"),
        "status": "approved",
        "source_doc": docx_path.name,
        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "sample_text": text,
    }
    bucket.append(entry)
    _save_lib(lib)
    print(f"Added approved example for {tag}: {entry['id']}")


def cmd_add_text(args):
    tag = args.tag.strip().upper()
    if tag not in SUPPORTED_TAGS:
        raise ValueError(f"Unsupported tag: {tag}")

    text_path = Path(args.text_file)
    if not text_path.exists():
        raise FileNotFoundError(f"Missing text file: {text_path}")

    text = text_path.read_text(encoding="utf-8", errors="ignore").strip()
    if not text:
        raise ValueError("Text file is empty")

    if _looks_generated(text):
        raise ValueError("Refusing generated fallback text. Only approved chapter samples are allowed.")

    lib = _load_lib()
    examples = lib.setdefault("examples", {})
    bucket = examples.setdefault(tag, [])
    if not isinstance(bucket, list):
        bucket = []
        examples[tag] = bucket

    entry = {
        "id": _normalize_id(args.example_id or f"{text_path.stem}_{tag}"),
        "status": "approved",
        "source_doc": text_path.name,
        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "sample_text": text,
    }
    bucket.append(entry)
    _save_lib(lib)
    print(f"Added approved example for {tag}: {entry['id']}")


def cmd_seed_from_logs(args):
    lib = _load_lib()
    examples = lib.setdefault("examples", {})

    min_score = int(args.min_score)
    max_per_tag = int(args.max_per_tag)
    require_direct = not bool(args.allow_non_direct)
    require_non_low = not bool(args.allow_low_score)

    existing_hashes = {tag: set() for tag in SUPPORTED_TAGS}
    for tag in SUPPORTED_TAGS:
        bucket = examples.get(tag, [])
        if not isinstance(bucket, list):
            continue
        for entry in bucket:
            if not isinstance(entry, dict):
                continue
            txt = str(entry.get("sample_text", "") or "").strip()
            if txt:
                existing_hashes[tag].add(_text_hash(txt))

    log_files = sorted(LOG_DIR.glob("*_mapped.log"), key=lambda p: p.stat().st_mtime, reverse=True)
    added = {tag: 0 for tag in SUPPORTED_TAGS}

    for mapped_log in log_files:
        sd_stem = mapped_log.stem
        rows = _parse_quality_rows_from_mapped_log(mapped_log)
        if not rows:
            continue

        docx_path = _find_latest_output_docx(sd_stem)
        if not docx_path or not docx_path.exists():
            continue

        for row in rows:
            tag = row.get("tag", "")
            if tag not in SUPPORTED_TAGS:
                continue
            if added[tag] >= max_per_tag:
                continue
            if row.get("overall", 0) < min_score:
                continue
            if require_direct and row.get("fill_type") != "direct_from_sd_chapter":
                continue
            if require_non_low and row.get("low_score", False):
                continue

            text = _extract_sdt_text(docx_path, tag)
            if not text:
                continue
            if _looks_generated(text):
                continue

            h = _text_hash(text)
            if h in existing_hashes[tag]:
                continue

            bucket = examples.setdefault(tag, [])
            if not isinstance(bucket, list):
                bucket = []
                examples[tag] = bucket

            entry = {
                "id": _normalize_id(f"seed_{docx_path.stem}_{tag}_{h}"),
                "status": "approved",
                "source_doc": docx_path.name,
                "seed_source": mapped_log.name,
                "quality_overall": row.get("overall", 0),
                "quality_fill_type": row.get("fill_type", ""),
                "created_at": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
                "sample_text": text,
            }
            bucket.append(entry)
            existing_hashes[tag].add(h)
            added[tag] += 1

        if all(added[t] >= max_per_tag for t in SUPPORTED_TAGS):
            break

    _save_lib(lib)

    total_added = sum(added.values())
    print(f"Seed complete. Added {total_added} example(s).")
    for tag in SUPPORTED_TAGS:
        if added[tag] > 0:
            print(f"- {tag}: +{added[tag]}")


def build_parser():
    parser = argparse.ArgumentParser(description="Manage reusable gold examples library per SDT tag.")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="List approved example counts per tag")
    p_list.set_defaults(func=cmd_list)

    p_add_docx = sub.add_parser("add-from-docx", help="Add approved sample from a filled SDT DOCX")
    p_add_docx.add_argument("--docx", required=True, help="Path to approved filled DOCX")
    p_add_docx.add_argument("--tag", required=True, help="Target SDT category tag")
    p_add_docx.add_argument("--source-tag", required=False, help="SDT tag in DOCX if different")
    p_add_docx.add_argument("--example-id", required=False, help="Optional stable ID")
    p_add_docx.set_defaults(func=cmd_add_from_docx)

    p_add_text = sub.add_parser("add-text", help="Add approved sample from plain text file")
    p_add_text.add_argument("--text-file", required=True, help="Path to approved text sample")
    p_add_text.add_argument("--tag", required=True, help="Target SDT category tag")
    p_add_text.add_argument("--example-id", required=False, help="Optional stable ID")
    p_add_text.set_defaults(func=cmd_add_text)

    p_seed = sub.add_parser("seed-from-logs", help="Seed library from high-quality direct SDT chapters in mapped logs")
    p_seed.add_argument("--min-score", type=int, default=80, help="Minimum overall quality score (default: 80)")
    p_seed.add_argument("--max-per-tag", type=int, default=3, help="Maximum new examples per tag per run (default: 3)")
    p_seed.add_argument("--allow-non-direct", action="store_true", help="Allow non-direct fill types as seed candidates")
    p_seed.add_argument("--allow-low-score", action="store_true", help="Allow low-score chapters as seed candidates")
    p_seed.set_defaults(func=cmd_seed_from_logs)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
