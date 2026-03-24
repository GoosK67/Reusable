import logging
from typing import List

from flask import Flask, jsonify, render_template, request

from modules.chapter_extractor import extract_chapters
from modules.classifier import classify_with_ollama
from modules.file_scanner import find_all_products, find_sd_files_for_product


CATEGORIES: List[str] = [
    "Executive Summary & Product Overview",
    "Scope Boundaries & Prerequisites",
    "Transition Operations & Governance",
    "Commercial & Risk Management",
    "Internal Presales Alignment",
    "Future Category 6",
    "Future Category 7",
]

app = Flask(__name__)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


@app.route("/")
def index():
    products = find_all_products()
    return render_template("index.html", products=products)


@app.route("/search", methods=["POST"])
def search():
    payload = request.get_json(silent=True) or {}
    selected_product = str(payload.get("product", "")).strip()
    selected_category = str(payload.get("category", "")).strip()

    if not selected_product:
        return jsonify({"error": "Product is verplicht."}), 400

    if selected_category not in CATEGORIES:
        return jsonify({"error": "Ongeldige categorie."}), 400

    sd_files = find_sd_files_for_product(selected_product)
    logging.info(
        "Search request | product=%s | category=%s | files=%d",
        selected_product,
        selected_category,
        len(sd_files),
    )

    matched_rows = []
    total_chapters = 0

    for doc_path in sd_files:
        try:
            chapters = extract_chapters(doc_path)
        except Exception as exc:
            logging.error("Skipping corrupt/unreadable document: %s | %s", doc_path, exc)
            continue

        for chapter in chapters:
            total_chapters += 1
            classification = classify_with_ollama(chapter["title"], chapter["text"])

            if classification["group"] == selected_category:
                matched_rows.append(
                    {
                        "document_name": doc_path.name,
                        "chapter_title": chapter["title"],
                        "chapter_preview": chapter["text"][:300],
                        "classification_group": classification["group"],
                        "reason": classification["reason"],
                    }
                )

    return jsonify(
        {
            "selected_product": selected_product,
            "selected_category": selected_category,
            "file_count": len(sd_files),
            "chapter_count": total_chapters,
            "match_count": len(matched_rows),
            "results": matched_rows,
        }
    )


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
