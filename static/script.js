const productSelect = document.getElementById("productSelect");
const categorySelect = document.getElementById("categorySelect");
const searchBtn = document.getElementById("searchBtn");
const statusText = document.getElementById("statusText");
const resultsBody = document.getElementById("resultsBody");

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}

function renderRows(rows) {
    if (!rows.length) {
        resultsBody.innerHTML = "<tr><td colspan='5'>Geen resultaten gevonden.</td></tr>";
        return;
    }

    resultsBody.innerHTML = rows
        .map((row) => {
            return `
                <tr>
                    <td>${escapeHtml(row.document_name)}</td>
                    <td>${escapeHtml(row.chapter_title)}</td>
                    <td>${escapeHtml(row.chapter_preview)}</td>
                    <td>${escapeHtml(row.classification_group)}</td>
                    <td>${escapeHtml(row.reason)}</td>
                </tr>
            `;
        })
        .join("");
}

async function runSearch() {
    const product = productSelect.value.trim();
    const category = categorySelect.value.trim();

    if (!product) {
        statusText.textContent = "Selecteer eerst een product.";
        return;
    }

    if (!category) {
        statusText.textContent = "Selecteer eerst een categorie.";
        return;
    }

    statusText.textContent = "Bezig met zoeken en classificeren...";
    resultsBody.innerHTML = "<tr><td colspan='5'>Laden...</td></tr>";

    try {
        const response = await fetch("/search", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({ product, category }),
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error || "Onbekende fout");
        }

        let statusMsg = `Bestanden: ${data.file_count} | Hoofdstukken: ${data.chapter_count} | Matches: ${data.match_count}`;

        if (data.file_count === 0) {
            statusMsg += " ⚠️  Geen SD-bestanden gevonden voor dit product.";
        } else if (data.chapter_count === 0) {
            statusMsg += " ⚠️  Geen Heading-gebaseerde hoofdstukken gevonden in de documenten.";
        } else if (data.match_count === 0) {
            statusMsg += " ⚠️  Geen hoofdstukken geclassificeerd onder deze categorie.";
        }

        statusText.textContent = statusMsg;
        console.log("[SD Classifier] response:", data);
        renderRows(data.results || []);
    } catch (error) {
        statusText.textContent = `Fout: ${error.message}`;
        resultsBody.innerHTML = "<tr><td colspan='5'>Er trad een fout op.</td></tr>";
    }
}

searchBtn.addEventListener("click", runSearch);
