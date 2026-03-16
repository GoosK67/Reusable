Dit project automatiseert het genereren van Presales Guides op basis van Cegeka‑Service Descriptions (SD’s).
Het systeem gebruikt pure Python (zonder LLM) om:

SD‑documenten (DOCX) in te lezen
headings, tekstblokken en tabellen te extraheren
relevante SD‑secties te mappen op Presales Guide velden
een standaard presales template (DOCX) automatisch in te vullen
voor 1 of 175 documenten tegelijk guides te genereren

Dit project is ideaal voor het batchgewijs genereren van grote aantallen presales guides.

📁 Projectstructuur
presales-generator/
│
├── input/                # Hier plaats je SD bestanden (.docx)
├── output/               # Hier komen de gegenereerde Presales Guides
├── templates/
│   └── presales_template.docx
│
├── extractor.py          # Extractie van headings, tekst en tabellen uit SD
├── mapper.py             # Mapping SD → Presales Guide velden
├── generator.py          # Template invulling (python-docx)
├── main.py               # Batch runner
│
├── requirements.txt
└── README.md


🚀 Functionaliteit
✔ Extractor (zonder LLM)
Gebruikt python-docx om:

hoofdstukken te detecteren
tekst onder vaste headings te verzamelen
SKU‑tabellen & SLA‑tabellen te herkennen
alle output naar een JSON‑achtig Python dict te structureren

✔ Mapping Engine
Zet SD‑secties om naar de velden van de Presales Guide:

Product Summary
Value Proposition
Product Description
Scope / Requirements
SLA
SKU’s
Client Responsibilities
…

✔ Template Generator
Vervangt placeholders zoals:
<PRODUCT_SUMMARY>
<VALUE_PROP>
<SKUS>
<OPS_SUPPORT>

met gegenereerde inhoud.

🔧 Installatie
1. Python installeren
Zorg dat je Python 3.10+ geïnstalleerd hebt.
2. Dependencies installeren
pip install -r requirements.txt

3. Input en template voorzien

plaats SD (.docx) bestanden in /input
plaats het presales template in /templates/presales_template.docx


▶ Hoe gebruik je het systeem?
Voer gewoon dit uit:
python main.py

Het script:

leest alle DOCX‑bestanden in /input
runt extractie → mapping → template generation
zet de finale Presales Guides in /output


🧩 Requirements (requirements.txt)
python-docx

(Meer libraries kunnen worden toegevoegd wanneer het project uitbreidt.)

🧪 Testing
Je kunt test‑SD’s in /input plaatsen en main.py uitvoeren om te valideren dat:

extractie correcte secties bevat
SKU‑tabellen goed omgezet worden
mapping coherent is
template correct gevuld wordt


💡 Tips voor gebruik in Visual Studio
Nieuwe Git repository maken
Je kunt in Visual Studio een Git‑repo aanmaken via:
Git → Create Git Repository
 [learn.microsoft.com]

Kies GitHub of lokale opslag
Configureer eventueel .gitignore
Klik Create and Push om naar GitHub te publiceren
 [learn.microsoft.com]


🧭 Roadmap

🔄 Optioneel: LLM‑extractor integratie
🖥 HITL Review UI (Streamlit)
🗂 SharePoint export pipeline
🧱 JSON‑schema validatie
🪄 Template extensies


📜 Licentie
Interne Cegeka engineering code – niet bestemd voor externe distributie.