#!/usr/bin/env python
"""Deep-extraction presales guide generator for IBM Power On Premise SD."""

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from difflib import SequenceMatcher
import os

SD_PATH = (
    r"C:\Users\koengo\Cegeka\Product Management - Product Management Library"
    r"\Business Line - Cloud and Digital Platforms"
    r"\[0.1] Cegeka IBM Power Services & Solutions"
    r"\SD - IBM Power on Premise [DV0.9].docx"
)

RULES = {
    "Product Summary":                ["service summary","introduction","overview","summary"],
    "Product Description":            ["standard services","description","approach"],
    "Key Features & Functionalities": ["features","functionalities","capabilities"],
    "Scope / Out-of-Scope":           ["out of scope","scope","excluded"],
    "Requirements & Prerequisites":   ["eligibility","prerequisites","conditions","pre-requisite"],
    "Operational Support":            ["operational services","incident management","operations"],
    "Terms & Conditions":             ["terms","governance","contractual","conditions"],
    "SLA & KPI Management":           ["service support","sla","service level agreement"],
    "Pricing Elements":               ["order","billing","sku","pricing"],
    "Client Responsibilities":        ["responsibility matrix","raci","responsible"],
}

def hlevel(p):
    n = p.style.name if p.style else ""
    if n.startswith("Heading"):
        parts = n.split()
        return int(parts[1]) if len(parts)==2 and parts[1].isdigit() else 1
    return 0

def tbl_text(t):
    rows=[]
    for row in t.rows:
        u=[]
        for c in row.cells:
            s=c.text.strip()
            if not u or s!=u[-1]: u.append(s)
        line=" | ".join(x for x in u if x)
        if line: rows.append(line)
    return "\n".join(rows)

def extract(path):
    doc=Document(path)
    secs,stack,chunks={},{},{}
    for child in doc.element.body.iterchildren():
        if isinstance(child,CT_P):
            p=Paragraph(child,doc); text=p.text.strip(); lvl=hlevel(p)
            if lvl in (1,2,3) and text:
                stack[lvl]=text
                for k in list(stack):
                    if k>lvl: del stack[k]
                secs[text]={"level":lvl,"parent":stack.get(lvl-1,""),"section_text":""}
                chunks[text]=[]
                continue
            cur=stack.get(3) or stack.get(2) or stack.get(1)
            if cur and text: chunks[cur].append(text)
        elif isinstance(child,CT_Tbl):
            t=Table(child,doc); cur=stack.get(3) or stack.get(2) or stack.get(1)
            if cur:
                tt=tbl_text(t)
                if tt: chunks[cur].append("[TABLE]\n"+tt)
    for title,cc in chunks.items():
        if title in secs: secs[title]["section_text"]="\n".join(cc)

    # Aggregate H3 children into empty H2 parents so the parent can be used for mapping
    for title, v in list(secs.items()):
        if v["level"] == 2 and not v["section_text"].strip():
            children_text = "\n\n".join(
                f"### {child_t}\n{child_v['section_text']}"
                for child_t, child_v in secs.items()
                if child_v["level"] == 3
                and child_v["parent"] == title
                and child_v["section_text"].strip()
            )
            if children_text:
                secs[title]["section_text"] = children_text

    return secs

def norm(t): return " ".join(str(t or "").strip().lower().split())
def score(title,kws,body=""):
    best,nt=0.0,norm(title)
    for kw in kws:
        nk=norm(kw)
        if nk in nt: best=max(best,1.0)
        else: best=max(best,SequenceMatcher(None,nt,nk).ratio())
    if body:
        nb=norm(body)
        for kw in kws:
            nk=norm(kw)
            if nk and nk in nb: best=max(best,0.5 if " " in nk else 0.3)
    return best

def map_secs(secs):
    pairs=[(t,v["section_text"]) for t,v in secs.items()]
    matched={}
    for field,kws in RULES.items():
        bs,bt,bxt=0.0,None,""
        bcs,bct,bcxt=0.0,None,""
        for t,text in pairs:
            s=score(t,kws,body=text)
            if s>=bs: bs,bt,bxt=s,t,text
            if text.strip() and s>=bcs: bcs,bct,bcxt=s,t,text
        if not bxt.strip() and bcs>=0.75: bt,bxt,bs=bct,bcxt,bcs
        if bs>=0.40 and bxt.strip():
            matched[field]={"src":bt,"text":bxt,"score":round(bs,3)}
    return matched

SKIP=["fix me","delete","option 1","option 2","option 3","option 4",
      "one or both","can be choosen","fill in","xxx",
      "for service level","for disaster recovery","this section describes"]
def clean(text):
    return "\n".join(l for l in (x.strip() for x in text.splitlines())
                     if l and len(l)>=10 and not any(p in l.lower() for p in SKIP))

def sd(matched,field,fb):
    if field in matched:
        m=matched[field]; t=clean(m["text"])
        if t.strip(): return f"*Bron SD: {m['src']} (score {m['score']})*\n\n{t}"
    return fb

def guide(secs,matched):
    P=[]
    P.append("""# 1. Presales Instructions & Checks

**Voor gebruik door:** Cegeka Account Managers en Solution Architects

**Pre-sales checklist**
- [ ] Klant heeft interesse bevestigd in on-premise of hybride IBM Power omgeving
- [ ] Due Diligence vragenlijst ingevuld en besproken
- [ ] Account Manager gebrieft over servicegrenzen en SLA-verplichtingen
- [ ] Juridische en procurement-vereisten geidentificeerd
- [ ] Prijsgoedkeuring Sales Management voor deals > 100K EUR ARR
- [ ] IBM licenties en onderhoudsstatus in kaart gebracht
- [ ] Contactpersonen bij klant geidentificeerd (IT, Finance, Business)

**Contacteer voor presentatie**
- Design Authority voor niet-standaard configuraties
- IBM Alliance Manager voor enterprise-pricing trajecten""")

    P.append(sd(matched,"Product Summary","""# 2. Product Summary

Cegeka IBM Power On Premise levert een volledig beheerde IBM Power-infrastructuur
- gehost op locatie bij de klant of in een Cegeka-datacenter - met end-to-end
beheer door gecertificeerde IBM Power-specialisten.

De service stelt organisaties in staat om bedrijfskritische workloads (SAP, IBM i,
AIX, Oracle) te draaien op bewezen IBM Power-hardware met hoge beschikbaarheid,
voorspelbare performantie en een duidelijk SLA-kader - zonder de operationele last
van in-house infrastructuurbeheer.

**Ideaal voor:** financiele instellingen, overheidsinstanties, nutsmaatschappijen
en enterprise-omgevingen met hoge eisen op vlak van betrouwbaarheid, security en
compliance.""").replace("\n","\n"))

    P.append("""# 3. Understanding the Client Needs

**Identificeer de nood van de klant voor de presentatie.**

**Zakelijke drivers**
- Welke workloads vereisen hoge performance? (ERP, core banking, SAP, IBM i)
- Zijn er compliance-vereisten die publieke cloud verhinderen?
- Wil de klant consolideren of verouderde hardware vervangen?
- TCO-doelstellingen over een horizon van 3-5 jaar?
- Open voor fully managed model, of volledige controle?

**Ontdekkingsvragen**
- Huidige serveromgeving? (IBM AIX / IBM i / Linux on Power / mix)
- Hoe bedrijfskritisch zijn de workloads? (RTO/RPO-eisen)
- Wie beheert de huidige omgeving? (intern / derde partij / gemengd)
- Bijhorende IBM-licenties en onderhoudscontracten?
- Wanneer loopt het huidige hardware-contract af?""")

    P.append("""# 4. Product Description""")

    arch_fb="""## 4.1 Architectural Description

De IBM Power On Premise architectuur is gebaseerd op dedicated, single-tenant
IBM Power-hardware (POWER9/POWER10) geinstalleerd op een door de klant of Cegeka
aangewezen locatie.

**Componenten**
- IBM Power-server (POWER9 of POWER10 afhankelijk van workload)
- Operating system: IBM AIX, IBM i, of Linux on Power
- Storage: intern of verbonden via Fibre Channel / iSCSI
- Netwerk: klant-beheerd of optioneel via Cegeka Network Services
- Management: IBM HMC, Cegeka remote management tooling + monitoring

**Beheermodel**
Cegeka NOC beheert de hardware- en OS-laag via een beveiligde managementverbinding
(VPN/MPLS). Alle activiteiten zijn traceerbaar via maandelijkse rapportage."""
    P.append(sd(matched,"Product Description",arch_fb))

    feat_fb="""## 4.2 Key Features & Functionalities

- **Volledig beheerd IBM Power platform** - hardware, OS, patching, beschikbaarheid
- **24/7 proactieve monitoring** - incidentdetectie voor klantimpact via Dynatrace
- **Incident & Problem Management** - ITIL-conform, gedefinieerde respons- en resolutietijden
- **Patch & Lifecycle Management** - gestructureerde patchcyclus met klantcommunicatie
- **Capacity Management** - periodieke capaciteitsanalyse en uitbreidingsadvies
- **Security Hardening** - OS-hardening, vulnerability scanning, compliancerapportage
- **Change Management** - gestructureerd via ServiceNow met klantgoedkeuring
- **Responsibility Matrix** - heldere RACI tussen Cegeka en klant
- **Maandelijkse rapportage** - SLA-prestaties, incidentoverzicht, capaciteitsevolutie"""
    P.append(sd(matched,"Key Features & Functionalities",feat_fb))

    scope_fb="""## 4.3 Scope / Out-of-Scope

Alles wat niet expliciet beschreven is in deze service description valt standaard
buiten scope. De service omvat de hardware- en OS-laag. Applicatiebeheer,
end-user support en netwerkinfrastructuur vallen buiten scope tenzij expliciet
gecontracteerd."""
    P.append(sd(matched,"Scope / Out-of-Scope",scope_fb))

    req_fb="""## 4.4 Requirements & Prerequisites

Een beveiligde managementverbinding (VPN of dedicated link) tussen de klant en het
Cegeka NOC is vereist. IBM hardware dient onder een geldig IBM-onderhoudscontract
te vallen. De klant stelt een benoemde contactpersoon beschikbaar voor change-
goedkeuring en escalaties."""
    P.append(sd(matched,"Requirements & Prerequisites",req_fb))

    vp_fb="""# 5. Value Proposition

Cegeka IBM Power On Premise biedt enterprise-grade infrastructuur zonder de
complexiteit van in-house beheer.

**Kernvoordelen**
- **Verlaagde operationele kosten** - geen nood aan dure IBM Power-specialisten intern
- **Voorspelbaar OPEX-model** - vaste maandelijkse kosten afgestemd op capaciteit
- **Hogere beschikbaarheid** - SLA-backed service met proactieve incidentpreventie
- **Snellere probleemresolutie** - 24/7 NOC-expertresponse
- **Schaalbare capaciteit** - hardware uitbreidbaar zonder operationele onderbreking
- **Compliance-ready** - ISO 27001 en sectorspecifieke compliancevereisten
- **Focus op kernactiviteiten** - IT-afdeling richt zich op business value"""
    P.append(sd(matched,"Value Proposition",vp_fb))

    diff_fb="""# 6. Key Differentiators

- **Gecertificeerd IBM Business Partner** - hoogste IBM Partner-status, dedicated Power-specialisten
- **Multi-generatie IBM Power expertise** - POWER7 tot POWER10, IBM i en AIX
- **Hybride cloud enablement** - IBM Power on-premise koppelen aan Cegeka Cloud of public cloud
- **Pan-Europese delivery** - IBM Power managed services in BE, NL, DE, CZ, SK, RO
- **ITSM-integratie** - directe koppeling met ServiceNow
- **Bewezen referentiebase** - klantenportfolio in finance, utilities en publieke sector
- **IBM i specialisatie** - zeldzame expertise in IBM i (AS/400) omgevingen"""
    P.append(sd(matched,"Key Differentiators",diff_fb))

    P.append("""# 7. Transition & Transformation

**Transitiefasen**
1. **Assessment & Design** (2-3 weken) - as-is inventarisatie, capaciteitsplanning en serviceontwerp
2. **Migratieplan** (1 week) - mijlpalen, eigenaarschap en rollback-procedures
3. **Infrastructuurbouw** (2-4 weken) - hardware gestagd, geconfigureerd en gevalideerd
4. **Hypercare** (30 dagen) - intensieve fase na go-live met dedicated engineering support
5. **Steady-state** - overdracht aan Cegeka Managed Services met gedocumenteerde runbooks

**Typische looptijd:** 6-12 weken afhankelijk van complexiteit""")

    cr_fb="""# 8. Client Responsibilities

Om de gecontracteerde serviceniveaus te garanderen, verbindt de klant zich ertoe:

- Tijdige toegang bieden tot hardware, datacenterfaciliteiten en netwerkverbindingen
- Actuele contactlijst bijhouden voor incidentescalatie en wijzigingsgoedkeuring
- Change requests goedkeuren of weigeren binnen afgesproken termijnen
- IBM hardware-onderhoudscontracten in stand houden
- Cegeka informeren over geplande maintenance windows of infrastructuurwijzigingen
- Geldige softwarelicenties onderhouden voor alle software op beheerde systemen"""
    P.append(sd(matched,"Client Responsibilities",cr_fb))

    ops_fb="""# 9. Operational Support

Cegeka biedt 24/7/365 operationele support via haar gecentraliseerde NOC.
Alle IBM Power omgevingen worden proactief gemonitord. Incidenten worden beheerd
via ITIL-conform proces en geintegreerd met de klant-ITSM."""
    P.append(sd(matched,"Operational Support",ops_fb))

    tac_fb="""# 10. Terms & Conditions

Deze service valt onder de Cegeka Algemene Voorwaarden en bijbehorende SLA.

**Commerciele kernvoorwaarden**
- **Minimale contractduur:** 36 maanden (standaard); 12 maanden op aanvraag
- **Opzegtermijn:** 6 maanden voor aflopen contractperiode
- **Indexering:** jaarlijkse aanpassing conform AGORIA / CPI
- **Wijzigingsverzoeken:** via formeel change management, kunnen prijsimpact hebben
- **Aansprakelijkheid:** beperkt tot 12 maanden servicefees"""
    P.append(sd(matched,"Terms & Conditions",tac_fb))

    sla_fb="""# 11. SLA & KPI Management

| KPI | Standaard | Premium |
|-----|-----------|---------|
| Beschikbaarheid | 99.5% | 99.9% |
| Incident P1 response | 30 min | 15 min |
| Incident P1 resolutie | 4 uur | 2 uur |
| Change lead time | 5 werkdagen | 2 werkdagen |
| Maandelijkse rapportage | Inbegrepen | Inbegrepen |

SLA-credits worden toegekend bij overschrijding van gecommitteerde serviceniveaus."""
    P.append(sd(matched,"SLA & KPI Management",sla_fb))

    price_fb="""# 12. Pricing Elements

**Basis servicefee** - maandelijks terugkerende kost voor hardware-/OS-beheer, monitoring en SLA

**Optionele uitbreidingen**
- Backup as a Service (BaaS)
- Disaster Recovery as a Service (DRaaS)
- Extended monitoring en APM
- Security Services (vulnerability scanning, compliance)

**Eenmalige kosten**
- Initieel setup en transitiefee
- Maatwerkkoppelingen (ITSM, netwerk)

Contacteer uw Cegeka Account Manager voor een gepersonaliseerde offerte."""
    P.append(sd(matched,"Pricing Elements",price_fb))

    return "\n\n---\n\n".join(P)


# == Run ==
print("Extracting SD (H1/H2/H3 deep)...")
secs=extract(SD_PATH)
print(f"  {len(secs)} sections found")
for t,v in secs.items():
    print(f"  H{v['level']}  {t!r:<50} ({len(v['section_text'])} chars)")

print("\nMapping...")
matched=map_secs(secs)
for f,m in matched.items():
    print(f"  OK  {f:<40} <- {m['src']!r}")
for f in RULES:
    if f not in matched: print(f"  --  {f:<40} <- fallback")

g=guide(secs,matched)
os.makedirs("output",exist_ok=True)
out="output/IBM Power On Premise - Presales Guide.md"
with open(out,"w",encoding="utf-8") as f:
    f.write("# Cegeka Presales Guide - IBM Power On Premise\n\n")
    f.write("> Gegenereerd: 16 maart 2026 | Bron: SD - IBM Power on Premise [DV0.9]\n\n---\n\n")
    f.write(g)
print(f"\nSaved: {out}\n\n{'='*70}\n")
print(g)
