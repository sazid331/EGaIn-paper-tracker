#!/usr/bin/env python3
"""Search OpenAlex for recent EGaIn-related papers and update a monthly Excel tracker."""

import csv
import json
import re
import requests
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

ROOT_DIR = Path(__file__).parent
REPORTS_DIR = ROOT_DIR / "reports"
SEARCH_DAYS = 8
OPENALEX_API = "https://api.openalex.org/works"

SEARCH_QUERIES = [
    "EGaIn",
    "eutectic gallium indium",
    "eutectic gallium-indium",
    "EGaIn liquid metal",
    "EGaIn soft electronics",
    "EGaIn microfluidics",
    "EGaIn stretchable",
    "EGaIn sensor",
    "EGaIn thermal interface",
    "gallium indium eutectic alloy",
]

RELEVANCE_TERMS = [
    "egain", "eutectic gallium", "gallium.indium", "ga.in eutectic",
    "galinstan", "liquid metal alloy", "liquid.metal.sensor",
    "liquid.metal.stretchable", "liquid.metal.soft",
    "liquid.metal.flexible", "liquid.metal.electrode",
    "liquid.metal.thermal", "liquid.metal.microfluid",
    "liquid.metal.droplet", "liquid.metal.circuit",
    "liquid.metal.antenna", "liquid.metal.reconfigurable",
]

TOPIC_KEYWORDS = {
    "Sensors & Wearables": [
        "sensor", "wearable", "strain sens", "pressure sens", "motion detect",
        "gesture", "health monitor", "bioelectron", "e-skin", "electronic skin",
    ],
    "Soft Robotics": [
        "soft robot", "actuator", "pneumatic", "gripper", "locomotion",
        "shape morphing", "soft machine",
    ],
    "Thermal Management": [
        "thermal interface", "thermal conduct", "heat transfer", "heat sink",
        "cooling", "thermal management", "thermal pad",
    ],
    "Microfluidics": [
        "microfluid", "microchannel", "droplet", "lab.on.chip", "fluidic",
        "micropump", "valve",
    ],
    "Electronics & Circuits": [
        "circuit", "antenna", "reconfigurable", "interconnect", "electrode",
        "conductive", "wire", "transistor", "switch",
    ],
    "Stretchable & Flexible Electronics": [
        "stretchable", "flexible electron", "deformable", "elastic",
        "conformable", "textile", "fiber",
    ],
    "Biomedical": [
        "biomedical", "drug delivery", "implant", "biosens", "wound",
        "tissue", "cell", "antimicrobial", "therap", "diagnos",
    ],
    "Surface & Interface Science": [
        "surface", "oxide", "wetting", "interfac", "self.assembled",
        "thin film", "coating", "corrosion", "oxidation",
    ],
    "Energy": [
        "energy harvest", "battery", "supercapacitor", "triboelectric",
        "solar", "photovoltaic", "fuel cell", "thermoelectric",
    ],
    "Fabrication & Processing": [
        "3d print", "additive manufactur", "patterning", "lithograph",
        "inject", "embed", "encapsulat", "sintering", "composite",
    ],
}

MAILTO = "rssazid331@gmail.com"
CSV_HEADERS = [
    "Title", "Authors", "DOI", "Journal", "Publication Date",
    "Topics", "Summary", "Cited By", "Open Access URL", "Date Added",
]


def get_monthly_paths() -> tuple:
    now = datetime.now()
    base = f"egain_papers_{now.year}_{now.month:02d}"
    return REPORTS_DIR / f"{base}.xlsx", REPORTS_DIR / f"{base}.csv"


def build_openalex_params(query: str, from_date: str) -> dict:
    return {
        "search": query,
        "filter": f"from_publication_date:{from_date}",
        "sort": "publication_date:desc",
        "per_page": 25,
        "mailto": MAILTO,
    }


def reconstruct_abstract(inverted_index: dict) -> str:
    if not inverted_index:
        return ""
    word_positions = []
    for word, positions in inverted_index.items():
        for pos in positions:
            word_positions.append((pos, word))
    word_positions.sort()
    return " ".join(w for _, w in word_positions)


def is_relevant(title: str, abstract: str) -> bool:
    text = (title + " " + abstract).lower()
    for term in RELEVANCE_TERMS:
        if re.search(term, text):
            return True
    return False


def classify_topics(title: str, abstract: str) -> list:
    text = (title + " " + abstract).lower()
    matched = []
    for topic, keywords in TOPIC_KEYWORDS.items():
        for kw in keywords:
            if re.search(kw, text):
                matched.append(topic)
                break
    return matched if matched else ["General / Other"]


def fetch_papers(from_date: str) -> dict:
    seen_ids = set()
    papers = {}

    for query in SEARCH_QUERIES:
        params = build_openalex_params(query, from_date)
        try:
            resp = requests.get(OPENALEX_API, params=params, timeout=30)
            resp.raise_for_status()
            results = resp.json().get("results", [])
        except (requests.RequestException, json.JSONDecodeError):
            continue

        for work in results:
            oa_id = work.get("id", "")
            if oa_id in seen_ids:
                continue
            seen_ids.add(oa_id)

            doi = work.get("doi", "") or ""
            title = work.get("title", "") or ""
            if not title:
                continue

            abstract = reconstruct_abstract(
                work.get("abstract_inverted_index") or {}
            )

            if not is_relevant(title, abstract):
                continue

            authors_list = []
            for authorship in (work.get("authorships") or [])[:5]:
                name = (authorship.get("author") or {}).get("display_name", "")
                if name:
                    authors_list.append(name)
            if len(work.get("authorships") or []) > 5:
                authors_list.append("et al.")
            authors = ", ".join(authors_list)

            source = (work.get("primary_location") or {}).get("source") or {}
            journal = source.get("display_name", "") or ""

            pub_date = work.get("publication_date", "") or ""
            cited_by = work.get("cited_by_count", 0)
            oa_url = work.get("open_access", {}).get("oa_url", "") or ""
            topics = classify_topics(title, abstract)

            papers[oa_id] = {
                "title": title,
                "authors": authors,
                "doi": doi,
                "journal": journal,
                "publication_date": pub_date,
                "abstract": abstract,
                "topics": topics,
                "cited_by_count": cited_by,
                "open_access_url": oa_url,
            }

    return papers


def load_all_existing_dois() -> set:
    dois = set()
    if not REPORTS_DIR.exists():
        return dois
    for f in REPORTS_DIR.glob("egain_papers_*.csv"):
        try:
            with open(f, newline="", encoding="utf-8") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    doi = (row.get("DOI") or "").strip().lower()
                    if doi:
                        dois.add(doi)
        except Exception:
            continue
    import openpyxl
    for f in REPORTS_DIR.glob("egain_papers_*.xlsx"):
        try:
            wb = openpyxl.load_workbook(f)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True):
                if row[0]:
                    dois.add(str(row[0]).strip().lower())
            wb.close()
        except Exception:
            continue
    return dois


def generate_summary(title: str, abstract: str) -> str:
    if not abstract:
        return f"Paper on: {title[:200]}"
    sentences = abstract.replace(". ", ".\n").split("\n")
    key_sentences = []
    for s in sentences:
        s = s.strip()
        if not s:
            continue
        lower = s.lower()
        if any(kw in lower for kw in [
            "we present", "we report", "we demonstrate", "we propose",
            "this study", "this paper", "this work", "here we", "in this",
            "results show", "our results", "we found", "findings",
            "novel", "new approach", "first time",
        ]):
            key_sentences.append(s)
        if len(key_sentences) >= 2:
            break
    if not key_sentences:
        key_sentences = sentences[:2]
    summary = " ".join(key_sentences)
    if len(summary) > 500:
        summary = summary[:497] + "..."
    return summary


def create_or_load_workbook(excel_path: Path):
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    if excel_path.exists():
        return openpyxl.load_workbook(excel_path)

    wb = openpyxl.Workbook()
    ws = wb.active
    now = datetime.now()
    ws.title = f"{now.strftime('%B %Y')}"

    headers = [
        "Title", "Authors", "DOI", "Journal", "Publication Date",
        "Topics", "Abstract", "Summary", "Cited By", "Open Access URL", "Date Added",
    ]

    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    thin_border = Border(bottom=Side(style="thin", color="B4C6E7"))

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    ws.row_dimensions[1].height = 30
    widths = {"A": 65, "B": 45, "C": 45, "D": 35, "E": 18,
              "F": 30, "G": 90, "H": 65, "I": 12, "J": 45, "K": 15}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    wb.save(excel_path)
    return wb


def append_papers(papers: dict, existing_dois: set, excel_path: Path, csv_path: Path) -> tuple:
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    new_papers = []
    for paper in papers.values():
        doi_lower = paper["doi"].strip().lower()
        if doi_lower and doi_lower in existing_dois:
            continue
        new_papers.append(paper)
        if doi_lower:
            existing_dois.add(doi_lower)

    if not new_papers:
        return 0, new_papers

    wb = create_or_load_workbook(excel_path)
    ws = wb.active
    today = datetime.now().strftime("%Y-%m-%d")

    body_font = Font(name="Calibri", size=11)
    title_font = Font(name="Calibri", size=11, bold=True)
    doi_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
    topic_font = Font(name="Calibri", size=10, italic=True, color="2F5496")
    row_border = Border(bottom=Side(style="thin", color="D9E2F3"))
    light_fill = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")

    csv_exists = csv_path.exists()
    csv_file = open(csv_path, "a", newline="", encoding="utf-8")
    writer = csv.writer(csv_file)
    if not csv_exists:
        writer.writerow(CSV_HEADERS)

    for paper in new_papers:
        summary = generate_summary(paper["title"], paper["abstract"])
        topics_str = ", ".join(paper["topics"])

        row_data = [
            paper["title"], paper["authors"], paper["doi"], paper["journal"],
            paper["publication_date"], topics_str, paper["abstract"], summary,
            paper["cited_by_count"], paper["open_access_url"], today,
        ]
        ws.append(row_data)
        row_num = ws.max_row
        use_stripe = (row_num % 2 == 0)
        ws.row_dimensions[row_num].height = 45

        font_map = {1: title_font, 3: doi_font, 6: topic_font}
        for col in range(1, 12):
            cell = ws.cell(row=row_num, column=col)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = row_border
            cell.font = font_map.get(col, body_font)
            if use_stripe:
                cell.fill = light_fill

        writer.writerow([
            paper["title"], paper["authors"], paper["doi"], paper["journal"],
            paper["publication_date"], topics_str, summary,
            paper["cited_by_count"], paper["open_access_url"], today,
        ])

    csv_file.close()
    wb.save(excel_path)
    wb.close()
    return len(new_papers), new_papers


def load_all_papers_for_readme() -> list:
    """Load all papers from all CSV files for the README dashboard."""
    all_papers = []
    if not REPORTS_DIR.exists():
        return all_papers
    for f in sorted(REPORTS_DIR.glob("egain_papers_*.csv")):
        try:
            with open(f, newline="", encoding="utf-8") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    all_papers.append(row)
        except Exception:
            continue
    return all_papers


def generate_readme(new_papers: list):
    all_papers = load_all_papers_for_readme()
    now = datetime.now()

    # Topic distribution across all papers
    topic_counter = Counter()
    for p in all_papers:
        for t in (p.get("Topics") or "General / Other").split(", "):
            topic_counter[t.strip()] += 1

    # Journal distribution
    journal_counter = Counter()
    for p in all_papers:
        j = p.get("Journal", "").strip()
        if j:
            journal_counter[j] += 1

    # Monthly counts
    monthly_counts = Counter()
    for p in all_papers:
        date_str = p.get("Publication Date", "")
        if date_str and len(date_str) >= 7:
            monthly_counts[date_str[:7]] += 1

    # Recent papers (last 20)
    recent = all_papers[-20:]
    recent.reverse()

    # Build README
    lines = []
    lines.append("# EGaIn Research Paper Tracker")
    lines.append("")
    lines.append("> Automated weekly tracker for **eutectic gallium-indium (EGaIn)** research papers.")
    lines.append(f"> Powered by [OpenAlex](https://openalex.org/) | Updated every Friday at 4 AM ET")
    lines.append("")
    lines.append(f"**Last updated:** {now.strftime('%B %d, %Y')}")
    lines.append(f"**Total papers tracked:** {len(all_papers)}")
    lines.append("")

    # Quick stats
    lines.append("---")
    lines.append("")
    lines.append("## At a Glance")
    lines.append("")

    this_month_key = now.strftime("%Y-%m")
    this_month_count = monthly_counts.get(this_month_key, 0)
    lines.append(f"| This Month | All Time | Top Topic | Top Journal |")
    lines.append(f"|:---:|:---:|:---:|:---:|")
    top_topic = topic_counter.most_common(1)[0][0] if topic_counter else "—"
    top_journal = journal_counter.most_common(1)[0][0] if journal_counter else "—"
    lines.append(f"| {this_month_count} papers | {len(all_papers)} papers | {top_topic} | {top_journal} |")
    lines.append("")

    # Research trends by topic
    lines.append("## Research Trends by Topic")
    lines.append("")
    if topic_counter:
        lines.append("| Topic | Papers | Share |")
        lines.append("|:---|:---:|:---:|")
        total = len(all_papers)
        for topic, count in topic_counter.most_common():
            pct = (count / total * 100) if total else 0
            bar = "█" * max(1, round(pct / 5))
            lines.append(f"| {topic} | {count} | {bar} {pct:.0f}% |")
        lines.append("")

    # Monthly activity
    if monthly_counts:
        lines.append("## Monthly Activity")
        lines.append("")
        lines.append("| Month | Papers |")
        lines.append("|:---|:---:|")
        for month_key in sorted(monthly_counts.keys(), reverse=True)[:12]:
            count = monthly_counts[month_key]
            bar = "█" * max(1, count)
            lines.append(f"| {month_key} | {bar} {count} |")
        lines.append("")

    # Top journals
    if journal_counter:
        lines.append("## Top Journals")
        lines.append("")
        lines.append("| Journal | Papers |")
        lines.append("|:---|:---:|")
        for journal, count in journal_counter.most_common(10):
            lines.append(f"| {journal} | {count} |")
        lines.append("")

    # Recent papers
    lines.append("## Recent Papers")
    lines.append("")
    if recent:
        for p in recent:
            title = p.get("Title", "Untitled")
            doi = p.get("DOI", "")
            journal = p.get("Journal", "")
            pub_date = p.get("Publication Date", "")
            topics = p.get("Topics", "")
            summary = p.get("Summary", "")
            authors = p.get("Authors", "")

            if doi:
                lines.append(f"### [{title}]({doi})")
            else:
                lines.append(f"### {title}")

            meta_parts = []
            if pub_date:
                meta_parts.append(f"📅 {pub_date}")
            if journal:
                meta_parts.append(f"📰 *{journal}*")
            if topics:
                meta_parts.append(f"🏷️ {topics}")
            lines.append(" | ".join(meta_parts))
            lines.append("")

            if authors:
                lines.append(f"**Authors:** {authors}")
                lines.append("")
            if summary:
                lines.append(f"> {summary}")
                lines.append("")
            lines.append("---")
            lines.append("")
    else:
        lines.append("*No papers tracked yet. Check back after the first scheduled run.*")
        lines.append("")

    # Monthly reports index
    lines.append("## Monthly Reports")
    lines.append("")
    lines.append("| Month | CSV (view on GitHub) | Excel (download) |")
    lines.append("|:---|:---:|:---:|")
    for f in sorted(REPORTS_DIR.glob("egain_papers_*.csv"), reverse=True):
        stem = f.stem
        xlsx = f.with_suffix(".xlsx")
        # Parse month from filename
        parts = stem.replace("egain_papers_", "").split("_")
        if len(parts) == 2:
            try:
                month_label = datetime(int(parts[0]), int(parts[1]), 1).strftime("%B %Y")
            except ValueError:
                month_label = stem
        else:
            month_label = stem
        csv_link = f"[📄 View](reports/{f.name})"
        xlsx_link = f"[📥 Download](reports/{xlsx.name})" if xlsx.exists() else "—"
        lines.append(f"| {month_label} | {csv_link} | {xlsx_link} |")
    lines.append("")

    lines.append("---")
    lines.append("")
    lines.append("*This tracker runs automatically via a scheduled Claude Code agent.*")
    lines.append("*Papers are sourced from OpenAlex and filtered for EGaIn relevance.*")

    readme_path = ROOT_DIR / "README.md"
    readme_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"README.md updated with dashboard ({len(all_papers)} total papers)")


def main():
    REPORTS_DIR.mkdir(exist_ok=True)

    excel_path, csv_path = get_monthly_paths()
    from_date = (datetime.now() - timedelta(days=SEARCH_DAYS)).strftime("%Y-%m-%d")
    print(f"Monthly report: {excel_path.name}")
    print(f"Searching for EGaIn papers published since {from_date}...")

    papers = fetch_papers(from_date)
    print(f"Found {len(papers)} relevant papers from OpenAlex")

    existing_dois = load_all_existing_dois()
    print(f"Already tracking {len(existing_dois)} papers across all months")

    added, new_papers = append_papers(papers, existing_dois, excel_path, csv_path)
    print(f"Added {added} new papers to {excel_path.name}")

    generate_readme(new_papers)

    if added == 0:
        print("No new papers found — Excel unchanged.")
    else:
        print(f"Excel updated with {added} new entries.")

    return added


if __name__ == "__main__":
    main()
