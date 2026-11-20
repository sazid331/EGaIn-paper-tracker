#!/usr/bin/env python3
"""Search OpenAlex for recent EGaIn-related papers and update a monthly Excel tracker."""

import json
import re
import requests
from datetime import datetime, timedelta
from pathlib import Path

REPORTS_DIR = Path(__file__).parent / "reports"
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

MAILTO = "rssazid331@gmail.com"


def get_monthly_excel_path() -> Path:
    now = datetime.now()
    filename = f"egain_papers_{now.year}_{now.month:02d}.xlsx"
    return REPORTS_DIR / filename


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

            papers[oa_id] = {
                "title": title,
                "authors": authors,
                "doi": doi,
                "journal": journal,
                "publication_date": pub_date,
                "abstract": abstract,
                "cited_by_count": cited_by,
                "open_access_url": oa_url,
            }

    return papers


def load_existing_dois(excel_path: Path) -> set:
    import openpyxl

    if not excel_path.exists():
        return set()
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    dois = set()
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True):
        if row[0]:
            dois.add(str(row[0]).strip().lower())
    wb.close()
    return dois


def load_all_existing_dois() -> set:
    """Load DOIs from all monthly reports to avoid duplicates across months."""
    dois = set()
    if not REPORTS_DIR.exists():
        return dois
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
        "Title",
        "Authors",
        "DOI",
        "Journal",
        "Publication Date",
        "Abstract",
        "Summary",
        "Cited By",
        "Open Access URL",
        "Date Added",
    ]

    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    thin_border = Border(
        bottom=Side(style="thin", color="B4C6E7"),
    )

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    ws.row_dimensions[1].height = 30

    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 35
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 90
    ws.column_dimensions["G"].width = 65
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 45
    ws.column_dimensions["J"].width = 15

    wb.save(excel_path)
    return wb


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


def append_papers(papers: dict, existing_dois: set, excel_path: Path) -> int:
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
        return 0

    wb = create_or_load_workbook(excel_path)
    ws = wb.active
    today = datetime.now().strftime("%Y-%m-%d")

    body_font = Font(name="Calibri", size=11)
    title_font = Font(name="Calibri", size=11, bold=True)
    doi_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
    row_border = Border(bottom=Side(style="thin", color="D9E2F3"))
    light_fill = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")

    for i, paper in enumerate(new_papers):
        summary = generate_summary(paper["title"], paper["abstract"])
        row_data = [
            paper["title"],
            paper["authors"],
            paper["doi"],
            paper["journal"],
            paper["publication_date"],
            paper["abstract"],
            summary,
            paper["cited_by_count"],
            paper["open_access_url"],
            today,
        ]
        ws.append(row_data)
        row_num = ws.max_row

        use_stripe = (row_num % 2 == 0)
        ws.row_dimensions[row_num].height = 45

        for col in range(1, 11):
            cell = ws.cell(row=row_num, column=col)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = row_border
            if col == 1:
                cell.font = title_font
            elif col == 3:
                cell.font = doi_font
            else:
                cell.font = body_font
            if use_stripe:
                cell.fill = light_fill

    wb.save(excel_path)
    wb.close()
    return len(new_papers)


def main():
    REPORTS_DIR.mkdir(exist_ok=True)

    excel_path = get_monthly_excel_path()
    from_date = (datetime.now() - timedelta(days=SEARCH_DAYS)).strftime("%Y-%m-%d")
    print(f"Monthly report: {excel_path.name}")
    print(f"Searching for EGaIn papers published since {from_date}...")

    papers = fetch_papers(from_date)
    print(f"Found {len(papers)} relevant papers from OpenAlex")

    existing_dois = load_all_existing_dois()
    print(f"Already tracking {len(existing_dois)} papers across all months")

    added = append_papers(papers, existing_dois, excel_path)
    print(f"Added {added} new papers to {excel_path.name}")

    if added == 0:
        print("No new papers found — Excel unchanged.")
    else:
        print(f"Excel updated with {added} new entries.")

    return added


if __name__ == "__main__":
    main()
