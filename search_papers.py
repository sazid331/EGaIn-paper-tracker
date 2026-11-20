#!/usr/bin/env python3
"""Search OpenAlex for recent EGaIn-related papers and update an Excel tracker."""

import json
import requests
from datetime import datetime, timedelta
from pathlib import Path

EXCEL_FILE = Path(__file__).parent / "egain_papers.xlsx"
SEARCH_DAYS = 4  # overlap by 1 day to avoid missing papers at boundaries
OPENALEX_API = "https://api.openalex.org/works"

SEARCH_QUERIES = [
    "EGaIn",
    "eutectic gallium indium",
    "liquid metal gallium indium",
    "gallium indium alloy soft electronics",
    "gallium indium microfluidics",
    "liquid metal stretchable sensor",
    "liquid metal soft robotics",
]

MAILTO = "rssazid331@gmail.com"


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


def fetch_papers(from_date: str) -> dict:
    """Fetch papers from OpenAlex, deduplicated by DOI."""
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


def load_existing_dois() -> set:
    """Load DOIs already in the Excel file to avoid duplicates."""
    import openpyxl

    if not EXCEL_FILE.exists():
        return set()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    dois = set()
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True):
        if row[0]:
            dois.add(str(row[0]).strip().lower())
    wb.close()
    return dois


def create_or_load_workbook():
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill

    if EXCEL_FILE.exists():
        return openpyxl.load_workbook(EXCEL_FILE)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EGaIn Papers"

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

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 80
    ws.column_dimensions["G"].width = 60
    ws.column_dimensions["H"].width = 10
    ws.column_dimensions["I"].width = 40
    ws.column_dimensions["J"].width = 15

    wb.save(EXCEL_FILE)
    return wb


def generate_summary(title: str, abstract: str) -> str:
    """Generate a brief summary from title and abstract."""
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


def append_papers(papers: dict, existing_dois: set) -> int:
    from openpyxl.styles import Alignment

    wb = create_or_load_workbook()
    ws = wb.active
    added = 0
    today = datetime.now().strftime("%Y-%m-%d")

    for paper in papers.values():
        doi_lower = paper["doi"].strip().lower()
        if doi_lower and doi_lower in existing_dois:
            continue

        summary = generate_summary(paper["title"], paper["abstract"])

        row = [
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
        ws.append(row)

        for col in range(1, 11):
            cell = ws.cell(row=ws.max_row, column=col)
            cell.alignment = Alignment(wrap_text=True, vertical="top")

        added += 1
        if doi_lower:
            existing_dois.add(doi_lower)

    wb.save(EXCEL_FILE)
    wb.close()
    return added


def main():
    from_date = (datetime.now() - timedelta(days=SEARCH_DAYS)).strftime("%Y-%m-%d")
    print(f"Searching for EGaIn papers published since {from_date}...")

    papers = fetch_papers(from_date)
    print(f"Found {len(papers)} papers from OpenAlex")

    existing_dois = load_existing_dois()
    print(f"Already tracking {len(existing_dois)} papers")

    added = append_papers(papers, existing_dois)
    print(f"Added {added} new papers to {EXCEL_FILE.name}")

    if added == 0:
        print("No new papers found — Excel unchanged.")
    else:
        print(f"Excel updated with {added} new entries.")

    return added


if __name__ == "__main__":
    main()
