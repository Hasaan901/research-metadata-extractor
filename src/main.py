import os
import gradio as gr
import requests
import pandas as pd
import re
from bs4 import BeautifulSoup
from datetime import datetime
from functools import lru_cache

# ==========================================================
# 📁 Configuration (Change these as needed)
# ==========================================================
# Get the script's directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

hec_excel_path = os.path.join(DATA_DIR, "national_journals_2024_25.xlsx")
impact_excel_path = os.path.join(DATA_DIR, "if2025_jcr.xlsx")  # User needs to provide this
export_path = os.path.join(OUTPUT_DIR, "submitted_metadata.xlsx")

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==========================================================
# 📊 Load Excel Files
# ==========================================================

@lru_cache(maxsize=1)
def load_impact_data():
    """Load JCR Impact Factor data."""
    if not os.path.exists(impact_excel_path):
        print(f"⚠️ Warning: Impact Factor file '{impact_excel_path}' not found. Impact factor features will be limited.")
        return pd.DataFrame(columns=["ISSN", "eISSN", "JIF 2024", "JIF Quartile"])
    
    df = pd.read_excel(impact_excel_path)
    df["ISSN"] = df["ISSN"].astype(str).str.strip()
    df["eISSN"] = df["eISSN"].astype(str).str.strip()
    return df

@lru_cache(maxsize=1)
def load_hec_journals():
    """Load HEC journals from Excel file."""
    if not os.path.exists(hec_excel_path):
        print(f"⚠️ Warning: HEC Journals file '{hec_excel_path}' not found.")
        return pd.DataFrame(columns=["ISSN", "HEC Category"])

    df = pd.read_excel(hec_excel_path)
    
    # Use the expected column names
    issn_p_col = "ISSN P"
    issn_e_col = "ISSN E"
    category_col = "Category"

    hec_data = []
    for _, row in df.iterrows():
        # Extract just the letter from "Y category" -> "Y"
        cat_value = str(row[category_col]).strip().upper()
        category = cat_value[0] if cat_value and cat_value[0] in ['X', 'Y', 'W'] else 'W'

        # Add Print ISSN
        if pd.notna(row[issn_p_col]):
            issn_p = str(row[issn_p_col]).strip()
            if issn_p and issn_p != 'nan':
                hec_data.append({"ISSN": issn_p, "HEC Category": category})

        # Add Electronic ISSN
        if pd.notna(row[issn_e_col]):
            issn_e = str(row[issn_e_col]).strip()
            if issn_e and issn_e != 'nan':
                hec_data.append({"ISSN": issn_e, "HEC Category": category})

    return pd.DataFrame(hec_data)

# ==========================================================
# 🧹 Helpers
# ==========================================================
def clean_html(html):
    return BeautifulSoup(html, "html.parser").get_text(" ", strip=True) if html else None

def extract_doi(text):
    match = re.search(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", text, re.I)
    return match.group(0) if match else text.strip()

def format_publication_date(date_parts):
    if not date_parts or not date_parts[0]:
        return None

    parts = date_parts[0]
    month_names = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]

    if len(parts) == 1:
        return str(parts[0])
    elif len(parts) == 2:
        return f"{month_names[parts[1]-1]} {parts[0]}"
    elif len(parts) >= 3:
        return f"{month_names[parts[1]-1]} {parts[2]}, {parts[0]}"

    return str(parts[0])

# ==========================================================
# 🔎 Cached API Calls
# ==========================================================
@lru_cache(maxsize=1024)
def fetch_crossref(doi):
    try:
        r = requests.get(f"https://api.crossref.org/works/{doi}", timeout=7)
        return r.json().get("message") if r.status_code == 200 else None
    except:
        return None

@lru_cache(maxsize=1024)
def fetch_openalex(doi):
    try:
        r = requests.get(f"https://api.openalex.org/works/doi:{doi}", timeout=7)
        return r.json() if r.status_code == 200 else None
    except:
        return None

@lru_cache(maxsize=1024)
def fetch_semantic(doi):
    try:
        r = requests.get(f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}", 
                         params={"fields": "abstract,tldr"}, timeout=7)
        return r.json() if r.status_code == 200 else None
    except:
        return None

@lru_cache(maxsize=1024)
def fetch_unpaywall(doi):
    try:
        r = requests.get(f"https://api.unpaywall.org/v2/{doi}", 
                         params={"email": "research@university.edu"}, timeout=7)
        return r.json() if r.status_code == 200 else None
    except:
        return None

# ==========================================================
# 📊 Category & Impact Lookups
# ==========================================================
def get_if_quartile(issn_list):
    impact_df = load_impact_data()
    for iss in issn_list:
        match = impact_df[(impact_df["ISSN"] == iss) | (impact_df["eISSN"] == iss)]
        if not match.empty:
            row = match.iloc[0]
            return row["JIF 2024"], row["JIF Quartile"]
    return None, None

def get_hec_category(issn_list):
    hec_df = load_hec_journals()
    for iss in issn_list:
        match = hec_df[hec_df["ISSN"] == iss]
        if not match.empty:
            return match.iloc[0]["HEC Category"]
    return "W"

# ==========================================================
# 🔥 MASTER METADATA BUILDER
# ==========================================================
def fetch_metadata(doi_raw):
    doi = extract_doi(doi_raw)
    cross = fetch_crossref(doi)
    if not cross:
        return {"Error": f"DOI not found -> {doi}"}

    # Basic Info
    title = cross.get("title", [None])[0]
    journal = cross.get("container-title", [None])[0]
    issn_list = cross.get("ISSN", [])

    # Date
    pub_date = None
    for field in ["published", "published-print", "issued"]:
        if cross.get(field):
            pub_date = format_publication_date(cross[field].get("date-parts"))
            if pub_date: break
    
    year = None
    for field in ["issued", "published"]:
        if cross.get(field, {}).get("date-parts"):
            year = cross[field]["date-parts"][0][0]
            if year: break
    
    pub_date = pub_date or "Not Available"

    # Authors & Affiliations
    authors = [f"{a.get('given', '')} {a.get('family', '')}".strip() for a in cross.get("author", [])]
    affs = {af.get("name", "").strip() for a in cross.get("author", []) for af in a.get("affiliation", []) if af.get("name")}

    # Abstract
    abstract = clean_html(cross.get("abstract"))
    if not abstract:
        s2 = fetch_semantic(doi)
        if s2: abstract = s2.get("abstract")
    if not abstract:
        unpaywall = fetch_unpaywall(doi)
        if unpaywall: abstract = clean_html(unpaywall.get("abstract"))
    if not abstract:
        oa = fetch_openalex(doi)
        if oa and oa.get("abstract_inverted_index"):
            try:
                words = sorted([(pos, word) for word, positions in oa["abstract_inverted_index"].items() for pos in positions])
                abstract = " ".join([w[1] for w in words])
            except: pass
    
    abstract = abstract or "Not Available"

    # Keywords
    keywords = list(cross.get("subject", []))
    oa = fetch_openalex(doi)
    if oa and oa.get("doi", "").lower().replace("https://doi.org/", "") == doi.lower():
        keywords += [c["display_name"] for c in oa.get("concepts", [])[:5]]
    keywords = ", ".join(list(dict.fromkeys(keywords))) if keywords else "Not Available"

    # IF + HEC
    jif, quartile = get_if_quartile(issn_list)
    hec_cat = get_hec_category(issn_list)

    # APA Citation
    lead = authors[0] if authors else "Unknown"
    apa = f"{lead} et al. ({year}). {title}. {journal}, {cross.get('volume')}({cross.get('issue')}), {cross.get('page')}. https://doi.org/{doi}"

    return {
        "DOI": doi, "Title": title, "Authors": ", ".join(authors),
        "Affiliations": ", ".join(affs) if affs else "Not Available",
        "Journal": journal, "ISSN": ", ".join(issn_list),
        "Volume": cross.get("volume"), "Issue": cross.get("issue"), "Pages": cross.get("page"),
        "Publication Date": pub_date, "Year": year, "Abstract": abstract, "Keywords": keywords,
        "Impact Factor": jif if jif else "Not Found",
        "JIF Quartile": quartile if quartile else "Not Found",
        "HEC Category": hec_cat, "APA Citation": apa
    }

# ==========================================================
# EXCEL EXPORT
# ==========================================================
def export_to_excel(all_data, role):
    df = pd.DataFrame(all_data)
    df["Submitted By"] = role
    df["Fetched On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df.to_excel(export_path, index=False)
    return export_path

# ==========================================================
# GRADIO UI
# ==========================================================
def process(dois_text, role):
    dois = [d.strip() for d in re.split(r"[, \n]", dois_text) if d.strip()]
    results, html = [], ""
    admin_only_fields = ["Impact Factor", "JIF Quartile", "HEC Category", "APA Citation"]

    for d in dois:
        data = fetch_metadata(d)
        results.append(data)
        html += "<div style='padding:10px;border:1px solid #eee;margin:10px;border-radius:8px;background:#f9f9f9;'>"
        for k, v in data.items():
            if role != "admin" and k in admin_only_fields: continue
            html += f"<p><b>{k}:</b> {v}</p>"
        html += "</div>"

    file = export_to_excel(results, role)
    return html, file

def main():
    ui = gr.Interface(
        fn=process,
        inputs=[
            gr.Textbox(lines=6, label="Enter DOI(s)", placeholder="Enter DOIs separated by comma or newline..."),
            gr.Dropdown(["admin", "user"], value="admin", label="Access Level")
        ],
        outputs=[gr.HTML(label="Fetched Metadata"), gr.File(label="Export Excel")],
        title="📚 Research Portal DOI Extractor",
        description="Extract metadata from DOIs and lookup JCR Impact Factors/HEC Categories automatically.",
        theme=gr.themes.Soft()
    )
    ui.launch()

if __name__ == "__main__":
    main()
