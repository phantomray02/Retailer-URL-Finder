import io
import json
import re
import urllib.parse
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Retailer URL Finder", layout="wide")

RETAILER_RULES = {
    "heb": {
        "label": "H-E-B",
        "aliases": ["heb", "h-e-b", "h e b"],
        "search_prefix": "https://www.heb.com/search?q=",
        "domain": "heb.com",
        "product_patterns": ["/product-detail/"],
        "reject_patterns": ["/search", "/category", "/browse", "/coupons", "/weekly-ad"],
        "id_in_url_preferred": True,
    },
    "cvs": {
        "label": "CVS",
        "aliases": ["cvs", "cvs pharmacy"],
        "search_prefix": "https://www.cvs.com/search?searchTerm=",
        "domain": "cvs.com",
        "product_patterns": ["/shop/", "prodid"],
        "reject_patterns": ["/search", "/category", "/content", "/account", "/weeklyad"],
        "id_in_url_preferred": False,
    },
    "walgreens": {
        "label": "Walgreens",
        "aliases": ["walgreens", "wags"],
        "search_prefix": "https://www.walgreens.com/search/results.jsp?Ntt=",
        "domain": "walgreens.com",
        "product_patterns": ["/store/c/", "id=prod", "-product"],
        "reject_patterns": ["/search", "/topic", "/offers", "/photo", "/storelocator"],
        "id_in_url_preferred": False,
    },
    "walmart": {
        "label": "Walmart",
        "aliases": ["walmart", "wal-mart"],
        "search_prefix": "https://www.walmart.com/search?q=",
        "domain": "walmart.com",
        "product_patterns": ["/ip/"],
        "reject_patterns": ["/search", "/browse", "/cp/", "/brand/", "/shop/"],
        "id_in_url_preferred": False,
    },
    "target": {
        "label": "Target",
        "aliases": ["target"],
        "search_prefix": "https://www.target.com/s?searchTerm=",
        "domain": "target.com",
        "product_patterns": ["/p/", "/a-"],
        "reject_patterns": ["/s?", "/c/", "/cat", "/circle", "/weekly-ad"],
        "id_in_url_preferred": False,
    },
    "kroger": {
        "label": "Kroger",
        "aliases": ["kroger"],
        "search_prefix": "https://www.kroger.com/search?query=",
        "domain": "kroger.com",
        "product_patterns": ["/p/"],
        "reject_patterns": ["/search", "/pl/", "/cl/", "/weeklyad"],
        "id_in_url_preferred": False,
    },
    "samsclub": {
        "label": "Sam's Club",
        "aliases": ["samsclub", "sam's club", "sams club", "sam club"],
        "search_prefix": "https://www.samsclub.com/s/",
        "domain": "samsclub.com",
        "product_patterns": ["/p/"],
        "reject_patterns": ["/s/", "/b/", "/c/", "/club"],
        "id_in_url_preferred": False,
    },
    "amazon": {
        "label": "Amazon",
        "aliases": ["amazon"],
        "search_prefix": "https://www.amazon.com/s?k=",
        "domain": "amazon.com",
        "product_patterns": ["/dp/", "/gp/product/"],
        "reject_patterns": ["/s?", "/stores/", "/b?", "/hz/", "/gp/bestsellers"],
        "id_in_url_preferred": False,
    },
}


def normalize_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def normalize_retailer(value):
    raw = normalize_text(value)
    compact = re.sub(r"[^a-z0-9]", "", raw)
    for key, rules in RETAILER_RULES.items():
        aliases = [key] + rules.get("aliases", []) + [rules.get("label", "")]
        for alias in aliases:
            alias_norm = normalize_text(alias)
            alias_compact = re.sub(r"[^a-z0-9]", "", alias_norm)
            if raw == alias_norm or compact == alias_compact:
                return key
    return ""


def build_search_url(retailer_key, search_term):
    rules = RETAILER_RULES[retailer_key]
    encoded = urllib.parse.quote_plus(str(search_term).strip())
    return f"{rules['search_prefix']}{encoded}"


def load_input(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file, engine="openpyxl")


def prep_dataframe(df):
    df = df.copy()
    if "Retailer" not in df.columns or "Search Term" not in df.columns:
        raise ValueError("Input must have columns named Retailer and Search Term.")

    rows = []
    for i, row in df.iterrows():
        retailer = row.get("Retailer", "")
        search_term = row.get("Search Term", "")
        retailer_key = normalize_retailer(retailer)
        if not retailer_key or pd.isna(search_term) or str(search_term).strip() == "":
            search_url = ""
        else:
            search_url = build_search_url(retailer_key, search_term)
        rows.append({
            "row_id": int(i) + 1,
            "retailer": str(retailer),
            "retailer_key": retailer_key,
            "search_term": "" if pd.isna(search_term) else str(search_term),
            "search_url": search_url,
        })
    return rows


def results_to_excel(results):
    df = pd.DataFrame(results)
    ordered = [
        "row_id", "retailer", "search_term", "search_url", "product_url",
        "status", "confidence", "matched_title", "notes", "captured_at"
    ]
    cols = [c for c in ordered if c in df.columns] + [c for c in df.columns if c not in ordered]
    df = df[cols]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="URL Results")
    output.seek(0)
    return output


st.title("Retailer URL Finder")
st.caption("Separate Streamlit + browser extension workflow. Streamlit prepares the job and exports results; the extension searches live retailer pages in your browser.")

with st.expander("Input format", expanded=True):
    st.markdown("""
Upload an Excel or CSV with these exact columns:

```text
Retailer | Search Term
HEB | 5002549
Walgreens | U by Kotex CleanWear Pads 32ct
CVS | Cottonelle wipes
```
""")

sample_df = pd.DataFrame([
    {"Retailer": "HEB", "Search Term": "5002549"},
    {"Retailer": "Walgreens", "Search Term": "U by Kotex CleanWear Pads 32ct"},
    {"Retailer": "CVS", "Search Term": "Cottonelle wipes"},
])
sample_buffer = io.BytesIO()
with pd.ExcelWriter(sample_buffer, engine="openpyxl") as writer:
    sample_df.to_excel(writer, index=False, sheet_name="Input")
sample_buffer.seek(0)
st.download_button("Download sample input Excel", sample_buffer, "sample_input.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

uploaded = st.file_uploader("Upload input Excel or CSV", type=["xlsx", "xlsm", "csv"])

if uploaded:
    try:
        input_df = load_input(uploaded)
        job_rows = prep_dataframe(input_df)
        st.subheader("Prepared Search Job")
        st.dataframe(pd.DataFrame(job_rows), use_container_width=True)

        job_payload = {
            "version": "1.0",
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "rules": RETAILER_RULES,
            "rows": job_rows,
        }
        job_json = json.dumps(job_payload, indent=2)

        st.markdown("### Step 1: Copy this job JSON into the browser extension")
        st.text_area("Job JSON", job_json, height=260)
        st.download_button("Download job JSON", job_json, "retailer_url_job.json", mime="application/json")

        st.markdown("### Step 2: Paste extension results here")
        results_text = st.text_area("Results JSON from extension", height=220, placeholder="Paste the results JSON copied by the extension.")
        if results_text.strip():
            try:
                parsed = json.loads(results_text)
                results = parsed.get("results", parsed if isinstance(parsed, list) else [])
                st.success(f"Loaded {len(results)} result rows.")
                st.dataframe(pd.DataFrame(results), use_container_width=True)
                excel = results_to_excel(results)
                st.download_button(
                    "Download results Excel",
                    excel,
                    "retailer_url_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as exc:
                st.error(f"Could not parse results JSON: {exc}")
    except Exception as exc:
        st.error(str(exc))

st.markdown("---")
st.markdown("### Supported retailers in V1")
st.write(", ".join([rules["label"] for rules in RETAILER_RULES.values()]))
