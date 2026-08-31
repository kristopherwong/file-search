import os
import re
import fitz  # PyMuPDF
import pandas as pd
import argparse
from tqdm import tqdm
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from urllib.parse import urlparse
import json

# === REGEX ===
url_regex = re.compile(r"https?://[^\s<>)\"']+")

def get_base_url(url):
    try:
        parsed = urlparse(url)
        return parsed.netloc.lower()
    except:
        return ""
    
import os

def get_unique_output_folder(base_name="run"):
    os.makedirs("output", exist_ok=True)
    i = 1
    folder = os.path.join("output", base_name)
    while os.path.exists(folder):
        folder = os.path.join("output", f"{base_name}_{i}")
        i += 1
    os.makedirs(folder)
    return folder

def save_large_json(data, base_filename, folder, threshold=5000):
    """Save large JSON data to external file if size exceeds threshold."""
    json_str = json.dumps(data, indent=2, ensure_ascii=False)

    if len(json_str) <= threshold:
        return json_str

    os.makedirs(folder, exist_ok=True)
    safe_name = re.sub(r'[^\w\-_.]', '_', base_filename)[:80]  # sanitize + truncate
    path = os.path.join(folder, f"{safe_name}.json")

    with open(path, "w", encoding="utf-8") as f:
        f.write(json_str)

    return f"[See {os.path.basename(path)}]"

def parse_args():
    parser = argparse.ArgumentParser(description="Extract links, keywords, and Bates numbers from PDFs.")
    parser.add_argument("path", help="Path to a single PDF file to process (default). Use --folder to instead treat this as a folder of PDFs to scan recursively.")
    parser.add_argument("--folder", action="store_true", help="Treat 'path' as a folder and recursively scan it for PDF files, instead of a single PDF file.")
    parser.add_argument("--output", default="pdf_extraction_output.xlsx", help="Output Excel file path")
    parser.add_argument("--link-annotations", action="store_true", help="Extract embedded link annotations from PDFs")
    parser.add_argument("--text-urls", action="store_true", help="Extract URLs from visible page text")
    parser.add_argument("--keywords", nargs="*", default=[], help="Search for specific keywords")
    parser.add_argument("--keywords-file", help="Path to file with one keyword per line")
    parser.add_argument("--bates-footer-prefix", help="Prefix string to identify Bates numbers in the bottom-right footer of each page (e.g. 'MyCompany')")
    parser.add_argument("--bates-body-prefix", help="Prefix string to identify Bates numbers anywhere in a page's visible text/content, not just the footer (e.g. 'MyCompany')")
    parser.add_argument("--context-window", type=int, default=100, help="Number of characters prepending and appending matched string that are returned in 'Context' parameter of 'References' column")

    return parser.parse_args()


def extract_context(text, match_start, match_end, window):
    return text[max(0, match_start - window):match_end + window]

def clean_url_string(url: str) -> str:
    # Remove newlines and collapse spaces
    return re.sub(r"\s+", "", url.replace('\n', '').replace('\r', ''))

def clean_context_string(text: str) -> str:
    text = text.replace('\n', ' ').replace('\r', '')
    text = re.sub(r"[•■·]", "", text)
    return re.sub(r'\s+', ' ', text).strip()

def format_excel(filepath):
    wb = load_workbook(filepath)
    for ws in wb.worksheets:
        # Bold headers
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # Wrap text and auto-fit width
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                if cell.value:
                    cell.alignment = Alignment(wrap_text=True)
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 4, 80)

    wb.save(filepath)
    
def format_references_to_json(refs):
    return json.dumps(refs, indent=2, ensure_ascii=False)

def main():
    args = parse_args()
    output_dir = get_unique_output_folder(os.path.splitext(os.path.basename(args.output))[0])
    excel_path = os.path.join(output_dir, os.path.basename(args.output))
    
    # Load keywords from file if specified
    if args.keywords_file:
        try:
            with open(args.keywords_file, "r", encoding="utf-8") as f:
                file_keywords = [line.strip() for line in f if line.strip()]
                args.keywords.extend(file_keywords)
        except Exception as e:
            print(f"⚠️ Failed to load keywords from {args.keywords_file}: {e}")

    link_results = []
    keyword_results = []

    if args.folder:
        pdf_files = []
        for root, _, files in os.walk(args.path):
            for f in files:
                if f.lower().endswith(".pdf"):
                    pdf_files.append(os.path.join(root, f))
    else:
        if not os.path.isfile(args.path):
            print(f"❌ {args.path} is not a file. Use --folder to scan a directory of PDFs instead.")
            return
        pdf_files = [args.path]

    for filepath in tqdm(pdf_files, desc="Processing PDFs", unit="file"):
        filename = os.path.basename(filepath)
        try:
            doc = fitz.open(filepath)

            for page_num, page in enumerate(doc, start=1):
                need_text = args.text_urls or args.keywords or args.bates_body_prefix
                if need_text:
                    text = page.get_text("text")
                    # Clean full page text before matching (safe replacement for display only)
                    cleaned_text = text.replace('\r', '').replace('\n', ' ')

                # Extract per-page Bates number from the bottom-right footer
                bates_id_footer = None
                if args.bates_footer_prefix:
                    br_text = page.get_textbox((page.rect.width - 200, page.rect.height - 100, page.rect.width, page.rect.height))
                    footer_pattern = re.compile(rf"{re.escape(args.bates_footer_prefix)}[_\-]?\d+", re.IGNORECASE)
                    footer_matches = footer_pattern.findall(br_text)
                    if footer_matches:
                        bates_id_footer = footer_matches[-1]  # Last match in the footer

                # Extract Bates number anywhere in the page's visible content
                bates_id_body = None
                if args.bates_body_prefix:
                    body_pattern = re.compile(rf"{re.escape(args.bates_body_prefix)}[_\-]?\d+", re.IGNORECASE)
                    body_matches = body_pattern.findall(text)
                    if body_matches:
                        bates_id_body = body_matches[-1]  # Last match on the page

                bates_fields = {}
                if args.bates_footer_prefix:
                    bates_fields["Bates ID (Footer)"] = bates_id_footer
                if args.bates_body_prefix:
                    bates_fields["Bates ID (Body)"] = bates_id_body

                # Link annotations
                # --- Link annotations ---
                if args.link_annotations:
                    for link in page.get_links():
                        uri = link.get("uri")
                        if uri:
                            link_results.append({
                                "Filename": filename,
                                "Page": page_num,
                                "Match Type": "link",
                                "Matched String": uri,
                                "Context": "",
                                **bates_fields,
                                "Base URL": get_base_url(uri)
                            })


                # Text-based search
                if args.text_urls or args.keywords:
                    if args.text_urls:
                        for match in url_regex.finditer(cleaned_text):
                            match_start = match.start()
                            match_end = match.end()
                            context = clean_context_string(
                                extract_context(text, match_start, match_end, args.context_window)
                            )
                            raw_url = match.group()
                            url = clean_url_string(raw_url)
                            link_results.append({
                                "Filename": filename,
                                "Page": page_num,
                                "Match Type": "text_url",
                                "Matched String": url,
                                "Context": context,
                                **bates_fields,
                                "Base URL": get_base_url(url)
                            })

                    for keyword in args.keywords:
                        for match in re.finditer(re.escape(keyword), text, flags=re.IGNORECASE):
                            match_start = match.start()
                            match_end = match.end()
                            context = clean_context_string(
                                extract_context(text, match_start, match_end, args.context_window)
                            )
                            keyword_results.append({
                                "Filename": filename,
                                "Page": page_num,
                                "Match Type": "keyword",
                                "Matched String": match.group(),
                                "Context": context,
                                **bates_fields
                            })

            doc.close()

        except Exception as e:
            print(f"❌ Error processing {filename}: {e}")


    df_links=pd.DataFrame()
    df_keywords=pd.DataFrame()

    if link_results:
        # Create dataframe from raw link results
        df_links_raw = pd.DataFrame(link_results)

        # Group by exact URL
        grouped_links = df_links_raw.groupby("Matched String")
        bates_columns = [c for c in ("Bates ID (Footer)", "Bates ID (Body)") if c in df_links_raw.columns]

        merged_link_rows = []

        for url, group in grouped_links:
            base_url = group["Base URL"].iloc[0]
            match_type = group["Match Type"].iloc[0]

            references = group.apply(
                lambda row: {
                    "Filename": row["Filename"],
                    "Page": row["Page"],
                    "Context": row["Context"],
                    **{col: row[col] for col in bates_columns}
                }, axis=1
            ).tolist()

            merged_link_rows.append({
                "Matched String": url,
                "Base URL": base_url,
                "Match Type": match_type,
                "Reference Count": len(references),
                "References": references
            })

        df_links = pd.DataFrame(merged_link_rows)
        df_links.sort_values(by=["Base URL", "Reference Count"], ascending=[True, False], inplace=True)

    if keyword_results:
        # Create dataframe from raw keyword results
        df_keywords_raw = pd.DataFrame(keyword_results)

        # Group by exact matched keyword
        grouped_keywords = df_keywords_raw.groupby("Matched String")
        bates_columns = [c for c in ("Bates ID (Footer)", "Bates ID (Body)") if c in df_keywords_raw.columns]

        merged_keyword_rows = []

        for keyword, group in grouped_keywords:
            match_type = group["Match Type"].iloc[0]

            references = group.apply(
                lambda row: {
                    "Filename": row["Filename"],
                    "Page": row["Page"],
                    "Context": row["Context"],
                    **{col: row[col] for col in bates_columns}
                }, axis=1
            ).tolist()

            merged_keyword_rows.append({
                "Matched String": keyword,
                "Match Type": match_type,
                "Reference Count": len(references),
                "References": references
            })

        df_keywords = pd.DataFrame(merged_keyword_rows)
        df_keywords.sort_values(by=["Reference Count", "Matched String"], ascending=[False, True], inplace=True)

    json_folder = os.path.join(output_dir, "references_json")
    if not df_links.empty or not df_keywords.empty:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            if not df_links.empty:
                df_links["References"] = df_links.apply(
                    lambda row: save_large_json(
                        row["References"],
                        base_filename=f"{row['Matched String'][:50].strip().replace('/', '_')}_links",
                        folder=json_folder
                    ),
                    axis=1
                )
                df_links.to_excel(writer, sheet_name="URLs", index=False)
            if not df_keywords.empty:
                df_keywords["References"] = df_keywords.apply(
                    lambda row: save_large_json(
                        row["References"],
                        base_filename=f"{row['Matched String'][:50].strip().replace('/', '_')}_keywords",
                        folder=json_folder
                    ),
                    axis=1
)
                df_keywords.to_excel(writer, sheet_name="Keywords", index=False)

        format_excel(excel_path)
        print(f"📘 Excel saved to {excel_path}")
    else:
        print(f"❌ no references or links found")



if __name__ == "__main__":
    main()
