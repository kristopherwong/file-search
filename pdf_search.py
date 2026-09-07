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
# Generic Bates-style reference number: a short letter prefix followed by a
# padded run of digits (e.g. "ACME-000123", "SMITH_0001234", "ABC000123").
generic_bates_regex = re.compile(r"\b[A-Za-z]{1,10}[_\-]?\d{3,10}\b")

# Known source code file extensions (used with --source-code)
SOURCE_CODE_EXTENSIONS = [
    "py", "js", "ts", "jsx", "tsx", "c", "h", "cpp", "hpp", "cc", "cs",
    "java", "rb", "go", "rs", "swift", "kt", "kts", "scala", "php",
    "sh", "bash", "ps1", "m", "mm", "lua", "pl", "r",
]

def build_file_pattern(extensions):
    ext_alts = "|".join(re.escape(e) for e in extensions)
    return re.compile(rf"\b[\w.\-/\\]+\.(?:{ext_alts})\b", re.IGNORECASE)

def _is_file_path(name):
    return ("/" in name) or ("\\" in name)

def _file_basename(name):
    return re.split(r"[\\/]", name.strip())[-1].lower()

def reconstruct_wrapped_path(text, match_start, match_end, max_lines=20, max_chars=512):
    """Recover a file path that the source text wrapped across multiple lines.

    `text` is the raw page text (with its newlines). The regex match is only the
    last segment of a longer path, e.g.

        ".../com/acme/app/\nUtils.java"

    which `build_file_pattern` matches merely as "Utils.java".
    Working backwards, we fold in preceding path lines, reconstructing the full
    path. A line is folded when its trailing token meets one of three break
    rules: it ends in a directory separator (/ or \\); it ends in a hyphen or
    underscore that splits a path word (and the match already has a separator);
    or the PDF split a word mid-name, in which case the line above is itself a
    path (contains a slash or backslash), the trailing token ends in a plain
    letter, and the
    match is a filename tail (e.g. "Provider.java" split as "...Provid" +
    "er.java"). Folding always stops at a line boundary introduced by a comma
    ("..., prefix") — that token is the prefix of the *next* entry in a
    comma-separated citation list (e.g. "..., acme-Release_…/" wrapping), not
    a continuation of this path, so it is never folded. Folding stays bounded
    (max_lines/max_chars) and uses only each
    line's final contiguous path-token, so prose ("see ...", "...files") and
    stray "dir/" + unrelated-name pairs are not over-joined. Returns
    (reconstructed_path, reconstructed_start).
    """
    full = text[match_start:match_end]
    if not full:
        return full, match_start

    pos = match_start
    folded_lines = 0
    while folded_lines < max_lines:
        # line boundary = the newline immediately before the current segment.
        line_start = text.rfind("\n", 0, pos)
        if line_start == -1:
            break
        # The *preceding* line is the text above that boundary: from after the
        # earlier newline up to (not including) line_start.
        prev_nl = text.rfind("\n", 0, line_start)
        preceding = text[prev_nl + 1:line_start].strip()
        # Take the trailing path-token of that line (longest whitespace-free
        # run of path characters ending at the line's end). This folds a real
        # wrapped path even when a bit of prose precedes it ("see C:\repo\..."),
        # because the space before the token breaks the run — while still
        # refusing lines whose ending is a bare word or prose ("...files").
        m = re.search(r"[A-Za-z0-9_.\-\\/]+\Z", preceding)
        seg = m.group(0) if m else ""
        # Is the preceding fragment introduced by a comma? If so, it is a list-
        # item prefix in a comma-separated citation list ("...cpp, tagx-",
        # "...h, tagx-") rather than a plain folder-path fragment — it belongs to
        # *this* match, and folding it must terminate the climb (what is further
        # up on its own line is a different file, not a continuation of ours).
        comma_boundary = bool(seg) and re.search(r",\s*\Z",
                                                preceding[:len(preceding) - len(seg)])
        # A real wrapped path can break at three kinds of token boundaries:
        #   - directory separator ("/" or "\")  — always fold (seg is a dir)
        #   - hyphen or underscore              — fold ONLY if the segment
        #     contains a letter AND the match we are joining to already has
        #     a separator (i.e. it is part of a multi-component path, not a
        #     bare filename). The letter check keeps us from folding a
        #     numbered page stamp like "-72-" attached to the path, since a
        #     real path break ("some-\ncomponent-dir/") always has letters
        #     on the segment we are folding.
        #   - mid-word (seg ends in a plain letter) — the PDF split a filename
        #     across the break (...Provid + er.java). Recognized when the line
        #     above is a multi-component path (contains "/" or "\\"), yet its
        #     broken trailing token has no dot (a real name fragment, not a
        #     complete "...impl.java" file on its own line), and the match is a
        #     filename tail ("er.java"). A dot-less segment is the key guard: it
        #     stops two *different* files on consecutive lines from being glued.
        #
        # Two extra guards stop gluing a *neighbouring* citation onto ours:
        #   * full_starts_with_tag — the path we are attaching to already begins
        #     with a root tag (a "tagx-..." prefix, or prose glued onto one such
        #     as "wordtagx-..."). Once a citation carries such a prefix, a
        #     word/letter fragment further left is the tail of a *different* file
        #     on the line above (e.g. ".../inc/Fragment" + "tagx-.../Result.cpp"
        #     = two files), not a continuation of ours, so only a clean directory
        #     boundary (slash) may be folded onto a tagged path.
        #   * the hyphen rule additionally requires the folded fragment to be a
        #     real path (contain a slash) or a clean list-item tag (comma-
        #     introduced). A bare hyphen word that is neither (prose glued
        #     directly onto a tag prefix) is rejected.
        full_starts_with_tag = bool(re.match(r"^[A-Za-z][A-Za-z0-9_]*-", full))
        ok = (
            seg and len(seg) >= 2
            and len(seg) + len(full) <= max_chars
            and (
                seg[-1] in "/\\"
                or (seg[-1] in "-_"
                    and re.search(r"[A-Za-z]", seg)
                    and re.search(r"[\\/]", full)
                    and (
                        re.search(r"[\\/]", seg)
                        or comma_boundary
                        or re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*[-_]", seg) is not None
                    ))
                or (
                    seg[-1].isalpha()
                    and full[0].isalpha()
                    and "." in full
                    and "." not in seg
                    and re.search(r"[\\/]", seg)
                    and not full_starts_with_tag
                )
            )
        )
        if ok:
            # seg already ends in one of the valid break characters ("/" "\"
            # for dir boundaries, "-" or "_" when the word itself was split),
            # so direct concatenation is the right joining.
            full = seg + full
            pos = line_start
            folded_lines += 1
            # A comma-introduced fragment is the prefix of *this* entry in a
            # comma-separated citation list ("...cpp, tagx-" then our path on the
            # line below). We fold exactly that one prefix and stop: climbing past
            # it would drag in the ", tagx-" belonging to the *previous* entries,
            # gluing tagx-tagx-tagx-... onto our path.
            if comma_boundary:
                break
        else:
            break
    return full, pos

def _common_suffix_path(paths):
    """Return the shared trailing directory+file sequence (as a '/'-joined
    relative path) if every path in `paths` ends in an identical run of path
    components longer than just the basename itself; otherwise None.

    Distinct files that merely share a basename (e.g. ".../app/A/Utils.java"
    vs ".../lib/B/Utils.java") yield only the basename as a common suffix and
    therefore return None, so they stay ambiguous. Files that are really the
    same file written with a differing leading prefix (e.g.
    "projA/x/src/Foo.cpp" vs "projB/x/src/Foo.cpp") share "x/src/Foo.cpp" and
    therefore collapse to one identity."""
    if len(paths) < 2:
        return None
    parts_list = [p.replace("\\", "/").split("/") for p in paths]
    min_len = min(len(p) for p in parts_list)
    common = []
    for i in range(1, min_len + 1):
        anchor = parts_list[0][-i]
        if all(p[-i] == anchor for p in parts_list):
            common.insert(0, anchor)
        else:
            break
    if len(common) > 1:
        return "/".join(common)
    return None

def canonicalize_file_citations(matched_strings):
    """Collapse a bare filename (no directory) into a full-path citation when a
    single full path ends in the same basename (e.g. "Utils.java" ->
    ".../com/acme/app/Utils.java"). Returns a {name: representative} map;
    names absent from the map are left unchanged.

    A bare name with exactly one matching full path is merged onto it. A bare
    name with several matching full paths is merged only when those paths all
    share a common trailing directory+file sequence longer than the basename
    (i.e. they are provably the same file, differing merely in a leading
    prefix/scanning artifact); if the full paths are genuinely distinct files
    that only share a basename, the name is left bare so it is never wrongly
    combined.

    Two full-path variants are ALSO folded when one is a mid-word truncation
    of the other: the shorter is a *proper* suffix of the longer one and the
    character immediately before it inside the longer one is an ordinary letter.
    That is how the PDF split a long folder word across a line (e.g.
    "…superlongcomponent/src/Widget.cpp" also appearing as the truncated
    "…component/src/Widget.cpp") — both cite the same file. Requiring the
    boundary to be a letter (isalnum) is deliberate: it lets a same-file
    truncation fold while leaving genuinely different files that share a tail —
    which detach at a path separator or a hyphen/_ word joiner ("…/src",
    "…a-b-a_server/…") — untouched, so they are never wrongly combined.
    Hyphen-tagged prose glues (e.g. "member-…Widget.cpp") are therefore left
    as-is here; they are indistinguishable from distinct files on the string
    alone and must be resolved at capture time (real PDF tokens), not inferred."""
    distinct = set(matched_strings)
    paths_by_basename = {}
    for s in distinct:
        if _is_file_path(s):
            paths_by_basename.setdefault(_file_basename(s), set()).add(s)
    mapping = {}
    for s in distinct:
        if _is_file_path(s):
            continue
        candidates = paths_by_basename.get(_file_basename(s), set())
        if len(candidates) == 1:
            mapping[s] = next(iter(candidates))
        elif len(candidates) > 1 and _common_suffix_path(candidates) is not None:
            mapping[s] = max(candidates, key=len)

    # Full-path citations in the same basename group can still be duplicates of
    # each other when one is a mid-name truncation of another (the PDF split a
    # long word across a line and we recovered only the tail, e.g. the folder
    # "…superlongcomponent/src/Widget.cpp" also appearing, truncated mid-word, as
    # "…component/src/Widget.cpp" — both are the same file). A citation counts as
    # the truncated tail of another when it
    # is a *proper* suffix of that other and the character right before the tail
    # there is an ordinary letter: that proves the split landed mid-word, not at
    # the start of a path component. A genuinely different location that merely
    # shares the tail basename detaches at a "/" (boundary non-alnum), so it is
    # left on its own — distinct files sharing a basename are never merged onto
    # each other.
    for base, cands in paths_by_basename.items():
        n = [c.replace("\\", "/") for c in cands]
        for c in n:
            best = None
            for a in n:
                if a == c or len(a) <= len(c):
                    continue
                if not a.endswith(c):
                    continue
                # Boundary guard: the character immediately before the short
                # citation inside the longer one must be an ordinary letter.
                # That proves the split landed mid-word (a truncation, e.g.
                # "…component/src/…" -> "…superlongcomponent/src/…"), not
                # at a component boundary — a genuinely different location that
                # merely shares the tail basename detaches at "/" (or a "-" /
                # "_" word joiner already handled by the wrap logic) and must
                # stay separate, so distinct files sharing a basename are never
                # merged onto each other.
                if not a[-len(c) - 1].isalnum():
                    continue
                if best is None or len(a) < len(best):
                    best = a
            if best is None:
                continue
            original_c = next(o for o in cands if o.replace("\\", "/") == c)
            target = next(o for o in cands if o.replace("\\", "/") == best)
            mapping.setdefault(original_c, target)
    return mapping

def resolve_bare_citation_via_hint(text, bare_name, after_index, window=400, max_lines=20):
    """Attribute a bare filename to a nearby full-path citation that names the
    same file. When a PDF cites a file by name only and then gives its full
    location on the spot — e.g.

        ...is shown for example by VoiceUtil.kt.
        See a/b/c/app/VoiceUtil.kt (implementing ...)

    the bare match can be safely merged onto that full path even when other files
    share the basename. We scan forward within a bounded window for a token that
    ends in the same basename (case-insensitive), reconstruct its possibly-wrapped
    path, and keep the longest one that is a genuine multi-component path. Returns
    (full_path, end_index_just_past_it), or (None, after_index) when no such
    full-path hint is present (so the caller falls back to the global ambiguity
    guard without guessing)."""
    if not bare_name or _is_file_path(bare_name):
        return None, after_index
    base = _file_basename(bare_name)
    limit = after_index + window
    token_re = re.compile(r"[A-Za-z0-9_.\-\\/]+")
    best = None
    for m in token_re.finditer(text, after_index, limit):
        if _file_basename(m.group()) != base:
            continue
        full, _ = reconstruct_wrapped_path(text, m.start(), m.end(), max_lines=max_lines)
        if not _is_file_path(full):
            continue
        if best is None or len(full) > len(best[0]):
            best = (full, m.end())
    if best is None:
        return None, after_index
    return best[0], best[1]


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
    parser.add_argument("--bates-body", action="store_true", help="Search each page's visible text/content for any Bates-style reference numbers (any prefix), not just the footer. Useful for capturing Bates numbers cited from other document sets.")
    parser.add_argument("--source-code", action="store_true", help="Search each page's visible text for citations of common source code file names (e.g. .py, .js, .java, .c, .ts, ...).")
    parser.add_argument("--file-ext", nargs="+", default=[], help="Search each page's visible text for citations of file names with a specific extension (e.g. --file-ext py txt).")
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
    bates_results = []
    file_results = []

    # Build a combined set of extensions to search for, plus a single regex.
    ext_set = set()
    if args.source_code:
        ext_set.update(SOURCE_CODE_EXTENSIONS)
    for e in args.file_ext:
        e = e.lstrip(".")
        if e:
            ext_set.add(e.lower())
    file_ext_regex = build_file_pattern(sorted(ext_set)) if ext_set else None

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
                need_text = args.text_urls or args.keywords or args.bates_body or file_ext_regex is not None
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

                bates_fields = {}
                if args.bates_footer_prefix:
                    bates_fields["Bates ID (Footer)"] = bates_id_footer

                # Every Bates-style reference number found anywhere in the page's
                # visible content, regardless of prefix, with surrounding context
                if args.bates_body:
                    for match in generic_bates_regex.finditer(text):
                        match_start = match.start()
                        match_end = match.end()
                        context = clean_context_string(
                            extract_context(text, match_start, match_end, args.context_window)
                        )
                        bates_results.append({
                            "Filename": filename,
                            "Page": page_num,
                            "Match Type": "bates_body",
                            "Matched String": match.group(),
                            "Context": context,
                            **bates_fields
                        })

                # File-name citations (source code and/or specific extensions)
                if file_ext_regex is not None:
                    for match in file_ext_regex.finditer(text):
                        matched_string = match.group()
                        match_start = match.start()
                        match_end = match.end()
                        # A wrapped path appears on the next line(s); fold any
                        # separator-terminated preceding lines back in so the
                        # Matched String is the full path, not just the last
                        # segment. This also widens the context window start.
                        full_path, full_start = reconstruct_wrapped_path(text, match_start, match_end)
                        if full_path != matched_string:
                            matched_string = full_path
                            match_start = full_start
                        # A bare filename (no directory) may be named in the body
                        # and pointed at by a nearby full path with the same
                        # basename ("...example by X.kt. See a/b/X.kt"). When such
                        # a hint is present, attribute the bare match to it so it
                        # merges even if the basename is shared by other files.
                        if not _is_file_path(matched_string):
                            hinted, _ = resolve_bare_citation_via_hint(text, matched_string, match_end)
                            if hinted:
                                matched_string = hinted
                        context = clean_context_string(
                            extract_context(text, match_start, match_end, args.context_window)
                        )
                        file_results.append({
                            "Filename": filename,
                            "Page": page_num,
                            "Match Type": "file_citation",
                            "Matched String": matched_string,
                            "Context": context,
                            **bates_fields
                        })

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
    df_bates=pd.DataFrame()
    df_files=pd.DataFrame()

    if link_results:
        # Create dataframe from raw link results
        df_links_raw = pd.DataFrame(link_results)

        # Group by exact URL
        grouped_links = df_links_raw.groupby("Matched String")
        bates_columns = [c for c in ("Bates ID (Footer)",) if c in df_links_raw.columns]

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
        bates_columns = [c for c in ("Bates ID (Footer)",) if c in df_keywords_raw.columns]

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

    if bates_results:
        # Create dataframe from raw body Bates number results
        df_bates_raw = pd.DataFrame(bates_results)

        # Group by exact matched Bates number
        grouped_bates = df_bates_raw.groupby("Matched String")
        bates_columns = [c for c in ("Bates ID (Footer)",) if c in df_bates_raw.columns]

        merged_bates_rows = []

        for bates_number, group in grouped_bates:
            references = group.apply(
                lambda row: {
                    "Filename": row["Filename"],
                    "Page": row["Page"],
                    "Context": row["Context"],
                    **{col: row[col] for col in bates_columns}
                }, axis=1
            ).tolist()

            merged_bates_rows.append({
                "Matched String": bates_number,
                "Reference Count": len(references),
                "References": references
            })

        df_bates = pd.DataFrame(merged_bates_rows)
        df_bates.sort_values(by=["Reference Count", "Matched String"], ascending=[False, True], inplace=True)

    if file_results:
        # Collapse bare filenames into their single matching full-path citation
        # (e.g. "Utils.java" -> ".../com/acme/app/Utils.java") so they no
        # longer appear as separate entries.
        canonical_map = canonicalize_file_citations(
            [r["Matched String"] for r in file_results]
        )
        for r in file_results:
            if r["Matched String"] in canonical_map:
                r["Matched String"] = canonical_map[r["Matched String"]]

        # Create dataframe from raw file-citation results
        df_files_raw = pd.DataFrame(file_results)

        grouped_files = df_files_raw.groupby("Matched String")
        bates_columns = [c for c in ("Bates ID (Footer)",) if c in df_files_raw.columns]

        merged_file_rows = []

        for file_name, group in grouped_files:
            references = group.apply(
                lambda row: {
                    "Filename": row["Filename"],
                    "Page": row["Page"],
                    "Context": row["Context"],
                    **{col: row[col] for col in bates_columns}
                }, axis=1
            ).tolist()

            merged_file_rows.append({
                "Matched String": file_name,
                "Match Type": group["Match Type"].iloc[0],
                "Reference Count": len(references),
                "References": references
            })

        df_files = pd.DataFrame(merged_file_rows)
        df_files.sort_values(by=["Reference Count", "Matched String"], ascending=[False, True], inplace=True)

    json_folder = os.path.join(output_dir, "references_json")
    if not df_links.empty or not df_keywords.empty or not df_bates.empty or not df_files.empty:
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
            if not df_bates.empty:
                df_bates["References"] = df_bates.apply(
                    lambda row: save_large_json(
                        row["References"],
                        base_filename=f"{row['Matched String'][:50].strip().replace('/', '_')}_bates",
                        folder=json_folder
                    ),
                    axis=1
                )
                df_bates.to_excel(writer, sheet_name="Bates Numbers", index=False)
            if not df_files.empty:
                df_files["References"] = df_files.apply(
                    lambda row: save_large_json(
                        row["References"],
                        base_filename=f"{row['Matched String'][:50].strip().replace('/', '_').replace(chr(92), '_')}_files",
                        folder=json_folder
                    ),
                    axis=1
                )
                df_files.to_excel(writer, sheet_name="File Citations", index=False)

        format_excel(excel_path)
        print(f"[OK] Excel saved to {excel_path}")
    else:
        print("[ERR] no references or links found")



if __name__ == "__main__":
    main()
