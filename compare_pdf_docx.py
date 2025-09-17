#!/usr/bin/env python3
"""
compare_pdf_docx.py

Compare a PDF (rendered from LaTeX/Quarto) and a DOCX (judges' corrections)
and produce similarity metrics + a simple HTML diff report.

Dependencies (install with pip):
    pip install pdfminer.six python-docx rapidfuzz nltk

Notes:
 - The script focuses on textual similarity, not exact layout similarity.
 - It reports several complementary metrics:
     * Character-level ratio (difflib.SequenceMatcher)
     * Token-based fuzzy ratio (rapidfuzz.token_sort_ratio)
     * Sentence coverage: how many PDF sentences find a close match in DOCX
 - Outputs a small HTML report with unmatched PDF sentences and a side-by-side diff.
"""

import sys
import re
import argparse
from pathlib import Path
from difflib import SequenceMatcher, HtmlDiff, unified_diff

# The heavy-lifting text extraction functions are imported lazily so the script
# can still be inspected/run for writing without installing all packages.
def extract_text_pdf(path):
    try:
        from pdfminer.high_level import extract_text
    except Exception as e:
        raise RuntimeError("pdfminer.six is required. Install: pip install pdfminer.six") from e
    return extract_text(path)

def extract_text_docx(path):
    try:
        from docx import Document
    except Exception as e:
        raise RuntimeError("python-docx is required. Install: pip install python-docx") from e
    doc = Document(path)
    paragraphs = [p.text for p in doc.paragraphs]
    return "\n\n".join(paragraphs)

def normalize_text(text, remove_page_numbers=True, remove_multiple_newlines=True):
    # Basic normalization: Unicode fixes, unify quotes, collapse spaces, lowercase.
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    # Common unicode quote fixes
    replacements = {
        '\u201c': '"', '\u201d': '"', '\u2018': "'", '\u2019': "'",
        '\u2014': '-', '\u2013': '-',
    }
    for a,b in replacements.items():
        text = text.replace(a,b)
    # Remove running headers/footers patterns like "Page 3", "3/20", or isolated numbers on lines
    if remove_page_numbers:
        # lines that only contain a page number or "page X" or "X / Y"
        text = re.sub(r'(?m)^\s*(page\s*)?\d+\s*(/\s*\d+)?\s*$', '', text, flags=re.IGNORECASE)
        # trailing/leading repeated header/footer short phrases (heuristic)
        text = re.sub(r'(?m)^\s*(Chapter|CHAPTER|Resumen|Abstract|References)\s*$', '', text)
    # Collapse many newlines
    if remove_multiple_newlines:
        text = re.sub(r'\n{3,}', '\n\n', text)
    # Collapse multiple spaces and strip
    text = re.sub(r'[ \t]{2,}', ' ', text)
    text = text.strip()
    return text

def split_sentences(text):
    # Use simple split on double newline or period followed by newline. If nltk is available, use it.
    try:
        import nltk
        nltk.data.find('tokenizers/punkt')
    except Exception:
        try:
            import nltk
            nltk.download('punkt')
        except Exception:
            pass
    try:
        from nltk.tokenize import sent_tokenize
        sents = sent_tokenize(text)
        return [s.strip() for s in sents if s.strip()]
    except Exception:
        # fallback heuristic: split on double newline, then on ". " if necessary
        parts = []
        for block in text.split('\n\n'):
            block = block.strip()
            if not block:
                continue
            if len(block) > 500:
                # further split long blocks by sentence-like punctuation
                parts += re.split(r'(?<=[\.\?\!])\s+', block)
            else:
                parts.append(block)
        return [p.strip() for p in parts if p.strip()]

def compute_similarity_metrics(a, b):
    # Character-level
    char_ratio = SequenceMatcher(None, a, b).ratio() * 100.0
    # Token-based fuzzy ratios (if rapidfuzz available)
    token_ratio = None
    try:
        from rapidfuzz import fuzz
        token_ratio = fuzz.token_sort_ratio(a, b)
    except Exception:
        token_ratio = None
    return {'char_ratio': round(char_ratio,2), 'token_ratio': token_ratio}

def sentence_coverage(pdf_sents, doc_sents, threshold=80):
    """
    For each PDF sentence, find the best matching DOC sentence using rapidfuzz.
    Return count of matches >= threshold and list of unmatched sentences.
    """
    try:
        from rapidfuzz import fuzz
    except Exception:
        raise RuntimeError("rapidfuzz is required for sentence matching. Install: pip install rapidfuzz")
    matched = []
    unmatched = []
    matches = []
    for s in pdf_sents:
        best = 0
        best_j = None
        for j, t in enumerate(doc_sents):
            score = fuzz.token_sort_ratio(s, t)
            if score > best:
                best = score
                best_j = j
        matches.append((s, best, best_j))
        if best >= threshold:
            matched.append((s, best, best_j))
        else:
            unmatched.append((s, best))
    return {'n_pdf': len(pdf_sents), 'n_doc': len(doc_sents),'matched': matched, 'unmatched': unmatched}

def make_html_report(pdf_path, doc_path, metrics, coverage, pdf_text, doc_text, out_path):
    hd = HtmlDiff(tabsize=4, wrapcolumn=80)
    # Create a simple HTML with top metrics then a side-by-side diff of full texts
    html = ['<html><head><meta charset="utf-8"><title>PDF vs DOCX comparison</title></head><body>']
    html.append(f'<h1>Comparison: {Path(pdf_path).name} ↔ {Path(doc_path).name}</h1>')
    html.append('<h2>Metrics</h2><ul>')
    html.append(f'<li>Character-level similarity: {metrics["char_ratio"]}%</li>')
    token_info = metrics["token_ratio"] if metrics["token_ratio"] is not None else "n/a (rapidfuzz not installed)"
    html.append(f'<li>Token-based similarity (token_sort_ratio): {token_info}</li>')
    html.append(f'<li>PDF sentences: {coverage["n_pdf"]}; DOC sentences: {coverage["n_doc"]}</li>')
    html.append(f'<li>Matched PDF sentences (threshold used): {len(coverage["matched"])}</li>')
    html.append(f'<li>Unmatched PDF sentences: {len(coverage["unmatched"])}</li>')
    html.append('</ul>')
    html.append('<h2>Top unmatched PDF sentences (score)</h2><ol>')
    for s, score in sorted(coverage['unmatched'], key=lambda x: x[1]):  # worst first
        snippet = (s[:300] + '...') if len(s) > 300 else s
        html.append(f'<li><b>{score}</b>: {snippet}</li>')
    html.append('</ol>')
    html.append('<h2>Full side-by-side diff (may be large)</h2>')
    html.append(hd.make_table(pdf_text.splitlines(), doc_text.splitlines(), fromdesc="PDF text", todesc="DOCX text", context=True))
    html.append('</body></html>')
    Path(out_path).write_text("\n".join(html), encoding='utf-8')
    return out_path

def main():
    p = argparse.ArgumentParser(description="Compare PDF and DOCX textual similarity")
    p.add_argument("pdf", help="path to PDF (export from Quarto/LaTeX)")
    p.add_argument("docx", help="path to DOCX (judges' file)")
    p.add_argument("--out", default="compare_report.html", help="HTML report path")
    p.add_argument("--sentence-threshold", default=80, type=int, help="Threshold (0-100) to count a PDF sentence as matched")
    args = p.parse_args()

    pdf_path = Path(args.pdf)
    doc_path = Path(args.docx)
    if not pdf_path.exists():
        print("PDF not found:", pdf_path); sys.exit(1)
    if not doc_path.exists():
        print("DOCX not found:", doc_path); sys.exit(1)

    print("Extracting text...")
    pdf_text = extract_text_pdf(str(pdf_path))
    doc_text = extract_text_docx(str(doc_path))

    print("Normalizing text...")
    pdf_text_n = normalize_text(pdf_text)
    doc_text_n = normalize_text(doc_text)

    print("Computing similarity metrics...")
    metrics = compute_similarity_metrics(pdf_text_n, doc_text_n)

    print("Splitting into sentences and computing coverage...")
    pdf_sents = split_sentences(pdf_text_n)
    doc_sents = split_sentences(doc_text_n)
    coverage = sentence_coverage(pdf_sents, doc_sents, threshold=args.sentence_threshold)

    print("Generating HTML report at", args.out)
    report_path = make_html_report(pdf_path, doc_path, metrics, coverage, pdf_text_n, doc_text_n, args.out)
    print("Done. Report:", report_path)
    print("Summary: char_ratio={char_ratio} token_ratio={token_ratio} matched_sentences={matched}/{total}".format(
        char_ratio=metrics['char_ratio'], token_ratio=metrics['token_ratio'],
        matched=len(coverage['matched']), total=coverage['n_pdf']
    ))

if __name__ == "__main__":
    main()
