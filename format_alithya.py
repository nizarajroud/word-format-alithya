"""
format_alithya.py  (v4)
=======================
Formate du contenu de content.md selon le template Alithya (BIXI Q&R).

Structure du document généré :
  - Pages de contenu  → header VIDE (aucun background)
  - Dernière page     → background Alithya complet (header4 du template)

Usage : python format_alithya.py
        python format_alithya.py mon_contenu.md          (fichier custom)
        python format_alithya.py content.md output.docx  (fichier + output custom)
"""

import os, re, zipfile, sys

# ─────────────────────────────────────────────────────────────────────────────
TEMPLATE = "template.docx"
OUTPUT   = "Questions-Réponses-BIXI-V2.docx"
CONTENT_FILE = "content.md"

# ─────────────────────────────────────────────────────────────────────────────
# XML HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def esc(text):
    return text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def xml_space(text):
    return ' xml:space="preserve"' if text and (text[0]==" " or text[-1]==" ") else ""

def run_xml(text, bold=False, italic=False, color=None, font="Arial Narrow", sz=20):
    parts = [f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>']
    if bold:   parts.append("<w:b/><w:bCs/>")
    if italic: parts.append("<w:i/>")
    if color:  parts.append(f'<w:color w:val="{color}"/>')
    parts.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    parts.append('<w:lang w:val="fr-CA"/>')
    sp = xml_space(text)
    return f'<w:r><w:rPr>{"".join(parts)}</w:rPr><w:t{sp}>{esc(text)}</w:t></w:r>'

# ─────────────────────────────────────────────────────────────────────────────
# PARAGRAPH BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def heading1(text):
    r = run_xml(text, bold=True, color="1F96E3", font="Arial", sz=28)
    return (
        "<w:p><w:pPr>"
        '<w:pStyle w:val="Titre1"/>'
        "<w:numPr>"
        '<w:ilvl w:val="0"/>'
        '<w:numId w:val="13"/>'
        "</w:numPr>"
        '<w:spacing w:line="360" w:lineRule="auto"/>'
        f"</w:pPr>{r}</w:p>"
    )

def heading2(text):
    r = run_xml(text, bold=True, color="1F96E3", font="Arial", sz=24)
    return (
        "<w:p><w:pPr>"
        '<w:pStyle w:val="Titre2"/>'
        '<w:spacing w:line="360" w:lineRule="auto"/>'
        f"</w:pPr>{r}</w:p>"
    )

def list_item(text):
    runs = run_xml(text, font="Arial Narrow", color="002957", sz=20)
    return (
        "<w:p><w:pPr>"
        '<w:pStyle w:val="Paragraphedeliste"/>'
        "<w:numPr>"
        '<w:ilvl w:val="0"/>'
        '<w:numId w:val="5"/>'
        "</w:numPr>"
        '<w:spacing w:line="360" w:lineRule="auto"/>'
        f"</w:pPr>{runs}</w:p>"
    )

def list_item_split(label, rest):
    """List item with a bold label followed by normal text."""
    # label may already have brackets e.g. "[Inventaire]" or just "Inventaire"
    if not (label.startswith("[") and label.endswith("]")):
        label = f"[{label}]"
    runs = (
        run_xml(label + " ", bold=True, font="Arial Narrow", color="002957", sz=20) +
        run_xml(rest, font="Arial Narrow", color="002957", sz=20)
    )
    return (
        "<w:p><w:pPr>"
        '<w:pStyle w:val="Paragraphedeliste"/>'
        "<w:numPr>"
        '<w:ilvl w:val="0"/>'
        '<w:numId w:val="5"/>'
        "</w:numPr>"
        '<w:spacing w:line="360" w:lineRule="auto"/>'
        f"</w:pPr>{runs}</w:p>"
    )

def aside_para(text, is_first=False, is_last=False):
    before = "120" if is_first else "40"
    after  = "120" if is_last  else "40"
    r = run_xml(text, italic=True, font="Arial Narrow", color="002957", sz=20)
    return (
        "<w:p><w:pPr>"
        "<w:pBdr>"
        '<w:left w:val="single" w:sz="24" w:space="4" w:color="1F96E3"/>'
        "</w:pBdr>"
        '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
        f'<w:spacing w:before="{before}" w:after="{after}" w:line="360" w:lineRule="auto"/>'
        '<w:ind w:left="360" w:right="360"/>'
        f"</w:pPr>{r}</w:p>"
    )

def normal_para(text):
    r = run_xml(text, font="Arial Narrow", color="002957", sz=20)
    return (
        "<w:p><w:pPr>"
        '<w:spacing w:line="360" w:lineRule="auto"/>'
        f"</w:pPr>{r}</w:p>"
    )

def empty_para():
    return '<w:p><w:pPr><w:spacing w:after="80"/></w:pPr></w:p>'

# ─────────────────────────────────────────────────────────────────────────────
# CONTENT PARSER
# ─────────────────────────────────────────────────────────────────────────────

def strip_bold(text):
    """Remove **bold** markers from text, returning (label, rest) if bold prefix found."""
    # Match **[Tag]** or **Tag** at start of string
    m = re.match(r"^\*\*(.+?)\*\*\s*(.*)", text, re.DOTALL)
    if m:
        return m.group(1), m.group(2)
    return None, text

def strip_inline_md(text):
    """Strip inline markdown: **bold**, *italic*, `code`."""
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*",     r"\1", text)
    text = re.sub(r"`(.+?)`",        r"\1", text)
    return text

def strip_emoji(text):
    """Remove leading emoji characters from a string."""
    return re.sub(
        r"^[\U00002000-\U0001FFFF\U00002300-\U000027BF"
        r"\U0001F000-\U0001FFFF\u2600-\u27BF\uFE00-\uFE0F\u200D]+\s*",
        "", text
    ).strip()

def parse_content(content):
    parts = []
    lines = content.splitlines()
    in_aside = False
    aside_lines = []

    def flush_aside():
        nonlocal aside_lines
        if not aside_lines:
            return
        n = len(aside_lines)
        for i, al in enumerate(aside_lines):
            if al.startswith("|") and al.endswith("|"):
                cols = [c.strip() for c in al.strip("|").split("|")]
                parts.append(aside_para("  |  ".join(cols), is_first=(i==0), is_last=(i==n-1)))
            else:
                # Strip inline markdown inside aside
                parts.append(aside_para(strip_inline_md(al), is_first=(i==0), is_last=(i==n-1)))
        aside_lines.clear()

    for line in lines:
        s = line.strip()

        # ── Aside: HTML-style <aside> or custom [ASIDE] ──
        if s in ("<aside>", "[ASIDE]"):
            in_aside = True; continue
        if s in ("</aside>", "[/ASIDE]"):
            flush_aside(); in_aside = False
            parts.append(empty_para()); continue
        if in_aside:
            if s: aside_lines.append(s)
            continue

        if not s:
            parts.append(empty_para()); continue

        # ── Headings: ### or ## or # (strip leading emoji) ──
        if s.startswith("### "):
            parts.append(heading1(strip_emoji(s[4:].strip())))
        elif s.startswith("## "):
            parts.append(heading1(strip_emoji(s[3:].strip())))
        elif s.startswith("# "):
            parts.append(heading1(strip_emoji(s[2:].strip())))

        # ── Numbered list: "1. **[Tag]** text" or "1. text" ──
        elif re.match(r"^\d+\.\s+", s):
            body = re.match(r"^\d+\.\s+(.*)", s).group(1)
            label, rest = strip_bold(body)
            if label:
                parts.append(list_item_split(label, rest))
            else:
                parts.append(list_item(strip_inline_md(body)))

        # ── Normal paragraph ──
        else:
            parts.append(normal_para(strip_inline_md(s)))

    flush_aside()
    return "\n".join(parts)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # Optional CLI arguments: content_file, output_file
    content_file = sys.argv[1] if len(sys.argv) > 1 else CONTENT_FILE
    output_file  = sys.argv[2] if len(sys.argv) > 2 else OUTPUT

    # Read content from .md file
    if not os.path.exists(content_file):
        print(f"❌  Fichier introuvable : {content_file}")
        print(f"    Crée un fichier '{content_file}' avec ton contenu.")
        sys.exit(1)
    with open(content_file, encoding="utf-8") as f:
        content = f.read()

    # Read original document.xml
    with zipfile.ZipFile(TEMPLATE, "r") as z:
        orig_doc = z.read("word/document.xml").decode("utf-8")

    # ── Split template body into individual paragraphs ────────────────────────
    # Template has 3 sections. Each section ends with a paragraph whose <w:pPr>
    # contains a <w:sectPr>. The final (body) sectPr sits OUTSIDE any paragraph.
    #
    # Para indices (0-based):
    #   0–21   → Section 1: first-page cover    ← KEEP as-is
    #   22     → Section 1 break paragraph (pPr contains sectPr, rId11/12/13)
    #   23–29  → Section 2: title/date page     ← KEEP as-is
    #   30     → Section 2 break paragraph (pPr contains sectPr, rId14/15)
    #   31+    → Section 3: last page + body sectPr  ← KEEP as-is
    #
    # Strategy: keep sections 1 & 2 intact, inject generated content
    # right after para 30, then append section 3 unchanged.
    # Generated pages reuse sec2's header/footer (rId14=logo header, rId15=page-num footer)
    # by cloning the sec2 break paragraph as the section-end for generated content.
    # ─────────────────────────────────────────────────────────────────────────

    body_start = orig_doc.find("<w:body>")
    body_end   = orig_doc.find("</w:body>")
    doc_header = orig_doc[:body_start]

    # Extract full body (includes trailing body sectPr)
    body = orig_doc[body_start + len("<w:body>"):body_end]

    # Split into paragraph tokens (each starts with <w:p)
    # The body sectPr at the very end is NOT in a <w:p>, so it appears as a trailing token
    paras = re.split(r'(?=<w:p[ >])', body)
    # paras[0] may be empty or whitespace, paras[1..N] are paragraphs
    # The last token after para 31 contains the body sectPr

    # Find section-break paragraphs (those whose pPr contains a sectPr)
    sec_break_indices = [
        i for i, p in enumerate(paras)
        if re.search(r'<w:pPr>.*?<w:sectPr', p, re.DOTALL)
    ]
    # sec_break_indices[0] = end of section 1 (cover page break para)
    # sec_break_indices[1] = end of section 2 (content pages break para)
    idx_sec1_break = sec_break_indices[0]
    idx_sec2_break = sec_break_indices[1]

    # Section 1: everything up to and including the section 1 break paragraph
    section1 = "".join(paras[0:idx_sec1_break + 1])

    # Section 2 body (paras between sec1 break and sec2 break) may contain
    # leftover empty paragraphs from the template (e.g. deleted content).
    # We DROP them entirely so there is no blank page between cover and content.

    # The sec2 break paragraph carries rId14/rId15 (logo header + page-number footer).
    # We reuse it verbatim as the section-end for our generated content pages.
    sec2_break_para = paras[idx_sec2_break]

    # Section 3 (last page + body sectPr): everything after sec2 break
    section3 = "".join(paras[idx_sec2_break + 1:])

    # ── Assemble final document ───────────────────────────────────────────────
    generated = parse_content(content)

    new_doc = (
        doc_header
        + "<w:body>"
        + section1           # cover page — untouched
        + generated          # generated content starts immediately on page 2
        + sec2_break_para    # section break: logo header + page-number footer
        + section3           # last page — untouched
        + "</w:body></w:document>"
    )

    # ── Write output (all files from template, only document.xml replaced) ───
    if os.path.exists(output_file):
        os.remove(output_file)

    with zipfile.ZipFile(TEMPLATE, "r") as zin, \
         zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, new_doc.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))

    print(f"✅  Document généré : {output_file}")

if __name__ == "__main__":
    main()