"""
format_alithya.py  (v5)
=======================
Génère un document Word formaté Alithya à partir d'un fichier Markdown.

Utilise template_ref.docx (extrait du document de référence) qui contient :
  - Section 0 : Page de couverture
  - Section 1 : Table des matières (champ TOC)
  - Section 2 : Placeholder contenu (<!-- CONTENT -->)
  - Section 3 : Dernière page

Les styles Heading1/2/3, TexteAlithya et ListParagraph sont définis dans le
template. La numérotation hiérarchique (1, 1.1, 1.1.1) est gérée par
abstractNum 95 via numId=1.

Usage : python format_alithya.py
        python format_alithya.py content.md
        python format_alithya.py content.md output.docx
"""

import os, re, zipfile, sys
from dotenv import load_dotenv

load_dotenv()

# ─────────────────────────────────────────────────────────────────────────────
TEMPLATE     = os.getenv("TEMPLATE", "template_ref.docx")
OUTPUT       = os.getenv("OUTPUT", "generated.docx")
CONTENT_FILE = os.getenv("CONTENT_FILE", "content.md")
NUMID_LIST   = os.getenv("NUMID_LIST", "5")
LINE_SPACING = os.getenv("LINE_SPACING", "360")

# ─────────────────────────────────────────────────────────────────────────────
# XML HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def esc(t):
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def _sp(t):
    return ' xml:space="preserve"' if t and (t[0]==" " or t[-1]==" ") else ""

def run(text, bold=False, italic=False):
    rpr = ""
    if bold or italic:
        parts = []
        if bold:   parts.append("<w:b/><w:bCs/>")
        if italic: parts.append("<w:i/>")
        rpr = f"<w:rPr>{''.join(parts)}</w:rPr>"
    return f'<w:r>{rpr}<w:t{_sp(text)}>{esc(text)}</w:t></w:r>'

# ─────────────────────────────────────────────────────────────────────────────
# PARAGRAPH BUILDERS — rely on template styles for all formatting
# ─────────────────────────────────────────────────────────────────────────────

_SPACING = f'<w:spacing w:line="{LINE_SPACING}" w:lineRule="auto"/>'

def heading(text, level=1):
    style = f"Heading{level}"
    ind = '<w:ind w:left="0" w:firstLine="0"/>' if level > 1 else ""
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{_SPACING}{ind}</w:pPr>{run(text)}</w:p>'

def texte(text):
    return f'<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>{_SPACING}</w:pPr>{run(text)}</w:p>'

def list_item(text):
    return (
        '<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{NUMID_LIST}"/></w:numPr>'
        f'{_SPACING}</w:pPr>{run(text)}</w:p>'
    )

def list_item_split(label, rest):
    if not (label.startswith("[") and label.endswith("]")):
        label = f"[{label}]"
    return (
        '<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{NUMID_LIST}"/></w:numPr>'
        f'{_SPACING}</w:pPr>{run(label + " ", bold=True)}{run(rest)}</w:p>'
    )

def empty_para():
    return '<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/></w:pPr></w:p>'

# ─────────────────────────────────────────────────────────────────────────────
# CONTENT PARSER
# ─────────────────────────────────────────────────────────────────────────────

_RE_NUM = re.compile(r"^[\d]+(?:\.[\d]+)*\.?\s+")
_RE_EMOJI = re.compile(
    r"^[\U00002000-\U0001FFFF\U00002300-\U000027BF"
    r"\U0001F000-\U0001FFFF\u2600-\u27BF\uFE00-\uFE0F\u200D]+\s*"
)

def clean_heading(text):
    text = _RE_EMOJI.sub("", text).strip()
    text = _RE_NUM.sub("", text).strip()
    return text

def strip_bold(text):
    m = re.match(r"^\*\*(.+?)\*\*\s*(.*)", text, re.DOTALL)
    return (m.group(1), m.group(2)) if m else (None, text)

def strip_md(text):
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    return text

def parse_content(content):
    lines = content.splitlines()
    parts = []
    in_aside = False
    aside_buf = []

    # Detect heading levels to normalize (e.g. if only ### used, map to level 1)
    md_levels = set()
    for line in lines:
        s = line.strip()
        if s.startswith("### "):   md_levels.add(3)
        elif s.startswith("## "):  md_levels.add(2)
        elif s.startswith("# "):   md_levels.add(1)
    lvl_map = {md: i + 1 for i, md in enumerate(sorted(md_levels))}

    def flush_aside():
        nonlocal aside_buf
        for al in aside_buf:
            if al.startswith("|") and al.endswith("|"):
                cols = [c.strip() for c in al.strip("|").split("|")]
                parts.append(texte("  |  ".join(cols)))
            else:
                parts.append(texte(strip_md(al)))
        aside_buf.clear()

    for line in lines:
        s = line.strip()

        if s in ("<aside>", "[ASIDE]"):
            in_aside = True; continue
        if s in ("</aside>", "[/ASIDE]"):
            flush_aside(); in_aside = False; continue
        if in_aside:
            if s: aside_buf.append(s)
            continue
        if not s:
            parts.append(empty_para()); continue

        # Headings
        if s.startswith("### "):
            parts.append(heading(clean_heading(s[4:]), lvl_map.get(3, 3)))
        elif s.startswith("## "):
            parts.append(heading(clean_heading(s[3:]), lvl_map.get(2, 2)))
        elif s.startswith("# "):
            parts.append(heading(clean_heading(s[2:]), lvl_map.get(1, 1)))
        # Numbered list
        elif re.match(r"^\d+\.\s+", s):
            body = re.match(r"^\d+\.\s+(.*)", s).group(1)
            label, rest = strip_bold(body)
            if label:
                parts.append(list_item_split(label, rest))
            else:
                parts.append(list_item(strip_md(body)))
        # Normal paragraph
        else:
            parts.append(texte(strip_md(s)))

    flush_aside()
    return "\n".join(parts)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    content_file = sys.argv[1] if len(sys.argv) > 1 else CONTENT_FILE
    output_file  = sys.argv[2] if len(sys.argv) > 2 else OUTPUT

    if not os.path.exists(content_file):
        print(f"❌  Fichier introuvable : {content_file}")
        sys.exit(1)

    with open(content_file, encoding="utf-8") as f:
        content = f.read()

    with zipfile.ZipFile(TEMPLATE, "r") as z:
        orig_doc = z.read("word/document.xml").decode("utf-8")

    # Replace the <!-- CONTENT --> placeholder with generated paragraphs
    generated = parse_content(content)
    new_doc = orig_doc.replace("<!-- CONTENT -->", generated)

    if os.path.exists(output_file):
        os.remove(output_file)

    with zipfile.ZipFile(TEMPLATE, "r") as zin, \
         zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc.encode("utf-8")
            zout.writestr(item, data)

    print(f"✅  Document généré : {output_file}")
    print(f"    ⚠️  Ouvrir dans Word et faire Ctrl+A → F9 pour mettre à jour la table des matières")

if __name__ == "__main__":
    main()
