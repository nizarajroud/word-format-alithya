"""
format_alithya_v2.py
====================
Génère un document Word formaté Alithya à partir d'une page Notion.

Usage : python format_alithya_v2.py <notion_url> [output.docx]
"""

import os, re, sys, zipfile, json, urllib.request
from dotenv import load_dotenv

load_dotenv()

TEMPLATE     = os.getenv("TEMPLATE", "template_ref.docx")
OUTPUT       = os.getenv("OUTPUT", "generated.docx")
NUMID_LIST   = os.getenv("NUMID_LIST", "5")
LINE_SPACING = os.getenv("LINE_SPACING", "360")
NOTION_TOKEN = os.getenv("NOTION_TOKEN")

# ── XML helpers ──────────────────────────────────────────────────────────────

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

# ── Paragraph builders ───────────────────────────────────────────────────────

_SPACING = f'<w:spacing w:line="{LINE_SPACING}" w:lineRule="auto"/>'

def heading(text, level=1):
    style = f"Heading{level}"
    ind = '<w:ind w:left="0" w:firstLine="0"/>' if level > 1 else ""
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{_SPACING}{ind}</w:pPr>{run(text)}</w:p>'

def texte(text):
    return f'<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>{_SPACING}</w:pPr>{run(text)}</w:p>'

def texte_runs(runs_xml):
    return f'<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>{_SPACING}</w:pPr>{runs_xml}</w:p>'

def list_item(runs_xml):
    return (
        '<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{NUMID_LIST}"/></w:numPr>'
        f'{_SPACING}</w:pPr>{runs_xml}</w:p>'
    )

def numbered_item(runs_xml):
    return list_item(runs_xml)  # same style, Word numbering handles it

def empty_para():
    return '<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/></w:pPr></w:p>'

# ── Notion API ───────────────────────────────────────────────────────────────

def notion_api(endpoint):
    url = f"https://api.notion.com/v1/{endpoint}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": "2022-06-28",
    })
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())

def extract_page_id(url):
    """Extract 32-char hex page ID from a Notion URL."""
    # Remove shell escapes and query string
    clean = url.replace("\\", "").split("?")[0].rstrip("/")
    m = re.search(r"([0-9a-f]{32})$", clean)
    if not m:
        print(f"❌  Impossible d'extraire l'ID de la page : {url}")
        sys.exit(1)
    raw = m.group(1)
    return f"{raw[:8]}-{raw[8:12]}-{raw[12:16]}-{raw[16:20]}-{raw[20:]}"

def get_all_blocks(page_id):
    blocks = []
    cursor = None
    while True:
        ep = f"blocks/{page_id}/children?page_size=100"
        if cursor:
            ep += f"&start_cursor={cursor}"
        data = notion_api(ep)
        blocks.extend(data["results"])
        if not data.get("has_more"):
            break
        cursor = data["next_cursor"]
    return blocks

# ── Rich text → Word XML runs ───────────────────────────────────────────────

_RE_EMOJI = re.compile(
    r"^[\U00002000-\U0001FFFF\U00002300-\U000027BF"
    r"\U0001F000-\U0001FFFF\u2600-\u27BF\uFE00-\uFE0F\u200D]+\s*"
)
_RE_NUM = re.compile(r"^[\d]+(?:\.[\d]+)*\.?\s+")

def clean_heading_text(text):
    text = _RE_EMOJI.sub("", text).strip()
    text = _RE_NUM.sub("", text).strip()
    return text

def rich_text_to_runs(rt_list):
    """Convert Notion rich_text array to Word XML runs."""
    parts = []
    for rt in rt_list:
        text = rt.get("plain_text", "")
        if not text:
            continue
        ann = rt.get("annotations", {})
        bold = ann.get("bold", False)
        italic = ann.get("italic", False)
        parts.append(run(text, bold=bold, italic=italic))
    return "".join(parts)

def rich_text_plain(rt_list):
    return "".join(rt.get("plain_text", "") for rt in rt_list)

# ── Notion blocks → Word XML ────────────────────────────────────────────────

def blocks_to_xml(blocks):
    # Detect heading levels for normalization
    levels = set()
    for b in blocks:
        if b["type"] == "heading_1": levels.add(1)
        elif b["type"] == "heading_2": levels.add(2)
        elif b["type"] == "heading_3": levels.add(3)
    lvl_map = {lv: i + 1 for i, lv in enumerate(sorted(levels))}

    parts = []
    for b in blocks:
        btype = b["type"]

        if btype in ("heading_1", "heading_2", "heading_3"):
            md_level = int(btype[-1])
            rt = b[btype]["rich_text"]
            text = clean_heading_text(rich_text_plain(rt))
            parts.append(heading(text, lvl_map.get(md_level, md_level)))

        elif btype == "paragraph":
            rt = b["paragraph"]["rich_text"]
            if not rt:
                parts.append(empty_para())
            else:
                parts.append(texte_runs(rich_text_to_runs(rt)))

        elif btype == "bulleted_list_item":
            rt = b["bulleted_list_item"]["rich_text"]
            parts.append(list_item(rich_text_to_runs(rt)))

        elif btype == "numbered_list_item":
            rt = b["numbered_list_item"]["rich_text"]
            parts.append(numbered_item(rich_text_to_runs(rt)))

        elif btype == "callout":
            rt = b["callout"]["rich_text"]
            parts.append(texte_runs(rich_text_to_runs(rt)))

        elif btype == "divider":
            parts.append(empty_para())

        # skip images, unsupported blocks

    return "\n".join(parts)

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python format_alithya_v2.py <notion_url> [output.docx]")
        sys.exit(1)

    if not NOTION_TOKEN:
        print("❌  NOTION_TOKEN manquant dans .env")
        sys.exit(1)

    notion_url  = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else OUTPUT

    page_id = extract_page_id(notion_url)
    print(f"📄  Page Notion : {page_id}")

    blocks = get_all_blocks(page_id)
    print(f"📦  {len(blocks)} blocs récupérés")

    generated = blocks_to_xml(blocks)

    with zipfile.ZipFile(TEMPLATE, "r") as z:
        orig_doc = z.read("word/document.xml").decode("utf-8")

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
