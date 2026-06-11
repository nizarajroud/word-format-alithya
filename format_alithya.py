"""
format_alithya.py  (v6)
=======================
Génère un document Word formaté Alithya à partir d'un fichier Markdown.

v6 additions:
  - Support des images Markdown: ![alt](path.png)
  - Support des tableaux Markdown: | col | col |
  - Images résolues relativement au fichier content.md

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
# IMAGE TRACKING — collected during parse, embedded after generation
# ─────────────────────────────────────────────────────────────────────────────
_images = []  # list of (rel_id, filepath)
_image_counter = 0

# ─────────────────────────────────────────────────────────────────────────────
# HYPERLINK TRACKING (v6) — collected during parse, added to rels
# ─────────────────────────────────────────────────────────────────────────────
_hyperlinks = []  # list of (rel_id, url)
_hyperlink_counter = 0

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

def hyperlink_run(url, display_text=None):
    """Generate a Word hyperlink with blue underline styling."""
    global _hyperlink_counter
    _hyperlink_counter += 1
    rel_id = f"rIdLink{_hyperlink_counter}"
    _hyperlinks.append((rel_id, url))
    text = display_text or url
    return (
        f'<w:hyperlink r:id="{rel_id}">'
        f'<w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
        f'<w:t{_sp(text)}>{esc(text)}</w:t></w:r></w:hyperlink>'
    )

# ─────────────────────────────────────────────────────────────────────────────
# PARAGRAPH BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

_SPACING = f'<w:spacing w:line="{LINE_SPACING}" w:lineRule="auto"/>'

def heading(text, level=1):
    style = f"Heading{level}"
    ind = '<w:ind w:left="0" w:firstLine="0"/>' if level > 1 else ""
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{_SPACING}{ind}</w:pPr>{run(text)}</w:p>'

def texte(text):
    # Detect inline URLs and convert to hyperlinks
    url_pattern = re.compile(r'(https?://[^\s)<]+)')
    parts_list = url_pattern.split(text)
    if len(parts_list) == 1:
        return f'<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>{_SPACING}</w:pPr>{run(text)}</w:p>'
    runs = ""
    for part in parts_list:
        if url_pattern.match(part):
            runs += hyperlink_run(part)
        elif part:
            runs += run(part)
    return f'<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>{_SPACING}</w:pPr>{runs}</w:p>'

def list_item(text):
    # Detect pattern: "label text (https://url)" → hyperlink with label
    url_paren = re.compile(r'^(.+?)\s*\((https?://[^\s)]+)\)$')
    m = url_paren.match(text)
    if m:
        label = m.group(1)
        url = m.group(2)
        runs = hyperlink_run(url, label)
    else:
        # Detect bare inline URLs
        url_pattern = re.compile(r'(https?://[^\s)<]+)')
        parts_list = url_pattern.split(text)
        if len(parts_list) == 1:
            runs = run(text)
        else:
            runs = ""
            for part in parts_list:
                if url_pattern.match(part):
                    runs += hyperlink_run(part)
                elif part:
                    runs += run(part)
    return (
        '<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{NUMID_LIST}"/></w:numPr>'
        f'{_SPACING}</w:pPr>{runs}</w:p>'
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
# IMAGE SUPPORT (v6)
# ─────────────────────────────────────────────────────────────────────────────

def _get_image_size(filepath):
    """Get image dimensions in EMU (English Metric Units). 1 inch = 914400 EMU."""
    try:
        import struct
        with open(filepath, 'rb') as f:
            header = f.read(32)
        # PNG
        if header[:8] == b'\x89PNG\r\n\x1a\n':
            w = struct.unpack('>I', header[16:20])[0]
            h = struct.unpack('>I', header[20:24])[0]
        # JPEG
        elif header[:2] == b'\xff\xd8':
            f = open(filepath, 'rb')
            f.seek(2)
            while True:
                marker, size = struct.unpack('>HH', f.read(4))
                if marker in (0xFFC0, 0xFFC2):
                    f.read(1)
                    h, w = struct.unpack('>HH', f.read(4))
                    break
                f.seek(size - 2, 1)
            f.close()
        else:
            w, h = 800, 600
    except Exception:
        w, h = 800, 600

    # Scale to max 15cm width (5669291 EMU) maintaining aspect ratio
    max_width_emu = 5669291
    dpi = 96
    width_emu = int(w / dpi * 914400)
    height_emu = int(h / dpi * 914400)
    if width_emu > max_width_emu:
        ratio = max_width_emu / width_emu
        width_emu = max_width_emu
        height_emu = int(height_emu * ratio)
    return width_emu, height_emu

def image_paragraph(filepath):
    """Generate OOXML for an inline image paragraph."""
    global _image_counter
    _image_counter += 1
    rel_id = f"rIdImg{_image_counter}"
    w_emu, h_emu = _get_image_size(filepath)
    _images.append((rel_id, filepath))

    return (
        f'<w:p><w:pPr><w:pStyle w:val="TexteAlithya"/><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr/><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{w_emu}" cy="{h_emu}"/>'
        f'<wp:docPr id="{_image_counter}" name="Image{_image_counter}"/>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{_image_counter}" name="Image{_image_counter}"/>'
        f'<pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w_emu}" cy="{h_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing></w:r></w:p>'
    )

# ─────────────────────────────────────────────────────────────────────────────
# TABLE SUPPORT (v6)
# ─────────────────────────────────────────────────────────────────────────────

def table_xml(rows):
    """Generate OOXML table from parsed markdown rows (list of lists)."""
    num_cols = len(rows[0]) if rows else 0
    col_width = 9000 // max(num_cols, 1)

    xml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>'
    xml += '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    xml += '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    xml += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    xml += '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    xml += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    xml += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    xml += '</w:tblBorders></w:tblPr>'

    # Grid
    xml += '<w:tblGrid>'
    for _ in range(num_cols):
        xml += f'<w:gridCol w:w="{col_width}"/>'
    xml += '</w:tblGrid>'

    for i, row in enumerate(rows):
        xml += '<w:tr>'
        for cell in row:
            cell_text = cell.strip()
            is_header = (i == 0)
            bold_tag = "<w:b/><w:bCs/>" if is_header else ""
            shading = '<w:shd w:val="clear" w:color="auto" w:fill="4472C4"/>' if is_header else ""
            color = '<w:color w:val="FFFFFF"/>' if is_header else ""
            xml += f'<w:tc><w:tcPr><w:tcW w:w="{col_width}" w:type="dxa"/>{shading}</w:tcPr>'
            xml += f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr>{bold_tag}{color}<w:sz w:val="18"/></w:rPr>'
            xml += f'<w:t{_sp(cell_text)}>{esc(cell_text)}</w:t></w:r></w:p></w:tc>'
        xml += '</w:tr>'
    xml += '</w:tbl>'
    return xml

# ─────────────────────────────────────────────────────────────────────────────
# CONTENT PARSER
# ─────────────────────────────────────────────────────────────────────────────

_RE_NUM = re.compile(r"^[\d]+(?:\.[\d]+)*\.?\s+")
_RE_EMOJI = re.compile(
    r"^[\U00002000-\U0001FFFF\U00002300-\U000027BF"
    r"\U0001F000-\U0001FFFF\u2600-\u27BF\uFE00-\uFE0F\u200D]+\s*"
)
_RE_IMAGE = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)$")
_RE_TABLE_ROW = re.compile(r"^\|(.+)\|$")
_RE_TABLE_SEP = re.compile(r"^\|[\s\-:|]+\|$")

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

def parse_content(content, content_dir):
    # Strip YAML frontmatter if present
    if content.startswith('---'):
        end = content.find('---', 3)
        if end > 0:
            content = content[end+3:].strip()

    lines = content.splitlines()
    parts = []
    in_aside = False
    aside_buf = []
    table_buf = []

    # Detect heading levels to normalize
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

    def flush_table():
        nonlocal table_buf
        if not table_buf:
            return
        rows = []
        for tl in table_buf:
            if _RE_TABLE_SEP.match(tl):
                continue
            cols = [c.strip() for c in tl.strip("|").split("|")]
            rows.append(cols)
        if rows:
            parts.append(table_xml(rows))
        table_buf.clear()

    for line in lines:
        s = line.strip()

        # Table continuation
        if table_buf and _RE_TABLE_ROW.match(s):
            table_buf.append(s)
            continue
        elif table_buf:
            flush_table()

        if s in ("<aside>", "[ASIDE]"):
            in_aside = True; continue
        if s in ("</aside>", "[/ASIDE]"):
            flush_aside(); in_aside = False; continue
        if in_aside:
            if s: aside_buf.append(s)
            continue
        if not s:
            continue  # Skip empty lines — no empty paragraphs in output

        # Image
        m_img = _RE_IMAGE.match(s)
        if m_img:
            alt_text = m_img.group(1)
            img_path = m_img.group(2)
            full_path = os.path.join(content_dir, img_path)
            if os.path.exists(full_path):
                parts.append(image_paragraph(full_path))
            else:
                parts.append(texte(f"[Image: {alt_text} — fichier non trouvé: {img_path}]"))
            continue

        # Table start
        if _RE_TABLE_ROW.match(s):
            table_buf.append(s)
            continue

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

    # Flush remaining table
    flush_table()
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

    content_dir = os.path.dirname(os.path.abspath(content_file))

    with open(content_file, encoding="utf-8") as f:
        content = f.read()

    with zipfile.ZipFile(TEMPLATE, "r") as z:
        orig_doc = z.read("word/document.xml").decode("utf-8")

    # Parse content (also collects images)
    generated = parse_content(content, content_dir)
    new_doc = orig_doc.replace("<!-- CONTENT -->", generated)

    # Replace cover page title and date (safe text-only replacement)
    from datetime import datetime
    # Extract title from YAML frontmatter or first heading
    title_match = re.search(r'^title:\s*(.+)$', content, re.MULTILINE)
    if not title_match:
        title_match = re.search(r'^# (.+)$', content, re.MULTILINE)
    doc_title = title_match.group(1).strip() if title_match else "Document"
    today_date = datetime.now().strftime("%-d %B %Y").replace(
        "January","janvier").replace("February","février").replace("March","mars").replace(
        "April","avril").replace("May","mai").replace("June","juin").replace(
        "July","juillet").replace("August","août").replace("September","septembre").replace(
        "October","octobre").replace("November","novembre").replace("December","décembre")

    # Replace the exact template title text (safe: replace full run content)
    new_doc = new_doc.replace(
        '>Analyse</w:t>',
        f'>{esc(doc_title)}</w:t>', 1
    )
    new_doc = new_doc.replace(
        '> - Migration et Modernisation CDBAC</w:t>',
        '></w:t>', 1
    )
    new_doc = new_doc.replace(
        '>-BNC</w:t>',
        '></w:t>', 1
    )
    # Replace date
    new_doc = new_doc.replace(
        '>12</w:t>',
        f'>{today_date.split(" ")[0]}</w:t>', 1
    )
    new_doc = new_doc.replace(
        '> janvier 2026</w:t>',
        f'> {" ".join(today_date.split(" ")[1:])}</w:t>', 1
    )

    # Ensure drawing namespaces are declared
    if _images or _hyperlinks:
        if 'xmlns:wp=' not in new_doc:
            new_doc = new_doc.replace(
                '<w:document ',
                '<w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
                1
            )
        if 'xmlns:r=' not in new_doc.split('<w:body')[0]:
            new_doc = new_doc.replace(
                '<w:document ',
                '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
                1
            )

    if os.path.exists(output_file):
        os.remove(output_file)

    with zipfile.ZipFile(TEMPLATE, "r") as zin, \
         zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zout:

        # Read existing rels
        rels_path = "word/_rels/document.xml.rels"
        rels_xml = zin.read(rels_path).decode("utf-8") if rels_path in zin.namelist() else ""

        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc.encode("utf-8")
            elif item.filename == rels_path and (_images or _hyperlinks):
                # Add image and hyperlink relationships
                insert_point = rels_xml.rfind("</Relationships>")
                new_rels = ""
                for rel_id, filepath in _images:
                    ext = os.path.splitext(filepath)[1].lower()
                    img_name = f"media/{rel_id}{ext}"
                    new_rels += f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{img_name}"/>'
                for rel_id, url in _hyperlinks:
                    new_rels += f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{esc(url)}" TargetMode="External"/>'
                rels_xml = rels_xml[:insert_point] + new_rels + rels_xml[insert_point:]
                data = rels_xml.encode("utf-8")
            elif item.filename == "[Content_Types].xml" and _images:
                ct_xml = data.decode("utf-8")
                for ext in set(os.path.splitext(fp)[1].lower().lstrip('.') for _, fp in _images):
                    content_type = "image/png" if ext == "png" else "image/jpeg"
                    if f'Extension="{ext}"' not in ct_xml:
                        ct_xml = ct_xml.replace(
                            "</Types>",
                            f'<Default Extension="{ext}" ContentType="{content_type}"/></Types>'
                        )
                data = ct_xml.encode("utf-8")
            elif item.filename == "word/settings.xml":
                # Force TOC update on open
                settings_xml = data.decode("utf-8")
                if '<w:updateFields' not in settings_xml:
                    settings_xml = settings_xml.replace(
                        '</w:settings>',
                        '<w:updateFields w:val="true"/></w:settings>'
                    )
                data = settings_xml.encode("utf-8")
            zout.writestr(item, data)

        # Add image files to the docx archive
        for rel_id, filepath in _images:
            ext = os.path.splitext(filepath)[1].lower()
            img_name = f"word/media/{rel_id}{ext}"
            with open(filepath, "rb") as img_f:
                zout.writestr(img_name, img_f.read())

    print(f"✅  Document généré : {output_file}")
    if _images:
        print(f"    📷  {len(_images)} image(s) intégrée(s)")
    print(f"    ⚠️  Ouvrir dans Word et faire Ctrl+A → F9 pour mettre à jour la table des matières")

if __name__ == "__main__":
    main()
