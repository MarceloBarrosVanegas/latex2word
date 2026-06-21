#!/usr/bin/env python3
r"""
LaTeX to Word (.docx) Converter
================================
Converts a LaTeX .tex file to a Word document preserving:
  - Title page  (title image, title, author, date)
  - Header / footer (logo + project name | page number)
  - Sections, subsections, subsubsections, paragraphs
  - Custom \newcommand macros and \setcounter values
  - Paragraphs with bold / italic / inline math / hyperlinks
  - Itemize and enumerate lists (including custom labels)
  - longtable and tabularx/tabular environments
  - Display math  \[ ... \]  rendered as centred italic text
  - Page breaks (\newpage, \clearpage)
  - Table of contents placeholder

Usage:
    python latex_to_docx.py [input.tex] [output.docx]
    Defaults: linea_base_en.tex  ->  linea_base_en.docx
"""

import re
import sys
import subprocess
import shutil
import tempfile
import unicodedata
from pathlib import Path
from copy import deepcopy

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from lxml import etree

# ─────────────────────────────────────────────────────────────────────────────
# Colours
# ─────────────────────────────────────────────────────────────────────────────
C_BLACK   = RGBColor(0x00, 0x00, 0x00)
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
C_LINK    = RGBColor(0x00, 0x56, 0xB3)
C_GRAY    = RGBColor(0x60, 0x60, 0x60)

HEX_HDR   = "1F497D"   # dark blue for table header rows
HEX_CAT   = "DDEEFF"   # light blue for category rows inside tables
HEX_ALT   = "F5F9FF"   # very light blue alternating row (unused but available)

FONT      = "Calibri"
PT_NORM   = 11
PT_SMALL  =  9

# Carpeta temporal para archivos intermedios (tablas, etc.)
TEMP_DIR: Path = Path(".")  # Se establece en main()
PT_FOOT   =  8

PANDOC_PATH: str = "pandoc"  # Se establece en main()
MATH_NS: str = "http://schemas.openxmlformats.org/officeDocument/2006/math"
OMML_CACHE: dict[str, object] = {}

# Ancho de texto para este documento (A4, margin=1in → 8.27 - 2 = 6.27in)
TEXT_WIDTH_INCHES  = 6.27
# Alto aproximado del cuerpo de texto (A4, margin ~1in arriba/abajo → 11.69 - 2 ≈ 9.5in)
TEXT_HEIGHT_INCHES = 9.50

# ─────────────────────────────────────────────────────────────────────────────
# XML helpers
# ─────────────────────────────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _set_cell_borders(cell, color="BBBBBB", sz="4"):
    tc    = cell._tc
    tcPr  = tc.get_or_add_tcPr()
    tcBdr = OxmlElement("w:tcBorders")
    for edge in ("top", "bottom", "left", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    sz)
        el.set(qn("w:color"), color)
        tcBdr.append(el)
    tcPr.append(tcBdr)


def _add_page_number_field(run):
    """Insert a PAGE field into a run (with proper begin/separate/end structure)."""
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)

    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " PAGE "
    run._r.append(instr)

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_sep)

    # Placeholder display text (updated by Word on open)
    t = OxmlElement("w:t")
    t.text = "1"
    run._r.append(t)

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_end)


def _add_tc_field(paragraph, text: str, identifier: str):
    """Add a hidden TC (Table of Contents Entry) field to a paragraph."""
    safe_text = text.replace('"', "'")

    run1 = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run1._r.append(fld_begin)

    run2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = f' TC "{safe_text}" \\f {identifier} '
    run2._r.append(instr)

    run3 = paragraph.add_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run3._r.append(fld_sep)

    run4 = paragraph.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run4._r.append(fld_end)


def _safe_page_break(doc: Document):
    """Add a page break only if the document does not already end with one.
    Also strips trailing empty paragraphs to avoid blank pages."""
    # Remove trailing completely empty paragraphs (no text, no runs)
    while doc.paragraphs:
        last_p = doc.paragraphs[-1]
        has_content = bool(last_p.text.strip()) or len(last_p.runs) > 0
        if not has_content:
            last_p._element.getparent().remove(last_p._element)
        else:
            break

    # Check if the last paragraph already ends with a page break
    if doc.paragraphs:
        last_p = doc.paragraphs[-1]
        for run in last_p.runs:
            for br in run._r.findall(qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    return  # Already a page break, skip

    doc.add_page_break()


def _insert_toc_field(doc: Document, instr: str, title: str = "", placeholder: str = ""):
    """Insert a Word TOC/LOT/LOF field."""
    from docx.oxml import parse_xml

    # NOTE: We intentionally do NOT set w:updateFields to true because that
    # triggers a Word dialog on every open asking whether to update fields.
    # The user can update once with Ctrl+A → F9, save, and the document
    # will open cleanly afterwards.

    if title:
        p = doc.add_paragraph()
        run = _run(p, title, bold=True, size=13)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()

    toc_para = doc.add_paragraph()
    toc_para.paragraph_format.space_before = Pt(12)
    toc_para.paragraph_format.space_after = Pt(12)

    toc_para._p.append(parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:fldChar w:fldCharType="begin"/></w:r>'
    ))
    toc_para._p.append(parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
    ))
    toc_para._p.append(parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:fldChar w:fldCharType="separate"/></w:r>'
    ))
    ph_text = placeholder or "Right-click and select Update Field to generate."
    toc_para._p.append(parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:rPr><w:color w:val="808080"/><w:i/></w:rPr>'
        f'<w:t>{ph_text}</w:t></w:r>'
    ))
    toc_para._p.append(parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:fldChar w:fldCharType="end"/></w:r>'
    ))


def add_toc(doc: Document, title: str = "TABLE OF CONTENTS"):
    """Insert a Table of Contents field (auto-updates on open)."""
    _insert_toc_field(
        doc, r'TOC \o "1-3" \h \z', title,
        "No table of contents entries found. Right-click and select Update Field to generate."
    )


def add_lof(doc: Document, title: str = "LIST OF FIGURES"):
    """Insert a List of Figures field (auto-updates on open)."""
    _insert_toc_field(
        doc, r'TOC \c "Figure" \h \z', title,
        "No figure entries found. Right-click and select Update Field to generate."
    )


def add_lot(doc: Document, title: str = "LIST OF TABLES"):
    """Insert a List of Tables field (auto-updates on open)."""
    _insert_toc_field(
        doc, r'TOC \c "Table" \h \z', title,
        "No table entries found. Right-click and select Update Field to generate."
    )

def add_table_caption(doc: Document, caption_text: str, table_num: int = 0):
    """
    Add a table caption with a hidden TC field for Word's List of Tables.
    Format: Table X: Caption text
    """
    clean_caption = strip_fmt(caption_text).replace('"', "'")
    full_caption = f"Table {table_num}: {clean_caption}" if table_num > 0 else f"Table: {clean_caption}"
    
    # Simple paragraph - no special style
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    
    # Visible caption
    r = p.add_run(full_caption)
    r.bold = True
    r.font.size = Pt(PT_SMALL)
    r.font.name = FONT
    r.font.color.rgb = C_BLACK
    
    # Hidden TC field for List of Tables
    tc_text = f"Tabla {table_num}\t{clean_caption}" if table_num > 0 else f"Tabla\t{clean_caption}"
    _add_tc_field(p, tc_text, "T")


def add_list_of_tables(doc: Document, title: str = "LIST OF TABLES"):
    r"""
    Insert a List of Tables field using TOA (Table of Authorities) approach.
    This creates a proper List of Tables that Word can update.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    
    # Title
    p = doc.add_paragraph()
    run = _run(p, title, bold=True, size=13)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # Create paragraph for the LOT field
    lot_para = doc.add_paragraph()
    lot_para.paragraph_format.space_before = Pt(12)
    lot_para.paragraph_format.space_after = Pt(12)
    
    # Field: TOA \h \c "Table"  (Table of Authorities for Tables)
    r = lot_para.add_run()
    
    # fldChar begin
    e1 = OxmlElement('w:fldChar')
    e1.set(qn('w:fldCharType'), 'begin')
    r._r.append(e1)
    
    # instrText - TOA (Table of Authorities) with \c category
    e2 = OxmlElement('w:instrText')
    e2.set(qn('xml:space'), 'preserve')
    e2.text = ' TOA \\h \\c "1" '
    r._r.append(e2)
    
    # fldChar separate
    e3 = OxmlElement('w:fldChar')
    e3.set(qn('w:fldCharType'), 'separate')
    r._r.append(e3)
    
    # Placeholder text
    e4 = OxmlElement('w:r')
    e4_pr = OxmlElement('w:rPr')
    e4_color = OxmlElement('w:color')
    e4_color.set(qn('w:val'), '808080')
    e4_pr.append(e4_color)
    e4_i = OxmlElement('w:i')
    e4_pr.append(e4_i)
    e4.append(e4_pr)
    e4_t = OxmlElement('w:t')
    e4_t.text = '[Press Ctrl+A, F9 to update List of Tables]'
    e4.append(e4_t)
    lot_para._p.append(e4)
    
    # fldChar end
    e5 = OxmlElement('w:fldChar')
    e5.set(qn('w:fldCharType'), 'end')
    r._r.append(e5)

    # Hint paragraph
    doc.add_paragraph()


# ─────────────────────────────────────────────────────────────────────────────
# Text helpers
# ─────────────────────────────────────────────────────────────────────────────

def _run(para, text: str, bold=False, italic=False,
         size=None, color=None, underline=False):
    if not text:
        return None
    r = para.add_run(text)
    r.font.name   = FONT
    r.font.size   = Pt(size or PT_NORM)
    r.bold        = bold
    r.italic      = italic
    r.underline   = underline
    if color:
        r.font.color.rgb = color
    return r


# ─────────────────────────────────────────────────────────────────────────────
# LaTeX text cleaning
# ─────────────────────────────────────────────────────────────────────────────

# Populated during preamble parsing
MACROS: dict[str, str] = {}


def _resolve_macros(text: str) -> str:
    """Replace custom commands with their resolved text.
    Sort by key length descending so \\PrestadorII replaces before \\PrestadorI.
    """
    for k in sorted(MACROS.keys(), key=len, reverse=True):
        text = text.replace(k, MACROS[k])
    return text


def _clean(text: str) -> str:
    """
    Convert LaTeX special chars/sequences to Unicode.
    Does NOT strip formatting commands — use strip_fmt() for plain text.
    """
    text = _resolve_macros(text)
    # Note: font-size switches (\large, \normalsize, etc.) and font switches
    # (\bfseries, \normalfont) are now handled by parse_inline so they can
    # affect formatting. They are no longer stripped here.
    # Remove \centering to avoid \c being treated as cedilla during accent processing.
    text = re.sub(r"\\centering\b", " ", text)
    # Convert LaTeX accent commands to Unicode (\'{a}, \'a, \~{n}, \^u, etc.)
    # MUST happen before replacing standalone ~ with non-breaking space
    def _accent_repl(m):
        accent_cmd = m.group(1)
        letter = m.group(2)
        combining = {
            "'": "\u0301", "`": "\u0300", "^": "\u0302",
            '"': "\u0308", "~": "\u0303", "c": "\u0327",
            "v": "\u030c",  # caron
        }.get(accent_cmd, "")
        return letter + combining if combining else m.group(0)

    text = re.sub(r"\\(['`^\"~cv])\{?([a-zA-Z])\}?", _accent_repl, text)
    text = unicodedata.normalize("NFC", text)
    # Now safe to replace standalone ~ (non-breaking space), ---, --, etc.
    text = text.replace("~",   "\u00a0")          # non-breaking space
    text = text.replace("---", "\u2014")           # em dash
    text = text.replace("--",  "\u2013")           # en dash
    text = text.replace("\\%", "%")
    text = text.replace("\\$", "$")
    text = text.replace("\\&", "&")
    text = text.replace("\\_", "_")
    text = text.replace("\\#", "#")
    text = text.replace("\\,", "\u202f")           # narrow no-break space
    text = re.sub(r"\\hspace\*?\{[^}]*\}", "\t", text)  # replace \hspace{...} with tab
    text = re.sub(r"\\vspace\*?\{[^}]*\}", "\t", text)  # replace \vspace{...} with tab
    # Collapse whitespace surrounding any tab into a single tab
    text = re.sub(r"[ \t]*\t[ \t]*", "\t", text)
    text = re.sub(r"\\[,;!: ]\s*", " ", text)
    text = re.sub(r"\\ ", " ", text)
    return text


def strip_fmt(text: str) -> str:
    """Return plain text: resolve macros + remove all LaTeX markup."""
    text = _clean(text)
    # Strip inline math $...$ — replace \cdot etc. with unicode, then strip $
    text = re.sub(r"\$\\cdot\$",             "·",   text)
    text = re.sub(r"\$\\approx\$",           "≈",   text)
    text = re.sub(r"\$\\times\$",            "×",   text)
    text = re.sub(r"\$([^$]*)\$",            r"\1", text)   # remaining $math$
    text = re.sub(r"\\textbf\{([^}]*)\}",    r"\1", text)
    text = re.sub(r"\\textit\{([^}]*)\}",    r"\1", text)
    text = re.sub(r"\\emph\{([^}]*)\}",      r"\1", text)
    text = re.sub(r"\\textsc\{([^}]*)\}",    r"\1", text)  # \textsc{...}
    text = re.sub(r"\\small\b",              "",    text)
    text = re.sub(r"\\large\b",              "",    text)
    text = re.sub(r"\\Large\b",              "",    text)
    text = re.sub(r"\\normalsize\b",         "",    text)
    text = re.sub(r"\\footnotesize\b",       "",    text)
    text = re.sub(r"\\scriptsize\b",         "",    text)
    text = re.sub(r"\\tiny\b",               "",    text)
    text = re.sub(r"\\huge\b",               "",    text)
    text = re.sub(r"\\Huge\b",               "",    text)
    text = re.sub(r"\\centering\b",          "",    text)
    text = re.sub(r"\\textbullet\b",         "•",   text)
    text = re.sub(r"\\href\{[^}]*\}\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\url\{([^}]*)\}",       r"\1", text)
    text = re.sub(r"\\ref\{[^}]*\}",         "",    text)
    text = re.sub(r"\\label\{[^}]*\}",       "",    text)
    text = re.sub(r"\\newline\b",            " ",   text)
    text = re.sub(r"\\hspace\*?\{[^}]*\}",  " ",   text)  # replace \hspace{...} with space
    text = re.sub(r"\\vspace\*?\{[^}]*\}",  " ",   text)  # replace \vspace{...} with space
    text = re.sub(r"\\rule\{[^}]*\}\{[^}]*\}", "", text)  # remove \rule{...}{...}
    # Line breaks with optional argument: \\[\bigskipamount] → space
    text = re.sub(r"\\\\(?:\[[^\]]*\])?", " ", text)
    text = re.sub(r"\\includegraphics(?:\[[^\]]*\])?\{[^}]*\}", "", text)
    text = re.sub(r"\\[a-zA-Z]+\*?\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\[a-zA-Z]+\*?",        "",    text)
    text = text.replace("{", "").replace("}", "")
    # Normalize internal whitespace (multi-line LaTeX → single line)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ─────────────────────────────────────────────────────────────────────────────
# LaTeX math → OMML (native Word equations) via Pandoc
# ─────────────────────────────────────────────────────────────────────────────

MATH_TEMP_COUNTER = [0]

def latex_to_omml_element(latex_code: str, temp_dir: Path, pandoc_path: str):
    r"""
    Convierte código LaTeX matemático a un elemento OMML usando Pandoc.
    `latex_code` debe incluir los delimitadores ($...$, \[...\], o entorno completo).
    Retorna el elemento XML (m:oMath o m:oMathPara) o None.
    Los resultados se cachean para evitar re-conversiones.
    """
    cache_key = latex_code
    if cache_key in OMML_CACHE:
        cached = OMML_CACHE[cache_key]
        if cached is not None:
            # lxml deepcopy is unreliable; serialize + parse for a true independent copy
            return parse_xml(etree.tostring(cached, encoding='utf-8'))
        return None

    MATH_TEMP_COUNTER[0] += 1
    n = MATH_TEMP_COUNTER[0]
    tex_file = temp_dir / f"math_{n:04d}.tex"
    out_file = temp_dir / f"math_{n:04d}.docx"

    # Expand custom macros so Pandoc can parse math that relies on them
    resolved_code = _resolve_macros(latex_code)
    # Pandoc mis-parses the European decimal brace syntax {,} in math mode,
    # splitting the number (e.g. $25{,}8^{\circ}$ becomes 25 , 8°).
    # Replace it with a dot for Pandoc and restore commas in the OMML output.
    resolved_code = resolved_code.replace("{,}", ".")
    # Pandoc also produces empty-base superscripts for unbraced exponents
    # like $m^3$; normalise to $m^{3}$ so the base is preserved.
    resolved_code = re.sub(r"(?<!\{)\^([A-Za-z0-9])(?![A-Za-z0-9])", r"^{\1}", resolved_code)
    tex_content = (
        r"\documentclass{article}" + "\n"
        r"\begin{document}" + "\n"
        + resolved_code + "\n"
        r"\end{document}" + "\n"
    )
    tex_file.write_text(tex_content, encoding="utf-8")

    try:
        subprocess.run(
            [pandoc_path, "-f", "latex", "-t", "docx", str(tex_file), "-o", str(out_file)],
            capture_output=True, text=True, timeout=30
        )
    except Exception:
        return None

    if not out_file.exists():
        return None

    try:
        pandoc_doc = Document(str(out_file))
        for para in pandoc_doc.paragraphs:
            for child in para._p:
                tag = child.tag
                if tag == f"{{{MATH_NS}}}oMath" or tag == f"{{{MATH_NS}}}oMathPara":
                    result = deepcopy(child)
                    # Restore European decimal commas and use the standard
                    # degree sign (U+00B0) instead of the ring operator.
                    for t_el in result.iter(f"{{{MATH_NS}}}t"):
                        if t_el.text:
                            txt = t_el.text
                            txt = txt.replace("\u2218", "\u00b0")
                            txt = re.sub(r"(\d)\.(\d)", r"\1,\2", txt)
                            t_el.text = txt
                    OMML_CACHE[cache_key] = result
                    return result
    except Exception:
        pass

    return None


# ─────────────────────────────────────────────────────────────────────────────
# Inline LaTeX → Word runs
# ─────────────────────────────────────────────────────────────────────────────

# Token patterns (order matters)
_TOK = re.compile(
    r"(\\textbf\{(?:[^{}]|\{[^{}]*\})*\})"           # \textbf{...}
    r"|(\\textit\{(?:[^{}]|\{[^{}]*\})*\})"           # \textit{...}
    r"|(\\emph\{(?:[^{}]|\{[^{}]*\})*\})"             # \emph{...}
    r"|(\\href\{[^}]*\}\{(?:[^{}]|\{[^{}]*\})*\})"   # \href{url}{text}
    r"|(\\url\{[^}]*\})"                               # \url{...}
    r"|(\$[^$]+\$)"                                    # $inline math$
    r"|(\\(?:label|ref|pageref|eqref)\{[^}]*\})"      # \label{...}, \ref{...} — skip
    r"|(\\(?:bfseries|normalfont|itshape|large|Large|LARGE|small|normalsize|footnotesize|scriptsize|tiny|huge|Huge)\b)"  # font switches
    r"|(\\[a-zA-Z]+\*?(?:\{[^}]*\})*)"                # generic \cmd{arg}...
    r"|([^\\${}]+)"                                    # plain text
)


def parse_inline(para, text: str, base_sz=PT_NORM, bold=False, italic=False, size=None):
    r"""Parse inline LaTeX and add styled runs to `para`.

    Supports formatting switches such as \bfseries, \normalfont, \itshape,
    \large, \normalsize, etc. by tracking current bold/italic/size state.
    """
    text = _resolve_macros(text)
    # Protect inline math ($...$) from _clean, which otherwise corrupts
    # commands like \circ, \cdot, etc. before Pandoc can process them.
    math_segments = []
    def _protect_math(m):
        math_segments.append(m.group(0))
        return f"\x00MATH{len(math_segments) - 1}\x00"
    text = re.sub(r"\$[^$]+\$", _protect_math, text)
    text = _clean(text)
    def _restore_math(m):
        return math_segments[int(m.group(1))]
    text = re.sub(r"\x00MATH(\d+)\x00", _restore_math, text)

    cur_bold = bool(bold)
    cur_italic = bool(italic)
    cur_size = size if size is not None else base_sz

    # LaTeX size switches mapped to approximate Word point sizes
    size_map = {
        "tiny": 6,
        "scriptsize": 8,
        "footnotesize": 9,
        "small": 10,
        "normalsize": PT_NORM,
        "large": 14,
        "Large": 16,
        "LARGE": 20,
        "huge": 24,
        "Huge": 28,
    }

    for m in _TOK.finditer(text):
        g0 = m.group(0)
        if m.group(1):   # \textbf
            inner_m = re.search(r"\\textbf\{((?:[^{}]|\{[^{}]*\})*)\}", g0)
            if inner_m:
                n_runs = len(para.runs)
                parse_inline(para, inner_m.group(1), base_sz=base_sz,
                             bold=cur_bold, italic=cur_italic, size=cur_size)
                for run in para.runs[n_runs:]:
                    run.bold = True
        elif m.group(2): # \textit
            inner_m = re.search(r"\\textit\{((?:[^{}]|\{[^{}]*\})*)\}", g0)
            if inner_m:
                n_runs = len(para.runs)
                parse_inline(para, inner_m.group(1), base_sz=base_sz,
                             bold=cur_bold, italic=cur_italic, size=cur_size)
                for run in para.runs[n_runs:]:
                    run.italic = True
        elif m.group(3): # \emph
            inner_m = re.search(r"\\emph\{((?:[^{}]|\{[^{}]*\})*)\}", g0)
            if inner_m:
                n_runs = len(para.runs)
                parse_inline(para, inner_m.group(1), base_sz=base_sz,
                             bold=cur_bold, italic=cur_italic, size=cur_size)
                for run in para.runs[n_runs:]:
                    run.italic = True
        elif m.group(4): # \href{url}{text}
            link_txt = re.sub(r"\\href\{[^}]*\}\{((?:[^{}]|\{[^{}]*\})*)\}", r"\1", g0)
            _run(para, _clean(strip_fmt(link_txt)), color=C_LINK, underline=True,
                 size=cur_size, bold=cur_bold, italic=cur_italic)
        elif m.group(5): # \url{...}
            url = re.sub(r"\\url\{([^}]*)\}", r"\1", g0)
            _run(para, url, color=C_LINK, underline=True,
                 size=cur_size, bold=cur_bold, italic=cur_italic)
        elif m.group(6): # $math$
            math_content = g0[1:-1]   # strip $
            # Plain numeric values (e.g. $-0,92$) are better as normal text
            # so they match the surrounding table/text formatting.
            if re.fullmatch(r"\s*[-+]?\d+(?:[.,]\d+)?\s*", math_content):
                plain_num = math_content.strip().replace(".", ",")
                _run(para, plain_num, size=cur_size, bold=cur_bold, italic=cur_italic)
            else:
                omml = None
                if TEMP_DIR and PANDOC_PATH:
                    try:
                        omml = latex_to_omml_element(g0, TEMP_DIR, PANDOC_PATH)
                    except Exception:
                        pass
                if omml is not None:
                    para._p.append(omml)
                else:
                    _run(para, _clean(strip_fmt(math_content)),
                         italic=True, size=cur_size, bold=cur_bold)
        elif m.group(7): # \label, \ref, \pageref — skip
            pass
        elif m.group(8): # font switches: \bfseries, \normalfont, \itshape, \large, etc.
            cmd = g0.lstrip("\\").strip()
            if cmd == "bfseries":
                cur_bold = True
            elif cmd == "normalfont":
                cur_bold = False
                cur_italic = False
            elif cmd == "itshape":
                cur_italic = True
            elif cmd in size_map:
                cur_size = size_map[cmd]
            # Other unrecognized switches become no-ops
        elif m.group(9): # generic \cmd
            plain = strip_fmt(g0)
            if plain:
                _run(para, _clean(plain), size=cur_size, bold=cur_bold, italic=cur_italic)
        else:            # plain text
            cleaned = _clean(g0)
            if cleaned:
                _run(para, cleaned, size=cur_size, bold=cur_bold, italic=cur_italic)


# ─────────────────────────────────────────────────────────────────────────────
# Preamble parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_preamble(src: str) -> dict:
    r"""
    Extract \newcommand definitions and \setcounter values.
    Populates global MACROS dict.
    Returns counters dict.
    """
    counters = {}

    # \setcounter{name}{value}
    for m in re.finditer(r"\\setcounter\{(\w+)\}\{(\d+)\}", src):
        counters[m.group(1)] = int(m.group(2))

    plazo = counters.get("PlazoTotal", 60)

    # \newcommand{\Name}{Definition}  — single-level braces
    for m in re.finditer(
        r"\\newcommand\{(\\[A-Za-z]+)\}\{((?:[^{}]|\{[^{}]*\})*)\}",
        src
    ):
        cmd, dfn = m.group(1), m.group(2)
        MACROS[cmd + "{}"] = dfn
        MACROS[cmd]        = dfn

    # \newglossaryentry{key}{name={name},description={...}}  -> macros \gls{key} and \Gls{key}
    pos = 0
    gls_keys = set()
    while True:
        idx = src.find("\\newglossaryentry{", pos)
        if idx == -1:
            break
        key_start = idx + len("\\newglossaryentry{") - 1  # points to '{' before key
        key = _extract_balanced_braces(src, key_start)
        if key is None:
            pos = key_start
            continue
        block_start = key_start + len(key) + 2
        if block_start >= len(src) or src[block_start] != "{":
            pos = block_start
            continue
        block = _extract_balanced_braces(src, block_start)
        if block is None:
            pos = block_start
            continue
        name_idx = block.find("name=")
        if name_idx != -1:
            name_start = name_idx + len("name=")
            if name_start < len(block) and block[name_start] == "{":
                name = _extract_balanced_braces(block, name_start)
                if name is not None:
                    MACROS[f"\\gls{{{key}}}"] = name
                    gls_keys.add(f"\\gls{{{key}}}")
                    if name:
                        cap_name = name[0].upper() + name[1:]
                        MACROS[f"\\Gls{{{key}}}"] = cap_name
                        gls_keys.add(f"\\Gls{{{key}}}")
        pos = block_start + len(block) + 2

    # Month / Year / generic placeholders
    date_info = {
        "\\MONTH":  "",
        "\\YEAR":   "",
        "\\month":  "",
        "\\year":   "",
    }
    for k, v in date_info.items():
        MACROS[k]       = v
        MACROS[k + "{}"] = v

    # Milestone day commands (arithmetic on PlazoTotal)
    milestones = {
        "\\DiaI":     str((8  * plazo + 50) // 100),
        "\\DiaII":    str((25 * plazo + 50) // 100),
        "\\DiaIII":   str((50 * plazo + 50) // 100),
        "\\DiaIV":    str((75 * plazo + 50) // 100),
        "\\DiaFinal": str(plazo),
        "\\arabic{PlazoTotal}": str(plazo),
    }
    for k, v in milestones.items():
        MACROS[k]       = v
        MACROS[k + "{}"] = v

    # Recursively resolve nested macro refs (one pass)
    for key in list(MACROS.keys()):
        for k2, v2 in list(MACROS.items()):
            MACROS[key] = MACROS[key].replace(k2, v2)

    # Strip remaining LaTeX from macro values (preserve glossary names so $math$ stays intact)
    for key in list(MACROS.keys()):
        if key not in gls_keys:
            MACROS[key] = strip_fmt(MACROS[key])

    return counters


# ─────────────────────────────────────────────────────────────────────────────
# Document setup
# ─────────────────────────────────────────────────────────────────────────────

def setup_document(doc: Document):
    """Page size (A4), margins, and base styles."""
    sec = doc.sections[0]
    sec.page_width    = Inches(8.27)
    sec.page_height   = Inches(11.69)
    sec.top_margin    = Cm(2.5)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin   = Cm(2.54)
    sec.right_margin  = Cm(2.54)

    doc.styles["Normal"].font.name = FONT
    doc.styles["Normal"].font.size = Pt(PT_NORM)
    doc.styles["Normal"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    heading_cfg = {
        "Heading 1": (14, True,  True),
        "Heading 2": (13, True,  False),
        "Heading 3": (12, True,  False),
        "Heading 4": (11, False, True),
        "Heading 5": (11, False, True),
    }
    for name, (sz, bold, is_caps) in heading_cfg.items():
        try:
            s = doc.styles[name]
            s.font.name  = FONT
            s.font.size  = Pt(sz)
            s.font.bold  = bold
            s.font.color.rgb = C_BLACK
            if is_caps:
                s.font.all_caps = True
        except Exception:
            pass


def add_header(doc: Document, logo_path: Path, right_text: str):
    """Left: logo image. Right: project name. With horizontal line below."""
    sec    = doc.sections[0]
    header = sec.header
    header.is_linked_to_previous = False
    
    # Clear any existing paragraphs
    for p in list(header.paragraphs):
        p._element.getparent().remove(p._element)
    
    # Create table with 2 columns spanning full width
    table = header.add_table(rows=1, cols=2, width=Inches(6.5))
    table.autofit = False
    
    # Left cell: Logo
    left_cell = table.cell(0, 0)
    left_cell.width = Inches(1.5)
    left_para = left_cell.paragraphs[0]
    left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    left_para.paragraph_format.space_after = Pt(0.2)  # Small space after
    
    if logo_path and logo_path.exists():
        run_img = left_para.add_run()
        run_img.add_picture(str(logo_path), height=Cm(0.8))
    
    # Right cell: Text aligned to right
    right_cell = table.cell(0, 1)
    right_cell.width = Inches(5.0)
    right_para = right_cell.paragraphs[0]
    right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right_para.paragraph_format.space_after = Pt(0.2)  # Small space after
    _run(right_para, right_text, size=PT_SMALL)
    
    # Add bottom border to last row's cells
    for cell in table.rows[0].cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        # Only bottom border
        for edge in ['top', 'left', 'right']:
            b = OxmlElement(f'w:{edge}')
            b.set(qn('w:val'), 'nil')
            tcBorders.append(b)
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:color'), '808080')
        tcBorders.append(bottom)
        tcPr.append(tcBorders)


def add_footer(doc: Document):
    """Centred page number."""
    sec    = doc.sections[0]
    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.clear()
    run = fp.add_run()
    _add_page_number_field(run)


# ─────────────────────────────────────────────────────────────────────────────
# Title page
# ─────────────────────────────────────────────────────────────────────────────

def add_title_page(doc: Document, title_img_path: Path | None,
                   title_img_width: float | None, title_img_height: float | None,
                   title: str, author: str, date_str: str):
    # Optional title image (e.g. \includegraphics inside \title)
    if title_img_path and title_img_path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        kwargs = {}
        if title_img_width is not None:
            kwargs["width"] = Inches(title_img_width)
        if title_img_height is not None:
            kwargs["height"] = Inches(title_img_height)
        if not kwargs:
            kwargs["width"] = Inches(4)
        p.add_run().add_picture(str(title_img_path), **kwargs)
        doc.add_paragraph()

    # Main title text
    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, title, bold=True, size=16)
        doc.add_paragraph()

    # Author
    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, author, bold=True, size=PT_NORM)
        doc.add_paragraph()

    # Date
    if date_str:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, date_str, bold=True, size=PT_NORM)


# ─────────────────────────────────────────────────────────────────────────────
# Environment extractor
# ─────────────────────────────────────────────────────────────────────────────

def _extract_balanced_braces(text: str, start: int):
    """Extract content inside balanced braces starting at `start` (which must be '{')."""
    if start >= len(text) or text[start] != "{":
        return None
    depth = 1
    i = start + 1
    while i < len(text) and depth > 0:
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
        i += 1
    if depth == 0:
        return text[start + 1 : i - 1]
    return None


def _extract_caption_content(text: str) -> str | None:
    """Extract the first \\caption{...} content, handling nested braces."""
    idx = text.find("\\caption{")
    if idx == -1:
        return None
    brace_start = idx + len("\\caption{") - 1  # position of '{'
    content = _extract_balanced_braces(text, brace_start)
    return content


def _remove_captions(text: str) -> str:
    """Remove all \\caption{...} occurrences, handling nested braces."""
    result = []
    pos = 0
    while True:
        idx = text.find("\\caption{", pos)
        if idx == -1:
            result.append(text[pos:])
            break
        result.append(text[pos:idx])
        brace_start = idx + len("\\caption{") - 1
        content = _extract_balanced_braces(text, brace_start)
        if content is None:
            result.append(text[idx:idx + len("\\caption")])
            pos = idx + len("\\caption")
            continue
        after = brace_start + len(content) + 2
        # skip trailing \\ if present
        if text[after:after + 2] == "\\\\":
            after += 2
        pos = after
    return "".join(result)


def extract_env(text: str, env: str, start: int = 0):
    r"""
    Find first balanced \begin{env}...\end{env} starting at `start`.
    Returns (abs_begin, abs_end, inner) or None.
    """
    bp = re.compile(r"\\begin\{" + re.escape(env) + r"\*?\}")
    ep = re.compile(r"\\end\{"   + re.escape(env) + r"\*?\}")

    m0 = bp.search(text, start)
    if not m0:
        return None

    depth = 1
    pos   = m0.end()
    while depth > 0 and pos < len(text):
        nb = bp.search(text, pos)
        ne = ep.search(text, pos)
        if ne is None:
            break
        if nb and nb.start() < ne.start():
            depth += 1
            pos = nb.end()
        else:
            depth -= 1
            if depth == 0:
                return (m0.start(), ne.end(), text[m0.end(): ne.start()])
            pos = ne.end()
    return None


def _strip_column_spec(text: str):
    """
    Remove the first balanced {col_spec} block from text.
    Returns (col_spec_content, rest_of_text).
    Handles nested braces like p{0.06\\textwidth}.
    """
    text = text.lstrip()
    if not text.startswith("{"):
        return "", text
    depth = 0
    for i, c in enumerate(text):
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return text[1:i], text[i + 1:]
    return "", text


def _parse_col_widths_dxa(col_spec: str, n_cols: int,
                           content_dxa: int = 9026) -> list:
    """
    Parse a LaTeX column spec and return column widths in DXA units.
    Handles p{frac\\textwidth}, X, l, r, c.
    Unknown/flexible columns share the remaining width equally.
    """
    fixed_fracs = []   # (index, fraction)
    flex_indices = []

    # Extract individual column descriptors (ignore | and @{...})
    col_spec_clean = re.sub(r"@\{[^}]*\}", "", col_spec)
    col_spec_clean = col_spec_clean.replace("|", "")

    tokens = re.findall(
        r"p\{([\d.]+)\\textwidth\}"   # p{0.XX\textwidth}
        r"|m\{([\d.]+)\\textwidth\}"  # m{0.XX\textwidth}
        r"|b\{([\d.]+)\\textwidth\}"  # b{0.XX\textwidth}
        r"|X"                          # X (tabularx flexible)
        r"|[lrcL]",                    # l r c or L
        col_spec_clean
    )

    cols = []
    for tok in tokens:
        if isinstance(tok, tuple):
            frac = tok[0] or tok[1] or tok[2]
            if frac:
                cols.append(float(frac))
            else:
                cols.append(None)  # X / l / r / c
        else:
            cols.append(None)

    # If we parsed wrong number of cols, fall back to equal widths
    if not cols or len(cols) != n_cols:
        equal = content_dxa // n_cols if n_cols else content_dxa
        return [equal] * n_cols

    fixed_total = sum(f for f in cols if f is not None)
    flex_count  = sum(1 for f in cols if f is None)
    remaining   = content_dxa - int(fixed_total * content_dxa)
    flex_each   = (remaining // flex_count) if flex_count else 0

    result = []
    for f in cols:
        if f is not None:
            result.append(int(f * content_dxa))
        else:
            result.append(max(flex_each, 500))  # minimum 500 DXA

    # Adjust rounding to match content width exactly
    diff = content_dxa - sum(result)
    if result:
        result[-1] += diff
    return result



def _is_rule(line: str) -> bool:
    return bool(re.match(r"\\(toprule|midrule|bottomrule|hline)\b", line.strip()))


def _is_category_row(line: str) -> bool:
    r"""Detect rows that are category/section headers (\multicolumn spanning all cols)."""
    return bool(re.match(r"\s*\\multicolumn\b", line.strip()) and "&" not in line)


def _expand_multicolumn(line: str) -> str:
    r"""
    Replace \multicolumn{n}{align}{text} with (text)(& repeat n-1 times).
    """
    def repl(m):
        n    = int(m.group(1))
        text = m.group(3)
        return text + " & " * (n - 1)

    return re.sub(
        r"\\multicolumn\{(\d+)\}\{([^}]*)\}\{((?:[^{}]|\{[^{}]*\})*)\}",
        repl, line
    )


# Contador global de tablas para usar tablas Pandoc pre-generadas
TABLE_COUNTER = [0]

# Contador global de figuras tikz
TIKZ_COUNTER = [0]

# Contador global de figuras (para captions numerados)
FIGURE_COUNTER = [0]

# Preámbulo LaTeX original (se establece en parse_body)
PREAMBLE_GLOBAL: str = ""

# Rutas de búsqueda de imágenes extraídas de \graphicspath{{...}}
GRAPHICSPATHS: list[Path] = []

# Opciones por defecto de enumitem extraídas del preámbulo (se establece en parse_body)
# Formato: {nivel: label_format, None: label_format_global}
ENUMITEM_DEFAULTS: dict[int | None, str] = {}

def apply_booktabs_style(table):
    """
    Aplica estilo LaTeX booktabs a una tabla:
    - Sin bordes verticales
    - Líneas horizontales: top (gruesa), mid (media entre filas), bottom (gruesa)
    """
    rows = list(table.rows)
    if not rows:
        return
    
    num_rows = len(rows)
    
    for row_idx, row in enumerate(rows):
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            
            # Sin bordes verticales (left/right)
            for edge in ['left', 'right']:
                b = OxmlElement(f'w:{edge}')
                b.set(qn('w:val'), 'nil')
                tcBorders.append(b)
            
            # Top border: primera fila = toprule (grueso)
            if row_idx == 0:
                b = OxmlElement('w:top')
                b.set(qn('w:val'), 'single')
                b.set(qn('w:sz'), '12')  # Grueso
                b.set(qn('w:color'), '000000')
                tcBorders.append(b)
            else:
                # Otras filas = sin línea arriba (la línea viene de la celda de arriba)
                b = OxmlElement('w:top')
                b.set(qn('w:val'), 'nil')
                tcBorders.append(b)
            
            # Bottom border:
            if row_idx == num_rows - 1:
                # Última fila = bottomrule (grueso)
                b = OxmlElement('w:bottom')
                b.set(qn('w:val'), 'single')
                b.set(qn('w:sz'), '12')
                b.set(qn('w:color'), '000000')
                tcBorders.append(b)
            elif row_idx == 0:
                # Después del header = midrule (media)
                b = OxmlElement('w:bottom')
                b.set(qn('w:val'), 'single')
                b.set(qn('w:sz'), '6')
                b.set(qn('w:color'), '000000')
                tcBorders.append(b)
            else:
                # Entre filas = línea ligera (cmidrule style)
                b = OxmlElement('w:bottom')
                b.set(qn('w:val'), 'single')
                b.set(qn('w:sz'), '4')  # Más delgada
                b.set(qn('w:color'), 'C0C0C0')  # Gris claro
                tcBorders.append(b)
            
            tcPr.append(tcBorders)
            
            # Asegurar que el texto tenga fuente Calibri y color negro
            for para in cell.paragraphs:
                para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.font.name = FONT
                    run.font.color.rgb = C_BLACK
                    if run.font.size is None:
                        run.font.size = Pt(PT_SMALL)

def normalize_text(text: str) -> str:
    """Normaliza texto para comparación (quita espacios extras, lowercase)."""
    return ' '.join(text.strip().lower().split())




# ─────────────────────────────────────────────────────────────────────────────
# enumitem label helpers
# ─────────────────────────────────────────────────────────────────────────────

def _int_to_roman(num: int, upper: bool = True) -> str:
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']
    res = ''
    for v, s in zip(val, syms):
        while num >= v:
            res += s
            num -= v
    return res if upper else res.lower()


def _int_to_alpha(num: int, upper: bool = True) -> str:
    res = ''
    while num > 0:
        num -= 1
        res = chr(ord('A') + (num % 26)) + res
        num //= 26
    return res if upper else res.lower()


def _format_enumitem_label(label_format: str, index: int) -> str:
    if not label_format:
        return str(index)
    fmt = label_format
    fmt = re.sub(r'\\roman\*', lambda m: _int_to_roman(index, upper=False), fmt)
    fmt = re.sub(r'\\Roman\*', lambda m: _int_to_roman(index, upper=True), fmt)
    fmt = re.sub(r'\\alph\*',  lambda m: _int_to_alpha(index, upper=False), fmt)
    fmt = re.sub(r'\\Alph\*',  lambda m: _int_to_alpha(index, upper=True), fmt)
    fmt = re.sub(r'\\arabic\*', lambda m: str(index), fmt)
    return fmt


def _extract_optional_arg(text: str):
    text = text.lstrip()
    if not text.startswith('['):
        return None, text
    depth = 0
    i = 0
    while i < len(text):
        c = text[i]
        if c == '[':
            depth += 1
        elif c == ']':
            depth -= 1
            if depth == 0:
                arg = text[1:i]
                rest = text[i + 1:].lstrip()
                return arg, rest
        elif c == "\\" and i + 1 < len(text):
            i += 1
        i += 1
    return None, text


def _parse_enumitem_label(optional_arg: str) -> str:
    if not optional_arg:
        return None
    m = re.search(r'(?:^|[,\[]\s*)label\s*=\s*(.+?)(?:\s*,\s*\w+\s*=|$)', optional_arg, re.DOTALL)
    if m:
        return m.group(1).strip()
    return None


def _extract_enumitem_defaults(preamble: str) -> dict[int | None, str]:
    r"""Extrae los labels por defecto de \setlist[enumerate]{...} del pre�mbulo."""
    defaults: dict[int | None, str] = {}
    # Busca \setlist[enumerate,...]{...} o \setlist[enumerate]{...}
    for m in re.finditer(r'\\setlist\s*\[\s*enumerate\s*(?:,\s*(\d+))?\s*\]\s*\{', preamble):
        level = int(m.group(1)) if m.group(1) else None
        content = _extract_balanced_braces(preamble, m.end() - 1)
        if content is None:
            continue
        label = _parse_enumitem_label(content)
        if label:
            defaults[level] = label
    return defaults


def _get_enumitem_default(level: int) -> str | None:
    """Devuelve el label por defecto para el nivel dado, o el global si no hay específico."""
    if level in ENUMITEM_DEFAULTS:
        return ENUMITEM_DEFAULTS[level]
    return ENUMITEM_DEFAULTS.get(None)


def is_duplicate_row(row_texts: list, seen_rows: list) -> bool:
    """
    Detecta si una fila es idéntica a una ya vista.
    Compara el contenido normalizado de todas las celdas.
    """
    normalized = tuple(normalize_text(t) for t in row_texts)
    
    for seen in seen_rows:
        if normalized == seen:
            return True
    return False


def is_continuation_row_cells(cells_texts: list) -> bool:
    """
    Detecta si una fila es de 'continued' o 'continues on next page'.
    Estas filas vienen de longtable y no queremos mostrarlas.
    Recibe una lista de textos de celdas.
    """
    for text in cells_texts:
        text = text.strip().lower()
        # Patrones de filas de continuación
        if any(pattern in text for pattern in [
            'continued',
            'continues on next page',
            'table continued',
            'continúa en la siguiente página',
            '(continued)'
        ]):
            return True
    return False


def is_continuation_row(row) -> bool:
    """Versión para filas de tabla de docx."""
    texts = [cell.text for cell in row.cells]
    return is_continuation_row_cells(texts)


def compile_tikz_to_png(tikz_code: str, preamble: str, temp_dir: Path, fig_num: int):
    """
    Compila un bloque tikzpicture a PNG mediante pdflatex + pdftoppm.
    Recorta el espacio en blanco sobrante con Pillow.
    Devuelve la ruta del PNG o None si falla.
    """
    # Limpiar preámbulo: quitar \documentclass y \usepackage{geometry} conflictivo
    clean_preamble = re.sub(r"\\documentclass\[.*?\]\{.*?\}\s*", "", preamble, flags=re.DOTALL)
    clean_preamble = re.sub(r"\\usepackage\[margin=[^\]]*\]\{geometry\}\s*", "", clean_preamble)

    tex_content = (
        r"\documentclass{article}" + "\n"
        r"\usepackage[margin=0pt]{geometry}" + "\n"
        r"\pagestyle{empty}" + "\n"
        + clean_preamble + "\n"
        r"\begin{document}" + "\n"
        r"\noindent" + "\n"
        + tikz_code + "\n"
        r"\end{document}" + "\n"
    )

    tex_file = temp_dir / f"tikz_{fig_num:02d}.tex"
    tex_file.write_text(tex_content, encoding="utf-8")

    # Buscar pdflatex
    pdflatex = "pdflatex"
    if sys.platform == "win32":
        for p in [
            r"C:\Program Files\MiKTeX\miktex\bin\x64\pdflatex.exe",
            r"C:\ProgramData\MiKTeX\miktex\bin\x64\pdflatex.exe",
            r"C:\Users\Alienware\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdflatex.exe",
        ]:
            if Path(p).exists():
                pdflatex = p
                break

    try:
        result = subprocess.run(
            [pdflatex, "-interaction=nonstopmode", str(tex_file.name)],
            cwd=str(temp_dir),
            capture_output=True,
            text=True,
            timeout=90,
        )
    except Exception as e:
        print(f"  [WARN] pdflatex falló para tikz {fig_num}: {e}")
        return None

    pdf_file = temp_dir / f"tikz_{fig_num:02d}.pdf"
    if not pdf_file.exists() or pdf_file.stat().st_size == 0:
        print(f"  [WARN] PDF no generado para tikz {fig_num}")
        return None

    # Buscar pdftoppm
    pdftoppm = "pdftoppm"
    if sys.platform == "win32":
        for p in [
            r"C:\Program Files\MiKTeX\miktex\bin\x64\pdftoppm.exe",
            r"C:\ProgramData\MiKTeX\miktex\bin\x64\pdftoppm.exe",
            r"C:\Users\Alienware\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdftoppm.exe",
        ]:
            if Path(p).exists():
                pdftoppm = p
                break

    png_file = temp_dir / f"tikz_{fig_num:02d}.png"
    try:
        subprocess.run(
            [pdftoppm, "-png", "-r", "300", "-singlefile", str(pdf_file), str(temp_dir / f"tikz_{fig_num:02d}")],
            capture_output=True,
            check=False,
            timeout=30,
        )
    except Exception as e:
        print(f"  [WARN] pdftoppm falló para tikz {fig_num}: {e}")
        return None

    if not png_file.exists():
        return None

    # Recortar espacio en blanco con Pillow
    try:
        from PIL import Image
        img = Image.open(str(png_file))
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGB")
        gray = img.convert("L")
        # Invertir para que el blanco sea 0 y lo demás >0
        bbox = gray.point(lambda x: 255 - x).getbbox()
        if bbox:
            img = img.crop(bbox)
            img.save(str(png_file))
    except Exception as e:
        print(f"  [WARN] No se pudo recortar tikz {fig_num}: {e}")

    return png_file


def _parse_graphics_size(opts_str: str, img_path: Path):
    r"""
    Parse \\includegraphics[...] options and return (width_inches, height_inches).
    Supports: width=0.65\\textwidth, width=\\textwidth, width=5cm, height=0.28\\textheight,
              height=\\textheight, height=3cm, scale=0.5, keepaspectratio.
    Reads natural image size for aspect ratio if needed.
    """
    width = None
    height = None
    scale = None
    keep_aspect = "keepaspectratio" in opts_str.lower().replace("-", "").replace(" ", "")

    # width = fraction \textwidth (fraction optional, defaults to 1.0)
    m = re.search(r"width=([\d.]*)\\?textwidth", opts_str)
    if m:
        frac = float(m.group(1)) if m.group(1) else 1.0
        width = frac * TEXT_WIDTH_INCHES
    else:
        m = re.search(r"width=([\d.]*)\\?linewidth", opts_str)
        if m:
            frac = float(m.group(1)) if m.group(1) else 1.0
            width = frac * TEXT_WIDTH_INCHES
        else:
            m = re.search(r"width=([\d.]+)cm", opts_str)
            if m:
                width = float(m.group(1)) / 2.54
            else:
                m = re.search(r"width=([\d.]+)mm", opts_str)
                if m:
                    width = float(m.group(1)) / 25.4
                else:
                    m = re.search(r"width=([\d.]+)in", opts_str)
                    if m:
                        width = float(m.group(1))

    # height = fraction \textheight (fraction optional, defaults to 1.0)
    m = re.search(r"height=([\d.]*)\\?textheight", opts_str)
    if m:
        frac = float(m.group(1)) if m.group(1) else 1.0
        height = frac * TEXT_HEIGHT_INCHES
    else:
        m = re.search(r"height=([\d.]+)cm", opts_str)
        if m:
            height = float(m.group(1)) / 2.54
        else:
            m = re.search(r"height=([\d.]+)mm", opts_str)
            if m:
                height = float(m.group(1)) / 25.4
            else:
                m = re.search(r"height=([\d.]+)in", opts_str)
                if m:
                    height = float(m.group(1))

    # scale
    m = re.search(r"scale=([\d.]+)", opts_str)
    if m:
        scale = float(m.group(1))

    # Natural size for aspect ratio / scale
    try:
        from PIL import Image
        img = Image.open(str(img_path))
        px_w, px_h = img.size
        nat_w = px_w / 96.0
        nat_h = px_h / 96.0
    except Exception:
        nat_w = 4.0
        nat_h = 3.0

    if scale is not None:
        width = nat_w * scale
        height = nat_h * scale
    elif width is not None and height is None:
        ratio = width / nat_w
        height = nat_h * ratio
    elif height is not None and width is None:
        ratio = height / nat_h
        width = nat_w * ratio
    elif width is not None and height is not None:
        # Both width and height were specified
        if keep_aspect:
            # Fit inside the width x height box while preserving aspect ratio
            scale_w = width / nat_w
            scale_h = height / nat_h
            scale = min(scale_w, scale_h)
            width = nat_w * scale
            height = nat_h * scale
        # Without keepaspectratio LaTeX distorts the image, so we keep
        # the explicit width and height as-is.

    return width, height


def _extract_graphicspaths(source: str) -> list[Path]:
    r"""Extrae rutas de \graphicspath{{dir1/}{dir2/}...} del preámbulo."""
    paths = []
    for m in re.finditer(r"\\graphicspath\s*\{((?:\s*\{[^}]*\}\s*)*)\}", source):
        block = m.group(1)
        for pm in re.finditer(r"\{([^}]*)\}", block):
            p = pm.group(1).strip().replace("\\", "/").strip("/")
            if p:
                paths.append(Path(p))
    return paths


def resolve_image_path(img_name: str, base_dir: Path) -> Path | None:
    """Busca una imagen respetando GRAPHICSPATHS y extensiones comunes."""
    name = img_name.strip().replace("\\", "/")
    # Lista de candidatos: nombre exacto + con extensiones comunes
    candidates = [name]
    if "." not in name.split("/")[-1]:
        for ext in (".png", ".jpg", ".jpeg", ".pdf", ".gif", ".bmp", ".eps"):
            candidates.append(name + ext)
    # Buscar en base_dir y en cada graphicspath relativo a base_dir
    search_dirs = [base_dir] + [base_dir / p for p in GRAPHICSPATHS]
    for cand in candidates:
        for d in search_dirs:
            p = d / cand
            if p.exists():
                return p
    return None


def _insert_image_with_size(para, img_path: Path, max_width_inches: float = 6.3, dpi: float = 300.0):
    """Inserta imagen respetando su tamaño natural (basado en píxeles / dpi).
    Si excede el ancho máximo, escala proporcionalmente."""
    try:
        from PIL import Image
        img = Image.open(str(img_path))
        px_w, px_h = img.size
        width = px_w / dpi
        height = px_h / dpi
        if width > max_width_inches:
            ratio = max_width_inches / width
            width = max_width_inches
            height = height * ratio
        para.add_run().add_picture(str(img_path), width=Inches(width), height=Inches(height))
        return True
    except Exception:
        return False


def _render_figure_as_table(doc: Document, inner: str, base_dir: Path) -> bool:
    r"""
    Renderiza una figura con múltiples \subfigure en una tabla de Word,
    respetando filas detectadas por \\ (salto de línea LaTeX) o \hfill.
    Todas las imágenes de una misma fila comparten la misma altura,
    y la fila completa se escala proporcionalmente si excede el ancho de página.
    Devuelve True si se renderizó como tabla, False si no había subfigures.
    """
    # Patrón: \subfigure[caption]{\includegraphics[opts]{name}}
    subfig_pattern = (
        r"\\subfigure\s*\[((?:[^\[\]]|\\\[|\\)*)\]\s*\{+\s*"
        r"\s*\\includegraphics(?:\[([^\]]*)\])?\{([^}]*)\}"
        r"(?:\s*\\label\{[^}]*\})?\s*\}+"
    )
    all_subfigs = list(re.finditer(subfig_pattern, inner))
    
    if len(all_subfigs) <= 1:
        return False  # Usar renderizado normal para 0 o 1 subfigure
    
    # Dividir inner en filas usando \\ como separador (aproximado),
    # permitiendo argumento opcional como \\[6pt] entre filas
    row_texts = re.split(r"\\\\(?:\s*\[[^\]]*\])?\s*(?=\\subfigure)", inner)
    
    row_groups = []
    for rt in row_texts:
        sf_in_row = list(re.finditer(subfig_pattern, rt))
        if sf_in_row:
            row_groups.append(sf_in_row)
    
    if not row_groups:
        return False
    
    max_cols = max(len(g) for g in row_groups)
    if max_cols == 0:
        return False
    
    # ── Primera pasada: calcular tamaños deseados de cada imagen ─────────────
    DEFAULT_HEIGHT = 2.5  # pulgadas, si ninguna imagen especifica tamaño
    
    def _get_natural_size(img_path):
        try:
            from PIL import Image
            img = Image.open(str(img_path))
            px_w, px_h = img.size
            return px_w / 96.0, px_h / 96.0
        except Exception:
            return 4.0, 3.0
    
    row_data = []  # lista de filas; cada fila es lista de dicts
    for subfigs in row_groups:
        row = []
        for sf in subfigs:
            opts = sf.group(2) or ""
            img_name = sf.group(3)
            img_path = resolve_image_path(img_name, base_dir)
            nat_w, nat_h = _get_natural_size(img_path)
            
            w, h = _parse_graphics_size(opts, img_path)
            
            # Si no se especificó nada, usar altura default
            if w is None and h is None:
                h = DEFAULT_HEIGHT
                w = nat_w * (h / nat_h)
            elif w is not None and h is None:
                h = nat_h * (w / nat_w)
            elif h is not None and w is None:
                w = nat_w * (h / nat_h)
            
            row.append({
                "sf": sf,
                "img_path": img_path,
                "img_name": img_name,
                "nat_w": nat_w,
                "nat_h": nat_h,
                "desired_w": w,
                "desired_h": h,
            })
        row_data.append(row)
    
    # ── Segunda pasada: calcular altura común y escalar filas si es necesario ─
    final_rows = []
    for row in row_data:
        # Altura común = mínimo de alturas deseadas (respeta la más restrictiva)
        common_h = min(item["desired_h"] for item in row)
        
        # Recalcular anchos a la altura común
        scaled_row = []
        total_w = 0.0
        for item in row:
            new_w = item["nat_w"] * (common_h / item["nat_h"])
            scaled_row.append({**item, "final_w": new_w, "final_h": common_h})
            total_w += new_w
        
        # Si la fila completa excede el ancho de texto, escalar proporcionalmente
        if total_w > TEXT_WIDTH_INCHES:
            scale = TEXT_WIDTH_INCHES / total_w
            for item in scaled_row:
                item["final_w"] *= scale
                item["final_h"] *= scale
        
        final_rows.append(scaled_row)
    
    # ── Tercera pasada: crear la tabla e insertar imágenes con tamaños finales ─
    table = doc.add_table(rows=len(row_groups), cols=max_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False
    
    subfig_idx = 0
    
    for ri, row in enumerate(final_rows):
        n_cols = len(row)
        for ci, item in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = ""
            cell.width = Inches(item["final_w"])
            
            caption_text = strip_fmt(item["sf"].group(1))
            img_path = item["img_path"]
            img_name = item["img_name"]
            
            # Párrafo para la imagen
            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if img_path and img_path.exists():
                try:
                    p_img.add_run().add_picture(
                        str(img_path),
                        width=Inches(item["final_w"]),
                        height=Inches(item["final_h"]),
                    )
                except Exception:
                    _run(p_img, f"[Figure: {img_name}]", italic=True, size=PT_SMALL)
            else:
                _run(p_img, f"[Figure: {img_name}]", italic=True, size=PT_SMALL)
            
            # Caption del subfigure
            if caption_text:
                p_cap = cell.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                letter = chr(ord('a') + subfig_idx)
                _run(p_cap, f"({letter}) {caption_text}", italic=True, size=PT_SMALL)
            
            subfig_idx += 1
        
        # Celdas vacías sobrantes
        for ci in range(n_cols, max_cols):
            cell = table.cell(ri, ci)
            cell.text = ""
    
    # Quitar bordes a toda la tabla
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for edge in ('top', 'bottom', 'left', 'right'):
                b = OxmlElement(f'w:{edge}')
                b.set(qn('w:val'), 'nil')
                tcBorders.append(b)
            tcPr.append(tcBorders)
    
    return True


def render_table_with_pandoc(doc: Document, tab_inner: str, caption: str = ""):
    """
    Inserta tabla pre-generada por Pandoc y aplica estilo booktabs.
    Las tablas deben estar en TEMP_DIR: tabla_XX_pandoc.docx
    Ajusta automaticamente el tamaño de fuente segun el numero de columnas.
    """
    global TEMP_DIR
    TABLE_COUNTER[0] += 1
    table_num = TABLE_COUNTER[0]
    
    pandoc_table_file = TEMP_DIR / f'tabla_{table_num:02d}_pandoc.docx'
    
    if not pandoc_table_file.exists():
        print(f'  [WARN] No se encuentra tabla Pandoc: {pandoc_table_file}')
        return False
    
    try:
        # Cargar documento con tabla Pandoc
        pandoc_doc = Document(str(pandoc_table_file))
        
        if not pandoc_doc.tables:
            print(f'  [WARN] El archivo no tiene tablas: {pandoc_table_file}')
            return False
        
        # Agregar caption si existe
        if caption:
            add_table_caption(doc, caption, table_num)
        
        # Copiar primera tabla del archivo Pandoc
        source_table = pandoc_doc.tables[0]
        
        # Detectar numero de columnas y ajustar fuente automaticamente
        ncols = len(source_table.columns)
        font_size = PT_SMALL  # Default 9pt
        if ncols > 7:
            font_size = 7  # Muy ancha -> fuente pequena
        elif ncols > 5:
            font_size = 8  # Moderadamente ancha
        
        # Filtrar filas de continuación y duplicados (headers repetidos)
        valid_rows = []
        seen_headers = set()  # Para detectar headers duplicados
        
        for row in source_table.rows:
            row_texts = [cell.text.strip() for cell in row.cells]
            row_text = ' '.join(row_texts).lower()
            
            # Detectar filas de continuación
            if 'continued' in row_text or 'continues on next page' in row_text:
                continue
            
            # Detectar headers duplicados (fila completamente idéntica a una vista antes)
            row_hash = tuple(normalize_text(t) for t in row_texts)
            if row_hash in seen_headers:
                continue  # Skip fila duplicada
            seen_headers.add(row_hash)
            
            valid_rows.append(row)
        
        if not valid_rows:
            print(f'  [WARN] Tabla {table_num} solo tiene filas de continuación')
            return False
        
        # Crear nueva tabla solo con filas válidas
        ncols = len(source_table.columns)
        new_table = doc.add_table(rows=len(valid_rows), cols=ncols)
        new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Detectar numero de columnas y ajustar fuente automaticamente
        font_size = PT_SMALL  # Default 9pt
        if ncols > 7:
            font_size = 7  # Muy ancha -> fuente pequena
        elif ncols > 5:
            font_size = 8  # Moderadamente ancha
        
        # Copiar contenido celda por celda PRESERVANDO formato
        for i, source_row in enumerate(valid_rows):
            new_row = new_table.rows[i]
            
            # Detectar si esta fila tiene celdas combinadas (mismo texto en todas)
            row_texts = [cell.text.strip() for cell in source_row.cells]
            all_same = len(set(row_texts)) == 1 and row_texts[0] and ncols > 1
            
            if all_same:
                # Merge todas las celdas de esta fila
                merged_cell = new_row.cells[0]
                merged_cell.text = row_texts[0]
                for run in merged_cell.paragraphs[0].runs:
                    run.font.name = FONT
                    run.font.size = Pt(font_size)
                    run.bold = True
                # Merge con la ultima celda
                merged_cell.merge(new_row.cells[-1])
            else:
                # Copiar normalmente
                for j, source_cell in enumerate(source_row.cells):
                    if j < len(new_row.cells):
                        new_cell = new_row.cells[j]
                        new_cell.text = ""
                        
                        for source_para in source_cell.paragraphs:
                            has_math = any(child.tag.startswith(f"{{{MATH_NS}}}") for child in source_para._p)
                            if not source_para.text.strip() and not has_math:
                                continue
                            new_para = new_cell.paragraphs[0] if new_cell.paragraphs else new_cell.add_paragraph()
                            new_para.text = source_para.text
                            new_para.alignment = source_para.alignment
                            # Copy OMML math elements (Pandoc produces these inline)
                            if has_math:
                                for child in source_para._p:
                                    if child.tag.startswith(f"{{{MATH_NS}}}"):
                                        cloned = etree.fromstring(etree.tostring(child))
                                        new_para._p.append(cloned)
                            for source_run in source_para.runs:
                                for run in new_para.runs:
                                    run.font.name = FONT
                                    run.font.size = Pt(font_size)
                                    if source_run.bold:
                                        run.bold = True
                                    if source_run.italic:
                                        run.italic = True
        
        # Aplicar estilo booktabs (solo líneas horizontales)
        apply_booktabs_style(new_table)
        
        doc.add_paragraph()  # espacio después
        print(f'  [OK] Tabla {table_num} insertada ({len(valid_rows)} de {len(source_table.rows)} filas)')
        return True
        
    except Exception as e:
        print(f'  [ERROR] Insertando tabla {table_num}: {e}')
        import traceback
        traceback.print_exc()
        return False


def _is_p_only_table(col_spec: str) -> bool:
    """Return True if the table has a single p/m/b column (no l/c/r/X).

    In single-column p{} tables, \\ inside the cell is a line break rather
    than a row separator, so Pandoc's row splitting produces too many rows.
    """
    if not col_spec:
        return False
    spec_clean = re.sub(r">\s*\{[^}]*\}", "", col_spec)
    spec_clean = re.sub(r"@[^{]*", "", spec_clean)
    spec_clean = re.sub(r"\|+", "", spec_clean)
    p_cols = re.findall(r"[pmb]\{[^}]*\}", spec_clean)
    other_cols = re.findall(r"[lcrX]", spec_clean)
    return len(p_cols) == 1 and not other_cols


def render_table(doc: Document, tab_inner: str, caption: str = "",
                 ncols_hint: int = 0, col_widths_dxa: list = None,
                 col_spec: str = ""):
    """
    Convert LaTeX tabular inner content to a Word table.
    First tries Pandoc, falls back to manual rendering.
    Tables with inline math ($...$) skip Pandoc because Pandoc places
    OMML on separate paragraphs; fallback manual rendering keeps math inline.
    Vertical p-only tables are also rendered manually because Pandoc treats
    in-cell \\ as row separators.
    """
    p_only = _is_p_only_table(col_spec)
    has_inline_math = bool(re.search(r"\$[^$]+\$", tab_inner))
    if not p_only and not has_inline_math and render_table_with_pandoc(doc, tab_inner, caption):
        return
    if p_only or has_inline_math:
        # Consume the Pandoc counter so subsequent tables stay in sync
        TABLE_COUNTER[0] += 1

    # Fallback: manual rendering
    if caption:
        add_table_caption(doc, caption, TABLE_COUNTER[0])

    # Clean tab_inner
    cleaned_inner = re.sub(r"p\{[^}]*\}", "", tab_inner)
    cleaned_inner = re.sub(r"m\{[^}]*\}", "", cleaned_inner)
    cleaned_inner = re.sub(r"b\{[^}]*\}", "", cleaned_inner)

    # Detect "vertical" tables where every column is p/m/b (no l/c/r/X).
    # In such tables \\ inside a cell is a line break, not a row separator.
    p_only = _is_p_only_table(col_spec)

    if p_only:
        # Split rows by horizontal rules; keep \\ as in-cell line breaks
        raw_rows = re.split(r"\\(?:toprule|midrule|bottomrule|hline)\b", cleaned_inner)
    else:
        raw_rows = re.split(r"\\\\", cleaned_inner)

    rows_data = []
    is_header_row = []
    
    for raw in raw_rows:
        raw = raw.strip()
        # Remove optional spacing arguments from \\[...] that ended up at row start
        raw = re.sub(r"^\[[^\]]*\]\s*", "", raw).strip()
        if not raw or raw.startswith("%"):
            continue
        # Strip table rules that may be embedded at the start of a row (e.g. \midrule\nBloque)
        raw = re.sub(r"\\(toprule|midrule|bottomrule|hline)\b\s*", "", raw).strip()
        raw = re.sub(r"\\addlinespace\b\s*", "", raw).strip()
        if not raw:
            continue
        if _is_rule(raw):
            continue
        # Skip rows that are only layout commands (e.g. \centering before table rows)
        if re.match(r"^\s*\\(?:centering|noindent|par\b|medskip|bigskip|smallskip)\s*$", raw):
            continue
            
        is_hdr = bool(re.match(r"\s*\\rowcolor\{blue!20\}", raw))
        raw = re.sub(r"\\rowcolor\{[^}]*\}", "", raw).strip()
        if not raw:
            continue

        raw = _expand_multicolumn(raw)
        cells = [c.strip() for c in raw.split("&")]
        cells = [c for c in cells if not re.match(r"^p\{", c.strip())]
        # Remove only trailing empty cells; keep leading/middle empty cells
        # because they are needed for \multirow alignment
        while cells and not cells[-1].strip():
            cells.pop()
        
        # Filtrar filas de continuación (longtable)
        if cells and is_continuation_row_cells(cells):
            continue
        
        # Filtrar duplicados exactos
        if cells and is_duplicate_row(cells, [tuple(r) for r in rows_data]):
            continue
        
        if cells:
            rows_data.append(cells)
            is_header_row.append(is_hdr)

    if not rows_data:
        return

    ncols = ncols_hint or max(len(r) for r in rows_data)
    
    for i, row in enumerate(rows_data):
        if len(row) < ncols:
            row.extend([""] * (ncols - len(row)))
        elif len(row) > ncols:
            rows_data[i] = row[:ncols]

    if not col_widths_dxa or len(col_widths_dxa) != ncols:
        col_widths_dxa = _parse_col_widths_dxa("", ncols)

    table = doc.add_table(rows=len(rows_data), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    header_row_idx = 0

    for ri, cells in enumerate(rows_data):
        tr = table.rows[ri]
        for ci in range(ncols):
            cell = tr.cells[ci]
            cell_text = cells[ci] if ci < len(cells) else ""
            cell_text = re.sub(r"\\label\{[^}]*\}", "", cell_text)
            cell_text = re.sub(r"\\hspace\*?\{[^}]*\}", "", cell_text)

            # For vertical p-only tables, treat \\ inside a cell as line breaks
            if p_only:
                cell_lines = [ln.strip() for ln in re.split(r"\\\\(?:\[[^\]]*\])?", cell_text) if ln.strip()]
            else:
                cell_lines = [cell_text]

            for line_idx, line in enumerate(cell_lines):
                if line_idx == 0:
                    p = cell.paragraphs[0]
                else:
                    p = cell.add_paragraph()
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if (ri == header_row_idx) else WD_ALIGN_PARAGRAPH.LEFT

                if ri == header_row_idx:
                    if '$' in line:
                        parse_inline(p, line, base_sz=PT_SMALL)
                        for run in p.runs:
                            run.bold = True
                    else:
                        _run(p, strip_fmt(line), bold=True, size=PT_SMALL)
                else:
                    parse_inline(p, line, base_sz=PT_SMALL)

    apply_booktabs_style(table)
    doc.add_paragraph()


# ─────────────────────────────────────────────────────────────────────────────
# List renderer
# ─────────────────────────────────────────────────────────────────────────────

def _get_list_number_abstract_num_id(doc: Document):
    """Find the abstractNumId used by the 'List Number' style in this document."""
    try:
        # 1. Try to find an existing paragraph with List Number and read its numId
        list_number_num_id = None
        for para in doc.paragraphs:
            if para.style and para.style.name == "List Number":
                pPr = para._p.find(qn('w:pPr'))
                if pPr is not None:
                    numPr = pPr.find(qn('w:numPr'))
                    if numPr is not None:
                        numId_el = numPr.find(qn('w:numId'))
                        if numId_el is not None:
                            list_number_num_id = int(numId_el.get(qn('w:val')))
                            break

        # 2. If no paragraph exists yet, create a temp one to get the numId
        if list_number_num_id is None:
            temp = doc.add_paragraph(style="List Number")
            pPr = temp._p.find(qn('w:pPr'))
            if pPr is not None:
                numPr = pPr.find(qn('w:numPr'))
                if numPr is not None:
                    numId_el = numPr.find(qn('w:numId'))
                    if numId_el is not None:
                        list_number_num_id = int(numId_el.get(qn('w:val')))
            # Remove temp paragraph
            temp._element.getparent().remove(temp._element)

        # 3. Look up the abstractNumId for this numId in numbering.xml
        if list_number_num_id is not None:
            numbering = doc.part.numbering_part.numbering_definitions._numbering
            for num in numbering.findall(qn('w:num')):
                if int(num.get(qn('w:numId'))) == list_number_num_id:
                    abstractNumId_el = num.find(qn('w:abstractNumId'))
                    if abstractNumId_el is not None:
                        return int(abstractNumId_el.get(qn('w:val')))
        return None
    except Exception:
        return None


def _split_items(inner: str) -> list[str]:
    r"""Split list inner content by \item only at the current nesting level.

    This avoids treating \item commands inside nested itemize/enumerate
    environments as top-level item separators.
    """
    items = []
    current = []
    depth = 0
    i = 0
    n = len(inner)
    while i < n:
        m_begin = re.match(r"\\begin\{(itemize|enumerate)\*?\}", inner[i:])
        if m_begin:
            depth += 1
            current.append(m_begin.group(0))
            i += m_begin.end()
            continue
        m_end = re.match(r"\\end\{(itemize|enumerate)\*?\}", inner[i:])
        if m_end:
            depth = max(0, depth - 1)
            current.append(m_end.group(0))
            i += m_end.end()
            continue
        if depth == 0 and inner.startswith(r"\item", i):
            items.append("".join(current))
            current = []
            i += len(r"\item")
            continue
        current.append(inner[i])
        i += 1
    if current:
        items.append("".join(current))
    # Drop leading empty chunk before the first \item
    if items and not items[0].strip():
        items = items[1:]
    return items


def render_list(doc: Document, inner: str, ordered: bool = False, depth: int = 0, base_dir: Path = Path("."), label_format: str = None):
    r"""
    Recursively parse itemize / enumerate and add list paragraphs.
    Handles nested lists, custom \item[label] labels,
    and block environments (figure, table, center, etc.) inside items.
    """
    bullet_style  = "List Bullet"  if not ordered else "List Number"
    # Try to use indented styles for nesting
    if depth > 0:
        bullet_style = "List Bullet 2" if not ordered else "List Number 2"

    # Split on \item boundaries, respecting nested itemize/enumerate
    items_raw = _split_items(inner)
    
    item_index = 0  # Counter for auto-numbering when no custom label

    # Environments that can appear inside an \item and need special handling
    NESTED_ENV_RE = re.compile(r"\\begin\{(itemize|enumerate|figure|table\*?|longtable|center|tikzpicture)\b")

    for item_raw in items_raw:
        item_raw = item_raw.strip()
        # Normalize internal newlines/spaces from multi-line LaTeX source
        item_raw = re.sub(r"\s+", " ", item_raw)
        if not item_raw:
            continue

        # Custom label:  [label] text
        label_m = re.match(r"^\[([^\]]*)\]\s*", item_raw)
        custom_label = None
        has_custom_label = False
        is_enumitem_label = False
        if label_m:
            custom_label = _clean(label_m.group(1))
            item_raw     = item_raw[label_m.end():]
            has_custom_label = True
        else:
            item_index += 1  # Only increment for auto-numbered items
            if ordered and label_format:
                custom_label = _format_enumitem_label(label_format, item_index)
                has_custom_label = True
                is_enumitem_label = True

        # When custom label exists, use normal paragraph (not list style) to avoid double numbering
        # This also makes text use 'Normal' style (justified, more width)
        if has_custom_label:
            para_style = "Normal"  # Use Normal style for custom labels (justified, more width)
        else:
            para_style = bullet_style

        # Helper to render a paragraph of item text
        def _render_item_text(text: str, use_label: bool, continuation: bool = False):
            if not text and not use_label:
                return
            if ordered and not has_custom_label and not use_label and not continuation:
                # Manual numbering ensures each enumerate restarts at 1
                p = doc.add_paragraph(style="Normal")
                p.paragraph_format.left_indent = Inches(0.5 * (depth + 1))
                p.paragraph_format.first_line_indent = Inches(-0.25)
                run = p.add_run(f"{item_index}.  ")
                run.font.name = FONT
                run.font.size = Pt(PT_NORM)
                run.font.color.rgb = C_BLACK
                parse_inline(p, text)
                _maybe_add_tab_stop(p, text)
            else:
                if continuation:
                    # Continuation paragraph: normal style, same indentation as list items, no number
                    p = doc.add_paragraph(style="Normal")
                    p.paragraph_format.left_indent = Inches(0.5 * (depth + 1))
                else:
                    p = doc.add_paragraph(style=para_style)
                    if has_custom_label:
                        if is_enumitem_label:
                            p.paragraph_format.left_indent = Inches(0.5 * (depth + 1))
                            p.paragraph_format.first_line_indent = Inches(-0.35)
                        else:
                            p.paragraph_format.left_indent = Inches(0.3 * (depth + 1))
                if use_label and custom_label:
                    n_runs_before = len(p.runs)
                    parse_inline(p, custom_label)
                    for run in p.runs[n_runs_before:]:
                        # enumitem default labels should match LaTeX formatting (not bold)
                        if not is_enumitem_label:
                            run.bold = True
                        run.font.name = FONT
                    _run(p, "  ", size=PT_NORM)
                parse_inline(p, text)
                _maybe_add_tab_stop(p, text)

        # Process the item extracting nested environments in order
        remaining = item_raw
        label_used = False
        is_continuation = False  # True for text that follows a nested environment

        while remaining:
            env_match = NESTED_ENV_RE.search(remaining)
            if not env_match:
                _render_item_text(remaining, use_label=not label_used, continuation=is_continuation)
                break

            pre_text = remaining[:env_match.start()].strip()
            env_name = env_match.group(1)

            result = extract_env(remaining, env_name, env_match.start())
            if not result:
                _render_item_text(remaining, use_label=not label_used, continuation=is_continuation)
                break

            _, end_pos, env_inner = result
            env_text = remaining[env_match.start():end_pos]
            post_text = remaining[end_pos:].strip()

            # Render text before the environment
            _render_item_text(pre_text, use_label=not label_used, continuation=is_continuation)
            label_used = True

            # Render the environment itself
            if env_name in ("itemize", "enumerate"):
                opt_arg, env_inner_clean = _extract_optional_arg(env_inner)
                nested_label = _parse_enumitem_label(opt_arg) if opt_arg else None
                # Apply enumitem default label for nested enumerate if no explicit option
                if env_name == "enumerate" and nested_label is None:
                    nested_label = _get_enumitem_default(depth + 2)
                render_list(doc, env_inner_clean, ordered=(env_name == "enumerate"), depth=depth + 1, base_dir=base_dir, label_format=nested_label)
            else:
                _walk(doc, env_text, base_dir)

            # Continue scanning the remainder of the item for more nested environments
            remaining = post_text
            is_continuation = True

        # If item was empty but had a label, make sure at least the label is rendered
        if not remaining and not label_used and custom_label:
            _render_item_text("", use_label=True)
def parse_body(doc: Document, source: str, base_dir: Path):
    """Walk the entire source and render into `doc`."""
    global PREAMBLE_GLOBAL
    
    # Reiniciar contadores
    TABLE_COUNTER[0] = 0
    TIKZ_COUNTER[0] = 0
    FIGURE_COUNTER[0] = 0

    # Guardar preámbulo para compilar tikzpictures
    preamble_end = source.find("\\begin{document}")
    PREAMBLE_GLOBAL = source[:preamble_end] if preamble_end != -1 else ""

    # Extraer rutas de imágenes del preámbulo
    global GRAPHICSPATHS
    GRAPHICSPATHS = _extract_graphicspaths(PREAMBLE_GLOBAL)

    # Extraer opciones por defecto de enumitem del preámbulo
    global ENUMITEM_DEFAULTS
    ENUMITEM_DEFAULTS = _extract_enumitem_defaults(PREAMBLE_GLOBAL)

    # ── remove comments ──────────────────────────────────────────────────────
    # Elimina comentarios (% hasta fin de línea) pero preserva \% (porcentaje literal)
    source = re.sub(r"(?m)(?<!\\)%.*$", "", source)

    # ── title block ──────────────────────────────────────────────────────────
    def _extract_one(cmd, text):
        # Find command \cmd not followed by another letter
        m = re.search(r"\\" + cmd + r"(?![a-zA-Z])", text)
        if not m:
            return ""
        brace_idx = text.find("{", m.end())
        if brace_idx == -1:
            return ""
        content = _extract_balanced_braces(text, brace_idx)
        return content if content is not None else ""

    title_raw  = _extract_one("title",  source)
    author_raw = _extract_one("author", source)
    date_raw   = _extract_one("date",   source)

    # Extract optional title image (\includegraphics inside \title)
    title_img_path = None
    title_img_width = None
    title_img_height = None
    img_m = re.search(r"\\includegraphics(?:\[([^\]]*)\])?\{([^}]*)\}", title_raw)
    if img_m:
        opts = img_m.group(1) or ""
        img_name = img_m.group(2)
        title_img_path = resolve_image_path(img_name, base_dir)
        if title_img_path and title_img_path.exists():
            title_img_width, title_img_height = _parse_graphics_size(opts, title_img_path)
        # Remove the \includegraphics from title_raw for text extraction
        title_raw = title_raw[:img_m.start()] + title_raw[img_m.end():]

    # Use the full cleaned title text (no hardcoded defaults)
    title_text = strip_fmt(title_raw).strip()
    # If title is empty after cleaning, try \textbf as fallback
    if not title_text:
        tb_m = re.search(r"\\textbf\{([^}]+)\}", title_raw)
        if tb_m:
            title_text = strip_fmt(tb_m.group(1))

    author_text = strip_fmt(author_raw)
    date_text   = strip_fmt(date_raw)

    has_title_content = bool(title_text or author_text or date_text or (title_img_path and title_img_path.exists()))
    if has_title_content:
        add_title_page(doc, title_img_path, title_img_width, title_img_height, title_text, author_text, date_text)
        _safe_page_break(doc)

    # ── Table of Contents ────────────────────────────────────────────────────
    add_toc(doc)
    _safe_page_break(doc)
    
    # ── isolate the document body ─────────────────────────────────────────────
    begin_doc = re.search(r"\\begin\{document\}", source)
    end_doc   = re.search(r"\\end\{document\}",   source)
    if not begin_doc:
        return
    body = source[begin_doc.end(): end_doc.start() if end_doc else len(source)]

    # Remove \title, \author, \date blocks from body so _walk doesn't re-render them
    def _remove_cmd_block(cmd: str, text: str) -> str:
        m = re.search(r"\\" + cmd + r"(?![a-zA-Z])", text)
        if not m:
            return text
        brace_idx = text.find("{", m.end())
        if brace_idx == -1:
            return text
        content = _extract_balanced_braces(text, brace_idx)
        if content is None:
            return text
        end_idx = brace_idx + 1 + len(content) + 1
        start_idx = m.start()
        return text[:start_idx] + text[end_idx:]

    for cmd in ("title", "author", "date"):
        body = _remove_cmd_block(cmd, body)

    # ── resolve \\ref{...} to numbers ────────────────────────────────────────
    labels = pre_scan_labels(body)
    body = resolve_refs(body, labels)

    # ── strip fancyhdr commands that leak into body text ──────────────────────
    body = strip_fancyhdr_cmds(body)

    # ── scan bibliography ─────────────────────────────────────────────────────
    bib_map, bib_inner, body, bib_title = pre_scan_bibliography(body)

    # ── resolve \\cite{...} to numbers ───────────────────────────────────────
    body = resolve_cites(body, bib_map)

    # ── scan figures and tables for lists ────────────────────────────────────
    figures = pre_scan_figures(body)
    tables = pre_scan_tables(body)

    # ── walk the body ─────────────────────────────────────────────────────────
    _walk(doc, body, base_dir, figures, tables)

    # ── render bibliography ───────────────────────────────────────────────────
    if bib_inner is not None:
        _safe_page_break(doc)
        if bib_title:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(bib_title.upper())
            run.font.name = FONT
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = C_BLACK
        _render_bibliography(doc, bib_inner, bib_map)


def strip_fancyhdr_cmds(text: str) -> str:
    """Remove fancyhdr setup commands that leak into the document body."""
    # Single-line fancyhdr / header-footer commands
    text = re.sub(
        r"(?m)^\s*\\(?:pagestyle|thispagestyle|fancyhf|rhead|lhead|chead|fancyfoot|rfoot|lfoot|cfoot|headrulewidth|footrulewidth)\b.*(?:\n|\r\n?)",
        "", text,
    )
    # \renewcommand{\headrulewidth}{...} or {\footrulewidth}{...}
    text = re.sub(
        r"(?m)^\s*\\renewcommand\s*\*?\s*\{\s*\\(?:headrulewidth|footrulewidth)\s*\}\s*\{[^}]*\}\s*(?:\n|\r\n?)",
        "", text,
    )
    # \setlength\headheight{...}
    text = re.sub(
        r"(?m)^\s*\\setlength\s*\\headheight\s*\{[^}]*\}\s*(?:\n|\r\n?)",
        "", text,
    )
    return text


# Significant structural tokens
_STRUCT = re.compile(
    r"\\(?:chapter|section|subsection|subsubsection|paragraph|begin|end|newpage|clearpage)\*?\b"
    r"|\\appendix\b"
    r"|\\maketitle\b|\\tableofcontents\b|\\listoffigures\b|\\listoftables\b"
)

# Display math
_DISP_MATH = re.compile(r"\\\[(.*?)\\\]", re.DOTALL)


def pre_scan_labels(text: str) -> dict:
    """Scan text to map \\label{...} to figure/table numbers."""
    labels = {}
    fig_counter = 0
    table_counter = 0

    # Figures (including figure*)
    fig_pattern = re.compile(r"\\begin\{figure\*?\}")
    pos = 0
    while True:
        m = fig_pattern.search(text, pos)
        if not m:
            break
        idx = m.start()
        result = extract_env(text, "figure", idx)
        if result:
            _, end_pos, inner = result
            if "\\caption{" in inner:
                fig_counter += 1
                for label_m in re.finditer(r"\\label\{([^}]*)\}", inner):
                    labels[label_m.group(1)] = str(fig_counter)
            pos = end_pos
        else:
            pos = idx + 1

    # Tables
    for env_name in ("table", "table*", "longtable"):
        pos = 0
        search_str = f"\\begin{{{env_name}}}"
        while True:
            idx = text.find(search_str, pos)
            if idx == -1:
                break
            result = extract_env(text, env_name, idx)
            if result:
                _, end_pos, inner = result
                if "\\caption{" in inner:
                    table_counter += 1
                    for label_m in re.finditer(r"\\label\{([^}]*)\}", inner):
                        labels[label_m.group(1)] = str(table_counter)
                pos = end_pos
            else:
                pos = idx + 1

    # Equations (all types, in order of appearance)
    eq_counter = 0
    pos = 0
    eq_envs = ("equation", "equation*", "align", "align*", "gather", "gather*", "multline", "multline*")
    while pos < len(text):
        next_idx = len(text)
        is_env = False
        env_name = None

        for en in eq_envs:
            idx = text.find(f"\\begin{{{en}}}", pos)
            if idx != -1 and idx < next_idx:
                next_idx = idx
                is_env = True
                env_name = en

        # Use _DISP_MATH regex to find display math (ignores \\[ like in \\[\bigskipamount])
        dm_match = _DISP_MATH.search(text, pos)
        idx_dm = dm_match.start() if dm_match else -1
        if idx_dm != -1 and idx_dm < next_idx:
            next_idx = idx_dm
            is_env = False

        if next_idx == len(text):
            break

        if is_env:
            result = extract_env(text, env_name, next_idx)
            if result:
                _, end_pos, inner = result
                # align/gather/etc. may contain multiple equations (one per \\ line);
                # number each labeled line sequentially.
                lines = re.split(r"\\\\(?:\[[^\]]*\])?", inner)
                for line in lines:
                    label_m = re.search(r"\\label\{([^}]*)\}", line)
                    if label_m:
                        eq_counter += 1
                        labels[label_m.group(1)] = str(eq_counter)
                pos = end_pos
            else:
                pos = next_idx + 1
        else:
            dm = _DISP_MATH.match(text, next_idx)
            if not dm:
                break
            block = dm.group(0)
            label_m = re.search(r"\\label\{([^}]*)\}", block)
            if label_m:
                eq_counter += 1
                labels[label_m.group(1)] = str(eq_counter)
            pos = dm.end()

    # Section / chapter labels
    sec_counters = {"chapter": 0, "section": 0, "subsection": 0, "subsubsection": 0, "paragraph": 0}
    has_chapters = bool(re.search(r"\\chapter\b", text))
    appendix_pos = text.find("\\appendix")

    sec_matches = list(re.finditer(
        r"\\(chapter|section|subsection|subsubsection|paragraph)\*?\s*\{[^{}]*\}",
        text
    ))

    for i, m in enumerate(sec_matches):
        cmd = m.group(1)
        start_pos = m.end()
        end_pos = sec_matches[i + 1].start() if i + 1 < len(sec_matches) else len(text)

        if cmd == "chapter":
            sec_counters["chapter"] += 1
            sec_counters["section"] = 0
            sec_counters["subsection"] = 0
            sec_counters["subsubsection"] = 0
            sec_counters["paragraph"] = 0
        elif cmd == "section":
            sec_counters["section"] += 1
            sec_counters["subsection"] = 0
            sec_counters["subsubsection"] = 0
            sec_counters["paragraph"] = 0
        elif cmd == "subsection":
            sec_counters["subsection"] += 1
            sec_counters["subsubsection"] = 0
            sec_counters["paragraph"] = 0
        elif cmd == "subsubsection":
            sec_counters["subsubsection"] += 1
            sec_counters["paragraph"] = 0
        elif cmd == "paragraph":
            sec_counters["paragraph"] += 1

        appendix_mode = appendix_pos != -1 and m.start() > appendix_pos

        if not has_chapters:
            if cmd == "section":
                num_str = f"{sec_counters['section']}"
            elif cmd == "subsection":
                num_str = f"{sec_counters['section']}.{sec_counters['subsection']}"
            elif cmd == "subsubsection":
                num_str = f"{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}"
            elif cmd == "paragraph":
                num_str = f"{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}.{sec_counters['paragraph']}"
            else:
                num_str = ""
        else:
            ch_label = chr(ord('A') + sec_counters['chapter'] - 1) if appendix_mode else str(sec_counters['chapter'])
            if cmd == "chapter":
                num_str = ch_label
            elif cmd == "section":
                num_str = f"{ch_label}.{sec_counters['section']}"
            elif cmd == "subsection":
                num_str = f"{ch_label}.{sec_counters['section']}.{sec_counters['subsection']}"
            elif cmd == "subsubsection":
                num_str = f"{ch_label}.{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}"
            elif cmd == "paragraph":
                num_str = f"{ch_label}.{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}.{sec_counters['paragraph']}"
            else:
                num_str = ""

        # Look for the first \label within a short window after the section command
        search_end = min(end_pos, start_pos + 300)
        label_m = re.search(r"\\label\{([^}]*)\}", text[start_pos:search_end])
        if label_m and label_m.group(1) not in labels:
            labels[label_m.group(1)] = num_str

    return labels


def pre_scan_figures(text: str) -> list[tuple[int, str]]:
    """Return list of (number, plain_caption) for every figure environment."""
    figures = []
    for env_name in ("figure", "figure*"):
        pos = 0
        search_str = f"\\begin{{{env_name}}}"
        while True:
            idx = text.find(search_str, pos)
            if idx == -1:
                break
            result = extract_env(text, env_name, idx)
            if result:
                _, end_pos, inner = result
                # Remove subfigure environments so we only keep the main caption
                inner_no_sub = inner
                sub_pos = 0
                while True:
                    sub_idx = inner_no_sub.find("\\begin{subfigure}", sub_pos)
                    if sub_idx == -1:
                        break
                    sub_res = extract_env(inner_no_sub, "subfigure", sub_idx)
                    if sub_res:
                        _, sub_end, _ = sub_res
                        inner_no_sub = inner_no_sub[:sub_idx] + inner_no_sub[sub_end:]
                    else:
                        sub_pos = sub_idx + 1
                cap = _extract_caption_content(inner_no_sub)
                if cap is not None:
                    cap_clean = strip_fmt(cap)
                    cap_clean = re.sub(r"\\label\{[^}]*\}", "", cap_clean).strip()
                    if cap_clean:
                        figures.append((len(figures) + 1, cap_clean))
                pos = end_pos
            else:
                pos = idx + 1
    return figures


def pre_scan_tables(text: str) -> list[tuple[int, str]]:
    """Return list of (number, plain_caption) for every table environment."""
    tables = []
    for env_name in ("table", "table*", "longtable"):
        pos = 0
        search_str = f"\\begin{{{env_name}}}"
        while True:
            idx = text.find(search_str, pos)
            if idx == -1:
                break
            result = extract_env(text, env_name, idx)
            if result:
                _, end_pos, inner = result
                cap = _extract_caption_content(inner)
                if cap is not None:
                    cap_clean = strip_fmt(cap)
                    cap_clean = re.sub(r"\\label\{[^}]*\}", "", cap_clean).strip()
                    if cap_clean:
                        tables.append((len(tables) + 1, cap_clean))
                pos = end_pos
            else:
                pos = idx + 1
    return tables


def resolve_refs(text: str, labels: dict) -> str:
    """Replace \\ref{...} and \\eqref{...} with the corresponding number from labels."""
    def repl(m):
        label_name = m.group(1)
        return labels.get(label_name, f"[ref:{label_name}]")
    text = re.sub(r"\\eqref\{([^}]*)\}", repl, text)
    text = re.sub(r"\\ref\{([^}]*)\}", repl, text)
    return text


def pre_scan_bibliography(text: str) -> tuple[dict, str | None, str, str]:
    """
    Extract \\begin{thebibliography}...\\end{thebibliography} from text.
    Returns (bib_map, bib_inner, text_without_bibliography, bib_title).
    bib_map: {cite_key: number}
    """
    pattern = r"\\begin\{thebibliography\}\{[^}]*\}(.*?)\\end\{thebibliography\}"
    m = re.search(pattern, text, re.DOTALL)
    if not m:
        return {}, None, text, ""

    bib_inner = m.group(1)
    text_without = text[:m.start()] + text[m.end():]

    bib_map = {}
    counter = 1
    for bm in re.finditer(r"\\bibitem\{([^}]+)\}", bib_inner):
        key = bm.group(1).strip()
        if key and key not in bib_map:
            bib_map[key] = counter
            counter += 1

    # Look for \\renewcommand{\\refname}{...} or \\renewcommand{\\bibname}{...}
    # in the text before thebibliography
    prefix = text[max(0, m.start()-500):m.start()]
    title_m = re.search(r"\\renewcommand\{\\(?:refname|bibname)\}\{((?:[^{}]|\{[^{}]*\})*)\}", prefix)
    bib_title = title_m.group(1) if title_m else "Referencias"

    # Remove the \\renewcommand from text_without so it doesn't render as stray text
    if title_m:
        rc_pattern = r"\\renewcommand\{\\(?:refname|bibname)\}\{" + re.escape(title_m.group(1)) + r"\}"
        text_without = re.sub(rc_pattern, "", text_without)

    return bib_map, bib_inner, text_without, bib_title


def resolve_cites(text: str, bib_map: dict) -> str:
    """Replace \\cite{key1,key2} with [n1,n2] using bib_map."""
    def repl(m):
        keys = [k.strip() for k in m.group(1).split(",")]
        nums = [str(bib_map.get(k, k)) for k in keys if k]
        return "[" + ", ".join(nums) + "]" if nums else ""

    text = re.sub(r"\\cite\{([^}]+)\}", repl, text)
    return text


def _render_bibliography(doc: Document, bib_inner: str, bib_map: dict):
    """
    Render thebibliography inner content as a numbered list of paragraphs.
    Each \\bibitem{key} becomes a paragraph with a hanging indent.
    """
    # Split by \bibitem{...}
    items = re.split(r"\\bibitem\{([^}]+)\}", bib_inner)
    if len(items) < 2:
        return

    i = 1
    while i < len(items):
        key = items[i].strip()
        text = items[i + 1] if i + 1 < len(items) else ""
        num = bib_map.get(key, "")
        if text:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
            _run(p, f"[{num}] ", size=PT_NORM)
            # Strip leading whitespace/newlines
            text = text.strip()
            if text:
                parse_inline(p, text, base_sz=PT_NORM)
        i += 2


def _insert_list_of_figures(doc: Document, figures: list[tuple[int, str]]):
    """Insert a real Word 'List of Figures' TOC field."""
    if not figures:
        return
    _insert_toc_field(
        doc,
        r'TOC \f F \h',
        "ÍNDICE DE FIGURAS",
        "No hay entradas de figuras. Haga clic derecho y seleccione Actualizar campo."
    )


def _insert_list_of_tables(doc: Document, tables: list[tuple[int, str]]):
    """Insert a real Word 'List of Tables' TOC field."""
    if not tables:
        return
    _insert_toc_field(
        doc,
        r'TOC \f T \h',
        "ÍNDICE DE TABLAS",
        "No hay entradas de tablas. Haga clic derecho y seleccione Actualizar campo."
    )


def _walk(doc: Document, text: str, base_dir: Path, figures: list[tuple[int, str]] | None = None, tables: list[tuple[int, str]] | None = None):
    """
    Scan `text` linearly.  At each structural command do the right thing,
    otherwise accumulate plain/inline text and flush as a paragraph.
    """
    pending: list[str] = []

    # Section / chapter counters
    sec_counters = {"chapter": 0, "section": 0, "subsection": 0, "subsubsection": 0, "paragraph": 0}
    appendix_mode = False
    has_chapters = False

    # Current font state propagated across paragraph breaks
    fmt_bold = False
    fmt_italic = False
    fmt_size = PT_NORM

    def flush():
        nonlocal pending, fmt_bold, fmt_italic, fmt_size
        raw = " ".join(pending).strip()
        if raw:
            _add_para(doc, raw, init_bold=fmt_bold, init_italic=fmt_italic, init_size=fmt_size)
            fmt_bold, fmt_italic, fmt_size = _scan_format_switches(raw, fmt_bold, fmt_italic, fmt_size)
        pending = []
    
    def format_section_number(cmd: str) -> str:
        """Generate section number like 1, 1.1, 1.1.1, or A.1"""
        ch = sec_counters['chapter']
        if not has_chapters:
            # No chapters → flat numbering as before
            if cmd == "section":
                return f"{sec_counters['section']}"
            elif cmd == "subsection":
                return f"{sec_counters['section']}.{sec_counters['subsection']}"
            elif cmd == "subsubsection":
                return f"{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}"
            elif cmd == "paragraph":
                return f"{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}.{sec_counters['paragraph']}"
            return ""
        # With chapters → hierarchical numbering
        ch_label = chr(ord('A') + ch - 1) if appendix_mode else str(ch)
        if cmd == "chapter":
            return ch_label
        elif cmd == "section":
            return f"{ch_label}.{sec_counters['section']}"
        elif cmd == "subsection":
            return f"{ch_label}.{sec_counters['section']}.{sec_counters['subsection']}"
        elif cmd == "subsubsection":
            return f"{ch_label}.{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}"
        elif cmd == "paragraph":
            return f"{ch_label}.{sec_counters['section']}.{sec_counters['subsection']}.{sec_counters['subsubsection']}.{sec_counters['paragraph']}"
        return ""

    pos = 0
    n   = len(text)

    while pos < n:
        # Display math
        dm = _DISP_MATH.match(text, pos)
        if dm:
            flush()
            p = doc.add_paragraph()
            omml = None
            if TEMP_DIR and PANDOC_PATH:
                try:
                    omml = latex_to_omml_element(dm.group(0), TEMP_DIR, PANDOC_PATH)
                except Exception:
                    pass
            if omml is not None:
                p._p.append(omml)
            else:
                math_text = strip_fmt(dm.group(1))
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _run(p, math_text, italic=True, size=PT_NORM)
            pos = dm.end()
            continue

        # Page break
        pb = re.match(r"\\(?:newpage|clearpage)\b", text[pos:])
        if pb:
            flush()
            _safe_page_break(doc)
            pos += pb.end()
            continue

        # Skip \maketitle, \tableofcontents (already handled)
        skip = re.match(r"\\(?:maketitle|tableofcontents)\b", text[pos:])
        if skip:
            pos += skip.end()
            continue
        
        # \listoffigures
        lof_m = re.match(r"\\listoffigures\b", text[pos:])
        if lof_m:
            flush()
            if figures:
                _insert_list_of_figures(doc, figures)
            pos += lof_m.end()
            continue
        
        # \listoftables
        lot_m = re.match(r"\\listoftables\b", text[pos:])
        if lot_m:
            flush()
            if tables:
                _insert_list_of_tables(doc, tables)
            pos += lot_m.end()
            continue
        
        # \appendix — switch to appendix mode
        app_m = re.match(r"\\appendix\b", text[pos:])
        if app_m:
            appendix_mode = True
            pos += app_m.end()
            continue
        
        # Chapter headings
        ch_m = re.match(
            r"\\(chapter)\*?\s*\{((?:[^{}]|\{[^{}]*\})*)\}",
            text[pos:]
        )
        if ch_m:
            flush()
            _safe_page_break(doc)
            has_chapters = True
            sec_counters["chapter"] += 1
            sec_counters["section"] = 0
            sec_counters["subsection"] = 0
            sec_counters["subsubsection"] = 0
            sec_counters["paragraph"] = 0
            htxt = strip_fmt(ch_m.group(2))
            num_str = format_section_number("chapter")
            prefix = "Apéndice " if appendix_mode else "Capítulo "
            full_title = f"{prefix}{num_str}  {htxt}"
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(full_title)
            run.font.name = FONT
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.all_caps = True
            run.font.color.rgb = C_BLACK
            pos += ch_m.end()
            continue

        # Section headings
        sec_m = re.match(
            r"\\(section|subsection|subsubsection|paragraph)(\*)?\s*\{((?:[^{}]|\{[^{}]*\})*)\}",
            text[pos:]
        )
        if sec_m:
            flush()
            cmd      = sec_m.group(1)
            starred  = sec_m.group(2) is not None
            htxt     = strip_fmt(sec_m.group(3))

            # Starred versions (\section*) are not numbered and do not update counters
            if not starred:
                # Update counters
                if cmd == "section":
                    sec_counters["section"] += 1
                    sec_counters["subsection"] = 0
                    sec_counters["subsubsection"] = 0
                    sec_counters["paragraph"] = 0
                elif cmd == "subsection":
                    sec_counters["subsection"] += 1
                    sec_counters["subsubsection"] = 0
                    sec_counters["paragraph"] = 0
                elif cmd == "subsubsection":
                    sec_counters["subsubsection"] += 1
                    sec_counters["paragraph"] = 0
                elif cmd == "paragraph":
                    sec_counters["paragraph"] += 1

            # Get level and formatted number
            if has_chapters:
                level = {"section": 2, "subsection": 3,
                         "subsubsection": 4, "paragraph": 5}[cmd]
                sizes = [16, 14, 13, 12, 11]
            else:
                level = {"section": 1, "subsection": 2,
                         "subsubsection": 3, "paragraph": 4}[cmd]
                sizes = [16, 14, 12, 11]
            num_str = "" if starred else format_section_number(cmd)

            # Create heading with number - use explicit style for TOC recognition
            full_title = f"{num_str}  {htxt}".strip()
            style_name = f"Heading {level}"
            p = doc.add_paragraph(style=style_name)
            run = p.add_run(full_title)
            run.font.name = FONT
            run.font.size = Pt(sizes[level-1])
            run.font.bold = True
            run.font.color.rgb = C_BLACK
            pos += sec_m.end()
            continue

        # \begin{env}
        env_m = re.match(r"\\begin\{(\w+)\*?\}", text[pos:])
        if env_m:
            env_name = env_m.group(1)
            flush()
            result = extract_env(text, env_name, pos)
            if result is None:
                pos += env_m.end()
                continue
            _, end_pos, inner = result

            if env_name in ("itemize", "enumerate", "description"):
                # Extract optional arguments like [label=..., leftmargin=...]
                opt_arg, inner = _extract_optional_arg(inner)
                label_format = _parse_enumitem_label(opt_arg) if opt_arg else None
                # Apply enumitem default label for enumerate if no explicit option
                if env_name == "enumerate" and label_format is None:
                    label_format = _get_enumitem_default(1)
                render_list(doc, inner, ordered=(env_name == "enumerate"), base_dir=base_dir, label_format=label_format)

            elif env_name in ("table", "table*"):
                cap_txt = _extract_caption_content(inner)
                caption = cap_txt if cap_txt is not None else ""
                # Find tabular/tabularx inside
                for tenv in ("tabularx", "tabular"):
                    tr = extract_env(inner, tenv)
                    if tr:
                        _, _, tab_inner = tr
                        # remove column spec: first {...}
                        col_spec, tab_inner = _strip_column_spec(tab_inner)
                        # remove \resizebox wrapper
                        tab_inner = re.sub(r"\\resizebox\{[^}]*\}\{[^}]*\}\{?\s*", "", tab_inner)
                        tab_inner = re.sub(r"\s*\}?\s*$", "", tab_inner)
                        render_table(doc, tab_inner, caption, col_spec=col_spec)
                        break

            elif env_name == "longtable":
                cap_txt = _extract_caption_content(inner)
                caption = cap_txt if cap_txt is not None else ""
                # strip column spec
                col_spec, inner = _strip_column_spec(inner)
                # strip \endfirsthead...\endhead  and  \endfoot...\endlastfoot
                inner = re.sub(r"\\endfirsthead.*?\\endhead",     "", inner, flags=re.DOTALL)
                inner = re.sub(r"\\endfoot.*?\\endlastfoot",       "", inner, flags=re.DOTALL)
                inner = _remove_captions(inner)
                render_table(doc, inner, caption, col_spec=col_spec)

            elif env_name in ("tabular", "tabularx"):
                # tabular/tabularx directo (sin \begin{table} wrapper)
                col_spec, inner = _strip_column_spec(inner)
                # remove \resizebox wrapper
                inner = re.sub(r"\\resizebox\{[^}]*\}\{[^}]*\}\{?\s*", "", inner)
                inner = re.sub(r"\s*\}?\s*$", "", inner)
                render_table(doc, inner, caption="", col_spec=col_spec)

            elif env_name == "center":
                n_para_before = len(doc.paragraphs)
                n_tables_before = len(doc.tables)
                _walk(doc, inner, base_dir)
                # Center all paragraphs and tables added inside this center environment
                for p in doc.paragraphs[n_para_before:]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for t in doc.tables[n_tables_before:]:
                    t.alignment = WD_TABLE_ALIGNMENT.CENTER

            elif env_name == "tikzpicture":
                tikz_code = r"\begin{tikzpicture}" + inner + r"\end{tikzpicture}"
                png_path = compile_tikz_to_png(tikz_code, PREAMBLE_GLOBAL, TEMP_DIR, TIKZ_COUNTER[0])
                TIKZ_COUNTER[0] += 1
                if png_path and png_path.exists():
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if not _insert_image_with_size(p, png_path, max_width_inches=6.3):
                        _run(p, "[DIAGRAM]", italic=True, color=C_GRAY)
                else:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    _run(p, "[DIAGRAM — see PDF version]", italic=True, color=C_GRAY)

            elif env_name == "figure":
                # Intentar renderizar como tabla si hay múltiples \subfigure
                rendered_as_table = _render_figure_as_table(doc, inner, base_dir)
                
                if not rendered_as_table:
                    # Procesar TODAS las \includegraphics dentro del entorno figure
                    img_matches = list(re.finditer(r"\\includegraphics(?:\[([^\]]*)\])?\{([^}]*)\}", inner))
                    if img_matches:
                        for img_m in img_matches:
                            opts = img_m.group(1) or ""
                            img_name = img_m.group(2)
                            img_path = resolve_image_path(img_name, base_dir)
                            p = doc.add_paragraph()
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            if img_path and img_path.exists():
                                try:
                                    from PIL import Image
                                    img = Image.open(str(img_path))
                                    px_w, px_h = img.size
                                    nat_w = px_w / 96.0
                                    nat_h = px_h / 96.0
                                    w, h = _parse_graphics_size(opts, img_path)
                                    # Si solo especificaron height, calcular width proporcional
                                    if h is not None and w is None:
                                        w = nat_w * (h / nat_h)
                                    # Si no se especificó nada, usar ancho por defecto
                                    if w is None:
                                        w = 4.0
                                        h = nat_h * (w / nat_w)
                                    # Si excede el ancho de página, escalar proporcionalmente
                                    if w > TEXT_WIDTH_INCHES:
                                        scale = TEXT_WIDTH_INCHES / w
                                        w *= scale
                                        h *= scale
                                    p.add_run().add_picture(str(img_path), width=Inches(w), height=Inches(h))
                                except Exception:
                                    _run(p, f"[Figure: {img_name}]", italic=True)
                            else:
                                _run(p, f"[Figure: {img_name}]", italic=True)
                    else:
                        # Check for tikzpicture inside figure
                        tikz_result = extract_env(inner, "tikzpicture")
                        if tikz_result:
                            _, _, tikz_inner = tikz_result
                            tikz_code = r"\begin{tikzpicture}" + tikz_inner + r"\end{tikzpicture}"
                            png_path = compile_tikz_to_png(tikz_code, PREAMBLE_GLOBAL, TEMP_DIR, TIKZ_COUNTER[0])
                            TIKZ_COUNTER[0] += 1
                            if png_path and png_path.exists():
                                p = doc.add_paragraph()
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                if not _insert_image_with_size(p, png_path, max_width_inches=6.3):
                                    _run(p, "[DIAGRAM]", italic=True, color=C_GRAY)
                            else:
                                p = doc.add_paragraph()
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                _run(p, "[DIAGRAM — see PDF version]", italic=True, color=C_GRAY)
                
                # Caption principal del figure (siempre, independientemente del modo de renderizado)
                fig_cap_txt = _extract_caption_content(inner)
                if fig_cap_txt is not None:
                    FIGURE_COUNTER[0] += 1
                    cp = doc.add_paragraph()
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        cp.style = "Caption"
                    except Exception:
                        pass
                    fig_clean = strip_fmt(fig_cap_txt).replace('"', "'")
                    fig_caption = f"Figure {FIGURE_COUNTER[0]}: " + fig_clean
                    _run(cp, fig_caption,
                         italic=True, size=PT_SMALL, color=C_BLACK)
                    # Hidden TC field for List of Figures
                    tc_text = f"Figura {FIGURE_COUNTER[0]}\t{fig_clean}"
                    _add_tc_field(cp, tc_text, "F")

            elif env_name in ("equation", "equation*", "align", "align*", "gather", "gather*", "multline", "multline*", "eqnarray", "eqnarray*"):
                full_env = f"\\begin{{{env_name}}}" + inner + f"\\end{{{env_name}}}"
                # Pandoc cannot convert display math containing \label{...};
                # strip labels before sending to Pandoc.
                full_env_for_pandoc = re.sub(r"\\label\{[^}]*\}", "", full_env)
                p = doc.add_paragraph()
                omml = None
                if TEMP_DIR and PANDOC_PATH:
                    try:
                        omml = latex_to_omml_element(full_env_for_pandoc, TEMP_DIR, PANDOC_PATH)
                    except Exception:
                        pass
                if omml is not None:
                    p._p.append(omml)
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    _run(p, strip_fmt(inner), italic=True, size=PT_NORM)

            elif env_name == "tcolorbox":
                # title from options [title={...}]
                opt_str = text[pos + env_m.end():]
                opt_m   = re.match(r"\[([^\]]*)\]", opt_str)
                box_title = ""
                if opt_m:
                    tm = re.search(r"title=\{?([^,}\]]+)\}?", opt_m.group(1))
                    if tm:
                        box_title = strip_fmt(tm.group(1))
                if box_title:
                    p = doc.add_paragraph()
                    _run(p, box_title, bold=True, size=PT_NORM)
                _walk(doc, inner, base_dir)

            elif env_name in ("document",):
                _walk(doc, inner, base_dir)
                end_pos = len(text)   # consumed everything

            # else: unknown env — walk its inner content
            else:
                _walk(doc, inner, base_dir)

            pos = end_pos
            continue

        # \end{...} — should not appear here if extract_env worked, but skip it
        end_m = re.match(r"\\end\{[^}]*\}", text[pos:])
        if end_m:
            pos += end_m.end()
            continue

        # Double blank line → paragraph break
        blank = re.match(r"\n[ \t]*\n", text[pos:])
        if blank:
            flush()
            pos += blank.end()
            continue

        # Advance to next structural boundary or gather text chunk
        nxt = _STRUCT.search(text, pos)
        # Also stop at display math and double blanks
        dm2  = _DISP_MATH.search(text, pos)
        nxt2 = re.search(r"\n[ \t]*\n", text[pos:])

        candidates = [c for c in [
            nxt.start()          if nxt  else None,
            dm2.start()          if dm2  else None,
            pos + nxt2.start()   if nxt2 else None,
        ] if c is not None]

        stop = min(candidates) if candidates else n
        if stop == pos:
            # Avoid infinite loop
            pos += 1
            continue

        chunk = text[pos:stop]
        # Convert LaTeX line breaks \\[...] to paragraph breaks
        chunk = re.sub(r"\\\\(?:\[[^\]]*\])?", "\n\n", chunk)
        # Convert LaTeX newlines to paragraph breaks
        chunk = chunk.replace("\\newline", "\n\n")
        pos   = stop

        # Collect non-empty lines
        _VAR_LINE = re.compile(r"^(\$[^$]+\$|\\(?:gls|Gls)\{[^}]*\})")
        prev_was_var = False
        for line in chunk.split("\n"):
            line = line.strip()
            if not line:
                flush()
                prev_was_var = False
                continue
            # skip pure layout/control commands
            if re.match(
                r"\\(?:maketitle|tableofcontents|vspace|hspace|noindent|centering"
                r"|par\b|medskip|bigskip|smallskip|clearpage|newpage"
                r"|label|ref|pageref|addcontentsline|setlength"
                r"|pagestyle|thispagestyle|fancyhf|lhead|rhead|cfoot|headheight"
                r"|usepackage|documentclass|title|author|date"
                r"|usetikzlibrary|tcbuselibrary|hypersetup|rowcolors"
                r"|arrayrulecolor|cellcolor|newcounter|setcounter"
                r"|newcommand|renewcommand|setlist)\b",
                line
            ):
                continue
            is_var = bool(_VAR_LINE.match(line))
            if is_var and prev_was_var:
                flush()
            pending.append(line)
            prev_was_var = is_var

    flush()


def _maybe_add_tab_stop(para, text: str):
    """Add a tab stop if the cleaned text contains tab characters."""
    if "\t" in _clean(text):
        para.paragraph_format.tab_stops.add_tab_stop(Inches(1.5))


def _parse_leading_switches(text: str, bold: bool, italic: bool, size: int):
    r"""Extract initial font switches (\bfseries, \large, etc.) and update state."""
    size_map = {
        "tiny": 6, "scriptsize": 8, "footnotesize": 9, "small": 10,
        "normalsize": PT_NORM, "large": 14, "Large": 16, "LARGE": 20,
        "huge": 24, "Huge": 28,
    }
    pattern = re.compile(
        r"^\s*\\(bfseries|normalfont|itshape|"
        r"tiny|scriptsize|footnotesize|small|normalsize|large|Large|LARGE|huge|Huge)\b\s*"
    )
    while True:
        m = pattern.match(text)
        if not m:
            break
        cmd = m.group(1)
        if cmd == "bfseries":
            bold = True
        elif cmd == "normalfont":
            bold = False
            italic = False
        elif cmd == "itshape":
            italic = True
        elif cmd in size_map:
            size = size_map[cmd]
        text = text[m.end():]
    return bold, italic, size, text


def _scan_format_switches(text: str, bold: bool, italic: bool, size: int):
    r"""Scan text for font switches and return the final formatting state."""
    size_map = {
        "tiny": 6, "scriptsize": 8, "footnotesize": 9, "small": 10,
        "normalsize": PT_NORM, "large": 14, "Large": 16, "LARGE": 20,
        "huge": 24, "Huge": 28,
    }
    pattern = re.compile(
        r"\\(bfseries|normalfont|itshape|"
        r"tiny|scriptsize|footnotesize|small|normalsize|large|Large|LARGE|huge|Huge)\b"
    )
    for m in pattern.finditer(text):
        cmd = m.group(1)
        if cmd == "bfseries":
            bold = True
        elif cmd == "normalfont":
            bold = False
            italic = False
        elif cmd == "itshape":
            italic = True
        elif cmd in size_map:
            size = size_map[cmd]
    return bold, italic, size


def _add_para(doc: Document, raw: str, init_bold=False, init_italic=False, init_size=None):
    r"""Create justified paragraph(s) from raw inline LaTeX.

    Handles paragraph breaks (\\ or blank lines) and propagates font switches
    such as \bfseries / \large across the generated paragraphs.
    """
    raw = raw.strip()
    if not raw:
        return None
    paragraphs = re.split(r"\n\s*\n", raw)
    cur_bold = bool(init_bold)
    cur_italic = bool(init_italic)
    cur_size = init_size if init_size is not None else PT_NORM
    last_p = None
    for para_text in paragraphs:
        para_text = para_text.strip()
        if not para_text:
            continue
        cur_bold, cur_italic, cur_size, para_text = _parse_leading_switches(
            para_text, cur_bold, cur_italic, cur_size
        )
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        parse_inline(p, para_text, base_sz=PT_NORM,
                     bold=cur_bold, italic=cur_italic, size=cur_size)
        _maybe_add_tab_stop(p, para_text)
        last_p = p
    return last_p


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def extract_tables_to_temp(source: str, temp_dir: Path) -> int:
    """Extrae tablas del LaTeX en orden de aparicion y las guarda en carpeta temporal."""
    # Buscar todos los entornos de tabla en orden de aparicion
    # El source tiene backslashes simples (ya que es texto leido del archivo)
    pattern = r"\\begin\{(longtable|tabular|tabularx)\*?\}(.*?)\\end\{\1\*?\}"
    table_count = 0
    
    for match in re.finditer(pattern, source, re.DOTALL):
        table_count += 1
        env_name = match.group(1)
        content = match.group(2)
        
        # Remover \resizebox si existe (como en la tabla de presupuesto)
        # \resizebox{width}{height}{content} -> solo content
        content = re.sub(r"\\resizebox\{[^}]*\}\{[^}]*\}\{?", "", content)
        content = content.rstrip("}")  # Remover el cierre del resizebox si existe
        
        # Guardar archivo .tex (keep inline math $...$ intact for Pandoc)
        tex_file = temp_dir / f"tabla_{table_count:02d}_{env_name}.tex"
        tex_file.write_text(f"\\begin{{{env_name}}}" + content + f"\\end{{{env_name}}}", encoding="utf-8")
    
    return table_count


def convert_tables_with_pandoc(temp_dir: Path, pandoc_path: str = "pandoc") -> bool:
    """Convierte todas las tablas .tex en temp_dir a .docx con Pandoc."""
    tex_files = list(temp_dir.glob("tabla_*.tex"))
    if not tex_files:
        return False
    
    for tex_file in tex_files:
        # Nombre de salida: tabla_XX_pandoc.docx
        num_match = re.search(r'tabla_(\d+)_', tex_file.name)
        if num_match:
            num = num_match.group(1)
            out_file = temp_dir / f"tabla_{num}_pandoc.docx"
            cmd = [pandoc_path, "-f", "latex", "-t", "docx", str(tex_file), "-o", str(out_file)]
            try:
                subprocess.run(cmd, capture_output=True, check=False)
            except Exception:
                pass  # Si falla, continuamos
    return True


def expand_inputs(source: str, base_dir: Path, root_dir: Path = None, depth: int = 0, max_depth: int = 10) -> str:
    r"""Expand \input{file} directives by inlining the referenced files.
    Also handles \IfFileExists{file}{true}{false} by evaluating file existence
    against both the local directory and the project's root directory.
    """
    if depth > max_depth:
        return source
    
    root_dir = root_dir or base_dir
    
    def _resolve_file(filename: str):
        """Find file and return (content, file_dir) or (None, None)."""
        for d in (base_dir, root_dir):
            for ext in ("", ".tex"):
                cand = d / (filename + ext)
                if cand.exists():
                    return cand.read_text(encoding="utf-8"), cand.parent
        return None, None
    
    # 1. Handle \IfFileExists{file}{true}{false}
    def replace_iff(m):
        filename = m.group(1).strip()
        true_br = m.group(2)
        false_br = m.group(3) if m.group(3) is not None else ""
        content, _ = _resolve_file(filename)
        if content is not None:
            return true_br
        return false_br
    
    source = re.sub(
        r"\\IfFileExists\s*\{([^}]+)\}\s*\{((?:[^{}]|\{[^{}]*\})*)\}\s*\{((?:[^{}]|\{[^{}]*\})*)\}",
        replace_iff, source
    )
    
    # 2. Handle \input{file}
    def replace_input(m):
        filename = m.group(1).strip()
        content, file_dir = _resolve_file(filename)
        if content is not None:
            return expand_inputs(content, file_dir, root_dir, depth + 1, max_depth)
        return m.group(0)
    
    return re.sub(r"\\input\{([^}]+)\}", replace_input, source)


def main():
    global TEMP_DIR, PANDOC_PATH
    
    in_path  = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("linea_base_en.tex")
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else in_path.with_suffix(".docx")

    if not in_path.exists():
        print(f"Error: file not found: {in_path}", file=sys.stderr)
        sys.exit(1)

    # Crear carpeta temporal para archivos intermedios
    temp_dir = tempfile.mkdtemp(prefix="latex_conv_")
    TEMP_DIR = Path(temp_dir)
    
    try:
        raw_source = in_path.read_text(encoding="utf-8")
        # Strip comments BEFORE expanding inputs, so commented-out \input / \IfFileExists are ignored
        source = re.sub(r"(?m)(?<!\\)%.*$", "", raw_source)
        source = expand_inputs(source, in_path.parent, root_dir=in_path.parent)
        print(f"[INFO] Inputs expandidos: {len(source)} caracteres")
        
        # Buscar pandoc
        pandoc_exe = "pandoc"
        if sys.platform == "win32":
            pandoc_paths = [
                r"C:\Program Files\Pandoc\pandoc.exe",
                r"C:\ProgramData\chocolatey\bin\pandoc.exe",
            ]
            for pp in pandoc_paths:
                if Path(pp).exists():
                    pandoc_exe = pp
                    break
        PANDOC_PATH = pandoc_exe

        # 1. Extraer tablas a carpeta temporal
        print("[1/3] Extrayendo tablas...")
        n_tables = extract_tables_to_temp(source, TEMP_DIR)
        print(f"       {n_tables} tablas encontradas")
        
        # 2. Convertir tablas con Pandoc
        if n_tables > 0:
            print("[2/3] Convirtiendo tablas con Pandoc...")
            convert_tables_with_pandoc(TEMP_DIR, pandoc_exe)

        # 3. Parse preamble → populate MACROS
        preamble_end = source.find("\\begin{document}")
        preamble = source[:preamble_end] if preamble_end != -1 else source
        parse_preamble(preamble)
        
        # Extract header text from \rhead{...}
        header_text = ""
        # Find all \rhead occurrences and pick the first one with actual text
        for rhead_match in re.finditer(r"\\rhead\s*\{", preamble):
            start = rhead_match.end()
            # Find matching } (simple brace balance)
            depth = 1
            i = start
            while i < len(preamble) and depth > 0:
                if preamble[i] == '{':
                    depth += 1
                elif preamble[i] == '}':
                    depth -= 1
                i += 1
            if depth == 0:
                raw_header = preamble[start:i-1]
                raw_header = re.sub(r"\\small\s*", "", raw_header)
                raw_header = raw_header.replace('%', '').strip()
                # Remove \includegraphics commands (they are rendered as images, not text)
                raw_header = re.sub(r"\\includegraphics(?:\[[^\]]*\])?\{[^}]*\}", "", raw_header)
                candidate = strip_fmt(raw_header).strip()
                if candidate and not candidate.startswith('[') and 'icono' not in candidate.lower():
                    header_text = candidate
                    break

        # Clean header text: remove image path artifacts
        if header_text.strip().startswith('[') or 'icono' in header_text.lower():
            header_text = ""

        # 4. Build document
        print("[3/3] Generando documento Word...")
        doc = Document()
        setup_document(doc)

        # Buscar logo
        script_dir = Path(__file__).parent
        # Prefer the project's own icon/logo
        logo_path = in_path.parent / "icono.png"
        if not logo_path.exists():
            logo_path = in_path.parent / "icono.jpg"
        if not logo_path.exists():
            logo_path = script_dir / "images" / "logo.jpg"
        if not logo_path.exists():
            logo_path = script_dir / "images" / "logo.png"
        if not logo_path.exists():
            logo_path = in_path.parent / "00_figs" / "logo.jpg"
        add_header(doc, logo_path, header_text)
        add_footer(doc)

        parse_body(doc, source, in_path.parent)

        doc.save(str(out_path))
        print(f"[OK] Guardado: {out_path}")
        
    finally:
        # Limpiar carpeta temporal
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()