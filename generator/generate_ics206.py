from __future__ import annotations

from copy import deepcopy
from io import BytesIO
import base64
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from docx.table import Table
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches


def _unique_cells(row):
    out, seen = [], set()
    for c in row.cells:
        key = id(c._tc)
        if key not in seen:
            seen.add(key)
            out.append(c)
    return out


def _cell_text(cell) -> str:
    return "\n".join(p.text for p in cell.paragraphs).strip()


def _row_first_text(row) -> str:
    cells = _unique_cells(row)
    return _cell_text(cells[0]) if cells else ""


def _replace_cell(cell, value: Any, *, preserve_font: bool = True) -> None:
    value = "" if value is None else str(value)
    if not cell.paragraphs:
        p = cell.add_paragraph()
    else:
        p = cell.paragraphs[0]
    if not p.runs:
        p.add_run("")
    p.runs[0].text = value
    for r in p.runs[1:]:
        r.text = ""
    for p2 in cell.paragraphs[1:]:
        for r in p2.runs:
            r.text = ""
    if not preserve_font:
        for p3 in cell.paragraphs:
            for r in p3.runs:
                r.font.name = "Arial"



def _replace_responder_list(cell, responders: list[Any]) -> None:
    """Render all EMS responders in one value cell with 4 pt breathing room.

    The first responder has 4 pt before it, each subsequent responder has
    4 pt before it (creating 4 pt between responders), and the final
    responder has 4 pt after it.
    """
    responders = [str(x).strip() for x in responders if str(x).strip()]

    # Remove existing paragraphs except the first so the template's first
    # paragraph properties remain the formatting baseline.
    if not cell.paragraphs:
        cell.add_paragraph()
    first = cell.paragraphs[0]
    for extra in list(cell.paragraphs[1:]):
        extra._element.getparent().remove(extra._element)

    paragraphs = [first]
    while len(paragraphs) < max(1, len(responders)):
        paragraphs.append(cell.add_paragraph())

    for i, paragraph in enumerate(paragraphs):
        text = responders[i] if i < len(responders) else ""
        if not paragraph.runs:
            paragraph.add_run("")
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.space_before = Pt(4) if responders else Pt(0)
        paragraph.paragraph_format.space_after = Pt(4) if responders and i == len(responders) - 1 else Pt(0)
        for run in paragraph.runs:
            run.font.name = "Arial"

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

def _replace_combined_name_address(cell, name: str, address: str) -> None:
    # Keep the validated cell formatting while rebuilding the two visible lines.
    if not cell.paragraphs:
        p = cell.add_paragraph()
    else:
        p = cell.paragraphs[0]
    if not p.runs:
        p.add_run("")
    p.runs[0].text = name or ""
    if len(p.runs) == 1:
        p.add_run("")
    p.runs[1].text = ("\n" + address) if address else ""
    for r in p.runs[2:]:
        r.text = ""
    for p2 in cell.paragraphs[1:]:
        for r in p2.runs:
            r.text = ""



def _replace_phone_numbers(cell, phone1: str, phone2: str) -> None:
    values = [_format_phone(x) for x in (phone1, phone2) if str(x).strip()]
    for p in cell.paragraphs:
        for r in p.runs:
            r.text = ""
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if values:
        p.add_run(values[0])
    if len(values) > 1:
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(0)
        p2.add_run(values[1])
    keep = 2 if len(values) > 1 else 1
    for extra in list(cell.paragraphs[keep:]):
        extra._element.getparent().remove(extra._element)


def _format_phone(value: Any) -> str:
    """Normalize North American phone numbers without changing short codes.

    Ten-digit numbers are rendered as ``(###) ###-####``. An eleven-digit
    number beginning with 1 is rendered as ``+1 (###) ###-####``. Short
    emergency/service codes such as 911, international numbers, and other
    nonstandard values are preserved exactly as entered. Common extension
    forms are retained as ``ext. ####``.
    """
    original = "" if value is None else str(value).strip()
    if not original:
        return ""

    phone_pattern = re.compile(
        r"(?<!\d)"
        r"(?P<country>\+?1[\s.\-]?)?"
        r"(?:\((?P<area_paren>\d{3})\)|(?P<area_plain>\d{3}))"
        r"[\s.\-]*(?P<prefix>\d{3})[\s.\-]*(?P<line>\d{4})"
        r"(?:\s*(?:ext\.?|x)\s*(?P<extension>\d+))?"
        r"(?!\d)",
        re.IGNORECASE,
    )

    def replace(match: re.Match[str]) -> str:
        area = match.group("area_paren") or match.group("area_plain")
        country = "+1 " if match.group("country") else ""
        formatted = f"{country}({area}) {match.group('prefix')}-{match.group('line')}"
        extension = match.group("extension")
        return f"{formatted} ext. {extension}" if extension else formatted

    # Replace phone-number portions in place so labels and short codes survive,
    # for example ``CWICC 509-884-3743`` and ``911/509-662-5111``.
    return phone_pattern.sub(replace, original)


def _replace_travel_cell(cell, value: Any) -> None:
    """Render hospital travel time without allowing unit words to split.

    Each whitespace-delimited token gets its own paragraph so Word may stack
    values as ``1 / hr / 28 / mins`` in the narrow travel-time columns.  The
    cell itself is marked no-wrap, which keeps unit words such as ``mins``,
    ``hr`` and ``hrs`` intact instead of breaking them into ``min`` / ``s``.
    """
    text = "" if value is None else str(value).strip()
    tokens = text.split() if text else [""]

    if not cell.paragraphs:
        cell.add_paragraph()
    first = cell.paragraphs[0]
    for extra in list(cell.paragraphs[1:]):
        extra._element.getparent().remove(extra._element)

    paragraphs = [first]
    while len(paragraphs) < len(tokens):
        paragraphs.append(cell.add_paragraph())

    for paragraph, token in zip(paragraphs, tokens):
        if not paragraph.runs:
            paragraph.add_run("")
        paragraph.runs[0].text = token
        for run in paragraph.runs[1:]:
            run.text = ""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.name = "Arial"
            # The Air/Ground travel-time columns are extremely narrow. Word may
            # still split a unit token such as "mins" into "min" / "s" even
            # when it is the only text in the paragraph. Reduce only the unit
            # token size so the complete unit fits on one line; the number and
            # unit may remain on separate lines.
            if token.lower() in {"min", "mins", "hr", "hrs"}:
                run.font.size = Pt(6)

    # Prevent Word from splitting a token in the middle of the word. Because
    # each token is already in its own paragraph, numbers and units may still
    # appear on separate lines exactly as intended.
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:noWrap")):
        tcPr.remove(old)
    tcPr.append(OxmlElement("w:noWrap"))


def _inline_picture_to_anchor(inline, x_emu: int, y_emu: int) -> None:
    """Convert an inline picture to the floating placement used by the approved reference."""
    graphic = deepcopy(inline.graphic)
    extent = deepcopy(inline.extent)
    doc_pr = deepcopy(inline.docPr)

    anchor = OxmlElement("wp:anchor")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "0")
    anchor.set("distR", "0")
    anchor.set("simplePos", "0")
    anchor.set("relativeHeight", "251658240")
    anchor.set("behindDoc", "0")
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1")

    simple_pos = OxmlElement("wp:simplePos")
    simple_pos.set("x", "0")
    simple_pos.set("y", "0")
    anchor.append(simple_pos)

    pos_h = OxmlElement("wp:positionH")
    pos_h.set("relativeFrom", "column")
    off_h = OxmlElement("wp:posOffset")
    off_h.text = str(x_emu)
    pos_h.append(off_h)
    anchor.append(pos_h)

    pos_v = OxmlElement("wp:positionV")
    pos_v.set("relativeFrom", "page")
    off_v = OxmlElement("wp:posOffset")
    off_v.text = str(y_emu)
    pos_v.append(off_v)
    anchor.append(pos_v)

    anchor.append(extent)

    effect_extent = OxmlElement("wp:effectExtent")
    for side in ("l", "t", "r", "b"):
        effect_extent.set(side, "0")
    anchor.append(effect_extent)

    anchor.append(OxmlElement("wp:wrapNone"))
    anchor.append(doc_pr)

    frame_pr = OxmlElement("wp:cNvGraphicFramePr")
    locks = OxmlElement("a:graphicFrameLocks")
    locks.set("noChangeAspect", "1")
    frame_pr.append(locks)
    anchor.append(frame_pr)
    anchor.append(graphic)

    inline.getparent().replace(inline, anchor)


def _replace_part_children(target_part, source_part) -> None:
    """Replace header/footer XML children with deep copies from another part."""
    target = target_part._element
    source = source_part._element
    for child in list(target):
        target.remove(child)
    for child in list(source):
        target.append(deepcopy(child))


def _add_logo_to_header(header, raw: bytes) -> None:
    title_p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    run = title_p.add_run()
    shape = run.add_picture(BytesIO(raw), width=Inches(1.15))
    inline = shape._inline

    # Preserve the known-working v0.10.87 logo geometry.
    x = 228600
    y = 45720
    _inline_picture_to_anchor(inline, x, y)


def _apply_optional_logo(doc: Document, incident: dict[str, Any]) -> None:
    """Apply logo with extra clearance beginning on page 2 only.

    When a logo is used:
      - Page 1 uses a dedicated first-page header with ONE blank return.
      - Pages 2+ use the normal repeating header with TWO blank returns.
    This matches the user's validated logo/header reference, including the
    larger continuation-page clearance beginning on page 2.
    """
    data_url = (incident.get("incident") or {}).get("logoDataUrl") or ""
    if not data_url or "," not in data_url:
        return

    try:
        raw = base64.b64decode(data_url.split(",", 1)[1])
        section = doc.sections[0]

        # Capture the original approved repeating header BEFORE inserting any logo.
        default_header = section.header
        first_header = section.first_page_header

        # Create a true different-first-page layout. This is the Word-native way
        # to make spacing start on page 2 rather than changing page 1.
        section.different_first_page_header_footer = True

        # First page header starts as an exact copy of the original approved header.
        _replace_part_children(first_header, default_header)

        # Because "Different First Page" also affects the footer, preserve the
        # existing footer on page 1 as well.
        _replace_part_children(section.first_page_footer, section.footer)

        # Add the same v0.10.87 logo to both header variants.
        _add_logo_to_header(first_header, raw)
        _add_logo_to_header(default_header, raw)

        # Normalize both header variants first so repeated generation cannot
        # accumulate additional blank returns.
        while first_header.paragraphs and not first_header.paragraphs[-1].text.strip():
            p_el = first_header.paragraphs[-1]._element
            p_el.getparent().remove(p_el)
        while default_header.paragraphs and not default_header.paragraphs[-1].text.strip():
            p_el = default_header.paragraphs[-1]._element
            p_el.getparent().remove(p_el)

        def _append_header_blank(header, source_p):
            blank = header.add_paragraph()
            if source_p._p.pPr is not None:
                blank_p = blank._p
                if blank_p.pPr is not None:
                    blank_p.remove(blank_p.pPr)
                blank_p.insert(0, deepcopy(source_p._p.pPr))

        # Exact logo/header spacing validated against:
        #   3 ICS206_Little_Giant_Fire_20260823.docx
        #
        # PAGE 1: one blank return after the CUI line.
        first_title_p = first_header.paragraphs[0]
        _append_header_blank(first_header, first_title_p)

        # PAGES 2+: two blank returns after the CUI line. Because this is the
        # normal repeating header, the same spacing automatically carries to
        # every continuation page, not just page 2.
        default_title_p = default_header.paragraphs[0]
        _append_header_blank(default_header, default_title_p)
        _append_header_blank(default_header, default_title_p)

    except Exception:
        return

def _apply_signatures(doc: Document, incident: dict[str, Any]) -> None:
    sig = incident.get("_signatures") or {}
    if not any(sig.get(k) for k in ("medlSignature", "medlTime", "safetySignature", "safetyTime")):
        return

    def clear(cell):
        for p in cell.paragraphs:
            for r in p.runs:
                r.text = ""

    def add_sig(cell, data_url):
        clear(cell)
        p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if data_url and "," in data_url:
            try:
                raw = base64.b64decode(data_url.split(",", 1)[1])
                p.add_run().add_picture(BytesIO(raw), width=Inches(1.55), height=Inches(0.34))
            except Exception:
                pass

    def add_time(cell, value):
        clear(cell)
        p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if value:
            p.add_run(str(value))

    # Locate the footer by its labels so variable numbers of Locations do not matter.
    for table in reversed(doc.tables):
        for i, row in enumerate(table.rows[:-1]):
            labels = " | ".join(_cell_text(c) for c in _unique_cells(row)).lower()
            if "prepared by" in labels and "reviewed by" in labels:
                cells = _unique_cells(table.rows[i + 1])
                if len(cells) >= 4:
                    add_sig(cells[0], sig.get("medlSignature"))
                    add_time(cells[1], sig.get("medlTime", ""))
                    add_sig(cells[2], sig.get("safetySignature"))
                    add_time(cells[3], sig.get("safetyTime", ""))
                return


def _set_arial_everywhere(doc: Document) -> None:
    try:
        doc.styles["Normal"].font.name = "Arial"
    except Exception:
        pass
    for table in doc.tables:
        for row in table.rows:
            seen = set()
            for cell in row.cells:
                key = id(cell._tc)
                if key in seen:
                    continue
                seen.add(key)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = "Arial"


def _set_small_hospital_labels(doc: Document) -> None:
    t = doc.tables[0]
    for row in t.rows:
        seen = set()
        for cell in row.cells:
            key = id(cell._tc)
            if key in seen:
                continue
            seen.add(key)
            if _cell_text(cell) in {"Lat:", "Long:", "VHF:"}:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(8)
                        r.font.bold = True


def _add_cant_split(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))


def _set_cell_border_none(cell, side: str) -> None:
    """Remove one cell border side while preserving all other borders."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    edge = tcBorders.find(qn(f"w:{side}"))
    if edge is None:
        edge = OxmlElement(f"w:{side}")
        tcBorders.append(edge)

    edge.set(qn("w:val"), "nil")


def _remove_cell_top_border(cell) -> None:
    _set_cell_border_none(cell, "top")


def _remove_cell_bottom_border(cell) -> None:
    _set_cell_border_none(cell, "bottom")


def _shrink_row_height(row, factor: float = 1 / 3) -> None:
    """Force a compact spacer row to roughly one-third normal blank-row height."""
    trPr = row._tr.get_or_add_trPr()
    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)

    # 60 twips = 3 pt. This is intentionally compact and avoids Word expanding
    # the blank spacer back toward a normal text-row height.
    trHeight.set(qn("w:val"), "60")
    trHeight.set(qn("w:hRule"), "exact")

    # Remove paragraph spacing/line-height pressure inside the blank row.
    for p in row._tr.xpath(".//w:p"):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)

        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "20")
        spacing.set(qn("w:lineRule"), "exact")

        # Ensure any empty paragraph mark cannot force the row taller.
        rPr = pPr.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            pPr.append(rPr)
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), "2")
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            rPr.append(szCs)
        szCs.set(qn("w:val"), "2")


def _set_keep_next(row, enabled: bool = True) -> None:
    # Apply directly to every paragraph XML in the row, including paragraphs
    # inside merged/continuation cells. This is required to keep a full
    # Division/Location block together across page boundaries in Word.
    for p in row._tr.xpath('.//w:p'):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)
        existing = pPr.find(qn("w:keepNext"))
        if enabled and existing is None:
            pPr.append(OxmlElement("w:keepNext"))
        elif not enabled and existing is not None:
            pPr.remove(existing)


def _protect_block(rows: Iterable, include_last_keep: bool = False) -> None:
    rows = list(rows)
    for i, row in enumerate(rows):
        _add_cant_split(row)
        _set_keep_next(row, include_last_keep or i < len(rows) - 1)


HEADER_GRAY = "D9D9D9"


def _set_cell_fill(cell, fill: str = HEADER_GRAY) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def _style_header_row(row, *, font_size: float | None = None) -> None:
    """Apply the one approved gray and centered header treatment."""
    for cell in _unique_cells(row):
        _set_cell_fill(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.name = "Arial"
                run.font.bold = True
                if font_size is not None:
                    run.font.size = Pt(font_size)


def _style_detail_label(cell) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(8)
            r.font.bold = True


def _style_large_identifier(cell) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(14)
            r.font.bold = True


def _format_date(d: str) -> str:
    if not d:
        return ""
    try:
        return datetime.strptime(d, "%Y-%m-%d").strftime("%m/%d/%Y")
    except ValueError:
        return d


def _operational_period(incident: dict[str, Any]) -> str:
    s = incident.get("operationalPeriodStart") or {}
    e = incident.get("operationalPeriodEnd") or {}
    if isinstance(s, str):
        return s
    return f"{_format_date(s.get('date',''))} {s.get('time','')} - {_format_date(e.get('date',''))} {e.get('time','')}".strip()


def _format_travel(raw: Any) -> str:
    if raw is None:
        return ""
    text = str(raw).strip()
    if not text:
        return ""
    if text.upper().replace("/", "") == "NA":
        return "NA"
    try:
        mins = int(round(float(text)))
    except ValueError:
        return text
    if mins < 60:
        return f"{mins} mins"
    h, m = divmod(mins, 60)
    if m == 0:
        return "1 hr" if h == 1 else f"{h} hrs"
    return f"1 hr {m} mins" if h == 1 else f"{h} hrs {m} mins"


def _final_ground(h: dict[str, Any]) -> str:
    return h.get("finalGroundDisplay") or _format_travel(
        h.get("finalGroundMin") or h.get("groundTimeOverride") or h.get("calculatedGroundMin")
    )


def _final_air(h: dict[str, Any]) -> str:
    return h.get("finalAirDisplay") or _format_travel(
        h.get("finalAirMin") or h.get("airTimeOverride") or h.get("calculatedAirMin")
    )


def _remove_rows_between(table, start_idx: int, end_idx: int) -> None:
    # end exclusive; remove backwards
    for idx in range(end_idx - 1, start_idx - 1, -1):
        table._tbl.remove(table.rows[idx]._tr)


def _insert_rows_before(table, marker_row, row_elements: list) -> None:
    for element in row_elements:
        marker_row._tr.addprevious(deepcopy(element))


def _find_row(table, predicate) -> int:
    for i, row in enumerate(table.rows):
        if predicate(row):
            return i
    raise ValueError("Required template row not found")


def _build_table0(doc: Document, ref: Document, data: dict[str, Any]) -> None:
    # Use the entire approved Table 0 as the formatting authority, including
    # its exact tblGrid. Copying only individual rows into a different grid
    # can change Ambulance/Air/Hospital widths and break the first Branch
    # header even when the row XML itself is correct.
    old_t = doc.tables[0]
    ref_t = ref.tables[0]
    clone = deepcopy(ref_t._tbl)
    parent = old_t._element.getparent()
    parent.replace(old_t._element, clone)
    t = Table(clone, doc._body)
    rt = ref_t
    incident = data.get("incident", {})

    # Incident header values.
    cells = _unique_cells(t.rows[1])
    _replace_cell(cells[0], incident.get("name", ""))
    _replace_cell(cells[1], _operational_period(incident))
    for c in cells:
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Ground ambulance rows.
    air_header_idx = _find_row(t, lambda r: _row_first_text(r) == "Air Ambulance Services")
    ground_start = 4
    _remove_rows_between(t, ground_start, air_header_idx)
    air_header_idx = _find_row(t, lambda r: _row_first_text(r) == "Air Ambulance Services")
    ground = data.get("groundAmbulances") or [{}]
    source_yes = rt.rows[4]._tr
    source_no = rt.rows[6]._tr
    new = []
    for g in ground:
        template = source_yes if str(g.get("als", "")).lower() == "yes" else source_no
        new.append(template)
    _insert_rows_before(t, t.rows[air_header_idx], new)

    # Air ambulance rows.
    hosp_header_idx = _find_row(t, lambda r: _row_first_text(r) == "Hospitals")
    air_data_start = _find_row(t, lambda r: _row_first_text(r) == "Name" and any("Type of Aircraft" in _cell_text(c) for c in _unique_cells(r))) + 1
    _remove_rows_between(t, air_data_start, hosp_header_idx)
    hosp_header_idx = _find_row(t, lambda r: _row_first_text(r) == "Hospitals")
    air = data.get("airAmbulances") or [{}]
    _insert_rows_before(t, t.rows[hosp_header_idx], [rt.rows[9]._tr for _ in air])

    # Hospital blocks.
    note_idx = _find_row(t, lambda r: _row_first_text(r).startswith("*Travel times are from"))
    hosp_col_header = _find_row(t, lambda r: any("GPS Datum" in _cell_text(c) for c in _unique_cells(r)))
    _remove_rows_between(t, hosp_col_header + 1, note_idx)
    note_idx = _find_row(t, lambda r: _row_first_text(r).startswith("*Travel times are from"))
    hospitals = data.get("hospitals") or [{}]
    blocks = []
    for h in hospitals:
        # The reference has a validated No block (14-16) and Yes block (17-19).
        rows = rt.rows[17:20] if str(h.get("helipad", "")).lower() == "yes" else rt.rows[14:17]
        blocks.extend([r._tr for r in rows])
    _insert_rows_before(t, t.rows[note_idx], blocks)

    # Save/reopen is handled by caller before data-row population.


def _populate_table0_rows(doc: Document, data: dict[str, Any]) -> None:
    t = doc.tables[0]
    incident = data.get("incident", {})

    # Ground rows are between ambulance headings and Air Ambulance Services.
    ground_header = _find_row(t, lambda r: _row_first_text(r) == "Ambulance Services")
    air_header = _find_row(t, lambda r: _row_first_text(r) == "Air Ambulance Services")
    ground_rows = t.rows[ground_header + 2:air_header]
    for row, g in zip(ground_rows, data.get("groundAmbulances") or [{}]):
        cells = _unique_cells(row)
        _replace_cell(cells[0], g.get("name", ""))
        _replace_cell(cells[1], g.get("address", ""))
        phone = _format_phone(g.get("phone", ""))
        freq = g.get("emsFrequency", "")
        combined = phone if not freq else f"{phone}\n{freq}" if phone else freq
        _replace_cell(cells[2], combined)

    # Air rows.
    hospital_header = _find_row(t, lambda r: _row_first_text(r) == "Hospitals")
    air_col_header = _find_row(t, lambda r: _row_first_text(r) == "Name" and any("Type of Aircraft" in _cell_text(c) for c in _unique_cells(r)))
    air_rows = t.rows[air_col_header + 1:hospital_header]
    for row, a in zip(air_rows, data.get("airAmbulances") or [{}]):
        cells = _unique_cells(row)
        _replace_cell(cells[0], a.get("name", ""))
        _replace_cell(cells[1], _format_phone(a.get("phone", "")))
        _replace_cell(cells[2], a.get("aircraftCapability", ""))

    # Hospitals.
    hospital_col_header = _find_row(t, lambda r: any("GPS Datum" in _cell_text(c) for c in _unique_cells(r)))
    note_idx = _find_row(t, lambda r: _row_first_text(r).startswith("*Travel times are from"))
    hosp_rows = t.rows[hospital_col_header + 1:note_idx]
    hospitals = data.get("hospitals") or [{}]
    for i, h in enumerate(hospitals):
        block = hosp_rows[i * 3:(i + 1) * 3]
        if len(block) < 3:
            break
        for j, row in enumerate(block):
            cells = _unique_cells(row)
            _replace_combined_name_address(cells[0], h.get("name", ""), h.get("address", ""))
            # cells: name/address, sublabel, coordinate, air, ground, phone, box yes, box no, care
            coord = h.get("latitudeDDM", "") if j == 0 else h.get("longitudeDDM", "") if j == 1 else h.get("vhf", "")
            _replace_cell(cells[2], coord)
            _replace_travel_cell(cells[3], _final_air(h))
            _replace_travel_cell(cells[4], _final_ground(h))
            _replace_phone_numbers(cells[5], h.get("phone", ""), h.get("phone2", ""))
            _replace_cell(cells[-1], h.get("levelOfCare", ""))

    # Travel note: origin appears only in the note, not in the Travel Time title box.
    origin = (incident.get("transportTimeOrigin") or {}).get("name", "") or "ICP"
    note_idx = _find_row(t, lambda r: _row_first_text(r).startswith("*Travel times are from"))
    _replace_cell(_unique_cells(t.rows[note_idx])[0], f"*Travel times are from {origin} to the facility and DO NOT include response time to the incident.")
    # The reference document carries keep-with-next on this note. Pages then
    # binds the note to the following Branch table and can create a nearly
    # empty page. The note is the final row of the ambulance/hospital table and
    # must paginate independently.
    _set_keep_next(t.rows[note_idx], False)

    # All full-width service headers share one explicit gray and are centered.
    # Setting these values in generated OOXML avoids application-specific style
    # inheritance differences between Word, Pages, and LibreOffice.
    for title in ("Ambulance Services", "Air Ambulance Services", "Hospitals"):
        idx = _find_row(t, lambda r, expected=title: _row_first_text(r) == expected)
        _style_header_row(t.rows[idx], font_size=8)


def _clear_table_rows(table) -> None:
    for row in list(table.rows):
        table._tbl.remove(row._tr)


def _build_division_rows(doc: Document, ref: Document, data: dict[str, Any]) -> None:
    # IMPORTANT: from Branch/Division down, use the current user-approved
    # formatting reference as the structural source. Preserve its merged
    # cells, widths, alignment, and row properties rather than recreating them.
    old_t = doc.tables[1]
    ref_t = ref.tables[1]

    clone = deepcopy(ref_t._tbl)
    parent = old_t._element.getparent()
    parent.replace(old_t._element, clone)
    t = Table(clone, doc._body)

    # Use the approved Division block itself as the row template.
    source_rows = [deepcopy(r._tr) for r in ref_t.rows[:4]]
    branch_row = deepcopy(ref_t.rows[28]._tr)
    division_header_row = deepcopy(ref_t.rows[29]._tr)
    _clear_table_rows(t)

    branches_used = bool(data.get("branchesUsed"))
    groups = data.get("branches", []) if branches_used else [{"branchName": "", "divisions": data.get("divisions", [])}]

    # Preserve the original validated Word structure: the first Branch heading
    # remains at the bottom of Table 0 and the Division column header moves to
    # Table 1 with the first Division block.
    t0 = doc.tables[0]
    branch_row_idx = _find_row(t0, lambda r: _row_first_text(r).startswith("Branch"))
    division_header_idx = branch_row_idx + 1

    if branches_used and groups:
        _replace_cell(_unique_cells(t0.rows[branch_row_idx])[0], groups[0].get("branchName", ""))
        _style_header_row(t0.rows[branch_row_idx], font_size=11)
        t0._tbl.remove(t0.rows[division_header_idx]._tr)
    else:
        t0._tbl.remove(t0.rows[division_header_idx]._tr)
        t0._tbl.remove(t0.rows[branch_row_idx]._tr)

    for gi, group in enumerate(groups):
        divisions = group.get("divisions", [])

        if branches_used and gi > 0:
            t._tbl.append(deepcopy(branch_row))
            generated_branch = t.rows[-1]
            _replace_cell(_unique_cells(generated_branch)[0], group.get("branchName", ""))
            _style_header_row(generated_branch, font_size=11)
            _add_cant_split(generated_branch)
            _set_keep_next(generated_branch, True)

        for di, div in enumerate(divisions):
            header_start = None
            if di == 0:
                header_start = len(t.rows)
                t._tbl.append(deepcopy(division_header_row))
                _style_header_row(t.rows[-1], font_size=7)

            start = len(t.rows)
            for x in source_rows:
                t._tbl.append(deepcopy(x))
            # EMS responders stay combined in one approved responder row.
            # The legacy blank spacer row below EMS Responders & Capability
            # is removed completely. 4 pt paragraph spacing provides the
            # breathing room before/between/after responders instead.
            responders = div.get("responders") or []
            row0 = t.rows[start]
            row0_cells = _unique_cells(row0)
            _replace_cell(row0_cells[0], div.get("name", ""))
            _style_large_identifier(row0_cells[0])
            _replace_responder_list(row0_cells[-1], responders)

            # Remove the template spacer row (original row 2 of the 4-row
            # Division block). Equipment and Medical Emergency Channel then
            # follow the EMS responder row directly.
            t._tbl.remove(t.rows[start + 1]._tr)
            count = 3
            rows = t.rows[start:start + count]

            _replace_cell(_unique_cells(rows[-2])[-1], div.get("equipment", ""))
            _replace_cell(_unique_cells(rows[-1])[-1], div.get("channel", ""))

            # Create the optional Helispots row only when the Division has
            # explicitly enabled Helispots. The new row clones the approved
            # Medical Emergency Channel row so the borders, widths, and merged
            # cell structure stay consistent with the current ICS-206 format.
            helispots_enabled = str(div.get("helispotsUsed", "")).strip().lower() == "yes"
            helispots_value = str(div.get("helispots", "") or "").strip()
            if helispots_enabled:
                helispot_xml = deepcopy(rows[-1]._tr)
                rows[-1]._tr.addnext(helispot_xml)
                helispot_row = t.rows[start + count]
                hcells = _unique_cells(helispot_row)
                if len(hcells) >= 2:
                    _replace_cell(hcells[1], "Helispots:")
                    _style_detail_label(hcells[1])
                _replace_cell(hcells[-1], helispots_value)
                count += 1
                rows = t.rows[start:start + count]

            # Optional Division response-time pair. When enabled, BOTH rows are
            # always printed together, after Helispots (if present), in this order:
            # Ground Response Time To, then Air Response Time To.
            response_times_enabled = str(div.get("responseTimesUsed", "")).strip().lower() == "yes"
            if response_times_enabled:
                for label, response_value in (
                    ("Ground Response Time To:", str(div.get("groundResponseTimeTo", "") or "").strip()),
                    ("Air Response Time To:", str(div.get("airResponseTimeTo", "") or "").strip()),
                ):
                    response_xml = deepcopy(rows[-1]._tr)
                    rows[-1]._tr.addnext(response_xml)
                    response_row = t.rows[start + count]
                    rcells = _unique_cells(response_row)
                    if len(rcells) >= 2:
                        _replace_cell(rcells[1], label)
                        _style_detail_label(rcells[1])
                    _replace_cell(rcells[-1], response_value)
                    count += 1
                    rows = t.rows[start:start + count]

            # Approved Branch/Division formatting:
            # identifier centered 14 pt bold; all detail labels 8 pt bold,
            # left justified and bottom aligned; detail values left justified.
            for rr in rows:
                cells = _unique_cells(rr)
                if len(cells) >= 2:
                    _style_detail_label(cells[1])
                if cells:
                    for p in cells[-1].paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # The EMS Responders & Capability title remains merged, left
            # justified, and aligned to the TOP of its cell.
            ems_cells = _unique_cells(rows[0])
            if len(ems_cells) >= 2:
                ems_cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                for p in ems_cells[1].paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    # v0.10.124: add 4 pt of breathing room above the
                    # EMS Responders & Capability title text while keeping
                    # the merged title cell top-aligned and left justified.
                    p.paragraph_format.space_before = Pt(4)

            # Retain the approved multi-row Word block and keep all of its rows
            # together. The first Division also carries its column header.
            if header_start is not None:
                _protect_block(t.rows[header_start:start + count])
            else:
                _protect_block(rows)


def _remove_table(doc: Document, table) -> None:
    table._element.getparent().remove(table._element)


def _build_locations(doc: Document, data: dict[str, Any]) -> None:
    # Table 2 is Location header, Table 3 is a blank location + Prepared/Reviewed footer.
    header = doc.tables[2]
    template = doc.tables[3]
    template_xml = deepcopy(template._tbl)
    _remove_table(doc, template)

    locations = data.get("locations") or [{}]
    marker = header._tbl
    # addnext reverses order when repeatedly used against same marker; advance marker each time.
    for _ in locations:
        clone = deepcopy(template_xml)
        marker.addnext(clone)
        marker = clone


def _populate_locations(doc: Document, data: dict[str, Any]) -> None:
    locations = data.get("locations") or [{}]
    blocks = doc.tables[3:3 + len(locations)]

    # Keep Location header with first block.
    _style_header_row(doc.tables[2].rows[0], font_size=8)
    _protect_block(doc.tables[2].rows, include_last_keep=True)

    for i, (tbl, loc) in enumerate(zip(blocks, locations)):
        last = i == len(locations) - 1
        # Non-final location blocks must not carry Prepared/Reviewed footer.
        if not last:
            for _ in range(2):
                tbl._tbl.remove(tbl.rows[-1]._tr)

        contacts = loc.get("pointsOfContact")
        if not isinstance(contacts, list):
            single = loc.get("pointOfContact", "")
            contacts = [x.strip() for x in str(single).split("|") if x.strip()] if single else []
        responders = loc.get("responders")
        if not isinstance(responders, list):
            single = loc.get("respondersCapability", "")
            responders = [x.strip() for x in str(single).split("|") if x.strip()] if single else []

        # Base rows: 0 contact, 1 spacer, 2 responder, 3 equip, 4 channel, [5-6 footer]
        spacer_xml = deepcopy(tbl.rows[1]._tr)
        # Extra contacts go before the original spacer.
        for extra in reversed(contacts[1:]):
            tbl.rows[1]._tr.addprevious(deepcopy(spacer_xml))
            _replace_cell(_unique_cells(tbl.rows[1])[-1], extra)

        # Locate EMS/equipment/channel dynamically after contact expansion.
        ems_idx = _find_row(tbl, lambda r: any(_cell_text(c).startswith("EMS Responders") for c in _unique_cells(r)))
        equip_idx = _find_row(tbl, lambda r: any(_cell_text(c).startswith("Equipment Available") for c in _unique_cells(r)))
        chan_idx = _find_row(tbl, lambda r: any(_cell_text(c).startswith("Medical Emergency") for c in _unique_cells(r)))

        # Extra responder continuation rows before equipment.
        for extra in responders[1:]:
            marker = tbl.rows[equip_idx]
            marker._tr.addprevious(deepcopy(spacer_xml))
            _replace_cell(_unique_cells(tbl.rows[equip_idx])[-1], extra)
            equip_idx += 1
            chan_idx += 1

        # Populate primary values.
        _replace_cell(_unique_cells(tbl.rows[0])[0], loc.get("location", ""))
        _style_large_identifier(_unique_cells(tbl.rows[0])[0])
        _replace_cell(_unique_cells(tbl.rows[0])[-1], contacts[0] if contacts else "")
        ems_idx = _find_row(tbl, lambda r: any(_cell_text(c).startswith("EMS Responders") for c in _unique_cells(r)))
        _replace_cell(_unique_cells(tbl.rows[ems_idx])[-1], responders[0] if responders else "")
        equip_idx = _find_row(tbl, lambda r: any(_cell_text(c).startswith("Equipment Available") for c in _unique_cells(r)))
        chan_idx = _find_row(tbl, lambda r: any(_cell_text(c).startswith("Medical Emergency") for c in _unique_cells(r)))
        _replace_cell(_unique_cells(tbl.rows[equip_idx])[-1], loc.get("equipment", ""))
        _replace_cell(_unique_cells(tbl.rows[chan_idx])[-1], loc.get("medicalEmergencyChannel", ""))

        # Approved Location formatting mirrors the validated reference:
        # location name centered 14 pt bold; labels 8 pt bold/left/bottom;
        # entered values left justified.
        footer_start = len(tbl.rows) - 2 if last else len(tbl.rows)
        for rr in tbl.rows[:footer_start]:
            cells = _unique_cells(rr)
            if len(cells) >= 2 and _cell_text(cells[1]):
                _style_detail_label(cells[1])
            if cells:
                for p in cells[-1].paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Prepared By is a header, not an input. Keep the data row blank.
        if last:
            _style_header_row(tbl.rows[-2], font_size=8)
            for c in _unique_cells(tbl.rows[-1]):
                _replace_cell(c, "")

        # Individual location block cannot split; the final block also carries the footer.
        _protect_block(tbl.rows)


def generate_ics206(
    incident: dict[str, Any],
    template_path: str | Path,
    output_path: str | Path,
    component_reference_path: str | Path | None = None,
) -> str:
    """Generate a validated IMPS ICS-206 DOCX from one incident dictionary."""
    template_path = Path(template_path)
    output_path = Path(output_path)
    if component_reference_path is None:
        component_reference_path = template_path.parent / "ICS206WF_IMPS_Component_Reference.docx"
    ref_path = Path(component_reference_path)

    doc = Document(template_path)
    ref = Document(ref_path)

    _build_table0(doc, ref, incident)
    # Reopen to refresh python-docx row collections after OOXML row insertions.
    tmp1 = output_path.with_suffix(".structure.tmp.docx")
    doc.save(tmp1)
    doc = Document(tmp1)

    _populate_table0_rows(doc, incident)
    _build_division_rows(doc, ref, incident)
    _build_locations(doc, incident)

    tmp2 = output_path.with_suffix(".locations.tmp.docx")
    doc.save(tmp2)
    doc = Document(tmp2)

    _populate_locations(doc, incident)
    _set_small_hospital_labels(doc)
    _set_arial_everywhere(doc)

    _apply_optional_logo(doc, incident)
    _apply_signatures(doc, incident)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)

    for p in (tmp1, tmp2):
        try:
            p.unlink()
        except FileNotFoundError:
            pass
    return str(output_path)
