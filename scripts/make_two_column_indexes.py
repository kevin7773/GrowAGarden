import os
import re
import sys
import zipfile
from hashlib import sha1


DOCX = sys.argv[1] if len(sys.argv) > 1 else "grow_a_garden_pet_guide_two_column_indexes.docx"
INDEX_BOOKMARK = 'w:name="pet-index-by-passive-ability-category"'
TIER_INDEX_BOOKMARK = 'w:name="pet-index-by-tier"'


def main():
    tmp = DOCX + ".tmp"
    with zipfile.ZipFile(DOCX, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    xml = files["word/document.xml"].decode("utf-8")
    xml = make_bookmarks_word_safe(xml)
    xml = center_top_navigation(xml)
    if 'w:name="pets_a_z"' in xml:
        xml = table_pet_entries(xml)
        xml = close_bookmark_after_next_paragraph(xml, "pets_a_z")
    else:
        xml = float_pet_images(xml)
    section_matches = list(re.finditer(r"<w:sectPr[\s\S]*?</w:sectPr>", xml))
    if len(section_matches) != 1:
        raise RuntimeError(f"Expected one document section, found {len(section_matches)}")

    first_index_match = find_bookmark_start(xml, INDEX_BOOKMARK)
    if not first_index_match:
        raise RuntimeError("Could not find the passive index bookmark")
    bookmark_start = first_index_match.start()

    original_section = ensure_page_geometry(section_matches[-1].group(0))
    first_section = original_section.replace(
        "<w:sectPr>", '<w:sectPr><w:type w:val="nextPage" />', 1
    )
    page_break = '<w:p><w:r><w:br w:type="page" /></w:r></w:p>\n    '
    section_break = f"{page_break}<w:p><w:pPr>{first_section}</w:pPr></w:p>\n    "
    xml = xml[:bookmark_start] + section_break + xml[bookmark_start:]

    tier_index_match = find_bookmark_start(xml, TIER_INDEX_BOOKMARK)
    if not tier_index_match:
        raise RuntimeError("Could not find the tier index bookmark")
    tier_bookmark_start = tier_index_match.start()
    page_break = (
        '<w:p><w:r><w:br w:type="page" /></w:r></w:p>\n    '
    )
    xml = xml[:tier_bookmark_start] + page_break + xml[tier_bookmark_start:]

    section_matches = list(re.finditer(r"<w:sectPr[\s\S]*?</w:sectPr>", xml))
    final_match = section_matches[-1]
    final_section = ensure_page_geometry(final_match.group(0))
    columns = '<w:cols w:num="2" w:space="360" />'
    if re.search(r"<w:cols\b[^>]*/>", final_section):
        final_section = re.sub(r"<w:cols\b[^>]*/>", columns, final_section, count=1)
    else:
        final_section = final_section.replace("</w:sectPr>", columns + "</w:sectPr>", 1)

    xml = xml[: final_match.start()] + final_section + xml[final_match.end() :]
    files["word/document.xml"] = xml.encode("utf-8")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    os.replace(tmp, DOCX)
    print("Patched index section to two columns")


def float_pet_images(xml):
    index_match = find_bookmark_start(xml, INDEX_BOOKMARK)
    index_pos = index_match.start() if index_match else len(xml)
    pets_match = find_bookmark_start(xml, 'w:name="pets-a-z"')
    pets_pos = pets_match.start() if pets_match else 0

    prefix = xml[:pets_pos]
    body = xml[pets_pos:index_pos]
    indexes = xml[index_pos:]

    body = clear_wrapping_before_pet_headings(body)
    body = re.sub(
        r"<w:p><w:pPr><w:pStyle w:val=\"ImageCaption\" /></w:pPr>[\s\S]*?</w:p>\s*",
        "",
        body,
    )

    def convert_inline(match):
        inner = match.group(1)
        extent = extract_tag(inner, "wp:extent")
        effect_extent = extract_tag(inner, "wp:effectExtent")
        doc_pr = extract_tag(inner, "wp:docPr")
        graphic = extract_block(inner, "a:graphic")
        if not (extent and doc_pr and graphic):
            return match.group(0)
        if not effect_extent:
            effect_extent = '<wp:effectExtent b="0" l="0" r="0" t="0" />'
        return (
            '<wp:anchor distT="0" distB="0" distL="114300" distR="228600" '
            'simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" '
            'layoutInCell="1" allowOverlap="0">'
            '<wp:simplePos x="0" y="0" />'
            '<wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f"{extent}{effect_extent}"
            '<wp:wrapSquare wrapText="right" />'
            f"{doc_pr}{graphic}</wp:anchor>"
        )

    body = re.sub(r"<wp:inline>([\s\S]*?)</wp:inline>", convert_inline, body)
    return prefix + body + indexes


def table_pet_entries(xml):
    index_match = find_bookmark_start(xml, INDEX_BOOKMARK)
    pets_match = find_bookmark_start(xml, 'w:name="pets-a-z"')
    if not index_match or not pets_match:
        return xml
    index_pos = index_match.start()
    pets_pos = pets_match.start()

    prefix = xml[:pets_pos]
    body = xml[pets_pos:index_pos]
    indexes = xml[index_pos:]
    start_pattern = re.compile(
        r'<w:bookmarkStart\b[^>]*?/>\s*<w:p><w:pPr><w:pStyle w:val="Heading3" />',
        re.S,
    )
    starts = [match.start() for match in start_pattern.finditer(body)]
    if not starts:
        return xml

    body_prefix = body[: starts[0]]
    rebuilt = [body_prefix]
    for idx, start in enumerate(starts):
        end = starts[idx + 1] if idx + 1 < len(starts) else len(body)
        rebuilt.append(pet_entry_table(body[start:end]))
    return prefix + "".join(rebuilt) + indexes


def pet_entry_table(entry):
    start_match = re.match(r'\s*(<w:bookmarkStart\b[^>]*?/>)\s*', entry, re.S)
    if not start_match:
        return entry
    bookmark_start = start_match.group(1)
    remainder = entry[start_match.end():]
    heading_match = re.match(r'\s*(<w:p><w:pPr><w:pStyle w:val="Heading3" />[\s\S]*?</w:p>)\s*', remainder, re.S)
    if not heading_match:
        return entry
    heading = force_left_justified(add_keep_next(heading_match.group(1)))
    remainder = remainder[heading_match.end():]

    paragraphs = re.findall(r"<w:p>[\s\S]*?</w:p>", remainder)
    rest = re.sub(r"<w:p>[\s\S]*?</w:p>", "", remainder)
    image = ""
    details = []
    bookmark_ends = re.findall(r"<w:bookmarkEnd\b[^>]*/>", rest)
    for paragraph in paragraphs:
        if 'w:val="ImageCaption"' in paragraph:
            continue
        ends = re.findall(r"<w:bookmarkEnd\b[^>]*/>", paragraph)
        if ends:
            bookmark_ends.extend(ends)
            paragraph = re.sub(r"<w:bookmarkEnd\b[^>]*/>", "", paragraph)
        if not image and "<w:drawing>" in paragraph:
            image = paragraph
        else:
            details.append(paragraph)
    bookmark_end_xml = "".join(bookmark_ends)
    details_xml = "".join(details) or "<w:p/>"
    if not image:
        return single_cell_pet_entry(bookmark_start, heading, details_xml, bookmark_end_xml)

    inner = (
        '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa" />'
        '<w:tblLayout w:type="fixed" />'
        '<w:tblBorders><w:top w:val="nil" /><w:left w:val="nil" />'
        '<w:bottom w:val="nil" /><w:right w:val="nil" />'
        '<w:insideH w:val="nil" /><w:insideV w:val="nil" /></w:tblBorders>'
        '</w:tblPr><w:tblGrid><w:gridCol w:w="1200" /><w:gridCol w:w="7800" /></w:tblGrid>'
        '<w:tr><w:trPr><w:cantSplit /></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa" /><w:vAlign w:val="center" />{cell_margins()}</w:tcPr>{image}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="7800" w:type="dxa" />{cell_margins()}</w:tcPr>{details_xml}</w:tc>'
        '</w:tr></w:tbl>'
    )
    outer = (
        '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa" />'
        '<w:tblLayout w:type="fixed" />'
        '<w:tblBorders><w:top w:val="nil" /><w:left w:val="nil" />'
        '<w:bottom w:val="nil" /><w:right w:val="nil" />'
        '<w:insideH w:val="nil" /><w:insideV w:val="nil" /></w:tblBorders>'
        '</w:tblPr><w:tblGrid><w:gridCol w:w="9360" /></w:tblGrid>'
        '<w:tr><w:trPr><w:cantSplit /></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa" />{cell_margins()}</w:tcPr>'
        f'{bookmark_start}{heading}{inner}{bookmark_end_xml}<w:p/></w:tc></w:tr></w:tbl>'
    )
    return outer


def single_cell_pet_entry(bookmark_start, heading, details_xml, bookmark_end_xml):
    outer = (
        '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa" />'
        '<w:tblLayout w:type="fixed" />'
        '<w:tblBorders><w:top w:val="nil" /><w:left w:val="nil" />'
        '<w:bottom w:val="nil" /><w:right w:val="nil" />'
        '<w:insideH w:val="nil" /><w:insideV w:val="nil" /></w:tblBorders>'
        '</w:tblPr><w:tblGrid><w:gridCol w:w="9360" /></w:tblGrid>'
        '<w:tr><w:trPr><w:cantSplit /></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa" />{cell_margins()}</w:tcPr>'
        f'{bookmark_start}{heading}{details_xml}{bookmark_end_xml}<w:p/></w:tc></w:tr></w:tbl>'
    )
    return outer


def add_keep_next(paragraph):
    if "<w:keepNext" in paragraph:
        return paragraph
    return paragraph.replace("<w:pPr>", "<w:pPr><w:keepNext />", 1)


def force_left_justified(paragraph):
    if "<w:jc " in paragraph:
        return re.sub(r"<w:jc\b[^>]*/>", '<w:jc w:val="left" />', paragraph, count=1)
    return paragraph.replace("</w:pPr>", '<w:jc w:val="left" /></w:pPr>', 1)


def cell_margins():
    return (
        '<w:tcMar><w:top w:w="60" w:type="dxa" /><w:left w:w="60" w:type="dxa" />'
        '<w:bottom w:w="60" w:type="dxa" /><w:right w:w="120" w:type="dxa" /></w:tcMar>'
    )


def clear_wrapping_before_pet_headings(body):
    clear_break = '<w:p><w:r><w:br w:type="textWrapping" w:clear="all" /></w:r></w:p>'
    pattern = r'(<w:bookmarkStart\b[^>]*?/>\s*<w:p><w:pPr><w:pStyle w:val="Heading3" />)'
    return re.sub(pattern, clear_break + r"\1", body)


def extract_tag(text, tag):
    match = re.search(rf"<{re.escape(tag)}\b[^>]*/>", text)
    return match.group(0) if match else ""


def extract_block(text, tag):
    match = re.search(rf"<{re.escape(tag)}\b[\s\S]*?</{re.escape(tag)}>", text)
    return match.group(0) if match else ""


def make_bookmarks_word_safe(xml):
    names = re.findall(r'<w:bookmarkStart\b[^>]*\bw:name="([^"]+)"', xml)
    mapping = {}
    used = set()
    for name in names:
        if name not in mapping:
            safe = sanitize_bookmark(name)
            base = safe
            n = 2
            while safe in used:
                suffix = f"_{n}"
                safe = base[: 40 - len(suffix)] + suffix
                n += 1
            mapping[name] = safe
            used.add(safe)

    def replace_bookmark(match):
        return f'{match.group(1)}{mapping[match.group(2)]}"'

    xml = re.sub(r'(<w:bookmarkStart\b[^>]*\bw:name=")([^"]+)"', replace_bookmark, xml)

    def replace_anchor(match):
        anchor = match.group(2)
        return f'{match.group(1)}{mapping.get(anchor, sanitize_bookmark(anchor))}"'

    xml = re.sub(r'(<w:hyperlink\b[^>]*\bw:anchor=")([^"]+)"', replace_anchor, xml)
    return xml


def center_top_navigation(xml):
    first_section = xml.find('w:name="pets_a_z"')
    if first_section < 0:
        first_section = xml.find('w:name="cosmetics"')
    if first_section < 0:
        return xml

    prefix = xml[:first_section]
    suffix = xml[first_section:]

    def center_nav_paragraph(match):
        para = match.group(0)
        is_label = "By Passive Ability:" in para or "By Rarity:" in para
        if "<w:hyperlink" not in para and not is_label:
            return para
        if "<w:jc " in para:
            centered = para
        else:
            centered = para.replace("</w:pPr>", '<w:jc w:val="center" /></w:pPr>', 1)
        if is_label:
            centered = re.sub(r"<w:rPr>([\s\S]*?)</w:rPr>", ensure_bold_run_props, centered)
            if "<w:rPr>" not in centered:
                centered = centered.replace(
                    "<w:r>",
                    '<w:r><w:rPr><w:b /><w:bCs /><w:sz w:val="28" /><w:szCs w:val="28" /></w:rPr>',
                    1,
                )
        return centered

    prefix = re.sub(r"<w:p><w:pPr>[\s\S]*?</w:p>", center_nav_paragraph, prefix)
    return prefix + suffix


def ensure_bold_run_props(match):
    props = match.group(1)
    if "<w:b" not in props:
        props = "<w:b />" + props
    if "<w:bCs" not in props:
        props = "<w:bCs />" + props
    if "<w:sz" not in props:
        props += '<w:sz w:val="28" /><w:szCs w:val="28" />'
    return f"<w:rPr>{props}</w:rPr>"


def sanitize_bookmark(name):
    safe = re.sub(r"[^A-Za-z0-9_]", "_", name)
    safe = re.sub(r"_+", "_", safe).strip("_")
    if not safe or not safe[0].isalpha():
        safe = "b_" + safe
    if len(safe) > 40:
        digest = sha1(name.encode("utf-8")).hexdigest()[:8]
        safe = safe[:31].rstrip("_") + "_" + digest
    return safe


def safe_bookmark_attr(attr):
    name = re.search(r'w:name="([^"]+)"', attr).group(1)
    return f'w:name="{sanitize_bookmark(name)}"'


def find_bookmark_start(xml, attr):
    name = re.search(r'w:name="([^"]+)"', attr).group(1)
    safe_name = sanitize_bookmark(name)
    return re.search(
        rf'<w:bookmarkStart\b[^>]*\bw:name="{re.escape(safe_name)}"[^>]*/>',
        xml,
    )


def close_bookmark_after_next_paragraph(xml, bookmark_name):
    safe_name = sanitize_bookmark(bookmark_name)
    start_match = re.search(
        rf'<w:bookmarkStart\b[^>]*\bw:id="([^"]+)"[^>]*\bw:name="{re.escape(safe_name)}"[^>]*/>',
        xml,
    )
    if not start_match:
        return xml
    bookmark_id = start_match.group(1)
    end_pattern = rf'<w:bookmarkEnd\b[^>]*\bw:id="{re.escape(bookmark_id)}"[^>]*/>\s*'
    end_match = re.search(end_pattern, xml)
    if not end_match:
        return xml
    end_tag = end_match.group(0).strip()
    xml = re.sub(end_pattern, "", xml, count=1)

    paragraph_end = xml.find("</w:p>", start_match.end())
    if paragraph_end < 0:
        return xml + end_tag
    insert_at = paragraph_end + len("</w:p>")
    return xml[:insert_at] + end_tag + xml[insert_at:]


def ensure_page_geometry(section):
    if "<w:pgSz" not in section:
        section = section.replace(
            "<w:sectPr>",
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840" />'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
            'w:left="1440" w:header="720" w:footer="720" w:gutter="0" />',
            1,
        )
    return section


if __name__ == "__main__":
    main()
