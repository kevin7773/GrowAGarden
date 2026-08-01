import re
import sys
import zipfile
import xml.etree.ElementTree as ET


DOCX = sys.argv[1] if len(sys.argv) > 1 else "grow_a_garden_ultimate_pet_guide_alphabetical.docx"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def main():
    with zipfile.ZipFile(DOCX) as zf:
        xml_bytes = zf.read("word/document.xml")
        xml = xml_bytes.decode("utf-8")

    root = ET.fromstring(xml_bytes)
    parents = {child: parent for parent in root.iter() for child in parent}

    body = xml[: xml.find("Pet Index")]
    names = re.findall(
        r'<w:pStyle w:val="Heading3" />.*?<w:t[^>]*>([^<]+)</w:t>',
        body,
        re.S,
    )
    missing_wrappers = []
    for name in names:
        i = body.find(f">{name}<")
        table_start = body.rfind("<w:tbl", 0, i)
        table_end = body.find("</w:tbl>", i)
        next_heading = body.find('<w:pStyle w:val="Heading3"', i + 1)
        block = body[table_start : table_end + 8 if table_end != -1 else len(body)] if table_start != -1 else ""
        if table_start == -1 or (next_heading != -1 and next_heading < table_end) or "<w:cantSplit" not in block:
            missing_wrappers.append(name)

    starts = {}
    ends = {}
    for el in root.iter(f"{{{NS['w']}}}bookmarkStart"):
        starts[el.attrib[f"{{{NS['w']}}}id"]] = el
    for el in root.iter(f"{{{NS['w']}}}bookmarkEnd"):
        ends[el.attrib[f"{{{NS['w']}}}id"]] = el

    bookmark_mismatches = []
    for bookmark_id, start in starts.items():
        end = ends.get(bookmark_id)
        bookmark_name = start.attrib.get(f"{{{NS['w']}}}name", "")
        if end is None:
            bookmark_mismatches.append((bookmark_id, bookmark_name, "missing_end"))
            continue
        start_tc = nearest(start, parents, f"{{{NS['w']}}}tc")
        end_tc = nearest(end, parents, f"{{{NS['w']}}}tc")
        if (start_tc is None) != (end_tc is None):
            bookmark_mismatches.append((bookmark_id, bookmark_name, "crosses_table_cell_boundary"))
        elif start_tc is not None and start_tc is not end_tc:
            bookmark_mismatches.append((bookmark_id, bookmark_name, "different_table_cells"))

    key_images = ["Cicada.png", "Fire_Wisp.png", "Newt.png", "Nightjar.png", "Water_Buffalo.png"]

    print(f"pet_heading_count={len(names)}")
    print(f"alphabetical={names == sorted(names, key=sort_key)}")
    print(f"entries_missing_own_table_or_cantsplit={len(missing_wrappers)}")
    print(f"missing_wrappers={missing_wrappers[:20]}")
    print(f"bookmark_table_boundary_mismatches={len(bookmark_mismatches)}")
    print(f"bookmark_mismatch_sample={bookmark_mismatches[:10]}")
    print(f"new_images_in_docx={all(image in xml for image in key_images)}")
    print(f"wiki_page_count={xml.count('Wiki Page:')}")


def nearest(el, parents, tag):
    current = el
    while current in parents:
        current = parents[current]
        if current.tag == tag:
            return current
    return None


def sort_key(value):
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


if __name__ == "__main__":
    main()
