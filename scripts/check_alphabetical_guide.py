import re
import sys
import zipfile


DOCX = sys.argv[1] if len(sys.argv) > 1 else "grow_a_garden_ultimate_pet_guide_alphabetical.docx"

with zipfile.ZipFile(DOCX) as zf:
    xml = zf.read("word/document.xml").decode("utf-8")
    rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")

bookmarks = set(re.findall(r'<w:bookmarkStart\b[^>]*\bw:name="([^"]+)"', xml))
anchors = re.findall(r'<w:hyperlink\b[^>]*\bw:anchor="([^"]+)"', xml)
missing = sorted(set(anchors) - bookmarks)

body_start = xml.find('w:name="pets_a_z"')
index_start = xml.find('w:name="pet_index_by_passive_ability_category"')
tier_start = xml.find('w:name="pet_index_by_tier"')
body_xml = xml[body_start:index_start]
heading_names = re.findall(
    r'<w:p><w:pPr>[\s\S]*?<w:pStyle w:val="Heading3" />[\s\S]*?</w:pPr><w:r><w:t xml:space="preserve">([^<]+)</w:t></w:r></w:p>',
    body_xml,
)
sorted_names = sorted(heading_names, key=lambda s: re.sub(r"[^a-z0-9]+", " ", s.lower()).strip())
top_nav = xml[:body_start]
tier_break_window = xml[max(0, tier_start - 250):tier_start]
source_pos = xml.find("Source:")

print(f"zip_body_start={body_start}")
print(f"pet_count={len(heading_names)}")
print(f"alphabetical_body={heading_names == sorted_names}")
print(f"first_five={heading_names[:5]}")
print(f"last_five={heading_names[-5:]}")
print(f"section_count={len(re.findall(r'<w:sectPr', xml))}")
print(f"two_column_section={'<w:cols w:num=\"2\"' in xml}")
print(f"internal_link_count={len(anchors)}")
print(f"missing_anchor_targets={len(missing)}")
print(f"top_nav_centered_paragraphs={top_nav.count('<w:jc w:val=\"center\" />')}")
print(f"top_nav_internal_links={len(re.findall(r'<w:hyperlink[^>]+w:anchor=', top_nav))}")
print(f"top_nav_has_passive_label={'By Passive Ability:' in top_nav}")
print(f"top_nav_has_rarity_label={'By Rarity:' in top_nav}")
print(f"tier_index_page_break={'<w:br w:type=\"page\"' in tier_break_window}")
print(f"source_after_indexes={source_pos > index_start}")
print(f"source_before_body={'Source:' in xml[:body_start]}")
print(f"source_external_hyperlink={'growagarden.fandom.com/wiki/Grow_a_Garden_Wiki' in rels}")
print(f"floating_image_count={xml.count('<wp:anchor')}")
print(f"inline_image_count={xml.count('<wp:inline')}")
print(f"image_caption_count={len(re.findall(r'w:val=\"ImageCaption\"', xml))}")
print(f"clear_before_pet_heading_count={body_xml.count('w:clear=\"all\"')}")
print(f"cant_split_row_count={body_xml.count('<w:cantSplit')}")
print(f"pet_entry_table_count={body_xml.count('<w:tbl><w:tblPr><w:tblW w:w=\"9360\"')}")
print(f"centered_image_cell_count={body_xml.count('<w:vAlign w:val=\"center\" />')}")
