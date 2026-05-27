import re
import sys
import zipfile


DOCX = sys.argv[1] if len(sys.argv) > 1 else "grow_a_garden_pet_guide_two_column_indexes.docx"

with zipfile.ZipFile(DOCX) as zf:
    xml = zf.read("word/document.xml").decode("utf-8")

tiers = ["Common", "Uncommon", "Rare", "Legendary", "Mythical", "Divine", "Prismatic"]
tier_start = xml.find('w:name="pet_index_by_tier"')
positions = [(tier, xml.find(f">{tier}</w:t>", tier_start)) for tier in tiers]
bookmarks = set(re.findall(r'<w:bookmarkStart\b[^>]*\bw:name="([^"]+)"', xml))
anchors = re.findall(r'<w:hyperlink\b[^>]*\bw:anchor="([^"]+)"', xml)
missing = sorted(set(anchors) - bookmarks)
tier_break_window = xml[max(0, tier_start - 250):tier_start]
top_nav = xml[:xml.find('w:name="cosmetics"')]

print(f"section_count={len(re.findall(r'<w:sectPr', xml))}")
print(f"two_column_section={'<w:cols w:num=\"2\"' in xml}")
print(f"internal_link_count={len(anchors)}")
print(f"missing_anchor_targets={len(missing)}")
print(f"top_nav_centered_paragraphs={top_nav.count('<w:jc w:val=\"center\" />')}")
print(f"top_nav_internal_links={len(re.findall(r'<w:hyperlink[^>]+w:anchor=', top_nav))}")
print(f"tier_index_page_break={'<w:br w:type=\"page\"' in tier_break_window}")
print(f"tier_heading_positions={positions}")
print(f"tier_order_ok={all(positions[i][1] < positions[i + 1][1] for i in range(len(positions) - 1))}")
print(f"floating_image_count={xml.count('<wp:anchor')}")
print(f"inline_image_count={xml.count('<wp:inline')}")
print(f"image_caption_count={len(re.findall(r'w:val=\"ImageCaption\"', xml))}")
