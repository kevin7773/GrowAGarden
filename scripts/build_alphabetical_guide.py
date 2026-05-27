import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(".")
SOURCE_MD = ROOT / "grow_a_garden_pet_guide_by_passive.md"
OUT_MD = ROOT / "grow_a_garden_pet_guide_alphabetical.md"
BANNER_SOURCE = ROOT / "assets" / "grow-a-garden-pets-tier-list.jpg"
BANNER_OUT = ROOT / "assets" / "ultimate-pet-guide-banner.png"


PASSIVE_INDEX_TITLE = "## Pet Index (by Passive Ability Category)"


def main():
    make_banner()
    text = SOURCE_MD.read_text(encoding="utf-8")
    body_text, index_text = text.split(PASSIVE_INDEX_TITLE, 1)
    index_text = PASSIVE_INDEX_TITLE + index_text
    pet_blocks = parse_pet_blocks(body_text)

    lines = [
        "# Grow a Garden Ultimate Pet Guide",
        "",
        "![](assets/ultimate-pet-guide-banner.png){width=6.5in}",
        "",
    ]
    lines.extend(top_nav_lines())
    lines.extend(
        [
            "",
            "## Pets A-Z {#pets-a-z}",
            "",
        ]
    )

    for _, block in sorted(pet_blocks, key=lambda item: sort_key(item[0])):
        lines.append(normalize_pet_block(block).strip())
        lines.append("")
        lines.append("")

    lines.append('<div style="page-break-before: always;"></div>')
    lines.append("")
    lines.append(index_text.strip())
    lines.append("")
    lines.append("Source: [https://growagarden.fandom.com/wiki/Grow_a_Garden_Wiki](https://growagarden.fandom.com/wiki/Grow_a_Garden_Wiki)")
    lines.append("")

    OUT_MD.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {OUT_MD}")
    print(f"Wrote {BANNER_OUT}")
    print(f"Pet count: {len(pet_blocks)}")


def parse_pet_blocks(markdown):
    blocks = []
    current_title = None
    current = []
    for line in markdown.splitlines():
        if line.startswith("## "):
            continue
        match = re.match(r"^###\s+(.+?)\s+\{#[^}]+\}\s*$", line)
        if match:
            if current_title and current:
                blocks.append((current_title, "\n".join(current)))
            current_title = match.group(1)
            current = [line]
        elif current_title:
            current.append(line)
    if current_title and current:
        blocks.append((current_title, "\n".join(current)))
    return blocks


def normalize_pet_block(block):
    lines = block.splitlines()
    out = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if line in ("- **Hatch Chance:**", "- **Obtaining Method:**"):
            field = line
            raw_items = []
            i += 1
            while i < len(lines):
                next_line = lines[i]
                if next_line == "":
                    break
                if next_line.startswith("- **"):
                    break
                raw_items.append(next_line)
                i += 1
            items = split_field_items(raw_items)
            out.append(field)
            if len(items) > 1:
                out.extend(f"  - {item}" for item in items)
            elif len(items) == 1:
                if raw_items and raw_items[0].lstrip().startswith("- "):
                    out.append(f"  - {items[0]}")
                else:
                    out.append(f"  {items[0]}")
            if i < len(lines) and lines[i] == "":
                out.append("")
                i += 1
            continue
        out.append(line)
        i += 1
    return "\n".join(out)


def split_field_items(raw_items):
    parts = []
    for raw in raw_items:
        item = raw.strip()
        if item.startswith("- "):
            item = item[2:].strip()
        if not item:
            continue
        item = item.replace("`n", "\n").replace("\\n", "\n")
        item = re.sub(r"(?<=%)(?=[A-Z])", "\n", item)
        item = re.sub(r"(?<=Chest)(?=[A-Z])", "\n", item)
        item = re.sub(r"(?<=Sack)(?=[A-Z])", "\n", item)
        for part in item.splitlines():
            part = part.strip()
            if part:
                parts.append(part)
    return parts


def top_nav_lines():
    passive_rows = [
        ["Cosmetics", "Crafting", "Egg Helpers", "Gear"],
        ["Harvest or Seed Helpers", "Levellers", "Mutators", "Pet Helpers"],
        ["Plant Growth", "Player Interactive", "Special"],
    ]
    tier_rows = [
        ["Common", "Uncommon", "Rare", "Legendary"],
        ["Mythical", "Divine", "Prismatic"],
    ]
    lines = []
    lines.append("<div style='text-align:center; font-size:14pt; font-weight:bold'>By Passive Ability:</div>")
    lines.append("")
    for row in passive_rows:
        lines.append(center_link_row(row, suffix="-index"))
        lines.append("")
    lines.append("<div style='text-align:center; font-size:14pt; font-weight:bold'>By Rarity:</div>")
    lines.append("")
    for row in tier_rows:
        lines.append(center_link_row(row, prefix="tier-"))
        lines.append("")
    return lines


def center_link_row(labels, prefix="", suffix=""):
    links = []
    for label in labels:
        slug = slugify(label)
        links.append(f"[{label}](#{prefix}{slug}{suffix})")
    spacer = "&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;"
    return f"<div style='text-align:center'>{spacer.join(links)}</div>"


def slugify(value):
    slug = re.sub(r"[^a-z0-9\s-]", "", value.lower())
    slug = re.sub(r"\s+", "-", slug)
    return slug.strip("-")


def sort_key(title):
    return re.sub(r"[^a-z0-9]+", " ", title.lower()).strip()


def make_banner():
    BANNER_OUT.parent.mkdir(parents=True, exist_ok=True)
    source = Image.open(BANNER_SOURCE).convert("RGB")
    banner_w, banner_h = 1800, 520
    src_w, src_h = source.size
    crop = source.crop((0, 160, src_w, src_h - 170)).resize((banner_w, banner_h))
    overlay = Image.new("RGBA", crop.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    draw.rectangle((0, 0, banner_w, banner_h), fill=(25, 123, 47, 45))
    draw.rectangle((0, int(banner_h * 0.67), banner_w, banner_h), fill=(41, 111, 25, 55))
    banner = Image.alpha_composite(crop.convert("RGBA"), overlay)

    title = "Grow a Garden\nUltimate Pet Guide"
    font_path = Path("C:/Windows/Fonts/LuckiestGuy-Regular.ttf")
    if not font_path.exists():
        font_path = Path("C:/Windows/Fonts/Baloo-Regular.ttf")
    font = ImageFont.truetype(str(font_path), 112)
    small_font = ImageFont.truetype(str(font_path), 90)

    text_layer = Image.new("RGBA", banner.size, (0, 0, 0, 0))
    text_draw = ImageDraw.Draw(text_layer)
    lines = title.split("\n")
    fonts = [font, small_font]
    line_boxes = [text_draw.textbbox((0, 0), line, font=fonts[i], stroke_width=6) for i, line in enumerate(lines)]
    total_h = sum(box[3] - box[1] for box in line_boxes) + 12
    y = (banner_h - total_h) // 2 - 8
    for i, line in enumerate(lines):
        box = line_boxes[i]
        tw = box[2] - box[0]
        x = (banner_w - tw) // 2
        color = (255, 255, 255, 255) if i == 0 else (120, 232, 93, 255)
        shadow = Image.new("RGBA", banner.size, (0, 0, 0, 0))
        shadow_draw = ImageDraw.Draw(shadow)
        shadow_draw.text((x + 12, y + 14), line, font=fonts[i], fill=(0, 70, 38, 210), stroke_width=7, stroke_fill=(0, 70, 38, 210))
        shadow = shadow.filter(ImageFilter.GaussianBlur(2))
        text_layer = Image.alpha_composite(text_layer, shadow)
        text_draw = ImageDraw.Draw(text_layer)
        text_draw.text((x, y), line, font=fonts[i], fill=color, stroke_width=7, stroke_fill=(34, 91, 36, 255))
        text_draw.text((x - 3, y - 4), line, font=fonts[i], fill=(255, 255, 255, 95))
        y += (box[3] - box[1]) + 12

    banner = Image.alpha_composite(banner, text_layer)
    border = ImageDraw.Draw(banner)
    border.rounded_rectangle((10, 10, banner_w - 10, banner_h - 10), radius=28, outline=(255, 255, 255, 170), width=6)
    banner.convert("RGB").save(BANNER_OUT, quality=95)


if __name__ == "__main__":
    main()
