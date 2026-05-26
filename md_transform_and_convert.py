import re
from pathlib import Path

p = Path('grow_a_garden_pet_guide.md')
s = p.read_text(encoding='utf-8')

# Insert tier links after the Source line
m = re.search(r'(?m)^Source:.*$', s)
if m:
    start = m.end()
    # find all level-2 headings
    tiers = re.findall(r'(?m)^##\s+(.*)$', s)
    if tiers:
        # create slug function similar to pandoc
        def slug(t):
            t2 = t.strip().lower()
            t2 = re.sub(r"[^a-z0-9\s-]", '', t2)
            t2 = re.sub(r"[\s]+", '-', t2)
            return t2
        links = ' | '.join(f'[{tier}](#{slug(tier)})' for tier in tiers)
        insert = '\n' + links + '\n\n'
        s = s[:start] + insert + s[start:]

# Transform each pet block (### PetName ... until next ## or ###)
pattern = re.compile(r'(?m)^(###\s+(.+?)\r?\n)((?:(?!^###\s|^##\s).*(?:\r?\n))*)', re.MULTILINE)

def replace_block(m):
    heading = m.group(2).strip()
    body = m.group(3)
    # find image URL
    img_match = re.search(r'- Image:\s*!\[.*?\]\((.*?)\)', body)
    if not img_match:
        return m.group(0)  # no change
    img_url = img_match.group(1).strip()
    # remove the image line from body
    body_clean = re.sub(r'(?m)^- Image:.*\r?\n', '', body)
    # build html table
    new = []
    new.append('<table>')
    new.append('<tr>')
    new.append(f'<td valign="top"><img src="{img_url}" width="120" height="120" /></td>')
    new.append('<td valign="top">')
    new.append(f'<h3>{heading}</h3>')
    new.append(body_clean.rstrip())
    new.append('</td>')
    new.append('</tr>')
    new.append('</table>\n')
    return '\n'.join(new)

s2 = pattern.sub(replace_block, s)

out = Path('grow_a_garden_pet_guide_for_docx.md')
out.write_text(s2, encoding='utf-8')
print('WROTE', out)

# Now call pandoc to convert
import subprocess
cmd = ['pandoc', str(out), '-o', 'grow_a_garden_pet_guide_updated.docx', '--resource-path=.']
print('RUNNING:', ' '.join(cmd))
try:
    subprocess.check_call(cmd)
    print('CONVERSION_OK')
except subprocess.CalledProcessError as e:
    print('CONVERSION_FAILED', e)
