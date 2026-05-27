import sys
import zipfile


docx = sys.argv[1]
needle = sys.argv[2]
before = int(sys.argv[3]) if len(sys.argv) > 3 else 500
after = int(sys.argv[4]) if len(sys.argv) > 4 else 3000

with zipfile.ZipFile(docx) as zf:
    xml = zf.read("word/document.xml").decode("utf-8")

index = xml.find(needle)
print(index)
print(xml[max(0, index - before): index + after])
