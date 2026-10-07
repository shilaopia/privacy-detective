"""Extract full text from GB/T 39770-2021 PDF and write to a text file."""
import fitz
import os
import config

pdf_path = os.path.join(config.BASE, "National Standard", "归档（隐私政策中有体现）", "APP服务方面", "GBT+39770-2021.pdf")
out_path = os.path.join(config.BASE, "gb39770_full_text.txt")

doc = fitz.open(pdf_path)
print(f"Total pages: {doc.page_count}")

full_text = []
for i in range(doc.page_count):
    text = doc[i].get_text()
    full_text.append(f"\n=== PAGE {i+1} ===\n{text}")

doc.close()

result = "\n".join(full_text)
with open(out_path, "w", encoding="utf-8") as f:
    f.write(result)

print(f"Saved {len(result)} chars to {out_path}")
print(f"First 500 chars:\n{result[:500]}")
