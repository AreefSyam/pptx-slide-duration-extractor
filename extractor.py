import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os

# === File setup ===
pptx_filename = "your-file.pptx"
excel_output = "slide_durations.xlsx"

# Get absolute paths
script_dir = os.path.dirname(os.path.abspath(__file__))
pptx_path = os.path.join(script_dir, pptx_filename)
output_path = os.path.join(script_dir, excel_output)

rows = []

with zipfile.ZipFile(pptx_path, 'r') as pptx:
    slide_files = [
        f for f in pptx.namelist()
        if re.match(r"ppt/slides/slide\d+\.xml$", f)
    ]

    # Sort slides by number
    slide_files = sorted(
        slide_files,
        key=lambda x: int(re.search(r"slide(\d+)\.xml", x).group(1))
    )

    for idx, slide_file in enumerate(slide_files, start=1):
        with pptx.open(slide_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()

            # Find transition element (no namespace version)
            transition = None
            for elem in root.iter():
                if 'transition' in elem.tag:
                    transition = elem
                    break

            if transition is not None:
                adv_tm = transition.attrib.get('advTm')
                if adv_tm:
                    duration_sec = round(int(adv_tm) / 1000, 2)
                else:
                    duration_sec = ''
            else:
                duration_sec = ''

            rows.append([idx, duration_sec])

# Create DataFrame and save to Excel
df = pd.DataFrame(rows, columns=["number page", "duration"])
df.to_excel(output_path, index=False)

print(f"Slide durations exported to Excel: {output_path}")
print("Done")
print("Script completed successfully.")