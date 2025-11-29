# PPTX to Excel Extraction (Python)

## 🚀 Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
```

### 2. Install dependencies

Make sure you have Python 3 installed.

```bash
pip install -r requirements.txt
```

### 3. Run the script

Place your `.pptx` file in the same folder as the script.
Then run:

```bash
python extract_slide_duration.py
```

---

# PPTX to Excel Extraction (Python)

A simple Python tool that extracts slide auto-advance durations from a PowerPoint (`.pptx`) file and exports them into an Excel file.

---

## 🧾 Purpose

This script reads the XML inside a PPTX (a ZIP archive), finds each slide's `<transition>` element (if present), extracts the `advTm` attribute (advance time in milliseconds), converts it to seconds, and writes a table of slide numbers and durations to an Excel file.

---

## ⚙️ Prerequisites

- Python 3.8 or newer
- `pandas` and `openpyxl` installed

Install dependencies:

```bash
pip install pandas openpyxl
```

Or use the provided `requirements.txt`:

```bash
pip install -r requirements.txt
```

---

## 📁 Files

- `extract_durations.py` — main script (example name)
- `requirements.txt` — dependency list
- `slide_durations.xlsx` — output generated after running

---

## 🚀 Quick Start

1. Clone the repo and `cd` into it:

```bash
git clone <your-repo-url>
cd pptx-slide-duration-extractor
```

2. Put your PPTX file in the project folder (or update the script to point to a different path). The example uses `your-slide.pptx`.

3. Run the script:

```bash
python extract_durations.py
```

4. Find `slide_durations.xlsx` in the same folder. It contains two columns: `number page` and `duration` (seconds).

---

## 🔎 Script: Step-by-step explanation

Below explains the exact flow inside the script.

### 1. Setup file paths

```python
pptx_filename = "your-slide.pptx"
excel_output = "slide_durations.xlsx"

script_dir = os.path.dirname(os.path.abspath(__file__))
pptx_path = os.path.join(script_dir, pptx_filename)
output_path = os.path.join(script_dir, excel_output)
```

- The script resolves absolute paths so it can be run from anywhere, assuming the PPTX is next to the script.

### 2. Open PPTX as ZIP and collect slide XML files

```python
with zipfile.ZipFile(pptx_path, 'r') as pptx:
    slide_files = [
        f for f in pptx.namelist()
        if re.match(r"ppt/slides/slide\d+\.xml$", f)
    ]
```

- A PPTX file is a ZIP archive; slide content lives under `ppt/slides/slideN.xml`.
- The script collects only those files.

### 3. Sort slide files numerically

```python
slide_files = sorted(
    slide_files,
    key=lambda x: int(re.search(r"slide(\d+)\.xml", x).group(1))
)
```

- Ensures `slide2.xml` comes before `slide10.xml` by extracting the slide number and sorting numerically.

### 4. Parse each slide XML and find `<transition>`

```python
for idx, slide_file in enumerate(slide_files, start=1):
    with pptx.open(slide_file) as f:
        tree = ET.parse(f)
        root = tree.getroot()

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
```

- Walks the XML tree looking for a tag name that contains `transition` (namespace-agnostic).
- If `advTm` exists, it is returned in milliseconds; the script converts it to seconds (rounded to 2 decimals).
- Stores the result in a list of rows.

### 5. Export results to Excel

```python
df = pd.DataFrame(rows, columns=["number page", "duration"])
df.to_excel(output_path, index=False)
```

- Uses `pandas` to create a table and write it to an `.xlsx` file.

### 6. Completion messages

The script prints the path to the generated Excel and a simple success message.

---

## 🛠 Troubleshooting

- **FileNotFoundError**: Check that `your-slide.pptx` exists in the same folder or update `pptx_filename` to an absolute path.
- **Blank durations**: If slides don’t show timings, they might not have auto-advance set. The script only reads `advTm` on `<transition>` elements. Some PPTX files use namespaces — see "Improvements".
- **Excel write errors**: Make sure `slide_durations.xlsx` is not open in Excel or another program.

---

## ✅ Suggested improvements

- Add command-line arguments (e.g., `--input`, `--output`) using `argparse`.
- Use namespace-aware XML parsing to handle slides where tags include XML namespaces.
- Add CSV export option.
- Create unit tests for different PPTX variants.

---

## 📄 License

MIT License — free to use and modify.

---

## 🧩 Contributing

Open an issue or submit a pull request. Include sample PPTX files if adding support for new transition formats.
