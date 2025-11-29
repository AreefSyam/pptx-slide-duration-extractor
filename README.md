## 🧠 **Purpose of the Script**

This Python script:

1. Reads a PowerPoint file (`.pptx`)
2. Looks inside each slide to extract **timing durations** (if available)
3. Outputs this information into an **Excel file** with:
   - `number page`: the slide number
   - `duration`: how long the slide is shown (in seconds)

---

## 🧩 **Step-by-step Breakdown**

### 📁 1. File Setup

```python
pptx_filename = "your-slide.pptx"
excel_output = "slide_durations.xlsx"
```

- Sets filenames for the input PowerPoint and output Excel file.

```python
script_dir = os.path.dirname(os.path.abspath(__file__))
pptx_path = os.path.join(script_dir, pptx_filename)
output_path = os.path.join(script_dir, excel_output)
```

- Automatically sets the **absolute path** to both files based on where the script is saved and run.
- Ensures the script and `.pptx` file are in the **same folder**.

---

### 📦 2. Open and Read PPTX

```python
with zipfile.ZipFile(pptx_path, 'r') as pptx:
```

- PowerPoint files (`.pptx`) are actually ZIP files containing XML.
- This opens the file like a ZIP archive.

```python
    slide_files = [
        f for f in pptx.namelist()
        if re.match(r"ppt/slides/slide\d+\.xml$", f)
    ]
```

- Gets only the files that are actual slide contents (like `slide1.xml`, `slide2.xml`, etc.)

```python
    slide_files = sorted(
        slide_files,
        key=lambda x: int(re.search(r"slide(\d+)\.xml", x).group(1))
    )
```

- Sorts the slides **numerically**, so `slide10.xml` doesn't come before `slide2.xml`.

---

### 🔍 3. Loop Through Slides & Extract Duration

```python
    for idx, slide_file in enumerate(slide_files, start=1):
```

- Goes through each slide one by one, keeping track of the slide number (`idx`).

```python
        with pptx.open(slide_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
```

- Parses the slide's XML content.

```python
            transition = None
            for elem in root.iter():
                if 'transition' in elem.tag:
                    transition = elem
                    break
```

- Searches for the `<transition>` tag. This is where the timing is stored if the slide has **auto-advance** enabled.

```python
            if transition is not None:
                adv_tm = transition.attrib.get('advTm')
                if adv_tm:
                    duration_sec = round(int(adv_tm) / 1000, 2)
                else:
                    duration_sec = ''
            else:
                duration_sec = ''
```

- If `advTm` (advance time) exists, it converts it from milliseconds to seconds.
- If not, leaves it blank.

```python
            rows.append([idx, duration_sec])
```

- Stores the slide number and its duration into a list called `rows`.

---

### 📊 4. Export to Excel

```python
df = pd.DataFrame(rows, columns=["number page", "duration"])
df.to_excel(output_path, index=False)
```

- Converts the list to a DataFrame (like a table).
- Exports it to `slide_durations.xlsx`.

---

### ✅ 5. Print Success Message

```python
print(f"Slide durations exported to Excel: {output_path}")
print("Done")
print("Script completed successfully.")
```

- Prints a confirmation message that the process is complete.

---

## 🔚 Summary

| Task                  | Done? ✅ |
| --------------------- | -------- |
| Read PPTX as ZIP      | ✅       |
| Parse each slide      | ✅       |
| Extract duration info | ✅       |
| Save to Excel         | ✅       |
| Simple, clean script  | ✅✅     |

---
