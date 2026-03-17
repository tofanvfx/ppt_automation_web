# Worksheet Generator — DOCX → PPTX

Converts a structured Word document into an Aveti Learning worksheet slide,
with the logo, header, and footer locked to the template and all question
content dynamically generated.

---

## Files

| File | Purpose |
|------|---------|
| `worksheet_generator.py` | Main Python script — reads DOCX, writes PPTX |
| `sample_worksheet.docx` | Sample input document (edit this to change content) |
| `create_docx.js` | Node.js script that generated `sample_worksheet.docx` |
| `worksheet_test.pptx` | Template file (header/logo/footer preserved automatically) |
| `output_worksheet.pptx` | Sample output — the generated worksheet |

---

## Quick Start

```bash
# Install dependencies (once)
pip install python-pptx python-docx

# Generate a worksheet
python worksheet_generator.py sample_worksheet.docx my_output.pptx worksheet_test.pptx
```

---

## DOCX Input Format

The input `.docx` file must follow this structure exactly.

### 1 — Metadata (required, at the top)

```
Chapter: Chapter 1 – Wonderful World of Science
Class: 6
Subject: Science
Worksheet: Worksheet 3 Key (Everyday Scientists & Collaboration)
```

### 2 — Section headers

Use letters A–F followed by a period and the section name:

```
A. MULTIPLE CHOICE QUESTIONS (MCQ)
B. FILL IN THE BLANKS (FIB)
C. TRUE OR FALSE (T/F)
D. ASSERTION–REASON
E. SHORT ANSWER QUESTIONS (SA)
F. LONG ANSWER QUESTION (LA)
```

Recognised section types and their formatting behaviour:

| Prefix | Type | Answer format |
|--------|------|---------------|
| MCQ    | Multiple Choice | `✔ Correct Answer: …` |
| FIB    | Fill in the Blank | `✔ Answer: …` |
| T/F    | True or False | `Answer: True/False – explanation` |
| AR     | Assertion–Reason | `✔ Correct Answer: (A/B/C/D) …` |
| SA     | Short Answer | `✔ Answer: …` |
| LA     | Long Answer | `✔ Answer:` + numbered lines |

### 3 — Questions

```
Q1. Question text here?
```

### 4 — MCQ / Assertion-Reason options

```
a) Option text
b) Option text
c) Option text
d) Option text
```

### 5 — Assertion-Reason questions

```
Q7. (Optional intro text)
Assertion: The assertion statement here.
Reason: The reason statement here.
a) Both Assertion and Reason are true, Reason is the correct explanation.
b) Both are true but Reason is NOT the correct explanation.
c) Assertion is true but Reason is false.
d) Assertion is false but Reason is true.
Answer: (A) Both Assertion and Reason are true and Reason is correct.
```

### 6 — Answers

**Single-line answer:**
```
Answer: The answer text here.
```

**True/False answer (verdict highlighted in red/green):**
```
Answer: True – Explanation of why it is true.
Answer: False – Explanation of why it is false.
```

**Multi-line answer (Long Answer):**
```
Answer:
Step one of the answer.
Step two of the answer.
Step three of the answer.
```

---

## How It Works

1. **Parse** — reads the DOCX, extracts metadata, sections, questions, options and answers
2. **Build blocks** — converts every element into a flat list of renderable blocks with dynamically estimated heights (avoids overflow for long text)
3. **Flow** — distributes blocks into left and right columns with keep-with-next logic (no orphaned section headers)
4. **Render** — places text boxes on the slide at computed Y positions with exact colour and bold formatting matching the template
5. **Save** — writes the output PPTX with the original header, logo, and footer unchanged

---

## Customisation

To change slide colours, edit the constants at the top of `worksheet_generator.py`:

```python
C_SECTION = RGBColor(0x00, 0x70, 0xC0)   # Blue  – section headers
C_ANSWER  = RGBColor(0x1F, 0x6B, 0x1F)   # Green – answer text
C_FALSE   = RGBColor(0xFF, 0x00, 0x00)   # Red   – "False" verdict
```

To change font size, edit `FS = 127000` (EMU; 127000 = 10pt).

---

## Generating a New Sample DOCX

The `create_docx.js` script is provided if you want to generate a pre-formatted sample DOCX programmatically:

```bash
npm install -g docx
node create_docx.js
# → creates sample_worksheet.docx
```
