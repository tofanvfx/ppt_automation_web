# DOCX to PowerPoint Generation Guide

This tool automates the creation of educational PowerPoint presentations from a structured Word document (DOCX). This guide explains how to format your Word document to ensure everything transitions perfectly to your slides.

---

## 1. General Structure

The script reads your Word document line-by-line. Each **Slide** is defined by a section starting with a **Layout Marker**.

- **Layout Marker**: A name in square brackets, e.g., `[sst_content_page_01]`.
- **Content**: All text and images appearing _after_ a marker will be placed on that specific slide.
- **Section Break**: A new layout marker automatically ends the previous slide and starts a new one.

---

## 2. Global Presentation Info (Title Slide)

To fill in the metadata on your title slide, use a simple Table or a list with colons at the very beginning of your document.

**Supported Metadata Fields:**

- `CLASS:` (e.g., Class - 10)
- `SUBJECT:` (e.g., Science)
- `CHAPTER NUMBER:` (e.g., Chapter - 1)
- `CHAPTER NAME:` (e.g., Chemical Reactions)
- `LESSON:` (e.g., Lesson - 1)
- `TOPIC:` (e.g., Displacement Reactions)

---

## 3. Slide Layouts & Examples

### A. General Content Slides

Used for standard teaching material with text and images.

- `[sst_content_page_01]` (Standard Layout)
- `[sst_content_page_02]` (Alternative Layout)
- `[sst_deafult_page]` (New Default Layout)
- `[math_deafult_page]` (Math Default Layout)

**Example:**

```text
[sst_content_page_01]
TOPIC: Photosynthesis
Plants use sunlight to produce food.
• Chlorophyll is essential.
• Oxygen is released.
[Insert Image Here]
```

### B. Quiz Slides

Automatically formats questions and options.

- `[sst_quiztime_page]`

**Example:**

```text
[sst_quiztime_page]
Question: What is the capital of France?
Options: Paris, London, Berlin, Madrid
```

_Note: The script automatically chooses a layout based on whether your options are short or long._

### C. Activity Pages

Specially designed for interactive tasks.

- `[sst_activity_page]`

**Features:**

- **Dynamic Resizing**: The yellow activity box will automatically expand downwards if you add a lot of text.
- **Multi-Image Support**: If you add multiple images, the script will arrange them cleanly in square slots without stretching.

**Example:**

```text
[sst_activity_page]
Task: Identify the parts of the flower shown below.
[Image 1] [Image 2]
```

### D. Discussion Pages

- `[sst_discussion_page]`

**Example:**

```text
[sst_discussion_page]
Why do we need to conserve water?
```

---

## 4. Advanced Elements

### 📐 Equations (Math)

You can use native Word equations (**Insert > Equation**). These will be imported into PowerPoint as **native math objects**, meaning they remain high-quality and fully editable in PowerPoint.

### 🖼️ Images

- **Single Image**: Will be placed in the designated picture placeholder.
- **Multiple Images**: Supported in Activity and Content pages. They are automatically scaled and aligned.

### 🎭 Special Overlays

You can add special floating elements to slides using these markers:

- `[add_syr]` or `[syr]`: Adds a "Stay Your Resources" branding/icon overlay.
- `[add_question (Your Text)]`: Adds a floating "Ask a Question" box with your specific text.

---

## 5. Summary of Markers

| Marker                  | Purpose                           |
| :---------------------- | :-------------------------------- |
| `[sst_lo_page]`         | Learning Objectives slide         |
| `[math_lo_page]`        | Math-specific Learning Objectives |
| `[sst_content_page_01]` | Standard content slide            |
| `[sst_deafult_page]`    | New Default Layout                |
| `[math_deafult_page]`   | Math Default Layout               |
| `[sst_quiztime_page]`   | Quiz slide with options           |
| `[sst_activity_page]`   | Interactive activity slide        |
| `[sst_discussion_page]` | Discussion/Question slide         |
| `[sst_notedown_page]`   | Final summary/notedown slide      |
| `[homework]`            | Homework assignment slide         |

---

## 6. Pro Tips

1. **Bullet Points**: Use standard bullet points in Word (`•` or `-`). Indenting them in Word (Tab) will preserve the nesting levels in PowerPoint.
2. **Bold/Italics**: Basic formatting is preserved in many placeholders.
3. **No Stretch**: Images are always scaled to "fit" their area without being stretched or distorted.


