---
name: python-ppt-generator
description: A tool to generate professional, fully editable, McKinsey-style native PowerPoint (.pptx) presentations using Python. Supports multiple layouts (cover, two-column, three-column, timeline, matrix).
---

# Python PPT Generator (V3)

Generate native, 100% editable `.pptx` files using `python-pptx`. Unlike HTML-based converters, this skill generates true PowerPoint files with advanced layouts (columns, timelines, matrices) and automatic font scaling.

## How to use

Pass a JSON file containing the presentation structure to the script. 

`python generate_ppt.py /path/to/slides_data.json /path/to/output.pptx`

### JSON Format Expected:

**Cover Layout:**
```json
{
  "layout": "cover",
  "title": "Main Title",
  "subtitle": "Subtitle here"
}
```

**Columns Layout (two-column or three-column):**
```json
{
  "layout": "two-column",
  "action_title": "Slide Title",
  "takeaway": "Bottom insight box text",
  "columns": [
    {
      "title": "Column 1 Header",
      "bullets": ["Point 1: explanation", "Point 2"]
    },
    {
      "title": "Column 2 Header",
      "bullets": ["Point A: explanation", "Point B"]
    }
  ]
}
```

**Timeline Layout:**
```json
{
  "layout": "timeline",
  "action_title": "Slide Title",
  "steps": [
    {"title": "Phase 1\n(2020-2022)", "desc": "Description here"},
    {"title": "Phase 2\n(2023-2025)", "desc": "Description here"}
  ]
}
```

**Matrix Layout (2x2 Grid):**
```json
{
  "layout": "matrix",
  "action_title": "Slide Title",
  "quadrants": [
    {"title": "Top Left", "desc": "Description here"},
    {"title": "Top Right", "desc": "Description here"},
    {"title": "Bottom Left", "desc": "Description here"},
    {"title": "Bottom Right", "desc": "Description here"}
  ]
}
```

The script automatically applies business styling, bolds text before colons (`：` or `:`) in bullets, and uses `Text to Fit Shape` to prevent text overflow.
