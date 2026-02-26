---
name: python-ppt-generator
description: A tool to generate professional, fully editable, McKinsey-style native PowerPoint (.pptx) presentations using Python. It uses a single unified text box for easy editing and applies automated business styling (bolding keywords, formatting citations).
---

# Python PPT Generator

Generate native, 100% editable `.pptx` files using `python-pptx`. Unlike Marp or HTML-based converters, this skill generates true PowerPoint files where all text is contained in a single, properly indented text box per slide.

## How to use

Pass a JSON file containing the presentation structure to the script. 

`python generate_ppt.py /path/to/slides_data.json /path/to/output.pptx`

### JSON Format Expected:

```json
[
  {
    "title": "Main Title",
    "subtitle": "Subtitle here",
    "is_cover": true
  },
  {
    "title": "1. First Slide",
    "bullets": [
      "Key point: detailed explanation [Citation: Source]",
      "Another point without citation"
    ]
  }
]
```

The script will automatically:
1. Apply a professional business theme (light gray background, deep blue headers).
2. Group all bullets into a single text frame for seamless editing.
3. Bold the text before any colon (`：` or `:`) in a bullet point.
4. Format citations (`[Citation: ...]`) as smaller, gray, italicized text.

