# Python PPT Generator

A robust, automated tool to generate professional, fully editable, consulting-style native PowerPoint (`.pptx`) presentations using Python. 

It supports multiple high-density business layouts (covers, multi-columns, timelines, matrices, image-text) and automatically handles text scaling, color theming, and vector icon integration.

## Features

* **Multiple Layouts**: `cover`, `two-column`, `three-column`, `timeline` (auto-stretching), `matrix` (2x2), and `image-text`.
* **Auto Vector Icons**: Fetches high-quality SVG icons from Iconify API (e.g., `lucide:shield`), automatically recolors them to match the corporate theme, and inserts them as native `.pptx` elements.
* **Smart Typography**: Leverages PPT's auto-fit to prevent text overflow. Automatically bolds text before colons in bullet points.
* **Fully Editable**: Generates 100% native `.pptx` shapes and text boxes. No uneditable background images.

## Installation

Requires Python 3.x and the following packages:

```bash
pip install python-pptx cairosvg requests pillow
```
*(Note: `cairosvg` may require system-level Cairo graphics libraries to be installed).*

## Usage

Pass a JSON file containing the presentation structure to the script:

```bash
python generate_ppt.py data.json output.pptx
```

## Author
**Fang Min**  
Email: 130218391+Marathon07@users.noreply.github.com

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
