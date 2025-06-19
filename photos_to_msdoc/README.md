# Photos to Word Document

A Python script to create a Microsoft Word document with images from a folder.  
Each image is placed on its own page, preceded by the image filename as a title.

---

## Features

- Supports common image formats: JPG, PNG, GIF, BMP, TIFF, WebP - Change it at script level.
- Automatically scales images to fit the page while preserving aspect ratio.
- Adds image filename as a centered title above each image.
- Each image saves to a new page for easy printing.
- Sets custom page margins and default font style.

---

## Requirements

- Python 3.6+
- [`python-docx`](https://python-docx.readthedocs.io/en/latest/)
- [`Pillow`](https://python-pillow.org/)

Install dependencies with:

```bash
pip install -r requirements.txt

