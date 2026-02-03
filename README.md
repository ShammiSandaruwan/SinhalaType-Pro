# SinhalaType Pro

SinhalaType Pro is a standalone desktop application that converts Sinhala Unicode text into Legacy FM Font encoding (specifically FMAbhaya) and inserts it directly into Adobe Photoshop as a text layer.

## Features

- **Unicode to FM Conversion**: Robust conversion logic including Kombuwa reordering and Rephaya/Rakaaransaya handling.
- **Photoshop Integration**: One-click "Send to Photoshop" functionality using COM automation.
- **Modern GUI**: Built with CustomTkinter for a sleek, dark-themed interface.
- **Always on Top**: Keeps the window accessible while working in other apps.
- **Clipboad Integration**: Auto-copy converted text.

## Usage

1.  Ensure Adobe Photoshop is open and a document is active.
2.  Run the application:
    ```bash
    python sinhala_type_pro.py
    ```
3.  Type or paste Sinhala Unicode text into the input box.
4.  Click **Send to Photoshop** to create a new text layer in Photoshop.
5.  Or click **Copy Text** to copy the legacy encoded text to your clipboard.

## Requirements

- Python 3.10+
- Adobe Photoshop (for integration features)

## Installation

Install dependencies:

```bash
pip install -r requirements.txt
```
