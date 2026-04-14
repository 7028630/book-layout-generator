# book-layout-generator
Book layout generator with Streamlit
# 📖 Book Layout Generator

A Streamlit app that generates a landscape Word document
formatted as a book — each printed sheet contains two
half-pages side by side. Print double-sided, cut down
the centre, stack, and you have a book.

## Features

- Landscape layout (A4, Letter, A5)
- Two book pages per printed sheet
- Adjustable footnote area per page
- Separator line between text and footnotes (adjustable)
- Per-page footnote numbering
- Font, size, line spacing and indent controls
- Sheet overview preview
- Download as .docx

## Setup

### 1. Clone the repo

git clone https://github.com/YOUR_USERNAME/book-layout-generator.git
cd book-layout-generator

### 2. Create a virtual environment (optional but recommended)

python -m venv venv
source venv/bin/activate        # Mac/Linux
venv\Scripts\activate           # Windows

### 3. Install dependencies

pip install -r requirements.txt

### 4. Run

streamlit run app.py

## How to Print as a Book

1. Open the generated .docx in Word
2. File → Print → Print on Both Sides (flip on long edge)
3. Cut each sheet down the vertical centre
4. Stack the half-sheets in page order
5. Bind or staple

## Notes

- Footnotes are per-page (numbered 1, 2, 3... per half-page)
- The footnote area size is fixed per document (set via slider)
  so plan your content length accordingly
- For booklet page ordering (so pages assemble in sequence
  after cutting) you need to order pages manually or use
  a dedicated imposition tool
