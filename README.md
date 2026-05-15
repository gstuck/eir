# Convos with Colleagues

A standalone site of 11 roundtable discussion guides for math teachers.

## Structure
- `index.html` — homepage
- `*.html` — convo detail pages, about, tips
- `styles.css` — single stylesheet
- `pdfs/` — merged convo guide PDFs (one per convo)
- `pdfs/sparks/` — standalone spark PDFs (for PDF-spark convos)

## Local preview
```
python3 -m http.server
```
Then open http://localhost:8000

## Deploy
Drop this folder onto Vercel (or any static host).
