# IATA Air Waybill (Excel to PDF) Converter

A web-based software application that converts standard Air Waybill spreadsheets (`.xlsx`, `.xls`, `.csv`) into locked single-page IATA Air Waybill PDFs.

## Features
- **Batch Processing:** Upload multiple spreadsheets at once and download individual PDFs or a single `.zip` batch.
- **Locked Single-Page Format:** Preserves the exact IATA grid layout, shaded headers, legal clauses, and widened `Nature and Quantity of Goods` column (46% width).
- **Clean Handling Information:** Leaves the Handling Information box free of unwanted U.S. export clauses.
- **Smart Excel Parsing:** Automatically converts Excel serial dates (e.g., `46333.0` -> `07-NOV-2026`), maps Prepaid (`PP`) vs. Collect (`CC`) charges, and places Dhaka agent details inside the cargo description block.

## Deployment Instructions (GitHub + Streamlit Cloud)
1. Create a new **Private** repository on GitHub (e.g., `awb-pdf-converter`).
2. Upload `app.py`, `template.html`, `requirements.txt`, `packages.txt`, and `README.md` into the repository.
3. Go to [share.streamlit.io](https://share.streamlit.io), connect your GitHub repository, set `app.py` as the main file path, and click **Deploy**.
