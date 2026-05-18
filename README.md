# Invoice Submittal Package Generator

A local web app that merges a filled-in cover letter with an invoice PDF into a single downloadable package.

---

## Prerequisites

### 1. Node.js (v18+)
Download from https://nodejs.org or install via Homebrew:
```bash
brew install node
```

### 2. LibreOffice
Required to convert the filled DOCX cover letter to PDF.

**macOS:**
```bash
brew install --cask libreoffice
```
Or download the installer from https://www.libreoffice.org/download/download/

**Ubuntu / Debian Linux:**
```bash
sudo apt update && sudo apt install libreoffice
```

---

## Setup

```bash
# 1. Move into the project directory
cd invoice-submittal

# 2. Install Node dependencies
npm install

# 3. Start the server
npm start
```

Open your browser to **http://localhost:3000**

---

## Customizing Contractors & Senders

Edit the top of `server.js`:

```js
const CONTRACTORS = [
  'Acme Construction Co.',
  'BuildRight Inc.',
  // add more here…
];

const SENDERS = {
  'Alice Johnson':  'signatures/alice_johnson.png',
  'Bob Martinez':   'signatures/bob_martinez.png',
  // add more senders here…
};
```

---

## Signature Images

Each sender entry maps to a PNG file in the `/signatures` folder.

- File format: PNG (transparent background recommended)
- Suggested size: ~450 × 150 px (displays at 150 × 50 px in the document)
- Naming must match exactly what's in `SENDERS` in `server.js`

Example mapping:
```
'Alice Johnson' → signatures/alice_johnson.png
```

### DOCX Template Placeholder

To embed a signature image in the cover letter, add this placeholder
in your .docx template at the spot where the signature should appear:

```
{%signature}
```

If no signature file exists for the selected sender, or if the placeholder
is absent, the app falls back gracefully and the rest of the document
renders normally.

---

## DOCX Template Placeholders

Your cover letter template must use these exact placeholders:

| Placeholder          | Filled with            |
|----------------------|------------------------|
| `{{today_date}}`     | Today's Date field     |
| `{{contractor_name}}`| Contractor Name field  |
| `{{invoice_number}}` | Invoice Number field   |
| `{{invoice_amount}}` | Invoice Amount field   |
| `{{date_range}}`     | Date Range field       |
| `{{sender_name}}`    | Sender Name field      |
| `{%signature}`       | Signature image (PNG)  |

---

## Output Filename Format

```
{Contractor Name} Invoice Submittal Package - PE {MM-DD-YYYY} {Invoice Number}.pdf
```

Example:
```
Acme Construction Co. Invoice Submittal Package - PE 05-15-2026 INV-2025-042.pdf
```

---

## PDF Generation Flow

1. User uploads invoice PDF + DOCX template
2. Server fills all `{{placeholders}}` using docxtemplater
3. Signature image injected via `{%signature}` if file exists
4. LibreOffice converts filled DOCX → PDF (in a temp directory)
5. pdf-lib merges: cover letter PDF first, invoice PDF second
6. Final PDF streamed back to browser as a download — nothing saved to disk

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| "LibreOffice not found" | Install LibreOffice; restart the server |
| "Template rendering error" | Check that all `{{placeholder}}` names match exactly |
| Signature not appearing | Verify the PNG file exists in `/signatures` with the exact name |
| Invoice PDF fails to load | Ensure the PDF is not encrypted or password-protected |
| Port already in use | `PORT=3001 npm start` to use a different port |
