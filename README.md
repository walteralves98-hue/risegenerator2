# ADRA RISE Activation Studio

A static, no-login, browser-based emergency activation proposal drafting tool built from the supplied RISE Pillar 1 templates and NEMP examples.

## Deploy to GitHub + Vercel

1. Extract this ZIP. Upload **the contents** of the `ADRA_RISE_Generator` folder to a new GitHub repository (keep `index.html`, `app.js` and `templates/` at repository root).
2. In Vercel, choose **Add New > Project**, import the GitHub repository.
3. Framework Preset: **Other**. Build Command: leave empty. Output Directory: leave empty. Deploy.
4. Visit the Vercel URL. Choose a country, enter verified response details, and click **Generate full activation package**.

You can also test locally with `python -m http.server 8000` from the project folder, then open `http://localhost:8000`.

## Important limits

- No login or database. Draft inputs remain in browser memory until exported to JSON; use Save Draft to retain them.
- Third-party JavaScript libraries (docx, ExcelJS, JSZip) load from a CDN, so document export requires internet access. User data is processed locally in the browser.
- The original uploaded donor templates are included unchanged. The generated Word documents and spreadsheets are **editable planning drafts** and do not reproduce every field or styling detail of the original donor forms.
- The generated budget planning workbook does not replace the original RISE detailed budget template. Indirect cost and MTDC exclusions must be reviewed by finance.
- RISE indicators C00, C07, C08 are included; additional sector-specific indicators must be selected and verified.
- Honduras, Mexico and Ecuador Scenario 1 reference values are historical NEMP examples, not verified emergency needs or market prices.
- Proposal data should be treated as sensitive. No authentication is included. Do not use a publicly accessible shared computer for sensitive drafts.
- CI, DHR, technical, safeguarding, compliance and finance review/approval are not automated.


## Upload NEMP from any country
Open Activation setup and upload PDF (text-based), DOCX, or XLSX. The browser extracts text, offers an editable country/intervention review, and links the approved extraction to the activation. No country list restriction. Source extract and review metadata are included in generated ZIP. Scanned PDFs require prior OCR. This is a human-reviewed source import, not automatic verified data extraction; targets and budget remain user inputs.

## Library loading troubleshooting
JSZip is now bundled locally. DOCX, ExcelJS, PDF.js and Mammoth are loaded with fallback CDNs (jsDelivr then unpkg). If both CDNs are blocked, package generation still cannot run. For fully offline/restricted-network operation, vendor the remaining JavaScript distributions under vendor/ and update index.html to use those local paths. Deploy the entire folder, including vendor/.
