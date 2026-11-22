# RUAS AI Question Paper Generator (Web)

A browser-based version of the RUAS AI Question Paper Generator. It runs entirely in the browser, so it can be hosted for free on **GitHub Pages**. Uploaded papers never leave the user's computer.

## Features
- Upload previous question papers (`.docx` and `.pdf`)
- Extracts questions from Word tables (Q No | Question | Marks | CO), grouping sub-parts such as 1a/1b
- Generates a new paper by picking one question per question number from across all uploaded papers, avoiding repeats when regenerating
- Uses the first `.docx` as the template (header, logos, formatting kept); the **last table** in it becomes the question table
- Keeps equations and images from the original question rows, including images copied from other papers
- Bloom's taxonomy level, quality score and improvement suggestions for each question, colour-coded
- Dashboard: average score, Bloom distribution, CO distribution
- **Authorization key + QR code**: the generated paper is encrypted (AES-256-GCM) into a `.qp` file. It can only be opened with the 20-character key, in section 3 of the page.

## Deploy on GitHub Pages
1. Create a new repository on GitHub (for example `ruas-qp-generator`).
2. Upload **all** files in this folder, keeping the `lib/` folder:
   ```
   index.html  app.js  style.css  README.md  .nojekyll  lib/
   ```
   (On github.com use **Add file → Upload files** and drag the whole folder contents in, or use git:)
   ```bash
   git init
   git add .
   git commit -m "RUAS QP Generator web app"
   git branch -M main
   git remote add origin https://github.com/<your-username>/ruas-qp-generator.git
   git push -u origin main
   ```
3. In the repository go to **Settings → Pages**. Under *Build and deployment*, set **Source: Deploy from a branch**, **Branch: `main`**, folder **`/ (root)`**, then **Save**.
4. After a minute or two your app is live at:
   ```
   https://<your-username>.github.io/ruas-qp-generator/
   ```

## Run locally
Open `index.html` directly in a browser, or serve the folder:
```bash
python -m http.server 8000   # then open http://localhost:8000
```

## How to use
1. **Upload** the previous papers. The first `.docx` is the template.
2. **Generate.** Set how many questions you want and click *Generate / Regenerate*.
3. Download the **Protected QP (.qp)**, the **Key (.txt)** and the **QR code (.png)**. Send the key or QR separately from the file.
4. To open a paper later, use **Open a protected question paper**: choose the `.qp` file, enter the key and the `.docx` is downloaded.

## Notes
- PDF extraction is approximate (text only, no formatting). Word `.docx` papers give the best results.
- Bundled libraries in `lib/`: JSZip, pdf.js (Mozilla) and qrcode-generator, each under its own open-source licence.


## Demo sample papers
Three anonymised-for-demo papers for EED103A (Introduction to Electrical Engineering, SEE Jan 2026) are included:
`SAMPLE_SET-1_EED103A.docx`, `SAMPLE_SET-2_EED103A.docx`, `SAMPLE_SET-3_EED103A.docx`.
On the page you can view each set, download one or all three (.zip), load all three into the generator,
or clear the samples from the page (remembered in this browser; "Show demo samples again" brings them back).
To remove the samples permanently from the website, delete the three SAMPLE_*.docx files from the repository.
