# Morning package — read me first

Good morning. This folder contains everything you need to:

1. Bring your live site up to date with all the recent changes (About, Anatomy, Tips, refined styles)
2. Wire up the auto-regenerate workflow so future content edits happen via the spreadsheet — no code touching

The goal: by the end of this README, you upload one set of files to your `eir` repo and walk away. Going forward, you edit the spreadsheet and the site updates itself.

## What's in this folder

```
morning-package/
├── about.html                  ← updated page
├── anatomy.html                ← NEW page (Citizen Math hype)
├── tips.html                   ← updated page
├── index.html                  ← updated
├── styles.css                  ← updated
├── (11 convo detail pages)     ← regenerated, fresh
├── generate_pages.py           ← the build script (used by the Action)
├── convos.xlsx                 ← your spreadsheet (source of truth)
├── .github/
│   └── workflows/
│       └── build.yml           ← the GitHub Action
└── README.md                   ← this file
```

Everything in this folder should end up in your `eir` repo. The folder structure matters — `.github/workflows/build.yml` needs to live in those exact subfolders, not at the root.

## The upload (10 minutes, all-at-once)

### Step 1: Upload the regular files

1. Open the unzipped `morning-package` folder on your computer.
2. Go to your `eir` repo on GitHub.
3. Click **"Add file" → "Upload files."**
4. **Select all the files except the `.github` folder** — i.e., all the `.html`, `.css`, `.py`, and `.xlsx` files — and drag them onto the upload area.
5. Wait for the uploads to finish.
6. Scroll to "Commit changes." Leave the defaults. Click **"Commit changes."**

GitHub will overwrite the existing files with these new versions. Your live site (`gstuck.github.io/eir/`) will update within 1-2 minutes.

At this point you can refresh your live site and check that the new About, Anatomy, and Tips pages look right.

### Step 2: Add the GitHub Action

The `.github/workflows/build.yml` file lives in a subfolder, which makes drag-and-drop awkward. Easier to create it directly in the GitHub UI.

1. In your repo, click **"Add file" → "Create new file."**
2. In the filename box, type exactly: `.github/workflows/build.yml`
   - As you type the slashes, GitHub will automatically turn them into folders. That's expected.
3. Open `.github/workflows/build.yml` from the morning-package folder on your computer in any text editor (TextEdit, Notepad, VS Code — anything).
4. Select all, copy. Paste into the GitHub editor.
5. Scroll down. Click **"Commit changes."**

### Step 3: Watch the Action run

GitHub now sees the workflow file and is allowed to run it.

1. In your repo, click the **"Actions"** tab in the top nav.
2. You'll likely see a workflow run for "Regenerate site pages from spreadsheet" — it might be running (yellow dot), already finished (green check), or queued.
3. If you don't see it run automatically, click into the workflow name in the left sidebar and click **"Run workflow"** in the top right to trigger it manually.
4. Click into the run to see logs if you're curious. It should say "Loaded 11 convos" and either "No HTML changes" (most likely, since we just uploaded fresh pages) or commit a small whitespace cleanup.

That's the whole setup.

## Your new editing workflow

Whenever you want to update convo content:

1. **Download `convos.xlsx`** from the repo.
   - Click the file in the repo, then click the download icon (looks like a downward arrow).
2. **Edit it locally** in Excel, Numbers, or Google Sheets (re-save as .xlsx if you use Sheets).
3. **Re-upload it** to replace the old version.
   - Go to the repo root → **"Add file" → "Upload files"** → drag the new .xlsx → commit.
4. **Wait ~30 seconds.** The Action will run automatically (you can watch in the Actions tab if you want).
5. **Refresh your live site.** New content will be there.

## What this workflow touches and doesn't touch

**Generated automatically from the spreadsheet** (don't hand-edit these — they'll be overwritten):
- `index.html`
- The 11 convo detail pages (e.g. `right-answers.html`)

**Hand-edited only** (edit directly in GitHub if you want to change these):
- `about.html`
- `anatomy.html`
- `tips.html`
- `styles.css`
- Anything in `/pdfs/`

**Manual one-off** (the spreadsheet workflow doesn't help here):
- Adding or replacing a PDF in `/pdfs/`. Same drag-and-drop flow as always.

## Adding a 12th convo (if you ever want to)

1. Add a row to `convos.xlsx` with Convo # = 12 and all the columns filled in.
2. Upload the spreadsheet to GitHub.
3. The Action will create `your-new-slug.html` automatically (slug derived from the Short Title — e.g. "My New Topic" → `my-new-topic.html`).
4. Manually upload the merged PDF to `pdfs/your-new-slug.pdf` (and, if it's a PDF spark, the standalone spark to `pdfs/sparks/your-new-slug-spark.pdf`).

## A few decisions baked in (in case you want to change them)

- **Locked URLs for the first 11 convos.** Even if you rename a "Short Title" in the spreadsheet, the URL stays stable. This protects the existing PDFs from breaking. To rename a URL, you'd edit `LOCKED_SLUGS` in `generate_pages.py` and rename the matching PDFs.
- **Theme order on the homepage** (Deepening Understanding → Discourse → Real World) is hardcoded in `THEME_ORDER` near the top of `generate_pages.py`. New themes you add to the spreadsheet would appear at the bottom unless you add them to that list.
- **Theme-lesson callouts** (e.g. "Big Foot Conspiracy by Citizen Math" under the Discourse theme) are also in `generate_pages.py` as `THEME_LESSONS`. If you add a new theme, edit this dict to give it a lesson callout, or it'll show no callout.
- **The Action only commits when HTML actually changed**, so it won't clutter your commit history with empty re-runs.

## If something looks off

- **Action shows a red X (failed):** Click into it and read the error. The most likely cause is a bad spreadsheet — a missing column, an unexpected cell value, etc. Fix the spreadsheet, re-upload.
- **Site doesn't reflect changes after 1-2 minutes:** Hard-refresh the page (Cmd+Shift+R on Mac, Ctrl+Shift+R on Windows) to bypass browser cache. GitHub Pages itself can take a beat to push edge updates.
- **You want to force a re-run:** Actions tab → click the workflow → "Run workflow" button (top right) → click the green button.

That's it. Send any questions when you're up.
