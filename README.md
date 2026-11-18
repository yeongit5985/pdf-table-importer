# PDF Table Importer (Excel add-in)

Pick several PDF files, preview the tables found in them, and import the ones you want into Excel.

- Reads PDFs **inside the task pane** (pdf.js). Files never leave your computer.
- Finds tables automatically, or treats each whole page as one table.
- Three output modes: a new sheet per table, all tables on one sheet, or **one combined table** (repeated headers dropped, with source file and page columns). Combined mode suits the same report split across many PDFs or pages.
- Converts numbers like `1,234.50`, `(99)` and `120-`. Keeps IDs (`00123`), dates and codes as text.

It only works on PDFs that contain real text. Scanned PDFs need OCR first.

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Tells Excel where the add-in lives. Not needed on GitHub, but harmless there. |
| `taskpane.html / .css / .js` | The add-in itself (static files, no build step) |
| `commands.html` | Required by the ribbon button |
| `assets/` | Icons |
| `samples/` | Two test PDFs |
| `set-url.ps1` | Writes your GitHub Pages address into the manifest |

## 1. Publish on GitHub Pages

1. On github.com, create a **public** repository named `pdf-table-importer`.
2. Click **Add file → Upload files** and drag in everything from this folder (keep the `assets` folder). Commit.
3. Open **Settings → Pages**. Under *Build and deployment*, pick **Deploy from a branch**, then branch `main` and folder `/ (root)`. Save.
4. After about a minute, check that `https://<your-user>.github.io/pdf-table-importer/taskpane.html` opens.

## 2. Point the manifest at your site

In PowerShell, in this folder:

```powershell
.\set-url.ps1 -GitHubUser <your-user>
```

## 3. Load it into Excel desktop (Windows)

1. Make a folder, e.g. `C:\ExcelAddins`, and copy `manifest.xml` into it.
2. Right-click the folder → **Properties → Sharing → Share…** → Share. Note the network path, e.g. `\\YOUR-PC\ExcelAddins`.
3. In Excel: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Paste the network path into *Catalog Url*, click **Add catalog**, tick **Show in Menu**, then OK. Restart Excel.
4. **Home → Add-ins → More Add-ins** (or **Insert → My Add-ins**) → **SHARED FOLDER** tab → *PDF Table Importer* → **Add**.
5. Find the **Import PDF Tables** button on the **Data** tab.

## Using it

1. Choose or drop PDF files, then click **Find tables**.
2. Untick tables you don't want. Use **Preview** to check them.
3. Pick an output mode and click **Import selected tables**.

If tables come out wrong, open **Detection settings**:

| Problem | Try |
|---|---|
| Two columns merged into one | Lower **Column gap** (e.g. 0.5) |
| One column split in two | Raise **Column gap** (e.g. 1.2–1.5) |
| Small tables missed, or junk detected | Change **Min. rows** |
| Long cell text spills onto extra rows | Tick **Join wrapped text** |
| Table not found at all | Use **Whole page as one table** and remove the extra rows in Excel |

## Updating

Upload the changed files to the repository again. If Excel keeps showing the old version, close Excel and delete the contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
