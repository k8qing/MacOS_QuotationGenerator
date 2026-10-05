# Keith's Quotation Generator (Word + PDF)

## Using the app
1. Open `KeithsQuotationGenerator.exe` (keep it with `products.csv`, `images/` and `templates.json`).
2. Fill in the quotation details (Quotation Number, Date, Company, Attention To, Contact, Email, Address).
3. **Add items** - three ways:
   * **Add item box**: type part of a name, pick it, press **+ Add** (or Enter). ▼ lists the whole catalog.
   * **New item...**: an item that is not in the CSV - fill in name, SRP, description, features, pictures...
     Tick *Also save this item to the catalog CSV* to keep it for next time.
   * **Load a template** (below).
4. **Edit the lines**: double-click *Qty*, *Item*, *SRP* or *Total*; tick the *Optional* checkbox;
   double-click the **#** cell (or press **Details...**) to edit the description, key features, specifications,
   in-the-box list and pictures of that line. Edits only affect this quotation, never the CSV.
   * Total = Qty x SRP unless you type your own Total (marked `*`).
   * In *Item*, type `\n` for a line break.
5. **Extra order summaries** (optional): the first ORDER SUMMARY table always lists *every* item, optional ones
   included. Press **Add...** next to *Extra summaries* to add another table with only the items you tick
   (e.g. the drone + 5 batteries, without the optional accessories). Give each one a banner title and, if you
   like, its own discounted price. Add as many as you need; **Edit...** / **Remove** manage them. If you remove
   an item from the quotation it disappears from the extra summaries too.
6. **Package title** (e.g. `MATRICE 4 THERMAL`) - shown in the ORDER SUMMARY banner as
   "MATRICE 4 THERMAL PACKAGE" and used in the file name.
7. Optional: tick **Add a Total Discounted Price** and enter the amount (off by default).
8. **Generate Quotation** -> three files, all named
   `Q-26-626 - 20261002 - MATRICE 4 THERMAL - PNP Mindoro`
   (quotation number - date as YYYYMMDD - package title - company name, or *Attention To* when there is no company;
   blank parts are skipped; change the pattern with `FILENAME_PARTS` in `quotation_config.py`):
   * `.docx` - the Word quotation
   * `.pdf` - the PDF quotation
   * `.qgen` - the **editable file**: every input (details, items and your edits, package title, discount, extra
     summaries, the wording used) plus what was generated (file names, totals) and a copy of the pictures.
     Press **Open saved quotation...** (top right) to load it and carry on editing, then generate again.
     If the terms/bank wording changed since it was saved, you are asked which wording to use.

## Templates
A template = package title + a ready-made list of items (with quantities, prices, optional ticks, edits).
* **Save as...** stores what is in the table now.   **Load** replaces the table with the template.
* To edit a template: **Load** it, change the lines / title, press **Update**.   **Rename** / **Delete** are there too.
* Templates keep the items exactly as saved - if a price changes in the CSV later, edit it in the template and press **Update**.
* A sample "Matrice 4 Thermal Package" (1 drone + 2 batteries) is included. Stored in `templates.json`.

## products.csv = your catalog
One row per item you sell. Items only START with the CSV values. Press **Reload** after saving the CSV.
Columns: `quantity, name, summary_name, srp, images, description, key_features, specifications, in_the_box`
* `quantity` = starting quantity (1 if blank).  *Optional is no longer a CSV column* - tick it in the app
  (an old `optional` column is simply ignored).
* `name`: `\n` = line break. ORDER SUMMARY shows the first line of the name (or `summary_name` if filled).
* `images`: file names inside `images/`, separated by `;`.   Lists: separate bullets with ` | `.

## Pictures
Any size or format works; every picture is fitted into the same 2.3" x 1.8" box and compressed
(`IMAGE_BOX_W_IN`, `IMAGE_BOX_H_IN` in `generate_quotation.py`). Pictures added in the app are copied into `images/`.

## Layout settings (generate_quotation.py)
`FONT`/`SIZE` = Arial 10 everywhere.  `HEADER_DISTANCE` = gap above the logo (1.0" from the top of the page).  The body start (`MARGIN_TOP`) follows
automatically (`HEADER_DISTANCE + LETTERHEAD_HEIGHT_IN`), so the gap under the letterhead stays the same.

## Other files next to the exe
| File | Purpose |
|---|---|
| `terms_and_company.json` | terms, delivery table, payment/warranty, bank details, letterhead lines, Prepared-by (button in *Settings*). Markup: `**bold**` `~~italic~~` `__underline__` `^^red^^` `@@link@@` |
| `app_settings.json` | remembered folders/options (automatic) |

PDF creation uses Microsoft Word if installed (Windows/Mac), otherwise free LibreOffice. The Word file is always created.

## Mac (MacBook) version
The app also runs as a Mac app. A Mac app must be built **on a Mac** (or by GitHub), and an Apple-Silicon build
(M1/M2/M3/M4) does not run on an older Intel MacBook, so build for the Mac type you have.

**Build it on your MacBook**
1. Install Python from https://www.python.org/downloads/macos/ (the Apple/Xcode Python will not work well).
2. Unzip this folder, then double-click **`build_mac.command`**
   (first time: right-click it -> Open -> Open, because macOS blocks scripts from the internet.
   If it says it has no permission, open Terminal, type `chmod +x ` and drag the file onto the window, press Enter, then retry.)
3. Wait a few minutes. The folder `release` opens with **KeithsQuotationGenerator.app** - drag it to *Applications*.

**Or let GitHub build it** (no Python needed): upload this folder to a GitHub repo, open the *Actions* tab ->
*Build apps* -> *Run workflow*, then download *Mac-AppleSilicon* or *Mac-Intel* from the finished run.

**First launch:** an app that was downloaded is unsigned, so macOS says it cannot verify the developer.
Right-click the app -> *Open* -> *Open* (only needed once). If it says the app is "damaged", open Terminal and run
`xattr -dr com.apple.quarantine /Applications/KeithsQuotationGenerator.app`.

**Where your files are on a Mac:** `~/Documents/Keiths Quotation Generator/` - `products.csv`, `images/`,
`templates.json`, `terms_and_company.json`, and the `output/` folder with the exported files. They are created on
first launch (Settings tab -> *Open folder* takes you there). Add your own pictures/CSV there.

**PDF on a Mac:** the PDF is made with Microsoft Word if it is installed (macOS asks once to let the app control
Word - click OK), otherwise with free LibreOffice (https://www.libreoffice.org). With neither, you still get the
Word file and the `.qgen` file; Pages cannot be used for the PDF step.

## Building the .exe (Windows)
* Install Python 3.10+ (tick "Add to PATH"), then double-click `build_exe.bat`. Result: `release\KeithsQuotationGenerator.exe`.
* No Python on the PC? Upload this folder to GitHub and run the "Build Windows EXE" action (`.github/workflows`).

## Code map
* `quotation_app.py` - the window (SearchPicker, ItemDialog, QuotationApp)
* `generate_quotation.py` - builds the document (`generate()` is the single entry point)
* `template_store.py` - reads/writes `templates.json`
* `quotation_config.py` - default wording / paths / file-name pattern
* `build_exe.py` / `build_exe.bat` / `build_mac.command` - build the Windows .exe or Mac .app
