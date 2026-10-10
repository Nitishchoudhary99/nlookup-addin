# DocBrisk (C# / Excel-DNA)   version 3.3

A compiled Excel add-in (.xll): 90 tools and 32 worksheet functions.
It is built the same way as the NLOOKUP project (Excel-DNA, .NET Framework 4.8, C# 7.3).

## What is new in 3.3: everything XLBooster has, done better

The comparison is in `../docs/COMPETITOR-XLBOOSTER.md`. New ribbon groups **Navigate** and **Formulas** sit right after
the DocBrisk group, because they are used every day.

| New | Plan | What it does |
|---|---|---|
| Sheet Navigator | Free | Ctrl+Shift+J: type a few letters, jump to any sheet of any open workbook (hidden ones too); "Back" returns to the last sheet. |
| Workbook Index | Free | Replaces Sheet Index. Rebuilt in place, lists hidden sheets, sizes, last cell, charts, tab colour; optional "Index" button on every sheet. |
| Focus Cell | Free | A soft cross over the active row and column. A see-through, click-through window over the grid: nothing is written to the file, Undo keeps working. On / off, remembered. |
| Sheet Manager | Free | Show, hide, very hide, rename (prefix / suffix / replace), tab colour, move, sort, delete (with backup) many sheets at once. |
| Smart Validation | Free | Dropdowns (typed, or the sorted different values of a column), no repeats, number / date ranges, not-in-future dates, mobile, PIN, PAN, GSTIN, IFSC, Aadhaar, email, text length. Plain Excel validation. |
| Formula Library | Free | 41 recipes; it asks for the cells and writes the formula, then offers to fill down. |
| AI Formula | Free 3 a day / Pro 100 a day | Your own words -> a formula for the active cell, with a short explanation. Goes through docbrisk.com with the user's DocBrisk key, so **no Anthropic key is needed on the PC**; the website checks the key, counts the uses and asks Claude. Sends the request and the column headers; 3 sample rows only if ticked; the window shows what is sent. Without any DocBrisk key it links to the free-key form (name + email). |
| Swap Ranges | Free | Swaps two blocks of the same size, formulas included. |
| Split to PDFs | Pro | One PDF per value of a column, header on every page, fit to width. |
| File Organizer | Pro | List a folder's files in a sheet, rename them from the sheet, or sort them into sub-folders by type / month / first word. Every change is logged to `_DocBrisk_file_log.csv`. |
| Web Table | Free | A table as one HTML page: search, sort, filters per column (`>1000` works on numbers), totals of the rows shown, CSV, print, dark mode, phone friendly. |
| Table to SQL / Python | Pro | CREATE TABLE + INSERT (SQL Server, MySQL, PostgreSQL, SQLite) or a pandas DataFrame, with column types. |

### Build notes for 3.3

- No new NuGet packages. Two Framework references were added: `System.Net.Http` and `System.Web.Extensions` (JSON).
- AI Formula needs the docbrisk.com Worker with the AI proxy (repo `docbrisk-seo`, section "AI FORMULA PROXY" in
  `seo-worker.js`) and the secret `ANTHROPIC_API_KEY` on that Worker. Until then the add-in shows
  "AI Formula is not switched on on docbrisk.com yet".
- For a test Worker, set the registry value `HKCU\Software\DocBrisk\SmartPro\AiProxyUrl` to its https URL.

### Status of version 3.3

- Compiles with warnings as errors; logic tests pass under Mono (ribbon XML valid with all 90 buttons, tips table,
  AI reply parsing, SQL / Python export, formula recipes).
- The Web Table page was run in Chromium (search, filters, sort, totals, theme, phone width, no script errors).
- AI Formula was run end to end: the add-in's own code -> the real Worker in `wrangler dev` -> a stand-in for the
  Anthropic API (Pro, free, daily limit, expired and invalid keys). Not yet against the live Anthropic API.
- **Not run in Excel:** Navigator, Index, Focus Cell (overlay position at different zoom levels and frozen panes),
  Sheet Manager, Validation (especially the PAN / GSTIN formulas), Formula Library, File Organizer, Split to PDFs.
  Test these on Windows before publishing.

## What is new in 3.2: a reason to come back every day

Why users did not stay is written up in `../docs/PAIN-POINTS-AND-GROWTH.md`. In short: many tools, but nothing that
shows the value the user gets, no daily hook, and Pro was only mentioned at the moment a tool locked. 3.2 adds:

| New | Plan | What it does |
|---|---|---|
| Smart Lookup | Free | VLOOKUP without the #N/A. Brings columns from a second table even when keys differ by spaces, non-breaking spaces, capitals, a leading apostrophe or numbers stored as text. A Match column says how each row was found; the summary says how many rows VLOOKUP would have missed, and how many keys appear twice. When keys are not found it estimates how many Fuzzy Merge (Pro) could match. |
| `=SMARTLOOKUP(value, lookup_range, return_range, [if_not_found], [fuzzy])` | Free | The same in a formula. `fuzzy` TRUE takes the closest name (90% alike or more) when nothing matches exactly. |
| My Impact | Free | Time saved (an estimate per tool, at most 3 runs of a tool counted a day), the streak of working days, the tools used most, milestones, and the Pro tools that fit the user's own work. |
| Milestones and bonus runs | Free | 10 milestones (10 runs, 1 hour saved, a 5-day streak ...). Each gives bonus runs that work on any Pro tool after its own 3 free runs are used. |
| Pro Week | - | For the first 7 days on a PC every Pro tool is open, without a key and without using the free runs (a "reverse trial"). Cards on day 1, the day before the end and after the end show what the user did with Pro. Existing users get their Pro Week when they first use 3.2. `ProPlan.WelcomeProDays` = 0 switches it off. |
| Tips after a tool | - | A small card in the corner of the screen (it does not take the focus and closes by itself) shows the Pro tool that takes the job one step further, for example Clean -> Smart Analyst, Compare Two Lists -> Reconcile. At most 2 a day, 3 tool runs apart, the same tip 3 times ever; "No tips" turns them off. On the last free runs of a Pro tool it says what Pro members get. Pro members see, at most once a week, what Pro has saved them. |
| Status bar greeting | - | When Excel starts: the streak, the time saved, the Pro Week day or the bonus runs waiting. |
| DocBrisk Pro window | - | Shows the user's own time saved and what it is worth, the price per day, "Open now" during the Pro Week, and the bonus runs. |

The settings for the offer are in `ProPlan.cs`: `WelcomeProDays` (7), `HourValueRupees` (400: what one hour of office
time is worth, for "your time saved is worth about ..."; 0 hides those lines). The minutes per tool are in `Habit.cs`
(`Minutes`); the tips (which free tool leads to which Pro tool) are in `Nudge.cs` (`Next`), with their limits at the top.

New files: `Habit.cs`, `Nudge.cs`, `ImpactWindow.cs`, `SmartLookup.cs`. Everything is kept in the registry under
`HKCU\Software\DocBrisk\SmartPro`; nothing is sent anywhere.

### Status of version 3.2

- The whole add-in now **compiles** (on Linux with the .NET 8 SDK and the .NET Framework 4.8 reference assemblies:
  `tests/CompileCheck.csproj`, warnings are errors) and GitHub Actions builds it on every push.
- The logic of 3.2 (lookup keys, `SMARTLOOKUP`, time saved, streaks, milestones, bonus runs, Pro Week, the licence
  gate) is **run** by `tests/LogicTests.cs` under Mono: `tests/run-tests.sh`.
- The windows (My Impact, the tip card) and Smart Lookup writing to a sheet have **not been run in Excel** yet.
  Check them once on Windows at 100% and 150% screen scaling before publishing.

## What is new in 3.1

**Keep your own `LicenceKeys.cs`** (the one with your public key X and Y). Every other file can be replaced.

| New | Plan | What it does |
|---|---|---|
| Explain Formula | Free | The formula of the active cell in plain words, step by step, with the value of every part right now, the step where an error starts, and warnings (VLOOKUP without FALSE, typed column numbers, volatile functions, links to other files). |
| DocBrisk Pro window | - | Replaces the old key box. Shows the plan and days left, every Pro tool with its free runs left, the two plans with prices, a Get Pro button that opens docbrisk.com, and the key box. A key copied on the website is picked up from the clipboard and filled in: pay, copy, Activate. |
| Smart Analyst | Pro | One click: a written analysis of a table. Who drives the total (top value, top 3, how many give 80%), latest month against the month before and against last year, best and weakest month, trend, unusual rows, negatives, and what to do next. With tables and charts; the findings can be copied for an email. |
| Duplicate Payments | Pro | Invoices that may have been entered or paid twice. High: same party, amount and invoice number (also when written differently: INV-001 / inv 1; ABC Pvt Ltd / ABC Private Limited). Medium: invoice numbers one character apart. Low: same amount within 7 days, or one invoice number with different amounts. The data is not changed. |
| Mail to Each Party | Pro | One Outlook email for every party with only its own rows, as a table in the email, an Excel file or a PDF. The address comes from the column that holds email addresses. Saved in Drafts to check, or sent at once. A log sheet lists what happened. |

Smaller changes: a free trial run of a Pro tool now shows the runs that are left on the status bar instead of a message box;
the reminder before the plan ends offers to open the renew page; the About box shows the plan; the fix for the build
error CS0136 in `DashLayout.cs` (`units` / `hUnits`) is included.

The smart tools work by fixed rules on the PC. Nothing is sent to DocBrisk or to an AI service, so they also work offline.

### The Pro package

- Free for ever: every dashboard tool (One-Click, Studio, Edit, Web Dashboard, BI Export), Explain Formula, cleaning,
  everyday tools, split / combine, India pack, all worksheet functions.
- Pro: 18 tools (`ProIds` in `Design.cs`). Each has 3 free runs on a PC, then it opens the DocBrisk Pro window.
- Prices shown in the window are in `ProPlan.cs` (`YearRupees` 999, `MonthRupees` 99). Keep them the same as `PRO_PLANS`
  in the website Worker. The five lines of the offer are in `ProPlan.Promise`: keep only lines that are true for your offer.
- The key is the DBK1 key the website issues after payment. It is checked on the PC with the public key; no call to the
  website is made, so a refund does not switch a key off before its expiry day.

## Build (about 5 minutes)

1. Unzip the folder and open `SmartPro.csproj` in Visual Studio 2022.
2. Build > Build Solution. NuGet downloads Excel-DNA on the first build (internet needed).
3. Open `bin\Debug\net48` (and its `publish` subfolder). Use the file ending
   `-AddIn64-packed.xll` for 64-bit Excel, or `-AddIn-packed.xll` for 32-bit Excel.
4. Double-click that file, or in Excel: File > Options > Add-ins > Go > Browse.
5. A DOCBRISK tab appears after Home. Ctrl + Shift + Q opens Quick Find.

Every `.cs` file in the folder is compiled; there is nothing to add to the project file.

## What is new in 3.0: dashboards you design yourself

All of it is in the free group **Dashboards (Free)**.

| Button | What it does |
|---|---|
| One-Click Dashboard | A sensible dashboard from any table in one click, in the look you used last. |
| Dashboard Studio | The designer, with a live preview drawn from your own numbers. |
| Edit Dashboard | Opens the dashboard on the active sheet in the Studio again and updates it in place. |
| Web Dashboard | The dashboard as one HTML file with working filters. On a dashboard sheet it takes that dashboard; on a table it asks whether to design it first. |
| More > Refresh Dashboard | Builds the dashboard again from its table as it is now. The filters you set are kept. |
| More > Dashboard Style | A quick change of theme. |

What can be chosen in Dashboard Studio:

- **Designs**: 78 ready designs (every theme with three card looks and layouts), shown as small pictures of *your* dashboard. "Surprise me" picks one.
- **Cards** (up to 8): the column; total, average, count, largest, smallest or different values; label; format, decimals,
  text in front and behind; colour; compare the latest month with the month before or with the same month last year;
  a small trend line; a target.
- **Charts** (up to 10) of 12 kinds: column, bar, line, area, pie, doughnut, stacked column, stacked bar, 100% stacked,
  combo (columns + line on its own scale), Pareto, table with bars. For each: group by a column or by month, quarter,
  fiscal year, year, weekday or day; the number and how it is added up; colours by a second column; top N and what
  happens to the rest; order; values, % of total or running total; title; data labels; legend; width and height;
  colour; trend line; smooth line.
- **Filters** (up to 6 slicers): on the left, on the right, as a strip on top, or none. Which date column, and the month
  the fiscal year starts.
- **Look**: 26 themes (8 of them dark) or your own colours for every part, 6 card looks, rounded corners, shadows,
  9 fonts, spacing, Indian (L / Cr) or international (K / M / B) units, a currency sign.
- **Page**: title, line under it, footer, your logo, 8 layouts, standard or wide page, cards in a row, key insights,
  hidden row and column headings, print setup.
- **Save design / Open design**: a design is a small `.dbdash` file. Open it on any table that has the same columns.

A click on a card or a chart in the preview opens its settings.

The dashboard in Excel is made of pivot tables, pivot charts and slicers on three sheets (the dashboard and two hidden
helper sheets), so a click on a filter button changes every card and every chart. The design is stored on the dashboard
sheet, so Edit, Refresh and Web Dashboard know what to build. Dashboards made with version 2 are recognised and
converted the first time they are refreshed or edited.

The web page has the same cards, charts, filters and insights. Whoever opens it can click a bar, a slice or a filter
button to filter, switch between the 26 themes and the 6 card looks, change the kind of a chart, save the numbers of a
chart as CSV, and print or save as PDF. It also works on a phone. The file carries the rows it shows (up to 100,000).

## Change the design

- `Design.cs`: `AppName`, `TabLabel`, and the `Tools` list (caption, group, icon, tooltip for every ribbon button).
- `DashThemes.cs`: the themes. One line per theme; add a line to add a theme. Also the fonts and the names of the card looks.
- `DashLook.cs`: font sizes and colour rules shared by Excel, the preview and the web page.
- `DashAuto.cs`: what One-Click Dashboard picks for a table, and the list of ready designs.
- `WebTemplate.cs`: the HTML, the style and the script of the web page (inside the C# text every `"` is written twice).

## Files

| File | What is inside |
|---|---|
| `Design.cs` | names, ribbon buttons, colours of header rows |
| `AddIn.cs` | start-up, ribbon, the tool runner, Quick Find |
| `Xl.cs` | Excel helpers and constants, message boxes, the picker window, backups |
| `Logic.cs` | pure logic: number/date parsing, rupee words, GSTIN, similarity, column profiling |
| `DashSpec.cs` | the design of a dashboard (cards, charts, filters, look) and how it is saved as text |
| `DashThemes.cs` | the 26 themes, fonts, card looks, colour arithmetic |
| `DashData.cs` | the table in memory, adding up for cards and charts, number formats, titles, key insights |
| `DashLayout.cs` | where every card, chart and filter goes on the grid of the sheet; the 8 layouts |
| `DashAuto.cs` | the automatic design for a table, repair of a design for a changed table, the 78 ready designs |
| `DashLook.cs` | sizes and colour rules; what the preview has already worked out |
| `DashRender.cs` | draws the preview of Dashboard Studio |
| `DashPlan.cs` | the plan of a build: helper table, pivot tables, formulas (no Excel calls, so it can be tested alone) |
| `Dashboard.cs` | builds the dashboard in Excel from a plan; Refresh, Edit, Dashboard Style; Data Health Scan |
| `DashboardStudio.cs` | the Dashboard Studio window |
| `WebDash.cs` | the rows and the design as data for the web page |
| `WebTemplate.cs` | the web page itself |
| `ExportTools.cs` | the Web Dashboard button and BI Export (CSV + measures for Power BI / Looker Studio) |
| `LicenceKeys.cs` | your public key and buy link; keep your own copy when you update |
| `Licence.cs` | checks a key offline, free trial runs, the key on the clipboard |
| `ProCentre.cs` | the DocBrisk Pro window |
| `Habit.cs` | time saved, streak, milestones, bonus runs, Pro Week |
| `Nudge.cs` | the tip cards after a tool and their limits |
| `ImpactWindow.cs` | the My Impact window |
| `SmartLookup.cs` | Smart Lookup and the SMARTLOOKUP function |
| `NavTools.cs` | Sheet Navigator, Workbook Index, Sheet Manager, Swap Ranges |
| `FocusCell.cs` | the Focus Cell overlay |
| `ValidationTools.cs` | Smart Validation |
| `FormulaLibrary.cs` | the 41 formula recipes |
| `AiFormula.cs` | AI Formula (through the docbrisk.com proxy) |
| `FileTools.cs` | File Organizer |
| `ExportMore.cs` | Web Table, Table to SQL / Python |
| `ProPlan.cs` | the prices and the words of the Pro offer |
| `SmartTools.cs` | Explain Formula, Smart Analyst, Duplicate Payments, Mail to Each Party |
| `Settings.cs` | saved choices such as the look of the last dashboard |
| `FuzzyMerge.cs` | Fuzzy Merge: joins two tables on nearly-equal names with a similarity %, indexed and multi-threaded |
| `DailyTools.cs` | Sum Finder, Reconcile, Report to Table |
| `Assistant.cs` | Assistant window: suggestions for the current table or dashboard and a searchable list of every tool |
| `InsightTools.cs` | Why It Changed (period comparison with reasons) and Forecast |
| `EverydayTools.cs` | Paste to visible cells, delete rows where, blank row at change, merge same values, join columns, text tools, math on cells, hide formula errors, lock references |
| `WorkbookTools.cs` | Make table neat, quick summary, compare two lists, colour count, values-only copy, unhide everything, protect all, links, file diet, sort sheets, sheets from a list |
| `AskTools.cs` | Ask Your Data: a request in plain words becomes a table and a chart |
| `ProTools.cs` | ABC Analysis, Ageing Report, Formula Risk Scan, Snapshot / What Changed |
| `CleanTools.cs` | Smart Clean, text to numbers/dates, trim, blank rows, fill down, unmerge, formulas to values |
| `AnalyseTools.cs` | Compare by Key, Fuzzy Duplicates, Outliers, Duplicates, Unpivot |
| `BookTools.cs` | split, combine, sheet index, PDFs, India pack, backups, function list |
| `Functions.cs` | RUPEEWORDS, INRFORMAT, INRSHORT, FYEAR, FYQUARTER, GSTINVALID, GSTSTATE, PANVALID, SIMILARITY, FUZZYLOOKUP, EXTRACTNUM, TEXTBETWEEN |

## Status of version 3.1

The four new files of 3.1 (`SmartTools.cs`, `ProCentre.cs`, `ProPlan.cs`, the new `Licence.cs`) were written without a
C# compiler, without Excel and without Outlook: they have **not been compiled and not been run**. Build once and send
the lines of the Error List if there are any.

## Status of version 3.0

The dashboard files were written without a C# compiler and without Excel at hand, so they have **not been compiled
and not been run in Excel yet**. What was done instead:

- every file was checked by a C# type checker written for this project (names, types, overloads, C# 7.3 rules);
- the logic was translated to JavaScript and run: the numbers of cards and charts, the layouts, the build plan, the
  formulas written to the sheet (against a formula evaluator), the whole build against a model of Excel's objects
  (pivot tables, charts, slicers, shapes) including refresh, edit, filters and a dozen ways Excel can refuse a step;
- Dashboard Studio was run on a model of Windows Forms: every control, every button, long random click sequences,
  and the layout of the window at 100% to 200% screen scaling;
- the web page was run in a real browser (Chromium) and its numbers compared with the add-in's own, with and without filters.

A model of Excel is not Excel. Expect a few compile errors or run-time surprises on the first build: copy the lines
from Visual Studio's Error List (or the file `last-error.txt` that the add-in writes) and send them to be fixed.
The 16 files that were not touched (cleaning, analysis, reconcile, licence and so on) are exactly as in version 2.4.

## Pro subscription

`Licence.cs` checks DocBrisk keys (format DBK1, the same keys as NLOOKUP) offline.

1. Paste the public key X and Y into `PubX` and `PubY` in `LicenceKeys.cs` (the same values the NLOOKUP add-in uses).
2. `ProIds` in `Design.cs` lists the Pro tools. Each Pro tool gives 3 free runs per PC, then opens the DocBrisk Pro window.
   Every dashboard tool is free.
3. A Pro key is a key with tier `P` and an expiry day. When the date passes, Pro tools lock again and Free tools keep working.
4. The website issues Pro keys after payment: `issueLicense(env, 'P', name, expiryDay)` in the Worker, where
   `expiryDay` = days since 2026-01-01 for the last paid day.
5. Users enter or renew the key with the DocBrisk Pro button (first group of the tab).
