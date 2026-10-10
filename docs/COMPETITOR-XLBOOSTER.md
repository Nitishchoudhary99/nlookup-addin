# DocBrisk vs XLBooster

XLBooster (xlbooster.com, by Satish Dhawale / SkillCourse) is an Indian Excel add-in for "Excel Operators, MIS
Executives and Data Analysts". The data below comes from its website (home, features and "What's new" pages,
read on 10-Oct-2026) and its YouTube and social posts. Their claims and testimonials are their own marketing and
were not checked.

## 1. What XLBooster sells

| | XLBooster |
|---|---|
| Price | ₹2,399 a year (shown as "₹4,788" struck through), ₹9,999 lifetime (₹19,999 struck through), GST included. A pop-up gives ₹1,400 off "for the next 15 minutes" |
| Free plan | Basic plan: the locked tools show a lock and cannot be used at all |
| Technology | VBA `.xlam` add-in (they say it gets past office laptop restrictions); Windows desktop Excel only |
| Licence | One PC at a time, online activation, internet needed to install |
| Positioning | "India's first AI-powered Excel automation tool", "save 100 minutes a day" |

Their tools: Workbook Index, Fast Navigation, Focus Cell, Ask AI / AI Assistant, AI Tools, All AI Formulas (20+),
All XL Formulas, Highlight Borders, Validation Tools (dropdowns, email validation), Data Utilities, Smart Data Cleaner,
Sheet Manager, Manage Files / Files Manager (batch rename), Auto File Organizer, Export Reports (PDF / Excel),
Smart Data Bulk Exporter (filter by unique values, export each as PDF / Excel / sheet), Smart Data Manager,
Protection Tools, Export as Web Table, Smart Chart (K / M labels, publish to their public GitHub for 30 days),
Smart Dashboard, Email Automation (mail merge with PDF), Excel to Python / SQL, Swap Column.

## 2. Tool by tool

✅ = we have it, ➕ = added in DocBrisk 3.3, ⭐ = we are clearly ahead

| Need | XLBooster | DocBrisk 3.3 |
|---|---|---|
| Sheet index | Workbook Index | ➕ Workbook Index, rebuilt **in place** (links keep working), hidden sheets, sizes, tab colours, optional "Index" button on every sheet |
| Jump between sheets | Fast Navigation | ➕ Sheet Navigator, **Ctrl+Shift+J**, search across **all open workbooks**, hidden sheets, "Back" to the last sheet |
| Focus Cell | Focus Cell | ➕⭐ Focus Cell as a see-through overlay: **does not change the file and does not clear Undo** (highlighting done by VBA usually changes the sheet on every move, and any change made by a macro clears Excel's Undo; worth checking theirs in a demo) |
| Sheet Manager | Sheet Manager | ➕ Sheet Manager: show / hide / very hide / rename (prefix, suffix, replace) / tab colour / move / sort / delete with backup |
| Validation | Validation Tools, email validation | ➕⭐ Smart Validation: dropdowns (typed or from a column, kept sorted on a hidden sheet), no repeats, number and date ranges, **mobile, PIN, PAN, GSTIN, IFSC, Aadhaar**, email; circles the old cells that break the rule |
| Formulas library | All XL Formulas | ➕ Formula Library: 41 recipes that **ask for your cells and write the formula** (two-condition lookup, SUMIFS between dates, overdue days, GST, EMI, CAGR ...) |
| AI formulas | Ask AI, AI Formulas | ➕ AI Formula with Claude (Anthropic): your own words → formula for your columns + explanation. **You see exactly what is sent**; only headers by default |
| Lookups | "Fetch data without VLOOKUP" video | ⭐ Smart Lookup + `=SMARTLOOKUP`: matches through spaces, capitals, numbers stored as text; says how many rows VLOOKUP would miss; fuzzy option |
| Fuzzy matching | – | ⭐ Fuzzy Merge, Fuzzy Duplicates, FUZZYLOOKUP, SIMILARITY |
| Reconciliation | – | ⭐ Reconcile (books vs statement / GSTR-2B with tolerance), Sum Finder, Duplicate Payments, Compare by Key |
| Data cleaning | Data Utilities, Smart Data Cleaner | ✅⭐ Smart Clean (SAP-style numbers and dates), Report to Table, Text Tools, Data Health Scan |
| Bulk export by value | Smart Data Bulk Exporter | ✅➕ Split to Sheets / Files, **Split to PDFs** (header on every page, fit to width), Sheets to PDFs |
| Email | Email Automation (mail merge + PDF) | ✅ Mail to Each Party (table in mail, Excel or PDF, drafts or send, log), Email This Sheet |
| Files | Files Manager, Auto File Organizer | ➕ File Organizer: list files in a sheet, **rename from the sheet**, sort into folders by type / month / name, every change logged |
| Web table | Export as Web Table | ➕⭐ Web Table: search, sort, a filter under every column (`>200000` works), totals of the rows shown, CSV, print, dark mode, phone. Stays on your PC (XLBooster's chart publishing puts data on a public GitHub repo) |
| Excel to code | Excel to Python / SQL | ➕ Table to SQL / Python: SQL Server, MySQL, PostgreSQL, SQLite or pandas, with column types |
| Dashboards | Smart Dashboard, Smart Chart | ⭐ Free Dashboard Studio: 78 designs, 26 themes, 12 chart kinds, slicers, Edit / Refresh, Web Dashboard, BI export, L / Cr units |
| Analysis | AI analysis (not described) | ⭐ Smart Analyst, Why It Changed, Forecast, Ask Your Data, ABC, Ageing, Outliers |
| Formula help | – | ⭐ Explain Formula, Formula Risk Scan, Snapshot / What Changed |
| India | – | ⭐ Rupees in words, Indian commas, fiscal year, GSTIN check, 32 worksheet functions (GSTADD, PANVALID, AGEBUCKET ...) |
| Swap columns | Swap Column | ➕ Swap Ranges (formulas included) |
| Borders / formatting | Highlight Borders | ✅ Make Table Neat |
| Protection | Protection Tools | ✅ Protect / Unprotect all sheets |
| Habit and value shown | – | ⭐ My Impact (hours saved, streak, milestones, bonus runs), Pro Week |

## 3. Where DocBrisk is stronger, in one line each (for the website)

1. **Cheaper:** ₹999 a year (₹2.8 a day) against ₹2,399; ₹99 for one month; no countdown pop-ups.
2. **Try before you buy:** a 7-day Pro Week and 3 free runs of every Pro tool, instead of locked buttons.
3. **More free:** dashboards, Smart Lookup, Formula Library, Focus Cell, Navigator, Validation, Web Table are free.
4. **Faster and safer:** a compiled `.xll` (not VBA): big tables in seconds, Undo kept by Focus Cell, backups before changes.
5. **Private:** everything works offline and nothing is uploaded; the one AI tool shows exactly what it sends.
6. **Built for Indian finance work:** reconciliation, GST, fuzzy party names, duplicate payments, ageing, rupees in words.
7. **Shows its value:** "DocBrisk has saved you 18 hours" (My Impact), so renewing is an easy decision.

## 4. Where XLBooster is still ahead, and what to do

| Their edge | What to do |
|---|---|
| AI with no setup (the user does not need an API key) | Add an AI proxy on docbrisk.com that accepts a valid Pro key and calls Claude with your API key, then make AI Formula work for Pro users without their own key. Budget the cost per user per month first. |
| Hindi / Hinglish videos and a large YouTube audience | Make 60-second videos per tool in Hindi and English: Smart Lookup, Focus Cell, Reconcile, Split to PDFs, Web Table. Lead with "VLOOKUP without #N/A". |
| Lifetime plan | Consider a lifetime price (for example ₹4,999) once renewals are measured; it brings cash early but caps revenue per user. |
| Installs on locked office laptops as an `.xlam` | The `.xll` with the per-user installer needs no admin rights either; say so on the download page. |

## 5. Sources

- https://xlbooster.com/ , https://xlbooster.com/features , https://xlbooster.com/whats-new , https://xlbooster.com/about
- https://github.com/SatishDhawale/XLBooster_Free_Setup
- https://www.youtube.com/@xlboosters , XLBooster videos and posts by Satish Dhawale / SkillCourse (YouTube, LinkedIn, Facebook)
- https://learnmoreindia.in/xlbooster-india-first-ai-excel-tool/
