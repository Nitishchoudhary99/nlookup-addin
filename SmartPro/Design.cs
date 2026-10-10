using System.Security;
using System.Text;

namespace SmartPro
{
    // =====================================================================
    //  DESIGN FILE: everything a rename or restyle needs is in this file.
    //    AppName / TabLabel ... the product name
    //    colours .............. header rows and marks on a sheet (dashboard themes are in DashThemes.cs)
    //    Tools ................ ribbon buttons: caption, group, icon, tip
    // =====================================================================
    internal static class Design
    {
        public const string AppName = "DocBrisk";          // shown in every message box
        public const string TabLabel = "DocBrisk";         // ribbon tab
        public const string Version = "3.2";
        public const string Author = "Nitish Choudhary";     // shown on the ribbon and in About
        public const string FunctionCategory = "DocBrisk"; // category in Insert Function

        // colours of the header row and the marks that the tools write on a sheet
        // (the colours of a dashboard are its theme: see DashThemes.cs)
        public static readonly int Page = Rgb(241, 245, 249);
        public static readonly int Navy = Rgb(15, 23, 42);
        public static readonly int Slate = Rgb(71, 85, 105);
        public static readonly int White = Rgb(255, 255, 255);
        public static readonly int SubtleText = Rgb(203, 213, 225);
        public static readonly int GridLine = Rgb(226, 232, 240);
        public static readonly int[] Accent = { Rgb(37, 99, 235), Rgb(13, 148, 136), Rgb(217, 119, 6), Rgb(124, 58, 237), Rgb(225, 29, 72) };

        // highlight colours used by the tools
        public static readonly int Good = Rgb(187, 247, 208);
        public static readonly int Bad = Rgb(254, 202, 202);
        public static readonly int Changed = Rgb(254, 240, 138);
        public static readonly int Unusual = Rgb(253, 186, 116);

        public static int Rgb(int r, int g, int b) { return r | (g << 8) | (b << 16); }

        // PRO TOOLS: these need a Pro key after the free trial runs. Add or remove ids freely.
        // The prices and the words of the offer are in ProPlan.cs. Every dashboard tool stays free.
        public static readonly string[] ProIds = { "analyst", "variance", "forecast", "ask", "fuzzymerge", "reconcile", "sumfinder", "duppay", "mailparty", "compare", "fuzzy", "abc", "ageing", "audit", "snap", "diff", "splitfiles", "combfolder" };

        public static bool IsPro(string id) { return System.Array.IndexOf(ProIds, id) >= 0; }

        public static string CaptionOf(string id)
        {
            foreach (ToolDef t in Tools) if (t.Id == id) return t.Caption;
            return id;
        }

        // id, caption, group, Office icon, large button?, tooltip, search words (for Quick Find)
        public static readonly ToolDef[] Tools =
        {
            new ToolDef("assist", "Assistant", "DocBrisk", "FunctionWizard", true, "Looks at the table you are in and suggests what to do next, in plain words. Also lists and searches every tool.", "assistant suggest recommend help start what next guide"),
            new ToolDef("smartlookup", "Smart Lookup", "DocBrisk", "FunctionWizard", true, "VLOOKUP without the #N/A: brings columns from another table even when the keys differ by spaces, capitals or numbers stored as text. Every row says how it was matched; keys found twice are counted.", "lookup vlookup xlookup index match bring column join merge na n/a not found master map fetch nlookup"),
            new ToolDef("explain", "Explain Formula", "DocBrisk", "FormulaEvaluate", true, "Explains the formula in the active cell in plain words, step by step, with the value of every part right now, where an error starts and what to watch out for.", "explain formula understand meaning what does step evaluate why error help smart ai"),
            new ToolDef("impact", "My Impact", "DocBrisk", "AcceptInvitation", true, "The time DocBrisk has saved you, your streak of working days and your milestones. Every milestone adds bonus runs for the Pro tools.", "impact time saved hours streak milestones progress stats statistics bonus reward"),
            new ToolDef("licence", "DocBrisk Pro", "DocBrisk", "Lock", true, "Your plan and expiry, what Pro adds and the free runs left for every Pro tool. Paste your key here, or get Pro on docbrisk.com: a key you copy there is picked up by itself.", "licence license key pro activate subscription renew buy plan price upgrade"),
            new ToolDef("dash", "One-Click Dashboard", "Dashboards (Free)", "ChartInsert", true, "KPI cards, charts, filter buttons and written insights from any table, in one click. Change anything afterwards with Edit Dashboard.", "dashboard kpi chart pivot slicer report mis"),
            new ToolDef("custom", "Dashboard Studio", "Dashboards (Free)", "ChartInsert", true, "Design your own dashboard with a live preview: up to 8 KPI cards and 10 charts of 12 kinds, 6 filters, 26 colour themes or your own colours, 6 card looks, 8 layouts and your logo. Start from 78 ready designs; save a design and use it again.", "studio custom manual dashboard design suggest layout style kpi builder theme logo template colours"),
            new ToolDef("editdash", "Edit Dashboard", "Dashboards (Free)", "FormatPainter", true, "Opens the dashboard on this sheet in Dashboard Studio. Change its cards, charts, filters, colours or layout, then update it in place.", "edit change modify dashboard design customise customize cards charts colours layout"),
            new ToolDef("webdash", "Web Dashboard", "Dashboards (Free)", "HyperlinkInsert", true, "Your dashboard as one web page file that opens in any browser, also on a phone. Whoever opens it can filter by clicking, switch between 26 themes, change the kind of a chart and print it.", "web html browser online share dashboard export phone mobile"),
            new ToolDef("refresh", "Refresh Dashboard", "Dashboards (Free)", "Refresh", false, "Builds the open dashboard again from its source table as it is now, with the same design. New rows are picked up and the filters you set are kept.", "refresh update rebuild dashboard new data again", "More", "FormatPainter"),
            new ToolDef("biexport", "BI Export", "Dashboards (Free)", "FileSaveAs", false, "Clean CSV plus measure formulas and steps for Power BI and Looker Studio.", "power bi powerbi looker studio export csv dax google", "More", "FormatPainter"),
            new ToolDef("style", "Dashboard Style", "Dashboards (Free)", "FormatPainter", false, "A quick change of colours: pick one of 26 themes for the open dashboard, or the theme the next dashboards start with.", "style theme colour color look design dashboard", "More", "FormatPainter"),
            new ToolDef("analyst", "Smart Analyst", "Insights (Pro)", "SparklineLineInsert", true, "One click: a written analysis of your table. Who drives the total, the trend by month, unusual rows and what to do next, with tables and charts. Ready to paste into an email.", "smart analyst analysis auto insights summary commentary management report findings ai explain data"),
            new ToolDef("variance", "Why It Changed", "Insights (Pro)", "SortDialog", true, "Compares two periods and shows which customers, products or plants caused the rise or fall, with a chart and a plain-English summary.", "variance why changed compare months movers increase decrease analysis mom yoy reason"),
            new ToolDef("forecast", "Forecast", "Insights (Pro)", "ChartInsert", true, "Forecasts the next 3, 6 or 12 months from your monthly history, with a likely range and a chart.", "forecast predict next months trend projection plan future"),
            new ToolDef("ask", "Ask Your Data", "Insights (Pro)", "ReviewNewComment", true, "Type what you want, for example: amount by customer top 10. You get the table and a chart.", "ask question plain words chart query natural language"),
            new ToolDef("fuzzymerge", "Fuzzy Merge", "Match and Reconcile (Pro)", "MergeCenter", true, "Joins two tables on names that are nearly the same, with a similarity % for every row. Works in every Excel version.", "fuzzy merge match join lookup similar percentage power query vendor customer master"),
            new ToolDef("reconcile", "Reconcile", "Match and Reconcile (Pro)", "WindowSideBySide", true, "Matches two lists by reference and amount with a rounding tolerance: books vs party statement, bank or GST portal data.", "reconcile reconciliation match ledger bank statement gstr 2a 2b party books difference"),
            new ToolDef("sumfinder", "Sum Finder", "Match and Reconcile (Pro)", "AutoSum", true, "Which of these amounts add up to this total? Finds the invoices behind a payment and marks them.", "sum finder combination subset total payment invoices match add up"),
            new ToolDef("duppay", "Duplicate Payments", "Match and Reconcile (Pro)", "RemoveDuplicates", true, "Finds invoices that may have been entered or paid twice: same party and amount with the same, a nearly-same or a differently-written invoice number, or within a few days. Your data is not changed.", "duplicate payment invoice double paid twice audit vendor bill check"),
            new ToolDef("compare", "Compare by Key", "Match and Reconcile (Pro)", "WindowSideBySide", false, "Matches two tables by ID and lists added, removed and changed records.", "compare difference reconcile match old new", "Compare", "Spelling"),
            new ToolDef("fuzzy", "Fuzzy Duplicates", "Match and Reconcile (Pro)", "Spelling", false, "Finds look-alike names: typos, Pvt/Ltd, spacing, case.", "fuzzy similar duplicate vendor customer names typo", "Compare", "Spelling"),
            new ToolDef("snap", "Take Snapshot", "Match and Reconcile (Pro)", "FileSave", false, "Remembers every cell of this workbook so you can see what changed later.", "snapshot remember version time machine", "Compare", "Spelling"),
            new ToolDef("diff", "What Changed?", "Match and Reconcile (Pro)", "WindowSideBySide", false, "Lists every cell that changed since the snapshot.", "changed difference since snapshot track version", "Compare", "Spelling"),
            new ToolDef("mailparty", "Mail to Each Party", "Send (Pro)", "FileSendAsAttachment", true, "One Outlook email for every party with only its own rows: as a table in the email, an Excel file or a PDF. Saved in Drafts to check, or sent at once. Needs the Outlook desktop app.", "mail email outlook send each party customer vendor statement mail merge bulk reminder outstanding pdf"),
            new ToolDef("clean", "Smart Clean", "Clean", "ClearFormats", true, "Trims text and turns text numbers and text dates (SAP style) into real ones.", "clean fix trim numbers dates sap"),
            new ToolDef("report", "Report to Table", "Clean", "SortDialog", true, "Turns a printed-style report (repeated headers, totals, separator lines, gaps) into one clean table on a new sheet.", "report table sap clean repeated header subtotal flatten export"),
            new ToolDef("tonum", "Text to Numbers", "Clean", "CommaStyle", false, "Converts 1,234.50- / (1,234) / Rs. 500 / 12% text into numbers.", "convert text number sap minus", "Quick Fixes", "FillDown"),
            new ToolDef("todate", "Text to Dates", "Clean", "DateAndTimeInsert", false, "Converts 01.04.2025 / 01-Apr-25 / 2025-04-01 text into dates.", "convert text date sap", "Quick Fixes", "FillDown"),
            new ToolDef("trim", "Trim Spaces", "Clean", "Cut", false, "Removes leading, trailing, double and non-breaking spaces.", "trim spaces", "Quick Fixes", "FillDown"),
            new ToolDef("blankrows", "Remove Blank Rows", "Clean", "SheetRowsDelete", false, "Deletes rows that are completely empty inside the table.", "delete blank empty rows", "Quick Fixes", "FillDown"),
            new ToolDef("filldown", "Fill Blanks Down", "Clean", "FillDown", false, "Blank cells take the value above them.", "fill blanks down", "Quick Fixes", "FillDown"),
            new ToolDef("unmerge", "Unmerge and Fill", "Clean", "UnmergeCells", false, "Unmerges cells and repeats the value in every cell.", "unmerge merged cells fill", "Quick Fixes", "FillDown"),
            new ToolDef("tovalues", "Formulas to Values", "Clean", "PasteValues", false, "Replaces formulas with their results.", "formula values paste freeze", "Quick Fixes", "FillDown"),
            new ToolDef("health", "Data Health Scan", "Clean", "AcceptInvitation", true, "Scores the table out of 100 and lists what to fix in each column.", "health profile quality check score"),
            new ToolDef("neat", "Make Table Neat", "Everyday", "FormatPainter", true, "One click: styled header, thin borders, number formats, column widths, filter buttons and a frozen header row.", "neat format beautify table header borders autofit freeze filter tidy"),
            new ToolDef("summary", "Quick Summary", "Everyday", "SortDialog", true, "Pick a column to group by and a number to add up: rows, total, average, smallest and largest for each group, without building a pivot table.", "summary summarise group total count average pivot quick subtotal"),
            new ToolDef("listcompare", "Compare Two Lists", "Everyday", "WindowSideBySide", true, "Shows what is only in the first list, only in the second list, and in both.", "compare lists missing common difference match two columns"),
            new ToolDef("pastevisible", "Paste to Visible Cells", "Everyday", "PasteValues", false, "Pastes into filtered data: hidden and filtered-out rows are skipped.", "paste visible filtered cells skip hidden rows", "Rows and Cells", "SheetRowsInsert"),
            new ToolDef("delrows", "Delete Rows Where", "Everyday", "SheetRowsDelete", false, "Deletes every row where a column contains a text, is blank, is zero, or is below or above a number.", "delete rows condition contains blank zero remove where", "Rows and Cells", "SheetRowsInsert"),
            new ToolDef("blankchange", "Blank Row at Each Change", "Everyday", "SheetRowsInsert", false, "Adds an empty row whenever the value in a column changes, to separate groups in a report.", "blank row insert change group separate", "Rows and Cells", "SheetRowsInsert"),
            new ToolDef("mergesame", "Merge Same Values", "Everyday", "MergeCenter", false, "Merges equal cells that sit under each other, for a printed report.", "merge same equal cells report group", "Rows and Cells", "SheetRowsInsert"),
            new ToolDef("joincols", "Join Columns", "Everyday", "MergeCenter", false, "Joins two or more columns into one, with a separator of your choice.", "join concatenate combine columns text separator", "Rows and Cells", "SheetRowsInsert"),
            new ToolDef("texttools", "Text Tools", "Everyday", "Spelling", false, "Change case, add text in front or at the end, remove first or last characters, keep only digits or letters, remove spaces and line breaks.", "text upper lower proper case prefix suffix remove characters digits letters spaces line breaks", "Text and Formulas", "FunctionWizard"),
            new ToolDef("mathcells", "Math on Cells", "Everyday", "AutoSum", false, "Changes the numbers themselves: *1.18 adds GST, /100000 shows lakhs, +500, -10%. Formulas are extended, not replaced.", "math multiply divide add subtract percent gst lakhs numbers change", "Text and Formulas", "FunctionWizard"),
            new ToolDef("iferror", "Hide Formula Errors", "Everyday", "FunctionWizard", false, "Wraps the selected formulas in IFERROR so #N/A and #DIV/0! show as blank, 0 or your own text.", "iferror error na div hide wrap formula", "Text and Formulas", "FunctionWizard"),
            new ToolDef("dollar", "Lock References ($)", "Everyday", "Lock", false, "Changes the references in the selected formulas to $A$1, A$1, $A1 or A1 in one go.", "dollar absolute relative lock reference f4 formula", "Text and Formulas", "FunctionWizard"),
            new ToolDef("abc", "ABC Analysis", "Analyse", "SortDialog", false, "Classes items A / B / C by value with a Pareto chart.", "abc pareto 80 20 inventory class", "Reports", "Filter"),
            new ToolDef("ageing", "Ageing Report", "Analyse", "DateAndTimeInsert", false, "Outstanding by party in 0-30, 31-60, 61-90, 91-180 and 180+ day buckets.", "ageing aging receivable payable outstanding stock days", "Reports", "Filter"),
            new ToolDef("audit", "Formula Risk Scan", "Analyse", "ErrorChecking", false, "Finds broken, inconsistent, overwritten and risky formulas on this sheet.", "formula audit risk error inconsistent hardcoded link volatile", "Reports", "Filter"),
            new ToolDef("outliers", "Find Outliers", "Analyse", "Filter", false, "Highlights numbers far outside the normal range.", "outlier unusual anomaly", "Reports", "Filter"),
            new ToolDef("dups", "Highlight Duplicates", "Analyse", "RemoveDuplicates", false, "Colours every value that appears more than once.", "duplicate repeated highlight", "Reports", "Filter"),
            new ToolDef("unpivot", "Unpivot Table", "Analyse", "Refresh", false, "Turns a cross-tab (months across) into a flat list.", "unpivot flatten crosstab", "Reports", "Filter"),
            new ToolDef("colourcount", "Colour Count", "Analyse", "Filter", false, "Counts and adds up cells by their fill colour, including colours from conditional formatting.", "colour color count sum fill highlighted cells", "Reports", "Filter"),
            new ToolDef("splitsheets", "Split to Sheets", "Sheets and Files", "SheetInsert", false, "One sheet for each value of a column.", "split column sheets", "Split and Combine", "Consolidate"),
            new ToolDef("splitfiles", "Split to Files", "Sheets and Files", "FileSaveAs", false, "One workbook for each value of a column.", "split column files workbooks", "Split and Combine", "Consolidate"),
            new ToolDef("combsheets", "Combine Sheets", "Sheets and Files", "Consolidate", false, "Stacks all sheets into one, matching columns by header name.", "combine merge append sheets", "Split and Combine", "Consolidate"),
            new ToolDef("combfolder", "Combine Folder", "Sheets and Files", "FileOpen", false, "Stacks every Excel / CSV file in a folder into one sheet.", "combine merge append files folder csv", "Split and Combine", "Consolidate"),
            new ToolDef("index", "Sheet Index", "Sheets and Files", "HyperlinkInsert", false, "A clickable list of all sheets.", "index contents hyperlink sheets", "Split and Combine", "Consolidate"),
            new ToolDef("pdfs", "Sheets to PDFs", "Sheets and Files", "FileSaveAsPdfOrXps", false, "Saves every visible sheet as its own PDF.", "pdf export print sheets", "Split and Combine", "Consolidate"),
            new ToolDef("emailsheet", "Email This Sheet", "Sheets and Files", "FileSave", false, "Opens a new Outlook email with this sheet attached as PDF, as an Excel file with values only, or both. Needs the Outlook desktop app.", "email mail send outlook attach pdf sheet report", "Workbook", "FileSave"),
            new ToolDef("unhideall", "Unhide Everything", "Sheets and Files", "Filter", false, "Shows every hidden sheet, row and column in the workbook.", "unhide hidden sheets rows columns show all", "Workbook", "FileSave"),
            new ToolDef("protectall", "Protect All Sheets", "Sheets and Files", "Lock", false, "Protects every sheet with one password in one go.", "protect all sheets password lock", "Workbook", "FileSave"),
            new ToolDef("unprotectall", "Unprotect All Sheets", "Sheets and Files", "Lock", false, "Removes the protection from every sheet in one go.", "unprotect all sheets password unlock", "Workbook", "FileSave"),
            new ToolDef("links", "Links to Other Files", "Sheets and Files", "HyperlinkInsert", false, "Lists the other workbooks this file reads from and can break those links, keeping the values.", "links external workbook break update source", "Workbook", "FileSave"),
            new ToolDef("diet", "File Diet", "Sheets and Files", "ClearFormats", false, "Finds formatted but empty rows and columns that make a file big and slow, and removes them.", "file size big slow slim diet reduce unused rows used range", "Workbook", "FileSave"),
            new ToolDef("sortsheets", "Sort Sheet Tabs", "Sheets and Files", "SortDialog", false, "Puts the sheet tabs in A to Z order.", "sort sheets tabs order alphabetical", "Workbook", "FileSave"),
            new ToolDef("sheetsfromlist", "Sheets from a List", "Sheets and Files", "SheetInsert", false, "Creates one new sheet for every name in the selected cells.", "create sheets list names new many", "Workbook", "FileSave"),
            new ToolDef("valuescopy", "Values-Only Copy", "Sheets and Files", "PasteValues", false, "Makes a copy of the sheet or the whole workbook with values only: no formulas, no links. Safe to email.", "values only copy send email without formulas share", "Workbook", "FileSave"),
            new ToolDef("inrfmt", "Indian Commas", "India", "AccountingFormat", false, "Shows numbers as 12,34,567.", "lakh crore indian comma format rupee", "India Tools", "AccountingFormat"),
            new ToolDef("fycol", "Add Fiscal Year", "India", "SheetColumnsInsert", false, "Adds Fiscal Year (April-March) and Quarter columns next to a date column.", "fiscal financial year quarter fy april", "India Tools", "AccountingFormat"),
            new ToolDef("words", "Rupees in Words", "India", "FunctionWizard", false, "Writes the amount in words in the cell to the right.", "rupees words amount", "India Tools", "AccountingFormat"),
            new ToolDef("gstin", "Check GSTIN", "India", "DataValidation", false, "Validates format, state code and check digit.", "gst gstin validate", "India Tools", "AccountingFormat"),
            new ToolDef("repeat", "Repeat Last Tool", "DocBrisk by Nitish Choudhary", "Refresh", false, "Runs the tool you used last once more, for example on the next sheet.", "repeat last again redo tool", "Help", "Help"),
            new ToolDef("find", "Quick Find", "DocBrisk by Nitish Choudhary", "FindDialog", false, "Type what you want to do. Shortcut: Ctrl+Shift+Q.", "search find tool", "Help", "Help"),
            new ToolDef("funcs", "Function List", "DocBrisk by Nitish Choudhary", "Help", false, "Lists the worksheet functions with examples.", "functions formulas help", "Help", "Help"),
            new ToolDef("backup", "Backup Now", "DocBrisk by Nitish Choudhary", "FileSaveAs", false, "Saves a time-stamped copy of this workbook.", "backup save copy", "Help", "Help"),
            new ToolDef("openbackups", "Open Backups", "DocBrisk by Nitish Choudhary", "FileOpen", false, "Opens the folder with the automatic backups.", "backup folder restore undo", "Help", "Help"),
            new ToolDef("about", "About", "DocBrisk by Nitish Choudhary", "Info", false, "Version and shortcuts.", "about version", "Help", "Help")
        };

        static string Esc(string s) { return SecurityElement.Escape(s); }

        // Big buttons for the main jobs; related tools sit in drop-down menus that show a description under each name.
        public static string RibbonXml()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs>");
            sb.Append("<tab id='spTab' label='").Append(Esc(TabLabel)).Append("' insertAfterMso='TabHome'>");
            string grp = null, menu = null;
            int g = 0, m = 0;
            foreach (ToolDef t in Tools)
            {
                if (t.Group != grp)
                {
                    if (menu != null) sb.Append("</menu>");
                    menu = null;
                    if (grp != null) sb.Append("</group>");
                    g++;
                    grp = t.Group;
                    sb.Append("<group id='spG").Append(g).Append("' label='").Append(Esc(grp)).Append("'>");
                }
                if (t.Menu != menu)
                {
                    if (menu != null) sb.Append("</menu>");
                    menu = t.Menu;
                    if (menu != null)
                    {
                        m++;
                        sb.Append("<menu id='spM").Append(m).Append("' label='").Append(Esc(menu))
                          .Append("' imageMso='").Append(t.MenuIcon ?? t.Icon)
                          .Append("' size='large' itemSize='large' screentip='").Append(Esc(menu)).Append("'>");
                    }
                }
                string pro = IsPro(t.Id) ? "  (Pro)" : "";
                sb.Append("<button id='").Append(t.Id).Append("' imageMso='").Append(t.Icon);
                if (menu != null)
                {
                    // inside a menu: the name with the description under it
                    sb.Append("' label='").Append(Esc(t.Caption + pro)).Append("' description='").Append(Esc(t.Tip));
                }
                else
                {
                    sb.Append("' label='").Append(Esc(t.Caption))
                      .Append("' size='").Append(t.Large ? "large" : "normal")
                      .Append("' screentip='").Append(Esc(t.Caption + pro))
                      .Append("' supertip='").Append(Esc(t.Tip + (pro.Length > 0 ? "  [Pro tool: " + Licence.FreeRuns + " free runs to try on every PC]" : "")));
                }
                sb.Append("' onAction='OnTool'/>");
            }
            if (menu != null) sb.Append("</menu>");
            if (grp != null) sb.Append("</group>");
            sb.Append("</tab></tabs></ribbon></customUI>");
            return sb.ToString();
        }
    }

    internal sealed class ToolDef
    {
        public readonly string Id, Caption, Group, Icon, Tip, Keys;
        public readonly string Menu, MenuIcon;        // set when the tool sits inside a drop-down menu
        public readonly bool Large;

        public ToolDef(string id, string caption, string group, string icon, bool large, string tip, string keys)
            : this(id, caption, group, icon, large, tip, keys, null, null)
        {
        }

        public ToolDef(string id, string caption, string group, string icon, bool large, string tip, string keys, string menu, string menuIcon)
        {
            Id = id; Caption = caption; Group = group; Icon = icon; Large = large; Tip = tip; Keys = keys;
            Menu = menu; MenuIcon = menuIcon;
        }
    }
}
