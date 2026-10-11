using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SmartPro
{
    // ASSISTANT: looks at the table the user is in and says what to do next, in plain words.
    // The same window lists every tool with a search box, so nobody has to hunt through the ribbon.
    internal static class Assistant
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public static void Show()
        {
            List<string[]> recs = new List<string[]>();          // tool id, what to do, why
            string about = "Click a cell inside your data and open the Assistant again to get suggestions.";
            try { about = Analyse(recs); } catch { }
            string run;
            using (AssistantForm f = new AssistantForm(about, recs))
            {
                f.ShowDialog(new ExcelWin());
                run = f.RunId;
            }
            if (run != null) Runner.Run(run);
        }

        static void Add(List<string[]> recs, string id, string why)
        {
            recs.Add(new[] { id, Design.CaptionOf(id) + (Design.IsPro(id) ? "  (Pro)" : ""), why });
        }

        static string N(int n) { return n.ToString("#,##0", Inv); }

        // fills recs and returns one line that says what was looked at
        static string Analyse(List<string[]> recs)
        {
            dynamic app = Xl.App;
            if (app.ActiveWorkbook == null) return "Open a workbook to get suggestions.";
            // on a dashboard sheet: what can be done with the dashboard
            DashInfo dash = null;
            try { dash = Dashboard.ReadInfo(app.ActiveSheet); } catch { }
            if (dash != null)
            {
                Add(recs, "editdash", "Change the cards, charts, filters, colours or layout of this dashboard, with a live preview.");
                Add(recs, "refresh", "Build it again from its table as it is now. New rows are picked up; the filters you set are kept.");
                Add(recs, "webdash", "The same dashboard as one web page file that opens on any PC or phone, without Excel.");
                Add(recs, "style", "A quick change of colours: one of 26 themes.");
                return "This sheet is the dashboard '" + dash.Name + "', made from the sheet '" + dash.SourceSheet + "'.";
            }
            dynamic rng;
            try
            {
                dynamic sel = app.Selection;
                long cnt = Xl.CellCount(sel);
                rng = cnt > 1 ? sel.Areas[1] : sel.CurrentRegion;
                rng = app.Intersect(rng, rng.Worksheet.UsedRange);
            }
            catch { return "Click a cell inside your data to get suggestions."; }
            if (rng == null) return "Click a cell inside your data to get suggestions.";
            int rows = (int)rng.Rows.Count, cols = (int)rng.Columns.Count;
            if (rows < 2) return "Click a cell inside a table with a header row and data rows to get suggestions.";
            bool sampled = rows > 20001;
            dynamic part = sampled ? rng.Resize[20001, cols] : rng;
            object[,] v = Xl.ReadTyped(part);
            bool merged = false;
            try
            {
                object mc = part.MergeCells;                 // false, true, or "mixed" when only some cells are merged
                merged = !(mc is bool) || (bool)mc;
            }
            catch { }

            int nr = v.GetLength(0), nc = v.GetLength(1);
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 20000);
            int textNum = 0, textDate = 0, spaces = 0, errors = 0, nNum = 0, nCat = 0, nDate = 0;
            string nameCol = null, firstCat = null, firstNum = null;
            for (int c = 1; c <= nc; c++)
            {
                string k = ci[c].Kind;
                if (k == "num")
                {
                    nNum++;
                    textNum += ci[c].TextNum;
                    if (firstNum == null) firstNum = ci[c].Hdr;
                }
                else if (k == "date") nDate++;
                else if (k == "cat")
                {
                    nCat++;
                    if (firstCat == null) firstCat = ci[c].Hdr;
                }
                textDate += ci[c].TextDate;
                spaces += ci[c].Spaces;
                errors += ci[c].Errors;
                if (nameCol == null && (k == "cat" || k == "text") && ci[c].Distinct >= 40) nameCol = ci[c].Hdr;
            }
            int blankRows = 0, repeatHead = 0, totals = 0;
            for (int r = hr + 1; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r))
                {
                    blankRows++;
                    continue;
                }
                if (!Profiler.RowIsData(v, r, nc))
                {
                    totals++;
                    continue;
                }
                int same = 0, filled = 0;
                for (int c = 1; c <= nc; c++)
                {
                    if (Xl.IsBlank(v[r, c])) continue;
                    filled++;
                    if (!Xl.IsBlank(v[hr, c]) && string.Equals(Xl.Str(v[r, c]).Trim(), Xl.Str(v[hr, c]).Trim(), StringComparison.OrdinalIgnoreCase)) same++;
                }
                if (same >= 2 && same >= 0.6 * filled) repeatHead++;
            }

            // most useful first: fix the data, then build on it
            if (repeatHead > 0 || totals >= 2)
                Add(recs, "report", "This looks like a printed report: " + N(repeatHead) + " repeated header rows and " + N(totals) + " total rows. Get one clean table from it.");
            if (textNum + textDate + spaces > 0)
                Add(recs, "clean", N(textNum) + " numbers stored as text, " + N(textDate) + " dates stored as text and " + N(spaces) + " cells with extra spaces can be fixed in one go.");
            if (merged) Add(recs, "unmerge", "The table has merged cells. They break sorting, filters and pivot tables.");
            if (blankRows > 0) Add(recs, "blankrows", N(blankRows) + " completely empty rows sit inside the table.");
            if (nNum > 0 && (nCat > 0 || nDate > 0))
            {
                Add(recs, "dash", N(nNum) + " number columns and " + N(nCat + nDate) + " columns to group by: a dashboard with KPI cards, charts and slicers can be built now.");
                Add(recs, "custom", "Design it yourself with a live preview: cards, charts, filters, 26 themes or your own colours, layouts and your logo. 78 ready designs to start from.");
                Add(recs, "webdash", "A web page file with KPI cards, charts and filters that opens on any PC or phone, without Excel.");
            }
            if (nNum > 0 && (nCat > 0 || nDate > 0)) Add(recs, "analyst", "A written analysis of this table in one click: who drives '" + firstNum + "', the trend by month, unusual rows and what to do next.");
            if (nCat > 0 && nNum > 0) Add(recs, "abc", "See which '" + firstCat + "' values make up 80% of '" + firstNum + "'.");
            if (nDate > 0 && nNum > 0 && nCat > 0) Add(recs, "ageing", "The table has dates, amounts and parties: age the amounts into 0-30, 31-60, 61-90 day buckets.");
            if (nameCol != null) Add(recs, "fuzzy", "'" + nameCol + "' holds many different names. Check it for look-alike duplicates such as typos or Pvt / Ltd differences.");
            if (errors > 0) Add(recs, "audit", N(errors) + " cells show an error value. Scan the sheet's formulas for risks.");
            Add(recs, "health", "A full check of every column with a score out of 100.");

            string sheet = (string)rng.Worksheet.Name;
            string addr = (string)rng.Address[false, false];
            return "Looked at '" + sheet + "'!" + addr + ":  " + N(rows - hr) + " rows, " + N(cols) + " columns" + (sampled ? "  (first 20,000 rows checked)" : "") + ".";
        }
    }

    internal sealed class AssistantForm : Form
    {
        readonly ListView recList = new ListView();
        readonly ListView toolList = new ListView();
        readonly TextBox search = new TextBox();
        ListView active;
        public string RunId;

        public AssistantForm(string about, List<string[]> recs)
        {
            Text = Design.AppName + " Assistant";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(860, 610);
            Font = new Font("Segoe UI", 9.5f);

            Label top = new Label();
            top.Text = about;
            top.SetBounds(12, 10, 836, 22);
            top.ForeColor = Color.FromArgb(71, 85, 105);
            Controls.Add(top);

            AddTitle("Suggested for this table   (double-click to run)", 12, 38);
            recList.SetBounds(12, 62, 836, 196);
            Setup(recList, "Do this", 200, "Why", 610);
            foreach (string[] r in recs)
            {
                ListViewItem it = new ListViewItem(new[] { r[1], r[2] });
                it.Tag = r[0];
                recList.Items.Add(it);
            }
            if (recs.Count == 0) recList.Items.Add(new ListViewItem(new[] { "No suggestions yet", "Click a cell inside your data, then open the Assistant again." }));

            AddTitle("All tools   (type what you want to do)", 12, 268);
            search.SetBounds(12, 292, 836, 26);
            toolList.SetBounds(12, 324, 836, 232);
            Setup(toolList, "Tool", 200, "What it does", 610);

            Button run = new Button();
            run.Text = "Run";
            run.SetBounds(668, 566, 85, 32);
            Button close = new Button();
            close.Text = "Close";
            close.SetBounds(763, 566, 85, 32);
            close.DialogResult = DialogResult.Cancel;
            Controls.Add(recList);
            Controls.Add(search);
            Controls.Add(toolList);
            Controls.Add(run);
            Controls.Add(close);
            AcceptButton = run;
            CancelButton = close;

            active = recs.Count > 0 ? recList : toolList;
            recList.Enter += delegate { active = recList; };
            toolList.Enter += delegate { active = toolList; };
            search.Enter += delegate { active = toolList; };
            search.TextChanged += delegate { FillTools(); };
            recList.DoubleClick += delegate { RunFrom(recList); };
            toolList.DoubleClick += delegate { RunFrom(toolList); };
            run.Click += delegate
            {
                if (active.SelectedItems.Count == 0 && active.Items.Count > 0) active.Items[0].Selected = true;
                RunFrom(active);
            };
            FillTools();
            if (recs.Count > 0) recList.Items[0].Selected = true;
        }

        void AddTitle(string text, int x, int y)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(x, y, 836, 20);
            l.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            Controls.Add(l);
        }

        static void Setup(ListView lv, string c1, int w1, string c2, int w2)
        {
            lv.View = View.Details;
            lv.FullRowSelect = true;
            lv.HideSelection = false;
            lv.MultiSelect = false;
            lv.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            lv.ShowItemToolTips = true;
            lv.Columns.Add(c1, w1);
            lv.Columns.Add(c2, w2);
        }

        // every word typed must appear in the tool's name, group, search words or description
        void FillTools()
        {
            string[] words = search.Text.Trim().ToLowerInvariant().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            toolList.BeginUpdate();
            toolList.Items.Clear();
            foreach (ToolDef t in Design.Tools)
            {
                if (t.Id == "assist") continue;
                string hay = (t.Caption + " " + t.Group + " " + t.Keys + " " + t.Tip).ToLowerInvariant();
                bool ok = true;
                foreach (string w in words)
                {
                    if (hay.IndexOf(w, StringComparison.Ordinal) < 0)
                    {
                        ok = false;
                        break;
                    }
                }
                if (!ok) continue;
                ListViewItem it = new ListViewItem(new[] { t.Caption + (Design.IsPro(t.Id) ? "  (Pro)" : ""), t.Tip });
                it.Tag = t.Id;
                it.ToolTipText = t.Tip;
                toolList.Items.Add(it);
            }
            toolList.EndUpdate();
        }

        void RunFrom(ListView lv)
        {
            if (lv.SelectedItems.Count == 0) return;
            string id = lv.SelectedItems[0].Tag as string;
            if (id == null) return;
            RunId = id;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
