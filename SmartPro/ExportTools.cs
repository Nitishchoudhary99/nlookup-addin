using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SmartPro
{
    // Web Dashboard: one HTML file with KPI cards, charts and filters that opens in any browser.
    // BI Export: a clean CSV plus ready measure formulas for Power BI and Looker Studio.
    internal static class ExportTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        sealed class Model
        {
            public string[] H;
            public ColInfo[] Ci;
            public int Hr, DateCol;
            public List<int> Meas = new List<int>(), Dims = new List<int>();
        }

        // for BI Export: the first date, up to 4 number columns, up to 4 columns to group by
        static Model Analyse(object[,] v)
        {
            Model m = new Model();
            m.Hr = Profiler.FindHeaderRow(v);
            m.H = Profiler.MakeHeaders(v, m.Hr);
            m.Ci = Profiler.Profile(v, m.Hr, m.H, 50000);
            int nc = v.GetLength(1);
            for (int c = 1; c <= nc; c++) if (m.Ci[c].Kind == "date" && m.DateCol == 0) m.DateCol = c;
            for (int rank = 1; rank <= 3; rank++)
                for (int c = 1; c <= nc; c++)
                    if (m.Ci[c].Kind == "num" && m.Meas.Count < 4 && Logic.MeasureRank(m.Ci[c].Hdr) == rank) m.Meas.Add(c);
            for (int c = 1; c <= nc; c++)
                if (m.Ci[c].Kind == "cat" && m.Ci[c].Distinct >= 2 && m.Dims.Count < 4) m.Dims.Add(c);
            return m;
        }

        // =================================================================
        //  WEB DASHBOARD
        //  On a dashboard sheet: that dashboard as a web page, with its cards, charts, filters and colours.
        //  On any other sheet: pick a table; the page gets the automatic design, or one made in Dashboard Studio.
        //  The page itself is WebTemplate.Html; WebDash puts the rows and the design into it.
        // =================================================================
        public static void WebDashboard()
        {
            DashInfo info = null;
            try { info = Dashboard.ReadInfo(Xl.App.ActiveSheet); } catch { }
            dynamic src;
            DataTab t;
            DashSpec s;
            string sheet, name;
            if (info != null)
            {
                src = Dashboard.SourceOf(info);
                if (src == null) return;
                t = Dashboard.Load(src);
                if (t == null) return;
                s = Dashboard.SpecOf(info, t);
                sheet = info.SourceSheet;
                name = info.Name;
            }
            else
            {
                src = Xl.PickRange("Web Dashboard");
                if (src == null) return;
                t = Dashboard.Load(src);
                if (t == null) return;
                sheet = (string)src.Worksheet.Name;
                name = "Dashboard " + sheet;
                s = DashAuto.Suggest(t, DashAuto.SavedLook());
                s.Title = DashText.AutoTitle(sheet, s);
                DialogResult how = Ui.YesNoCancel("Design the web dashboard first?\n\n" +
                                                  "Yes:  open Dashboard Studio and choose the cards, charts, filters and colours.\n" +
                                                  "No:   use the automatic design.\n\n" +
                                                  "Tip: on a dashboard sheet this button makes the web page of that dashboard.");
                if (how == DialogResult.Cancel) return;
                if (how == DialogResult.Yes)
                {
                    string where = sheet + "!" + (string)src.Address[false, false];
                    s = DashboardStudio.Open(t, s, where, "Make web page");
                    if (s == null) return;
                }
            }
            string source = sheet + "!" + (string)src.Address[false, false];
            string folder = Ui.PickFolder("Choose the folder for the web dashboard file");
            if (folder == null) return;
            string path = Path.Combine(folder, Xl.SafeFileName(name) + ".html");
            int rows, left;
            Xl.Status("writing the web dashboard...");
            try
            {
                string html = WebDash.Page(t, s, source, "built " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm", Inv), out rows, out left);
                File.WriteAllText(path, html, new UTF8Encoding(false));
            }
            finally
            {
                try { Xl.App.StatusBar = false; } catch { }
            }
            try { System.Diagnostics.Process.Start(path); } catch { }
            Ui.Info("Web dashboard saved and opened in your browser:\n" + path + "\n\n" +
                    s.Cards.Count.ToString(Inv) + " KPI cards, " + s.Charts.Count.ToString(Inv) + " charts, " + rows.ToString("#,##0", Inv) + " rows." +
                    (left > 0 ? "\nOnly the first " + rows.ToString("#,##0", Inv) + " rows are in the page; " + left.ToString("#,##0", Inv) + " were left out to keep the file small." : "") +
                    "\n\nIn the page a click on a bar, a slice or a filter button filters everything. At the top right anyone can change the colours (" +
                    DashThemes.Count.ToString(Inv) + " themes), the look of the cards and the kind of each chart, and print the page or save it as a PDF." +
                    "\n\nThe file contains the data it shows, so share it only with people who may see that data.");
        }

        // =================================================================
        //  BI EXPORT (Power BI / Looker Studio)
        // =================================================================
        static string Csv(string s)
        {
            if (s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0 || s.StartsWith(" ", StringComparison.Ordinal) || s.EndsWith(" ", StringComparison.Ordinal))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        public static void BiExport()
        {
            dynamic src = Xl.PickRange("BI Export");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 2)
            {
                Ui.Info("Select a table with a header row and data rows.");
                return;
            }
            string sheet = (string)src.Worksheet.Name;
            string folder = Ui.PickFolder("Choose the folder for the Power BI / Looker Studio files");
            if (folder == null) return;
            string info;
            Xl.Status("writing the export...");
            bool ok = WriteBi(v, Xl.SafeFileName(sheet), folder, out info);
            try { Xl.App.StatusBar = false; } catch { }
            Ui.Info(info);
            if (ok)
            {
                try { System.Diagnostics.Process.Start("explorer.exe", "\"" + folder + "\""); } catch { }
            }
        }

        static bool WriteBi(object[,] v, string name, string folder, out string info)
        {
            Model m = Analyse(v);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            List<int> cols = new List<int>();
            for (int c = 1; c <= nc; c++) if (m.Ci[c].Kind != "empty") cols.Add(c);
            if (cols.Count == 0)
            {
                info = "The selection has no data.";
                return false;
            }
            string csvPath = Path.Combine(folder, name + "_data.csv");
            string guidePath = Path.Combine(folder, name + "_PowerBI_Looker_guide.txt");
            string monName = Profiler.FreeName(m.H, "Month"), fyName = Profiler.FreeName(m.H, "Fiscal Year");
            int n = 0;
            using (StreamWriter w = new StreamWriter(csvPath, false, new UTF8Encoding(true)))
            {
                StringBuilder line = new StringBuilder();
                for (int i = 0; i < cols.Count; i++)
                {
                    if (i > 0) line.Append(',');
                    line.Append(Csv(m.H[cols[i]]));
                }
                if (m.DateCol > 0) line.Append(',').Append(Csv(monName)).Append(',').Append(Csv(fyName));
                w.WriteLine(line.ToString());
                for (int r = m.Hr + 1; r <= nr; r++)
                {
                    if (!Profiler.RowIsData(v, r, nc)) continue;
                    line.Length = 0;
                    for (int i = 0; i < cols.Count; i++)
                    {
                        int c = cols[i];
                        object x = v[r, c];
                        string kind = m.Ci[c].Kind;
                        double d;
                        DateTime dt;
                        if (i > 0) line.Append(',');
                        if (kind == "num")
                        {
                            if (Logic.TryNum(x, out d)) line.Append(d.ToString("R", Inv));
                        }
                        else if (kind == "date")
                        {
                            if (Logic.TryDate(x, out dt)) line.Append(dt.ToString("yyyy-MM-dd", Inv));
                        }
                        else if (!Xl.IsBlank(x)) line.Append(Csv(Xl.Str(x).Trim()));
                    }
                    if (m.DateCol > 0)
                    {
                        DateTime dd;
                        if (Logic.TryDate(v[r, m.DateCol], out dd)) line.Append(',').Append(dd.ToString("yyyy-MM", Inv)).Append(',').Append(Logic.FYLabel(dd, 4));
                        else line.Append(",,");
                    }
                    w.WriteLine(line.ToString());
                    n++;
                }
            }

            StringBuilder g = new StringBuilder();
            g.AppendLine("DOCBRISK - BI EXPORT");
            g.AppendLine("Data file: " + Path.GetFileName(csvPath));
            g.AppendLine("Rows: " + n.ToString("#,##0", Inv) + "    Columns: " + (cols.Count + (m.DateCol > 0 ? 2 : 0)));
            g.AppendLine("Numbers are plain numbers, dates are yyyy-mm-dd, total and blank rows are removed.");
            g.AppendLine();
            g.AppendLine("POWER BI DESKTOP");
            g.AppendLine("1. Home > Get data > Text/CSV. Choose the data file, then Load.");
            g.AppendLine("2. Rename the table to Data.");
            g.AppendLine("3. For each line below: New measure, paste the line, press Enter.");
            foreach (int c in m.Meas)
            {
                bool avg = Logic.MeasureIsAvg(m.H[c]);
                g.AppendLine("     " + (avg ? "Average " : "Total ") + m.H[c] + " = " + (avg ? "AVERAGE" : "SUM") + "('Data'[" + m.H[c].Replace("]", "]]") + "])");
            }
            g.AppendLine("     Rows = COUNTROWS('Data')");
            g.AppendLine("4. Suggested visuals:");
            g.AppendLine("     Cards: the measures above");
            if (m.DateCol > 0) g.AppendLine("     Line chart: " + (m.Meas.Count > 0 ? m.H[m.Meas[0]] : "Rows") + " by " + monName);
            foreach (int c in m.Dims) g.AppendLine("     Bar chart: " + (m.Meas.Count > 0 ? m.H[m.Meas[0]] : "Rows") + " by " + m.H[c] + "   (also a good slicer)");
            g.AppendLine();
            g.AppendLine("LOOKER STUDIO");
            g.AppendLine("1. Create > Data source > File Upload, and upload the data file.");
            g.AppendLine("   (Or import the file into Google Sheets and connect that sheet.)");
            g.AppendLine("2. Check the field types: dates as Date, the number columns as Number.");
            g.AppendLine("3. Add scorecards for the numbers, a time series by date, and bar charts by the columns listed above.");
            g.AppendLine();
            g.AppendLine("WHAT THIS EXPORT IS");
            g.AppendLine("A clean data file and measure formulas. It is not a finished Power BI or Looker Studio report:");
            g.AppendLine("the report itself is built inside those products, using the steps above.");
            File.WriteAllText(guidePath, g.ToString(), new UTF8Encoding(true));

            info = "Two files saved in:\n" + folder + "\n\n" + Path.GetFileName(csvPath) + "  (" + n.ToString("#,##0", Inv) + " clean rows)\n" + Path.GetFileName(guidePath) +
                   "  (steps and measure formulas)\n\nThis is a data export with formulas, not a finished Power BI or Looker Studio report.";
            return true;
        }
    }
}
