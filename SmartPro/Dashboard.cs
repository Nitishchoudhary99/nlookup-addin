using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace SmartPro
{
    // what is known about a dashboard sheet that is already in the workbook
    internal sealed class DashInfo
    {
        public dynamic Sheet;
        public string Name = "", SourceSheet = "", SourceAddress = "", DataSheet = "", PivotSheet = "";
        public DashSpec Spec;                         // null when the design could not be read, and for an older dashboard
        public bool Older;                            // made by a version before 3.0
        public List<int> OldMeasures = new List<int>(), OldDims = new List<int>();
        public string OldLook = "";
    }

    // a rectangle on a sheet, in points
    internal sealed class Pts
    {
        public double L, T, W, H;

        public Pts(double l, double t, double w, double h)
        {
            L = l;
            T = t;
            W = w;
            H = h;
        }
    }

    // =====================================================================
    //  THE DASHBOARD BUILDER
    //  One-Click Dashboard, Dashboard Studio, Edit and Refresh all end here. A dashboard is drawn from a design
    //  (DashSpec): DashPlan works out the helper table, the pivot tables and the formulas, DashLayout says where
    //  every card, chart and filter goes, and this file carries it out in Excel.
    //
    //  Three sheets are made: the dashboard, and two hidden helper sheets (a clean copy of the columns that are
    //  used, and the pivot tables). Every number on the dashboard comes from a pivot table, so every card and
    //  chart follows the filter buttons. The design is kept in a hidden column of the dashboard sheet.
    // =====================================================================
    internal static class Dashboard
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        public const int MetaCol = 190;               // column of the dashboard sheet that keeps the design
        const string OldTag = "DOCBRISK-DASH";        // what dashboards made before version 3.0 have in cell BA1
        const int RowsCellRow = 3, CardRow0 = 4;      // helper cells on the pivot sheet: the row count, then one row for each card
        const int PivotRow = 6, PivotCol = 13;        // where the first pivot table starts on the pivot sheet
        const double Radius = 7.0;                    // corner radius of cards and panels, in points

        // everything one build needs
        sealed class Job
        {
            public dynamic Wb, Ws, WsP;
            public DashSpec S;
            public DataTab T;
            public DashPlan Plan;
            public DashGrid Grid;
            public bool Shapes;                       // cards and panels are drawn as shapes (rounded corners, shadows)
            public string PRef = "";                  // 'pivot sheet'! in front of a cell address
            public int Cards, Charts, Filters;
            public List<string> Problems = new List<string>();
            public StringBuilder Log = new StringBuilder();
        }

        // =================================================================
        //  ONE-CLICK DASHBOARD
        // =================================================================
        public static void Build()
        {
            dynamic src = Xl.PickRange("One-Click Dashboard");
            if (src == null) return;
            DataTab t = Load(src);
            if (t == null) return;
            DashSpec s = DashAuto.Suggest(t, DashAuto.SavedLook());
            s.Title = DashText.AutoTitle((string)src.Worksheet.Name, s);
            string msg;
            using (new Speed())
            {
                msg = Create(src, t, s, null);
            }
            Ui.Info(msg + "\n\nTo change cards, charts, colours or layout: Edit Dashboard.");
        }

        // reads the table; says why when it cannot be used and returns null
        internal static DataTab Load(dynamic src)
        {
            int rows = (int)src.Rows.Count, cols = (int)src.Columns.Count;
            if (rows < 3 || cols < 2)
            {
                Ui.Info("Select a table with a header row, at least 2 columns and 2 data rows.");
                return null;
            }
            Xl.Status("reading the data...");
            object[,] v = Xl.ReadTyped(src);
            DataTab t = DataTab.Load(v, 0);
            try { Xl.App.StatusBar = false; } catch { }
            if (t.Rows < 2)
            {
                Ui.Info("There are not enough data rows under the header.");
                return null;
            }
            if (!DashAuto.CanBuild(t))
            {
                Ui.Info("No category or date column was found to group the data by.\nRun Data Health Scan to see how each column was read.");
                return null;
            }
            return t;
        }

        // =================================================================
        //  REFRESH, EDIT, STYLE: they work on the dashboard sheet that is open
        // =================================================================
        public static void Refresh()
        {
            DashInfo info = ActiveDashboard("Refresh Dashboard");
            if (info == null) return;
            dynamic src = SourceOf(info);
            if (src == null) return;
            DataTab t = Load(src);
            if (t == null) return;
            DashSpec s = SpecOf(info, t);
            string addr = (string)src.Address[false, false];
            if (!Ui.Ok("Build this dashboard again from '" + info.SourceSheet + "'!" + addr + " as it is now?\n\n" +
                       "Cards, charts, colours and layout stay the same. The filters you have set are kept where possible.")) return;
            string msg;
            using (new Speed())
            {
                msg = Create(src, t, s, info);
            }
            Ui.Info("Dashboard refreshed.\n\n" + msg);
        }

        public static void Edit()
        {
            DashInfo info = ActiveDashboard("Edit Dashboard");
            if (info == null) return;
            dynamic src = SourceOf(info);
            if (src == null) return;
            DataTab t = Load(src);
            if (t == null) return;
            DashSpec s = SpecOf(info, t);
            string where = info.SourceSheet + "!" + (string)src.Address[false, false];
            DashSpec done = DashboardStudio.Open(t, s, where, true);
            if (done == null) return;
            string msg;
            using (new Speed())
            {
                msg = Create(src, t, done, info);
            }
            Ui.Info("Dashboard updated.\n\n" + msg);
        }

        // the colours of the open dashboard, or (on any other sheet) of the next dashboards
        public static void ChooseStyle()
        {
            DashInfo info = null;
            try { info = ReadInfo(Xl.App.ActiveSheet); } catch { }
            DashSpec look = info != null && info.Spec != null ? info.Spec : DashAuto.SavedLook();
            string[] names = DashThemes.Names();
            string hint = (info != null ? "Pick the colours for this dashboard." : "Pick the colours for the next dashboards.") + "  Now: " + look.Theme.Name;
            int pick = Ui.Pick("Dashboard Style", hint, names, null);
            if (pick < 0) return;
            string font = look.Theme.Font;
            if (info == null)
            {
                look.Theme = DashThemes.Get(pick);
                look.Theme.Font = font;
                DashAuto.SaveLook(look);
                Ui.Info("Style set to " + DashThemes.NameOf(pick) + ".\nOne-Click Dashboard and Web Dashboard use it from now on.\n\n" +
                        "Open a dashboard sheet first to change the colours of that dashboard.");
                return;
            }
            dynamic src = SourceOf(info);
            if (src == null) return;
            DataTab t = Load(src);
            if (t == null) return;
            DashSpec s = SpecOf(info, t);
            s.Theme = DashThemes.Get(pick);
            s.Theme.Font = font;
            string msg;
            using (new Speed())
            {
                msg = Create(src, t, s, info);
            }
            Ui.Info("Style changed to " + DashThemes.NameOf(pick) + ".\n\n" + msg);
        }

        static DashInfo ActiveDashboard(string tool)
        {
            DashInfo info = null;
            try { info = ReadInfo(Xl.App.ActiveSheet); } catch { }
            if (info == null) Ui.Info("Open the dashboard sheet first, then click " + tool + ".\n\nThis works on dashboards made with One-Click Dashboard or Dashboard Studio.");
            return info;
        }

        // the table a dashboard was made from, as it is now: rows added under it are included
        internal static dynamic SourceOf(DashInfo info)
        {
            dynamic wb = info.Sheet.Parent;
            dynamic src;
            try { src = wb.Worksheets[info.SourceSheet].Range[info.SourceAddress]; }
            catch
            {
                Ui.Info("The source sheet '" + info.SourceSheet + "' was not found in this workbook, so the dashboard cannot be built again.");
                return null;
            }
            try
            {
                dynamic reg = src.CurrentRegion;
                if ((int)reg.Row == (int)src.Row && (int)reg.Column == (int)src.Column && (int)reg.Columns.Count == (int)src.Columns.Count &&
                    (int)reg.Rows.Count >= (int)src.Rows.Count) src = reg;
            }
            catch { }
            return src;
        }

        // the design of a dashboard, repaired for the table as it is now
        internal static DashSpec SpecOf(DashInfo info, DataTab t)
        {
            DashSpec s;
            if (info.Spec != null) s = info.Spec.Clone();
            else if (info.Older)
            {
                // "style,layout,cards" of the older settings
                DashSpec look = new DashSpec();
                string[] lp = info.OldLook.Split(',');
                if (lp.Length == 3)
                {
                    int style, layout, cards;
                    int.TryParse(lp[0].Trim(), NumberStyles.Integer, Inv, out style);
                    int.TryParse(lp[1].Trim(), NumberStyles.Integer, Inv, out layout);
                    int.TryParse(lp[2].Trim(), NumberStyles.Integer, Inv, out cards);
                    look.Theme = DashThemes.Get(style < 0 || style >= 12 ? 0 : style);
                    int[] map = { 0, 1, 2, 6, 4, 3 };
                    look.Arrange = layout >= 0 && layout < map.Length ? map[layout] : 0;
                    look.CardStyle = cards == 1 ? 1 : 0;
                }
                s = DashAuto.FromPicks(t, look, info.OldMeasures, info.OldDims);
                s.Title = DashText.AutoTitle(info.SourceSheet, s);
            }
            else
            {
                s = DashAuto.Suggest(t, DashAuto.SavedLook());
                s.Title = DashText.AutoTitle(info.SourceSheet, s);
            }
            DashAuto.Fix(s, t);
            return s;
        }

        static List<int> ParseInts(string s)
        {
            List<int> list = new List<int>();
            foreach (string part in (s ?? "").Split(','))
            {
                int n;
                if (int.TryParse(part.Trim(), NumberStyles.Integer, Inv, out n) && n >= 1) list.Add(n);
            }
            return list;
        }

        // null when the sheet is not a dashboard
        internal static DashInfo ReadInfo(dynamic ws)
        {
            if (ws == null) return null;
            try
            {
                string tag = Xl.Str((object)ws.Cells[1, MetaCol].Value2);
                if (tag == DashSpec.Magic)
                {
                    DashInfo info = new DashInfo();
                    info.Sheet = ws;
                    info.Name = (string)ws.Name;
                    info.SourceSheet = Xl.Str((object)ws.Cells[2, MetaCol].Value2);
                    info.SourceAddress = Xl.Str((object)ws.Cells[3, MetaCol].Value2);
                    info.DataSheet = Xl.Str((object)ws.Cells[4, MetaCol].Value2);
                    info.PivotSheet = Xl.Str((object)ws.Cells[5, MetaCol].Value2);
                    int n;
                    int.TryParse(Xl.Str((object)ws.Cells[6, MetaCol].Value2), NumberStyles.Integer, Inv, out n);
                    List<string> pieces = new List<string>();
                    for (int i = 0; i < n && i < 4000; i++) pieces.Add(Xl.Str((object)ws.Cells[7 + i, MetaCol].Value2));
                    string text = DashMeta.Unpack(pieces);
                    info.Spec = text == null ? null : DashSpec.FromText(text);
                    // a title that was typed over on the sheet is kept
                    if (info.Spec != null)
                    {
                        string shown = Xl.Str((object)ws.Cells[8, MetaCol + 1].Value2);
                        if (shown.Length > 0)
                        {
                            string now = Xl.Str((object)ws.Range[shown].Cells[1, 1].Value2).Trim();
                            if (now.Length > 0) info.Spec.Title = now;
                        }
                    }
                    return info;
                }
                string old = Xl.Str((object)ws.Range["BA1"].Value2);
                if (old == OldTag)
                {
                    DashInfo info = new DashInfo();
                    info.Sheet = ws;
                    info.Name = (string)ws.Name;
                    info.Older = true;
                    info.SourceSheet = Xl.Str((object)ws.Range["BA2"].Value2);
                    info.SourceAddress = Xl.Str((object)ws.Range["BA3"].Value2);
                    info.OldMeasures = ParseInts(Xl.Str((object)ws.Range["BA4"].Value2));
                    info.OldDims = ParseInts(Xl.Str((object)ws.Range["BA5"].Value2));
                    string[] hp = Xl.Str((object)ws.Range["BA6"].Value2).Split('|');
                    if (hp.Length == 2)
                    {
                        info.DataSheet = hp[0];
                        info.PivotSheet = hp[1];
                    }
                    info.OldLook = Xl.Str((object)ws.Range["BA7"].Value2);
                    return info;
                }
            }
            catch { }
            return null;
        }

        static void WriteMeta(Job j, string srcSheet, string srcAddr, string dataSheet, string pivotSheet, string titleAddress)
        {
            dynamic ws = j.Ws;
            List<string> pieces = DashMeta.Pack(j.S.ToText());
            ws.Range[ws.Cells[1, MetaCol], ws.Cells[7 + pieces.Count, MetaCol + 1]].NumberFormat = "@";
            ws.Cells[1, MetaCol].Value2 = DashSpec.Magic;
            ws.Cells[2, MetaCol].Value2 = srcSheet;
            ws.Cells[3, MetaCol].Value2 = srcAddr;
            ws.Cells[4, MetaCol].Value2 = dataSheet;
            ws.Cells[5, MetaCol].Value2 = pivotSheet;
            ws.Cells[6, MetaCol].Value2 = pieces.Count.ToString(Inv);
            for (int i = 0; i < pieces.Count; i++) ws.Cells[7 + i, MetaCol].Value2 = pieces[i];
            ws.Cells[1, MetaCol + 1].Value2 = "DocBrisk keeps the design of this dashboard in these two columns. Do not change them.";
            ws.Cells[8, MetaCol + 1].Value2 = titleAddress;
            Xl.Try(delegate { ws.Range[ws.Columns[MetaCol - 1], ws.Columns[MetaCol + 2]].Hidden = true; });
        }

        // =================================================================
        //  THE BUILD
        //  old = the dashboard this one takes the place of (Refresh, Edit, Style), or null for a new one.
        //  The new sheets are made first; the old ones are only taken away when everything worked.
        // =================================================================
        internal static string Create(dynamic src, DataTab t, DashSpec s, DashInfo old)
        {
            dynamic wb = src.Worksheet.Parent;
            string srcSheet = (string)src.Worksheet.Name;
            string srcAddr = (string)src.Address[false, false];
            List<string> notes = DashAuto.Fix(s, t);
            if (s.Title.Trim().Length == 0) s.Title = DashText.AutoTitle(srcSheet, s);

            Job j = new Job();
            j.Wb = wb;
            j.S = s;
            j.T = t;
            List<string> made = new List<string>();
            Dictionary<string, List<string>> picks = old != null ? ReadPicks(wb, old.Name) : null;
            try
            {
                Xl.Status("working out the dashboard...");
                j.Plan = DashPlan.Build(t, s);
                DashCache cache = new DashCache();
                j.Grid = cache.Grid(t, s);
                List<string> insights = cache.Insights(t, s);

                // ---- helper sheet 1: a clean copy of the columns the dashboard uses ----
                Xl.Status("writing the helper table...");
                dynamic wsD = Xl.AddSheet(wb, "SP_Data");
                made.Add((string)wsD.Name);
                int nr = j.Plan.Rows + 1, nc = j.Plan.Headers.Count;
                dynamic rngData = wsD.Range[wsD.Cells[1, 1], wsD.Cells[nr, nc]];
                wsD.Range[wsD.Cells[1, 1], wsD.Cells[1, nc]].NumberFormat = "@";
                for (int c = 0; c < nc; c++)
                {
                    if (j.Plan.IsText[c]) wsD.Range[wsD.Cells[2, c + 1], wsD.Cells[nr, c + 1]].NumberFormat = "@";
                }
                rngData.Value2 = j.Plan.Data;

                // ---- helper sheet 2: the pivot tables ----
                Xl.Status("building the pivot tables...");
                dynamic wsP = Xl.AddSheet(wb, "SP_Pivots");
                made.Add((string)wsP.Name);
                j.WsP = wsP;
                j.PRef = DashFormula.Sheet((string)wsP.Name);
                dynamic cache0 = MakeCache(wb, wsD, rngData);
                int col = PivotCol;
                foreach (PlanPivot p in j.Plan.Pivots)
                {
                    wsP.Cells[2, col].Value2 = "'" + p.Purpose;
                    MakePivot(cache0, wsP, p, col, j.Plan);
                    col += p.Columns + 1;
                }
                WriteCardCells(j);

                // ---- the dashboard sheet ----
                Xl.Status("drawing the dashboard...");
                dynamic ws = Xl.AddSheet(wb, "Dashboard");
                made.Add((string)ws.Name);
                j.Ws = ws;
                SetupSheet(j);
                j.Shapes = s.Rounded || s.Shadow;
                if (j.Shapes)
                {
                    // can shapes be drawn at all? (a protected workbook, an odd Excel build)
                    try
                    {
                        dynamic probe = ws.Shapes.AddShape(C.msoShapeRectangle, 1.0, 1.0, 4.0, 4.0);
                        probe.Delete();
                    }
                    catch (Exception ex)
                    {
                        j.Shapes = false;
                        Problem(j, "Rounded corners and shadows were left out: this workbook does not let shapes be drawn.", ex);
                    }
                }
                string titleAddress = "";
                foreach (Box b in j.Grid.Boxes)
                {
                    try
                    {
                        if (b.Kind == Box.Header) titleAddress = DrawHeader(j, b, srcSheet + "!" + srcAddr);
                        else if (b.Kind == Box.Card && b.Index < s.Cards.Count) DrawCard(j, b, b.Index);
                        else if (b.Kind == Box.Chart && b.Index < s.Charts.Count) DrawChart(j, b, b.Index);
                        else if (b.Kind == Box.Insights) DrawInsights(j, b, insights);
                        else if (b.Kind == Box.Footer) DrawFooter(j, b);
                    }
                    catch (Exception ex)
                    {
                        string what = b.Kind == Box.Card ? "KPI card " + (b.Index + 1).ToString(Inv) : b.Kind == Box.Chart ? "Chart " + (b.Index + 1).ToString(Inv) :
                                      b.Kind == Box.Header ? "The title" : b.Kind == Box.Insights ? "The key insights" : "The footer";
                        Problem(j, what + " could not be drawn.", ex);
                    }
                }
                Xl.Status("adding the filters...");
                AddSlicers(j, picks);
                WriteMeta(j, srcSheet, srcAddr, (string)wsD.Name, (string)wsP.Name, titleAddress);

                // ---- finish ----
                Xl.Try(delegate { wsP.Calculate(); });
                wsD.Visible = C.xlSheetHidden;
                wsP.Visible = C.xlSheetHidden;
                if (old != null) Replace(j, old);
                else ws.Move(wb.Sheets[1]);
                ws.Activate();
                FinishWindow(j);
                if (s.PrintSetup) PrintSetup(j);
                Xl.Try(delegate { ws.Range["A1"].Select(); });
                DashAuto.SaveLook(s);
            }
            catch
            {
                // nothing half-made is left behind, and a dashboard that was being replaced is still there
                foreach (string nm in made)
                {
                    try { wb.Worksheets[nm].Delete(); } catch { }
                }
                throw;
            }
            if (j.Log.Length > 0)
            {
                try { File.WriteAllText(Path.Combine(Xl.BackupFolder(), "last-dashboard-notes.txt"), j.Log.ToString()); } catch { }
            }

            StringBuilder msg = new StringBuilder();
            msg.Append(t.Rows.ToString("#,##0", Inv)).Append(" rows used");
            if (t.Skipped > 0) msg.Append(", ").Append(t.Skipped.ToString("#,##0", Inv)).Append(" blank or total rows skipped");
            msg.Append(".\n").Append(j.Cards.ToString(Inv)).Append(" KPI cards, ").Append(j.Charts.ToString(Inv)).Append(" charts, ")
               .Append(j.Filters.ToString(Inv)).Append(" filters.");
            if (j.Filters > 0) msg.Append("\n\nClick a filter button: every card and chart follows it.");
            if (j.Plan.Undated > 0) msg.Append("\n\n").Append(j.Plan.Undated.ToString("#,##0", Inv)).Append(" rows have no date. Charts over time leave them out; the cards count them.");
            foreach (string n in notes) msg.Append("\n\n").Append(n);
            foreach (string n in j.Problems) msg.Append("\n\n").Append(n);
            return msg.ToString();
        }

        // something that is nice to have did not work: the dashboard is still built, and the reason is written down
        static void Problem(Job j, string text, Exception ex)
        {
            if (!j.Problems.Contains(text)) j.Problems.Add(text);
            j.Log.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", Inv)).Append("  ").Append(text).Append("\r\n");
            if (ex != null) j.Log.Append(ex.ToString()).Append("\r\n\r\n");
        }

        // the new dashboard takes the place, the name and the filter settings of the old one
        static void Replace(Job j, DashInfo old)
        {
            dynamic wb = j.Wb;
            dynamic ws = j.Ws;
            string name = old.Name;
            Xl.Try(delegate { ws.Move(old.Sheet); });
            DropSlicerCaches(wb, name);
            old.Sheet.Delete();
            foreach (string h in new[] { old.DataSheet, old.PivotSheet })
            {
                if (h == null || h.Length == 0) continue;
                try { wb.Worksheets[h].Delete(); } catch { }
            }
            Xl.Try(delegate { ws.Name = name; });
        }

        // ---------------------------------------------------------------- pivot tables
        static dynamic MakeCache(dynamic wb, dynamic wsD, dynamic rngData)
        {
            dynamic pc = null;
            string srcText = DashFormula.Sheet((string)wsD.Name) + (string)rngData.Address[true, true, C.xlR1C1];
            // version 14 (Excel 2010) is asked for: the running-total share of a Pareto chart needs it
            try { pc = wb.PivotCaches().Create(C.xlDatabase, rngData, C.xlPivotTableVersion14); } catch { }
            if (pc == null) { try { pc = wb.PivotCaches().Create(C.xlDatabase, srcText, C.xlPivotTableVersion14); } catch { } }
            if (pc == null) { try { pc = wb.PivotCaches().Create(C.xlDatabase, rngData); } catch { } }
            if (pc == null) pc = wb.PivotCaches().Create(C.xlDatabase, srcText);
            return pc;
        }

        static void MakePivot(dynamic cache, dynamic wsP, PlanPivot p, int col, DashPlan plan)
        {
            dynamic dest = wsP.Cells[PivotRow, col];
            dynamic pt = null;
            try { pt = cache.CreatePivotTable(dest, p.Name, Xl.M, C.xlPivotTableVersion14); } catch { }
            if (pt == null) pt = cache.CreatePivotTable(dest, p.Name);
            p.Table = pt;
            Xl.Try(delegate { pt.ManualUpdate = true; });
            if (p.DatedOnly && plan.DatedField.Length > 0)
            {
                // only the rows that have a date
                try
                {
                    dynamic page = pt.PivotFields(plan.DatedField);
                    page.Orientation = C.xlPageField;
                    page.CurrentPage = DashPlan.Yes;
                }
                catch { }
            }
            dynamic rf = null, cf = null;
            if (p.RowField.Length > 0)
            {
                rf = pt.PivotFields(p.RowField);
                rf.Orientation = C.xlRowField;
            }
            if (p.ColField.Length > 0)
            {
                cf = pt.PivotFields(p.ColField);
                cf.Orientation = C.xlColumnField;
            }
            List<object> fields = new List<object>();
            foreach (PlanField f in p.Data)
            {
                dynamic df = pt.AddDataField(pt.PivotFields(f.Source), f.Caption, f.Function);
                fields.Add((object)df);
            }
            pt.RowGrand = false;
            pt.ColumnGrand = false;
            Xl.Try(delegate { pt.ManualUpdate = false; });

            // ---- the order of the rows and of the columns ----
            if (rf != null && p.Order != null)
            {
                p.ByHand = SetOrder(rf, p.Order);
                if (!p.ByHand && p.PlainRow.Length > 0)
                {
                    // the order could not be set by hand: the names that sort in time order take over
                    try
                    {
                        rf.Orientation = C.xlHidden;
                        rf = pt.PivotFields(p.PlainRow);
                        rf.Orientation = C.xlRowField;
                        p.RowField = p.PlainRow;
                    }
                    catch { }
                }
            }
            if (cf != null && p.ColOrder != null)
            {
                if (!SetOrder(cf, p.ColOrder) && p.PlainCol.Length > 0)
                {
                    try
                    {
                        cf.Orientation = C.xlHidden;
                        cf = pt.PivotFields(p.PlainCol);
                        cf.Orientation = C.xlColumnField;
                        p.ColField = p.PlainCol;
                    }
                    catch { }
                }
            }
            string cap = p.Data.Count > 0 ? p.Data[0].Caption : "";
            if (rf != null && !p.ByHand)
            {
                try
                {
                    if (p.Sort == 0) rf.AutoSort(C.xlAscending, p.RowField);
                    else if (p.Sort == 1) rf.AutoSort(C.xlDescending, p.RowField);
                    else if (p.Sort == 2) rf.AutoSort(C.xlDescending, cap);
                    else rf.AutoSort(C.xlAscending, cap);
                }
                catch { }
            }
            if (rf != null && p.TopN > 0)
            {
                try { rf.AutoShow(C.xlAutomatic, C.xlTop, p.TopN, cap); } catch { }
            }

            // ---- numbers shown as a share or as a running total ----
            for (int i = 0; i < p.Data.Count; i++)
            {
                PlanField f = p.Data[i];
                dynamic df = fields[i];
                if (f.Calc != 0)
                {
                    bool ok = false;
                    try
                    {
                        df.Calculation = f.Calc;
                        if (f.Calc != C.xlPercentOfTotal) df.BaseField = p.RowField;
                        ok = true;
                    }
                    catch { }
                    if (!ok && f.Calc == C.xlPercentRunningTotal)
                    {
                        // an older Excel: the plain running total is drawn instead of its share
                        try
                        {
                            df.Calculation = C.xlRunningTotal;
                            df.BaseField = p.RowField;
                            f.Calc = C.xlRunningTotal;
                            f.Format = "";
                            ok = true;
                        }
                        catch { }
                    }
                    if (!ok) f.Calc = 0;
                }
                if (f.Format.Length > 0)
                {
                    try { df.NumberFormat = f.Format; } catch { }
                }
            }

            // ---- where the numbers are, for the formulas ----
            try
            {
                dynamic body = pt.DataBodyRange;
                p.FirstRow = (int)body.Row;
                p.ValueCol = (int)body.Column;
                p.LabelCol = rf != null ? (int)pt.RowRange.Column : p.ValueCol;
            }
            catch
            {
                p.FirstRow = 0;
            }
        }

        // Puts the items of a pivot field in the given order and checks that Excel kept it.
        static bool SetOrder(dynamic pf, List<string> order)
        {
            try
            {
                int n = (int)pf.PivotItems().Count;
                Dictionary<string, object> byName = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                for (int i = 1; i <= n; i++)
                {
                    dynamic item = pf.PivotItems(i);
                    byName[Convert.ToString((object)item.Name, Inv)] = (object)item;
                }
                int pos = 0;
                foreach (string nm in order)
                {
                    object o;
                    if (!byName.TryGetValue(nm, out o)) continue;
                    pos++;
                    dynamic item = o;
                    item.Position = pos;
                }
                if (pos == 0) return false;
                pos = 0;
                foreach (string nm in order)
                {
                    object o;
                    if (!byName.TryGetValue(nm, out o)) continue;
                    pos++;
                    dynamic item = o;
                    if ((int)item.Position != pos) return false;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        static string Abs(int row, int col)
        {
            return "$" + Xl.ColLetter(col) + "$" + row.ToString(Inv);
        }

        static string AbsRange(int row1, int col1, int row2, int col2)
        {
            return Abs(row1, col1) + ":" + Abs(row2, col2);
        }

        // The helper cells of the cards, on the pivot sheet. Columns: A what, B the number, C the number as text,
        // D the note, E the note for a rise, F the note for a fall, G months, H the earlier month, I the change, J per row.
        static void WriteCardCells(Job j)
        {
            dynamic ws = j.WsP;
            DashPlan plan = j.Plan;
            DashSpec s = j.S;
            string[] heads = { "Helper cells of the dashboard", "Number", "As text", "Note", "Note (rise)", "Note (fall)", "Months", "Earlier month", "Change", "Per row" };
            for (int k = 0; k < heads.Length; k++) ws.Cells[2, k + 1].Value2 = heads[k];
            ws.Cells[1, 1].Value2 = "DocBrisk dashboard: the cards and charts read from this sheet. Do not change it.";
            PlanPivot kpi = plan.Pivots[plan.KpiPivot];
            dynamic kpiTable = kpi.Table;
            string kpiCell = (string)kpiTable.TableRange1.Cells[1, 1].Address;
            ws.Cells[RowsCellRow, 1].Value2 = "Rows";
            ws.Cells[RowsCellRow, 2].Formula = DashFormula.KpiValue(plan.RowsField.Caption, kpiCell);
            string rowsRef = Abs(RowsCellRow, 2);
            for (int i = 0; i < plan.Cards.Count; i++)
            {
                PlanCard pc = plan.Cards[i];
                CardSpec c = s.Cards[i];
                int r = CardRow0 + i;
                string v = Abs(r, 2);
                ws.Cells[r, 1].Value2 = "'" + pc.Label;
                if (pc.Field != null && pc.Field2 != null) ws.Cells[r, 2].Formula = DashFormula.KpiCalc(c.Op, pc.Field.Caption, pc.Field2.Caption, kpiCell);
                else if (pc.Field != null) ws.Cells[r, 2].Formula = DashFormula.KpiValue(pc.Field.Caption, kpiCell);
                else
                {
                    PlanPivot d = pc.DistinctPivot >= 0 ? plan.Pivots[pc.DistinctPivot] : null;
                    if (d != null && d.FirstRow > 0)
                        ws.Cells[r, 2].Formula = DashFormula.DistinctCount(AbsRange(d.FirstRow, d.LabelCol, d.FirstRow + plan.DistinctRows(pc.DistinctPivot), d.LabelCol));
                    else ws.Cells[r, 2].Value2 = DashEngine.Scalar(j.T, c.Col, c.Agg);
                }
                ws.Cells[r, 3].Formula = DashFormula.CardText(v, pc.Format);

                PlanPivot m = pc.MonthPivot >= 0 ? plan.Pivots[pc.MonthPivot] : null;
                if (c.Compare > 0 && m != null && m.FirstRow > 0)
                {
                    string labels = AbsRange(m.FirstRow, m.LabelCol, m.FirstRow + DashPlan.MonthRows - 1, m.LabelCol);
                    string values = AbsRange(m.FirstRow, m.ValueCol, m.FirstRow + DashPlan.MonthRows - 1, m.ValueCol);
                    ws.Cells[r, 7].Formula = DashFormula.MonthCount(labels);
                    ws.Cells[r, 8].Formula = DashFormula.Earlier(c.Compare, labels, values, Abs(r, 7));
                    ws.Cells[r, 9].Formula = DashFormula.Change(values, Abs(r, 7), Abs(r, 8));
                    ws.Cells[r, 4].Formula = DashFormula.ChangeNote(Abs(r, 9), c.Compare, 0);
                    ws.Cells[r, 5].Formula = DashFormula.ChangeNote(Abs(r, 9), c.Compare, 1);
                    ws.Cells[r, 6].Formula = DashFormula.ChangeNote(Abs(r, 9), c.Compare, -1);
                }
                else if (c.Compare > 0) ws.Cells[r, 4].Value2 = "' ";
                else if (c.Target > 0) ws.Cells[r, 4].Formula = DashFormula.TargetNote(v, c.Target, s.Unit);
                else if (c.Agg == AggKind.Sum && c.Col.Length > 0 && c.Op == CardOp.None)
                {
                    ws.Cells[r, 10].Formula = DashFormula.PerRow(v, rowsRef);
                    ws.Cells[r, 4].Formula = DashFormula.PerRowNote(Abs(r, 10), s.Unit == 0);
                }
                else ws.Cells[r, 4].Value2 = "'" + DashText.CardNote(c);
            }
            Xl.Try(delegate { ws.Range["A2:J2"].Font.Bold = true; });
        }

        // ---------------------------------------------------------------- the sheet
        static dynamic Rng(dynamic ws, int x, int y, int w, int h)
        {
            return ws.Range[ws.Cells[y + 1, x + 1], ws.Cells[y + h, x + w]];
        }

        // where a block of grid cells is on the sheet, in points
        static Pts At(dynamic ws, int x, int y, int w, int h)
        {
            dynamic r = ws.Range[ws.Cells[y + 1, x + 1], ws.Cells[y + h, x + w]];
            return new Pts(Convert.ToDouble((object)r.Left, Inv), Convert.ToDouble((object)r.Top, Inv),
                           Convert.ToDouble((object)r.Width, Inv), Convert.ToDouble((object)r.Height, Inv));
        }

        // The sheet becomes graph paper: small square cells, so that everything can be placed exactly.
        static void SetupSheet(Job j)
        {
            dynamic ws = j.Ws;
            ThemeSpec th = j.S.Theme;
            ws.Cells.Font.Name = th.Font;
            ws.Cells.Font.Size = 9;
            ws.Cells.Font.Color = th.Text;
            ws.Cells.Interior.Color = th.Page;
            dynamic cols = ws.Range[ws.Columns[1], ws.Columns[j.Grid.Cols + 4]];
            cols.ColumnWidth = 1.0;
            // column widths are counted in characters of the workbook's font: the width is corrected until a column is 9 points
            for (int k = 0; k < 4; k++)
            {
                double w = Convert.ToDouble((object)ws.Columns[1].Width, Inv);
                double chars = Convert.ToDouble((object)ws.Columns[1].ColumnWidth, Inv);
                if (w <= 0 || chars <= 0 || Math.Abs(w - DashLayout.CellPt) < 0.4) break;
                cols.ColumnWidth = chars * DashLayout.CellPt / w;
            }
            ws.Range[ws.Rows[1], ws.Rows[j.Grid.Rows + 4]].RowHeight = DashLayout.CellPt;
        }

        // one text or formula in a merged block of cells
        static dynamic TextCells(dynamic ws, int x, int y, int w, int h, string text, bool formula, double size, bool bold, int color, int hAlign, int vAlign)
        {
            dynamic r = ws.Range[ws.Cells[y + 1, x + 1], ws.Cells[y + h, x + w]];
            r.Merge();
            if (formula) ws.Cells[y + 1, x + 1].Formula = text;
            else ws.Cells[y + 1, x + 1].Value2 = "'" + text;
            r.Font.Size = size;
            r.Font.Bold = bold;
            r.Font.Color = color;
            r.HorizontalAlignment = hAlign;
            r.VerticalAlignment = vAlign;
            return r;
        }

        static void Edge(dynamic rng, int which, int color, int weight)
        {
            dynamic b = rng.Borders[which];
            b.LineStyle = C.xlContinuous;
            b.Weight = weight;
            b.Color = color;
        }

        // ---------------------------------------------------------------- shapes
        static void Shadow(dynamic sh, DashSpec s)
        {
            if (!s.Shadow)
            {
                Xl.Try(delegate { sh.Shadow.Visible = C.msoFalse; });
                return;
            }
            Xl.Try(delegate
            {
                dynamic sd = sh.Shadow;
                sd.Visible = C.msoTrue;
                sd.ForeColor.RGB = 0;
                sd.Transparency = s.Theme.Dark ? 0.55 : 0.84;
                sd.OffsetX = 0;
                sd.OffsetY = 1.5;
            });
            Xl.Try(delegate { sh.Shadow.Blur = 6; });
        }

        // the background of a card, a chart or the key insights
        static dynamic Panel(dynamic ws, Pts r, int fill, DashSpec s, string name)
        {
            dynamic sh = ws.Shapes.AddShape(s.Rounded ? C.msoShapeRoundedRectangle : C.msoShapeRectangle, r.L, r.T, r.W, r.H);
            sh.Fill.Solid();
            sh.Fill.ForeColor.RGB = fill;
            sh.Line.Visible = C.msoFalse;
            if (s.Rounded)
            {
                double least = Math.Min(r.W, r.H);
                Xl.Try(delegate { sh.Adjustments.Item[1] = least > 0 ? Math.Min(0.5, Radius / least) : 0.1; });
            }
            Shadow(sh, s);
            Xl.Try(delegate { sh.Name = name; });
            return sh;
        }

        // a small coloured bar: the top line or the side bar of a card
        static dynamic Strip(dynamic ws, Pts r, int fill, bool rounded)
        {
            dynamic sh = ws.Shapes.AddShape(rounded ? C.msoShapeRoundedRectangle : C.msoShapeRectangle, r.L, r.T, r.W, r.H);
            sh.Fill.Solid();
            sh.Fill.ForeColor.RGB = fill;
            sh.Line.Visible = C.msoFalse;
            if (rounded) Xl.Try(delegate { sh.Adjustments.Item[1] = 0.5; });
            Xl.Try(delegate { sh.Shadow.Visible = C.msoFalse; });
            return sh;
        }

        // A text box without fill or border. link = the cell it shows ("='Sheet'!$C$4"), or null for a fixed text.
        static dynamic TextBox(dynamic ws, Pts r, string text, string link, string font, double size, bool bold, int color, int align, int anchor)
        {
            dynamic tb = ws.Shapes.AddTextbox(C.msoTextHorizontal, r.L, r.T, r.W, r.H);
            tb.Fill.Visible = C.msoFalse;
            tb.Line.Visible = C.msoFalse;
            Xl.Try(delegate { tb.Shadow.Visible = C.msoFalse; });
            dynamic tf = tb.TextFrame2;
            tf.MarginLeft = 0;
            tf.MarginRight = 0;
            tf.MarginTop = 0;
            tf.MarginBottom = 0;
            tf.WordWrap = C.msoFalse;
            tf.VerticalAnchor = anchor;
            Xl.Try(delegate { tf.AutoSize = 0; });
            if (link != null) tb.DrawingObject.Formula = link;
            else tf.TextRange.Text = text;
            dynamic f = tf.TextRange.Font;
            f.Size = size;
            f.Bold = bold ? C.msoTrue : C.msoFalse;
            f.Fill.ForeColor.RGB = color;
            Xl.Try(delegate { f.Name = font; });
            Xl.Try(delegate { tf.TextRange.ParagraphFormat.Alignment = align; });
            return tb;
        }

        static void DeleteAll(List<object> made)
        {
            foreach (object o in made)
            {
                dynamic d = o;
                try { d.Delete(); } catch { }
            }
            made.Clear();
        }

        // ---------------------------------------------------------------- title, key insights, footer
        // returns the address of the title cell, so a title typed over on the sheet can be read back later
        static string DrawHeader(Job j, Box b, string source)
        {
            dynamic ws = j.Ws;
            DashSpec s = j.S;
            ThemeSpec th = s.Theme;
            DashGrid g = j.Grid;
            // the band runs across the whole sheet, so it looks right on a wide screen too
            ws.Range[ws.Rows[b.Y + 1], ws.Rows[b.Y + b.H]].Interior.Color = th.Band;
            if (ColorMath.Lum(th.Band) > 0.9 && ColorMath.Lum(th.Page) > 0.9)
                Edge(ws.Range[ws.Rows[b.Y + b.H], ws.Rows[b.Y + b.H]], C.xlEdgeBottom, th.Grid, C.xlThin);
            int top = b.H >= 6 ? 1 : 0;                       // one free row above the title when there is room
            int logoCells = s.LogoData.Length > 0 ? 22 : 0;
            int width = g.Width - logoCells;
            dynamic title = TextCells(ws, g.Margin, b.Y + top, width, 3, DashRender.TitleText(s), false, DashLook.TitlePt, true, th.BandText, C.xlLeft, C.xlCenter);
            string sub = s.Subtitle.Trim().Length > 0 ? s.Subtitle.Trim()
                       : DashText.AutoSubtitle(source, j.T.Rows) + "   |   built " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm", Inv);
            TextCells(ws, g.Margin, b.Y + top + 3, width, 2, sub, false, DashLook.SubtitlePt, false, ColorMath.Mix(th.Band, th.BandText, 0.72), C.xlLeft, C.xlCenter);
            if (s.LogoData.Length > 0)
            {
                try { PlaceLogo(j, b); }
                catch (Exception ex) { Problem(j, "The logo could not be placed.", ex); }
            }
            return (string)ws.Cells[b.Y + top + 1, g.Margin + 1].Address[false, false];
        }

        static void PlaceLogo(Job j, Box b)
        {
            dynamic ws = j.Ws;
            DashGrid g = j.Grid;
            string path = Path.Combine(Path.GetTempPath(), "docbrisk-logo-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(path, Convert.FromBase64String(j.S.LogoData));
            try
            {
                Pts area = At(ws, g.Margin + g.Width - 22, b.Y + 1, 22, Math.Max(1, b.H - 2));
                dynamic pic = ws.Shapes.AddPicture(path, C.msoFalse, C.msoTrue, area.L, area.T, -1, -1);
                Xl.Try(delegate { pic.LockAspectRatio = C.msoTrue; });
                double pw = Convert.ToDouble((object)pic.Width, Inv), ph = Convert.ToDouble((object)pic.Height, Inv);
                if (pw > 0 && ph > 0)
                {
                    double k = Math.Min(area.W / pw, area.H / ph);
                    pic.Height = ph * k;
                    pic.Width = pw * k;
                    pic.Left = area.L + area.W - pw * k;
                    pic.Top = area.T + (area.H - ph * k) / 2;
                }
                Xl.Try(delegate { pic.Name = "DB_Logo"; });
            }
            finally
            {
                try { File.Delete(path); } catch { }
            }
        }

        static void DrawInsights(Job j, Box b, List<string> lines)
        {
            if (lines.Count == 0) return;
            dynamic ws = j.Ws;
            DashSpec s = j.S;
            ThemeSpec th = s.Theme;
            if (j.Shapes)
            {
                List<object> made = new List<object>();
                try
                {
                    Pts r = At(ws, b.X, b.Y, b.W, b.H);
                    double u = r.H / b.H, uw = r.W / b.W;
                    made.Add((object)Panel(ws, r, th.Card, s, "DB_Insights"));
                    made.Add((object)TextBox(ws, new Pts(r.L + uw, r.T + u * 0.6, r.W - 2 * uw, 2 * u), "KEY INSIGHTS", null, th.Font, DashLook.InsightHeadPt, true, th.Muted,
                                             C.msoAlignLeft, C.msoAnchorMiddle));
                    for (int i = 0; i < lines.Count; i++)
                    {
                        made.Add((object)TextBox(ws, new Pts(r.L + uw, r.T + u * (2.8 + 2 * i), r.W - 2 * uw, 2 * u), Bullet + lines[i], null, th.Font, DashLook.InsightPt, false,
                                                 th.Text, C.msoAlignLeft, C.msoAnchorMiddle));
                    }
                    return;
                }
                catch (Exception ex)
                {
                    DeleteAll(made);
                    Problem(j, "Some rounded panels could not be drawn; plain cells were used there.", ex);
                }
            }
            Rng(ws, b.X, b.Y, b.W, b.H).Interior.Color = th.Card;
            TextCells(ws, b.X + 1, b.Y + 1, b.W - 2, 2, "KEY INSIGHTS", false, DashLook.InsightHeadPt, true, th.Muted, C.xlLeft, C.xlCenter);
            for (int i = 0; i < lines.Count; i++)
                TextCells(ws, b.X + 1, b.Y + 3 + 2 * i, b.W - 2, 2, Bullet + lines[i], false, DashLook.InsightPt, false, th.Text, C.xlLeft, C.xlCenter);
        }

        const string Bullet = "\u2022  ";

        static void DrawFooter(Job j, Box b)
        {
            TextCells(j.Ws, b.X, b.Y, b.W, b.H, DashRender.FooterText(j.S), false, DashLook.FooterPt, false, j.S.Theme.Muted, C.xlLeft, C.xlCenter);
        }

        // ---------------------------------------------------------------- KPI cards
        static void DrawCard(Job j, Box b, int i)
        {
            DashSpec s = j.S;
            CardSpec c = s.Cards[i];
            PlanCard pc = j.Plan.Cards[i];
            CardLook k = DashLook.Card(s, ColorMath.Pick(s.Theme, c.Color, i));
            PlanPivot m = pc.MonthPivot >= 0 ? j.Plan.Pivots[pc.MonthPivot] : null;
            bool compare = c.Compare > 0 && m != null && m.FirstRow > 0;
            bool spark = c.Spark && m != null && m.Table != null;
            if (j.Shapes)
            {
                List<object> made = new List<object>();
                try
                {
                    CardShapes(j, b, i, k, compare, spark, made);
                    j.Cards++;
                    return;
                }
                catch (Exception ex)
                {
                    DeleteAll(made);
                    Problem(j, "Some rounded panels could not be drawn; plain cells were used there.", ex);
                }
            }
            CardCells(j, b, i, k, compare, spark);
            j.Cards++;
        }

        // the number of a card as it reads when no filter is set: its length decides the size of the font
        static string ValueNow(Job j, int i)
        {
            CardSpec c = j.S.Cards[i];
            double v = DashEngine.Scalar(j.T, c.Col, c.Agg);
            if (c.Op > 0) v = CardOp.Apply(c.Op, v, DashEngine.Scalar(j.T, c.Col2, c.Agg2));
            return NumFmt.Apply(v, j.Plan.Cards[i].Format);
        }

        // the note under the number: 0 = none, 1 = a fixed text, 2 = a formula
        static int NoteKind(CardSpec c, bool compare, out string fixedText)
        {
            fixedText = "";
            if (compare) return 2;
            if (c.Compare > 0) return 0;
            if (c.Target > 0) return 2;
            if (c.Agg == AggKind.Sum && c.Col.Length > 0 && c.Op == CardOp.None) return 2;
            fixedText = DashText.CardNote(c);
            return fixedText.Length > 0 ? 1 : 0;
        }

        static void CardShapes(Job j, Box b, int i, CardLook k, bool compare, bool spark, List<object> made)
        {
            dynamic ws = j.Ws;
            DashSpec s = j.S;
            ThemeSpec th = s.Theme;
            CardSpec c = s.Cards[i];
            int row = CardRow0 + i;
            string num = (i + 1).ToString(Inv);
            Pts r = At(ws, b.X, b.Y, b.W, b.H);
            double u = r.H / b.H, uw = r.W / b.W;
            dynamic sh = Panel(ws, r, k.Bg, s, "DB_Card" + num);
            made.Add((object)sh);
            if (k.Gradient)
            {
                Xl.Try(delegate
                {
                    sh.Fill.TwoColorGradient(C.msoGradientVertical, 1);
                    sh.Fill.ForeColor.RGB = k.Bg;
                    sh.Fill.BackColor.RGB = k.Bg2;
                });
            }
            if (k.LineKind == 1)
            {
                Pts lr = s.Rounded ? new Pts(r.L + uw, r.T + u * 0.45, uw * 3, Math.Max(2.0, u * 0.3)) : new Pts(r.L, r.T, r.W, Math.Max(2.0, u * 0.25));
                made.Add((object)Strip(ws, lr, k.Line, s.Rounded));
            }
            else if (k.LineKind == 2)
            {
                Pts lr = s.Rounded ? new Pts(r.L + uw * 0.5, r.T + u, Math.Max(2.0, uw * 0.3), r.H - 2 * u) : new Pts(r.L, r.T, Math.Max(2.0, uw * 0.3), r.H);
                made.Add((object)Strip(ws, lr, k.Line, s.Rounded));
            }
            else if (k.LineKind == 3)
            {
                sh.Line.Visible = C.msoTrue;
                sh.Line.ForeColor.RGB = ColorMath.Mix(th.Card, k.Line, 0.6);
                sh.Line.Weight = 1.25;
            }
            double pad = uw * (k.LineKind == 2 ? 1.4 : 1.0);
            int[] rows = DashLook.CardRows(b.H);
            double inner = r.W - pad - uw;
            double sparkW = spark ? inner * 0.38 : 0;
            string link = "=" + j.PRef;
            made.Add((object)TextBox(ws, new Pts(r.L + pad, r.T + rows[0] * u, inner, 2 * u), DashText.CardLabel(c).ToUpperInvariant(), null, th.Font, DashLook.CardLabelPt, true,
                                     k.Label, C.msoAlignLeft, C.msoAnchorMiddle));
            double valuePt = DashLook.FitFont(ValueNow(j, i), inner - sparkW, b.W < 16 ? DashLook.CardValueSmallPt : DashLook.CardValuePt);
            made.Add((object)TextBox(ws, new Pts(r.L + pad, r.T + rows[1] * u, inner - sparkW, 3 * u), null, link + Abs(row, 3), th.Font, valuePt, true, k.Value,
                                     C.msoAlignLeft, C.msoAnchorMiddle));
            string fixedText;
            int note = NoteKind(c, compare, out fixedText);
            Pts nr = new Pts(r.L + pad, r.T + rows[2] * u, inner, 2 * u);
            if (note == 1) made.Add((object)TextBox(ws, nr, fixedText, null, th.Font, DashLook.CardNotePt, false, k.Note, C.msoAlignLeft, C.msoAnchorTop));
            else if (note == 2 && compare && k.Plain)
            {
                // two boxes on the same spot: one shows a rise in green, the other a fall in red; only one of them has a text
                made.Add((object)TextBox(ws, nr, null, link + Abs(row, 5), th.Font, DashLook.CardNotePt, false, DashLook.UpColor(s), C.msoAlignLeft, C.msoAnchorTop));
                made.Add((object)TextBox(ws, nr, null, link + Abs(row, 6), th.Font, DashLook.CardNotePt, false, DashLook.DownColor(s), C.msoAlignLeft, C.msoAnchorTop));
            }
            else if (note == 2) made.Add((object)TextBox(ws, nr, null, link + Abs(row, 4), th.Font, DashLook.CardNotePt, false, k.Note, C.msoAlignLeft, C.msoAnchorTop));
            if (spark)
            {
                Pts sr = new Pts(r.L + pad + inner - sparkW + uw * 0.3, r.T + rows[1] * u + u * 0.2, sparkW - uw * 0.3, 3 * u - u * 0.4);
                try { made.Add((object)MiniChart(j, sr, j.Plan.Cards[i].MonthPivot, k.Plain ? k.Line : k.Value, "DB_Trend" + num)); }
                catch (Exception ex) { Problem(j, "The small trend line of a card could not be drawn.", ex); }
            }
        }

        static void CardCells(Job j, Box b, int i, CardLook k, bool compare, bool spark)
        {
            dynamic ws = j.Ws;
            DashSpec s = j.S;
            ThemeSpec th = s.Theme;
            CardSpec c = s.Cards[i];
            int row = CardRow0 + i;
            dynamic all = Rng(ws, b.X, b.Y, b.W, b.H);
            all.Interior.Color = k.Bg;
            if (k.LineKind == 1) Edge(all, C.xlEdgeTop, k.Line, C.xlThick);
            else if (k.LineKind == 2) Edge(all, C.xlEdgeLeft, k.Line, C.xlThick);
            else if (k.LineKind == 3)
            {
                int line = ColorMath.Mix(th.Card, k.Line, 0.6);
                foreach (int which in new[] { C.xlEdgeLeft, C.xlEdgeTop, C.xlEdgeBottom, C.xlEdgeRight }) Edge(all, which, line, C.xlMedium);
            }
            int[] rows = DashLook.CardRows(b.H);
            int inner = b.W - 2;
            int sparkCells = spark ? (int)Math.Round(inner * 0.38) : 0;
            string link = "=" + j.PRef;
            TextCells(ws, b.X + 1, b.Y + rows[0], inner, 2, DashText.CardLabel(c).ToUpperInvariant(), false, DashLook.CardLabelPt, true, k.Label, C.xlLeft, C.xlCenter);
            double valuePt = DashLook.FitFont(ValueNow(j, i), (inner - sparkCells) * DashLayout.CellPt, b.W < 16 ? DashLook.CardValueSmallPt : DashLook.CardValuePt);
            dynamic val = TextCells(ws, b.X + 1, b.Y + rows[1], inner - sparkCells, 3, link + Abs(row, 3), true, valuePt, true, k.Value, C.xlLeft, C.xlCenter);
            Xl.Try(delegate { val.ShrinkToFit = true; });
            string fixedText;
            int note = NoteKind(c, compare, out fixedText);
            if (note == 1) TextCells(ws, b.X + 1, b.Y + rows[2], inner, 2, fixedText, false, DashLook.CardNotePt, false, k.Note, C.xlLeft, C.xlAlignTop);
            else if (note == 2 && compare)
            {
                // the cell holds the change as a number; its number format writes the arrow and, on a light card, the colour
                dynamic nc = TextCells(ws, b.X + 1, b.Y + rows[2], inner, 2, link + Abs(row, 9), true, DashLook.CardNotePt, false, k.Note, C.xlLeft, C.xlAlignTop);
                bool ok = false;
                try
                {
                    nc.NumberFormat = DashFormula.ChangeFormat(c.Compare, k.Plain, th.Dark);
                    ok = true;
                }
                catch { }
                if (!ok) ws.Cells[b.Y + rows[2] + 1, b.X + 2].Formula = link + Abs(row, 4);
            }
            else if (note == 2) TextCells(ws, b.X + 1, b.Y + rows[2], inner, 2, link + Abs(row, 4), true, DashLook.CardNotePt, false, k.Note, C.xlLeft, C.xlAlignTop);
            if (spark && sparkCells > 0)
            {
                try { MiniChart(j, At(ws, b.X + 1 + inner - sparkCells, b.Y + rows[1], sparkCells, 3), j.Plan.Cards[i].MonthPivot, k.Plain ? k.Line : k.Value, "DB_Trend" + (i + 1).ToString(Inv)); }
                catch (Exception ex) { Problem(j, "The small trend line of a card could not be drawn.", ex); }
            }
        }

        // the small trend line of a card: a line chart of the months without axes, title or background
        static dynamic MiniChart(Job j, Pts r, int pivot, int color, string name)
        {
            dynamic pt = j.Plan.Pivots[pivot].Table;
            dynamic co = j.Ws.ChartObjects().Add(r.L, r.T, r.W, r.H);
            dynamic ch = co.Chart;
            ch.SetSourceData(pt.TableRange1);
            ch.ChartType = C.xlLine;
            Xl.Try(delegate { ch.ShowAllFieldButtons = false; });
            Xl.Try(delegate { ch.HasTitle = false; });
            Xl.Try(delegate { ch.HasLegend = false; });
            Xl.Try(delegate { ch.Axes(C.xlValue).HasMajorGridlines = false; });
            Xl.Try(delegate { ch.Axes(C.xlValue).Delete(); });
            Xl.Try(delegate { ch.Axes(C.xlCategory).Delete(); });
            Xl.Try(delegate
            {
                ch.ChartArea.Format.Fill.Visible = C.msoFalse;
                ch.ChartArea.Format.Line.Visible = C.msoFalse;
            });
            Xl.Try(delegate { ch.PlotArea.Format.Fill.Visible = C.msoFalse; });
            Xl.Try(delegate
            {
                dynamic ser = ch.SeriesCollection(1);
                ser.Format.Line.ForeColor.RGB = color;
                ser.Format.Line.Weight = 1.75;
                ser.MarkerStyle = C.xlNone;
            });
            Xl.Try(delegate { co.Name = name; });
            return co;
        }

        // ---------------------------------------------------------------- charts
        static int ExcelChartType(int kind, int points)
        {
            switch (kind)
            {
                case ChartKind.Bar: return C.xlBarClustered;
                case ChartKind.Line: return points > 30 ? C.xlLine : C.xlLineMarkers;
                case ChartKind.Area: return C.xlArea;
                case ChartKind.Pie: return C.xlPie;
                case ChartKind.Doughnut: return C.xlDoughnut;
                case ChartKind.StackedColumn: return C.xlColumnStacked;
                case ChartKind.StackedBar: return C.xlBarStacked;
                case ChartKind.Stacked100: return C.xlColumnStacked100;
                default: return C.xlColumnClustered;
            }
        }

        static void DrawChart(Job j, Box b, int i)
        {
            DashSpec s = j.S;
            ChartSpec c = s.Charts[i];
            PlanChart pc = j.Plan.Charts[i];
            PlanPivot pv = j.Plan.Pivots[pc.Pivot];
            if (pc.Result.Cats.Count == 0 || pc.Result.S.Count == 0 || pv.Table == null) return;
            if (c.Kind == ChartKind.Table)
            {
                TableCells(j, b, i);
                j.Charts++;
                return;
            }
            dynamic ws = j.Ws;
            ThemeSpec th = s.Theme;
            Pts r = At(ws, b.X, b.Y, b.W, b.H);
            bool panel = false;
            if (j.Shapes)
            {
                try
                {
                    Panel(ws, r, th.Card, s, "DB_Panel" + (i + 1).ToString(Inv));
                    panel = true;
                }
                catch (Exception ex) { Problem(j, "Some rounded panels could not be drawn; plain cells were used there.", ex); }
            }
            if (!panel) Rng(ws, b.X, b.Y, b.W, b.H).Interior.Color = th.Card;

            double inset = panel ? 3.0 : 0.0;
            dynamic co = ws.ChartObjects().Add(r.L + inset, r.T + inset, r.W - 2 * inset, r.H - 2 * inset);
            dynamic ch = co.Chart;
            dynamic pt = pv.Table;
            ch.SetSourceData(pt.TableRange1);
            ch.ChartType = ExcelChartType(c.Kind, pc.Result.Cats.Count);
            j.Charts++;
            Xl.Try(delegate { co.Name = "DB_Chart" + (i + 1).ToString(Inv); });
            Xl.Try(delegate { ch.ShowAllFieldButtons = false; });

            // ---- background and text ----
            Xl.Try(delegate
            {
                if (panel) ch.ChartArea.Format.Fill.Visible = C.msoFalse;
                else
                {
                    ch.ChartArea.Format.Fill.Visible = C.msoTrue;
                    ch.ChartArea.Format.Fill.ForeColor.RGB = th.Card;
                }
                ch.ChartArea.Format.Line.Visible = C.msoFalse;
            });
            Xl.Try(delegate { ch.PlotArea.Format.Fill.Visible = C.msoFalse; });
            if (!panel && s.Rounded) Xl.Try(delegate { co.RoundedCorners = true; });
            Xl.Try(delegate
            {
                ch.ChartArea.Font.Name = th.Font;
                ch.ChartArea.Font.Color = th.Text;
            });
            Xl.Try(delegate
            {
                ch.HasTitle = true;
                dynamic title = ch.ChartTitle;
                title.Text = pc.Title;
                title.Font.Size = DashLook.ChartTitlePt;
                title.Font.Bold = true;
                title.Font.Color = th.Text;
            });
            Xl.Try(delegate
            {
                ch.ChartTitle.Left = 6;
                ch.ChartTitle.Top = 4;
            });

            bool round = ChartKind.IsRound(c.Kind);
            int legend = DashLook.LegendPos(c, round ? pc.Result.Cats.Count : pc.Result.S.Count);
            if (legend == 1) Xl.Try(delegate { ch.HasLegend = false; });
            else
            {
                Xl.Try(delegate
                {
                    ch.HasLegend = true;
                    ch.Legend.Position = legend == 2 ? C.xlLegendRight : (legend == 3 ? C.xlLegendBottom : C.xlLegendTop);
                    ch.Legend.Font.Size = DashLook.LegendPt;
                    ch.Legend.Font.Color = th.Muted;
                });
            }
            if (round) StyleRound(j, ch, c, pc, pv);
            else StyleAxes(j, ch, c, pc, i, r.W);
        }

        // pie and doughnut: one colour for each slice, Others in grey, shares on the slices
        static void StyleRound(Job j, dynamic ch, ChartSpec c, PlanChart pc, PlanPivot pv)
        {
            DashSpec s = j.S;
            ThemeSpec th = s.Theme;
            int mode = DashLook.LabelMode(c);
            if (c.Kind == ChartKind.Doughnut) Xl.Try(delegate { ch.ChartGroups(1).DoughnutHoleSize = 60; });
            Xl.Try(delegate
            {
                dynamic ser = ch.SeriesCollection(1);
                List<string> names = PointNames(ser);
                int n = (int)ser.Points().Count;
                int[] colors = new int[n];
                for (int k = 0; k < n; k++)
                {
                    bool others = pc.OthersName.Length > 0 && (k < names.Count ? names[k] == pc.OthersName : k == n - 1);
                    colors[k] = others ? DashLook.OthersColor(s) : DashLook.PointColor(s, k);
                    dynamic p = ser.Points(k + 1);
                    p.Format.Fill.ForeColor.RGB = colors[k];
                    p.Format.Line.ForeColor.RGB = th.Card;
                    p.Format.Line.Weight = 1.5;
                }
                if (mode == 1) return;
                ser.HasDataLabels = true;
                dynamic dl = ser.DataLabels();
                dl.ShowCategoryName = false;
                dl.ShowSeriesName = false;
                dl.ShowPercentage = mode == 3;
                dl.ShowValue = mode != 3;
                if (mode != 3) dl.NumberFormat = pc.LabelFormat;
                dl.Font.Size = 8;
                dl.Font.Bold = true;
                for (int k = 0; k < n; k++) ser.Points(k + 1).DataLabel.Font.Color = ColorMath.Ink(colors[k]);
            });
            if (c.Kind == ChartKind.Pie && mode != 1) Xl.Try(delegate { ch.SeriesCollection(1).DataLabels().Position = C.xlLabelCenter; });
        }

        // the names of the points of a series, in the order Excel has them
        static List<string> PointNames(dynamic ser)
        {
            List<string> names = new List<string>();
            try
            {
                object xv = ser.XValues;
                Array arr = xv as Array;
                if (arr != null)
                {
                    foreach (object o in arr) names.Add(Xl.Str(o));
                }
            }
            catch { }
            return names;
        }

        // every chart with axes: column, bar, line, area, stacked, combo, Pareto
        static void StyleAxes(Job j, dynamic ch, ChartSpec c, PlanChart pc, int index, double width)
        {
            DashSpec s = j.S;
            ThemeSpec th = s.Theme;
            AggResult res = pc.Result;
            bool bar = c.Kind == ChartKind.Bar || c.Kind == ChartKind.StackedBar;
            bool line = ChartKind.IsLine(c.Kind);
            bool combo = c.Kind == ChartKind.Combo || c.Kind == ChartKind.Pareto;
            bool stacked = ChartKind.IsStacked(c.Kind);
            int mode = DashLook.LabelMode(c);
            int single = DashLook.ChartColor(s, c, index);

            // ---- axes ----
            Xl.Try(delegate
            {
                dynamic ax = ch.Axes(C.xlValue);
                ax.TickLabels.NumberFormat = pc.AxisFormat;
                ax.TickLabels.Font.Size = DashLook.AxisPt;
                ax.TickLabels.Font.Color = th.Muted;
            });
            Xl.Try(delegate { ch.Axes(C.xlValue).MajorGridlines.Format.Line.ForeColor.RGB = th.Grid; });
            Xl.Try(delegate { ch.Axes(C.xlValue).Format.Line.Visible = C.msoFalse; });
            if (pc.DisplayUnit != 0)
            {
                Xl.Try(delegate
                {
                    dynamic ax = ch.Axes(C.xlValue);
                    ax.DisplayUnit = pc.DisplayUnit;
                    ax.HasDisplayUnitLabel = false;
                });
            }
            Xl.Try(delegate
            {
                dynamic cx = ch.Axes(C.xlCategory);
                cx.TickLabels.Font.Size = DashLook.AxisPt;
                cx.TickLabels.Font.Color = th.Muted;
                cx.Format.Line.ForeColor.RGB = th.Grid;
            });
            if (!bar) Xl.Try(delegate { ch.Axes(C.xlCategory).TickLabels.Orientation = C.xlHorizontal; });
            if (!line) Xl.Try(delegate { ch.ChartGroups(1).GapWidth = 60; });

            // ---- colours ----
            int count = 0;
            Xl.Try(delegate { count = (int)ch.SeriesCollection().Count; });
            if (stacked)
            {
                for (int k = 1; k <= count; k++)
                {
                    int at = k;
                    Xl.Try(delegate
                    {
                        dynamic ser = ch.SeriesCollection(at);
                        string name = Xl.Str((object)ser.Name);
                        int color = pc.SplitOthersName.Length > 0 && name == pc.SplitOthersName ? DashLook.OthersColor(s) : DashLook.SeriesColor(s, c, index, at - 1);
                        ser.Format.Fill.ForeColor.RGB = color;
                    });
                }
                return;
            }
            if (count < 1) return;
            Xl.Try(delegate
            {
                dynamic ser = ch.SeriesCollection(1);
                if (line)
                {
                    ser.Format.Line.ForeColor.RGB = single;
                    ser.Format.Line.Weight = 2.25;
                    if (c.Kind == ChartKind.Area)
                    {
                        ser.Format.Fill.ForeColor.RGB = single;
                        ser.Format.Fill.Transparency = 0.6;
                    }
                }
                else ser.Format.Fill.ForeColor.RGB = single;
            });
            if (c.Kind == ChartKind.Line)
            {
                Xl.Try(delegate
                {
                    dynamic ser = ch.SeriesCollection(1);
                    if (res.Cats.Count <= 30)
                    {
                        ser.MarkerStyle = C.xlMarkerCircle;
                        ser.MarkerSize = 5;
                        ser.MarkerBackgroundColor = single;
                        ser.MarkerForegroundColor = single;
                    }
                    if (c.Smooth) ser.Smooth = true;
                });
            }
            if (line && c.Trend > 0 && res.Cats.Count > 2)
            {
                Xl.Try(delegate
                {
                    dynamic ser = ch.SeriesCollection(1);
                    dynamic tl = c.Trend == 2 ? ser.Trendlines().Add(C.xlMovingAvg, Xl.M, 3) : ser.Trendlines().Add(C.xlLinear);
                    tl.Format.Line.ForeColor.RGB = th.Muted;
                    tl.Format.Line.Weight = 1.25;
                    tl.Format.Line.DashStyle = C.msoLineDash;
                });
            }
            if (mode == 2)
            {
                Xl.Try(delegate
                {
                    dynamic ser = ch.SeriesCollection(1);
                    ser.HasDataLabels = true;
                    dynamic dl = ser.DataLabels();
                    dl.NumberFormat = pc.LabelFormat;
                    dl.Font.Size = 8;
                    dl.Font.Color = th.Text;
                });
                Xl.Try(delegate { ch.SeriesCollection(1).DataLabels().Position = line ? C.xlLabelAbove : C.xlLabelOutsideEnd; });
                // with many points only every k-th one keeps its label, so that the labels do not run into each other
                int points = res.Cats.Count;
                int every = bar ? 1 : DashLook.LabelEvery(points, width * 0.82, NumFmt.Compact(res.MaxAbs, s.Unit, res.S[0].Percent).Length + 1);
                if (every > 1)
                {
                    Xl.Try(delegate
                    {
                        dynamic ser = ch.SeriesCollection(1);
                        for (int k = 1; k <= points; k++)
                        {
                            if ((k - 1) % every != 0) ser.Points(k).HasDataLabel = false;
                        }
                    });
                }
            }

            // ---- the second number of a combo or Pareto chart: a line on its own axis ----
            if (combo && count >= 2)
            {
                int second = DashLook.SeriesColor(s, c, index, 1);
                bool share = c.Kind == ChartKind.Pareto && j.Plan.Pivots[pc.Pivot].Data.Count > 1 && j.Plan.Pivots[pc.Pivot].Data[1].Calc == C.xlPercentRunningTotal;
                Xl.Try(delegate
                {
                    dynamic ser = ch.SeriesCollection(2);
                    ser.ChartType = C.xlLineMarkers;
                    ser.AxisGroup = C.xlSecondary;
                    ser.Format.Line.ForeColor.RGB = second;
                    ser.Format.Line.Weight = 2.25;
                    ser.MarkerStyle = C.xlMarkerCircle;
                    ser.MarkerSize = 5;
                    ser.MarkerBackgroundColor = second;
                    ser.MarkerForegroundColor = second;
                });
                Xl.Try(delegate
                {
                    dynamic ax = ch.Axes(C.xlValue, C.xlSecondary);
                    ax.TickLabels.Font.Size = DashLook.AxisPt;
                    ax.TickLabels.Font.Color = th.Muted;
                    ax.Format.Line.Visible = C.msoFalse;
                    if (c.Kind == ChartKind.Pareto && !share)
                    {
                        // the running total itself is drawn: it gets the unit of the columns
                        ax.TickLabels.NumberFormat = pc.AxisFormat;
                        if (pc.DisplayUnit != 0)
                        {
                            ax.DisplayUnit = pc.DisplayUnit;
                            ax.HasDisplayUnitLabel = false;
                        }
                    }
                    else
                    {
                        ax.TickLabels.NumberFormat = pc.Axis2Format;
                        if (share)
                        {
                            ax.MinimumScale = 0;
                            ax.MaximumScale = 1;
                        }
                        else if (pc.DisplayUnit2 != 0)
                        {
                            ax.DisplayUnit = pc.DisplayUnit2;
                            ax.HasDisplayUnitLabel = false;
                        }
                    }
                });
            }
        }

        // A "table with bars" is drawn in cells: every row reads a row of the pivot table, so it follows the filters too.
        static void TableCells(Job j, Box b, int i)
        {
            dynamic ws = j.Ws;
            DashSpec s = j.S;
            ThemeSpec th = s.Theme;
            ChartSpec c = s.Charts[i];
            PlanChart pc = j.Plan.Charts[i];
            PlanPivot pv = j.Plan.Pivots[pc.Pivot];
            AggResult res = pc.Result;
            Rng(ws, b.X, b.Y, b.W, b.H).Interior.Color = th.Card;
            TextCells(ws, b.X + 1, b.Y, b.W - 2, 3, pc.Title, false, DashLook.ChartTitlePt, true, th.Text, C.xlLeft, C.xlCenter);
            if (pv.FirstRow <= 0) return;
            int values = res.S.Count - 1;                         // the last series is the share
            if (values < 1) return;
            int inner = b.W - 2;
            int nameW = (int)Math.Round(inner * (values >= 2 ? 0.34 : 0.42));
            int colW = (inner - nameW) / (values + 1);
            if (colW < 3) return;
            nameW = inner - colW * (values + 1);
            int fit = (b.H - 3 - 2 - 1) / 2;
            int rows = Math.Min(res.Cats.Count, fit);
            if (rows < 1) return;
            int x0 = b.X + 1, y0 = b.Y + 3;

            // ---- headings ----
            TextCells(ws, x0, y0, nameW, 2, DashText.DimName(c.Dim).ToUpperInvariant(), false, 8, true, th.Muted, C.xlLeft, C.xlCenter);
            for (int v = 0; v < values; v++)
                TextCells(ws, x0 + nameW + colW * v, y0, colW, 2, res.S[v].Name.ToUpperInvariant(), false, 8, true, th.Muted, C.xlAlignRight, C.xlCenter);
            TextCells(ws, x0 + nameW + colW * values, y0, colW, 2, "SHARE", false, 8, true, th.Muted, C.xlAlignRight, C.xlCenter);
            Edge(Rng(ws, x0, y0, inner, 2), C.xlEdgeBottom, th.Grid, C.xlThin);

            // ---- rows ----
            int last = pv.FirstRow + res.Cats.Count - 1;
            string labels = j.PRef + AbsRange(pv.FirstRow, pv.LabelCol, last, pv.LabelCol);
            for (int k = 0; k < rows; k++)
            {
                int y = y0 + 2 + 2 * k;
                TextCells(ws, x0, y, nameW, 2, DashFormula.TableName(labels, k + 1), true, DashLook.TablePt, false, th.Text, C.xlLeft, C.xlCenter);
                string nameCell = Abs(y + 1, x0 + 1);
                for (int v = 0; v < values; v++)
                {
                    string vals = j.PRef + AbsRange(pv.FirstRow, pv.ValueCol + v, last, pv.ValueCol + v);
                    double div = v == 0 ? pc.UnitDivisor : pc.Unit2Divisor;
                    dynamic cell = TextCells(ws, x0 + nameW + colW * v, y, colW, 2, DashFormula.TableValue(nameCell, vals, k + 1, div), true, DashLook.TablePt, v == 0, th.Text,
                                             C.xlAlignRight, C.xlCenter);
                    cell.NumberFormat = v == 0 ? pc.UnitFormat : pc.Unit2Format;
                }
                string first = j.PRef + AbsRange(pv.FirstRow, pv.ValueCol, last, pv.ValueCol);
                dynamic share = TextCells(ws, x0 + nameW + colW * values, y, colW, 2, DashFormula.TableShare(nameCell, first, k + 1), true, DashLook.TablePt, false, th.Muted,
                                          C.xlAlignRight, C.xlCenter);
                share.NumberFormat = "0.0%";
                Edge(Rng(ws, x0, y, inner, 2), C.xlEdgeBottom, th.Grid, C.xlThin);
            }

            // ---- bars behind the first number ----
            int accent = ColorMath.Mix(th.Card, DashLook.ChartColor(s, c, i), th.Dark ? 0.5 : 0.28);
            Xl.Try(delegate
            {
                dynamic area = Rng(ws, x0 + nameW, y0 + 2, colW, 2 * rows);
                dynamic bars = area.FormatConditions.AddDatabar();
                bars.BarColor.Color = accent;
                bars.MinPoint.Modify(C.xlConditionValueNumber, 0);
                bars.MaxPoint.Modify(C.xlConditionValueAutomaticMax);
                bars.BarFillType = 0;
                bars.BarBorder.Type = 0;
            });
        }

        // ---------------------------------------------------------------- filters
        // which buttons were switched on in the filters of a dashboard (only filters that were actually set)
        static Dictionary<string, List<string>> ReadPicks(dynamic wb, string sheetName)
        {
            Dictionary<string, List<string>> res = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            try
            {
                dynamic caches = wb.SlicerCaches;
                int n = (int)caches.Count;
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        dynamic sc = caches[i];
                        if (!OnSheet(sc, sheetName)) continue;
                        dynamic items = sc.SlicerItems;
                        int ni = (int)items.Count;
                        if (ni > 1000) continue;
                        List<string> chosen = new List<string>();
                        for (int k = 1; k <= ni; k++)
                        {
                            dynamic it = items[k];
                            if ((bool)it.Selected) chosen.Add(Xl.Str((object)it.Name));
                        }
                        if (chosen.Count > 0 && chosen.Count < ni) res[Xl.Str((object)sc.SourceName)] = chosen;
                    }
                    catch { }
                }
            }
            catch { }
            return res;
        }

        // true when the buttons of this filter sit on the sheet with this name
        static bool OnSheet(dynamic sc, string sheetName)
        {
            try
            {
                dynamic sls = sc.Slicers;
                if ((int)sls.Count == 0) return false;
                return string.Equals(Xl.Str((object)sls[1].Shape.Parent.Name), sheetName, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        static void DropSlicerCaches(dynamic wb, string sheetName)
        {
            try
            {
                dynamic caches = wb.SlicerCaches;
                for (int i = (int)caches.Count; i >= 1; i--)
                {
                    try
                    {
                        dynamic sc = caches[i];
                        if (OnSheet(sc, sheetName)) sc.Delete();
                    }
                    catch { }
                }
            }
            catch { }
        }

        static void ApplyPicks(dynamic sc, List<string> want)
        {
            HashSet<string> keep = new HashSet<string>(want, StringComparer.OrdinalIgnoreCase);
            dynamic items = sc.SlicerItems;
            int n = (int)items.Count;
            int found = 0;
            for (int k = 1; k <= n; k++) if (keep.Contains(Xl.Str((object)items[k].Name))) found++;
            // nothing of the old choice is left, or too many buttons to switch one by one
            if (found == 0 || found == n || n - found > 150) return;
            for (int k = 1; k <= n; k++)
            {
                dynamic it = items[k];
                if (!keep.Contains(Xl.Str((object)it.Name))) it.Selected = false;
            }
        }

        static void AddSlicers(Job j, Dictionary<string, List<string>> picks)
        {
            DashPlan plan = j.Plan;
            if (plan.Slicers.Count == 0) return;
            dynamic wb = j.Wb;
            PlanPivot kpi = plan.Pivots[plan.KpiPivot];
            dynamic first = kpi.Table;
            string style = null;
            try { style = SlicerStyle(wb, j.S); } catch { }
            if (style == null) style = DashLook.BuiltInSlicerStyle(j.S);
            bool failed = false;
            for (int i = 0; i < plan.Slicers.Count; i++)
            {
                Box b = j.Grid.Find(Box.Slicer, i);
                if (b == null) continue;
                PlanSlicer ps = plan.Slicers[i];
                dynamic sc = null;
                try { sc = wb.SlicerCaches.Add2(first, ps.Field); } catch { }
                if (sc == null) { try { sc = wb.SlicerCaches.Add(first, ps.Field); } catch { } }
                if (sc == null)
                {
                    failed = true;
                    continue;
                }
                // one filter drives every pivot table, so every card and chart follows it
                foreach (PlanPivot p in plan.Pivots)
                {
                    if (p == kpi || p.Table == null) continue;
                    dynamic pt = p.Table;
                    try { sc.PivotTables.AddPivotTable(pt); } catch { }
                }
                try
                {
                    Pts r = At(j.Ws, b.X, b.Y, b.W, b.H);
                    dynamic sl = sc.Slicers.Add(j.Ws, Xl.M, Xl.M, ps.Caption, r.T, r.L, r.W, r.H);
                    j.Filters++;
                    string st = style;
                    Xl.Try(delegate { sl.Style = st; });
                    if (j.S.SlicerPos == 1)
                    {
                        int columns = DashLook.SlicerColumns(ps.Items, b.W);
                        Xl.Try(delegate { sl.NumberOfColumns = columns; });
                    }
                    Xl.Try(delegate { sl.RowHeight = 15.75; });
                }
                catch (Exception ex)
                {
                    failed = true;
                    Problem(j, "The filter '" + ps.Caption + "' could not be placed.", ex);
                }
                List<string> want;
                if (picks != null && picks.TryGetValue(ps.Field, out want))
                {
                    try { ApplyPicks(sc, want); }
                    catch (Exception ex) { Problem(j, "The filter '" + ps.Caption + "' was reset: its old setting could not be put back.", ex); }
                }
            }
            if (failed && j.Filters == 0) Problem(j, "The filters could not be added. Filter buttons (slicers) need Excel 2010 or later.", null);
        }

        // A slicer style in the colours of the dashboard. Returns its name, or null when Excel did not take it
        // (the nearest style that comes with Excel is used then).
        static string SlicerStyle(dynamic wb, DashSpec s)
        {
            ThemeSpec th = s.Theme;
            int onFill = DashLook.SlicerOn(s), offFill = DashLook.SlicerOff(s), edge = th.Card;
            string name = "DocBrisk " + th.Name + " " + ((th.Card * 31 + th.Accent[0] * 17 + th.Text * 7 + th.Muted) & 0xFFFFFF).ToString("x6", Inv);
            try
            {
                dynamic have = wb.TableStyles[name];
                if (have != null) return name;
            }
            catch { }
            dynamic st = null;
            try
            {
                st = wb.TableStyles.Add(name);
                st.ShowAsAvailablePivotTableStyle = false;
                st.ShowAsAvailableTableStyle = false;
                st.ShowAsAvailableSlicerStyle = true;
                StyleElement(st, C.xlWholeTable, th.Card, th.Text, false, DashLook.SlicerEdge(s));
                StyleElement(st, C.xlHeaderRow, th.Card, th.Text, true, -1);
                StyleElement(st, C.xlSlicerSelectedItemWithData, onFill, DashLook.SlicerOnText(s), false, edge);
                StyleElement(st, C.xlSlicerSelectedItemWithNoData, ColorMath.Mix(onFill, th.Card, 0.5), th.Muted, false, edge);
                StyleElement(st, C.xlSlicerUnselectedItemWithData, offFill, DashLook.SlicerOffText(s), false, edge);
                StyleElement(st, C.xlSlicerUnselectedItemWithNoData, th.Card, ColorMath.Mix(th.Muted, th.Card, 0.45), false, edge);
                StyleElement(st, C.xlSlicerHoveredSelectedItemWithData, DashLook.SlicerHover(s), DashLook.SlicerOnText(s), false, edge);
                StyleElement(st, C.xlSlicerHoveredUnselectedItemWithData, ColorMath.Mix(offFill, th.Accent[0], 0.18), th.Text, false, edge);
                StyleElement(st, C.xlSlicerHoveredSelectedItemWithNoData, ColorMath.Mix(onFill, th.Card, 0.35), th.Muted, false, edge);
                StyleElement(st, C.xlSlicerHoveredUnselectedItemWithNoData, offFill, th.Muted, false, edge);
                // did Excel really take the colours?
                int back = Convert.ToInt32((object)st.TableStyleElements[C.xlSlicerSelectedItemWithData].Interior.Color, Inv);
                if (back != onFill) throw new InvalidOperationException("The slicer style was not taken.");
                return name;
            }
            catch
            {
                if (st != null)
                {
                    try { st.Delete(); } catch { }
                }
                return null;
            }
        }

        static void StyleElement(dynamic st, int which, int fill, int text, bool bold, int edge)
        {
            dynamic el = st.TableStyleElements[which];
            el.Interior.Color = fill;
            el.Font.Color = text;
            if (bold) el.Font.Bold = true;
            if (edge >= 0)
            {
                Xl.Try(delegate
                {
                    dynamic bs = el.Borders;
                    bs.LineStyle = C.xlContinuous;
                    bs.Weight = C.xlThin;
                    bs.Color = edge;
                });
            }
        }

        // ---------------------------------------------------------------- the window and printing
        static void FinishWindow(Job j)
        {
            DashSpec s = j.S;
            DashGrid g = j.Grid;
            Xl.Try(delegate { Xl.App.ActiveWindow.DisplayGridlines = false; });
            Xl.Try(delegate { Xl.App.ActiveWindow.DisplayHeadings = !s.HideHeadings; });
            // the zoom is set so that the dashboard fills the width of the window
            Xl.Try(delegate
            {
                dynamic win = Xl.App.ActiveWindow;
                double usable = Convert.ToDouble((object)win.UsableWidth, Inv);
                Pts all = At(j.Ws, 0, 0, g.Cols, 1);
                if (usable > 100 && all.W > 100)
                {
                    int zoom = (int)Math.Floor((usable - 24) / all.W * 20.0) * 5;
                    if (zoom < 60) zoom = 60;
                    if (zoom > 130) zoom = 130;
                    win.Zoom = zoom;
                }
            });
        }

        // one page wide, landscape; needs a printer to be installed, so every step may fail quietly
        static void PrintSetup(Job j)
        {
            dynamic ws = j.Ws;
            DashGrid g = j.Grid;
            Xl.Try(delegate { Xl.App.PrintCommunication = false; });
            Xl.Try(delegate
            {
                dynamic ps = ws.PageSetup;
                ps.PrintArea = (string)Rng(ws, 0, 0, g.Cols, g.Rows).Address;
                ps.Orientation = C.xlLandscape;
                ps.Zoom = false;
                ps.FitToPagesWide = 1;
                ps.FitToPagesTall = false;
                ps.CenterHorizontally = true;
                ps.LeftMargin = 18;
                ps.RightMargin = 18;
                ps.TopMargin = 18;
                ps.BottomMargin = 18;
            });
            Xl.Try(delegate { Xl.App.PrintCommunication = true; });
        }

        // =================================================================
        //  DATA HEALTH SCAN
        // =================================================================
        public static void HealthScan()
        {
            dynamic src = Xl.PickRange("Data Health Scan");
            if (src == null) return;
            int rows = (int)src.Rows.Count;
            if (rows < 2)
            {
                Ui.Info("Select a table with a header row and at least one data row.");
                return;
            }
            string msg;
            using (new Speed())
            {
                msg = HealthCore(src);
            }
            Ui.Info(msg);
        }

        static string HealthCore(dynamic src)
        {
            dynamic wb = src.Worksheet.Parent;
            Xl.Status("scanning...");
            object[,] v = Xl.ReadTyped(src);
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 0);
            int nr = v.GetLength(0) - hr, nc = v.GetLength(1);

            // repeated rows
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            int dupRows = 0;
            StringBuilder sb = new StringBuilder();
            for (int r = hr + 1; r <= v.GetLength(0); r++)
            {
                if (Xl.RowBlank(v, r)) continue;
                sb.Length = 0;
                for (int c = 1; c <= nc; c++) sb.Append(Xl.Key(v[r, c])).Append('\u001E');
                if (!seen.Add(sb.ToString())) dupRows++;
            }

            object[,] o = Xl.New2D(nc + 1, 12);
            string[] heads = { "Column", "Read as", "Filled", "Blanks", "Different values", "Numbers stored as text", "Dates stored as text",
                               "Extra spaces", "Error cells", "Smallest", "Largest", "What to do" };
            for (int j = 0; j < heads.Length; j++) o[1, j + 1] = heads[j];
            double issues = 0, fixable = 0;
            for (int c = 1; c <= nc; c++)
            {
                ColInfo k = ci[c];
                o[c + 1, 1] = k.Hdr;
                o[c + 1, 2] = k.KindLabel;
                if (nr > 0) o[c + 1, 3] = (double)(k.NonBlank - k.Errors) / nr;
                o[c + 1, 4] = (double)k.Blank;
                o[c + 1, 5] = (double)k.Distinct;
                o[c + 1, 6] = (double)k.TextNum;
                o[c + 1, 7] = (double)k.TextDate;
                o[c + 1, 8] = (double)k.Spaces;
                o[c + 1, 9] = (double)k.Errors;
                if (k.NumCount > 0) { o[c + 1, 10] = k.MinV; o[c + 1, 11] = k.MaxV; }
                string advice = "";
                if (k.Kind == "num" && k.TextNum > 0) advice += "Convert text to numbers; ";
                if (k.TextDate > 0) advice += "Convert text to dates; ";
                if (k.Spaces > 0) advice += "Trim spaces; ";
                if (k.Errors > 0) advice += "Fix formula errors; ";
                if (k.Kind == "empty") advice += "Empty column - delete it; ";
                else if (k.Blank > 0) advice += "Check the blanks; ";
                o[c + 1, 12] = advice.Length == 0 ? "OK" : advice.TrimEnd(' ', ';');
                issues += k.Blank + k.TextDate + k.Spaces + k.Errors;
                fixable += k.TextDate + k.Spaces;
                if (k.Kind == "num") { issues += k.TextNum; fixable += k.TextNum; }
            }
            double totalCells = (double)nr * nc;
            double score = totalCells > 0 ? 100 - 100 * issues / totalCells : 0;
            if (nr > 0) score -= Math.Min(10, 100.0 * dupRows / nr);
            score = Math.Max(0, Math.Round(score));
            string verdict = score >= 90 ? "Clean" : score >= 70 ? "Needs a little cleaning" : "Needs cleaning";

            dynamic ws = Xl.AddSheet(wb, "Data Health");
            ws.Range["A1"].Value2 = "Data Health: " + (string)src.Worksheet.Name + "!" + (string)src.Address[false, false];
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            object[,] top = Xl.New2D(5, 3);
            top[1, 1] = "Health score"; top[1, 2] = score; top[1, 3] = verdict;
            top[2, 1] = "Data rows"; top[2, 2] = (double)nr;
            top[3, 1] = "Repeated rows"; top[3, 2] = (double)dupRows;
            top[4, 1] = "Cells Smart Clean can fix"; top[4, 2] = fixable;
            top[5, 1] = "Header found in row"; top[5, 2] = (double)((int)src.Row + hr - 1);
            ws.Range["A2:C6"].Value2 = top;
            dynamic sc = ws.Range["B2"];
            sc.Font.Size = 16;
            sc.Font.Bold = true;
            sc.Font.Color = score >= 90 ? Design.Accent[1] : score >= 70 ? Design.Accent[2] : Design.Accent[4];
            Xl.WriteTable(ws.Range["A8"], o, 1);
            Xl.StyleHeader(ws.Range["A8:L8"]);
            ws.Range["C9"].Resize[nc, 1].NumberFormat = "0%";
            ws.Columns["A:L"].AutoFit();
            ws.Activate();
            return "Health score: " + score + " / 100  (" + verdict + ")\n\n" +
                   "Smart Clean can fix " + fixable.ToString("#,##0", Inv) + " cells automatically.\n" +
                   "The details are on the new 'Data Health' sheet.";
        }
    }
}
