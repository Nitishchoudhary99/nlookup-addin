using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SmartPro
{
    // EASY CUSTOMISATION for Dashboard Studio, for people who do not want to learn every setting:
    //   - the "Easy" tab (first tab): five plain questions on one screen; every answer shows at once in the preview
    //       1. the main number   2. break it down by (one chart each)   3. filter buttons   4. colours   5. title
    //   - right-click on any card, chart or filter in the preview: change it on the spot
    //       (kind of chart, group by, number, top 10, colour, width, move, copy, remove)
    //   - Undo / Redo for every change (buttons, Ctrl+Z, Ctrl+Y)
    // The other tabs still hold every setting; this file only changes the same design (DashSpec) in fewer steps.
    internal sealed partial class StudioForm
    {
        // tab positions: Easy is first
        const int TabEasy = 0, TabDesigns = 1, TabCards = 2, TabCharts = 3, TabFilters = 4, TabLook = 5, TabPage = 6;

        // Easy tab controls
        ComboBox easyNumber, easyHow;
        CheckedListBox easyBreak, easyFilters;
        TextBox easyTitle;
        readonly List<string> easyFields = new List<string>();      // what the two tick lists offer: date parts and columns
        readonly Button[] easyThemes = new Button[8];

        // undo / redo: the design as text before each change
        readonly Stack<string> undo = new Stack<string>(), redo = new Stack<string>();
        string lastText;
        DateTime lastPush = DateTime.MinValue;
        bool restoring;
        Button undoButton, redoButton;

        static readonly string[] HowNames = { "Total", "Average", "Count of rows" };
        ComboBox cardOp, cardCol2;

        // the number columns, with the rows first
        void FillNumbers(ComboBox box, string current)
        {
            List<string> items = new List<string>();
            items.Add(RowsChoice);
            items.AddRange(numberCols);
            Fill(box, items);
            int at = 0;
            for (int i = 1; i < items.Count; i++) if (string.Equals(items[i], current, StringComparison.OrdinalIgnoreCase)) at = i;
            Pick(box, at);
        }

        // ------------------------------------------------------------------ the Easy tab
        void BuildEasyTab()
        {
            NewTab("Easy");
            Note("Answer these and watch the preview. Right-click any chart or card in the preview to change it. Ctrl+Z undoes.", 2);

            Heading("1.  Which number matters most?");
            List<string> nums = new List<string>(numberCols);
            nums.Add(RowsChoice);
            easyNumber = Combo("Number", nums.ToArray(), EasyNumberChanged);
            easyHow = Combo("Show it as", HowNames, EasyNumberChanged);

            Heading("2.  Break it down by  (one chart each)");
            foreach (string f in groupFields)
            {
                // the useful date parts only; the rest stay on the Charts tab
                if (DatePart.Is(f) && f != DatePart.Month && f != DatePart.Quarter && f != DatePart.FiscalYear) continue;
                if (!DatePart.Is(f) && !EasyField(f)) continue;
                easyFields.Add(f);
            }
            easyBreak = TickList(Math.Min(S(150), S(22) * Math.Max(3, easyFields.Count) + S(6)), EasyBreakChanged);

            Heading("3.  Filter buttons for");
            easyFilters = TickList(Math.Min(S(130), S(22) * Math.Max(3, easyFields.Count) + S(6)), EasyFiltersChanged);

            Heading("4.  Colours");
            int n = Math.Min(easyThemes.Length, DashThemes.Count);
            for (int i = 0; i < n; i++)
            {
                ThemeSpec th = DashThemes.Get(i);
                Button b = new Button();
                b.FlatStyle = FlatStyle.Flat;
                b.FlatAppearance.BorderSize = 2;
                b.Cursor = Cursors.Hand;
                b.SetBounds(S(10) + (i % 4) * S(106), cy + (i / 4) * S(46), S(100), S(40));
                b.BackColor = Rgb(th.Band);
                b.ForeColor = Rgb(th.BandText);
                b.Text = th.Name;
                b.Font = new Font("Segoe UI", 7.5f, FontStyle.Bold);
                int which = i;
                b.Paint += delegate (object sender, PaintEventArgs e)
                {
                    // the theme's six colours as a strip at the bottom of the button
                    int w = Math.Max(1, (b.Width - 8) / 6);
                    for (int k = 0; k < 6; k++)
                        using (SolidBrush br = new SolidBrush(Rgb(th.Accent[k]))) e.Graphics.FillRectangle(br, 4 + k * w, b.Height - 9, w - 1, 5);
                };
                b.Click += delegate { Do(delegate { lookTheme.SelectedIndex = which; EasyThemeMarks(); }); };
                tips.SetToolTip(b, th.Name + (th.Dark ? " (dark)" : ""));
                page.Controls.Add(b);
                easyThemes[i] = b;
            }
            cy += S(96);
            Note("More themes, your own colours and fonts: the Look tab.", 1);

            Heading("5.  Title");
            easyTitle = Input("Title", delegate { if (pgTitle != null && pgTitle.Text != easyTitle.Text) pgTitle.Text = easyTitle.Text; });
            EndTab();
        }

        CheckedListBox TickList(int height, Action changed)
        {
            CheckedListBox l = new CheckedListBox();
            l.CheckOnClick = true;
            l.IntegralHeight = false;
            l.SetBounds(S(10), cy, S(Edge - 10), height);
            foreach (string f in easyFields) l.Items.Add(FieldLabel(f));
            page.Controls.Add(l);
            // ItemCheck fires before the tick changes: act after it
            l.ItemCheck += delegate { if (!loading) BeginInvoke((MethodInvoker)delegate { Do(changed); }); };
            cy += height + S(8);
            return l;
        }

        // the Easy lists offer columns people group by: not numbers (they are added up) and not IDs with many values
        bool EasyField(string name)
        {
            foreach (DataCol c in t.Cols)
            {
                if (!string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)) continue;
                if (c.IsNum) return false;
                if (c.Kind == "id") return false;
                return c.Distinct >= 2 && c.Distinct <= 60;
            }
            return true;
        }

        // "Order ID" -> "Orders", "Invoice No" -> "Invoices": the label of a count of different values
        internal static string CountLabel(string col)
        {
            string s = (col ?? "").Trim();
            string[] tails = { " ID", " Id", " id", " No", " No.", " Number", " Num", " Code", " #", "ID", "No" };
            foreach (string tail in tails)
            {
                if (s.Length > tail.Length + 1 && s.EndsWith(tail, StringComparison.Ordinal))
                {
                    s = s.Substring(0, s.Length - tail.Length).Trim();
                    break;
                }
            }
            if (s.Length == 0) return "Different values";
            if (s.EndsWith("s", StringComparison.OrdinalIgnoreCase)) return s;
            if (s.Length > 2 && s.EndsWith("y", StringComparison.OrdinalIgnoreCase) && "aeiouAEIOU".IndexOf(s[s.Length - 2]) < 0)
                return s.Substring(0, s.Length - 1) + "ies";
            return s + "s";
        }

        string FieldLabel(string field)
        {
            int i = GroupIndex(field);
            return i >= 0 && i < groupLabels.Count ? groupLabels[i] : field;
        }

        // the main number: the first card and every chart that showed the old main number follow it
        void EasyNumberChanged()
        {
            if (easyNumber.SelectedIndex < 0) return;
            string col = (string)easyNumber.Items[easyNumber.SelectedIndex];
            if (col == RowsChoice) col = "";
            int agg = col.Length == 0 || easyHow.SelectedIndex == 2 ? AggKind.Count : easyHow.SelectedIndex == 1 ? AggKind.Avg : AggKind.Sum;
            string oldCol = s.Cards.Count > 0 ? s.Cards[0].Col : (s.Charts.Count > 0 ? s.Charts[0].Col : "");
            int oldAgg = s.Cards.Count > 0 ? s.Cards[0].Agg : AggKind.Sum;
            if (s.Cards.Count == 0) s.Cards.Add(new CardSpec());
            // another card already shows the new number: the two swap, so no number is shown twice
            for (int k = 1; k < s.Cards.Count; k++)
            {
                if (!string.Equals(s.Cards[k].Col, col, StringComparison.OrdinalIgnoreCase) || s.Cards[k].Agg == AggKind.Distinct) continue;
                s.Cards[k].Col = oldCol;
                s.Cards[k].Agg = oldAgg;
                s.Cards[k].Label = "";
                break;
            }
            s.Cards[0].Col = col;
            s.Cards[0].Agg = agg;
            s.Cards[0].Label = "";
            foreach (ChartSpec c in s.Charts)
            {
                if (!string.Equals(c.Col, oldCol, StringComparison.OrdinalIgnoreCase)) continue;
                c.Col = col;
                c.Agg = agg;
                c.Title = "";
            }
            AfterQuickChange("The main number is now " + (agg == AggKind.Count ? "the count of rows" : HowNames[agg == AggKind.Avg ? 1 : 0].ToLowerInvariant() + " of " + col) + ".");
        }

        // one chart for every ticked field, in the order of the list; charts that stay keep their settings
        void EasyBreakChanged()
        {
            List<string> want = new List<string>();
            foreach (int i in easyBreak.CheckedIndices) want.Add(easyFields[i]);
            if (want.Count > DashSpec.MaxCharts)
            {
                SetHint("A dashboard takes " + DashSpec.MaxCharts + " charts at most.");
                want.RemoveRange(DashSpec.MaxCharts, want.Count - DashSpec.MaxCharts);
            }
            string col = s.Cards.Count > 0 ? s.Cards[0].Col : (numberCols.Count > 0 ? numberCols[0] : "");
            int agg = s.Cards.Count > 0 ? s.Cards[0].Agg : AggKind.Sum;
            List<ChartSpec> keep = new List<ChartSpec>();
            bool round = false;
            foreach (ChartSpec c in s.Charts) if (ChartKind.IsRound(c.Kind)) round = true;
            foreach (string f in want)
            {
                ChartSpec found = null;
                foreach (ChartSpec c in s.Charts) if (string.Equals(c.Dim, f, StringComparison.OrdinalIgnoreCase)) { found = c; break; }
                if (found == null)
                {
                    found = new ChartSpec();
                    found.Dim = f;
                    found.Col = col;
                    found.Agg = agg;
                    int distinct = Distinct(f);
                    if (DatePart.Is(f)) found.Kind = ChartKind.Line;
                    else if (distinct <= 6 && !round) { found.Kind = ChartKind.Doughnut; round = true; }
                    else found.Kind = distinct > 8 ? ChartKind.Bar : ChartKind.Column;
                    if (!DatePart.Is(f) && distinct > 12) found.Top = 10;
                }
                keep.Add(found);
            }
            // charts grouped by something the list does not offer (a split, a weekday ...) stay as they were
            foreach (ChartSpec c in s.Charts) if (easyFields.FindIndex(x => string.Equals(x, c.Dim, StringComparison.OrdinalIgnoreCase)) < 0 && keep.Count < DashSpec.MaxCharts) keep.Add(c);
            s.Charts.Clear();
            s.Charts.AddRange(keep);
            AfterQuickChange(s.Charts.Count + (s.Charts.Count == 1 ? " chart." : " charts."));
        }

        void EasyFiltersChanged()
        {
            List<string> want = new List<string>();
            foreach (int i in easyFilters.CheckedIndices) want.Add(easyFields[i]);
            if (want.Count > DashSpec.MaxSlicers)
            {
                SetHint("A dashboard takes " + DashSpec.MaxSlicers + " filters at most.");
                want.RemoveRange(DashSpec.MaxSlicers, want.Count - DashSpec.MaxSlicers);
            }
            List<string> others = new List<string>();
            foreach (string f in s.Slicers) if (easyFields.FindIndex(x => string.Equals(x, f, StringComparison.OrdinalIgnoreCase)) < 0) others.Add(f);
            s.Slicers.Clear();
            s.Slicers.AddRange(want);
            foreach (string f in others) if (s.Slicers.Count < DashSpec.MaxSlicers) s.Slicers.Add(f);
            if (s.SlicerPos == 3 && s.Slicers.Count > 0) s.SlicerPos = 0;
            AfterQuickChange(s.Slicers.Count == 0 ? "No filter buttons." : s.Slicers.Count + " filter " + (s.Slicers.Count == 1 ? "button." : "buttons."));
        }

        int Distinct(string field)
        {
            foreach (DataCol c in t.Cols) if (string.Equals(c.Name, field, StringComparison.OrdinalIgnoreCase)) return c.Distinct;
            return 12;
        }

        // a change made outside the normal tabs: every tab shows it, then the preview
        void AfterQuickChange(string hintText)
        {
            AllToControls();
            Changed();
            if (hintText != null) SetHint(hintText + "  Ctrl+Z undoes it.");
        }

        // the Easy tab shows the design as it is now (called from AllToControls)
        void EasyToControls()
        {
            if (easyNumber == null) return;
            bool was = loading;
            loading = true;
            string col = s.Cards.Count > 0 ? s.Cards[0].Col : "";
            int agg = s.Cards.Count > 0 ? s.Cards[0].Agg : AggKind.Count;
            int at = easyNumber.Items.Count - 1;
            for (int i = 0; i < easyNumber.Items.Count; i++) if (string.Equals((string)easyNumber.Items[i], col, StringComparison.OrdinalIgnoreCase)) at = i;
            Pick(easyNumber, at);
            Pick(easyHow, agg == AggKind.Avg ? 1 : agg == AggKind.Sum ? 0 : 2);
            for (int i = 0; i < easyFields.Count; i++)
            {
                bool chart = false, filter = false;
                foreach (ChartSpec c in s.Charts) if (string.Equals(c.Dim, easyFields[i], StringComparison.OrdinalIgnoreCase)) chart = true;
                foreach (string f in s.Slicers) if (string.Equals(f, easyFields[i], StringComparison.OrdinalIgnoreCase)) filter = true;
                easyBreak.SetItemChecked(i, chart);
                easyFilters.SetItemChecked(i, filter);
            }
            if (easyTitle.Text != s.Title) easyTitle.Text = s.Title;
            loading = was;
            EasyThemeMarks();
        }

        void EasyThemeMarks()
        {
            for (int i = 0; i < easyThemes.Length; i++)
            {
                if (easyThemes[i] == null) continue;
                bool on = string.Equals(DashThemes.NameOf(i), s.Theme.Name, StringComparison.OrdinalIgnoreCase);
                easyThemes[i].FlatAppearance.BorderColor = on ? Color.FromArgb(250, 204, 21) : Color.FromArgb(203, 213, 225);
                easyThemes[i].FlatAppearance.BorderSize = on ? 4 : 1;
            }
        }

        // ------------------------------------------------------------------ undo / redo
        void UndoButtons(Control parent)
        {
            undoButton = MakeButton(parent, "Undo", 0, S(10), S(70));
            redoButton = MakeButton(parent, "Redo", 0, S(10), S(70));
            tips.SetToolTip(undoButton, "Undo the last change (Ctrl+Z)");
            tips.SetToolTip(redoButton, "Redo (Ctrl+Y)");
            undoButton.Click += delegate { Do(Undo); };
            redoButton.Click += delegate { Do(Redo); };
            lastText = startText;
            UndoState();
        }

        // called on every change of the design; quick changes in a row (typing a title) are one step
        void RecordChange()
        {
            if (restoring || lastText == null) return;
            string now = s.ToText();
            if (now == lastText) return;
            if ((DateTime.Now - lastPush).TotalMilliseconds > 800 || undo.Count == 0)
            {
                undo.Push(lastText);
                if (undo.Count > 200) TrimStack(undo, 150);
                lastPush = DateTime.Now;
            }
            redo.Clear();
            lastText = now;
            UndoState();
        }

        static void TrimStack(Stack<string> st, int keep)
        {
            string[] a = st.ToArray();          // newest first
            st.Clear();
            for (int i = Math.Min(keep, a.Length) - 1; i >= 0; i--) st.Push(a[i]);
        }

        void Undo()
        {
            if (undo.Count == 0) { SetHint("Nothing to undo."); return; }
            redo.Push(lastText);
            Restore(undo.Pop(), "Undone.");
        }

        void Redo()
        {
            if (redo.Count == 0) { SetHint("Nothing to redo."); return; }
            undo.Push(lastText);
            Restore(redo.Pop(), "Done again.");
        }

        void Restore(string text, string hintText)
        {
            DashSpec back = DashSpec.FromText(text);
            if (back == null) return;
            restoring = true;
            try
            {
                s = back;
                DashAuto.Fix(s, t);
                lastText = text;
                lastPush = DateTime.MinValue;
                selKind = -1;
                selIndex = -1;
                AllToControls();
                Changed();
            }
            finally { restoring = false; }
            UndoState();
            SetHint(hintText);
        }

        void UndoState()
        {
            if (undoButton == null) return;
            undoButton.Enabled = undo.Count > 0;
            redoButton.Enabled = redo.Count > 0;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // a text box keeps its own Ctrl+Z
            if (!(ActiveControl is TextBox))
            {
                if (keyData == (Keys.Control | Keys.Z)) { Do(Undo); return true; }
                if (keyData == (Keys.Control | Keys.Y) || keyData == (Keys.Control | Keys.Shift | Keys.Z)) { Do(Redo); return true; }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ------------------------------------------------------------------ right-click in the preview
        internal void PreviewMenu(int kind, int index, Point screen)
        {
            if (loading) return;
            ContextMenuStrip m = new ContextMenuStrip();
            m.Font = new Font("Segoe UI", 9.5f);
            if (kind == Box.Chart && index >= 0 && index < s.Charts.Count) ChartMenu(m, index);
            else if (kind == Box.Card && index >= 0 && index < s.Cards.Count) CardMenu(m, index);
            else if (kind == Box.Slicer && index >= 0 && index < s.Slicers.Count)
            {
                string f = s.Slicers[index];
                Item(m.Items, "Remove the filter '" + FieldName(f) + "'", false, delegate { s.Slicers.RemoveAt(index); AfterQuickChange("Filter removed."); });
                Item(m.Items, "Filter settings...", false, delegate { tabs.SelectedIndex = TabFilters; });
            }
            else if (kind == Box.Header)
            {
                Item(m.Items, "Change the title...", false, delegate { tabs.SelectedIndex = TabEasy; easyTitle.Focus(); easyTitle.SelectAll(); });
                Item(m.Items, "Logo, footer and page...", false, delegate { tabs.SelectedIndex = TabPage; });
            }
            if (m.Items.Count > 0) m.Items.Add(new ToolStripSeparator());
            Item(m.Items, "Add a chart", s.Charts.Count >= DashSpec.MaxCharts, delegate { tabs.SelectedIndex = TabCharts; ChartAdd(); });
            Item(m.Items, "Add a KPI card", s.Cards.Count >= DashSpec.MaxCards, delegate { tabs.SelectedIndex = TabCards; CardAdd(); });
            Item(m.Items, "Undo   (Ctrl+Z)", undo.Count == 0, Undo);
            if (kind >= 0) SelectBox(kind, index);
            m.Closed += delegate { BeginInvoke((MethodInvoker)delegate { m.Dispose(); }); };
            m.Show(screen);
        }

        void ChartMenu(ContextMenuStrip m, int i)
        {
            ChartSpec c = s.Charts[i];
            m.Items.Add(new ToolStripLabel(ChartName(c)) { Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) });
            ToolStripMenuItem kinds = Sub(m.Items, "Chart type");
            int[] simple = { ChartKind.Column, ChartKind.Bar, ChartKind.Line, ChartKind.Area, ChartKind.Pie, ChartKind.Doughnut, ChartKind.Pareto, ChartKind.Table };
            foreach (int k in simple)
            {
                int kk = k;
                Item(kinds.DropDownItems, ChartKind.Names[k], false, delegate { c.Kind = kk; c.Title = ""; AfterQuickChange("Now a " + ChartKind.Names[kk].ToLowerInvariant() + " chart."); }).Checked = c.Kind == k;
            }
            ToolStripMenuItem by = Sub(m.Items, "Group by");
            foreach (string f in groupFields)
            {
                string ff = f;
                Item(by.DropDownItems, FieldLabel(f), false, delegate { c.Dim = ff; c.Title = ""; if (DatePart.Is(ff) && ChartKind.IsRound(c.Kind)) c.Kind = ChartKind.Line; AfterQuickChange("Grouped by " + FieldName(ff) + "."); }).Checked = string.Equals(c.Dim, f, StringComparison.OrdinalIgnoreCase);
            }
            ToolStripMenuItem num = Sub(m.Items, "Number");
            foreach (string col in numberCols)
            {
                string cc = col;
                Item(num.DropDownItems, col, false, delegate { c.Col = cc; if (c.Agg == AggKind.Count) c.Agg = AggKind.Sum; c.Title = ""; AfterQuickChange("Shows " + cc + "."); }).Checked = string.Equals(c.Col, col, StringComparison.OrdinalIgnoreCase);
            }
            Item(num.DropDownItems, "Count of rows", false, delegate { c.Col = ""; c.Agg = AggKind.Count; c.Title = ""; AfterQuickChange("Shows the count of rows."); }).Checked = c.Col.Length == 0;
            if (!DatePart.Is(c.Dim))
            {
                ToolStripMenuItem top = Sub(m.Items, "How many");
                foreach (int n in new[] { 5, 10, 15, -1, 0 })
                {
                    int nn = n;
                    Item(top.DropDownItems, n == 0 ? "Automatic" : n < 0 ? "All of them" : "Top " + n, false, delegate { c.Top = nn; AfterQuickChange(null); }).Checked = c.Top == n;
                }
            }
            ToolStripMenuItem col6 = Sub(m.Items, "Colour");
            ColourItems(col6, c.Color, delegate (int v) { c.Color = v; AfterQuickChange("Colour changed."); });
            ToolStripMenuItem wid = Sub(m.Items, "Width");
            string[] wn = { "Automatic", "A quarter", "A third", "Half", "Two thirds", "Full width" };
            int[] wv = { 0, 3, 4, 6, 8, 12 };
            for (int k = 0; k < wn.Length; k++)
            {
                int v = wv[k];
                Item(wid.DropDownItems, wn[k], false, delegate { c.Span = v; AfterQuickChange(null); }).Checked = c.Span == v;
            }
            Item(m.Items, "Show values on it", false, delegate { c.Labels = c.Labels == 2 ? 1 : 2; AfterQuickChange(null); }).Checked = c.Labels == 2;
            m.Items.Add(new ToolStripSeparator());
            Item(m.Items, "Move earlier", i == 0, delegate { Swap(s.Charts, i, i - 1); AfterQuickChange("Moved."); });
            Item(m.Items, "Move later", i == s.Charts.Count - 1, delegate { Swap(s.Charts, i, i + 1); AfterQuickChange("Moved."); });
            Item(m.Items, "Copy this chart", s.Charts.Count >= DashSpec.MaxCharts, delegate { s.Charts.Insert(i + 1, c.Clone()); AfterQuickChange("Copied: change the copy with a right-click."); });
            Item(m.Items, "Remove this chart", false, delegate { s.Charts.RemoveAt(i); selKind = -1; AfterQuickChange("Chart removed."); });
            Item(m.Items, "All chart settings...", false, delegate { tabs.SelectedIndex = TabCharts; if (i < chartList.Items.Count) chartList.SelectedIndex = i; });
        }

        void CardMenu(ContextMenuStrip m, int i)
        {
            CardSpec c = s.Cards[i];
            m.Items.Add(new ToolStripLabel(c.Col.Length == 0 ? "Count of rows" : (c.Agg == AggKind.Distinct ? "Different values of " : AggKind.Words[c.Agg] + " ") + c.Col) { Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) });
            ToolStripMenuItem num = Sub(m.Items, "Number");
            foreach (string col in numberCols)
            {
                string cc = col;
                Item(num.DropDownItems, col, false, delegate { c.Col = cc; if (c.Agg == AggKind.Count) c.Agg = AggKind.Sum; c.Label = ""; AfterQuickChange(null); }).Checked = string.Equals(c.Col, col, StringComparison.OrdinalIgnoreCase);
            }
            Item(num.DropDownItems, "Count of rows", false, delegate { c.Col = ""; c.Agg = AggKind.Count; c.Label = ""; AfterQuickChange(null); }).Checked = c.Col.Length == 0;
            if (c.Col.Length > 0)
            {
                ToolStripMenuItem how = Sub(m.Items, "Show as");
                foreach (int a in new[] { AggKind.Sum, AggKind.Avg, AggKind.Max, AggKind.Min })
                {
                    int aa = a;
                    Item(how.DropDownItems, AggKind.Names[a], false, delegate { c.Agg = aa; c.Label = ""; AfterQuickChange(null); }).Checked = c.Agg == a;
                }
            }
            // count of different values of any column, e.g. Order ID -> number of orders
            ToolStripMenuItem diff = Sub(m.Items, "Count different values of");
            foreach (string col in anyCols)
            {
                string cc = col;
                Item(diff.DropDownItems, col, false, delegate { c.Col = cc; c.Agg = AggKind.Distinct; c.Label = CountLabel(cc); c.Compare = 0; c.Spark = false; AfterQuickChange("The card counts the different values of " + cc + ": " + CountLabel(cc) + "."); }).Checked = c.Agg == AggKind.Distinct && string.Equals(c.Col, col, StringComparison.OrdinalIgnoreCase);
            }
            if (t.HasDate)
            {
                Item(m.Items, "Compare with the month before", false, delegate { c.Compare = c.Compare == 1 ? 0 : 1; AfterQuickChange(null); }).Checked = c.Compare == 1;
                Item(m.Items, "Compare with the same month last year", false, delegate { c.Compare = c.Compare == 2 ? 0 : 2; AfterQuickChange(null); }).Checked = c.Compare == 2;
                Item(m.Items, "Small trend line", false, delegate { c.Spark = !c.Spark; AfterQuickChange(null); }).Checked = c.Spark;
            }
            // calculate with a second number: Profit % of Net Sales, Gross Sales - Discount ...
            if (c.Col.Length > 0 && c.Agg != AggKind.Distinct)
            {
                ToolStripMenuItem calc = Sub(m.Items, "Calculate");
                string[] how = { "", "As % of", "Divided by", "Minus", "Plus" };
                for (int op = 1; op <= 4; op++)
                {
                    ToolStripMenuItem o = Sub(calc.DropDownItems, how[op]);
                    int oo = op;
                    foreach (string col in numberCols)
                    {
                        if (string.Equals(col, c.Col, StringComparison.OrdinalIgnoreCase)) continue;
                        string cc = col;
                        Item(o.DropDownItems, col, false, delegate
                        {
                            c.Op = oo;
                            c.Col2 = cc;
                            c.Agg2 = Logic.MeasureIsAvg(cc) ? AggKind.Avg : AggKind.Sum;
                            c.Label = "";
                            c.Compare = 0;
                            c.Spark = false;
                            AfterQuickChange("The card now shows " + DashText.CardLabel(c) + ".");
                        }).Checked = c.Op == oo && string.Equals(c.Col2, col, StringComparison.OrdinalIgnoreCase);
                    }
                }
                if (c.Op > 0) Item(calc.DropDownItems, "No calculation", false, delegate { c.Op = CardOp.None; c.Label = ""; AfterQuickChange(null); });
            }
            ToolStripMenuItem col6 = Sub(m.Items, "Colour");
            ColourItems(col6, c.Color, delegate (int v) { c.Color = v; AfterQuickChange("Colour changed."); });
            m.Items.Add(new ToolStripSeparator());
            Item(m.Items, "Move earlier", i == 0, delegate { Swap(s.Cards, i, i - 1); AfterQuickChange("Moved."); });
            Item(m.Items, "Move later", i == s.Cards.Count - 1, delegate { Swap(s.Cards, i, i + 1); AfterQuickChange("Moved."); });
            Item(m.Items, "Remove this card", false, delegate { s.Cards.RemoveAt(i); selKind = -1; AfterQuickChange("Card removed."); });
            Item(m.Items, "All card settings...", false, delegate { tabs.SelectedIndex = TabCards; if (i < cardList.Items.Count) cardList.SelectedIndex = i; });
        }

        void ColourItems(ToolStripMenuItem parent, int current, Action<int> set)
        {
            Item(parent.DropDownItems, "Automatic", false, delegate { set(-1); }).Checked = current < 0;
            for (int k = 0; k < 6; k++)
            {
                int kk = k;
                ToolStripMenuItem it = Item(parent.DropDownItems, "Colour " + (k + 1), false, delegate { set(kk); });
                it.Checked = current == k;
                Bitmap dot = new Bitmap(14, 14);
                using (Graphics g = Graphics.FromImage(dot))
                using (SolidBrush br = new SolidBrush(Rgb(s.Theme.Accent[k]))) g.FillRectangle(br, 0, 0, 14, 14);
                it.Image = dot;
            }
        }

        static ToolStripMenuItem Sub(ToolStripItemCollection items, string text)
        {
            ToolStripMenuItem it = new ToolStripMenuItem(text);
            items.Add(it);
            return it;
        }

        ToolStripMenuItem Item(ToolStripItemCollection items, string text, bool disabled, Action act)
        {
            ToolStripMenuItem it = new ToolStripMenuItem(text);
            it.Enabled = !disabled;
            it.Click += delegate { Do(act); };
            items.Add(it);
            return it;
        }

        static void Swap<T>(List<T> l, int a, int b)
        {
            if (a < 0 || b < 0 || a >= l.Count || b >= l.Count) return;
            T x = l[a];
            l[a] = l[b];
            l[b] = x;
        }

        static string FieldName(string f) { return DatePart.Is(f) ? DatePart.Label(f) : f; }

        static string ChartName(ChartSpec c)
        {
            string what = c.Col.Length == 0 ? "Rows" : c.Col;
            return ChartKind.Names[c.Kind] + ": " + what + " by " + FieldName(c.Dim);
        }
    }
}
