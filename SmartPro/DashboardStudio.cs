using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace SmartPro
{
    // =====================================================================
    //  DASHBOARD STUDIO
    //  One window to design a dashboard: pick a ready design made from your own table, then change anything:
    //  the KPI cards, the charts, the filters, the colours and the page. The preview on the right is drawn from
    //  the real numbers and shows what will be built. A design can be saved and used again on another table.
    // =====================================================================
    internal static class DashboardStudio
    {
        // ribbon button: design a new dashboard for a table
        public static void Run()
        {
            dynamic src = Xl.PickRange("Dashboard Studio");
            if (src == null) return;
            DataTab t = Dashboard.Load(src);
            if (t == null) return;
            string sheet = (string)src.Worksheet.Name;
            DashSpec s = DashAuto.Suggest(t, DashAuto.SavedLook());
            s.Title = DashText.AutoTitle(sheet, s);
            DashSpec done = Open(t, s, sheet + "!" + (string)src.Address[false, false], false);
            if (done == null) return;
            string msg;
            using (new Speed())
            {
                msg = Dashboard.Create(src, t, done, null);
            }
            Ui.Info("Dashboard ready.\n\n" + msg + "\n\nTo change it later: open the dashboard sheet and click Edit Dashboard.");
        }

        // shows the designer; returns the design to build, or null when the window was closed without building
        public static DashSpec Open(DataTab t, DashSpec spec, string source, bool editing)
        {
            return Show(t, spec, source, editing, editing ? "Update dashboard" : "Build dashboard");
        }

        // the same window for another job (the web dashboard): okText is written on the button that closes it
        public static DashSpec Open(DataTab t, DashSpec spec, string source, string okText)
        {
            return Show(t, spec, source, false, okText);
        }

        static DashSpec Show(DataTab t, DashSpec spec, string source, bool editing, string okText)
        {
            using (StudioForm f = new StudioForm(t, spec, source, editing, okText))
            {
                return f.ShowDialog(new ExcelWin()) == DialogResult.OK ? f.Result : null;
            }
        }

        // where saved designs are kept
        public static string DesignFolder()
        {
            string f = Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DocBrisk"), "Dashboards");
            try { Directory.CreateDirectory(f); } catch { }
            return f;
        }
    }

    internal sealed class StudioForm : Form
    {
        // ---- the design and the table ----
        readonly DataTab t;
        DashSpec s;
        readonly bool editing;
        readonly DashCache cache = new DashCache();
        readonly string startText;                  // the design as it was when the window opened
        public DashSpec Result;
        bool loading;                               // the controls are being filled: their change events do nothing
        float z = 1f;                               // screen scale: 1 at 100%, 1.5 at 150%
        int selKind = -1, selIndex = -1;            // what is selected in the preview
        int version;                                // goes up whenever the design changes in a way the small pictures show
        Image logo;
        string logoOf = "";
        readonly Font bold;

        // ---- columns the pickers offer ----
        readonly List<string> anyCols = new List<string>();         // every column
        readonly List<string> numberCols = new List<string>();      // columns that can be added up
        readonly List<string> groupFields = new List<string>();     // columns and date parts to group by
        readonly List<string> groupLabels = new List<string>();
        readonly List<int> topChoices = new List<int>();

        // ---- controls ----
        readonly TabControl tabs = new TabControl();
        readonly PreviewPanel preview;
        readonly GalleryPanel gallery;
        readonly Panel bottom = new Panel();
        readonly Label hint = new Label();
        readonly ToolTip tips = new ToolTip();
        Button okButton, cancelButton, saveButton, openButton, againButton;
        TabPage page;                               // the tab that is being filled
        int cy;                                     // the next free line on it

        ListBox cardList, chartList;
        ComboBox cardCol, cardAgg, cardFmt, cardDec, cardColor, cardCompare;
        TextBox cardLabel, cardPrefix, cardSuffix, cardTarget;
        CheckBox cardSpark;
        Panel cardSwatch;
        ComboBox chKind, chDim, chCol, chAgg, chSplit, chCol2, chAgg2, chTop, chOthers, chSort, chShowAs, chLabels, chLegend, chSpan, chHeight, chColor, chTrend;
        TextBox chTitle;
        CheckBox chSmooth;
        Panel chSwatch;
        Label chNote;
        CheckedListBox filterList;
        ComboBox filterPos, dateCol, fyStart;
        ComboBox lookTheme, lookCards, lookFont, lookDensity, lookUnit;
        CheckBox lookRounded, lookShadow;
        TextBox lookCurrency;
        readonly Button[] swatches = new Button[10];
        TextBox pgTitle, pgSubtitle, pgFooter;
        ComboBox pgArrange, pgWidth, pgPerRow;
        CheckBox pgInsights, pgHeadings, pgPrint;
        Label pgLogo;

        static readonly string[] SwatchNames = { "Colour 1", "Colour 2", "Colour 3", "Colour 4", "Colour 5", "Colour 6", "Page", "Title band", "Cards", "Text" };
        static readonly string[] ColorNames = { "Automatic", "Colour 1", "Colour 2", "Colour 3", "Colour 4", "Colour 5", "Colour 6", "Own colour..." };
        const string RowsChoice = "(Rows of the table)";
        const string NoneChoice = "(none)";
        const int Edge = 436;                       // right edge of the controls on a tab, at 100%

        public StudioForm(DataTab table, DashSpec spec, string sourceText, bool isEditing, string okText)
        {
            t = table;
            editing = isEditing;
            s = spec.Clone();
            cache.Source = sourceText;
            DashAuto.Fix(s, t);
            startText = s.ToText();

            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                z = g.DpiX / 96f;
            }
            if (z < 1f) z = 1f;
            if (z > 4f) z = 4f;
            Text = Design.AppName + " - Dashboard Studio  -  " + sourceText + "  (" + t.Rows.ToString("#,##0", CultureInfo.InvariantCulture) + " rows)";
            Font = new Font("Segoe UI", 9f);
            bold = new Font(Font, FontStyle.Bold);
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ShowIcon = false;
            Rectangle area = Screen.PrimaryScreen.WorkingArea;
            try { area = Screen.FromHandle(new ExcelWin().Handle).WorkingArea; } catch { }
            int w = Math.Min(S(1380), area.Width - S(24)), h = Math.Min(S(830), area.Height - S(46));
            ClientSize = new Size(Math.Max(w, 700), Math.Max(h, 460));
            MinimumSize = new Size(Math.Min(S(940), area.Width), Math.Min(S(560), area.Height));
            StartPosition = FormStartPosition.Manual;
            Location = new Point(area.X + Math.Max(0, (area.Width - Width) / 2), area.Y + Math.Max(0, (area.Height - Height) / 2));

            FillColumnLists();

            // ---- right: the live preview (added first, so that it gets the room the other two leave) ----
            preview = new PreviewPanel(this);
            preview.Dock = DockStyle.Fill;
            Controls.Add(preview);

            // ---- left: the tabs ----
            tabs.Dock = DockStyle.Left;
            tabs.Width = S(468);
            Controls.Add(tabs);
            gallery = new GalleryPanel(this, t.HasDate);
            BuildDesignsTab();
            BuildCardsTab();
            BuildChartsTab();
            BuildFiltersTab();
            BuildLookTab();
            BuildPageTab();

            // ---- bottom: a line of help and the buttons ----
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = S(50);
            Controls.Add(bottom);
            okButton = MakeButton(bottom, okText, 0, S(10), S(150));
            cancelButton = MakeButton(bottom, "Cancel", 0, S(10), S(90));
            saveButton = MakeButton(bottom, "Save design...", 0, S(10), S(112));
            openButton = MakeButton(bottom, "Open design...", 0, S(10), S(112));
            againButton = MakeButton(bottom, "Start again", 0, S(10), S(100));
            okButton.Font = bold;
            cancelButton.DialogResult = DialogResult.Cancel;
            CancelButton = cancelButton;
            hint.AutoEllipsis = true;
            hint.ForeColor = Color.FromArgb(71, 85, 105);
            hint.TextAlign = ContentAlignment.MiddleLeft;
            bottom.Controls.Add(hint);
            tips.SetToolTip(saveButton, "Keeps this design in a file. Open it later on any table that has the same columns.");
            tips.SetToolTip(openButton, "Opens a design you saved before and puts it on this table.");
            tips.SetToolTip(againButton, "Throws the changes away and starts from the design suggested for this table.");
            bottom.Resize += delegate { PlaceButtons(); };
            okButton.Click += delegate { Do(Accept); };
            saveButton.Click += delegate { Do(SaveDesign); };
            openButton.Click += delegate { Do(OpenDesign); };
            againButton.Click += delegate { Do(StartAgain); };
            PlaceButtons();

            AllToControls();
            SetHint(editing ? "This is the design of your dashboard. Change what you like; click a card or a chart in the preview to edit it."
                            : "Pick a design on the left, then change what you like. Click a card or a chart in the preview to edit it.");
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            PlaceButtons();
            preview.Relayout();
        }

        // Closing by accident (Esc, the cross) must not throw the work away without a question.
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (e.Cancel || DialogResult == DialogResult.OK) return;
            if (e.CloseReason == CloseReason.WindowsShutDown || e.CloseReason == CloseReason.TaskManagerClosing) return;
            bool changed;
            try { changed = s.ToText() != startText; }
            catch { changed = false; }
            if (!changed) return;
            string ask = "Close Dashboard Studio without using this design?\n\nThe changes you made in this window are lost.";
            if (MessageBox.Show(this, ask, Design.AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) e.Cancel = true;
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                if (logo != null) logo.Dispose();
                if (bold != null) bold.Dispose();
                tips.Dispose();
            }
        }

        // sizes are written for a 100% screen and scaled for the screen in use
        internal int S(int px)
        {
            return (int)Math.Round(px * z);
        }

        internal int Version { get { return version; } }

        // runs the code of a button or of a change; an error is shown instead of closing the window
        void Do(Action act)
        {
            if (loading) return;
            try { act(); }
            catch (Exception ex) { Oops(ex); }
        }

        void Oops(Exception ex)
        {
            loading = false;
            MessageBox.Show(this, "That did not work:\n" + ex.Message, Design.AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        void SetHint(string text)
        {
            hint.Text = text;
        }

        void PlaceButtons()
        {
            if (okButton == null || againButton == null) return;
            int right = bottom.ClientSize.Width - S(12);
            cancelButton.Left = right - cancelButton.Width;
            okButton.Left = cancelButton.Left - S(8) - okButton.Width;
            saveButton.Left = okButton.Left - S(22) - saveButton.Width;
            openButton.Left = saveButton.Left - S(8) - openButton.Width;
            againButton.Left = openButton.Left - S(8) - againButton.Width;
            hint.SetBounds(S(12), S(6), Math.Max(S(40), againButton.Left - S(24)), S(38));
        }

        // ------------------------------------------------------------------ what the table offers
        void FillColumnLists()
        {
            foreach (DataCol c in t.Cols)
            {
                anyCols.Add(c.Name);
                if (c.IsNum) numberCols.Add(c.Name);
            }
            if (t.HasDate)
            {
                for (int i = 0; i < DatePart.All.Length; i++)
                {
                    groupFields.Add(DatePart.All[i]);
                    groupLabels.Add(DatePart.Labels[i] + "  (from the date)");
                }
            }
            foreach (DataCol c in t.Cols)
            {
                if (!c.IsCat) continue;
                groupFields.Add(c.Name);
                groupLabels.Add(c.Name + "  (" + c.Distinct.ToString("#,##0", CultureInfo.InvariantCulture) + (c.Distinct == 1 ? " value)" : " values)"));
            }
            topChoices.AddRange(new int[] { 0, -1, 3, 5, 8, 10, 12, 15, 20, 25 });
        }

        int GroupIndex(string field)
        {
            for (int i = 0; i < groupFields.Count; i++)
            {
                if (string.Equals(groupFields[i], field, StringComparison.OrdinalIgnoreCase)) return i;
            }
            return -1;
        }

        static string TopName(int n)
        {
            return n == 0 ? "Automatic" : (n < 0 ? "All of them" : "Top " + n.ToString(CultureInfo.InvariantCulture));
        }

        // ------------------------------------------------------------------ building blocks for the tabs
        void NewTab(string title)
        {
            page = new TabPage(title);
            page.AutoScroll = true;
            page.UseVisualStyleBackColor = true;
            tabs.TabPages.Add(page);
            cy = S(10);
        }

        // a little room under the last control of a tab
        void EndTab()
        {
            Label l = new Label();
            l.SetBounds(S(10), cy, S(10), S(8));
            page.Controls.Add(l);
        }

        Button MakeButton(Control parent, string text, int x, int y, int width)
        {
            Button b = new Button();
            b.Text = text;
            b.SetBounds(x, y, width, S(30));
            parent.Controls.Add(b);
            return b;
        }

        Label Caption(string text, int x, int width)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(x, cy + S(4), width, S(20));
            page.Controls.Add(l);
            return l;
        }

        // a heading line across the tab
        void Heading(string text)
        {
            cy += S(4);
            Label l = new Label();
            l.Text = text;
            l.Font = bold;
            l.SetBounds(S(10), cy, S(Edge - 10), S(20));
            page.Controls.Add(l);
            cy += S(24);
        }

        // lines of smaller grey text
        Label Note(string text, int lines)
        {
            Label l = new Label();
            l.Text = text;
            l.ForeColor = Color.FromArgb(100, 116, 139);
            l.SetBounds(S(10), cy, S(Edge - 10), S(17) * lines);
            page.Controls.Add(l);
            cy += S(17) * lines + S(6);
            return l;
        }

        ComboBox NewCombo(int x, int width, string[] items, Action changed)
        {
            ComboBox c = new ComboBox();
            c.DropDownStyle = ComboBoxStyle.DropDownList;
            c.MaxDropDownItems = 16;
            c.SetBounds(x, cy, width, S(24));
            c.DropDownWidth = Math.Max(width, S(170));
            if (items != null) foreach (string it in items) c.Items.Add(it);
            page.Controls.Add(c);
            if (changed != null) c.SelectedIndexChanged += delegate { Do(changed); };
            CalmWheel(c);
            return c;
        }

        // The mouse wheel over a closed list must not change its value by accident: it scrolls the tab instead.
        static void CalmWheel(ComboBox c)
        {
            c.MouseWheel += delegate(object sender, MouseEventArgs e)
            {
                if (c.DroppedDown) return;
                HandledMouseEventArgs h = e as HandledMouseEventArgs;
                if (h != null) h.Handled = true;
                ScrollableControl p = c.Parent as ScrollableControl;
                if (p != null && p.AutoScroll) p.AutoScrollPosition = new Point(0, Math.Max(0, -p.AutoScrollPosition.Y - e.Delta));
            };
        }

        ComboBox Combo(string caption, string[] items, Action changed)
        {
            Caption(caption, S(10), S(118));
            ComboBox c = NewCombo(S(130), S(Edge - 130), items, changed);
            cy += S(29);
            return c;
        }

        // two short lists on one line: first the left one, then the right one, which ends the line
        ComboBox Half(bool right, string caption, string[] items, Action changed)
        {
            ComboBox c;
            if (!right)
            {
                Caption(caption, S(10), S(118));
                c = NewCombo(S(130), S(126), items, changed);
            }
            else
            {
                Caption(caption, S(266), S(62));
                c = NewCombo(S(328), S(Edge - 328), items, changed);
                cy += S(29);
            }
            return c;
        }

        TextBox Input(string caption, Action changed)
        {
            Caption(caption, S(10), S(118));
            TextBox b = new TextBox();
            b.SetBounds(S(130), cy, S(Edge - 130), S(24));
            page.Controls.Add(b);
            if (changed != null) b.TextChanged += delegate { Do(changed); };
            // Enter in a line of text does nothing (and does not ring the bell)
            b.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) e.SuppressKeyPress = true;
            };
            cy += S(29);
            return b;
        }

        CheckBox Check(string text, Action changed)
        {
            CheckBox c = new CheckBox();
            c.Text = text;
            c.SetBounds(S(130), cy, S(Edge - 130), S(22));
            page.Controls.Add(c);
            if (changed != null) c.CheckedChanged += delegate { Do(changed); };
            cy += S(26);
            return c;
        }

        // a small square next to a colour list that shows the colour in use; a click asks for an own colour
        Panel Swatch(ComboBox beside, Action clicked)
        {
            Panel p = new Panel();
            p.BorderStyle = BorderStyle.FixedSingle;
            p.Cursor = Cursors.Hand;
            beside.Width = S(Edge - 130 - 32);
            p.SetBounds(beside.Left + beside.Width + S(8), beside.Top, S(24), S(23));
            page.Controls.Add(p);
            tips.SetToolTip(p, "Click to choose your own colour");
            p.Click += delegate { Do(clicked); };
            return p;
        }

        // the list with Add / Remove / Up / Down / Copy beside it, as on the Cards and Charts tabs
        ListBox ItemList(int height, Action picked, Action add, Action remove, Action up, Action down, Action copy)
        {
            ListBox list = new ListBox();
            list.IntegralHeight = false;
            list.SetBounds(S(10), cy, S(Edge - 112), height);
            page.Controls.Add(list);
            int bx = S(Edge - 94), step = (height - S(25)) / 2;
            Button b1 = MakeButton(page, "Add", bx, cy, S(45));
            Button b5 = MakeButton(page, "Copy", bx + S(49), cy, S(45));
            Button b2 = MakeButton(page, "Remove", bx, cy + step, S(94));
            Button b3 = MakeButton(page, "Up", bx, cy + 2 * step, S(45));
            Button b4 = MakeButton(page, "Down", bx + S(49), cy + 2 * step, S(45));
            b1.Height = b2.Height = b3.Height = b4.Height = b5.Height = S(25);
            tips.SetToolTip(b1, "Adds a new one at the end of the list.");
            tips.SetToolTip(b5, "Adds a copy of the selected one, to change it a little.");
            tips.SetToolTip(b3, "Moves the selected one a place up: it comes earlier on the dashboard.");
            tips.SetToolTip(b4, "Moves the selected one a place down.");
            list.SelectedIndexChanged += delegate { Do(picked); };
            b1.Click += delegate { Do(add); };
            b2.Click += delegate { Do(remove); };
            b3.Click += delegate { Do(up); };
            b4.Click += delegate { Do(down); };
            b5.Click += delegate { Do(copy); };
            cy += height + S(8);
            return list;
        }

        // sets a list without an error when the value is not among its items
        static void Pick(ComboBox c, int index)
        {
            c.SelectedIndex = index >= 0 && index < c.Items.Count ? index : (c.Items.Count > 0 ? 0 : -1);
        }

        static void Fill(ComboBox c, List<string> items)
        {
            c.Items.Clear();
            foreach (string it in items) c.Items.Add(it);
        }

        static Color Rgb(int c)
        {
            return Color.FromArgb(ColorMath.R(c), ColorMath.G(c), ColorMath.B(c));
        }

        static int FromColor(Color c)
        {
            return ColorMath.Rgb(c.R, c.G, c.B);
        }

        // asks for a colour; -1 when cancelled
        int AskColor(int start)
        {
            using (ColorDialog d = new ColorDialog())
            {
                d.FullOpen = true;
                d.Color = Rgb(start);
                return d.ShowDialog(this) == DialogResult.OK ? FromColor(d.Color) : -1;
            }
        }

        static int ColorIndex(int code)
        {
            return code >= CardSpec.CustomColor ? 7 : (code >= 0 && code < 6 ? code + 1 : 0);
        }

        // ------------------------------------------------------------------ tab 1: ready designs
        void BuildDesignsTab()
        {
            NewTab("Designs");
            page.AutoScroll = false;
            Panel top = new Panel();
            top.Dock = DockStyle.Top;
            top.Height = S(62);
            Label note = new Label();
            note.Text = gallery.Count.ToString(CultureInfo.InvariantCulture) + " ready designs, drawn from your own table. Click one to use it. The other tabs change every part of it.";
            note.ForeColor = Color.FromArgb(100, 116, 139);
            note.SetBounds(S(6), S(6), S(318), S(52));
            top.Controls.Add(note);
            Button lucky = MakeButton(top, "Surprise me", S(334), S(14), S(102));
            tips.SetToolTip(lucky, "Picks one of the designs by chance.");
            lucky.Click += delegate { Do(gallery.PickByChance); };
            gallery.Dock = DockStyle.Fill;
            page.Controls.Add(gallery);
            page.Controls.Add(top);
        }

        internal void UseIdea(DesignIdea idea)
        {
            if (loading) return;
            try
            {
                DashAuto.Apply(s, idea);
                AllToControls();
                SetHint("Design: " + idea.Name + ".  Every chart has its automatic size again.");
            }
            catch (Exception ex) { Oops(ex); }
        }

        // true when the design in the window has the look of this ready design
        internal bool Matches(DesignIdea d)
        {
            return s.CardStyle == d.CardStyle && s.Arrange == d.Arrange && s.Rounded == d.Rounded && s.SlicerPos == d.SlicerPos && DashThemes.IndexOf(s.Theme.Name) == d.Theme;
        }

        // a small picture of the dashboard in one of the ready designs
        internal Bitmap Thumbnail(DesignIdea idea, int width, int height)
        {
            DashSpec d = s.Clone();
            DashAuto.Apply(d, idea);
            Bitmap bmp = new Bitmap(Math.Max(8, width), Math.Max(8, height));
            using (Graphics g = Graphics.FromImage(bmp))
            {
                DashRender.Paint(g, new RectangleF(0, 0, bmp.Width, bmp.Height), t, d, cache, -1, -1, null);
            }
            return bmp;
        }

        // ------------------------------------------------------------------ tab 2: KPI cards
        void BuildCardsTab()
        {
            NewTab("Cards");
            Note("The big numbers on top. Up to " + DashSpec.MaxCards + " cards. Each one can compare months, draw a small trend line or show how much of a target is reached.", 2);
            cardList = ItemList(S(100), CardPicked, CardAdd, CardRemove, delegate { CardMove(-1); }, delegate { CardMove(1); }, CardCopy);
            Heading("The selected card");
            cardCol = Combo("Number", null, CardColumnChanged);
            cardAgg = Combo("Show", null, CardEdited);
            cardLabel = Input("Label", CardEdited);
            tips.SetToolTip(cardLabel, "Leave empty for an automatic label such as Total Amount.");
            cardFmt = Half(false, "Format", new string[] { "Automatic", "Short with units", "Full number", "Percent" }, CardEdited);
            cardDec = Half(true, "Decimals", new string[] { "Automatic", "0", "1", "2", "3", "4" }, CardEdited);
            cardPrefix = Input("Text in front", CardEdited);
            cardSuffix = Input("Text behind", CardEdited);
            cardColor = Combo("Colour", ColorNames, CardColorChanged);
            cardSwatch = Swatch(cardColor, CardOwnColor);
            cardCompare = Combo("Compare", new string[] { "No comparison", "Latest month with the month before", "Latest month with the same month last year" }, CardEdited);
            cardSpark = Check("Small trend line of the months", CardEdited);
            cardTarget = Input("Target", CardEdited);
            tips.SetToolTip(cardTarget, "A number. The card then shows how much of the target is reached. Leave empty for no target.");
            if (!t.HasDate) Note("This table has no date column, so months cannot be compared and no trend line can be drawn.", 2);
            EndTab();
        }

        void CardsToList()
        {
            bool was = loading;
            loading = true;
            int keep = cardList.SelectedIndex;
            cardList.Items.Clear();
            foreach (CardSpec c in s.Cards) cardList.Items.Add(DashText.CardLabel(c));
            if (s.Cards.Count > 0) cardList.SelectedIndex = keep >= 0 && keep < s.Cards.Count ? keep : 0;
            loading = was;
            CardToControls();
        }

        CardSpec CurrentCard()
        {
            int i = cardList.SelectedIndex;
            return i >= 0 && i < s.Cards.Count ? s.Cards[i] : null;
        }

        // the ways a column can be shown: a text column can only be counted
        List<int> AggChoices(string col, bool forChart)
        {
            List<int> list = new List<int>();
            DataCol c = t.Find(col);
            if (c == null)
            {
                list.Add(AggKind.Count);
                return list;
            }
            if (c.IsNum)
            {
                list.Add(AggKind.Sum);
                list.Add(AggKind.Avg);
                list.Add(AggKind.Max);
                list.Add(AggKind.Min);
            }
            list.Add(AggKind.Count);
            if (!forChart && (c.IsCat || c.IsNum)) list.Add(AggKind.Distinct);
            return list;
        }

        void FillAgg(ComboBox box, string col, int agg, bool forChart)
        {
            List<int> choices = AggChoices(col, forChart);
            bool rows = t.Find(col) == null;
            box.Items.Clear();
            int at = 0;
            for (int i = 0; i < choices.Count; i++)
            {
                box.Items.Add(rows ? "Count the rows" : AggKind.Names[choices[i]]);
                if (choices[i] == agg) at = i;
            }
            box.Tag = choices;
            Pick(box, at);
        }

        static int AggOf(ComboBox box)
        {
            List<int> choices = box.Tag as List<int>;
            int i = box.SelectedIndex;
            return choices != null && i >= 0 && i < choices.Count ? choices[i] : AggKind.Count;
        }

        void FillColumns(ComboBox box, string current)
        {
            List<string> items = new List<string>();
            items.Add(RowsChoice);
            foreach (string c in anyCols) items.Add(c);
            Fill(box, items);
            int at = 0;
            for (int i = 1; i < items.Count; i++)
            {
                if (string.Equals(items[i], current, StringComparison.OrdinalIgnoreCase)) at = i;
            }
            Pick(box, at);
        }

        static string ColumnOf(ComboBox box)
        {
            return box.SelectedIndex <= 0 ? "" : (string)box.Items[box.SelectedIndex];
        }

        // The way to show a column that was just picked. What was chosen before stays when it was a choice;
        // a column that could only be counted before starts with its total (or its average, when it is a rate).
        int FirstAgg(string before, string col, int now, bool forChart)
        {
            List<int> ok = AggChoices(col, forChart);
            DataCol c = t.Find(col), was = t.Find(before);
            if (c == null || !c.IsNum) return ok.Contains(now) ? now : ok[0];
            int natural = Logic.MeasureIsAvg(col) ? AggKind.Avg : AggKind.Sum;
            if (was == null || !was.IsNum) return natural;
            if (now == AggKind.Sum || (now == AggKind.Avg && Logic.MeasureIsAvg(before))) return natural;
            return ok.Contains(now) ? now : natural;
        }

        void CardToControls()
        {
            bool was = loading;
            loading = true;
            CardSpec c = CurrentCard();
            bool on = c != null;
            foreach (Control x in new Control[] { cardCol, cardAgg, cardLabel, cardFmt, cardDec, cardPrefix, cardSuffix, cardColor, cardCompare, cardSpark, cardTarget, cardSwatch }) x.Enabled = on;
            if (on)
            {
                FillColumns(cardCol, c.Col);
                FillAgg(cardAgg, c.Col, c.Agg, false);
                cardLabel.Text = c.Label;
                Pick(cardFmt, c.Fmt);
                Pick(cardDec, c.Dec + 1);
                cardPrefix.Text = c.Prefix;
                cardSuffix.Text = c.Suffix;
                cardTarget.Text = c.Target > 0 ? c.Target.ToString("0.############", CultureInfo.InvariantCulture) : "";
                CardSmallParts(c);
            }
            else cardSwatch.BackColor = SystemColors.Control;
            loading = was;
        }

        // the parts of the card editor that follow from other choices
        void CardSmallParts(CardSpec c)
        {
            Pick(cardColor, ColorIndex(c.Color));
            cardSwatch.BackColor = Rgb(ColorMath.Pick(s.Theme, c.Color, cardList.SelectedIndex));
            Pick(cardCompare, c.Compare);
            cardSpark.Checked = c.Spark;
            DataCol col = t.Find(c.Col);
            bool months = t.HasDate && c.Agg != AggKind.Distinct && (c.Col.Length == 0 || c.Agg == AggKind.Count || (col != null && col.IsNum));
            cardCompare.Enabled = months;
            cardSpark.Enabled = months;
        }

        void CardColumnChanged()
        {
            CardSpec c = CurrentCard();
            if (c == null) return;
            string before = c.Col;
            c.Col = ColumnOf(cardCol);
            c.Agg = FirstAgg(before, c.Col, c.Agg, false);
            loading = true;
            FillAgg(cardAgg, c.Col, c.Agg, false);
            loading = false;
            CardEdited();
        }

        void CardColorChanged()
        {
            CardSpec c = CurrentCard();
            if (c == null) return;
            int i = cardColor.SelectedIndex;
            if (i == 7)
            {
                CardOwnColor();
                return;
            }
            c.Color = i - 1;
            CardEdited();
        }

        void CardOwnColor()
        {
            CardSpec c = CurrentCard();
            if (c == null) return;
            int picked = AskColor(ColorMath.Pick(s.Theme, c.Color, cardList.SelectedIndex));
            if (picked >= 0) c.Color = CardSpec.CustomColor + picked;
            CardEdited();
        }

        void CardEdited()
        {
            CardSpec c = CurrentCard();
            if (c == null) return;
            int i = cardList.SelectedIndex;
            c.Col = ColumnOf(cardCol);
            c.Agg = c.Col.Length == 0 ? AggKind.Count : AggOf(cardAgg);
            c.Label = cardLabel.Text.Trim();
            c.Fmt = Math.Max(0, cardFmt.SelectedIndex);
            c.Dec = Math.Max(0, cardDec.SelectedIndex) - 1;
            c.Prefix = cardPrefix.Text;
            c.Suffix = cardSuffix.Text;
            if (cardCompare.Enabled) c.Compare = Math.Max(0, cardCompare.SelectedIndex);
            if (cardSpark.Enabled) c.Spark = cardSpark.Checked;
            double target;
            string tt = cardTarget.Text.Trim().Replace(",", "").Replace(" ", "");
            c.Target = double.TryParse(tt, NumberStyles.Float, CultureInfo.InvariantCulture, out target) && target > 0 && !double.IsInfinity(target) ? target : 0;
            DashAuto.Fix(s, t);
            loading = true;
            if (i >= 0 && i < cardList.Items.Count) cardList.Items[i] = DashText.CardLabel(c);
            CardSmallParts(c);
            loading = false;
            SelectBox(Box.Card, i);
            Changed();
        }

        void CardPicked()
        {
            CardToControls();
            SelectBox(Box.Card, cardList.SelectedIndex);
        }

        void CardAdd()
        {
            if (s.Cards.Count >= DashSpec.MaxCards)
            {
                SetHint("A dashboard takes " + DashSpec.MaxCards + " cards at most.");
                return;
            }
            CardSpec c = new CardSpec();
            // the first number column that has no card yet
            foreach (string col in numberCols)
            {
                bool used = false;
                foreach (CardSpec o in s.Cards) if (string.Equals(o.Col, col, StringComparison.OrdinalIgnoreCase)) used = true;
                if (used) continue;
                c.Col = col;
                c.Agg = Logic.MeasureIsAvg(col) ? AggKind.Avg : AggKind.Sum;
                break;
            }
            if (c.Col.Length == 0) c.Agg = AggKind.Count;
            s.Cards.Add(c);
            CardsToList();
            cardList.SelectedIndex = s.Cards.Count - 1;
            Changed();
        }

        void CardCopy()
        {
            CardSpec c = CurrentCard();
            if (c == null) return;
            if (s.Cards.Count >= DashSpec.MaxCards)
            {
                SetHint("A dashboard takes " + DashSpec.MaxCards + " cards at most.");
                return;
            }
            int at = cardList.SelectedIndex + 1;
            s.Cards.Insert(at, c.Clone());
            CardsToList();
            cardList.SelectedIndex = at;
            Changed();
        }

        void CardRemove()
        {
            int i = cardList.SelectedIndex;
            if (i < 0 || i >= s.Cards.Count) return;
            s.Cards.RemoveAt(i);
            CardsToList();
            SelectBox(Box.Card, cardList.SelectedIndex);
            Changed();
        }

        void CardMove(int by)
        {
            int i = cardList.SelectedIndex, k = i + by;
            if (i < 0 || k < 0 || k >= s.Cards.Count) return;
            CardSpec c = s.Cards[i];
            s.Cards.RemoveAt(i);
            s.Cards.Insert(k, c);
            CardsToList();
            cardList.SelectedIndex = k;
            Changed();
        }

        // ------------------------------------------------------------------ tab 3: charts
        void BuildChartsTab()
        {
            NewTab("Charts");
            chartList = ItemList(S(84), ChartPicked, ChartAdd, ChartRemove, delegate { ChartMove(-1); }, delegate { ChartMove(1); }, ChartCopy);
            chKind = Combo("Kind of chart", ChartKind.Names, ChartKindChanged);
            chDim = Combo("Group by", groupLabels.ToArray(), ChartEdited);
            chCol = Combo("Number", null, ChartColumnChanged);
            chAgg = Combo("Show", null, ChartEdited);
            chSplit = Combo("Colours by", null, ChartEdited);
            tips.SetToolTip(chSplit, "For stacked charts: one colour for each value of this column.");
            chCol2 = Combo("Second number", null, ChartColumn2Changed);
            tips.SetToolTip(chCol2, "Combo chart: the line. Table: a second column of numbers.");
            chAgg2 = Combo("Show second", null, ChartEdited);
            chTop = Half(false, "How many", null, ChartEdited);
            chOthers = Half(true, "The rest", new string[] { "Automatic", "As Others", "Leave out" }, ChartEdited);
            chSort = Half(false, "Order", new string[] { "Automatic", "Largest first", "Smallest first", "A to Z" }, ChartEdited);
            chShowAs = Half(true, "Values as", new string[] { "Numbers", "% of total", "Running total" }, ChartEdited);
            chLabels = Half(false, "Data labels", new string[] { "Automatic", "Off", "Values", "Percent" }, ChartEdited);
            chLegend = Half(true, "Legend", new string[] { "Automatic", "Off", "Right", "Bottom", "Top" }, ChartEdited);
            chSpan = Half(false, "Width", DashLayout.SpanNames, ChartEdited);
            chHeight = Half(true, "Height", DashLayout.HeightNames, ChartEdited);
            chTitle = Input("Title", ChartEdited);
            tips.SetToolTip(chTitle, "Leave empty for an automatic title such as Amount by Customer.");
            chColor = Combo("Colour", ColorNames, ChartColorChanged);
            chSwatch = Swatch(chColor, ChartOwnColor);
            chTrend = Half(false, "Trend line", new string[] { "None", "Straight line", "Average of 3" }, ChartEdited);
            chSmooth = new CheckBox();
            chSmooth.Text = "Smooth line";
            chSmooth.SetBounds(S(270), cy + S(1), S(Edge - 270), S(22));
            page.Controls.Add(chSmooth);
            chSmooth.CheckedChanged += delegate { Do(ChartEdited); };
            cy += S(29);
            chNote = Note("", 1);
            EndTab();
        }

        static string ChartText(ChartSpec c)
        {
            return ChartKind.Names[c.Kind < 0 || c.Kind >= ChartKind.Names.Length ? 0 : c.Kind] + ":  " + DashText.ChartTitle(c, null);
        }

        void ChartsToList()
        {
            bool was = loading;
            loading = true;
            int keep = chartList.SelectedIndex;
            chartList.Items.Clear();
            foreach (ChartSpec c in s.Charts) chartList.Items.Add(ChartText(c));
            if (s.Charts.Count > 0) chartList.SelectedIndex = keep >= 0 && keep < s.Charts.Count ? keep : 0;
            loading = was;
            ChartToControls();
        }

        ChartSpec CurrentChart()
        {
            int i = chartList.SelectedIndex;
            return i >= 0 && i < s.Charts.Count ? s.Charts[i] : null;
        }

        // the columns a chart can take its colours from: every column to group by, but not the one it is grouped by
        void FillSplit(string current, string dim)
        {
            List<string> items = new List<string>(), fields = new List<string>();
            items.Add(NoneChoice);
            fields.Add("");
            int at = 0;
            for (int i = 0; i < groupFields.Count; i++)
            {
                if (string.Equals(groupFields[i], dim, StringComparison.OrdinalIgnoreCase)) continue;
                items.Add(groupLabels[i]);
                fields.Add(groupFields[i]);
                if (string.Equals(groupFields[i], current, StringComparison.OrdinalIgnoreCase)) at = items.Count - 1;
            }
            Fill(chSplit, items);
            chSplit.Tag = fields;
            Pick(chSplit, at);
        }

        string SplitNow()
        {
            List<string> fields = chSplit.Tag as List<string>;
            int i = chSplit.SelectedIndex;
            return fields != null && i >= 0 && i < fields.Count ? fields[i] : "";
        }

        void FillTop(int current)
        {
            if (!topChoices.Contains(current)) topChoices.Add(current);
            chTop.Items.Clear();
            foreach (int n in topChoices) chTop.Items.Add(TopName(n));
            Pick(chTop, topChoices.IndexOf(current));
        }

        // only the choices that mean something for the kind of chart can be used
        void ChartEnable(ChartSpec c)
        {
            bool on = c != null;
            foreach (Control x in new Control[] { chKind, chDim, chCol, chAgg, chSplit, chCol2, chAgg2, chTop, chOthers, chSort, chShowAs, chTitle, chLabels, chLegend, chSpan, chHeight, chColor, chTrend, chSmooth, chSwatch })
                x.Enabled = on;
            if (!on)
            {
                chNote.Text = s.Charts.Count == 0 ? "No chart yet: click Add. A dashboard takes up to " + DashSpec.MaxCharts + " charts." : "";
                chSwatch.BackColor = SystemColors.Control;
                return;
            }
            bool stacked = ChartKind.IsStacked(c.Kind), round = ChartKind.IsRound(c.Kind), line = ChartKind.IsLine(c.Kind);
            bool second = ChartKind.HasSecond(c.Kind);
            bool running = c.ShowAs == 2 && DashEngine.CanShowAs(c.Kind);
            bool time = DatePart.Is(c.Dim);
            chCol2.Enabled = second;
            chAgg2.Enabled = second && c.Col2.Length > 0;
            chShowAs.Enabled = DashEngine.CanShowAs(c.Kind);
            chSort.Enabled = c.Kind != ChartKind.Pareto && !time;
            chTop.Enabled = !running;
            chOthers.Enabled = c.Kind != ChartKind.Pareto && !running;
            chColor.Enabled = !stacked && !round;
            chSwatch.Enabled = !stacked && !round;
            chSplit.Enabled = stacked || c.Kind == ChartKind.Column || c.Kind == ChartKind.Bar;
            chTrend.Enabled = line;
            chSmooth.Enabled = c.Kind == ChartKind.Line;
            chLegend.Enabled = c.Kind != ChartKind.Table;
            chLabels.Enabled = c.Kind != ChartKind.Table && !stacked;
            string note = "";
            if (c.Kind == ChartKind.Table) note = "A table is drawn in cells. Every row follows the filters.";
            else if (c.Kind == ChartKind.Pareto) note = "Largest first, with a line that adds the shares up to 100%.";
            else if (running) note = "A running total always shows every group.";
            else if (time) note = "Groups of the date are always shown in the order of time.";
            else if (c.Kind == ChartKind.Combo && c.Col2.Length == 0) note = "Pick a second number: it is drawn as a line with its own scale.";
            chNote.Text = note;
        }

        void ChartToControls()
        {
            bool was = loading;
            loading = true;
            ChartSpec c = CurrentChart();
            if (c != null)
            {
                Pick(chKind, c.Kind);
                Pick(chDim, GroupIndex(c.Dim));
                FillColumns(chCol, c.Col);
                FillAgg(chAgg, c.Col, c.Agg, true);
                FillSplit(c.Split, c.Dim);
                FillColumns(chCol2, c.Col2);
                FillAgg(chAgg2, c.Col2, c.Agg2, true);
                FillTop(c.Top);
                Pick(chOthers, c.Others);
                Pick(chSort, c.Sort);
                Pick(chShowAs, c.ShowAs);
                chTitle.Text = c.Title;
                Pick(chLabels, c.Labels);
                Pick(chLegend, c.Legend);
                Pick(chSpan, Math.Max(0, Array.IndexOf(DashLayout.SpanValues, c.Span)));
                Pick(chHeight, c.Height);
                Pick(chColor, ColorIndex(c.Color));
                chSwatch.BackColor = Rgb(DashLook.ChartColor(s, c, chartList.SelectedIndex));
                Pick(chTrend, c.Trend);
                chSmooth.Checked = c.Smooth;
            }
            ChartEnable(c);
            loading = was;
        }

        // A stacked chart needs a second column to take its colours from: the column with the fewest values is offered.
        string SplitFor(ChartSpec c)
        {
            string best = "";
            int fewest = int.MaxValue;
            foreach (DataCol d in t.Cols)
            {
                if (!d.IsCat || d.Distinct < 2 || string.Equals(d.Name, c.Dim, StringComparison.OrdinalIgnoreCase)) continue;
                if (d.Distinct < fewest)
                {
                    fewest = d.Distinct;
                    best = d.Name;
                }
            }
            if (best.Length > 0 && fewest <= 12) return best;
            if (t.HasDate && !DatePart.Is(c.Dim)) return DatePart.Year;
            return best;
        }

        void ChartKindChanged()
        {
            ChartSpec c = CurrentChart();
            if (c == null) return;
            int kind = Math.Max(0, chKind.SelectedIndex);
            if (ChartKind.IsStacked(kind) && c.Split.Length == 0)
            {
                c.Split = SplitFor(c);
                if (c.Split.Length == 0)
                {
                    SetHint("A stacked chart needs a second column to take its colours from. This table has none besides '" + DashText.DimName(c.Dim) + "'.");
                    loading = true;
                    Pick(chKind, c.Kind);
                    loading = false;
                    return;
                }
            }
            if (!ChartKind.IsStacked(kind)) c.Split = "";
            if (kind == ChartKind.Pareto && c.Agg != AggKind.Sum && c.Agg != AggKind.Count)
            {
                DataCol col = t.Find(c.Col);
                c.Agg = col != null && col.IsNum ? AggKind.Sum : AggKind.Count;
            }
            c.Kind = kind;
            loading = true;
            FillAgg(chAgg, c.Col, c.Agg, true);
            FillSplit(c.Split, c.Dim);
            loading = false;
            ChartEdited();
        }

        void ChartColumnChanged()
        {
            ChartSpec c = CurrentChart();
            if (c == null) return;
            string before = c.Col;
            c.Col = ColumnOf(chCol);
            c.Agg = FirstAgg(before, c.Col, c.Agg, true);
            loading = true;
            FillAgg(chAgg, c.Col, c.Agg, true);
            loading = false;
            ChartEdited();
        }

        void ChartColumn2Changed()
        {
            ChartSpec c = CurrentChart();
            if (c == null) return;
            string before = c.Col2;
            c.Col2 = ColumnOf(chCol2);
            c.Agg2 = FirstAgg(before, c.Col2, c.Agg2, true);
            loading = true;
            FillAgg(chAgg2, c.Col2, c.Agg2, true);
            loading = false;
            ChartEdited();
        }

        void ChartColorChanged()
        {
            ChartSpec c = CurrentChart();
            if (c == null) return;
            int i = chColor.SelectedIndex;
            if (i == 7)
            {
                ChartOwnColor();
                return;
            }
            c.Color = i - 1;
            ChartEdited();
        }

        void ChartOwnColor()
        {
            ChartSpec c = CurrentChart();
            if (c == null) return;
            int picked = AskColor(DashLook.ChartColor(s, c, chartList.SelectedIndex));
            if (picked >= 0) c.Color = CardSpec.CustomColor + picked;
            ChartEdited();
        }

        void ChartEdited()
        {
            ChartSpec c = CurrentChart();
            if (c == null) return;
            int i = chartList.SelectedIndex;
            c.Kind = Math.Max(0, chKind.SelectedIndex);
            int d = chDim.SelectedIndex;
            bool newDim = false;
            if (d >= 0 && d < groupFields.Count && !string.Equals(groupFields[d], c.Dim, StringComparison.OrdinalIgnoreCase))
            {
                c.Dim = groupFields[d];
                newDim = true;
            }
            c.Col = ColumnOf(chCol);
            c.Agg = c.Col.Length == 0 ? AggKind.Count : AggOf(chAgg);
            c.Split = SplitNow();
            if (string.Equals(c.Split, c.Dim, StringComparison.OrdinalIgnoreCase)) c.Split = "";
            if (c.Split.Length > 0 && !ChartKind.IsStacked(c.Kind))
            {
                // a column to take the colours from makes a plain column or bar chart a stacked one
                if (c.Kind == ChartKind.Column) c.Kind = ChartKind.StackedColumn;
                else if (c.Kind == ChartKind.Bar) c.Kind = ChartKind.StackedBar;
                else c.Split = "";
            }
            c.Col2 = ColumnOf(chCol2);
            c.Agg2 = c.Col2.Length == 0 ? AggKind.Count : AggOf(chAgg2);
            int top = chTop.SelectedIndex;
            c.Top = top >= 0 && top < topChoices.Count ? topChoices[top] : 0;
            c.Others = Math.Max(0, chOthers.SelectedIndex);
            c.Sort = Math.Max(0, chSort.SelectedIndex);
            c.ShowAs = Math.Max(0, chShowAs.SelectedIndex);
            c.Title = chTitle.Text.Trim();
            c.Labels = Math.Max(0, chLabels.SelectedIndex);
            c.Legend = Math.Max(0, chLegend.SelectedIndex);
            int span = chSpan.SelectedIndex;
            c.Span = span >= 0 && span < DashLayout.SpanValues.Length ? DashLayout.SpanValues[span] : 0;
            c.Height = Math.Max(0, chHeight.SelectedIndex);
            c.Trend = Math.Max(0, chTrend.SelectedIndex);
            c.Smooth = chSmooth.Checked;
            DashAuto.Fix(s, t);
            if (i < 0 || i >= s.Charts.Count || s.Charts[i] != c)
            {
                // the chart could not be kept as it was
                ChartsToList();
                Changed();
                return;
            }
            loading = true;
            if (i < chartList.Items.Count) chartList.Items[i] = ChartText(c);
            Pick(chKind, c.Kind);
            if (newDim) FillSplit(c.Split, c.Dim);
            Pick(chColor, ColorIndex(c.Color));
            chSwatch.BackColor = Rgb(DashLook.ChartColor(s, c, i));
            ChartEnable(c);
            loading = false;
            SelectBox(Box.Chart, i);
            Changed();
        }

        void ChartPicked()
        {
            ChartToControls();
            SelectBox(Box.Chart, chartList.SelectedIndex);
        }

        void ChartAdd()
        {
            if (s.Charts.Count >= DashSpec.MaxCharts)
            {
                SetHint("A dashboard takes " + DashSpec.MaxCharts + " charts at most.");
                return;
            }
            if (groupFields.Count == 0)
            {
                SetHint("This table has no column to group a chart by.");
                return;
            }
            ChartSpec c = new ChartSpec();
            // the first column that has no chart yet
            c.Dim = groupFields[0];
            foreach (string f in groupFields)
            {
                bool used = false;
                foreach (ChartSpec o in s.Charts) if (string.Equals(o.Dim, f, StringComparison.OrdinalIgnoreCase)) used = true;
                if (used) continue;
                c.Dim = f;
                break;
            }
            if (s.Charts.Count > 0)
            {
                c.Col = s.Charts[0].Col;
                c.Agg = s.Charts[0].Agg;
            }
            else if (numberCols.Count > 0)
            {
                c.Col = numberCols[0];
                c.Agg = Logic.MeasureIsAvg(c.Col) ? AggKind.Avg : AggKind.Sum;
            }
            else c.Agg = AggKind.Count;
            c.Kind = DatePart.Is(c.Dim) ? ChartKind.Line : ChartKind.Column;
            s.Charts.Add(c);
            ChartsToList();
            chartList.SelectedIndex = s.Charts.Count - 1;
            Changed();
        }

        void ChartCopy()
        {
            ChartSpec c = CurrentChart();
            if (c == null) return;
            if (s.Charts.Count >= DashSpec.MaxCharts)
            {
                SetHint("A dashboard takes " + DashSpec.MaxCharts + " charts at most.");
                return;
            }
            int at = chartList.SelectedIndex + 1;
            s.Charts.Insert(at, c.Clone());
            ChartsToList();
            chartList.SelectedIndex = at;
            Changed();
        }

        void ChartRemove()
        {
            int i = chartList.SelectedIndex;
            if (i < 0 || i >= s.Charts.Count) return;
            s.Charts.RemoveAt(i);
            ChartsToList();
            SelectBox(Box.Chart, chartList.SelectedIndex);
            Changed();
        }

        void ChartMove(int by)
        {
            int i = chartList.SelectedIndex, k = i + by;
            if (i < 0 || k < 0 || k >= s.Charts.Count) return;
            ChartSpec c = s.Charts[i];
            s.Charts.RemoveAt(i);
            s.Charts.Insert(k, c);
            ChartsToList();
            chartList.SelectedIndex = k;
            Changed();
        }

        // ------------------------------------------------------------------ tab 4: filters
        void BuildFiltersTab()
        {
            NewTab("Filters");
            Note("Filter buttons (slicers). A click on a button changes every card and every chart. Tick up to " + DashSpec.MaxSlicers + ".", 2);
            filterList = new CheckedListBox();
            filterList.CheckOnClick = true;
            filterList.IntegralHeight = false;
            filterList.SetBounds(S(10), cy, S(Edge - 10), S(220));
            foreach (string l in groupLabels) filterList.Items.Add(l, false);
            page.Controls.Add(filterList);
            cy += S(228);
            filterList.ItemCheck += OnFilterCheck;
            filterPos = Combo("Where", DashLayout.SlicerPosNames, FiltersEdited);
            Heading("Dates");
            List<string> dates = new List<string>();
            foreach (DataCol c in t.Cols) if (c.IsDate) dates.Add(c.Name);
            dateCol = Combo("Date column", dates.ToArray(), FiltersEdited);
            dateCol.Enabled = dates.Count > 1;
            fyStart = Combo("Fiscal year starts", new string[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" }, FiltersEdited);
            fyStart.Enabled = t.HasDate;
            Note(t.HasDate ? "Month, quarter and year are taken from the date column. Rows without a date are left out of charts over time."
                           : "This table has no date column, so there are no months, quarters or years to group by.", 2);
            EndTab();
        }

        void FiltersToControls()
        {
            bool was = loading;
            loading = true;
            for (int i = 0; i < groupFields.Count && i < filterList.Items.Count; i++)
            {
                bool on = false;
                foreach (string f in s.Slicers) if (string.Equals(f, groupFields[i], StringComparison.OrdinalIgnoreCase)) on = true;
                filterList.SetItemChecked(i, on);
            }
            Pick(filterPos, s.SlicerPos);
            if (dateCol.Items.Count > 0)
            {
                int at = 0;
                for (int i = 0; i < dateCol.Items.Count; i++) if (string.Equals((string)dateCol.Items[i], t.DateName, StringComparison.OrdinalIgnoreCase)) at = i;
                Pick(dateCol, at);
            }
            Pick(fyStart, s.FyStart - 1);
            loading = was;
        }

        void OnFilterCheck(object sender, ItemCheckEventArgs e)
        {
            if (loading) return;
            try
            {
                if (e.Index < 0 || e.Index >= groupFields.Count) return;
                string field = groupFields[e.Index];
                for (int i = s.Slicers.Count - 1; i >= 0; i--)
                {
                    if (string.Equals(s.Slicers[i], field, StringComparison.OrdinalIgnoreCase)) s.Slicers.RemoveAt(i);
                }
                if (e.NewValue == CheckState.Checked)
                {
                    if (s.Slicers.Count >= DashSpec.MaxSlicers)
                    {
                        e.NewValue = CheckState.Unchecked;
                        SetHint("A dashboard takes " + DashSpec.MaxSlicers + " filters at most.");
                    }
                    else s.Slicers.Add(field);
                }
                if (s.SlicerPos == 3 && s.Slicers.Count > 0)
                {
                    // filters were switched off: ticking one switches them on again
                    s.SlicerPos = 0;
                    loading = true;
                    Pick(filterPos, 0);
                    loading = false;
                }
                Changed();
            }
            catch (Exception ex) { Oops(ex); }
        }

        void FiltersEdited()
        {
            s.SlicerPos = Math.Max(0, filterPos.SelectedIndex);
            if (fyStart.SelectedIndex >= 0) s.FyStart = fyStart.SelectedIndex + 1;
            if (dateCol.SelectedIndex >= 0)
            {
                string name = (string)dateCol.Items[dateCol.SelectedIndex];
                if (!string.Equals(name, t.DateName, StringComparison.OrdinalIgnoreCase))
                {
                    s.DateCol = name;
                    t.UseDate(name);
                }
            }
            Changed();
        }

        // ------------------------------------------------------------------ tab 5: the look
        void BuildLookTab()
        {
            NewTab("Look");
            lookTheme = Combo("Theme", DashThemes.Names(), ThemeChanged);
            lookTheme.DrawMode = DrawMode.OwnerDrawFixed;
            lookTheme.ItemHeight = S(18);
            lookTheme.DrawItem += DrawThemeItem;
            cy += S(4);
            Caption("Colours", S(10), S(118));
            for (int i = 0; i < swatches.Length; i++)
            {
                Button b = new Button();
                b.FlatStyle = FlatStyle.Flat;
                b.Cursor = Cursors.Hand;
                b.SetBounds(S(130) + (i % 6) * S(36), cy + (i / 6) * S(36), S(30), S(30));
                page.Controls.Add(b);
                tips.SetToolTip(b, SwatchNames[i] + ": click to change");
                int which = i;
                b.Click += delegate { Do(delegate { SwatchClicked(which); }); };
                swatches[i] = b;
            }
            cy += S(76);
            Note("First row: the six colours of cards and charts. Second row: page, title band, cards, text. Click one to change it.", 2);
            lookCards = Combo("KPI cards", DashThemes.CardStyles, LookEdited);
            lookRounded = Check("Rounded corners", LookEdited);
            lookShadow = Check("Soft shadows", LookEdited);
            lookFont = Combo("Font", DashThemes.Fonts, LookEdited);
            lookDensity = Combo("Spacing", DashLayout.DensityNames, LookEdited);
            Heading("Numbers");
            lookUnit = Combo("Units", new string[] { "Indian: lakh and crore (12.5 L, 3.2 Cr)", "International: K, M, B", "Full numbers" }, LookEdited);
            lookCurrency = Input("Currency sign", LookEdited);
            tips.SetToolTip(lookCurrency, "For example Rs or a rupee sign. It is put in front of amounts on the cards.");
            EndTab();
        }

        // a theme in the list: its title band, its six colours and its name
        void DrawThemeItem(object sender, DrawItemEventArgs e)
        {
            try
            {
                e.DrawBackground();
                ThemeSpec th = e.Index >= 0 && e.Index < DashThemes.Count ? DashThemes.Get(e.Index) : s.Theme;
                string name = e.Index >= 0 && e.Index < lookTheme.Items.Count ? (string)lookTheme.Items[e.Index] : th.Name;
                int x = e.Bounds.X + S(3), y = e.Bounds.Y + S(3), h = Math.Max(4, e.Bounds.Height - S(6)), w = S(9);
                using (SolidBrush b = new SolidBrush(Rgb(th.Band)))
                {
                    e.Graphics.FillRectangle(b, x, y, w + S(4), h);
                }
                for (int i = 0; i < 6; i++)
                {
                    using (SolidBrush b = new SolidBrush(Rgb(th.Accent[i])))
                    {
                        e.Graphics.FillRectangle(b, x + w + S(4) + i * w, y, w, h);
                    }
                }
                using (SolidBrush b = new SolidBrush(Rgb(th.Page)))
                {
                    e.Graphics.FillRectangle(b, x + S(4) + 7 * w, y, w, h);
                }
                using (Pen p = new Pen(Color.FromArgb(148, 163, 184)))
                {
                    e.Graphics.DrawRectangle(p, x, y, S(4) + 8 * w, h);
                }
                using (SolidBrush ink = new SolidBrush(e.ForeColor))
                {
                    e.Graphics.DrawString(name, e.Font, ink, x + S(12) + 8 * w, e.Bounds.Y + (e.Bounds.Height - e.Font.Height) / 2f);
                }
                e.DrawFocusRectangle();
            }
            catch { }
        }

        int SwatchColor(int i)
        {
            ThemeSpec th = s.Theme;
            if (i < 6) return th.Accent[i];
            return i == 6 ? th.Page : (i == 7 ? th.Band : (i == 8 ? th.Card : th.Text));
        }

        void SwatchClicked(int i)
        {
            int picked = AskColor(SwatchColor(i));
            if (picked < 0) return;
            ThemeSpec th = s.Theme;
            if (i < 6) th.Accent[i] = picked;
            else if (i == 6)
            {
                th.Page = picked;
                th.Dark = ColorMath.Lum(picked) < 0.4;
            }
            else if (i == 7)
            {
                th.Band = picked;
                th.BandText = ColorMath.Ink(picked);
            }
            else if (i == 8)
            {
                th.Card = picked;
                // the text and the grid lines must stay readable on the new card colour
                if (Math.Abs(ColorMath.Lum(th.Text) - ColorMath.Lum(picked)) < 0.45) th.Text = ColorMath.Ink(picked);
                th.Muted = ColorMath.Mix(th.Text, picked, 0.42);
                th.Grid = ColorMath.Mix(picked, th.Text, 0.12);
            }
            else
            {
                th.Text = picked;
                th.Muted = ColorMath.Mix(picked, th.Card, 0.42);
                th.Grid = ColorMath.Mix(th.Card, picked, 0.12);
            }
            th.Name = "Own colours";
            LookToControls();
            RefreshSwatches();
            Changed();
            if (Math.Abs(ColorMath.Lum(th.Text) - ColorMath.Lum(th.Card)) < 0.3) SetHint("The text colour is close to the card colour, so words on cards and charts will be hard to read.");
            else SetHint(SwatchNames[i] + " changed. The theme is now your own; pick a theme again to go back.");
        }

        void LookToControls()
        {
            bool was = loading;
            loading = true;
            int ti = DashThemes.IndexOf(s.Theme.Name);
            lookTheme.SelectedIndex = ti >= 0 && ti < lookTheme.Items.Count ? ti : -1;
            lookTheme.Invalidate();
            for (int i = 0; i < swatches.Length; i++) swatches[i].BackColor = Rgb(SwatchColor(i));
            Pick(lookCards, s.CardStyle);
            lookRounded.Checked = s.Rounded;
            lookShadow.Checked = s.Shadow;
            int fi = Array.IndexOf(DashThemes.Fonts, s.Theme.Font);
            Pick(lookFont, fi < 0 ? 0 : fi);
            Pick(lookDensity, s.Density);
            Pick(lookUnit, s.Unit);
            lookCurrency.Text = s.Currency;
            loading = was;
        }

        void ThemeChanged()
        {
            int i = lookTheme.SelectedIndex;
            if (i < 0) return;
            string font = s.Theme.Font;
            s.Theme = DashThemes.Get(i);
            s.Theme.Font = font;
            loading = true;
            for (int k = 0; k < swatches.Length; k++) swatches[k].BackColor = Rgb(SwatchColor(k));
            loading = false;
            RefreshSwatches();
            Changed();
        }

        void LookEdited()
        {
            s.CardStyle = Math.Max(0, lookCards.SelectedIndex);
            s.Rounded = lookRounded.Checked;
            s.Shadow = lookShadow.Checked;
            if (lookFont.SelectedIndex >= 0) s.Theme.Font = DashThemes.Fonts[lookFont.SelectedIndex];
            s.Density = Math.Max(0, lookDensity.SelectedIndex);
            s.Unit = Math.Max(0, lookUnit.SelectedIndex);
            s.Currency = lookCurrency.Text.Trim();
            Changed();
        }

        // the colour squares on the Cards and Charts tabs follow the theme
        void RefreshSwatches()
        {
            CardSpec c = CurrentCard();
            if (c != null) cardSwatch.BackColor = Rgb(ColorMath.Pick(s.Theme, c.Color, cardList.SelectedIndex));
            ChartSpec ch = CurrentChart();
            if (ch != null) chSwatch.BackColor = Rgb(DashLook.ChartColor(s, ch, chartList.SelectedIndex));
        }

        // ------------------------------------------------------------------ tab 6: the page
        void BuildPageTab()
        {
            NewTab("Page");
            pgTitle = Input("Title", PageEdited);
            pgSubtitle = Input("Line under it", PageEdited);
            tips.SetToolTip(pgSubtitle, "Leave empty for the source, the number of rows and the time of building.");
            pgFooter = Input("Footer", PageEdited);
            tips.SetToolTip(pgFooter, "The small line at the bottom, for example the name of your firm.");
            Caption("Logo", S(10), S(118));
            Button pick = MakeButton(page, "Choose picture...", S(130), cy, S(140));
            Button drop = MakeButton(page, "Remove", S(276), cy, S(84));
            pick.Height = drop.Height = S(26);
            cy += S(32);
            pgLogo = Note("", 1);
            pick.Click += delegate { Do(PickLogo); };
            drop.Click += delegate { Do(DropLogo); };
            Heading("Layout");
            pgArrange = Combo("Charts", DashLayout.ArrangeNames, PageEdited);
            Button reset = MakeButton(page, "Give every chart its automatic size", S(130), cy, S(Edge - 130));
            reset.Height = S(26);
            cy += S(34);
            reset.Click += delegate { Do(ResetSizes); };
            pgWidth = Combo("Page width", DashLayout.WidthNames, PageEdited);
            pgPerRow = Combo("Cards in a row", new string[] { "Automatic", "2", "3", "4", "5", "6", "7", "8" }, PageEdited);
            pgInsights = Check("Key insights in words under the charts", PageEdited);
            Heading("In Excel");
            pgHeadings = Check("Hide the row numbers and column letters", PageEdited);
            pgPrint = Check("Set up printing: landscape, one page wide", PageEdited);
            EndTab();
        }

        void PageToControls()
        {
            bool was = loading;
            loading = true;
            pgTitle.Text = s.Title;
            pgSubtitle.Text = s.Subtitle;
            pgFooter.Text = s.Footer;
            pgLogo.Text = s.LogoData.Length > 0 ? "Logo: " + (s.LogoName.Length > 0 ? s.LogoName : "a picture") + ". It stands at the right of the title band." : "No logo.";
            Pick(pgArrange, s.Arrange);
            Pick(pgWidth, s.Width);
            Pick(pgPerRow, s.CardsPerRow <= 1 ? 0 : s.CardsPerRow - 1);
            pgInsights.Checked = s.Insights;
            pgHeadings.Checked = s.HideHeadings;
            pgPrint.Checked = s.PrintSetup;
            loading = was;
        }

        void PageEdited()
        {
            s.Title = pgTitle.Text;
            s.Subtitle = pgSubtitle.Text;
            s.Footer = pgFooter.Text;
            s.Arrange = Math.Max(0, pgArrange.SelectedIndex);
            s.Width = Math.Max(0, pgWidth.SelectedIndex);
            int per = pgPerRow.SelectedIndex;
            s.CardsPerRow = per <= 0 ? 0 : per + 1;
            s.Insights = pgInsights.Checked;
            s.HideHeadings = pgHeadings.Checked;
            s.PrintSetup = pgPrint.Checked;
            Changed();
        }

        void ResetSizes()
        {
            foreach (ChartSpec c in s.Charts)
            {
                c.Span = 0;
                c.Height = 0;
            }
            ChartToControls();
            Changed();
        }

        // The logo is kept inside the design as a small PNG, so it travels with the dashboard and with a saved design.
        void PickLogo()
        {
            string file;
            using (OpenFileDialog d = new OpenFileDialog())
            {
                d.Title = "Choose the logo";
                d.Filter = "Pictures|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files|*.*";
                if (d.ShowDialog(this) != DialogResult.OK) return;
                file = d.FileName;
            }
            string data;
            try
            {
                using (Image img = Image.FromFile(file))
                {
                    double k = Math.Min(1.0, Math.Min(480.0 / img.Width, 160.0 / img.Height));
                    int w = Math.Max(1, (int)Math.Round(img.Width * k)), h = Math.Max(1, (int)Math.Round(img.Height * k));
                    using (Bitmap small = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                    {
                        using (Graphics g = Graphics.FromImage(small))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            g.DrawImage(img, new Rectangle(0, 0, w, h));
                        }
                        using (MemoryStream ms = new MemoryStream())
                        {
                            small.Save(ms, ImageFormat.Png);
                            data = Convert.ToBase64String(ms.ToArray());
                        }
                    }
                }
            }
            catch
            {
                // GDI+ answers "out of memory" for a file it cannot read as a picture
                MessageBox.Show(this, "This file could not be read as a picture. Use a PNG or JPG file.", Design.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            s.LogoData = data;
            s.LogoName = Path.GetFileName(file);
            PageToControls();
            Changed();
        }

        void DropLogo()
        {
            s.LogoData = "";
            s.LogoName = "";
            PageToControls();
            Changed();
        }

        // the logo as a picture, read once from the design
        Image LogoImage()
        {
            if (s.LogoData.Length == 0) return null;
            if (logoOf == s.LogoData) return logo;
            if (logo != null)
            {
                logo.Dispose();
                logo = null;
            }
            logoOf = s.LogoData;
            try
            {
                using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(s.LogoData)))
                using (Image img = Image.FromStream(ms))
                {
                    logo = new Bitmap(img);
                }
            }
            catch
            {
                logo = null;
            }
            return logo;
        }

        // ------------------------------------------------------------------ the whole design to and from the window
        void AllToControls()
        {
            bool was = loading;
            loading = true;
            t.UseDate(s.DateCol);
            CardsToList();
            ChartsToList();
            FiltersToControls();
            LookToControls();
            PageToControls();
            loading = was;
            gallery.Invalidate();
            preview.Relayout();
        }

        // after every change of the design: the preview and the small pictures are drawn again
        void Changed()
        {
            version++;
            cache.Trim(600);
            gallery.Invalidate();
            preview.Relayout();
        }

        internal void SelectBox(int kind, int index)
        {
            if (selKind == kind && selIndex == index) return;
            selKind = kind;
            selIndex = index;
            preview.Redraw();
        }

        // a click in the preview opens the tab of what was clicked
        internal void PreviewClicked(int kind, int index)
        {
            if (loading) return;
            try
            {
                if (kind == Box.Card && index >= 0 && index < s.Cards.Count)
                {
                    tabs.SelectedIndex = 1;
                    if (index < cardList.Items.Count) cardList.SelectedIndex = index;
                }
                else if (kind == Box.Chart && index >= 0 && index < s.Charts.Count)
                {
                    tabs.SelectedIndex = 2;
                    if (index < chartList.Items.Count) chartList.SelectedIndex = index;
                }
                else if (kind == Box.Slicer) tabs.SelectedIndex = 3;
                else if (kind == Box.Header || kind == Box.Insights) tabs.SelectedIndex = 5;
                SelectBox(kind, index);
            }
            catch (Exception ex) { Oops(ex); }
        }

        // draws the dashboard 'width' pixels wide
        internal void PaintPreview(Graphics g, int width, List<HitBox> hits)
        {
            DashRender.Paint(g, new RectangleF(0, 0, width, 10), t, s, cache, selKind, selIndex, hits);
            Image img = LogoImage();
            if (img == null) return;
            DashGrid grid = cache.Grid(t, s);
            Box head = grid.Find(Box.Header, 0);
            if (head == null || grid.Cols < 1) return;
            float u = (float)width / grid.Cols;
            RectangleF area = DashRender.LogoBox(new RectangleF(head.X * u, head.Y * u, head.W * u, head.H * u), grid, u);
            if (area.Width < 2f || area.Height < 2f) return;
            float k = Math.Min(area.Width / img.Width, area.Height / img.Height);
            float w = img.Width * k, h = img.Height * k;
            InterpolationMode was = g.InterpolationMode;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(img, area.Right - w, area.Y + (area.Height - h) / 2f, w, h);
            g.InterpolationMode = was;
        }

        // the height of the dashboard when it is drawn 'width' pixels wide
        internal float PreviewHeight(int width)
        {
            DashGrid grid = cache.Grid(t, s);
            return grid.Cols < 1 ? 10f : grid.Rows * ((float)width / grid.Cols);
        }

        // ------------------------------------------------------------------ buttons at the bottom
        void Accept()
        {
            DashAuto.Fix(s, t);
            if (s.Cards.Count == 0 && s.Charts.Count == 0)
            {
                MessageBox.Show(this, "Add at least one KPI card or one chart.", Design.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Result = s;
            DialogResult = DialogResult.OK;
        }

        void StartAgain()
        {
            if (MessageBox.Show(this, "Start again from the design suggested for this table?\n\nThe cards, charts and filters you set here are replaced. Colours and layout stay.",
                                Design.AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            DashSpec look = new DashSpec();
            look.CopyLook(s);
            string title = s.Title, subtitle = s.Subtitle, footer = s.Footer, logoData = s.LogoData, logoName = s.LogoName;
            s = DashAuto.Suggest(t, look);
            s.Title = title;
            s.Subtitle = subtitle;
            s.Footer = footer;
            s.LogoData = logoData;
            s.LogoName = logoName;
            selKind = -1;
            selIndex = -1;
            version++;
            AllToControls();
            SetHint("Back to the design suggested for this table.");
        }

        void SaveDesign()
        {
            using (SaveFileDialog d = new SaveFileDialog())
            {
                d.Title = "Save this design";
                d.Filter = "DocBrisk dashboard design|*.dbdash";
                d.DefaultExt = "dbdash";
                d.AddExtension = true;
                d.OverwritePrompt = true;
                d.InitialDirectory = DashboardStudio.DesignFolder();
                d.FileName = Xl.SafeFileName(s.Title.Trim().Length > 0 ? s.Title.Trim() : "Dashboard design");
                if (d.ShowDialog(this) != DialogResult.OK) return;
                File.WriteAllText(d.FileName, s.ToText(), new System.Text.UTF8Encoding(false));
                SetHint("Design saved: " + d.FileName);
            }
        }

        void OpenDesign()
        {
            string file;
            using (OpenFileDialog d = new OpenFileDialog())
            {
                d.Title = "Open a saved design";
                d.Filter = "DocBrisk dashboard design|*.dbdash|All files|*.*";
                d.InitialDirectory = DashboardStudio.DesignFolder();
                if (d.ShowDialog(this) != DialogResult.OK) return;
                file = d.FileName;
            }
            DashSpec got = DashSpec.FromText(File.ReadAllText(file));
            if (got == null)
            {
                MessageBox.Show(this, "This file is not a DocBrisk dashboard design.", Design.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            List<string> notes = DashAuto.Fix(got, t);
            version++;
            selKind = -1;
            selIndex = -1;
            if (got.Cards.Count == 0 && got.Charts.Count == 0)
            {
                // none of its columns are in this table: only the look is taken over
                s.CopyLook(got);
                AllToControls();
                SetHint("This design was made for a table with other columns. Its colours and layout were taken over; your cards and charts stay.");
                return;
            }
            s = got;
            AllToControls();
            SetHint(notes.Count == 0 ? "Design opened: " + Path.GetFileName(file)
                                     : "Design opened. " + notes.Count.ToString(CultureInfo.InvariantCulture) + " part(s) were left out: their columns are not in this table.");
        }
    }

    // The live preview: the dashboard drawn to the width of the panel, with a scroll bar when it is longer.
    // It is drawn once into a picture; scrolling and uncovering only copy from that picture.
    internal sealed class PreviewPanel : Panel
    {
        readonly StudioForm owner;
        readonly List<HitBox> hits = new List<HitBox>();
        Bitmap canvas;
        bool stale = true;
        int laidOut = -1;                 // the width the scroll range was worked out for

        public PreviewPanel(StudioForm form)
        {
            owner = form;
            AutoScroll = true;
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Color.FromArgb(203, 213, 225);
        }

        // The dashboard is drawn this wide. The room of the scroll bar is always left free,
        // so that the bar coming or going cannot change the width (and with it the height) again.
        int PageWidth
        {
            get { return Math.Max(60, Width - SystemInformation.VerticalScrollBarWidth - 2); }
        }

        // the design changed: work out the height again and draw again
        public void Relayout()
        {
            stale = true;
            if (owner == null) return;
            int width = PageWidth;
            laidOut = width;
            int need;
            try { need = (int)Math.Ceiling(owner.PreviewHeight(width)); }
            catch { need = 0; }
            if (need < 0) need = 0;
            if (need > 30000) need = 30000;
            if (AutoScrollMinSize.Height != need || AutoScrollMinSize.Width != 0) AutoScrollMinSize = new Size(0, need);
            Invalidate();
        }

        // only the selection changed
        public void Redraw()
        {
            stale = true;
            Invalidate();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && canvas != null)
            {
                canvas.Dispose();
                canvas = null;
            }
            base.Dispose(disposing);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (owner != null && PageWidth != laidOut) Relayout();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            try
            {
                int width = PageWidth;
                if (canvas == null || stale || canvas.Width != width)
                {
                    int height = (int)Math.Ceiling(owner.PreviewHeight(width));
                    if (height < 1) height = 1;
                    if (height > 30000) height = 30000;
                    if (canvas == null || canvas.Width != width || canvas.Height != height)
                    {
                        if (canvas != null) canvas.Dispose();
                        canvas = null;
                        canvas = new Bitmap(width, height, PixelFormat.Format32bppPArgb);
                    }
                    hits.Clear();
                    using (Graphics cg = Graphics.FromImage(canvas))
                    {
                        owner.PaintPreview(cg, width, hits);
                    }
                    stale = false;
                }
                g.DrawImage(canvas, AutoScrollPosition.X, AutoScrollPosition.Y, canvas.Width, canvas.Height);
            }
            catch (Exception ex)
            {
                stale = true;
                g.Clear(Color.White);
                using (Font f = new Font("Segoe UI", 9f))
                {
                    g.DrawString("The preview could not be drawn:\n" + ex.Message, f, Brushes.Firebrick, new RectangleF(12, 12, Math.Max(50, ClientSize.Width - 24), 200));
                }
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();        // so that the mouse wheel scrolls the preview
            float x = e.X - AutoScrollPosition.X, y = e.Y - AutoScrollPosition.Y;
            // the last box drawn that holds the point
            for (int i = hits.Count - 1; i >= 0; i--)
            {
                if (hits[i].R.Contains(x, y))
                {
                    owner.PreviewClicked(hits[i].Kind, hits[i].Index);
                    return;
                }
            }
        }
    }

    // The gallery of ready designs: small pictures of the dashboard, drawn when they scroll into view.
    internal sealed class GalleryPanel : Panel
    {
        readonly StudioForm owner;
        readonly List<DesignIdea> ideas;
        readonly Bitmap[] pictures;
        readonly Random chance = new Random();
        int shown = -1;                   // the version of the design the pictures were drawn for
        int tileW = -1;                   // and their width

        public GalleryPanel(StudioForm form, bool hasDate)
        {
            owner = form;
            ideas = DashAuto.Ideas(hasDate);
            pictures = new Bitmap[ideas.Count];
            AutoScroll = true;
            DoubleBuffered = true;
            ResizeRedraw = true;
            BorderStyle = BorderStyle.FixedSingle;
            BackColor = Color.FromArgb(241, 245, 249);
        }

        public int Count { get { return ideas.Count; } }

        int Gap { get { return owner.S(8); } }
        int TileW { get { return Math.Max(40, (Width - SystemInformation.VerticalScrollBarWidth - 4 - 3 * Gap) / 2); } }
        int PicH { get { return (int)(TileW * 0.8); } }
        int TileH { get { return PicH + owner.S(20); } }

        void DropPictures()
        {
            for (int i = 0; i < pictures.Length; i++)
            {
                if (pictures[i] != null) pictures[i].Dispose();
                pictures[i] = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && pictures != null) DropPictures();
            base.Dispose(disposing);
        }

        Rectangle TileAt(int i)
        {
            int col = i % 2, row = i / 2;
            return new Rectangle(Gap + col * (TileW + Gap), Gap + row * (TileH + Gap), TileW, TileH);
        }

        // one of the designs by chance
        public void PickByChance()
        {
            if (ideas.Count == 0) return;
            int i = chance.Next(ideas.Count);
            owner.UseIdea(ideas[i]);
            AutoScrollPosition = new Point(0, Math.Max(0, TileAt(i).Top - Gap));
            Invalidate();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (owner == null || ideas == null) return;
            int rows = (ideas.Count + 1) / 2;
            Size need = new Size(0, Gap + rows * (TileH + Gap));
            if (AutoScrollMinSize != need) AutoScrollMinSize = need;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            try
            {
                // the small pictures are drawn again when the design changed or the tiles got another size
                if (shown != owner.Version || tileW != TileW)
                {
                    DropPictures();
                    shown = owner.Version;
                    tileW = TileW;
                }
                int dy = AutoScrollPosition.Y;
                using (Font f = new Font("Segoe UI", 7.5f))
                using (SolidBrush ink = new SolidBrush(Color.FromArgb(51, 65, 85)))
                using (Pen mark = new Pen(Color.FromArgb(37, 99, 235), 3f))
                using (Pen edge = new Pen(Color.FromArgb(203, 213, 225), 1f))
                using (StringFormat sf = new StringFormat(StringFormatFlags.NoWrap))
                {
                    sf.Trimming = StringTrimming.EllipsisCharacter;
                    for (int i = 0; i < ideas.Count; i++)
                    {
                        Rectangle r = TileAt(i);
                        r.Offset(0, dy);
                        if (r.Bottom < 0 || r.Top > ClientSize.Height) continue;
                        if (pictures[i] == null)
                        {
                            try { pictures[i] = owner.Thumbnail(ideas[i], TileW, PicH); }
                            catch { pictures[i] = new Bitmap(8, 8); }
                        }
                        g.DrawImage(pictures[i], r.X, r.Y, pictures[i].Width, pictures[i].Height);
                        bool on = owner.Matches(ideas[i]);
                        g.DrawRectangle(on ? mark : edge, r.X, r.Y, TileW - 1, PicH - 1);
                        g.DrawString((i + 1).ToString(CultureInfo.InvariantCulture) + ".  " + ideas[i].Name, f, ink, new RectangleF(r.X, r.Y + PicH + 2, TileW, owner.S(16)), sf);
                    }
                }
            }
            catch { }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();        // so that the mouse wheel scrolls the gallery
            for (int i = 0; i < ideas.Count; i++)
            {
                Rectangle r = TileAt(i);
                r.Offset(0, AutoScrollPosition.Y);
                if (!r.Contains(e.Location)) continue;
                owner.UseIdea(ideas[i]);
                Invalidate();
                return;
            }
        }
    }
}
