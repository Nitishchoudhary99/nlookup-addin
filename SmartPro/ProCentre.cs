using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SmartPro
{
    // DOCBRISK PRO window: the plan in use, what Pro adds, the free runs that are left, the plans and the box for the key.
    // A key copied on the DocBrisk website is picked up from the clipboard, so paying and activating takes two clicks.
    internal static class ProCentre
    {
        // reason = why the window opened by itself (a Pro tool whose free runs are used), or null.
        // True when a Pro plan is active when the window closes.
        public static bool Show(string reason)
        {
            using (ProForm f = new ProForm(reason))
            {
                f.ShowDialog(new ExcelWin());
            }
            return Licence.IsProNow();
        }
    }

    internal sealed class ProForm : Form
    {
        static readonly Color Navy = Color.FromArgb(15, 23, 42), Slate = Color.FromArgb(71, 85, 105), Blue = Color.FromArgb(37, 99, 235),
                              Green = Color.FromArgb(21, 128, 61), Red = Color.FromArgb(185, 28, 28), Soft = Color.FromArgb(241, 245, 249);
        readonly float z = 1f;                       // screen scaling: 1 = 100%, 1.5 = 150%
        readonly Label plan = new Label();
        readonly Label msg = new Label();
        readonly ListView list = new ListView();
        readonly TextBox keyBox = new TextBox();
        readonly Button buy = new Button();

        int S(int px) { return (int)Math.Round(px * z); }

        Label Lbl(Control parent, string text, int x, int y, int w, int hgt, float size, bool bold, Color color)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(S(x), S(y), S(w), S(hgt));
            l.Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular);
            l.ForeColor = color;
            l.BackColor = Color.Transparent;
            parent.Controls.Add(l);
            return l;
        }

        public ProForm(string reason)
        {
            try
            {
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) { z = g.DpiX / 96f; }
            }
            catch { z = 1f; }
            if (z < 1f) z = 1f;
            if (z > 3f) z = 3f;
            int y0 = reason == null ? 78 : 112;
            try
            {
                // never taller than the screen
                float fit = (Screen.PrimaryScreen.WorkingArea.Height - 48) / (float)(y0 + 520);
                if (fit < z && fit > 0.6f) z = fit;
            }
            catch { }
            AutoScaleMode = AutoScaleMode.None;
            Text = Design.AppName + " Pro";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);
            ClientSize = new Size(S(880), S(y0 + 520));

            // ---- header: the plan in use ----
            Panel head = new Panel();
            head.SetBounds(0, 0, S(880), S(78));
            head.BackColor = Navy;
            Controls.Add(head);
            Lbl(head, Design.AppName + " Pro", 18, 8, 600, 34, 16f, true, Color.White);
            plan.SetBounds(S(20), S(46), S(840), S(22));
            plan.ForeColor = Color.FromArgb(203, 213, 225);
            plan.BackColor = Color.Transparent;
            head.Controls.Add(plan);
            if (reason != null)
            {
                Label why = Lbl(this, reason, 0, 78, 880, 34, 9.5f, true, Color.FromArgb(120, 53, 15));
                why.BackColor = Color.FromArgb(254, 243, 199);
                why.Padding = new Padding(S(18), 0, 0, 0);
                why.TextAlign = ContentAlignment.MiddleLeft;
            }

            // ---- left: what Pro adds ----
            Lbl(this, "What Pro adds   (" + Design.ProIds.Length + " tools, each with " + Licence.FreeRuns + " free runs to try)", 18, y0 + 12, 548, 20, 9.5f, true, Navy);
            list.SetBounds(S(18), S(y0 + 36), S(548), S(322));
            list.View = View.Details;
            list.FullRowSelect = true;
            list.MultiSelect = false;
            list.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            list.ShowItemToolTips = true;
            list.Columns.Add("Pro tool", S(148));
            list.Columns.Add("What it does for you", S(282));
            list.Columns.Add("On this PC", S(94));
            Controls.Add(list);

            // ---- right: the plans ----
            Panel box = new Panel();
            box.SetBounds(S(582), S(y0 + 12), S(280), S(346));
            box.BackColor = Soft;
            Controls.Add(box);
            Lbl(box, "Choose your plan", 14, 12, 252, 20, 9.5f, true, Slate);
            Lbl(box, ProPlan.YearLine, 14, 38, 258, 26, 12f, true, Navy);
            Lbl(box, ProPlan.SavingLine, 14, 64, 260, 20, 8.5f, false, Green);
            Lbl(box, ProPlan.MonthLine, 14, 90, 258, 24, 10.5f, true, Navy);
            int py = 122;
            foreach (string line in ProPlan.Promise)
            {
                if (py > 212) break;
                Lbl(box, "•  " + line, 14, py, 260, 22, 9f, false, Slate);
                py += 22;
            }
            buy.SetBounds(S(14), S(244), S(252), S(44));
            buy.FlatStyle = FlatStyle.Flat;
            buy.FlatAppearance.BorderSize = 0;
            buy.BackColor = Blue;
            buy.ForeColor = Color.White;
            buy.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            buy.Cursor = Cursors.Hand;
            box.Controls.Add(buy);
            Lbl(box, "Pay on the website and copy your key. This window picks it up from the clipboard.", 14, 296, 258, 40, 8.5f, false, Slate);

            // ---- bottom: the key ----
            int y1 = y0 + 372;
            Lbl(this, "Already have a key?  Paste it here and press Activate:", 18, y1, 700, 20, 9.5f, true, Navy);
            keyBox.SetBounds(S(18), S(y1 + 26), S(620), S(27));
            Controls.Add(keyBox);
            Button paste = new Button();
            paste.Text = "Paste";
            paste.SetBounds(S(648), S(y1 + 25), S(90), S(29));
            Controls.Add(paste);
            Button act = new Button();
            act.Text = "Activate";
            act.SetBounds(S(746), S(y1 + 25), S(116), S(29));
            Controls.Add(act);
            msg.SetBounds(S(18), S(y1 + 60), S(844), S(40));
            Controls.Add(msg);
            Lbl(this, Design.AppName + " tools run on this PC so far: " + Settings.GetInt("Runs", 0).ToString("#,##0", CultureInfo.InvariantCulture), 18, y1 + 112, 600, 20, 8.5f, false, Slate);
            Button close = new Button();
            close.Text = "Close";
            close.SetBounds(S(776), S(y1 + 106), S(86), S(30));
            close.DialogResult = DialogResult.Cancel;
            Controls.Add(close);
            AcceptButton = act;
            CancelButton = close;

            buy.Click += delegate
            {
                Licence.OpenBuyPage();
                Say("The DocBrisk page is open in your browser. Pay there, copy your key and come back: it is filled in here by itself.", Slate);
            };
            paste.Click += delegate
            {
                string k = Licence.ClipboardKey();
                if (k == null)
                {
                    try { k = Clipboard.ContainsText() ? Clipboard.GetText().Trim() : null; }
                    catch { k = null; }
                }
                if (string.IsNullOrEmpty(k)) Say("There is no text on the clipboard. Copy your key first: it starts with DBK1-.", Red);
                else
                {
                    keyBox.Text = k.Length > 2000 ? k.Substring(0, 2000) : k;
                    Say("Pasted. Press Activate.", Slate);
                }
            };
            act.Click += delegate { UseKey(); };
            Activated += delegate { Sniff(); };      // also when the user comes back from the browser
            Fill();
            ActiveControl = keyBox;
        }

        void Say(string text, Color color)
        {
            msg.ForeColor = color;
            msg.Text = text;
        }

        // the plan line, the button and the list of Pro tools with the free runs that are left
        void Fill()
        {
            bool pro = Licence.IsProNow();
            plan.Text = Licence.PlanText();
            buy.Text = pro ? "Renew Pro on docbrisk.com" : "Get Pro on docbrisk.com";
            list.BeginUpdate();
            list.Items.Clear();
            foreach (ToolDef t in Design.Tools)
            {
                if (!Design.IsPro(t.Id)) continue;
                int left = Licence.TrialLeft(t.Id);
                string state = pro ? "Unlocked" : (left > 0 ? left + (left == 1 ? " free run left" : " free runs left") : "Free runs used");
                ListViewItem it = new ListViewItem(new[] { t.Caption, t.Tip, state });
                it.ToolTipText = t.Tip;
                if (!pro && left == 0) it.ForeColor = Red;
                list.Items.Add(it);
            }
            list.EndUpdate();
        }

        // a valid new key on the clipboard is filled in without asking
        void Sniff()
        {
            if (keyBox.Text.Trim().Length > 0) return;
            string k = Licence.NewClipboardKey();
            if (k == null) return;
            keyBox.Text = k;
            Say("A DocBrisk key was found on your clipboard and filled in. Press Activate.", Green);
        }

        void UseKey()
        {
            string t = keyBox.Text.Trim();
            if (t.Length == 0)
            {
                Say("Paste your key first. It starts with DBK1-.", Red);
                return;
            }
            string err = Licence.Accept(t);
            if (err != null)
            {
                Say(err, Red);
                return;
            }
            keyBox.Text = "";
            Fill();
            Say(Licence.IsProNow() ? "Key accepted. Pro is active: every Pro tool is unlocked on this PC." : "Key accepted (Free plan). The Pro tools still need a Pro key.", Green);
        }
    }
}
