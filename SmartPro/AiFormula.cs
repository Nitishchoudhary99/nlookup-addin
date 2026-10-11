using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace SmartPro
{
    // AI FORMULA: describe what you want in your own words (English or Hinglish), get one Excel formula for the active cell,
    // with a short explanation, using Claude.
    //   - It goes through docbrisk.com with the user's DocBrisk key: no Anthropic key is needed on the PC.
    //     The website checks the key and counts: Pro 100 a day, a free key 3 a day (set on the website), and asks
    //     Claude with its own key and its own fixed instructions (seo-worker.js, "AI FORMULA PROXY").
    //   - What is sent: the request, the column headers of the table around the active cell and the cell address.
    //     Three sample rows only when the user ticks the box. The window shows exactly what will be sent.
    //   - Every other DocBrisk tool works offline; this is the only one that needs the internet.
    internal static class AiFormula
    {
        public const string Model = "claude-opus-5-5";
        public const string ProxyUrl = "https://docbrisk.com/addins-api/ai/formula";
        public const string FreeKeyUrl = "https://docbrisk.com/addins#ax-key";

        // a test or staging Worker can be set in the registry value AiProxyUrl (HKCU\Software\DocBrisk\SmartPro)
        static string Endpoint
        {
            get
            {
                string u = Settings.Get("AiProxyUrl");
                return !string.IsNullOrEmpty(u) && (u.StartsWith("https://", StringComparison.OrdinalIgnoreCase) || u.StartsWith("http://127.0.0.1", StringComparison.Ordinal)) ? u : ProxyUrl;
            }
        }

        // the answer of the docbrisk.com proxy
        internal sealed class Answer
        {
            public string Formula, Why, Error, Code, Tier;
            public int Left = -1, Limit = -1;
        }

        public static void Run()
        {
            dynamic app = Xl.App;
            dynamic cell = app.ActiveCell;
            if (cell == null) { Ui.Info("Click the cell where the formula should go."); return; }
            string context = Describe(cell, false), sample = Describe(cell, true);
            using (AiFormulaForm f = new AiFormulaForm(((string)cell.Address).Replace("$", ""), context, sample))
            {
                if (f.ShowDialog(new ExcelWin()) != DialogResult.OK || string.IsNullOrEmpty(f.Formula)) return;
                Insert(cell, f.Formula, f.FillDown);
            }
        }

        static void Insert(dynamic cell, string formula, bool fillDown)
        {
            try { cell.Formula = formula; }
            catch
            {
                try { cell.Formula2 = formula; }                       // Excel 365 dynamic-array formulas
                catch (Exception ex) { Ui.Warn("Excel did not accept the formula:\n" + formula + "\n\n" + ex.Message + "\n\nIt was copied to the clipboard."); Clipboard.SetText(formula); return; }
            }
            if (!fillDown) return;
            try
            {
                dynamic region = cell.CurrentRegion;
                int last = (int)region.Row + (int)region.Rows.Count - 1;
                if (last > (int)cell.Row) cell.Worksheet.Range[cell, cell.Worksheet.Cells[last, (int)cell.Column]].FillDown();
            }
            catch { }
        }

        // the table around the cell: headers with their column letters and the data rows, in words
        static string Describe(dynamic cell, bool withSample)
        {
            StringBuilder sb = new StringBuilder();
            string addr = ((string)cell.Address).Replace("$", "");
            sb.Append("Sheet: ").Append((string)cell.Worksheet.Name).Append(". The formula goes in cell ").Append(addr).Append(".\n");
            try
            {
                dynamic region = cell.CurrentRegion;
                if (Xl.CellCount(region) <= 1)
                {
                    int left = (int)cell.Column > 1 ? (int)cell.Column - 1 : 1;
                    region = cell.Offset[0, -Math.Min(1, (int)cell.Column - 1)].CurrentRegion;
                }
                int r0 = (int)region.Row, c0 = (int)region.Column, nr = (int)region.Rows.Count, nc = Math.Min((int)region.Columns.Count, 40);
                object[,] v = Xl.ReadTyped(region.Resize[Math.Min(nr, 4), nc]);
                sb.Append("Table ").Append(Xl.A1(r0, c0)).Append(":").Append(Xl.A1(r0 + nr - 1, c0 + nc - 1)).Append(", header in row ").Append(r0)
                  .Append(", data in rows ").Append(r0 + 1).Append(" to ").Append(r0 + nr - 1).Append(".\nColumns:\n");
                for (int j = 1; j <= nc; j++)
                {
                    sb.Append("  ").Append(Xl.ColLetter(c0 + j - 1)).Append(" = ").Append(Xl.HeaderText(v[1, j], j));
                    if (v.GetLength(0) >= 2)
                    {
                        object x = v[2, j];
                        sb.Append(x is double ? " (number)" : x is DateTime ? " (date)" : Xl.IsBlank(x) ? "" : " (text)");
                    }
                    sb.Append("\n");
                }
                if (withSample && v.GetLength(0) >= 2)
                {
                    sb.Append("Sample rows:\n");
                    for (int r = 2; r <= v.GetLength(0); r++)
                    {
                        sb.Append("  row ").Append(r0 + r - 1).Append(": ");
                        for (int j = 1; j <= nc; j++) sb.Append(j > 1 ? " | " : "").Append(Xl.Str(v[r, j]));
                        sb.Append("\n");
                    }
                }
            }
            catch { sb.Append("(No table around the cell.)\n"); }
            return sb.ToString();
        }

        // through docbrisk.com with the DocBrisk key
        public static async Task<Answer> AskDocBrisk(string licenceKey, string request, string context)
        {
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }
            string body = new JavaScriptSerializer().Serialize(new Dictionary<string, string>
            {
                { "key", licenceKey }, { "request", request }, { "context", context }, { "client", "excel-" + Design.Version }
            });
            using (HttpClient http = new HttpClient())
            {
                http.Timeout = TimeSpan.FromSeconds(90);
                using (StringContent content = new StringContent(body, Encoding.UTF8, "application/json"))
                using (HttpResponseMessage res = await http.PostAsync(Endpoint, content).ConfigureAwait(false))
                {
                    string json = await res.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return ParseProxy((int)res.StatusCode, json);
                }
            }
        }

        public static Answer ParseProxy(int status, string json)
        {
            Answer a = new Answer();
            try
            {
                Dictionary<string, object> r = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
                a.Formula = Text(r, "formula");
                a.Why = Text(r, "why");
                a.Error = Text(r, "error");
                a.Code = Text(r, "code");
                a.Tier = Text(r, "tier");
                a.Left = Whole(r, "left");
                a.Limit = Whole(r, "limit");
            }
            catch
            {
                a.Error = "docbrisk.com sent an answer that could not be read (HTTP " + status + ").";
                a.Code = "bad_answer";
            }
            if (status != 200 && a.Error == null) a.Error = "docbrisk.com answered HTTP " + status + ".";
            if (a.Formula != null && !a.Formula.StartsWith("=", StringComparison.Ordinal)) a.Formula = null;
            return a;
        }

        static string Text(Dictionary<string, object> r, string name)
        {
            object v;
            return r != null && r.TryGetValue(name, out v) ? v as string : null;
        }

        static int Whole(Dictionary<string, object> r, string name)
        {
            object v;
            return r != null && r.TryGetValue(name, out v) && v is int ? (int)v : -1;
        }

        public static string[] Parse(string text)
        {
            string formula = null, why = "";
            foreach (string raw in (text ?? "").Replace("\r", "").Split('\n'))
            {
                string line = raw.Trim().Trim('`');
                if (line.StartsWith("FORMULA:", StringComparison.OrdinalIgnoreCase)) formula = line.Substring(8).Trim().Trim('`');
                else if (line.StartsWith("WHY:", StringComparison.OrdinalIgnoreCase)) why = line.Substring(4).Trim();
                else if (why.Length > 0 && line.Length > 0) why += " " + line;
                else if (formula == null && line.StartsWith("=", StringComparison.Ordinal)) formula = line;
            }
            if (formula != null && !formula.StartsWith("=", StringComparison.Ordinal)) formula = null;   // "none"
            return new[] { formula, why };
        }
    }

    internal sealed class AiFormulaForm : Form
    {
        static readonly Color Navy = Color.FromArgb(15, 23, 42), Slate = Color.FromArgb(71, 85, 105), Blue = Color.FromArgb(37, 99, 235), Red = Color.FromArgb(185, 28, 28);
        readonly TextBox ask = new TextBox(), result = new TextBox(), sent = new TextBox();
        readonly Label why = new Label(), status = new Label(), plan = new Label();
        readonly LinkLabel getKey = new LinkLabel(), enterKey = new LinkLabel();
        readonly CheckBox withSample = new CheckBox(), fill = new CheckBox();
        readonly Button go = new Button(), insert = new Button();
        readonly string context, sample;
        readonly float z = 1f;
        public string Formula;
        public bool FillDown;
        string leftNote = "";

        int S(int px) { return (int)Math.Round(px * z); }

        Label Lbl(string text, int x, int y, int w, int h, bool bold)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(S(x), S(y), S(w), S(h));
            l.Font = new Font("Segoe UI", 9.5f, bold ? FontStyle.Bold : FontStyle.Regular);
            l.ForeColor = bold ? Navy : Slate;
            Controls.Add(l);
            return l;
        }

        public AiFormulaForm(string cell, string context, string sample)
        {
            this.context = context;
            this.sample = sample;
            try { using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) { z = Math.Max(1f, Math.Min(2.5f, g.DpiX / 96f)); } } catch { }
            AutoScaleMode = AutoScaleMode.None;
            Text = Design.AppName + " - AI Formula   (for " + cell + ")";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);
            ClientSize = new Size(S(760), S(544));

            Lbl("What should the formula do?  (your own words)", 16, 12, 700, 22, true);
            ask.Multiline = true;
            ask.SetBounds(S(16), S(36), S(728), S(56));
            ask.Text = "";
            Controls.Add(ask);
            Lbl("Example: total amount for this customer in this month  |  days overdue, blank if paid  |  grade A above 90, B above 75, else C", 16, 94, 728, 20, false).Font = new Font("Segoe UI", 8.5f);

            withSample.Text = "Also send 3 sample rows (better formulas; only the headers are sent otherwise)";
            withSample.SetBounds(S(16), S(118), S(600), S(24));
            withSample.CheckedChanged += delegate { ShowSent(); };
            Controls.Add(withSample);

            go.Text = "Write the formula";
            go.SetBounds(S(584), S(114), S(160), S(32));
            go.FlatStyle = FlatStyle.Flat;
            go.FlatAppearance.BorderSize = 0;
            go.BackColor = Blue;
            go.ForeColor = Color.White;
            go.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            go.Click += async delegate { await Generate(); };
            Controls.Add(go);

            Lbl("Formula", 16, 156, 200, 22, true);
            result.SetBounds(S(16), S(180), S(728), S(28));
            result.Font = new Font("Consolas", 10.5f);
            Controls.Add(result);
            why.SetBounds(S(16), S(212), S(728), S(44));
            why.ForeColor = Slate;
            Controls.Add(why);

            fill.Text = "Fill it down to the end of the table";
            fill.Checked = true;
            fill.SetBounds(S(16), S(258), S(300), S(24));
            Controls.Add(fill);
            insert.Text = "Insert in " + cell;
            insert.Enabled = false;
            insert.SetBounds(S(584), S(254), S(160), S(32));
            insert.Click += delegate
            {
                if ((insert.Tag as string) == "pro")
                {
                    if (ProCentre.Show(null)) { insert.Tag = null; insert.Text = "Insert in " + cell; insert.Enabled = false; ShowPlan(); Say("Pro is active: 100 AI formulas a day.", Color.FromArgb(21, 128, 61)); }
                    return;
                }
                Formula = result.Text.Trim();
                FillDown = fill.Checked;
                if (!Formula.StartsWith("=", StringComparison.Ordinal)) { Say("A formula starts with =.", Red); return; }
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(insert);

            Lbl("What is sent to Claude (by Anthropic) with your request:", 16, 296, 728, 20, true);
            sent.Multiline = true;
            sent.ReadOnly = true;
            sent.ScrollBars = ScrollBars.Vertical;
            sent.Font = new Font("Consolas", 8.5f);
            sent.SetBounds(S(16), S(318), S(728), S(120));
            Controls.Add(sent);

            plan.SetBounds(S(16), S(446), S(728), S(22));
            plan.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            Controls.Add(plan);
            getKey.Text = "Get a free DocBrisk key";
            getKey.SetBounds(S(16), S(468), S(200), S(22));
            getKey.LinkClicked += delegate { try { System.Diagnostics.Process.Start(AiFormula.FreeKeyUrl); } catch { } Say("Fill in your name and email on the page, copy the key, then press 'Enter my key'.", Slate); };
            Controls.Add(getKey);
            enterKey.Text = "Enter my key";
            enterKey.SetBounds(S(224), S(468), S(120), S(22));
            enterKey.LinkClicked += delegate { ProCentre.Show(null); ShowPlan(); };
            Controls.Add(enterKey);
            status.SetBounds(S(16), S(496), S(630), S(40));
            Controls.Add(status);
            Button close = new Button();
            close.Text = "Close";
            close.DialogResult = DialogResult.Cancel;
            close.SetBounds(S(658), S(500), S(86), S(30));
            Controls.Add(close);
            CancelButton = close;
            ShowPlan();
            ActiveControl = ask;
        }

        // which way the request goes, and what that means for the user
        void ShowPlan()
        {
            string lic = Licence.StoredKey;
            bool hasLic = !string.IsNullOrEmpty(lic) && Licence.Current().Valid;
            if (hasLic && Licence.IsProNow()) { plan.Text = "Included in your Pro plan: 100 AI formulas a day. No other key needed."; plan.ForeColor = Color.FromArgb(21, 128, 61); }
            else if (hasLic) { plan.Text = "Free key: 3 AI formulas a day. Pro: 100 a day (" + ProPlan.PerDay + ")."; plan.ForeColor = Navy; }
            else { plan.Text = "Free with a DocBrisk key: 3 AI formulas a day. Get one in a minute (name + email)."; plan.ForeColor = Red; }
            getKey.Visible = enterKey.Visible = !hasLic;
            ShowSent();
        }

        void ShowSent()
        {
            string head = "(sent to docbrisk.com with your DocBrisk key, which checks your plan; docbrisk.com asks Claude)\r\n";
            sent.Text = head + (withSample.Checked ? sample : context).Replace("\n", "\r\n") + "\r\nRequest: (your words above)";
        }

        void Say(string s, Color c)
        {
            status.ForeColor = c;
            status.Text = s;
        }

        async Task Generate()
        {
            string req = ask.Text.Trim();
            if (req.Length < 3) { Say("Describe what the formula should do first.", Red); return; }
            string lic = Licence.StoredKey;
            if (string.IsNullOrEmpty(lic) || !Licence.Current().Valid)
            {
                Say("Get a free DocBrisk key first (link above): 3 AI formulas a day, free.", Red);
                return;
            }
            go.Enabled = false;
            insert.Enabled = false;
            Say("Writing the formula ...", Slate);
            try
            {
                string[] r;
                {
                    AiFormula.Answer a = await AiFormula.AskDocBrisk(lic, req, withSample.Checked ? sample : context);
                    if (a.Error != null)
                    {
                        Say(a.Error, Red);
                        if (a.Code == "limit" || a.Code == "expired" || a.Code == "not_pro") enterKey.Visible = getKey.Visible = false;
                        if ((a.Code == "limit" && a.Tier != "Pro") || a.Code == "expired" || a.Code == "not_pro")
                        {
                            insert.Text = "See Pro";
                            insert.Enabled = true;
                            insert.Tag = "pro";
                        }
                        return;
                    }
                    r = new[] { a.Formula, a.Why };
                    if (a.Left >= 0) leftNote = "  " + a.Left + " AI " + (a.Left == 1 ? "formula" : "formulas") + " left today" + (a.Tier == "Free" ? " (Pro: 100 a day)." : ".");
                }
                if (r[0] == null)
                {
                    result.Text = "";
                    why.Text = r[1];
                    Say("No formula for this request.", Red);
                }
                else
                {
                    result.Text = r[0];
                    why.Text = r[1];
                    insert.Enabled = true;
                    Say("Check the formula, change it if you like, then Insert." + leftNote, Color.FromArgb(21, 128, 61));
                }
            }
            catch (Exception ex) { Say("Could not reach docbrisk.com: " + ex.Message + "  (internet or office proxy?)", Red); }
            finally { go.Enabled = true; }
        }
    }
}
