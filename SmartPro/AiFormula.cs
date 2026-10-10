using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Anthropic;
using Anthropic.Models.Beta.Messages;

namespace SmartPro
{
    // AI FORMULA: describe what you want in your own words (English or Hinglish), get one Excel formula for the active cell,
    // with a short explanation, using Claude.
    //   - What is sent: the request, the column headers of the table around the active cell and the cell address.
    //     Three sample rows only when the user ticks the box. The window shows exactly what will be sent.
    //   - The API key is the user's own Anthropic key, kept encrypted for this Windows user (DPAPI).
    //   - Every other DocBrisk tool works offline; this is the only one that needs the internet.
    internal static class AiFormula
    {
        public const string Model = "claude-opus-5-5";

        const string SystemPrompt =
            "You write Microsoft Excel formulas for office users in India (accounts, MIS, sales, HR).\n" +
            "Reply in exactly this shape and nothing else:\n" +
            "FORMULA: <one formula starting with =, written for the given cell, with English function names and commas between arguments>\n" +
            "WHY: <one or two short sentences in plain English saying what it does and anything to check>\n" +
            "Rules: use the real column letters and rows from the table description. Prefer functions that work in Excel 2016 " +
            "(SUMIFS, COUNTIFS, INDEX/MATCH, IFERROR, TEXT, DATEDIF) unless the user says they have Excel 365; when an Excel 365 " +
            "function such as XLOOKUP, FILTER or UNIQUE is clearly better, use it and say so in WHY. The formula must be ready to " +
            "fill down. If the request cannot be done with a formula, write FORMULA: none and explain in WHY.";

        // ---- the key (encrypted for the Windows user) ----
        public static string Key
        {
            get
            {
                string s = Settings.Get("AiKey");
                if (string.IsNullOrEmpty(s)) return null;
                try { return Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(s), null, DataProtectionScope.CurrentUser)); }
                catch { return null; }
            }
            set
            {
                if (string.IsNullOrEmpty(value)) { Settings.Set("AiKey", ""); return; }
                Settings.Set("AiKey", Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(value.Trim()), null, DataProtectionScope.CurrentUser)));
            }
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

        // one request to Claude; returns the formula and the explanation
        public static async Task<string[]> Ask(string key, string request, string context)
        {
            AnthropicClient client = new AnthropicClient { ApiKey = key };
            MessageCreateParams p = new MessageCreateParams
            {
                Model = Model,
                MaxTokens = 4000,
                System = SystemPrompt,
                OutputConfig = new BetaOutputConfig { Effort = Effort.Low },
                // a request the main model declines is answered by a fallback model in the same call
                Betas = new List<Anthropic.Core.ApiEnum<string, Anthropic.Models.Beta.AnthropicBeta>> { "server-side-fallback-2026-06-01" },
                Fallbacks = new List<BetaFallbackParam> { new BetaFallbackParam { Model = "claude-opus-4-8" } },
                Messages = new List<BetaMessageParam>
                {
                    new BetaMessageParam { Role = Role.User, Content = context + "\nRequest: " + request }
                }
            };
            BetaMessage r = await client.Beta.Messages.Create(p).ConfigureAwait(false);
            if (r.StopReason == "refusal") return new[] { null, "The request was declined. Try describing the calculation differently." };
            string text = string.Join("", r.Content.Select(b => b.Value).OfType<BetaTextBlock>().Select(t => t.Text));
            return Parse(text);
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
        readonly TextBox ask = new TextBox(), keyBox = new TextBox(), result = new TextBox(), sent = new TextBox();
        readonly Label why = new Label(), status = new Label();
        readonly CheckBox withSample = new CheckBox(), fill = new CheckBox();
        readonly Button go = new Button(), insert = new Button();
        readonly string context, sample;
        readonly float z = 1f;
        public string Formula;
        public bool FillDown;

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
            ClientSize = new Size(S(760), S(560));

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
                Formula = result.Text.Trim();
                FillDown = fill.Checked;
                if (!Formula.StartsWith("=", StringComparison.Ordinal)) { Say("A formula starts with =.", Red); return; }
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(insert);

            Lbl("What is sent to Claude (Anthropic) with your request:", 16, 296, 728, 20, true);
            sent.Multiline = true;
            sent.ReadOnly = true;
            sent.ScrollBars = ScrollBars.Vertical;
            sent.Font = new Font("Consolas", 8.5f);
            sent.SetBounds(S(16), S(318), S(728), S(120));
            Controls.Add(sent);

            Lbl("Your Anthropic API key (console.anthropic.com). Kept encrypted on this PC for your Windows user:", 16, 448, 728, 20, false);
            keyBox.UseSystemPasswordChar = true;
            keyBox.SetBounds(S(16), S(470), S(600), S(26));
            string k = AiFormula.Key;
            if (k != null) keyBox.Text = k;
            Controls.Add(keyBox);
            Button forget = new Button();
            forget.Text = "Forget key";
            forget.SetBounds(S(624), S(468), S(120), S(30));
            forget.Click += delegate { AiFormula.Key = null; keyBox.Text = ""; Say("The key was removed from this PC.", Slate); };
            Controls.Add(forget);

            status.SetBounds(S(16), S(510), S(600), S(40));
            Controls.Add(status);
            Button close = new Button();
            close.Text = "Close";
            close.DialogResult = DialogResult.Cancel;
            close.SetBounds(S(658), S(516), S(86), S(30));
            Controls.Add(close);
            CancelButton = close;
            ShowSent();
            ActiveControl = ask;
        }

        void ShowSent()
        {
            sent.Text = (withSample.Checked ? sample : context).Replace("\n", "\r\n") + "\r\nRequest: (your words above)";
        }

        void Say(string s, Color c)
        {
            status.ForeColor = c;
            status.Text = s;
        }

        async Task Generate()
        {
            string req = ask.Text.Trim();
            string key = keyBox.Text.Trim();
            if (req.Length < 3) { Say("Describe what the formula should do first.", Red); return; }
            if (key.Length < 20) { Say("Paste your Anthropic API key first (it starts with sk-ant-).", Red); return; }
            AiFormula.Key = key;
            go.Enabled = false;
            insert.Enabled = false;
            Say("Writing the formula ...", Slate);
            try
            {
                string[] r = await AiFormula.Ask(key, req, withSample.Checked ? sample : context);
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
                    Say("Check the formula, change it if you like, then Insert.", Color.FromArgb(21, 128, 61));
                }
            }
            catch (Anthropic.Exceptions.AnthropicUnauthorizedException) { Say("The API key was not accepted. Check it on console.anthropic.com.", Red); }
            catch (Anthropic.Exceptions.AnthropicRateLimitException) { Say("Too many requests just now. Wait a minute and try again.", Red); }
            catch (Exception ex) { Say("Could not reach Claude: " + ex.Message + "  (internet or proxy?)", Red); }
            finally { go.Enabled = true; }
        }
    }
}
