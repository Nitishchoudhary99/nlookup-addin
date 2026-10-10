using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace SmartPro
{
    // WEB TABLE: the selected table as one HTML file: search, sort by any column, a filter under every header,
    //            totals of the rows shown, CSV download, print, dark mode, works on a phone. Nothing is uploaded.
    // TABLE TO CODE: the table as SQL (CREATE TABLE + INSERT, for SQL Server, MySQL, PostgreSQL or SQLite)
    //            or as Python (pandas), with column types worked out from the data.
    internal static class ExportMore
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        static bool ReadTable(string title, out object[,] v, out string[] h, out string name, out dynamic src)
        {
            v = null; h = null; name = null;
            src = Xl.PickRange(title, "Select the table, with its header row:");
            if (src == null) return false;
            v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 2) { Ui.Info("Select a table with a header row and at least one row of data."); return false; }
            int nc = v.GetLength(1);
            h = new string[nc + 1];
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int j = 1; j <= nc; j++)
            {
                string s = Xl.HeaderText(v[1, j], j), t = s;
                int k = 1;
                while (!seen.Add(t)) t = s + "_" + (++k);
                h[j] = t;
            }
            name = (string)src.Worksheet.Name;
            return true;
        }

        static string Json(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length + 2);
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': break;
                    case '\t': sb.Append("\\t"); break;
                    case '<': sb.Append("\\u003c"); break;
                    default:
                        if (c < 32) sb.Append("\\u").Append(((int)c).ToString("x4", Inv));
                        else sb.Append(c);
                        break;
                }
            }
            return sb.Append('"').ToString();
        }

        static string Html(string s) { return System.Security.SecurityElement.Escape(s ?? ""); }

        // ---------------------------------------------------------------
        public static void WebTable()
        {
            object[,] v; string[] h; string sheet; dynamic src;
            if (!ReadTable("Web Table", out v, out h, out sheet, out src)) return;
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr > 200001) { Ui.Info("A web table can hold up to 200,000 rows. Select fewer rows."); return; }
            string title = Ui.AskText("Title of the page:", "Web Table", sheet);
            if (title == null) return;
            // column kinds: n = number, d = date, t = text
            char[] kind = new char[nc + 1];
            for (int j = 1; j <= nc; j++)
            {
                int num = 0, date = 0, other = 0;
                for (int r = 2; r <= nr && r < 2000; r++)
                {
                    object x = v[r, j];
                    if (Xl.IsBlank(x)) continue;
                    if (x is double) num++; else if (x is DateTime) date++; else other++;
                }
                kind[j] = other == 0 && num > 0 && date == 0 ? 'n' : other == 0 && date > 0 && num == 0 ? 'd' : 't';
            }
            StringBuilder data = new StringBuilder();
            data.Append("[");
            bool firstRow = true;
            for (int r = 2; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r)) continue;
                data.Append(firstRow ? "[" : ",[");
                firstRow = false;
                for (int j = 1; j <= nc; j++)
                {
                    if (j > 1) data.Append(',');
                    object x = v[r, j];
                    if (Xl.IsBlank(x)) data.Append("null");
                    else if (x is double) data.Append(((double)x).ToString("R", Inv));
                    else if (x is DateTime) data.Append(Json(((DateTime)x).ToString("yyyy-MM-dd", Inv)));
                    else data.Append(Json(Xl.Str(x)));
                }
                data.Append(']');
            }
            data.Append("]");
            StringBuilder cols = new StringBuilder("[");
            for (int j = 1; j <= nc; j++) cols.Append(j > 1 ? "," : "").Append("{\"n\":").Append(Json(h[j])).Append(",\"k\":\"").Append(kind[j]).Append("\"}");
            cols.Append("]");

            string path;
            using (System.Windows.Forms.SaveFileDialog d = new System.Windows.Forms.SaveFileDialog())
            {
                d.Filter = "Web page (*.html)|*.html";
                d.FileName = Xl.SafeFileName(title) + ".html";
                if (d.ShowDialog(new ExcelWin()) != System.Windows.Forms.DialogResult.OK) return;
                path = d.FileName;
            }
            string page = WebTablePage.Text
                .Replace("%TITLE%", Html(title))
                .Replace("%SUB%", Html(sheet + "!" + ((string)src.Address).Replace("$", "") + "  ·  " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm", Inv) + "  ·  made with " + Design.AppName))
                .Replace("%COLS%", cols.ToString())
                .Replace("%DATA%", data.ToString());
            File.WriteAllText(path, page, new UTF8Encoding(false));
            try { System.Diagnostics.Process.Start(path); } catch { }
            Xl.Status("Web table saved: " + path + "  (one file; email it or put it on a shared drive).");
        }

        // ---------------------------------------------------------------
        public static void TableToCode()
        {
            object[,] v; string[] h; string sheet; dynamic src;
            if (!ReadTable("Table to SQL / Python", out v, out h, out sheet, out src)) return;
            string[] targets = { "SQL Server (T-SQL)", "MySQL / MariaDB", "PostgreSQL", "SQLite", "Python (pandas DataFrame)" };
            int t = Ui.Pick("Table to SQL / Python", "Which language?", targets, null);
            if (t < 0) return;
            string table = Ui.AskText("Name of the table / DataFrame:", "Table to SQL / Python", Ident(sheet));
            if (table == null) return;
            table = Ident(table);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            // column types from the data
            string[] type = new string[nc + 1];
            int[] maxLen = new int[nc + 1];
            for (int j = 1; j <= nc; j++)
            {
                bool allInt = true, anyNum = false, anyDate = false, anyText = false, anyBool = false;
                for (int r = 2; r <= nr; r++)
                {
                    object x = v[r, j];
                    if (Xl.IsBlank(x)) continue;
                    if (x is double) { anyNum = true; double d = (double)x; if (d != Math.Floor(d) || Math.Abs(d) > 2e9) allInt = false; }
                    else if (x is DateTime) anyDate = true;
                    else if (x is bool) anyBool = true;
                    else { anyText = true; maxLen[j] = Math.Max(maxLen[j], Xl.Str(x).Length); }
                }
                type[j] = anyText || (anyNum && anyDate) ? "text" : anyDate ? "date" : anyBool && !anyNum ? "bool" : anyNum ? (allInt ? "int" : "dec") : "text";
                if (type[j] == "text") maxLen[j] = Math.Max(maxLen[j], 1);
            }
            string code = t == 4 ? Python(v, h, type, table) : Sql(v, h, type, maxLen, table, t);
            string ext = t == 4 ? ".py" : ".sql";
            string path = Path.Combine(Path.GetTempPath(), Xl.SafeFileName(table) + ext);
            File.WriteAllText(path, code, new UTF8Encoding(false));
            try { System.Windows.Forms.Clipboard.SetText(code.Length > 5000000 ? code.Substring(0, 5000000) : code); } catch { }
            try { System.Diagnostics.Process.Start("notepad.exe", "\"" + path + "\""); } catch { }
            Xl.Status(DataRows(v).ToString("#,##0", Inv) + " rows written as " + targets[t] + ". Copied to the clipboard and opened in Notepad: " + path);
        }

        static int DataRows(object[,] v)
        {
            int n = 0;
            for (int r = 2; r <= v.GetLength(0); r++) if (!Xl.RowBlank(v, r)) n++;
            return n;
        }

        static string Ident(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in s.Trim()) sb.Append(char.IsLetterOrDigit(c) ? char.ToLowerInvariant(c) : '_');
            string r = sb.ToString().Trim('_');
            while (r.Contains("__")) r = r.Replace("__", "_");
            if (r.Length == 0) r = "data";
            if (char.IsDigit(r[0])) r = "t_" + r;
            return r;
        }

        static string Sql(object[,] v, string[] h, string[] type, int[] maxLen, string table, int dialect)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            Func<string, string> q = delegate (string n)
            {
                switch (dialect)
                {
                    case 0: return "[" + n.Replace("]", "]]") + "]";
                    case 1: return "`" + n.Replace("`", "``") + "`";
                    default: return "\"" + n.Replace("\"", "\"\"") + "\"";
                }
            };
            StringBuilder sb = new StringBuilder();
            sb.Append("-- ").Append(DataRows(v)).Append(" rows from Excel, made with ").Append(Design.AppName).Append("\n");
            sb.Append("CREATE TABLE ").Append(q(table)).Append(" (\n");
            for (int j = 1; j <= nc; j++)
            {
                string ty;
                switch (type[j])
                {
                    case "int": ty = dialect == 3 ? "INTEGER" : "BIGINT"; break;
                    case "dec": ty = dialect == 3 ? "REAL" : "DECIMAL(18,4)"; break;
                    case "date": ty = dialect == 0 ? "DATETIME2" : dialect == 1 ? "DATETIME" : dialect == 2 ? "TIMESTAMP" : "TEXT"; break;
                    case "bool": ty = dialect == 0 ? "BIT" : dialect == 3 ? "INTEGER" : "BOOLEAN"; break;
                    default:
                        int n = maxLen[j] <= 50 ? 50 : maxLen[j] <= 255 ? 255 : maxLen[j] <= 4000 ? 4000 : -1;
                        ty = dialect == 3 ? "TEXT" : n < 0 ? (dialect == 0 ? "NVARCHAR(MAX)" : "TEXT") : (dialect == 0 ? "NVARCHAR(" + n + ")" : "VARCHAR(" + n + ")");
                        break;
                }
                sb.Append("    ").Append(q(h[j])).Append(' ').Append(ty).Append(j < nc ? ",\n" : "\n");
            }
            sb.Append(");\n\n");
            const int batch = 500;
            int inBatch = 0;
            for (int r = 2; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r)) continue;
                if (inBatch == 0)
                {
                    sb.Append("INSERT INTO ").Append(q(table)).Append(" (");
                    for (int j = 1; j <= nc; j++) sb.Append(j > 1 ? ", " : "").Append(q(h[j]));
                    sb.Append(") VALUES\n");
                }
                else sb.Append(",\n");
                sb.Append("(");
                for (int j = 1; j <= nc; j++)
                {
                    if (j > 1) sb.Append(", ");
                    object x = v[r, j];
                    if (Xl.IsBlank(x) || Xl.IsErr(x)) sb.Append("NULL");
                    else if (type[j] == "int" || type[j] == "dec") sb.Append(x is double ? ((double)x).ToString("R", Inv) : "NULL");
                    else if (type[j] == "bool") sb.Append(x is bool ? ((bool)x ? (dialect == 2 ? "TRUE" : "1") : (dialect == 2 ? "FALSE" : "0")) : "NULL");
                    else if (type[j] == "date" && x is DateTime) sb.Append('\'').Append(((DateTime)x).ToString("yyyy-MM-dd HH:mm:ss", Inv)).Append('\'');
                    else sb.Append(dialect == 0 ? "N'" : "'").Append(Xl.Str(x).Replace("'", "''")).Append('\'');
                }
                sb.Append(")");
                if (++inBatch == batch) { sb.Append(";\n\n"); inBatch = 0; }
            }
            if (inBatch > 0) sb.Append(";\n");
            return sb.ToString();
        }

        static string Py(object x, string type)
        {
            if (Xl.IsBlank(x) || Xl.IsErr(x)) return "None";
            if (x is double) return ((double)x).ToString("R", Inv);
            if (x is bool) return (bool)x ? "True" : "False";
            if (x is DateTime) return "\"" + ((DateTime)x).ToString("yyyy-MM-dd HH:mm:ss", Inv) + "\"";
            return Json(Xl.Str(x));
        }

        static string Python(object[,] v, string[] h, string[] type, string table)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            StringBuilder sb = new StringBuilder();
            sb.Append("# ").Append(DataRows(v)).Append(" rows from Excel, made with ").Append(Design.AppName).Append("\nimport pandas as pd\n\n");
            sb.Append(table).Append(" = pd.DataFrame({\n");
            for (int j = 1; j <= nc; j++)
            {
                sb.Append("    ").Append(Json(h[j])).Append(": [");
                bool first = true;
                for (int r = 2; r <= nr; r++)
                {
                    if (Xl.RowBlank(v, r)) continue;
                    if (!first) sb.Append(", ");
                    first = false;
                    sb.Append(Py(v[r, j], type[j]));
                }
                sb.Append("],\n");
            }
            sb.Append("})\n");
            for (int j = 1; j <= nc; j++)
            {
                if (type[j] == "date") sb.Append(table).Append("[").Append(Json(h[j])).Append("] = pd.to_datetime(").Append(table).Append("[").Append(Json(h[j])).Append("])\n");
                if (type[j] == "int") sb.Append(table).Append("[").Append(Json(h[j])).Append("] = ").Append(table).Append("[").Append(Json(h[j])).Append("].astype(\"Int64\")\n");
            }
            sb.Append("\nprint(").Append(table).Append(".head())\nprint(").Append(table).Append(".describe(include=\"all\"))\n");
            return sb.ToString();
        }
    }

    // the page of a Web Table; %TITLE%, %SUB%, %COLS% and %DATA% are filled in
    internal static class WebTablePage
    {
        public const string Text = @"<!doctype html><html lang=""en""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1"">
<title>%TITLE%</title><style>
:root{--bg:#f8fafc;--card:#fff;--ink:#0f172a;--mute:#64748b;--line:#e2e8f0;--head:#0f172a;--headInk:#fff;--acc:#2563eb;--hit:#fef9c3}
@media (prefers-color-scheme:dark){:root:not([data-theme=light]){--bg:#0b1220;--card:#111a2e;--ink:#e2e8f0;--mute:#94a3b8;--line:#1e293b;--head:#1e293b;--headInk:#e2e8f0;--acc:#60a5fa;--hit:#3f3a10}}
:root[data-theme=dark]{--bg:#0b1220;--card:#111a2e;--ink:#e2e8f0;--mute:#94a3b8;--line:#1e293b;--head:#1e293b;--headInk:#e2e8f0;--acc:#60a5fa;--hit:#3f3a10}
*{box-sizing:border-box}body{margin:0;font:14px/1.4 'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--ink)}
header{padding:18px 16px 8px}h1{margin:0;font-size:20px}.sub{color:var(--mute);font-size:12px;margin-top:4px}
.bar{display:flex;flex-wrap:wrap;gap:8px;padding:8px 16px;align-items:center}
input[type=search]{flex:1;min-width:180px;padding:9px 12px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit}
button{padding:8px 12px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);cursor:pointer;font:inherit}
button.main{background:var(--acc);border-color:var(--acc);color:#fff}.count{color:var(--mute);font-size:13px;margin-left:auto}
.wrap{margin:0 16px 16px;overflow:auto;max-height:calc(100vh - 150px);border:1px solid var(--line);border-radius:10px;background:var(--card)}
table{border-collapse:collapse;width:100%}th,td{padding:7px 10px;border-bottom:1px solid var(--line);white-space:nowrap;text-align:left}
thead th{position:sticky;top:0;background:var(--head);color:var(--headInk);cursor:pointer;user-select:none;z-index:2}
thead tr.f th{top:35px;background:var(--card);padding:4px 6px;cursor:default}thead tr.f input{width:100%;min-width:60px;padding:4px 6px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--ink);font:12px inherit}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}tbody tr:hover{background:var(--hit)}tfoot td{font-weight:600;background:var(--bg);position:sticky;bottom:0}
.more{padding:10px;text-align:center;color:var(--mute)}@media print{.bar,thead tr.f{display:none}.wrap{max-height:none;overflow:visible;border:0}}
</style></head><body><header><h1>%TITLE%</h1><div class=""sub"">%SUB%</div></header>
<div class=""bar""><input type=""search"" id=""q"" placeholder=""Search every column ..."" aria-label=""Search""><button id=""csv"" class=""main"">Download CSV</button><button onclick=""print()"">Print / PDF</button><button id=""th"">Theme</button><span class=""count"" id=""cnt""></span></div>
<div class=""wrap""><table><thead><tr id=""hd""></tr><tr class=""f"" id=""fl""></tr></thead><tbody id=""bd""></tbody><tfoot><tr id=""ft""></tr></tfoot></table><div class=""more"" id=""more""></div></div>
<script>
const C=%COLS%,D=%DATA%;let rows=D.slice(),sortC=-1,sortD=1,limit=500;const filt=C.map(()=>'');
C.forEach((c,i)=>{c.dp=c.k==='n'&&D.some(r=>r[i]!=null&&r[i]%1!==0)?2:0});
const fmt=(v,k,i)=>v==null?'':k==='n'?Number(v).toLocaleString('en-IN',{minimumFractionDigits:C[i].dp,maximumFractionDigits:C[i].dp}):String(v);
const esc=s=>s.replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const hd=document.getElementById('hd'),fl=document.getElementById('fl');
C.forEach((c,i)=>{const th=document.createElement('th');th.textContent=c.n;if(c.k==='n')th.className='n';th.onclick=()=>{sortD=sortC===i?-sortD:1;sortC=i;run()};hd.appendChild(th);
const f=document.createElement('th'),inp=document.createElement('input');inp.placeholder='filter';inp.oninput=()=>{filt[i]=inp.value.toLowerCase();run()};f.appendChild(inp);fl.appendChild(f)});
function run(){const q=document.getElementById('q').value.toLowerCase();
rows=D.filter(r=>(!q||r.some(v=>v!=null&&String(v).toLowerCase().includes(q)))&&filt.every((f,i)=>{if(!f)return true;const v=r[i];if(v==null)return false;
if(C[i].k==='n'&&/^[<>]=?\s*-?[\d.]+$/.test(f)){const op=f.match(/^[<>]=?/)[0],n=parseFloat(f.slice(op.length));return op==='>'?v>n:op==='<'?v<n:op==='>='?v>=n:v<=n}return String(v).toLowerCase().includes(f)}));
if(sortC>=0)rows.sort((a,b)=>{const x=a[sortC],y=b[sortC];if(x==null)return 1;if(y==null)return -1;return(x>y?1:x<y?-1:0)*sortD});
draw()}
function draw(){const bd=document.getElementById('bd');bd.innerHTML=rows.slice(0,limit).map(r=>'<tr>'+r.map((v,i)=>'<td class=""'+(C[i].k==='n'?'n':'')+'"">'+esc(fmt(v,C[i].k,i))+'</td>').join('')+'</tr>').join('');
document.getElementById('cnt').textContent=rows.length.toLocaleString('en-IN')+' of '+D.length.toLocaleString('en-IN')+' rows';
const m=document.getElementById('more');m.innerHTML=rows.length>limit?'<button onclick=""limit+=1000;draw()"">Show more rows</button>':'';
document.getElementById('ft').innerHTML=C.map((c,i)=>'<td class=""'+(c.k==='n'?'n':'')+'"">'+(c.k==='n'?fmt(rows.reduce((s,r)=>s+(r[i]||0),0),'n',i):i===0?'Total':'')+'</td>').join('');
[...hd.children].forEach((th,i)=>th.textContent=C[i].n+(i===sortC?(sortD>0?' ▲':' ▼'):''))}
document.getElementById('q').oninput=run;
document.getElementById('csv').onclick=()=>{const q=v=>v==null?'':/[,""\n]/.test(String(v))?'""'+String(v).replace(/""/g,'""""')+'""':String(v);
const t=[C.map(c=>q(c.n)).join(',')].concat(rows.map(r=>r.map(q).join(','))).join('\r\n');const a=document.createElement('a');
a.href=URL.createObjectURL(new Blob(['﻿'+t],{type:'text/csv'}));a.download=document.title+'.csv';a.click()};
document.getElementById('th').onclick=()=>{const r=document.documentElement,d=r.dataset.theme==='dark'||(!r.dataset.theme&&matchMedia('(prefers-color-scheme:dark)').matches);r.dataset.theme=d?'light':'dark'};
run();
</script></body></html>";
    }
}
