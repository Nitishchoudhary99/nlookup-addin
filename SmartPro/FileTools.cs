using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace SmartPro
{
    // FILE ORGANIZER: the files of a folder as a table you can edit, then rename them all from that table,
    // or sort them into sub-folders by type, by month or by the start of their name.
    // Every move and rename is written to a log file in the folder, so nothing is lost and it can be undone by hand.
    internal static class FileTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const string ListTitle = "Files in: ";
        const string LogName = "_DocBrisk_file_log.csv";

        public static void Run()
        {
            string[] items =
            {
                "1.  List the files of a folder in a new sheet (name, type, size, date, link) - then type new names",
                "2.  Rename the files from that list (column 'New name')",
                "3.  Sort the files of a folder into sub-folders (by type, by month or by the start of the name)"
            };
            int pick = Ui.Pick(Design.AppName + " - File Organizer", "What do you want to do?", items, null);
            if (pick == 0) ListFiles();
            else if (pick == 1) RenameFromList();
            else if (pick == 2) SortIntoFolders();
        }

        static string Kind(string ext)
        {
            switch (ext.ToLowerInvariant())
            {
                case ".xlsx": case ".xls": case ".xlsm": case ".xlsb": case ".csv": return "Excel";
                case ".pdf": return "PDF";
                case ".doc": case ".docx": case ".rtf": case ".txt": return "Documents";
                case ".ppt": case ".pptx": return "Presentations";
                case ".jpg": case ".jpeg": case ".png": case ".gif": case ".bmp": case ".webp": case ".heic": return "Images";
                case ".zip": case ".rar": case ".7z": return "Archives";
                case ".mp4": case ".mov": case ".avi": case ".mkv": return "Videos";
                case ".xml": case ".json": return "Data";
                default: return "Other";
            }
        }

        // ---------------------------------------------------------------
        static void ListFiles()
        {
            string folder = Ui.PickFolder("Choose the folder to list");
            if (folder == null) return;
            bool sub = Ui.Ok("Include the files in sub-folders too?\n\nOK = yes   Cancel = only this folder");
            List<FileInfo> files = new List<FileInfo>();
            try
            {
                foreach (string f in Directory.GetFiles(folder, "*", sub ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly))
                {
                    FileInfo fi = new FileInfo(f);
                    if ((fi.Attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0) continue;
                    if (fi.Name.StartsWith("~$", StringComparison.Ordinal) || fi.Name == LogName) continue;
                    files.Add(fi);
                    if (files.Count > 200000) break;
                }
            }
            catch (Exception ex) { Ui.Warn("The folder could not be read completely:\n" + ex.Message); }
            if (files.Count == 0) { Ui.Info("No files were found in:\n" + folder); return; }
            files.Sort(delegate (FileInfo a, FileInfo b) { return string.Compare(a.FullName, b.FullName, StringComparison.OrdinalIgnoreCase); });
            object[,] o = Xl.New2D(files.Count + 1, 8);
            string[] hd = { "Folder", "File name", "Type", "Extension", "Size (KB)", "Modified", "Open", "New name" };
            for (int j = 0; j < hd.Length; j++) o[1, j + 1] = hd[j];
            for (int i = 0; i < files.Count; i++)
            {
                FileInfo fi = files[i];
                o[i + 2, 1] = fi.DirectoryName;
                o[i + 2, 2] = Path.GetFileNameWithoutExtension(fi.Name);
                o[i + 2, 3] = Kind(fi.Extension);
                o[i + 2, 4] = fi.Extension;
                o[i + 2, 5] = Math.Round(fi.Length / 1024.0, 1);
                o[i + 2, 6] = fi.LastWriteTime;
            }
            dynamic wb = Xl.App.ActiveWorkbook ?? Xl.App.Workbooks.Add();
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(wb, "Files");
                ws.Range["A1"].Value2 = ListTitle + folder;
                ws.Range["A1"].Font.Bold = true;
                ws.Range["A2"].Value2 = "Type new names (without the extension) in column H, then File Organizer > Rename the files from that list.";
                ws.Range["A2"].Font.Color = Design.Slate;
                Xl.WriteTable(ws.Range["A4"], o, 1);
                Xl.StyleHeader(ws.Range["A4:H4"]);
                ws.Range["F5"].Resize[files.Count, 1].NumberFormat = "dd-mmm-yyyy hh:mm";
                // a HYPERLINK formula per row opens the file; one write for the whole column
                ws.Range["G5"].Resize[files.Count, 1].Formula = "=HYPERLINK(A5&\"\\\"&B5&D5,\"open\")";
                ws.Range["H5"].Resize[files.Count, 1].Interior.Color = Design.Changed;
                ws.Range["A4"].Resize[files.Count + 1, 8].AutoFilter();
                ws.Columns["A:H"].AutoFit();
                ws.Columns["H"].ColumnWidth = 40;
                ws.Activate();
            }
            Xl.Status(files.Count.ToString("#,##0", Inv) + " files listed. Type new names in the yellow column H.");
        }

        // ---------------------------------------------------------------
        static void RenameFromList()
        {
            dynamic ws = Xl.App.ActiveSheet;
            string a1 = Convert.ToString((object)ws.Range["A1"].Value2, Inv) ?? "";
            if (!a1.StartsWith(ListTitle, StringComparison.Ordinal))
            {
                Ui.Info("Open the sheet made by 'List the files of a folder' first, and type the new names in column H.");
                return;
            }
            int last = (int)ws.Cells[ws.Rows.Count, 2].End[-4162].Row;      // xlUp
            if (last < 5) return;
            object[,] v = Xl.ReadTyped(ws.Range["A5"].Resize[last - 4, 8]);
            List<string[]> plan = new List<string[]>();
            HashSet<string> targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int clash = 0;
            for (int r = 1; r <= v.GetLength(0); r++)
            {
                string nn = Xl.Str(v[r, 8]).Trim();
                if (nn.Length == 0) continue;
                string dir = Xl.Str(v[r, 1]), name = Xl.Str(v[r, 2]), ext = Xl.Str(v[r, 4]);
                if (!nn.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) nn += ext;
                nn = Xl.SafeFileName(nn);
                string from = Path.Combine(dir, name + ext), to = Path.Combine(dir, nn);
                if (string.Equals(from, to, StringComparison.Ordinal)) continue;
                if (!targets.Add(to) || (File.Exists(to) && !string.Equals(from, to, StringComparison.OrdinalIgnoreCase))) { clash++; continue; }
                plan.Add(new[] { from, to, (r + 4).ToString(Inv) });
            }
            if (plan.Count == 0) { Ui.Info("No new names were found in column H" + (clash > 0 ? " (" + clash + " were skipped because the name is already taken)." : ".")); return; }
            if (!Ui.Ok("Rename " + plan.Count + " files?" + (clash > 0 ? "\n" + clash + " are skipped because the new name is already taken." : "") +
                       "\n\nEvery rename is written to " + LogName + " in its folder.")) return;
            int ok = 0, fail = 0;
            using (new Speed())
            {
                foreach (string[] p in plan)
                {
                    int row = int.Parse(p[2], Inv);
                    try
                    {
                        File.Move(p[0], p[1]);
                        Log(Path.GetDirectoryName(p[0]), "rename", p[0], p[1]);
                        ws.Cells[row, 2].Value2 = Path.GetFileNameWithoutExtension(p[1]);
                        ws.Cells[row, 8].Value2 = "";
                        ws.Cells[row, 8].Interior.Color = Design.Good;
                        ok++;
                    }
                    catch (Exception ex)
                    {
                        ws.Cells[row, 8].Interior.Color = Design.Bad;
                        try { ws.Cells[row, 8].AddComment(ex.Message); } catch { }
                        fail++;
                    }
                }
            }
            Ui.Info(ok + " files renamed." + (fail > 0 ? "\n" + fail + " could not be renamed (red, the reason is in the cell note); a file open in another program cannot be renamed." : ""));
        }

        // ---------------------------------------------------------------
        static void SortIntoFolders()
        {
            string folder = Ui.PickFolder("Choose the folder to tidy up");
            if (folder == null) return;
            string[] ways = { "By type: Excel, PDF, Images, Documents ...", "By month of the last change: 2026-04, 2026-05 ...", "By the first word of the file name (party, project ...)" };
            int how = Ui.Pick("Sort files into folders", "How should the files be grouped?", ways, null);
            if (how < 0) return;
            string[] files;
            try { files = Directory.GetFiles(folder, "*", SearchOption.TopDirectoryOnly); }
            catch (Exception ex) { Ui.Warn(ex.Message); return; }
            Dictionary<string, List<string>> groups = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (string f in files)
            {
                FileInfo fi = new FileInfo(f);
                if ((fi.Attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0 || fi.Name.StartsWith("~$", StringComparison.Ordinal) || fi.Name == LogName) continue;
                string g = how == 0 ? Kind(fi.Extension) : how == 1 ? fi.LastWriteTime.ToString("yyyy-MM", Inv) : FirstWord(fi.Name);
                List<string> l;
                if (!groups.TryGetValue(g, out l)) { l = new List<string>(); groups[g] = l; }
                l.Add(f);
            }
            if (groups.Count == 0) { Ui.Info("There are no files directly in that folder."); return; }
            StringBuilder preview = new StringBuilder();
            int shown = 0, total = 0;
            List<string> keys = new List<string>(groups.Keys);
            keys.Sort(StringComparer.OrdinalIgnoreCase);
            foreach (string k in keys)
            {
                total += groups[k].Count;
                if (shown++ < 12) preview.Append("   ").Append(k).Append(": ").Append(groups[k].Count).Append("\n");
            }
            if (keys.Count > 12) preview.Append("   ... and ").Append(keys.Count - 12).Append(" more folders\n");
            if (!Ui.Ok("Move " + total + " files into " + keys.Count + " folders inside\n" + folder + "\n\n" + preview +
                       "\nEvery move is written to " + LogName + ", so it can be undone.")) return;
            int moved = 0, fail = 0;
            foreach (string k in keys)
            {
                string dir = Path.Combine(folder, Xl.SafeFileName(k));
                try { Directory.CreateDirectory(dir); } catch { fail += groups[k].Count; continue; }
                foreach (string f in groups[k])
                {
                    string to = Path.Combine(dir, Path.GetFileName(f));
                    int n = 1;
                    while (File.Exists(to)) to = Path.Combine(dir, Path.GetFileNameWithoutExtension(f) + " (" + (++n) + ")" + Path.GetExtension(f));
                    try
                    {
                        File.Move(f, to);
                        Log(folder, "move", f, to);
                        moved++;
                    }
                    catch { fail++; }
                }
            }
            Ui.Info(moved + " files moved into " + keys.Count + " folders." + (fail > 0 ? "\n" + fail + " could not be moved (open in another program?)." : "") +
                    "\nThe log is in " + Path.Combine(folder, LogName));
        }

        static string FirstWord(string name)
        {
            string s = Path.GetFileNameWithoutExtension(name).Trim();
            int i = 0;
            while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '&')) i++;
            string w = s.Substring(0, i);
            return w.Length == 0 ? "Other" : char.ToUpperInvariant(w[0]) + w.Substring(1);
        }

        static void Log(string folder, string what, string from, string to)
        {
            try
            {
                string path = Path.Combine(folder, LogName);
                bool head = !File.Exists(path);
                using (StreamWriter w = new StreamWriter(path, true, Encoding.UTF8))
                {
                    if (head) w.WriteLine("When,Action,From,To");
                    w.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", Inv) + "," + what + ",\"" + from.Replace("\"", "\"\"") + "\",\"" + to.Replace("\"", "\"\"") + "\"");
                }
            }
            catch { }
        }
    }
}
