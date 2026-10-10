using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.IntelliSense;

namespace SmartPro
{
    // Runs when Excel loads / unloads the add-in
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            try { IntelliSenseServer.Install(); } catch { }
            try
            {
                ExcelAsyncUtil.QueueAsMacro(delegate
                {
                    try { Xl.App.OnKey("^+q", "SmartProQuickFind"); } catch { }
                    try { Xl.App.OnKey("^+j", "SmartProNavigator"); } catch { }
                    try { FocusCell.Restore(); } catch { }
                    // a returning user sees the streak and the time saved on the status bar (nothing to click away)
                    try
                    {
                        string hello = Habit.Greeting();
                        if (hello != null) Xl.Status(hello);
                    }
                    catch { }
                });
            }
            catch { }
        }

        public void AutoClose()
        {
            try { FocusCell.Stop(); } catch { }
            try { IntelliSenseServer.Uninstall(); } catch { }
        }
    }

    // The ribbon tab. The XML is generated from Design.Tools
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return Design.RibbonXml();
        }

        public void OnTool(IRibbonControl control)
        {
            Runner.Run(control.Id);
        }
    }

    public static class Commands
    {
        [ExcelCommand(Name = "SmartProQuickFind")]
        public static void QuickFind()
        {
            Runner.Run("find");
        }

        [ExcelCommand(Name = "SmartProNavigator")]
        public static void Navigator()
        {
            Runner.Run("navigator");
        }
    }

    internal static class Runner
    {
        static string lastTool;      // for Repeat Last Tool

        public static void Run(string id)
        {
            if (id == "repeat")
            {
                if (lastTool == null)
                {
                    Ui.Info("No tool has been used yet since Excel was opened.");
                    return;
                }
                id = lastTool;
            }
            else if (id != "assist" && id != "find" && id != "about" && id != "licence" && id != "funcs" && id != "impact") lastTool = id;
            try
            {
                bool needsBook = !(id == "about" || id == "openbackups" || id == "find" || id == "combfolder" || id == "fileorg" || id == "focus" || id == "funcs" || id == "licence" || id == "assist" || id == "repeat" || id == "impact");
                object book = Xl.App.ActiveWorkbook;
                if (needsBook && book == null)
                {
                    Ui.Info("Open a workbook first.");
                    return;
                }
                // Pro tools: a valid Pro key, or one of the free trial runs
                if (Design.IsPro(id) && !Licence.AllowPro(id, Design.CaptionOf(id))) return;
                if (Habit.IsWork(id)) Settings.Set("Runs", (Settings.GetInt("Runs", 0) + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                switch (id)
                {
                    case "assist": Assistant.Show(); break;
                    case "explain": SmartTools.ExplainFormula(); break;
                    case "smartlookup": SmartLookup.Run(); break;
                    case "impact": ImpactWindow.Show(); break;
                    case "analyst": SmartTools.SmartAnalyst(); break;
                    case "duppay": SmartTools.DuplicatePayments(); break;
                    case "mailparty": SmartTools.MailToEachParty(); break;
                    case "dash": Dashboard.Build(); break;
                    case "refresh": Dashboard.Refresh(); break;
                    case "emailsheet": WorkbookTools.EmailSheet(); break;
                    case "custom": DashboardStudio.Run(); break;
                    case "editdash": Dashboard.Edit(); break;
                    case "sumfinder": DailyTools.SumFinder(); break;
                    case "reconcile": DailyTools.Reconcile(); break;
                    case "report": DailyTools.ReportToTable(); break;
                    case "ask": AskTools.Ask(); break;
                    case "variance": InsightTools.Variance(); break;
                    case "forecast": InsightTools.Forecast(); break;
                    case "style": Dashboard.ChooseStyle(); break;
                    case "webdash": ExportTools.WebDashboard(); break;
                    case "biexport": ExportTools.BiExport(); break;
                    case "neat": WorkbookTools.MakeNeat(); break;
                    case "summary": WorkbookTools.QuickSummary(); break;
                    case "listcompare": WorkbookTools.CompareLists(); break;
                    case "colourcount": WorkbookTools.ColourCount(); break;
                    case "valuescopy": WorkbookTools.ValuesCopy(); break;
                    case "unhideall": WorkbookTools.UnhideAll(); break;
                    case "protectall": WorkbookTools.ProtectAll(true); break;
                    case "unprotectall": WorkbookTools.ProtectAll(false); break;
                    case "links": WorkbookTools.Links(); break;
                    case "diet": WorkbookTools.FileDiet(); break;
                    case "sortsheets": WorkbookTools.SortSheets(); break;
                    case "sheetsfromlist": WorkbookTools.SheetsFromList(); break;
                    case "pastevisible": EverydayTools.PasteToVisible(); break;
                    case "delrows": EverydayTools.DeleteRowsWhere(); break;
                    case "blankchange": EverydayTools.BlankRowsAtChange(); break;
                    case "mergesame": EverydayTools.MergeSameValues(); break;
                    case "joincols": EverydayTools.JoinColumns(); break;
                    case "texttools": EverydayTools.TextTools(); break;
                    case "mathcells": EverydayTools.MathOnCells(); break;
                    case "iferror": EverydayTools.WrapIfError(); break;
                    case "dollar": EverydayTools.ToggleDollar(); break;
                    case "health": Dashboard.HealthScan(); break;
                    case "abc": ProTools.AbcAnalysis(); break;
                    case "ageing": ProTools.AgeingReport(); break;
                    case "audit": ProTools.FormulaRiskScan(); break;
                    case "snap": ProTools.TakeSnapshot(); break;
                    case "diff": ProTools.WhatChanged(); break;
                    case "backup": BookTools.BackupNow(); break;
                    case "clean": CleanTools.Clean(true, true, true, false, "Smart Clean"); break;
                    case "tonum": CleanTools.Clean(false, true, false, true, "Text to Numbers"); break;
                    case "todate": CleanTools.Clean(false, false, true, true, "Text to Dates"); break;
                    case "trim": CleanTools.Clean(true, false, false, true, "Trim Spaces"); break;
                    case "blankrows": CleanTools.RemoveBlankRows(); break;
                    case "filldown": CleanTools.FillBlanksDown(); break;
                    case "unmerge": CleanTools.UnmergeAndFill(); break;
                    case "tovalues": CleanTools.FormulasToValues(); break;
                    case "compare": AnalyseTools.CompareByKey(); break;
                    case "fuzzy": AnalyseTools.FuzzyDuplicates(); break;
                    case "fuzzymerge": FuzzyMerge.Run(); break;
                    case "outliers": AnalyseTools.FindOutliers(); break;
                    case "dups": AnalyseTools.HighlightDuplicates(); break;
                    case "unpivot": AnalyseTools.Unpivot(); break;
                    case "splitsheets": BookTools.SplitByColumn(false); break;
                    case "splitfiles": BookTools.SplitByColumn(true); break;
                    case "combsheets": BookTools.CombineSheets(); break;
                    case "combfolder": BookTools.CombineFolder(); break;
                    case "index": NavTools.WorkbookIndex(); break;
                    case "navigator": NavTools.Navigator(); break;
                    case "focus": FocusCell.Toggle(); break;
                    case "sheetmgr": NavTools.SheetManager(); break;
                    case "swap": NavTools.SwapRanges(); break;
                    case "formulas": FormulaLibrary.Run(); break;
                    case "aiformula": AiFormula.Run(); break;
                    case "validation": ValidationTools.Run(); break;
                    case "fileorg": FileTools.Run(); break;
                    case "splitpdfs": BookTools.SplitByColumn(2); break;
                    case "webtable": ExportMore.WebTable(); break;
                    case "tablecode": ExportMore.TableToCode(); break;
                    case "pdfs": BookTools.SheetsToPdf(); break;
                    case "inrfmt": BookTools.IndianFormat(); break;
                    case "fycol": BookTools.AddFiscalYear(); break;
                    case "words": BookTools.RupeesInWords(); break;
                    case "gstin": BookTools.CheckGstin(); break;
                    case "find": QuickFind(); break;
                    case "openbackups": BookTools.OpenBackups(); break;
                    case "funcs": BookTools.FunctionList(); break;
                    case "licence": Licence.Activate(); break;
                    case "about": About(); break;
                    default: Ui.Warn("Unknown tool: " + id); break;
                }
                // after a free trial run of a Pro tool: how many are left, on the status bar (nothing to click away)
                string trial = Licence.TakeNote();
                if (trial != null) Xl.Status(trial);
                // the time saved, the streak and the milestones; then at most one small card in the corner (Nudge.cs)
                System.Collections.Generic.List<Habit.Milestone> reached = null;
                try { reached = Habit.Record(id); } catch { }
                Nudge.AfterRun(id, reached);
            }
            catch (Exception ex)
            {
                Speed.Restore();
                // keep the full details in a file, so the cause can be found from one screenshot or one file
                string log = null;
                try
                {
                    log = System.IO.Path.Combine(Xl.BackupFolder(), "last-error.txt");
                    System.IO.File.WriteAllText(log, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\nDocBrisk " + Design.Version + "\r\nTool: " + id + "\r\n\r\n" + ex.ToString());
                }
                catch { log = null; }
                Ui.Warn("This tool stopped (" + id + "):\n" + ex.Message + "\n\nTools that change your data save a backup first (Help > Open Backups)." +
                        (log != null ? "\n\nThe full error was saved here. Send this file if you need help:\n" + log : ""));
            }
        }

        // Command palette: type a few letters, press Enter
        static void QuickFind()
        {
            ToolDef[] all = Design.Tools;
            string[] items = new string[all.Length];
            string[] keys = new string[all.Length];
            for (int i = 0; i < all.Length; i++)
            {
                items[i] = all[i].Caption + "   -   " + all[i].Tip;
                keys[i] = all[i].Caption + " " + all[i].Group + " " + all[i].Keys;
            }
            int pick = Ui.Pick(Design.AppName + " - Quick Find", "Type what you want to do (split, compare, gst, dashboard ...)", items, keys);
            if (pick < 0) return;
            if (all[pick].Id == "find") return;
            Run(all[pick].Id);
        }

        static void About()
        {
            Ui.Info(Design.AppName + " for Excel  version " + Design.Version + "\nby " + Design.Author + "\ndocbrisk.com\n\n" +
                    Licence.PlanText() + "\n" + ProPlan.Short + "\n\n" +
                    "Free: Smart Lookup, dashboards you design yourself and edit later, web dashboards, Explain Formula, report cleaning, data health, " +
                    "everyday tools, split / combine and an India pack.\n" +
                    "Pro: Smart Analyst, Why It Changed, Forecast, Ask Your Data, fuzzy merge, reconcile, sum finder, duplicate payments, " +
                    "mail to each party, formula risk scan, snapshots, compare, ABC and ageing reports.\n\n" +
                    "Saved so far: about " + Habit.Hours(Habit.MinutesSaved) + " (My Impact)\n" +
                    "Find any tool:  Ctrl + Shift + Q      Jump to any sheet:  Ctrl + Shift + J\n" +
                    "Not sure which tool to use?  Open the Assistant (first button on the " + Design.TabLabel + " tab)");
        }
    }
}
