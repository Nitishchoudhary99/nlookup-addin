using System;
using System.Linq;
using System.Reflection;
using System.Globalization;
using ExcelDna.Integration;
// Runs the pure logic of the add-in (no Excel needed) under Mono or .NET Framework: see tests/run-tests.sh
class T {
  static Assembly A = Assembly.LoadFrom("SmartPro.dll");
  static int fails = 0;
  static object Call(string type, string m, params object[] a) {
    Type t = A.GetType("SmartPro." + type);
    foreach (MethodInfo mi in t.GetMethods(BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic))
      if (mi.Name == m && mi.GetParameters().Length == a.Length) return mi.Invoke(null, a);
    throw new Exception("no method " + type + "." + m);
  }
  static object Prop(string type, string p) {
    return A.GetType("SmartPro." + type).GetProperty(p, BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic).GetValue(null);
  }
  static void Eq(object got, object want, string what) {
    bool ok = object.Equals(got, want);
    if (!ok) fails++;
    Console.WriteLine((ok ? "ok   " : "FAIL ") + what + "  got=" + got + (ok ? "" : " want=" + want));
  }
  static void Set(string k, string v) { Call("Settings", "Set", k, v); }
  static string Day(int add) { return DateTime.Now.AddDays(add).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); }
  static void Main() {
    // ---- keys ----
    Eq(Call("SmartLookup","LooseKey"," ABC   Ltd "), "t:abc ltd", "loose spaces/nbsp/case");
    Eq(Call("SmartLookup","LooseKey","00123"), "n:123", "loose number text");
    Eq(Call("SmartLookup","LooseKey",123.0), "n:123", "loose number");
    Eq(Call("SmartLookup","LooseKey","'123"), "n:123", "loose apostrophe");
    Eq(Call("SmartLookup","LooseKey","1-2"), "t:1-2", "code stays text");
    Eq(Call("SmartLookup","LooseKey","1e5"), "t:1e5", "1e5 stays text");
    Eq(Call("SmartLookup","LooseKey","  "), null, "blank");
    Eq(Call("SmartLookup","StrictKey","00123"), "t:00123", "strict text");
    Eq(Call("SmartLookup","StrictKey",123.0), "n:123", "strict num");
    // ---- SMARTLOOKUP function (0-based arrays like Excel-DNA) ----
    Type lf = A.GetType("SmartPro.LookupFunctions");
    MethodInfo fn = lf.GetMethod("SmartLookupFn");
    object[,] keys = { { "Acme Private Limited" }, { 1001.0 }, { " Zeta Traders " } };
    object[,] vals = { { "A" }, { "B" }, { "Z" } };
    Eq(fn.Invoke(null, new object[] { "zeta traders", keys, vals, ExcelMissing.Value, ExcelMissing.Value }), "Z", "fn spaces/case");
    Eq(fn.Invoke(null, new object[] { "1001", keys, vals, ExcelMissing.Value, ExcelMissing.Value }), "B", "fn number as text");
    Eq(fn.Invoke(null, new object[] { "Nobody", keys, vals, ExcelMissing.Value, ExcelMissing.Value }), ExcelError.ExcelErrorNA, "fn not found NA");
    Eq(fn.Invoke(null, new object[] { "Nobody", keys, vals, "-", ExcelMissing.Value }), "-", "fn if_not_found");
    Eq(fn.Invoke(null, new object[] { "Acme Pvt Ltd", keys, vals, "-", true }), "A", "fn fuzzy Pvt/Private");
    Eq(fn.Invoke(null, new object[] { "Acme Pvt Ltd", keys, vals, "-", false }), "-", "fn no fuzzy");
    // ---- habit (uses the registry of the user running the test: Mono keeps it under ~/.mono/registry) ----
    foreach (string k in new[]{"H_First","H_Last","H_Streak","H_Best","H_Min","H_Day","H_DayMin","H_DayRuns","H_Month","H_MonthMin","H_Tools","H_ToolRuns","H_Miles","Bonus","H_ProMin","H_ProRuns","Runs","PW_Start","PW_End","PW_Tomorrow","N_Day","N_DayCount","N_LastRun","Trial_analyst"}) Set(k, "");
    for (int i = 0; i < 5; i++) { Set("Runs", (i+1).ToString()); Call("Habit","Record","clean"); }
    Eq(Prop("Habit","MinutesSaved"), 30.0, "clean x5 counts 3 x 10 min");
    Eq(Prop("Habit","Streak"), 1, "streak 1 first day");
    // streak continues from the previous working day
    DateTime prev = DateTime.Now.Date.AddDays(-1);
    while (prev.DayOfWeek == DayOfWeek.Saturday || prev.DayOfWeek == DayOfWeek.Sunday) prev = prev.AddDays(-1);
    Set("H_Last", prev.ToString("yyyy-MM-dd")); Set("H_Streak", "2");
    Set("Runs", "10");
    var reached = (System.Collections.IList)Call("Habit","Record","dash");
    Eq(Prop("Habit","Streak"), 3, "streak continues to 3");
    string names = ""; foreach (object m in reached) names += m.GetType().GetField("Id").GetValue(m) + ",";
    Eq(names, "runs10,streak3,hour1,", "milestones reached (75 min also passes 1 hour)");
    Eq(Prop("Habit","Bonus"), 4, "bonus runs given (1+1+2)");
    // broken streak
    Set("H_Last", DateTime.Now.Date.AddDays(-9).ToString("yyyy-MM-dd")); Set("H_Streak", "7");
    Eq(Prop("Habit","Streak"), 0, "streak broken after 9 days");
    Eq(Prop("Habit","BestStreak"), 3, "best streak kept");
    // ---- pro week & licence ----
    Set("H_First", Day(0));
    Eq(Prop("Habit","WelcomeDay"), 1, "pro week day 1");
    Eq(Call("Licence","AllowPro","analyst","Smart Analyst"), true, "pro week lets Pro tool run");
    Eq(Call("Licence","TrialUsed","analyst"), 0, "pro week does not use a free run");
    Set("H_First", Day(-6));
    Eq(Prop("Habit","WelcomeDay"), 7, "pro week day 7");
    Set("H_First", Day(-7));
    Eq(Prop("Habit","WelcomeDay"), 0, "pro week over on day 8");
    Eq(Prop("Habit","WelcomeOver"), true, "welcome over");
    for (int i = 0; i < 3; i++) Eq(Call("Licence","AllowPro","analyst","Smart Analyst"), true, "free run " + (i+1));
    Eq(Call("Licence","TrialUsed","analyst"), 3, "3 free runs used");
    Eq(Call("Licence","AllowPro","analyst","Smart Analyst"), true, "bonus run 1");
    Eq(Call("Licence","AllowPro","analyst","Smart Analyst"), true, "bonus run 2");
    Eq(Prop("Habit","Bonus"), 2, "two bonus runs used, two left");
    Console.WriteLine("note: " + Call("Licence","TakeNote"));
    Console.WriteLine("greeting: " + Call("Habit","Greeting"));
    Console.WriteLine("plan: " + Call("Licence","PlanText"));
    Console.WriteLine("perday: " + Prop("ProPlan","PerDay") + " | " + Prop("ProPlan","SavingLine"));
    Eq(Call("Habit","Hours",95.0), "1.6 hours", "hours fmt");
    Console.WriteLine("rupees 120min: " + Call("Habit","Rupees",120.0));

    // ---- 3.3: ribbon, tips table, AI parsing, exports, formula recipes ----
    string xml = (string)Call("Design","RibbonXml");
    var doc = new System.Xml.XmlDocument(); doc.LoadXml(xml);
    Eq(doc.GetElementsByTagName("button").Count, ((Array)A.GetType("SmartPro.Design").GetField("Tools", BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic).GetValue(null)).Length, "ribbon XML is valid and has every tool");
    Eq(Prop("Nudge","TipsOff") is bool, true, "tips table loads (no duplicate keys)");
    string[] p1 = (string[])Call("AiFormula","Parse","FORMULA: =SUMIFS($C$2:$C$500,$A$2:$A$500,A2)\nWHY: Adds the amount of this customer.");
    Eq(p1[0], "=SUMIFS($C$2:$C$500,$A$2:$A$500,A2)", "AI reply: formula");
    Eq(p1[1], "Adds the amount of this customer.", "AI reply: explanation");
    string[] p2 = (string[])Call("AiFormula","Parse","FORMULA: `=TODAY()-B2`\nWHY: Days since\nthe date.");
    Eq(p2[0] + "|" + p2[1], "=TODAY()-B2|Days since the date.", "AI reply: backticks and two-line why");
    string[] p3 = (string[])Call("AiFormula","Parse","FORMULA: none\nWHY: Needs a macro.");
    Eq(p3[0], null, "AI reply: none means no formula");
    Eq(Call("ExportMore","Ident","Sales Data 2026!"), "sales_data_2026", "SQL identifier");
    Eq(Call("ExportMore","Ident","2026 sales"), "t_2026_sales", "SQL identifier starting with a digit");
    object[,] t = (object[,])Array.CreateInstance(typeof(object), new[]{4,4}, new[]{1,1});
    t[1,1]="Party"; t[1,2]="Amount"; t[1,3]="Date"; t[1,4]="Qty";
    t[2,1]="O'Neil Traders"; t[2,2]=1250.5; t[2,3]=new DateTime(2026,4,1); t[2,4]=3.0;
    t[3,1]=null; t[3,2]=null; t[3,3]=null; t[3,4]=null;
    t[4,1]="ABC"; t[4,2]=99.0; t[4,3]=new DateTime(2026,5,2); t[4,4]=7.0;
    string[] h = { "", "Party", "Amount", "Date", "Qty" };
    string[] ty = { "", "text", "dec", "date", "int" };
    int[] ml = { 0, 14, 0, 0, 0 };
    string sql = (string)Call("ExportMore","Sql", t, h, ty, ml, "sales", 2);
    Eq(sql.Contains("\"Amount\" DECIMAL(18,4)") && sql.Contains("'O''Neil Traders', 1250.5, '2026-04-01 00:00:00', 3)") && !sql.Contains("NULL, NULL"), true, "SQL: types, quotes escaped, blank row skipped");
    Console.WriteLine(sql);
    string py = (string)Call("ExportMore","Python", t, h, ty, "sales");
    Eq(py.Contains("\"Party\": [\"O'Neil Traders\", \"ABC\"]") && py.Contains("astype(\"Int64\")"), true, "Python: lists without the blank row");
    var recipes = (Array)A.GetType("SmartPro.FormulaLibrary").GetField("All", BindingFlags.Static|BindingFlags.NonPublic).GetValue(null);
    int badRecipes = 0;
    foreach (object r in recipes) {
      string tpl = (string)r.GetType().GetField("Template").GetValue(r);
      int depth = 0; foreach (char ch in tpl) { if (ch=='(') depth++; if (ch==')') depth--; }
      bool kinds = System.Text.RegularExpressions.Regex.Matches(tpl, @"\{([^:}]+):").Cast<System.Text.RegularExpressions.Match>().All(m => "craN".Contains(m.Groups[1].Value) || m.Groups[1].Value=="n");
      if (!tpl.StartsWith("=") || !kinds) { badRecipes++; Console.WriteLine("bad recipe: " + tpl); }
    }
    Eq(badRecipes, 0, "formula recipes: start with = and use known placeholders");
    Eq(recipes.Length, 41, "41 formula recipes");
    Console.WriteLine(fails == 0 ? "ALL PASSED" : fails + " FAILED");
    Environment.Exit(fails == 0 ? 0 : 1);
  }
}
