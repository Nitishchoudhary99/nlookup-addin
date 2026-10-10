using System;
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
    Console.WriteLine(fails == 0 ? "ALL PASSED" : fails + " FAILED");
    Environment.Exit(fails == 0 ? 0 : 1);
  }
}
