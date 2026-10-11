// Calculated KPI cards (Profit % of Net Sales, Gross - Net ...): maths, saved design, Excel formula, preview, web page data.
// Run by run-tests.sh after LogicTests; writes calc-web.html for the browser check in calc-web.mjs.
using System; using System.Collections; using System.IO; using System.Reflection;
class CalcCardTest {
  static Assembly A = Assembly.LoadFrom("SmartPro.dll");
  static Type Ty(string n) { return A.GetType("SmartPro." + n); }
  static object S(string type, string m, params object[] a) { foreach (var mi in Ty(type).GetMethods(BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic)) if (mi.Name == m && mi.GetParameters().Length == a.Length) return mi.Invoke(null, a); throw new Exception(type + "." + m); }
  static void Set(object o, string f, object v) { o.GetType().GetField(f).SetValue(o, v); }
  static object Get(object o, string f) { return o.GetType().GetField(f).GetValue(o); }
  static int fails = 0;
  static void Eq(object got, object want, string what) { bool ok = Equals(got, want); if (!ok) fails++; Console.WriteLine((ok ? "ok   " : "FAIL ") + what + "  got=" + got + (ok ? "" : " want=" + want)); }
  static void Main() {
    // maths
    Eq(S("CardOp", "Apply", 1, 25.0, 100.0), 0.25, "% of");
    Eq(S("CardOp", "Apply", 1, 25.0, 0.0), 0.0, "% of zero is 0, not an error");
    Eq(S("CardOp", "Apply", 3, 100.0, 30.0), 70.0, "minus");
    // a table: Zone, Gross, Net, Profit
    var rnd = new Random(3); int n = 400; string[] zones = { "North", "South", "East", "West" };
    object[,] v = (object[,])Array.CreateInstance(typeof(object), new[] { n + 1, 5 }, new[] { 1, 1 });
    string[] h = { "Order Date", "Zone", "Gross Sales", "Net Sales", "Profit" }; for (int j = 0; j < 5; j++) v[1, j + 1] = h[j];
    double net = 0, profit = 0, gross = 0, netN = 0, profitN = 0;
    for (int r = 2; r <= n + 1; r++) { v[r, 1] = new DateTime(2026, 1, 1).AddDays(rnd.Next(200)); string z = zones[rnd.Next(4)]; v[r, 2] = z; double g = rnd.Next(1000, 90000); v[r, 3] = g; v[r, 4] = Math.Round(g * 0.9); v[r, 5] = Math.Round(g * (0.05 + rnd.NextDouble() * 0.2)); gross += g; net += (double)v[r, 4]; profit += (double)v[r, 5]; if (z == "North") { netN += (double)v[r, 4]; profitN += (double)v[r, 5]; } }
    object t = S("DataTab", "Load", v, 0);
    object spec = S("DashAuto", "Suggest", t, S("DashAuto", "SavedLook"));
    IList cards = (IList)Get(spec, "Cards");
    object c = Activator.CreateInstance(Ty("CardSpec"), true);
    Set(c, "Col", "Profit"); Set(c, "Agg", 0); Set(c, "Op", 1); Set(c, "Col2", "Net Sales"); Set(c, "Agg2", 0); Set(c, "Compare", 1); Set(c, "Spark", true);
    cards.Insert(0, c);
    object c2 = Activator.CreateInstance(Ty("CardSpec"), true);
    Set(c2, "Col", "Gross Sales"); Set(c2, "Agg", 0); Set(c2, "Op", 3); Set(c2, "Col2", "Net Sales"); Set(c2, "Agg2", 0);
    cards.Insert(1, c2);
    S("DashAuto", "Fix", spec, t);
    Eq((int)Get(c, "Compare") + "|" + (bool)Get(c, "Spark"), "0|False", "a calculated card has no month comparison or trend line");
    Eq(S("DashText", "CardLabel", c), "Profit % of Net Sales", "automatic label");
    Eq(S("DashText", "CardLabel", c2), "Gross Sales − Net Sales", "automatic label (minus)");
    Eq(S("DashText", "CardNote", c), "Profit ÷ Net Sales", "note under the number");
    // preview number = hand calculation
    object cache = Activator.CreateInstance(Ty("DashCache"), true);
    double pv = (double)S("DashRender", "CardValue", t, cache, c);
    Eq(Math.Round(pv, 6), Math.Round(profit / net, 6), "preview: Profit % of Net Sales = sum(Profit) / sum(Net Sales)");
    Console.WriteLine("     shown as: " + S("DashRender", "CardValueText", t, spec, cache, c) + "  and  " + S("DashRender", "CardValueText", t, spec, cache, c2));
    Eq(((string)S("DashRender", "CardValueText", t, spec, cache, c)).EndsWith("%"), true, "shown as a percent");
    Eq(Math.Round((double)S("DashRender", "CardValue", t, cache, c2)), Math.Round(gross - net), "preview: Gross - Net");
    // saved design keeps the calculation
    string text = (string)spec.GetType().GetMethod("ToText").Invoke(spec, null);
    object back = S("DashSpec", "FromText", text);
    object b0 = ((IList)Get(back, "Cards"))[0];
    Eq(Get(b0, "Op") + "|" + Get(b0, "Col2") + "|" + Get(b0, "Agg2"), "1|Net Sales|0", "save / open keeps the calculation");
    // a missing second column switches the calculation off with a note
    object bad = ((IList)Get(back, "Cards"))[1]; Set(bad, "Col2", "Discount");
    IList notes = (IList)S("DashAuto", "Fix", back, t);
    Eq(Get(bad, "Op"), 0, "missing second column: shows the first number only");
    Eq(notes.Count > 0, true, "and says why");
    // Excel: both numbers from the KPI pivot table, so the slicers change the card
    Eq(S("DashFormula", "KpiCalc", 1, "K1", "K1b", "$A$1"), "=IFERROR(IFERROR(GETPIVOTDATA(\"K1\",$A$1),0)/IFERROR(GETPIVOTDATA(\"K1b\",$A$1),0),0)", "Excel formula of the card");
    // web page with the same design
    int rows, left;
    string page = (string)S("WebDash", "Page", t, spec, "Sales!A1:E401", "test", null, null);
    File.WriteAllText("calc-web.html", page);
    File.WriteAllText("calc-expect.json", "{\"all\":" + (profit / net).ToString("R") + ",\"north\":" + (profitN / netN).ToString("R") + ",\"minus\":" + (gross - net).ToString("R") + "}");
    Console.WriteLine(fails == 0 ? "ALL PASSED" : fails + " FAILED");
    Environment.Exit(fails == 0 ? 0 : 1);
  }
}
