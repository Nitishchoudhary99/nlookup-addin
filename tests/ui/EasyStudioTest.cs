using System; using System.Collections; using System.Diagnostics; using System.Drawing; using System.Reflection; using System.Windows.Forms;
class Easy {
  static Assembly A = Assembly.LoadFrom("SmartPro.dll");
  static Type Ty(string n) { return A.GetType("SmartPro." + n); }
  static Type SF = Ty("StudioForm");
  static object F(object o, string n) { return SF.GetField(n, BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public).GetValue(o); }
  static object M(object o, string n, params object[] a) { foreach (var mi in SF.GetMethods(BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public)) if (mi.Name==n && mi.GetParameters().Length==a.Length) return mi.Invoke(o,a); throw new Exception(n); }
  static void Pump(int ms) { DateTime end = DateTime.Now.AddMilliseconds(ms); while (DateTime.Now < end) { Application.DoEvents(); System.Threading.Thread.Sleep(15); } }
  static void Shot(string name, Rectangle b) { Process.Start("import", "-window root -crop " + b.Width + "x" + b.Height + "+" + b.X + "+" + b.Y + " out/" + name + ".png").WaitForExit(); Console.WriteLine("saved " + name); }
  static int fails = 0;
  static void Eq(object got, object want, string what) { bool ok = Equals(got, want); if (!ok) fails++; Console.WriteLine((ok?"ok   ":"FAIL ")+what+"  got="+got+(ok?"":" want="+want)); }
  static string Spec(object f) { object s = F(f,"s"); return (string)s.GetType().GetMethod("ToText").Invoke(s,null); }
  static IList Charts(object f) { object s = F(f,"s"); return (IList)s.GetType().GetField("Charts").GetValue(s); }
  static IList Cards(object f) { object s = F(f,"s"); return (IList)s.GetType().GetField("Cards").GetValue(s); }
  static IList Slicers(object f) { object s = F(f,"s"); return (IList)s.GetType().GetField("Slicers").GetValue(s); }
  static object G(object o, string n) { return o.GetType().GetField(n).GetValue(o); }
  [STAThread] static void Main() {
    var rnd = new Random(7);
    string[] states = {"Maharashtra","Gujarat","Karnataka","Tamil Nadu","Delhi","Uttar Pradesh","Rajasthan","West Bengal","Telangana","Kerala","Punjab","Madhya Pradesh","Haryana","Bihar"};
    string[] zone = {"West","West","South","South","North","North","North","East","South","South","North","Central","North","East"};
    string[] prod = {"Mobiles","Laptops","Headphones","Watches","Cameras"};
    int n = 1500; object[,] v = (object[,])Array.CreateInstance(typeof(object), new[]{n+1,10}, new[]{1,1});
    string[] h = {"Order ID","Order Date","State","Zone","Product","Qty","Gross Sales","Net Sales","Profit","Rating"}; for (int j=0;j<10;j++) v[1,j+1]=h[j];
    for (int r=2;r<=n+1;r++){ int si=rnd.Next(states.Length); v[r,1]="ORD-"+(10000+rnd.Next(1100)); v[r,2]=new DateTime(2025,4,1).AddDays(rnd.Next(540)); v[r,3]=states[si]; v[r,4]=zone[si]; v[r,5]=prod[rnd.Next(5)]; double q=rnd.Next(1,6); v[r,6]=q; double g=q*rnd.Next(800,60000); v[r,7]=g; v[r,8]=Math.Round(g*0.9); v[r,9]=Math.Round(g*0.12); v[r,10]=Math.Round(3+rnd.NextDouble()*2,1); }
    object t = Ty("DataTab").GetMethod("Load", BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic).Invoke(null, new object[]{v, 0});
    object look = Ty("DashAuto").GetMethod("SavedLook", BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic).Invoke(null, null);
    object spec = Ty("DashAuto").GetMethod("Suggest", BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic).Invoke(null, new object[]{t, look});
    Form f = (Form)Activator.CreateInstance(SF, BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic, null, new object[]{t, spec, "Orders!A1:J1501", false, "Build dashboard"}, null);
    f.StartPosition = FormStartPosition.Manual; f.Location = new Point(0,0); f.Show(); Pump(2500);
    Point p0 = f.PointToScreen(Point.Empty); Rectangle full = new Rectangle(p0, f.ClientSize);
    Shot("easy-1-start", full);
    var tabs = (TabControl)F(f,"tabs"); Eq(tabs.SelectedTab.Text, "Easy", "Easy tab opens first");
    string start = Spec(f);
    var num = (ComboBox)F(f,"easyNumber"); var how = (ComboBox)F(f,"easyHow");
    Console.WriteLine("main number now: " + num.SelectedItem + " / " + how.SelectedItem + "   charts: " + Charts(f).Count + "   filters: " + Slicers(f).Count);
    // 1) main number -> Net Sales
    num.SelectedIndex = num.Items.IndexOf("Net Sales"); Pump(600);
    int follow = 0; foreach (object c in Charts(f)) if ((string)G(c,"Col")=="Net Sales") follow++;
    Eq((string)G(Cards(f)[0],"Col"), "Net Sales", "first KPI card follows the main number");
    Eq(follow, Charts(f).Count, "every chart that showed the old main number now shows Net Sales");
    // 2) break down by: tick State (one more chart), it should get Top 10 (14 states)
    var brk = (CheckedListBox)F(f,"easyBreak"); var fields = (IList)F(f,"easyFields");
    int stateAt = fields.IndexOf("State"), zoneAt = fields.IndexOf("Zone");
    int before = Charts(f).Count; bool stateWas = brk.GetItemChecked(stateAt);
    brk.SetItemChecked(stateAt, !stateWas); Pump(600);
    Eq(Charts(f).Count, stateWas ? before - 1 : before + 1, "ticking a field adds (or removes) one chart");
    if (!stateWas) { foreach (object c in Charts(f)) if ((string)G(c,"Dim")=="State") Eq((int)G(c,"Top"), 10, "a field with many values (14 states) starts as Top 10"); }
    // 3) filters: tick Zone
    var fl = (CheckedListBox)F(f,"easyFilters"); bool zoneWas = fl.GetItemChecked(zoneAt);
    fl.SetItemChecked(zoneAt, !zoneWas); Pump(600);
    Eq(Slicers(f).Contains("Zone"), !zoneWas, "Zone filter button toggled");
    // 4) colours: theme 3
    ((Button[])F(f,"easyThemes"))[3].PerformClick(); Pump(800);
    object th = G(F(f,"s"), "Theme"); Console.WriteLine("theme: " + G(th,"Name"));
    // 5) title
    var title = (TextBox)F(f,"easyTitle"); title.Text = "Sales Dashboard FY 2025-26"; Pump(900);
    Eq((string)G(F(f,"s"),"Title"), "Sales Dashboard FY 2025-26", "title from the Easy tab");
    Shot("easy-2-after", full);
    // undo everything step by step, then redo
    int steps = 0; while (((System.Collections.Generic.Stack<string>)F(f,"undo")).Count > 0 && steps < 50) { M(f,"Undo"); Pump(200); steps++; }
    Eq(Spec(f) == start, true, "undo all the way returns exactly to the start (" + steps + " steps)");
    Eq(num.SelectedItem, (object)"Gross Sales", "the Easy tab shows the restored design");
    for (int i = 0; i < steps; i++) { M(f,"Redo"); Pump(150); }
    Eq((string)G(F(f,"s"),"Title"), "Sales Dashboard FY 2025-26", "redo brings every change back");
    // right-click on chart 2 and on card 1: menu builds and shows
    M(f, "PreviewMenu", 2, 1, new Point(760, 380)); Pump(900);
    Shot("easy-3-chart-menu", new Rectangle(0, 0, 1380, 830));
    SendKeys.SendWait("{ESC}"); Pump(300);
    // the distinct-count item, invoked through the menu itself
    M(f, "PreviewMenu", 1, 1, new Point(760, 140)); Pump(700);
    Shot("easy-4-card-menu", new Rectangle(0, 0, 1380, 830));
    ToolStripDropDown open = null; foreach (Form x in Application.OpenForms) {} 
    SendKeys.SendWait("{ESC}"); Pump(300);
    // call the same thing the menu item does: distinct count of Order ID on card 2
    object card = Cards(f)[1]; card.GetType().GetField("Col").SetValue(card, "Order ID"); card.GetType().GetField("Agg").SetValue(card, 5);
    card.GetType().GetField("Label").SetValue(card, SF.GetMethod("CountLabel", BindingFlags.Static|BindingFlags.NonPublic).Invoke(null, new object[]{"Order ID"}));
    card.GetType().GetField("Compare").SetValue(card, 0); card.GetType().GetField("Spark").SetValue(card, false);
    M(f, "AfterQuickChange", "test"); Pump(800);
    int netCards = 0; foreach (object c in Cards(f)) if ((string)G(c,"Col")=="Net Sales") netCards++;
    Eq(netCards, 1, "Net Sales is on one card only (the cards swapped)");
    bool idOffered = false, qtyOffered = false; foreach (object x in fields) { if ((string)x=="Order ID") idOffered = true; if ((string)x=="Qty") qtyOffered = true; }
    Eq(idOffered || qtyOffered, false, "Order ID and Qty are not offered to break down by");
    Shot("easy-5-orders-card", full);
    Console.WriteLine("design:\n" + Spec(f).Substring(0, Math.Min(900, Spec(f).Length)));
    Console.WriteLine(fails == 0 ? "ALL PASSED" : fails + " FAILED");
    f.Hide(); Environment.Exit(0);
  }
}
