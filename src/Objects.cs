using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

// Things on a sheet beyond cell contents and formats: conditional formats, charts and
// tables. All journaled, so `cgr undo` reverts them.
static class Objects
{
    const int xlExpression = 2, xlValue = 2, xlCategory = 1, xlSrcRange = 1, xlYes = 1;
    static readonly object Missing = System.Reflection.Missing.Value;

    static Dictionary<string, object> Entry(string kind, string sheet, string cmd) => new Dictionary<string, object>
    {
        ["t"] = DateTime.Now.ToString("s"), ["cmd"] = cmd, ["kind"] = kind, ["sheet"] = sheet,
    };

    // ---- conditional formats ----

    // Relative references in a rule's formula are relative to the top-left cell of the
    // range it applies to, both in FormatConditions.Add and in Formula1. (Old VBA advice
    // says "relative to the active cell"; on Excel 16.0.20326 over COM, that's wrong.)

    public static void Cf(Xl x, dynamic ws, dynamic rng, List<string> args, string cmd, Out o)
    {
        string sheet = ws.Name;
        if (args.Count == 0) { CfList(x, ws, rng, o); return; }
        if (args[0] == "clear") { CfClear(x, ws, rng, cmd, o); return; }
        if (args[0] != "add" || args.Count < 3 || !args[1].StartsWith("="))
            throw new CgrError("usage: cf RANGE | cf RANGE add =FORMULA fill=#RRGGBB [color=#RRGGBB] [bold=1] [italic=1] | cf RANGE clear");
        string formula = args[1];
        var style = Style(args.Skip(2));
        dynamic fc = rng.FormatConditions.Add(xlExpression, Missing, formula);
        ApplyStyle(fc, style);
        var e = Entry("cf-add", sheet, cmd);
        e["applies"] = (string)fc.AppliesTo.Address;
        e["formula"] = formula;
        Journal.Push(x, e);
        o.Line($"conditional format on {Xl.FullOf(rng)}: {formula}  {string.Join(" ", args.Skip(2))}");
    }

    static Dictionary<string, object> Style(IEnumerable<string> pairs)
    {
        var style = new Dictionary<string, object>();
        foreach (var p in pairs)
        {
            int eq = p.IndexOf('=');
            string k = eq < 0 ? p : p.Substring(0, eq).ToLowerInvariant(), v = eq < 0 ? "" : p.Substring(eq + 1);
            switch (k)
            {
                case "fill": case "color": style[k] = Edit.Color(v); break;
                case "bold": case "italic": style[k] = Edit.Bool(v); break;
                default: throw new CgrError($"'{p}': conditional formats take fill= color= bold= italic=");
            }
        }
        return style;
    }

    static void ApplyStyle(dynamic fc, Dictionary<string, object> style)
    {
        foreach (var kv in style)
            switch (kv.Key)
            {
                case "fill": fc.Interior.Color = kv.Value; break;
                case "color": fc.Font.Color = kv.Value; break;
                case "bold": fc.Font.Bold = kv.Value; break;
                case "italic": fc.Font.Italic = kv.Value; break;
            }
    }

    static Dictionary<string, object> ReadStyle(dynamic fc)
    {
        var style = new Dictionary<string, object>();
        void Take(string key, Func<object> get)
        {
            try { object v = get(); if (v != null && !(v is DBNull) && !(v is bool b && !b)) style[key] = v; } catch (Exception) { }
        }
        Take("fill", () => fc.Interior.Color);
        Take("color", () => fc.Font.Color);
        Take("bold", () => fc.Font.Bold);
        Take("italic", () => fc.Font.Italic);
        return style;
    }

    static readonly Dictionary<int, string> CfTypes = new Dictionary<int, string>
    {
        [1] = "cell value", [2] = "formula", [3] = "color scale", [4] = "data bar", [5] = "top 10",
        [6] = "icon set", [8] = "unique values", [9] = "text", [10] = "blanks", [11] = "time period", [12] = "above average",
    };

    static void CfList(Xl x, dynamic ws, dynamic rng, Out o)
    {
        dynamic conds = rng.FormatConditions;
        int n = conds.Count;
        if (n == 0) { o.Line($"{Xl.FullOf(rng)}: no conditional formats"); return; }
        for (int i = 1; i <= n; i++)
        {
            dynamic fc = conds.Item(i);
            int type = fc.Type;
            string applies = ((string)fc.AppliesTo.Address).Replace("$", "");
            string what = CfTypes.TryGetValue(type, out var t) ? t : "type " + type;
            if (type == xlExpression) what = fc.Formula1;
            Dictionary<string, object> read = ReadStyle(fc);
            var style = read.Select(kv => kv.Key + "=" + Show(kv.Key, kv.Value));
            o.Line($"  [{i}] {applies}  {what}  {string.Join(" ", style)}");
        }
    }

    static string Show(string key, object v) =>
        key == "fill" || key == "color" ? Hex(Convert.ToInt32(v)) : v is bool b ? (b ? "1" : "0") : Convert.ToString(v);

    static string Hex(int bgr) => $"#{bgr & 0xFF:X2}{(bgr >> 8) & 0xFF:X2}{(bgr >> 16) & 0xFF:X2}";

    static void CfClear(Xl x, dynamic ws, dynamic rng, string cmd, Out o)
    {
        dynamic conds = rng.FormatConditions;
        var saved = new List<object>();
        int lost = 0;
        for (int i = 1; i <= (int)conds.Count; i++)
        {
            dynamic fc = conds.Item(i);
            string applies = fc.AppliesTo.Address;
            if ((int)fc.Type != xlExpression) { lost++; continue; }
            saved.Add(new Dictionary<string, object>
            {
                ["applies"] = applies,
                ["formula"] = (string)fc.Formula1,
                ["style"] = ReadStyle(fc),
            });
        }
        rng.FormatConditions.Delete();
        var e = Entry("cf-clear", (string)ws.Name, cmd);
        e["rules"] = saved;
        Journal.Push(x, e);
        o.Line($"cleared conditional formats on {Xl.FullOf(rng)}" + (lost > 0 ? $" ({lost} non-formula rule(s) can't be restored by undo)" : ""));
    }

    // ---- charts ----

    // "Sheet!Chart 1" by name, or "Sheet!1" by index.
    static dynamic FindChart(Xl x, string arg, out string sheet)
    {
        int bang = arg.LastIndexOf('!');
        if (bang < 0) throw new CgrError("charts are named Sheet!ChartName or Sheet!1 (see `cgr outline`)");
        sheet = arg.Substring(0, bang).Trim('\'');
        string name = arg.Substring(bang + 1);
        dynamic ws = x.Sheet(sheet);
        try { return int.TryParse(name, out int i) ? ws.ChartObjects(i) : ws.ChartObjects(name); }
        catch (System.Runtime.InteropServices.COMException) { throw new CgrError($"no chart '{name}' on {sheet}"); }
    }

    static readonly string[] ChartKeys = { "title", "ytitle", "xtitle", "yfmt" };

    static string GetChart(dynamic ch, string key)
    {
        switch (key)
        {
            case "title": return (bool)ch.HasTitle ? (string)ch.ChartTitle.Text : "";
            case "ytitle": return (bool)ch.Axes(xlValue).HasTitle ? (string)ch.Axes(xlValue).AxisTitle.Text : "";
            case "xtitle": return (bool)ch.Axes(xlCategory).HasTitle ? (string)ch.Axes(xlCategory).AxisTitle.Text : "";
            case "yfmt": return ch.Axes(xlValue).TickLabels.NumberFormat;
        }
        throw new CgrError("unknown chart property " + key);
    }

    static void SetChart(dynamic ch, string key, string v)
    {
        switch (key)
        {
            case "title":
                ch.HasTitle = v.Length > 0;
                if (v.Length > 0) ch.ChartTitle.Text = v;
                break;
            case "ytitle":
            case "xtitle":
                dynamic axis = ch.Axes(key == "ytitle" ? xlValue : xlCategory);
                axis.HasTitle = v.Length > 0;
                if (v.Length > 0) axis.AxisTitle.Text = v;
                break;
            case "yfmt":
                ch.Axes(xlValue).TickLabels.NumberFormat = v;
                break;
        }
    }

    public static void Chart(Xl x, string arg, List<string> pairs, string cmd, Out o)
    {
        dynamic co = FindChart(x, arg, out string sheet);
        dynamic ch = co.Chart;
        if (pairs.Count == 0)
        {
            dynamic tl = co.TopLeftCell;   // scalar properties only: see README, "Excel bugs"
            o.Line($"{Cells.Sheet(sheet)}!{co.Name}  at {Cells.Addr((int)tl.Row, (int)tl.Column)}  type {ch.ChartType}");
            foreach (var k in ChartKeys) o.Line($"  {k,-7} {GetChart(ch, k)}");
            foreach (dynamic s in ch.SeriesCollection()) o.Line($"  series {Cells.Trunc((string)s.Formula, 200)}");
            return;
        }
        var e = Entry("chart", sheet, cmd);
        e["chart"] = (string)co.Name;
        var before = new Dictionary<string, object>();
        foreach (var p in pairs)
        {
            int eq = p.IndexOf('=');
            string k = eq < 0 ? "" : p.Substring(0, eq).ToLowerInvariant();
            if (!ChartKeys.Contains(k)) throw new CgrError($"'{p}': charts take {string.Join("= ", ChartKeys)}=");
            before[k] = GetChart(ch, k);
            SetChart(ch, k, p.Substring(eq + 1));
        }
        e["before"] = before;
        Journal.Push(x, e);
        o.Line($"chart {Cells.Sheet(sheet)}!{co.Name}: {string.Join(" ", pairs)}");
    }

    // ---- tables ----

    public static void Table(Xl x, dynamic ws, dynamic rng, string name, string style, string cmd, Out o)
    {
        dynamic lo = ws.ListObjects.Add(xlSrcRange, rng, Missing, xlYes);
        if (name != null) lo.Name = name;
        lo.TableStyle = style ?? "TableStyleMedium2";
        var e = Entry("table", (string)ws.Name, cmd);
        e["table"] = (string)lo.Name;
        Journal.Push(x, e);
        o.Line($"table {lo.Name} on {Xl.FullOf(lo.Range)}, style {lo.TableStyle.Name}");
    }

    // ---- undo ----

    public static void Undo(Xl x, dynamic ws, string kind, Dictionary<string, object> e)
    {
        switch (kind)
        {
            case "cf-add":
            {
                string applies = (string)e["applies"], formula = (string)e["formula"];
                dynamic conds = ws.Range[applies].FormatConditions;
                for (int i = (int)conds.Count; i >= 1; i--)
                {
                    dynamic fc = conds.Item(i);
                    if ((int)fc.Type == xlExpression && (string)fc.AppliesTo.Address == applies
                        && (string)fc.Formula1 == formula)
                    {
                        fc.Delete();
                        return;
                    }
                }
                throw new CgrError($"the conditional format {formula} on {applies} is gone already");
            }
            case "cf-clear":
            {
                foreach (var rule in ((IEnumerable)e["rules"]).Cast<Dictionary<string, object>>())
                {
                    string applies = (string)rule["applies"];
                    dynamic fc = ws.Range[applies].FormatConditions.Add(xlExpression, Missing, (string)rule["formula"]);
                    ApplyStyle(fc, ((Dictionary<string, object>)rule["style"]).ToDictionary(kv => kv.Key, kv => Journal.FromJson(kv.Value)));
                }
                return;
            }
            case "chart":
            {
                dynamic ch = ws.ChartObjects((string)e["chart"]).Chart;
                foreach (var kv in (Dictionary<string, object>)e["before"]) SetChart(ch, kv.Key, (string)kv.Value);
                return;
            }
            case "table":
                dynamic lo = ws.ListObjects.Item((string)e["table"]);
                lo.TableStyle = "";   // else Unlist bakes the style (stripes etc.) into the cells
                lo.Unlist();
                return;
            default:
                throw new CgrError($"don't know how to undo '{kind}'");
        }
    }
}
