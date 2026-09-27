using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

// Writes into the live workbook. Excel drops its own undo history whenever a program
// changes a cell, so every write records what it replaced in a journal, and `undo`
// plays that back.
static class Edit
{
    public static void Set(Xl x, dynamic ws, dynamic rng, string value, string cmd, Out o)
    {
        string sheet = ws.Name;
        object toWrite = value;
        if (value == "-")
        {
            var rows = ReadTsv();
            if (rows.Count == 0) throw new CgrError("nothing on stdin");
            int w = rows.Max(r => r.Length);
            var grid = new object[rows.Count, w];
            for (int i = 0; i < rows.Count; i++)
                for (int j = 0; j < w; j++)
                    grid[i, j] = j < rows[i].Length ? rows[i][j] : "";
            int r0 = rng.Row, c0 = rng.Column;
            rng = ws.Range[Cells.Addr(r0, c0, r0 + rows.Count - 1, c0 + w - 1)];
            toWrite = grid;
        }
        var entry = Snapshot(ws, rng, "content", cmd);
        entry["fmt"] = RunsToJson(FormatRuns(ws, rng, Props["numfmt"]));
        Block before = Xl.Read(rng, false);
        Xl.SetFormulas(rng, toWrite);
        Block after = Xl.Read(rng, false);
        entry["after"] = Jagged(after.F);
        Journal.Push(x, entry);
        Report(before, after, o);
    }

    static void Report(Block before, Block after, Out o)
    {
        int cells = after.Rows * after.Cols, shown = 0, changed = 0;
        o.Line($"wrote {Cells.Full(after.Sheet, after.Row, after.Col, after.Row2, after.Col2)} ({cells} cell{(cells == 1 ? "" : "s")})");
        for (int i = 0; i < after.Rows; i++)
            for (int j = 0; j < after.Cols; j++)
            {
                if (before.Formula(i, j) == after.Formula(i, j)) continue;
                changed++;
                if (++shown > 12) continue;
                string was = before.Formula(i, j).Length == 0 ? "∅" : Cells.Trunc(before.Formula(i, j), 60);
                string now = after.IsFormula(i, j) ? Cells.Trunc(after.Formula(i, j), 80) + " → " + Cells.Raw(after.V[i, j]) : Cells.Raw(after.V[i, j]);
                o.Line($"  {Cells.Addr(after.Row + i, after.Col + j),-8} {was}  ⇒  {now}");
            }
        if (shown > 12) o.Line($"  … {shown - 12} more changed");
        if (changed == 0) o.Line("  (no cell changed)");
    }

    static List<string[]> ReadTsv()
    {
        var reader = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        var lines = reader.ReadToEnd().Replace("\r\n", "\n").Split('\n').ToList();
        if (lines.Count > 0 && lines[lines.Count - 1].Length == 0) lines.RemoveAt(lines.Count - 1);
        return lines.Select(l => l.Split('\t')).ToList();
    }

    // Formatting properties `fmt` understands. Get returns null when the range is mixed.
    class Prop
    {
        public Func<dynamic, object> Get;
        public Action<dynamic, object> Set;
        public Func<string, object> Parse;
    }

    const int xlNone = -4142;

    static readonly Dictionary<string, Prop> Props = new Dictionary<string, Prop>(StringComparer.OrdinalIgnoreCase)
    {
        ["numfmt"] = new Prop { Get = r => r.NumberFormat, Set = (r, v) => r.NumberFormat = v, Parse = s => s },
        ["bold"] = new Prop { Get = r => r.Font.Bold, Set = (r, v) => r.Font.Bold = v, Parse = Bool },
        ["italic"] = new Prop { Get = r => r.Font.Italic, Set = (r, v) => r.Font.Italic = v, Parse = Bool },
        ["indent"] = new Prop { Get = r => r.IndentLevel, Set = (r, v) => r.IndentLevel = v, Parse = s => int.Parse(s, Cells.Inv) },
        ["size"] = new Prop { Get = r => r.Font.Size, Set = (r, v) => r.Font.Size = v, Parse = s => double.Parse(s, Cells.Inv) },
        ["font"] = new Prop { Get = r => r.Font.Name, Set = (r, v) => r.Font.Name = v, Parse = s => s },
        ["color"] = new Prop { Get = r => r.Font.Color, Set = (r, v) => r.Font.Color = v, Parse = Color },
        ["fill"] = new Prop { Get = GetFill, Set = SetFill, Parse = s => s.Equals("none", StringComparison.OrdinalIgnoreCase) ? (object)"none" : Color(s) },
        ["width"] = new Prop { Get = r => r.ColumnWidth, Set = (r, v) => r.ColumnWidth = v, Parse = s => double.Parse(s, Cells.Inv) },
        ["wrap"] = new Prop { Get = r => r.WrapText, Set = (r, v) => r.WrapText = v, Parse = Bool },
        ["align"] = new Prop { Get = r => r.HorizontalAlignment, Set = (r, v) => r.HorizontalAlignment = v, Parse = Align },
    };

    internal static object Bool(string s) => s == "1" || s.Equals("true", StringComparison.OrdinalIgnoreCase) || s.Equals("yes", StringComparison.OrdinalIgnoreCase);

    // #RRGGBB; Excel stores colors as 0xBBGGRR.
    internal static object Color(string s)
    {
        s = s.TrimStart('#');
        if (s.Length != 6) throw new CgrError($"color '{s}' should be #RRGGBB");
        int rgb = Convert.ToInt32(s, 16);
        return (double)(((rgb >> 16) & 0xFF) | (rgb & 0xFF00) | ((rgb & 0xFF) << 16));
    }

    static object Align(string s)
    {
        switch (s.ToLowerInvariant())
        {
            case "left": return -4131;
            case "center": return -4108;
            case "right": return -4152;
            case "general": return 1;
            default: throw new CgrError("align is left, center, right or general");
        }
    }

    static object GetFill(dynamic r)
    {
        object index = r.Interior.ColorIndex;
        if (index == null || index is DBNull) return null;
        if (Convert.ToInt32(index) == xlNone) return "none";
        object color = r.Interior.Color;
        return color is DBNull ? null : color;
    }

    static void SetFill(dynamic r, object v)
    {
        if (v as string == "none") r.Interior.ColorIndex = xlNone;
        else r.Interior.Color = v;
    }

    public static void Fmt(Xl x, dynamic ws, dynamic rng, List<string> pairs, string cmd, Out o)
    {
        if (pairs.Count == 0) { ShowFormats(ws, rng, o); return; }
        var parsed = new List<(string key, Prop prop, object value)>();
        foreach (var p in pairs)
        {
            int eq = p.IndexOf('=');
            if (eq < 0 || !Props.TryGetValue(p.Substring(0, eq), out var prop))
                throw new CgrError($"'{p}': expected key=value with key one of {string.Join(", ", Props.Keys)}");
            parsed.Add((p.Substring(0, eq).ToLowerInvariant(), prop, prop.Parse(p.Substring(eq + 1))));
        }
        var entry = Snapshot(ws, rng, "format", cmd);
        var props = new Dictionary<string, object>();
        foreach (var (key, prop, value) in parsed)
        {
            props[key] = RunsToJson(FormatRuns(ws, rng, prop));
            prop.Set(rng, value);
        }
        entry["props"] = props;
        Journal.Push(x, entry);
        o.Line($"formatted {Xl.FullOf(rng)}: {string.Join(" ", pairs)}");
    }

    // `fmt RANGE` with no pairs: each property as runs of equal value.
    static void ShowFormats(dynamic ws, dynamic rng, Out o)
    {
        o.Line(Xl.FullOf(rng));
        foreach (var p in Props)
        {
            List<Run> runs = FormatRuns(ws, rng, p.Value);
            var shown = runs.Take(12).Select(r => runs.Count == 1 ? Describe(p.Key, r.Value) : $"{r.Addr}={Describe(p.Key, r.Value)}");
            o.Line($"  {p.Key,-7} {string.Join("  ", shown)}{(runs.Count > 12 ? $"  … (+{runs.Count - 12} runs)" : "")}");
        }
    }

    static string Describe(string prop, object v)
    {
        if (v == null) return "?";
        if ((prop == "color" || prop == "fill") && !(v is string))
        {
            int bgr = Convert.ToInt32(v);
            return $"#{bgr & 0xFF:X2}{(bgr >> 8) & 0xFF:X2}{(bgr >> 16) & 0xFF:X2}";
        }
        if (prop == "align")
            switch (Convert.ToInt32(v)) { case -4131: return "left"; case -4108: return "center"; case -4152: return "right"; case 1: return "general"; }
        if (v is bool b) return b ? "1" : "0";
        return Convert.ToString(v, Cells.Inv);
    }

    static List<Run> FormatRuns(dynamic ws, dynamic rng, Prop prop)
    {
        var runs = new List<Run>();
        int r = rng.Row, c = rng.Column;
        Xl.Runs(ws, r, c, r + (int)rng.Rows.Count - 1, c + (int)rng.Columns.Count - 1, prop.Get, runs);
        return Xl.Coalesce(runs);
    }

    static object[] RunsToJson(List<Run> runs) => runs.Select(r => (object)new object[] { r.Addr, r.Value }).ToArray();

    // Whole rows ("5:7") or whole columns ("C:D").
    static string RowsOrCols(dynamic rng)
    {
        if ((int)rng.Columns.Count == Cells.MaxCols) return "rows";
        if ((int)rng.Rows.Count == Cells.MaxRows) return "cols";
        throw new CgrError("insert/delete work on whole rows (Sheet!5:7) or columns (Sheet!C:D)");
    }

    public static void Insert(Xl x, dynamic ws, dynamic rng, string cmd, Out o)
    {
        string kind = RowsOrCols(rng), where = Xl.FullOf(rng);   // rng moves with the insert
        var entry = Snapshot(ws, rng, "insert", cmd);
        if (kind == "rows") rng.EntireRow.Insert(); else rng.EntireColumn.Insert();
        Journal.Push(x, entry);
        o.Line($"inserted {kind} {where}");
    }

    public static void Delete(Xl x, dynamic ws, dynamic rng, string cmd, Out o)
    {
        string kind = RowsOrCols(rng), where = Xl.FullOf(rng);   // rng is invalid after the delete
        var entry = Snapshot(ws, rng, "delete", cmd);
        dynamic content = x.Clip(ws, rng);
        if (content != null)
        {
            entry["contentAddr"] = Xl.AddrOf(content);
            entry["before"] = Jagged(Xl.Read(content, false).F);
        }
        if (kind == "rows") rng.EntireRow.Delete(); else rng.EntireColumn.Delete();
        Journal.Push(x, entry);
        o.Line($"deleted {kind} {where}  (undo restores contents, not formats or references broken by the delete)");
    }

    public static void AddSheet(Xl x, string name, string after, string cmd, Out o)
    {
        dynamic anchor = after != null ? x.Sheet(after) : x.Book.Worksheets[(int)x.Book.Worksheets.Count];
        dynamic ws = x.Book.Worksheets.Add(After: anchor);
        ws.Name = name;
        Journal.Push(x, new Dictionary<string, object> { ["t"] = DateTime.Now.ToString("s"), ["cmd"] = cmd, ["kind"] = "addsheet", ["sheet"] = name });
        o.Line($"added sheet {Cells.Sheet(name)} after {Cells.Sheet((string)anchor.Name)}");
    }

    static Dictionary<string, object> Snapshot(dynamic ws, dynamic rng, string kind, string cmd)
    {
        var e = new Dictionary<string, object>
        {
            ["t"] = DateTime.Now.ToString("s"),
            ["cmd"] = cmd,
            ["kind"] = kind,
            ["sheet"] = (string)ws.Name,
            ["addr"] = Xl.AddrOf(rng),
        };
        if (kind == "content") e["before"] = Jagged(Xl.Read(rng, false).F);
        return e;
    }

    static object[] Jagged(object[,] g)
    {
        var rows = new object[g.GetLength(0)];
        for (int i = 0; i < rows.Length; i++)
        {
            var row = new object[g.GetLength(1)];
            for (int j = 0; j < row.Length; j++) row[j] = g[i, j] ?? "";
            rows[i] = row;
        }
        return rows;
    }

    static object[,] Rect(object json)
    {
        var rows = ((IEnumerable)json).Cast<object>().Select(r => ((IEnumerable)r).Cast<object>().ToArray()).ToArray();
        var g = new object[rows.Length, rows.Length == 0 ? 0 : rows[0].Length];
        for (int i = 0; i < rows.Length; i++)
            for (int j = 0; j < rows[i].Length; j++) g[i, j] = Journal.FromJson(rows[i][j]);
        return g;
    }

    static bool Same(object[,] a, object[,] b)
    {
        if (a.GetLength(0) != b.GetLength(0) || a.GetLength(1) != b.GetLength(1)) return false;
        for (int i = 0; i < a.GetLength(0); i++)
            for (int j = 0; j < a.GetLength(1); j++)
                if ((a[i, j] as string ?? "") != (b[i, j] as string ?? "")) return false;
        return true;
    }

    public static void Undo(Xl x, int count, bool force, Out o)
    {
        var entries = Journal.Load(x);
        if (entries.Count == 0) { o.Line("nothing to undo"); return; }
        for (int k = 0; k < count && entries.Count > 0; k++)
        {
            var e = entries[entries.Count - 1];
            string kind = (string)e["kind"], sheet = (string)e["sheet"];
            dynamic ws = x.Sheet(sheet);
            switch (kind)
            {
                case "content":
                {
                    dynamic rng = ws.Range[(string)e["addr"]];
                    if (!force && !Same(Xl.Read(rng, false).F, Rect(e["after"])))
                        throw new CgrError($"{Cells.Sheet(sheet)}!{e["addr"]} changed since `{e["cmd"]}`; `cgr undo --force` overwrites it");
                    Xl.SetFormulas(rng, Rect(e["before"]));
                    RestoreRuns(ws, e["fmt"], Props["numfmt"]);
                    break;
                }
                case "format":
                    foreach (var p in (Dictionary<string, object>)e["props"]) RestoreRuns(ws, p.Value, Props[p.Key]);
                    break;
                case "insert":
                {
                    dynamic rng = ws.Range[(string)e["addr"]];
                    if ((int)rng.Columns.Count == Cells.MaxCols) rng.EntireRow.Delete(); else rng.EntireColumn.Delete();
                    break;
                }
                case "delete":
                {
                    dynamic rng = ws.Range[(string)e["addr"]];
                    if ((int)rng.Columns.Count == Cells.MaxCols) rng.EntireRow.Insert(); else rng.EntireColumn.Insert();
                    if (e.ContainsKey("contentAddr")) Xl.SetFormulas(ws.Range[(string)e["contentAddr"]], Rect(e["before"]));
                    break;
                }
                case "addsheet":
                {
                    bool alerts = x.App.DisplayAlerts;
                    x.App.DisplayAlerts = false;
                    try { ws.Delete(); } finally { x.App.DisplayAlerts = alerts; }
                    break;
                }
                default:
                    Objects.Undo(x, ws, kind, e);
                    break;
            }
            entries.RemoveAt(entries.Count - 1);
            Journal.Save(x, entries);
            o.Line($"undid: {e["cmd"]}");
        }
    }

    static void RestoreRuns(dynamic ws, object runs, Prop prop)
    {
        foreach (var run in ((IEnumerable)runs).Cast<object>())
        {
            var pair = ((IEnumerable)run).Cast<object>().ToArray();
            object v = Journal.FromJson(pair[1]);
            if (v != null) prop.Set(ws.Range[(string)pair[0]], v);
        }
    }

    public static void Log(Xl x, Out o)
    {
        var entries = Journal.Load(x);
        if (entries.Count == 0) { o.Line("(no journaled writes for this workbook)"); return; }
        for (int i = 0; i < entries.Count; i++)
            o.Line($"{entries.Count - i,3}  {entries[i]["t"]}  {entries[i]["cmd"]}");
    }

    // whatif IN=V[,V...] ... -- OUT ...: set inputs, recalculate, read outputs, put inputs back.
    public static void WhatIf(Xl x, List<string> args, Out o)
    {
        int sep = args.IndexOf("--");
        if (sep < 1 || sep == args.Count - 1) throw new CgrError("usage: whatif IN=V[,V...] ... -- OUT ...");
        var inputs = new List<(dynamic rng, string label, string[] values, object original)>();
        foreach (var a in args.Take(sep))
        {
            int eq = a.IndexOf('=');
            if (eq < 1) throw new CgrError($"'{a}': inputs look like Sheet!A1=value or Sheet!A1=v1,v2,v3");
            var (ws, rng) = x.Resolve(a.Substring(0, eq));
            inputs.Add((rng, Xl.FullOf(rng), a.Substring(eq + 1).Split(','), Xl.GetFormulas(rng)));
        }
        var sweeps = inputs.Where(i => i.values.Length > 1).ToList();
        if (sweeps.Count > 1) throw new CgrError("only one input can take several values");
        int scenarios = sweeps.Count == 1 ? sweeps[0].values.Length : 1;

        var outputs = new List<(string name, string label, dynamic cell, string fmt)>();
        foreach (var a in args.Skip(sep + 1))
        {
            var (ws, rng) = x.Resolve(a);
            Block u = x.Used((string)ws.Name);
            foreach (dynamic cell in rng.Cells)
            {
                if (outputs.Count >= 40) throw new CgrError("at most 40 output cells");
                int r = cell.Row, c = cell.Column;
                outputs.Add((Cells.Full((string)ws.Name, r, c), Look.Label(u, r, c).Trim(), cell, cell.NumberFormat as string));
            }
        }

        var baseline = outputs.Select(p => (object)p.cell.Value2).ToList();
        var results = new List<List<object>>();
        bool screen = x.App.ScreenUpdating;
        x.App.ScreenUpdating = false;
        try
        {
            for (int s = 0; s < scenarios; s++)
            {
                foreach (var inp in inputs) Xl.SetFormulas(inp.rng, inp.values.Length > 1 ? inp.values[s] : inp.values[0]);
                x.RecalcIfManual();
                results.Add(outputs.Select(p => (object)p.cell.Value2).ToList());
            }
        }
        finally
        {
            foreach (var inp in inputs) Xl.SetFormulas(inp.rng, inp.original);
            x.RecalcIfManual();
            x.App.ScreenUpdating = screen;
        }

        foreach (var inp in inputs.Where(i => i.values.Length == 1))
            o.Line($"with {inp.label} = {inp.values[0]}");
        if (scenarios == 1)
        {
            foreach (var (p, k) in outputs.Select((p, k) => (p, k)))
            {
                string delta = baseline[k] is double b0 && results[0][k] is double b1 ? "  Δ " + Cells.Formatted(b1 - b0, p.fmt) : "";
                o.Line($"{p.name,-22} {Cells.Show(baseline[k], p.fmt),14} → {Cells.Show(results[0][k], p.fmt),-14}{delta}  {p.label}");
            }
        }
        else
        {
            o.Line("outputs:");
            foreach (var (p, k) in outputs.Select((p, k) => (p, k))) o.Line($"  [{k + 1}] {p.name}  {p.label}  (now {Cells.Show(baseline[k], p.fmt)})");
            o.Line(sweeps[0].label + "\t" + string.Join("\t", outputs.Select((p, k) => $"[{k + 1}]")));
            for (int s = 0; s < scenarios; s++)
                o.Line(sweeps[0].values[s] + "\t" + string.Join("\t", results[s].Select((v, k) => Cells.Show(v, outputs[k].fmt))));
        }
    }
}

// One JSON-lines file per workbook under %LOCALAPPDATA%\comgrid\journal.
static class Journal
{
    static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

    // %LOCALAPPDATA%\comgrid\<kind>\<workbook name>-<hash of full path><ext>
    public static string FileFor(Xl x, string kind, string ext)
    {
        string full = x.Book.FullName;
        string dir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "comgrid", kind);
        Directory.CreateDirectory(dir);
        string hash;
        using (var sha = SHA1.Create())
            hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(full.ToLowerInvariant()))).Replace("-", "").Substring(0, 8);
        string safe = string.Concat(System.IO.Path.GetFileName(full).Select(ch => System.IO.Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch));
        return System.IO.Path.Combine(dir, $"{safe}-{hash}{ext}");
    }

    static string PathFor(Xl x) => FileFor(x, "journal", ".jsonl");

    public static void Push(Xl x, Dictionary<string, object> entry) =>
        File.AppendAllText(PathFor(x), Json.Serialize(entry) + "\n", new UTF8Encoding(false));

    public static List<Dictionary<string, object>> Load(Xl x)
    {
        string p = PathFor(x);
        if (!File.Exists(p)) return new List<Dictionary<string, object>>();
        return File.ReadAllLines(p, Encoding.UTF8).Where(l => l.Trim().Length > 0)
            .Select(l => Json.Deserialize<Dictionary<string, object>>(l)).ToList();
    }

    public static void Save(Xl x, List<Dictionary<string, object>> entries) =>
        File.WriteAllText(PathFor(x), string.Concat(entries.Select(e => Json.Serialize(e) + "\n")), new UTF8Encoding(false));

    // The serializer reads numbers back as decimal, which COM would pass as VT_DECIMAL.
    public static object FromJson(object v) => v is decimal d ? (object)(double)d : v;
}
