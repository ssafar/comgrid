using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

// "What did you change since I last looked?": a per-workbook snapshot of every formula
// (and constant), diffed by `cgr changes`. cgr's own writes must not show up as the
// user's edits, so around each write the user's edits so far are set aside as text and
// the snapshot is re-taken afterwards. That stays right even when an inserted row makes
// Excel rewrite references all over the workbook.
static class Snapshot
{
    class Snap
    {
        public string Time;
        public List<string> Pending = new List<string>();   // user edits set aside around cgr writes
        public Dictionary<string, Dictionary<string, string[]>> Sheets =
            new Dictionary<string, Dictionary<string, string[]>>(StringComparer.OrdinalIgnoreCase);   // sheet -> "r,c" -> [A1, R1C1]
        public List<string> Order = new List<string>();
    }

    static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

    static string PathFor(Xl x) => Journal.FileFor(x, "snapshots", ".json");

    static Snap Load(Xl x)
    {
        string p = PathFor(x);
        if (!File.Exists(p)) return null;
        var d = Json.Deserialize<Dictionary<string, object>>(File.ReadAllText(p, Encoding.UTF8));
        var s = new Snap { Time = (string)d["time"] };
        s.Pending = ((IEnumerable)d["pending"]).Cast<string>().ToList();
        foreach (var sheet in (Dictionary<string, object>)d["sheets"])
        {
            s.Order.Add(sheet.Key);
            s.Sheets[sheet.Key] = ((Dictionary<string, object>)sheet.Value)
                .ToDictionary(kv => kv.Key, kv => ((IEnumerable)kv.Value).Cast<string>().ToArray());
        }
        return s;
    }

    static void Save(Xl x, Snap s)
    {
        var sheets = new Dictionary<string, object>();
        foreach (var name in s.Order) sheets[name] = s.Sheets[name];
        var d = new Dictionary<string, object> { ["time"] = s.Time, ["pending"] = s.Pending, ["sheets"] = sheets };
        File.WriteAllText(PathFor(x), Json.Serialize(d), new UTF8Encoding(false));
    }

    // The live workbook, plus the blocks read (for current values).
    static Snap Take(Xl x, Dictionary<string, Block> blocks)
    {
        var s = new Snap { Time = DateTime.Now.ToString("s") };
        foreach (dynamic ws in x.Worksheets())
        {
            string name = ws.Name;
            Block b = Xl.Read(ws.UsedRange);
            blocks[name] = b;
            var cells = new Dictionary<string, string[]>();
            for (int i = 0; i < b.Rows; i++)
                for (int j = 0; j < b.Cols; j++)
                    if (!b.IsEmpty(i, j)) cells[(b.Row + i) + "," + (b.Col + j)] = new[] { b.Formula(i, j), b.R[i, j] as string ?? "" };
            s.Order.Add(name);
            s.Sheets[name] = cells;
        }
        return s;
    }

    public static void Changes(Xl x, bool reset, Out o)
    {
        Snap old = Load(x);
        var blocks = new Dictionary<string, Block>(StringComparer.OrdinalIgnoreCase);
        Snap now = Take(x, blocks);
        Save(x, now);
        if (old == null || reset)
        {
            o.Line($"baseline taken: {now.Sheets.Values.Sum(c => c.Count)} cells in {now.Order.Count} sheets; `cgr changes` reports edits from here on");
            return;
        }
        var lines = old.Pending.Concat(Diff(old, now, blocks)).ToList();
        if (lines.Count == 0) { o.Line($"no changes since {old.Time}"); return; }
        o.Line($"changes since {old.Time}:");
        foreach (var l in lines) o.Line(l);
    }

    // Around a cgr write: set the user's edits aside, then re-take after the write.
    public static IDisposable Around(Xl x)
    {
        Snap old = Load(x);
        if (old == null) return null;
        var blocks = new Dictionary<string, Block>(StringComparer.OrdinalIgnoreCase);
        var pending = old.Pending.Concat(Diff(old, Take(x, blocks), blocks)).ToList();
        return new After(() =>
        {
            Snap now = Take(x, new Dictionary<string, Block>(StringComparer.OrdinalIgnoreCase));
            now.Time = old.Time;
            now.Pending = pending;
            Save(x, now);
        });
    }

    class After : IDisposable
    {
        readonly Action action;
        public After(Action action) { this.action = action; }
        public void Dispose() => action();
    }

    static List<string> Diff(Snap old, Snap now, Dictionary<string, Block> blocks)
    {
        var lines = new List<string>();
        foreach (var name in now.Order)
        {
            var cur = now.Sheets[name];
            bool added = !old.Sheets.TryGetValue(name, out var was);
            if (added) was = new Dictionary<string, string[]>();
            var changed = cur.Keys.Union(was.Keys)
                .Where(k => !was.TryGetValue(k, out var a) | !cur.TryGetValue(k, out var b) || a[0] != b[0])
                .Select(k => { var rc = k.Split(','); return (r: int.Parse(rc[0]), c: int.Parse(rc[1]), key: k); })
                .OrderBy(p => p.c).ThenBy(p => p.r).ToList();
            if (added) lines.Add($"+ new sheet {Cells.Sheet(name)} ({cur.Count} cells)");
            if (changed.Count == 0) continue;
            if (!added) lines.Add($"{Cells.Sheet(name)}: {changed.Count} cell{(changed.Count == 1 ? "" : "s")}");

            // Down a column, cells whose old and new formulas are each "the same" in R1C1
            // (a fill, or a cleared block) print as one run.
            string[] Get(Dictionary<string, string[]> d, string k) => d.TryGetValue(k, out var v) ? v : new[] { "", "" };
            var runs = new List<(int r1, int r2, int c, string key)>();
            foreach (var p in changed)
            {
                var last = runs.Count > 0 ? runs[runs.Count - 1] : default;
                if (runs.Count > 0 && last.c == p.c && last.r2 == p.r - 1
                    && Get(was, last.key)[1] == Get(was, p.key)[1] && Get(cur, last.key)[1] == Get(cur, p.key)[1])
                    runs[runs.Count - 1] = (last.r1, p.r, p.c, last.key);
                else runs.Add((p.r, p.r, p.c, p.key));
            }
            blocks.TryGetValue(name, out Block blk);
            foreach (var run in runs.OrderBy(q => q.r1).ThenBy(q => q.c))
            {
                string before = Get(was, run.key)[0], after = Get(cur, run.key)[0];
                int n = run.r2 - run.r1 + 1;
                string value = "";
                if (after.StartsWith("=") && blk != null && blk.Contains(run.r1, run.c)) value = " → " + Cells.Raw(blk.V[run.r1 - blk.Row, run.c - blk.Col]);
                lines.Add($"  {Cells.Addr(run.r1, run.c, run.r2, run.c),-10}{(n > 1 ? $" ×{n}" : "")} {Show(before)}  ⇒  {Show(after)}{value}");
            }
        }
        foreach (var name in old.Order.Where(n => !now.Sheets.ContainsKey(n)))
            lines.Add($"- sheet {Cells.Sheet(name)} removed ({old.Sheets[name].Count} cells)");
        return lines;
    }

    static string Show(string formula) => formula.Length == 0 ? "∅" : Cells.Trunc(formula, 100);
}
