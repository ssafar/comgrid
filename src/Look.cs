using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

// Read-only commands: outline, dump, find, calc.
static class Look
{
    public static void Outline(Xl x, string only, Out o)
    {
        dynamic book = x.Book;
        bool saved = book.Saved;
        o.Line($"{book.FullName}{(saved ? "" : "  (unsaved changes)")}");
        o.Line($"calculation: {CalcMode(x)}   active: {Active(x)}");
        foreach (dynamic nm in book.Names)
        {
            bool visible = nm.Visible;
            o.Line($"name {nm.Name} {nm.RefersTo}{(visible ? "" : "  (hidden)")}");
        }
        foreach (dynamic ws in x.Worksheets())
        {
            string name = ws.Name;
            if (only != null && !string.Equals(name, only, StringComparison.OrdinalIgnoreCase)) continue;
            Block b = x.Used(name);
            int formulas = 0, constants = 0;
            var distinct = new HashSet<string>();
            for (int i = 0; i < b.Rows; i++)
                for (int j = 0; j < b.Cols; j++)
                {
                    if (b.IsFormula(i, j)) { formulas++; distinct.Add(b.R[i, j] as string); }
                    else if (!b.IsEmpty(i, j)) constants++;
                }
            int visibility = ws.Visible;
            o.Line();
            o.Line($"== {Cells.Sheet(name)}  {Cells.Addr(b.Row, b.Col, b.Row2, b.Col2)}  {formulas} formulas ({distinct.Count} distinct), {constants} constants"
                   + (visibility == -1 ? "" : visibility == 0 ? "  [hidden]" : "  [very hidden]"));
            Program.Debug("tables");
            foreach (dynamic lo in ws.ListObjects) o.Line($"  table {lo.Name}  {Xl.AddrOf(lo.Range)}");
            Program.Debug("chartobjects");
            foreach (dynamic co in ws.ChartObjects())
            {
                dynamic ch = co.Chart;
                bool hasTitle = ch.HasTitle;
                string title = hasTitle ? (string)ch.ChartTitle.Text : "(untitled)";
                // Don't touch .Rows/.Columns/.Cells on TopLeftCell: on Excel 16.0.20326 that
                // crashes Excel outright (0xc0000005). Scalar properties are fine.
                dynamic tl = co.TopLeftCell;
                o.Line($"  chart \"{co.Name}\" at {Cells.Addr((int)tl.Row, (int)tl.Column)}: {Cells.Trunc(title)}");
                Program.Debug("series");
                foreach (dynamic s in ch.SeriesCollection()) o.Line($"    {Cells.Trunc((string)s.Formula, 200)}");
            }
            Program.Debug("regions");
            foreach (var box in Regions(b)) Region(b, box, o);
        }
        foreach (dynamic cs in book.Charts) o.Line($"\n== chart sheet {cs.Name}");
    }

    static string CalcMode(Xl x)
    {
        int mode = x.App.Calculation;
        return mode == -4105 ? "automatic" : mode == Xl.xlCalculationManual ? "MANUAL (values may be stale; `cgr calc --now`)" : "semi-automatic";
    }

    static string Active(Xl x)
    {
        string s = Cells.Sheet((string)x.Book.ActiveSheet.Name);
        try
        {
            if ((string)x.App.ActiveWorkbook.FullName == (string)x.Book.FullName)
                s += ", selection " + Xl.AddrOf(x.App.Selection);
        }
        catch (Exception) { }
        return s;
    }

    // Islands of non-empty cells, joined through edges and corners like Excel's CurrentRegion.
    static List<int[]> Regions(Block b)
    {
        var seen = new bool[b.Rows, b.Cols];
        var boxes = new List<int[]>();
        var stack = new Stack<(int, int)>();
        for (int i = 0; i < b.Rows; i++)
            for (int j = 0; j < b.Cols; j++)
            {
                if (seen[i, j] || b.IsEmpty(i, j)) continue;
                var box = new[] { i, j, i, j };
                seen[i, j] = true;
                stack.Push((i, j));
                while (stack.Count > 0)
                {
                    var (ci, cj) = stack.Pop();
                    box[0] = Math.Min(box[0], ci); box[1] = Math.Min(box[1], cj);
                    box[2] = Math.Max(box[2], ci); box[3] = Math.Max(box[3], cj);
                    for (int di = -1; di <= 1; di++)
                        for (int dj = -1; dj <= 1; dj++)
                        {
                            int ni = ci + di, nj = cj + dj;
                            if (ni < 0 || nj < 0 || ni >= b.Rows || nj >= b.Cols || seen[ni, nj] || b.IsEmpty(ni, nj)) continue;
                            seen[ni, nj] = true;
                            stack.Push((ni, nj));
                        }
                }
                boxes.Add(box);
            }
        for (bool merged = true; merged;)
        {
            merged = false;
            for (int p = 0; p < boxes.Count && !merged; p++)
                for (int q = p + 1; q < boxes.Count && !merged; q++)
                {
                    int[] a = boxes[p], c = boxes[q];
                    if (a[0] > c[2] || c[0] > a[2] || a[1] > c[3] || c[1] > a[3]) continue;
                    boxes[p] = new[] { Math.Min(a[0], c[0]), Math.Min(a[1], c[1]), Math.Max(a[2], c[2]), Math.Max(a[3], c[3]) };
                    boxes.RemoveAt(q);
                    merged = true;
                }
        }
        boxes.Sort((a, c) => a[0] != c[0] ? a[0].CompareTo(c[0]) : a[1].CompareTo(c[1]));
        return boxes;
    }

    static void Region(Block b, int[] box, Out o)
    {
        int h = box[2] - box[0] + 1, w = box[3] - box[1] + 1, formulas = 0, constants = 0;
        var distinct = new HashSet<string>();
        for (int i = box[0]; i <= box[2]; i++)
            for (int j = box[1]; j <= box[3]; j++)
            {
                if (b.IsFormula(i, j)) { formulas++; distinct.Add(b.R[i, j] as string); }
                else if (!b.IsEmpty(i, j)) constants++;
            }
        string addr = Cells.Addr(b.Row + box[0], b.Col + box[1], b.Row + box[2], b.Col + box[3]);
        if ((formulas == 0 && h == 1) || h * w <= 4)
        {
            // Labels, and tiny islands like "Holding years | =Assumptions!C6 → 10", inline.
            var vals = new List<string>();
            for (int i = box[0]; i <= box[2]; i++)
                for (int j = box[1]; j <= box[3]; j++)
                    if (!b.IsEmpty(i, j))
                        vals.Add(b.IsFormula(i, j) ? $"{Cells.Trunc(b.Formula(i, j), 80)} → {Cells.Raw(b.V[i, j])}" : Cells.Raw(b.V[i, j]));
            o.Line($"  {addr,-10} {Cells.Trunc(string.Join(" | ", vals), 200)}");
            return;
        }
        o.Line($"  {addr,-10} {h}x{w}  {formulas} formulas ({distinct.Count} distinct), {constants} constants");
        var head = Enumerable.Range(box[1], w).Select(j => b.V[box[0], j] as string).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        bool hasHeader = head.Count >= 2;
        if (hasHeader) o.Line("             header: " + Cells.Trunc(string.Join(" | ", head.Select(s => Cells.Trunc(s, 30))), 240));
        var labels = RunLength(Enumerable.Range(box[0] + (hasHeader ? 1 : 0), h - (hasHeader ? 1 : 0))
            .Select(i => b.V[i, box[1]] as string).Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => Cells.Trunc(s, 40)));
        if (labels.Count >= 2)
            o.Line("             rows: " + string.Join(", ", labels.Take(10)) + (labels.Count > 10 ? $", … (+{labels.Count - 10})" : ""));
    }

    static List<string> RunLength(IEnumerable<string> items)
    {
        var list = new List<(string s, int n)>();
        foreach (var s in items)
            if (list.Count > 0 && list[list.Count - 1].s == s) list[list.Count - 1] = (s, list[list.Count - 1].n + 1);
            else list.Add((s, 1));
        return list.Select(p => p.n > 1 ? $"{p.s} ×{p.n}" : p.s).ToList();
    }

    public static void Dump(Xl x, dynamic ws, dynamic rng, bool grid, Out o)
    {
        string sheet = ws.Name;
        string asked = Xl.AddrOf(rng);
        dynamic clipped = x.Clip(ws, rng);
        if (clipped == null) { o.Line($"{Cells.Sheet(sheet)}!{asked}  (empty)"); return; }
        Block b = Xl.Read(clipped);
        o.Line($"{Cells.Full(sheet, b.Row, b.Col, b.Row2, b.Col2)}  ({b.Rows}x{b.Cols})");
        if (grid) Grid(b, o);
        else Collapsed(b, o);
    }

    // Tab-separated values as displayed (approximately), with row numbers and column letters.
    static void Grid(Block b, Out o)
    {
        string[,] fmts = Xl.NumberFormats(b);
        o.Line("\t" + string.Join("\t", Enumerable.Range(b.Col, b.Cols).Select(Cells.Col)));
        for (int i = 0; i < b.Rows; i++)
            o.Line((b.Row + i) + "\t" + string.Join("\t", Enumerable.Range(0, b.Cols).Select(j => Cells.Trunc(Cells.Show(b.V[i, j], fmts[i, j]), 40))));
    }

    class Item
    {
        public int R1, C1, R2, C2;   // block coordinates
        public string Kind, Key;     // f(ormula), c(onstant), seq(uence), row (of lone constants)
    }

    // One line per run of cells that share a formula (compared in R1C1, so filled-down
    // formulas match) or a constant, numeric sequences as "a, b, … z", and lone constants
    // in one row joined as "x | y | z".
    static void Collapsed(Block b, Out o)
    {
        int n = b.Rows, m = b.Cols;
        var key = new string[n, m];
        for (int i = 0; i < n; i++)
            for (int j = 0; j < m; j++)
            {
                string f = b.Formula(i, j);
                key[i, j] = f.Length == 0 ? null : f.StartsWith("=") ? "f" + (b.R[i, j] as string) : "c" + f;
            }

        var shapes = new Refs.Shape[n, m];
        var runs = new List<Item>();
        for (int j = 0; j < m; j++)
            for (int i = 0; i < n;)
            {
                if (key[i, j] == null) { i++; continue; }
                int len = SequenceLength(b, key, i, j);
                if (len >= 3) { runs.Add(new Item { R1 = i, C1 = j, R2 = i + len - 1, C2 = j, Kind = "seq" }); i += len; continue; }
                int k = i;
                if (key[i, j][0] == 'f')
                {
                    // Extend while each next formula is this one moved down a row, with every
                    // reference either fixed or moving, the same way all along the run. That
                    // also catches fixed references written without $ (R1C1 differs per row).
                    string modes = null;
                    while (k + 1 < n && key[k + 1, j] != null && key[k + 1, j][0] == 'f')
                    {
                        string rel = Refs.Relation(ShapeAt(b, shapes, k, j), ShapeAt(b, shapes, k + 1, j), 1, 0);
                        if (rel == null || (modes != null && rel != modes)) break;
                        modes = rel;
                        k++;
                    }
                    string runKey = "f" + string.Join("\u0002", Enumerable.Range(i, k - i + 1).Select(r => key[r, j]));
                    runs.Add(new Item { R1 = i, C1 = j, R2 = k, C2 = j, Kind = "f", Key = runKey });
                }
                else
                {
                    while (k + 1 < n && key[k + 1, j] == key[i, j]) k++;
                    runs.Add(new Item { R1 = i, C1 = j, R2 = k, C2 = j, Kind = "c", Key = key[i, j] });
                }
                i = k + 1;
            }

        // Identical runs in adjacent columns become one rectangle (formulas filled right).
        var open = new Dictionary<string, Item>();
        var merged = new List<Item>();
        foreach (var r in runs)
        {
            if (r.Kind != "seq")
            {
                string k = r.R1 + ":" + r.R2 + ":" + r.Key;
                if (open.TryGetValue(k, out var prev) && prev.C2 == r.C1 - 1) { prev.C2 = r.C1; continue; }
                open[k] = r;
            }
            merged.Add(r);
        }
        merged.Sort((p, q) => p.R1 != q.R1 ? p.R1.CompareTo(q.R1) : p.C1.CompareTo(q.C1));

        var items = new List<Item>();
        foreach (var r in merged)
        {
            bool lone = r.Kind == "c" && r.R1 == r.R2 && r.C1 == r.C2;
            var last = items.Count > 0 ? items[items.Count - 1] : null;
            if (lone && last != null && last.R1 == r.R1 && last.R2 == r.R1 && last.C2 == r.C1 - 1
                && (last.Kind == "row" || (last.Kind == "c" && last.C1 == last.C2)))
            {
                last.Kind = "row";
                last.C2 = r.C1;
                continue;
            }
            items.Add(r);
        }

        // Number formats are shown when they change from the previous line's.
        string lastFmt = "General";
        string Fmt(Item it)
        {
            string f = Format(b, it);
            if (f == null || f == lastFmt) return "";
            lastFmt = f;
            return $"  [{f}]";
        }

        if (items.Count == 0) o.Line("(empty)");
        foreach (var it in items)
        {
            string addr = Cells.Addr(b.Row + it.R1, b.Col + it.C1, b.Row + it.R2, b.Col + it.C2);
            int count = (it.R2 - it.R1 + 1) * (it.C2 - it.C1 + 1);
            object first = b.V[it.R1, it.C1], last = b.V[it.R2, it.C2];
            string times = count > 1 ? $"  ×{count}" : "";
            string text;
            switch (it.Kind)
            {
                case "f":
                    string values = count == 1 ? Cells.Raw(first) : AllSame(b, it) ? Cells.Raw(first) + " (all)" : Cells.Raw(first) + " … " + Cells.Raw(last);
                    text = Cells.Trunc(b.Formula(it.R1, it.C1), 200) + times + "  → " + values + Fmt(it);
                    break;
                case "seq":
                    object second = b.V[it.R1 + 1, it.C1];
                    text = $"{Cells.Raw(first)}, {Cells.Raw(second)}, … {Cells.Raw(last)}  (step {Cells.Num((double)second - (double)first)}, ×{count})" + Fmt(it);
                    break;
                case "row":
                    text = string.Join(" | ", Enumerable.Range(it.C1, it.C2 - it.C1 + 1).Select(j => Cells.Raw(b.V[it.R1, j])));
                    break;
                default:
                    text = Cells.Raw(first) + times + Fmt(it);
                    break;
            }
            o.Line($"{addr,-10} {text}");
        }
    }

    static Refs.Shape ShapeAt(Block b, Refs.Shape[,] shapes, int i, int j) =>
        shapes[i, j] ?? (shapes[i, j] = Refs.ShapeOf(b.Formula(i, j), b.Sheet));

    static bool AllSame(Block b, Item it)
    {
        object v0 = b.V[it.R1, it.C1];
        for (int i = it.R1; i <= it.R2; i++)
            for (int j = it.C1; j <= it.C2; j++)
                if (!Equals(b.V[i, j], v0)) return false;
        return true;
    }

    static int SequenceLength(Block b, string[,] key, int i, int j)
    {
        bool Num(int r) => r < b.Rows && key[r, j] != null && key[r, j][0] == 'c' && b.V[r, j] is double;
        if (!Num(i) || !Num(i + 1)) return 0;
        double d0 = (double)b.V[i, j], step = (double)b.V[i + 1, j] - d0;
        if (step == 0) return 0;
        int len = 2;
        while (Num(i + len) && Math.Abs((double)b.V[i + len, j] - (d0 + step * len)) <= 1e-9 * Math.Max(1, Math.Abs(d0 + step * len))) len++;
        return len;
    }

    // The number format of an item; null when it holds no numbers.
    static string Format(Block b, Item it)
    {
        bool numeric = false;
        for (int i = it.R1; i <= it.R2 && !numeric; i++)
            for (int j = it.C1; j <= it.C2 && !numeric; j++)
                numeric = b.V[i, j] is double;
        if (!numeric) return null;
        object nf = b.Ws.Range[Cells.Addr(b.Row + it.R1, b.Col + it.C1, b.Row + it.R2, b.Col + it.C2)].NumberFormat;
        return nf as string ?? "mixed formats";
    }

    // "[row label / column header]": nearest text to the left and above within the sheet.
    public static string Label(Block u, int r, int c)
    {
        if (!u.Contains(r, c)) return "";
        int i = r - u.Row, j = c - u.Col;
        string row = null, col = null;
        for (int k = j - 1; k >= 0 && row == null; k--)
            if (u.V[i, k] is string s && s.Trim().Length > 0) row = s;
        for (int k = i - 1; k >= 0 && col == null; k--)
            if (u.V[k, j] is string s && s.Trim().Length > 0) col = s;
        var parts = new[] { row, col }.Where(s => s != null).Select(s => Cells.Trunc(s, 40)).ToList();
        return parts.Count == 0 ? "" : "  [" + string.Join(" / ", parts) + "]";
    }

    public static void Find(Xl x, string text, Out o)
    {
        bool isNum = double.TryParse(text, NumberStyles.Float, Cells.Inv, out double target);
        int decimals = text.Contains(".") ? text.Length - text.IndexOf('.') - 1 : 0;
        int hits = 0;
        foreach (dynamic ws in x.Worksheets())
        {
            Block b = x.Used((string)ws.Name);
            for (int i = 0; i < b.Rows; i++)
                for (int j = 0; j < b.Cols; j++)
                {
                    string f = b.Formula(i, j);
                    if (f.Length == 0) continue;
                    object v = b.V[i, j];
                    bool match = f.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
                        || (v is string s && s.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                        || (isNum && v is double d && Math.Abs(d - target) <= 0.5 * Math.Pow(10, -decimals));
                    if (!match) continue;
                    hits++;
                    int r = b.Row + i, c = b.Col + j;
                    string formula = b.IsFormula(i, j) ? "  " + Cells.Trunc(f, 160) : "";
                    o.Line($"{Cells.Full(b.Sheet, r, c),-24} {Cells.Raw(v)}{formula}{Label(b, r, c)}");
                }
        }
        if (hits == 0) o.Line("(no matches)");
    }

    public static void Calc(Xl x, bool now, Out o)
    {
        if (now) x.App.Calculate();
        int state = x.App.CalculationState;
        o.Line($"calculation: {CalcMode(x)}; state: {(state == 0 ? "done" : state == 1 ? "calculating" : "pending")}");
    }
}
