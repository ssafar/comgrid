using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

// A reference found in a formula: a rectangle on a sheet, or a defined name.
class Ref
{
    public string Sheet, Name, Text;
    public int R1, C1, R2, C2;

    public int Count => (R2 - R1 + 1) * (C2 - C1 + 1);
    public bool Contains(string sheet, int r, int c) =>
        Name == null && string.Equals(Sheet, sheet, StringComparison.OrdinalIgnoreCase) && r >= R1 && r <= R2 && c >= C1 && c <= C2;
    public override string ToString() => Name ?? Cells.Full(Sheet, R1, C1, R2, C2);
}

// Formula reference parsing, and trace (precedents / dependents) built on it.
// Excel's own Precedents property stops at the sheet boundary, so we parse instead.
static class Refs
{
    static readonly Regex Strings = new Regex("\"(?:[^\"]|\"\")*\"");
    static readonly Regex RefRx = new Regex(
        @"(?<![\w.$'\]!])" +
        @"(?:(?<sh>'(?:[^']|'')+'|[A-Za-z_À-￿][\w.]*)!)?" +
        @"(?:(?<c1>\$?[A-Z]{1,3})(?<r1>\$?\d+)(?::(?<c2>\$?[A-Z]{1,3})(?<r2>\$?\d+))?" +
        @"|(?<cc1>\$?[A-Z]{1,3}):(?<cc2>\$?[A-Z]{1,3})" +
        @"|(?<rr1>\$?\d+):(?<rr2>\$?\d+))" +
        @"(?![\w(!.])");
    static readonly Regex Ident = new Regex(@"(?<![\w.$'\]!])[A-Za-z_\\][\w.]*(?![\w(!])");

    static int Int(string s) => int.Parse(s.TrimStart('$'), Cells.Inv);

    // A formula split into its text outside references ("skeleton") and the rectangles it
    // references, so formulas can be compared by what they point at rather than how the
    // references happen to be written.
    public class Shape
    {
        public string Skeleton;
        public List<Ref> Refs;
    }

    public static Shape ShapeOf(string formula, string ownSheet)
    {
        var sb = new StringBuilder();
        int pos = 0;
        foreach (Match s in Strings.Matches(formula))   // string literals stay as they are
        {
            sb.Append(RefRx.Replace(formula.Substring(pos, s.Index - pos), "\u0001")).Append(s.Value);
            pos = s.Index + s.Length;
        }
        sb.Append(RefRx.Replace(formula.Substring(pos), "\u0001"));
        return new Shape { Skeleton = sb.ToString(), Refs = Parse(formula, ownSheet, null) };
    }

    // How a's references relate to b's, where b sits (dr, dc) away from a: per reference,
    // 'f' if it points at the same cells, 's' if it moved with the cell. Null if neither.
    public static string Relation(Shape a, Shape b, int dr, int dc)
    {
        if (a.Skeleton != b.Skeleton || a.Refs.Count != b.Refs.Count) return null;
        var modes = new char[a.Refs.Count];
        for (int k = 0; k < modes.Length; k++)
        {
            Ref x = a.Refs[k], y = b.Refs[k];
            if (!string.Equals(x.Sheet, y.Sheet, StringComparison.OrdinalIgnoreCase)) return null;
            if (x.R1 == y.R1 && x.R2 == y.R2 && x.C1 == y.C1 && x.C2 == y.C2) modes[k] = 'f';
            else if (x.R1 + dr == y.R1 && x.R2 + dr == y.R2 && x.C1 + dc == y.C1 && x.C2 + dc == y.C2) modes[k] = 's';
            else return null;
        }
        return new string(modes);
    }

    public static List<Ref> Parse(string formula, string ownSheet, ICollection<string> names)
    {
        var list = new List<Ref>();
        if (formula == null || !formula.StartsWith("=")) return list;
        string f = Strings.Replace(formula, "\"\"");
        foreach (Match mt in RefRx.Matches(f))
        {
            string sh = ownSheet;
            if (mt.Groups["sh"].Success)
            {
                sh = mt.Groups["sh"].Value;
                if (sh.StartsWith("'")) sh = sh.Substring(1, sh.Length - 2).Replace("''", "'");
            }
            if (sh.StartsWith("[")) continue;   // another workbook
            var r = new Ref { Sheet = sh, Text = mt.Value };
            if (mt.Groups["c1"].Success)
            {
                r.C1 = Cells.ColNum(mt.Groups["c1"].Value); r.R1 = Int(mt.Groups["r1"].Value);
                r.C2 = mt.Groups["c2"].Success ? Cells.ColNum(mt.Groups["c2"].Value) : r.C1;
                r.R2 = mt.Groups["r2"].Success ? Int(mt.Groups["r2"].Value) : r.R1;
            }
            else if (mt.Groups["cc1"].Success)
            {
                r.C1 = Cells.ColNum(mt.Groups["cc1"].Value); r.C2 = Cells.ColNum(mt.Groups["cc2"].Value);
                r.R1 = 1; r.R2 = Cells.MaxRows;
            }
            else
            {
                r.R1 = Int(mt.Groups["rr1"].Value); r.R2 = Int(mt.Groups["rr2"].Value);
                r.C1 = 1; r.C2 = Cells.MaxCols;
            }
            if (r.R1 > r.R2) (r.R1, r.R2) = (r.R2, r.R1);
            if (r.C1 > r.C2) (r.C1, r.C2) = (r.C2, r.C1);
            if (r.R1 < 1 || r.R2 > Cells.MaxRows || r.C1 < 1 || r.C2 > Cells.MaxCols) continue;
            list.Add(r);
        }
        if (names != null && names.Count > 0)
            foreach (Match mt in Ident.Matches(RefRx.Replace(f, " ")))
                if (names.Contains(mt.Value)) list.Add(new Ref { Name = mt.Value, Text = mt.Value });
        return list;
    }

    class Ctx
    {
        public Xl X;
        public HashSet<string> Names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> NameRefersTo = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Seen = new HashSet<string>();
        public Out O;

        public Ctx(Xl x, Out o)
        {
            X = x; O = o;
            foreach (dynamic nm in x.Book.Names)
            {
                string name = nm.Name;
                int bang = name.IndexOf('!');   // sheet-scoped names are "Sheet!Name"
                string bare = bang >= 0 ? name.Substring(bang + 1) : name;
                Names.Add(bare);
                NameRefersTo[bare] = nm.RefersTo;
            }
        }

        public Block Used(string sheet)
        {
            try { return X.Used(sheet); } catch (CgrError) { return null; }
        }

        public string Describe(string sheet, int r, int c)
        {
            Block u = Used(sheet);
            string where = Cells.Full(sheet, r, c);
            if (u == null) return where + "  (no such sheet)";
            if (!u.Contains(r, c)) return where + "  (empty)";
            int i = r - u.Row, j = c - u.Col;
            string f = u.IsFormula(i, j) ? Cells.Trunc(u.Formula(i, j), 200) + "  → " : "";
            return $"{where}  {f}{Cells.Raw(u.V[i, j])}{Look.Label(u, r, c)}";
        }
    }

    public static void Trace(Xl x, dynamic ws, dynamic rng, int depth, bool dependents, Out o)
    {
        string sheet = ws.Name;
        int r = rng.Row, c = rng.Column;
        var ctx = new Ctx(x, o);
        o.Line(ctx.Describe(sheet, r, c));
        if (dependents) Dependents(ctx, sheet, r, c);
        else Precedents(ctx, sheet, r, c, depth, "  ");
    }

    static void Precedents(Ctx ctx, string sheet, int r, int c, int depth, string indent)
    {
        Block u = ctx.Used(sheet);
        if (u == null || !u.Contains(r, c) || !u.IsFormula(r - u.Row, c - u.Col)) return;
        var seenHere = new HashSet<string>();
        foreach (var rf in Parse(u.Formula(r - u.Row, c - u.Col), sheet, ctx.Names))
        {
            if (!seenHere.Add(rf.ToString())) continue;
            if (rf.Name != null)
            {
                ctx.O.Line($"{indent}← name {rf.Name} = {ctx.NameRefersTo[rf.Name]}");
                continue;
            }
            if (rf.Count == 1)
            {
                ctx.O.Line($"{indent}← {ctx.Describe(rf.Sheet, rf.R1, rf.C1)}");
                if (depth > 1 && ctx.Seen.Add(rf.ToString())) Precedents(ctx, rf.Sheet, rf.R1, rf.C1, depth - 1, indent + "  ");
            }
            else ctx.O.Line($"{indent}← {rf}  {Summary(ctx, rf)}");
        }
    }

    // "51 cells: 51 formulas (2 distinct) e.g. =Q15-S15; values 72990 … 889153  [header]"
    static string Summary(Ctx ctx, Ref rf)
    {
        Block u = ctx.Used(rf.Sheet);
        if (u == null) return "(no such sheet)";
        int formulas = 0, constants = 0;
        var distinct = new Dictionary<string, string>();
        object first = null, last = null;
        int r1 = Math.Max(rf.R1, u.Row), r2 = Math.Min(rf.R2, u.Row2), c1 = Math.Max(rf.C1, u.Col), c2 = Math.Min(rf.C2, u.Col2);
        for (int r = r1; r <= r2; r++)
            for (int c = c1; c <= c2; c++)
            {
                int i = r - u.Row, j = c - u.Col;
                if (u.IsEmpty(i, j)) continue;
                if (u.IsFormula(i, j))
                {
                    formulas++;
                    string key = u.R[i, j] as string;
                    if (!distinct.ContainsKey(key)) distinct[key] = u.Formula(i, j);
                }
                else constants++;
                if (first == null) first = u.V[i, j];
                last = u.V[i, j];
            }
        string size = rf.R2 == Cells.MaxRows || rf.C2 == Cells.MaxCols ? "" : $"{rf.Count} cells: ";
        string fs = formulas == 0 ? "" : $"{formulas} formulas ({distinct.Count} distinct) e.g. {Cells.Trunc(distinct.Values.First(), 100)}; ";
        string cs = constants == 0 ? "" : $"{constants} constants; ";
        string vals = first == null ? "empty" : $"values {Cells.Raw(first)} … {Cells.Raw(last)}";
        return size + fs + cs + vals + Look.Label(u, rf.R1, rf.C1);
    }

    static void Dependents(Ctx ctx, string sheet, int r, int c)
    {
        var hits = new List<(string sheet, int r, int c)>();
        foreach (dynamic ws in ctx.X.Worksheets())
        {
            string name = ws.Name;
            Block u = ctx.Used(name);
            for (int i = 0; i < u.Rows; i++)
                for (int j = 0; j < u.Cols; j++)
                    if (u.IsFormula(i, j) && Parse(u.Formula(i, j), name, null).Any(rf => rf.Contains(sheet, r, c)))
                        hits.Add((name, u.Row + i, u.Col + j));
        }
        if (hits.Count == 0) { ctx.O.Line("  (nothing reads this cell)"); return; }
        // Consecutive cells down a column print as one run. Hits are in row-major order,
        // so a column's run is interleaved with other columns; gather it explicitly.
        while (hits.Count > 0)
        {
            var h = hits[0];
            var run = hits.Where(p => p.sheet == h.sheet && p.c == h.c && p.r >= h.r).OrderBy(p => p.r).ToList();
            int len = 1;
            while (len < run.Count && run[len].r == run[len - 1].r + 1) len++;
            var members = new HashSet<(string, int, int)>(run.Take(len));
            hits.RemoveAll(members.Contains);
            if (len == 1) ctx.O.Line("  → " + ctx.Describe(h.sheet, h.r, h.c));
            else
            {
                Block u = ctx.Used(h.sheet);
                string f = u.Formula(h.r - u.Row, h.c - u.Col);
                ctx.O.Line($"  → {Cells.Full(h.sheet, h.r, h.c, h.r + len - 1, h.c)}  ×{len}  e.g. {Cells.Trunc(f, 160)}{Look.Label(u, h.r, h.c)}");
            }
        }
    }
}
