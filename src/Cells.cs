using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

// Addresses, value rendering and a small number-format renderer.
static class Cells
{
    public const int MaxRows = 1048576, MaxCols = 16384;
    public static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    // Strings longer than this are cut with an ellipsis; 0 disables (--full).
    public static int TextMax = 120;

    public static string Col(int c)
    {
        var sb = new StringBuilder();
        for (; c > 0; c = (c - 1) / 26) sb.Insert(0, (char)('A' + (c - 1) % 26));
        return sb.ToString();
    }

    public static int ColNum(string letters)
    {
        int n = 0;
        foreach (char ch in letters.ToUpperInvariant())
            if (ch != '$') n = n * 26 + (ch - 'A' + 1);
        return n;
    }

    public static string Addr(int r, int c) => Col(c) + r.ToString(Inv);

    public static string Addr(int r1, int c1, int r2, int c2)
    {
        if (r1 == r2 && c1 == c2) return Addr(r1, c1);
        if (r1 == 1 && r2 == MaxRows) return Col(c1) + ":" + Col(c2);
        if (c1 == 1 && c2 == MaxCols) return r1 + ":" + r2;
        return Addr(r1, c1) + ":" + Addr(r2, c2);
    }

    public static string Sheet(string name) =>
        Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_.]*$") && !Regex.IsMatch(name, @"^[A-Za-z]{1,3}[0-9]+$")
            ? name : "'" + name.Replace("'", "''") + "'";

    public static string Full(string sheet, int r1, int c1, int r2, int c2) => Sheet(sheet) + "!" + Addr(r1, c1, r2, c2);
    public static string Full(string sheet, int r, int c) => Sheet(sheet) + "!" + Addr(r, c);

    // COM hands back a scalar for one cell and a 1-based 2-D array otherwise; normalize to 0-based.
    public static object[,] Grid(object o, int rows, int cols)
    {
        var g = new object[rows, cols];
        if (o is object[,] a)
        {
            int lr = a.GetLowerBound(0), lc = a.GetLowerBound(1);
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    g[i, j] = a[lr + i, lc + j];
        }
        else g[0, 0] = o;
        return g;
    }

    public static string Trunc(string s, int max)
    {
        s = s.Replace("\r\n", "⏎").Replace('\n', '⏎').Replace('\t', ' ');
        return max > 0 && s.Length > max ? s.Substring(0, max - 1) + "…" : s;
    }

    public static string Trunc(string s) => Trunc(s, TextMax);

    // Value2 reports cell errors as VT_ERROR, which arrives as an int (0x800A0000 + code).
    static readonly Dictionary<int, string> Errors = new Dictionary<int, string>
    {
        [2000] = "#NULL!", [2007] = "#DIV/0!", [2015] = "#VALUE!", [2023] = "#REF!", [2029] = "#NAME?",
        [2036] = "#NUM!", [2042] = "#N/A", [2043] = "#GETTING_DATA", [2045] = "#SPILL!", [2046] = "#CONNECT!",
        [2047] = "#BLOCKED!", [2048] = "#UNKNOWN!", [2049] = "#FIELD!", [2050] = "#CALC!",
    };

    public static string Err(int v) => Errors.TryGetValue(v + 2146828288, out var s) ? s : "#ERR" + v;

    public static string Num(double d) => d.ToString("G12", Inv);

    // Unambiguous rendering: strings quoted, numbers raw.
    public static string Raw(object v)
    {
        switch (v)
        {
            case null: return "∅";
            case string s: return "\"" + Trunc(s) + "\"";
            case double d: return Num(d);
            case bool b: return b ? "TRUE" : "FALSE";
            case int e: return Err(e);
            case DateTime t: return t.ToString("yyyy-MM-dd HH:mm:ss", Inv);
            default: return Convert.ToString(v, Inv);
        }
    }

    // Display rendering, approximating the cell's number format.
    public static string Show(object v, string fmt)
    {
        switch (v)
        {
            case null: return "";
            case string s: return Trunc(s);
            case double d: return Formatted(d, fmt);
            default: return Raw(v);
        }
    }

    public static string Formatted(double d, string fmt)
    {
        if (string.IsNullOrEmpty(fmt) || fmt.Equals("General", StringComparison.OrdinalIgnoreCase)) return d.ToString("G10", Inv);
        var sections = Sections(fmt);
        string sec = sections[0];
        bool minus = d < 0;
        if (d < 0 && sections.Count > 1) { sec = sections[1]; d = -d; minus = false; }
        else if (d == 0 && sections.Count > 2) sec = sections[2];
        sec = Regex.Replace(sec, @"\[[^\]]*\]|_.|\*.", "");                 // colors/conditions, padding, fill
        string core = Regex.Replace(sec, "\"[^\"]*\"|\\\\.", "");           // without literals
        if (Regex.IsMatch(core, "[dmyhs]", RegexOptions.IgnoreCase) && !Regex.IsMatch(core, "[0#?]"))
        {
            try { return DateTime.FromOADate(d).ToString(Regex.IsMatch(core, "[hs]", RegexOptions.IgnoreCase) ? "yyyy-MM-dd HH:mm" : "yyyy-MM-dd", Inv); }
            catch (ArgumentException) { return Num(d); }
        }
        int first = sec.IndexOfAny("0#?".ToCharArray());
        if (first < 0) return Literal(sec);                                 // e.g. a "–" zero section
        int last = sec.LastIndexOfAny("0#?".ToCharArray());
        string span = sec.Substring(first, last - first + 1);
        if (minus) d = -d;   // the sign is added back below, outside any prefix like "$"
        if (core.Contains("%")) d *= 100;
        int dot = span.IndexOf('.');
        int decimals = dot < 0 ? 0 : span.Length - dot - 1;
        string net = (span.Contains(",") ? "#,##0" : "0") + (decimals > 0 ? "." + new string('0', decimals) : "");
        string num = d.ToString(net, Inv);
        if (Regex.IsMatch(span, "[Ee][+-]")) num = d.ToString("0." + new string('0', Math.Max(decimals, 0)) + "E+00", Inv);
        return (minus ? "-" : "") + Literal(sec.Substring(0, first)) + num + Literal(sec.Substring(last + 1));
    }

    static string Literal(string s) => Regex.Replace(s, "\"([^\"]*)\"|\\\\(.)", m => m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value);

    static List<string> Sections(string fmt)
    {
        var list = new List<string>();
        var sb = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < fmt.Length; i++)
        {
            char ch = fmt[i];
            if (ch == '"') quoted = !quoted;
            if (ch == '\\' && i + 1 < fmt.Length) { sb.Append(ch).Append(fmt[++i]); continue; }
            if (ch == ';' && !quoted) { list.Add(sb.ToString()); sb.Clear(); continue; }
            sb.Append(ch);
        }
        list.Add(sb.ToString());
        return list;
    }
}
