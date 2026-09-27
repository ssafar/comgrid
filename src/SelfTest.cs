using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// `cgr selftest`: checks for the parts that don't need Excel, so CI can run them.
static class SelfTest
{
    static int failures;

    static void Eq(string what, object got, object want)
    {
        if (Equals(got, want)) return;
        failures++;
        Console.WriteLine($"FAIL {what}\n  got:  {got}\n  want: {want}");
    }

    public static int Run()
    {
        Eq("Col 1", Cells.Col(1), "A");
        Eq("Col 28", Cells.Col(28), "AB");
        Eq("Col 16384", Cells.Col(16384), "XFD");
        Eq("ColNum $AI", Cells.ColNum("$AI"), 35);
        Eq("Addr range", Cells.Addr(5, 2, 7, 3), "B5:C7");
        Eq("Addr rows", Cells.Addr(5, 1, 7, Cells.MaxCols), "5:7");
        Eq("Sheet quoting", Cells.Sheet("Annual model"), "'Annual model'");
        Eq("Sheet cell-like name", Cells.Sheet("AB1"), "'AB1'");
        Eq("Err #N/A", Cells.Err(-2146826246), "#N/A");

        const string money = "\\$#,##0;(\\$#,##0);\"–\"";
        Eq("fmt money", Cells.Formatted(1234.5, money), "$1,235");
        Eq("fmt money negative", Cells.Formatted(-5, money), "($5)");
        Eq("fmt money zero", Cells.Formatted(0, money), "–");
        Eq("fmt percent", Cells.Formatted(0.0703, "0.00%"), "7.03%");
        Eq("fmt general", Cells.Formatted(0.07, "General"), "0.07");
        Eq("fmt date", Cells.Formatted(46291, "mm/dd/yyyy"), "2026-09-26");
        Eq("fmt plain negative", Cells.Formatted(-1.5, "0.0"), "-1.5");

        var refs = Refs.Parse("=SUM('Annual model'!T15:T65)+Assumptions!$C$6+A1+\"B2\"+LOG10(4)", "Comparison", null);
        Eq("refs count", refs.Count, 3);
        Eq("ref 1", refs.ElementAtOrDefault(0)?.ToString(), "'Annual model'!T15:T65");
        Eq("ref 2", refs.ElementAtOrDefault(1)?.ToString(), "Assumptions!C6");
        Eq("ref 3", refs.ElementAtOrDefault(2)?.ToString(), "Comparison!A1");
        Eq("whole column", Refs.Parse("=SUM(C:C)", "S", null).FirstOrDefault()?.ToString(), "S!C:C");
        var named = Refs.Parse("=Rate*A1", "S", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "rate" });
        Eq("defined name", named.Count(r => r.Name != null), 1);

        // A fixed reference written without $ still counts as "the same formula filled down".
        Refs.Shape a = Refs.ShapeOf("=$C$6*Properties!L6*(1+C16)", "M"), b = Refs.ShapeOf("=$C$6*Properties!L6*(1+C17)", "M");
        Eq("relation fixed/fixed/shift", Refs.Relation(a, b, 1, 0), "ffs");
        Eq("relation mismatch", Refs.Relation(a, Refs.ShapeOf("=$C$6*Properties!L6*(2+C17)", "M"), 1, 0), null);
        Eq("relation ignores refs in strings", Refs.ShapeOf("=\"A1\"&B2", "M").Refs.Count, 1);

        Eq("collapse", Collapse(new[,]
        {
            { "Year", "Value" },
            { "1", "=B1*2" },
            { "2", "=B2*2" },
            { "3", "=B3*2" },
        }), "A1:B1 \"Year\" | \"Value\"\nA2:A4 1, 2, … 3 (step 1, ×3)\nB2:B4 =B1*2 ×3 → 0 (all)");

        Console.WriteLine(failures == 0 ? "selftest: all passed" : $"selftest: {failures} failed");
        return failures == 0 ? 0 : 1;
    }

    // Runs the dump's collapsing over a fake block (formulas in A1 form; R1C1 derived for
    // the simple relative case), and returns its lines with spacing squeezed.
    static string Collapse(string[,] cells)
    {
        int n = cells.GetLength(0), m = cells.GetLength(1);
        var b = new Block { Sheet = "S", Row = 1, Col = 1, Rows = n, Cols = m, F = new object[n, m], R = new object[n, m], V = new object[n, m] };
        for (int i = 0; i < n; i++)
            for (int j = 0; j < m; j++)
            {
                string f = cells[i, j];
                b.F[i, j] = f;
                b.R[i, j] = f.StartsWith("=") ? System.Text.RegularExpressions.Regex.Replace(f, @"[A-Z]+\d+", "R[-1]C") : f;
                b.V[i, j] = f.StartsWith("=") ? (object)0.0 : double.TryParse(f, out double d) ? (object)d : f;
            }
        var old = Console.Out;
        var sw = new StringWriter();
        Console.SetOut(sw);
        try { Look.Collapsed(b, new Out(0)); }
        finally { Console.SetOut(old); }
        return string.Join("\n", sw.ToString().Replace("\r", "").Split('\n').Where(l => l.Length > 0)
            .Select(l => System.Text.RegularExpressions.Regex.Replace(l, " {2,}", " ")));
    }
}
