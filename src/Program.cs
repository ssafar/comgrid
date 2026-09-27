using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

// Prints at most Max lines, then says how many it held back.
class Out
{
    readonly int max;
    int lines, dropped;

    public Out(int max) { this.max = max; }

    public void Line(string s = "")
    {
        if (max > 0 && lines >= max) { dropped++; return; }
        Console.WriteLine(s);
        lines++;
    }

    public void Finish()
    {
        if (dropped > 0) Console.WriteLine($"… {dropped} more lines (raise with --max N; 0 = no limit)");
        dropped = 0;
    }
}

class Args
{
    static readonly HashSet<string> WithValue = new HashSet<string> { "-b", "--book", "--max", "--depth", "--after", "--timeout" };
    public readonly List<string> Pos = new List<string>();
    readonly Dictionary<string, string> opts = new Dictionary<string, string>();

    public Args(string[] argv)
    {
        bool rest = false;
        for (int i = 0; i < argv.Length; i++)
        {
            string s = argv[i];
            if (rest) Pos.Add(s);
            else if (s == "--") { Pos.Add(s); rest = true; }
            else if (WithValue.Contains(s) && i + 1 < argv.Length) opts[s] = argv[++i];
            else if (s.StartsWith("--") && s.Length > 2) opts[s] = "";
            else Pos.Add(s);
        }
    }

    public bool Has(string o) => opts.ContainsKey(o);
    public string Val(string o) => opts.TryGetValue(o, out var v) ? v : null;
    public int Int(string o, int dflt) => opts.TryGetValue(o, out var v) ? int.Parse(v) : dflt;
}

static class Program
{
    const string Usage = @"cgr - read and edit workbooks open in Excel, over COM (comgrid)

usage: cgr [-b BOOK] COMMAND [ARGS]

workbooks
  books                        workbooks open in any running Excel
  open FILE [--isolated]       open in the running Excel; --isolated: in a new hidden instance
  close [--save]               close the -b workbook (default: without saving)

reading
  outline [SHEET]              sheets, cell islands with headers and row labels, names, tables, charts
  dump RANGE [--grid]          formulas collapsed into runs, with values; --grid: values as a TSV grid
  sel [--grid]                 dump whatever is selected in Excel right now
  find TEXT                    cells whose formula or value contains TEXT (numbers match by value)
  trace CELL [--depth N]       what a cell's formula reads (N levels, default 1)
  trace CELL --dependents      which formulas read the cell
  calc [--now]                 calculation mode; --now recalculates

writing (straight into the live workbook; journaled so `undo` can revert)
  set RANGE VALUE              VALUE or =FORMULA into every cell (relative refs shift, like Ctrl+Enter)
  set CELL -                   TSV from stdin, top-left at CELL; =... cells are formulas
  fmt RANGE [KEY=VALUE ...]    no pairs: show formats; numfmt=0.0% bold=1 italic=0 color=#RRGGBB fill=#RRGGBB|none
                               width=12 wrap=1 align=left|center|right|general
  insert Sheet!5:7 | Sheet!C:D insert whole rows / columns
  delete Sheet!5:7 | Sheet!C:D delete them (undo restores contents only)
  addsheet NAME [--after SHEET]
  undo [N] [--force]           revert the last N writes; refuses if the cells changed since, unless --force
  log                          journaled writes, newest last

what-if (inputs are put back afterwards)
  whatif IN=V ... -- OUT ...   e.g. whatif Assumptions!C12=0.05 -- Comparison!J7:J11
  whatif IN=V1,V2,V3 -- OUT    sweep one input; prints a table

RANGE: Sheet!A1:B2, 'Sheet name'!A1, A1 (active sheet), a defined name, Sheet!C:C.
BOOK: a workbook name, a unique part of one, or a path (opened in Excel if not open).
      Default: $CGR_BOOK, else Excel's active workbook.
options: --max N (output lines, default 400; 0 = all)  --full (don't shorten text)
         --timeout MS (how long to retry while Excel is busy, default 8000)
";

    [STAThread]
    static int Main(string[] argv)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        var a = new Args(argv);
        if (a.Has("--version")) { Console.WriteLine("cgr " + BuildInfo.Version); return 0; }
        if (a.Pos.Count > 0 && a.Pos[0] == "selftest") return SelfTest.Run();
        if (a.Pos.Count == 0 || a.Has("--help") || a.Pos[0] == "help")
        {
            Console.Write(Usage);
            return a.Pos.Count == 0 && !a.Has("--help") ? 2 : 0;
        }
        BusyFilter.Register();
        if (a.Val("--timeout") != null) BusyFilter.TimeoutMs = a.Int("--timeout", 8000);
        Cells.TextMax = a.Has("--full") ? 0 : 120;
        var o = new Out(a.Int("--max", 400));
        try
        {
            Run(a, "cgr " + string.Join(" ", argv.Select(s => s.Contains(" ") || s.Length == 0 ? "\"" + s + "\"" : s)), o);
            o.Finish();
            return 0;
        }
        catch (Exception e)
        {
            o.Finish();
            if (e is System.Reflection.TargetInvocationException && e.InnerException != null) e = e.InnerException;
            string msg = e is CgrError ? e.Message
                : e is COMException ce && BusyFilter.IsBusy(ce) ? "Excel is busy (a cell is being edited or a dialog is open); try again"
                : e is COMException ce2 ? $"Excel said: {ce2.Message} (0x{ce2.HResult:X8})"
                : e is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ? "unexpected object from Excel: " + e.Message
                : e.GetType().Name + ": " + e.Message;
            Console.Error.WriteLine("cgr: " + msg);
            return 1;
        }
    }

    static readonly bool debug = Environment.GetEnvironmentVariable("CGR_DEBUG") != null;
    static readonly System.Diagnostics.Stopwatch clock = System.Diagnostics.Stopwatch.StartNew();

    // CGR_DEBUG=1 prints step markers to stderr, flushed, so a crash shows the last call.
    public static void Debug(string step)
    {
        if (debug) { Console.Error.WriteLine($"[{clock.ElapsedMilliseconds,6} ms] {step}"); Console.Error.Flush(); }
    }

    static void Need(List<string> p, int n, string usage)
    {
        if (p.Count < n) throw new CgrError("usage: cgr " + usage);
    }

    static void Run(Args a, string cmdline, Out o)
    {
        string cmd = a.Pos[0];
        var p = a.Pos.Skip(1).ToList();
        if (cmd == "books") { Xl.Books(o); return; }
        if (cmd == "open")
        {
            Need(p, 1, "open FILE [--isolated]");
            Xl.Open(p[0], a.Has("--isolated"), o);
            return;
        }
        if (cmd == "sel")
        {
            var (sx, sws, srng) = Xl.Selection();
            Look.Dump(sx, sws, srng, a.Has("--grid"), o);
            return;
        }
        var x = Xl.Attach(a.Val("-b") ?? a.Val("--book") ?? Environment.GetEnvironmentVariable("CGR_BOOK"));
        switch (cmd)
        {
            case "outline":
                Look.Outline(x, p.FirstOrDefault(), o);
                break;
            case "dump":
            {
                Need(p, 1, "dump RANGE [--grid]");
                var (ws, rng) = x.Resolve(p[0]);
                Look.Dump(x, ws, rng, a.Has("--grid"), o);
                break;
            }
            case "find":
                Need(p, 1, "find TEXT");
                Look.Find(x, string.Join(" ", p), o);
                break;
            case "trace":
            {
                Need(p, 1, "trace CELL [--depth N] [--dependents]");
                var (ws, rng) = x.Resolve(p[0]);
                Refs.Trace(x, ws, rng, a.Int("--depth", 1), a.Has("--dependents"), o);
                break;
            }
            case "calc":
                Look.Calc(x, a.Has("--now"), o);
                break;
            case "set":
            {
                Need(p, 2, "set RANGE VALUE   |   set CELL - < block.tsv");
                var (ws, rng) = x.Resolve(p[0]);
                Edit.Set(x, ws, rng, string.Join(" ", p.Skip(1)), cmdline, o);
                break;
            }
            case "fmt":
            {
                Need(p, 1, "fmt RANGE [KEY=VALUE ...]");
                var (ws, rng) = x.Resolve(p[0]);
                Edit.Fmt(x, ws, rng, p.Skip(1).ToList(), cmdline, o);
                break;
            }
            case "insert":
            case "delete":
            {
                Need(p, 1, cmd + " Sheet!5:7 | Sheet!C:D");
                var (ws, rng) = x.Resolve(p[0]);
                if (cmd == "insert") Edit.Insert(x, ws, rng, cmdline, o); else Edit.Delete(x, ws, rng, cmdline, o);
                break;
            }
            case "addsheet":
                Need(p, 1, "addsheet NAME [--after SHEET]");
                Edit.AddSheet(x, p[0], a.Val("--after"), cmdline, o);
                break;
            case "undo":
                Edit.Undo(x, p.Count > 0 ? int.Parse(p[0]) : 1, a.Has("--force"), o);
                break;
            case "log":
                Edit.Log(x, o);
                break;
            case "close":
                x.Close(a.Has("--save"), o);
                break;
            case "whatif":
                Edit.WhatIf(x, p, o);
                break;
            default:
                throw new CgrError($"unknown command '{cmd}' (cgr --help)");
        }
    }
}
