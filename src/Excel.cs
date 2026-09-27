using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;

class CgrError : Exception
{
    public CgrError(string message) : base(message) { }
}

// Excel rejects COM calls while the user is editing a cell or has a dialog open.
// Registering a message filter makes COM retry those calls instead of failing at once.
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

class BusyFilter : IOleMessageFilter
{
    public static int TimeoutMs = 8000;

    [DllImport("ole32.dll")]
    static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    public static void Register() => CoRegisterMessageFilter(new BusyFilter(), out _);

    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) => 0;
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) => dwTickCount < TimeoutMs ? 150 : -1;
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) => 2;

    public static bool IsBusy(COMException e) => (uint)e.HResult == 0x80010001 || (uint)e.HResult == 0x8001010A;
}

// A rectangle of cells read in bulk: one COM call per property for the whole block.
class Block
{
    public string Sheet;
    public dynamic Ws;
    public int Row, Col, Rows, Cols;
    public object[,] F, R, V;   // Formula2 (A1), Formula2R1C1, Value2; 0-based

    public int Row2 => Row + Rows - 1;
    public int Col2 => Col + Cols - 1;
    public string Formula(int i, int j) => F[i, j] as string ?? "";
    public bool IsFormula(int i, int j) => Formula(i, j).StartsWith("=");
    public bool IsEmpty(int i, int j) => Formula(i, j).Length == 0;
    public bool Contains(int r, int c) => r >= Row && r <= Row2 && c >= Col && c <= Col2;
}

struct Run
{
    public int R1, C1, R2, C2;
    public object Value;
    public string Addr => Cells.Addr(R1, C1, R2, C2);
}

class Xl
{
    public dynamic App, Book;
    readonly Dictionary<string, Block> used = new Dictionary<string, Block>(StringComparer.OrdinalIgnoreCase);

    [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
    [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx ctx);

    static dynamic ActiveApp()
    {
        try { return Marshal.GetActiveObject("Excel.Application"); }
        catch (COMException) { return null; }
    }

    // The active instance's workbooks, plus file-backed workbooks in the running object
    // table, which is how workbooks in other Excel instances are reachable.
    public static List<dynamic> AllBooks(dynamic app)
    {
        var books = new List<dynamic>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (app != null)
            foreach (dynamic b in app.Workbooks)
                if (seen.Add((string)b.FullName)) books.Add(b);
        GetRunningObjectTable(0, out var rot);
        CreateBindCtx(0, out var ctx);
        rot.EnumRunning(out var e);
        var m = new IMoniker[1];
        while (e.Next(1, m, IntPtr.Zero) == 0)
        {
            try
            {
                m[0].GetDisplayName(ctx, null, out string name);
                if (!Regex.IsMatch(name, @"\.(xls[xmb]?|csv)$", RegexOptions.IgnoreCase)) continue;
                rot.GetObject(m[0], out object o);
                dynamic b = o;
                if (seen.Add((string)b.FullName)) books.Add(b);
            }
            catch (COMException) { }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
        }
        return books;
    }

    public static void Books(Out o)
    {
        dynamic app = ActiveApp();
        string active = null;
        try { if (app != null && app.ActiveWorkbook != null) active = app.ActiveWorkbook.FullName; } catch (COMException) { }
        List<dynamic> books = AllBooks(app);
        if (books.Count == 0) { o.Line("(no workbooks open)"); return; }
        foreach (dynamic b in books)
        {
            string full = b.FullName;
            bool saved = b.Saved;
            o.Line($"{(full == active ? "*" : " ")} {b.Name}  {(full == (string)b.Name ? "(never saved)" : full)}{(saved ? "" : "  (unsaved changes)")}");
        }
    }

    public static Xl Attach(string arg)
    {
        dynamic app = ActiveApp();
        if (string.IsNullOrEmpty(arg))
        {
            if (app == null) throw new CgrError("no Excel running (pass -b FILE to open one)");
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) throw new CgrError("Excel has no workbook open");
            return new Xl { App = app, Book = wb };
        }
        List<dynamic> books = AllBooks(app);
        string full = null;
        try { full = Path.GetFullPath(arg); } catch (Exception) { }
        dynamic hit = books.FirstOrDefault(b => string.Equals((string)b.FullName, full, StringComparison.OrdinalIgnoreCase))
                   ?? books.FirstOrDefault(b => string.Equals((string)b.Name, arg, StringComparison.OrdinalIgnoreCase));
        if (hit == null)
        {
            var partial = books.Where(b => ((string)b.Name).IndexOf(arg, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            if (partial.Count == 1) hit = partial[0];
            else if (partial.Count > 1) throw new CgrError($"'{arg}' matches several workbooks: " + string.Join(", ", partial.Select(b => (string)b.Name)));
        }
        if (hit != null) return new Xl { App = hit.Application, Book = hit };
        if (full != null && File.Exists(full))
        {
            if (app == null)
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = true;
                app.UserControl = true;   // keeps Excel alive after cgr exits
            }
            dynamic wb = app.Workbooks.Open(full);
            Console.Error.WriteLine($"cgr: opened {full} in Excel");
            return new Xl { App = app, Book = wb };
        }
        throw new CgrError($"no open workbook matches '{arg}'" + (books.Count > 0 ? "; open: " + string.Join(", ", books.Select(b => (string)b.Name)) : ""));
    }

    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

    public static uint Pid(dynamic app)
    {
        GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out uint pid);
        return pid;
    }

    // Opens a file in the running Excel, or with isolated=true in a new hidden instance
    // of its own (reachable afterwards with -b, since workbooks register in the ROT).
    public static void Open(string path, bool isolated, Out o)
    {
        string full = Path.GetFullPath(path);
        if (!File.Exists(full)) throw new CgrError($"no such file: {full}");
        dynamic app = isolated ? null : ActiveApp();
        if (app == null)
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            app.Visible = !isolated;
            app.DisplayAlerts = !isolated;
            app.UserControl = true;   // keeps Excel alive after cgr exits
        }
        dynamic wb = app.Workbooks.Open(full);
        o.Line($"opened {wb.Name} in {(isolated ? "a new hidden" : "the running")} Excel (pid {Pid(app)})");
    }

    // Closes a workbook without saving; an instance left hidden and empty is shut down.
    public void Close(bool save, Out o)
    {
        dynamic app = App;
        string name = Book.Name;
        Book.Close(save);
        bool gone = false;
        if (!(bool)app.Visible && (int)app.Workbooks.Count == 0) { app.Quit(); gone = true; }
        o.Line($"closed {name}{(save ? " (saved)" : " without saving")}{(gone ? "; its hidden Excel quit" : "")}");
    }

    // Excel's current selection, wherever it is.
    public static (Xl, dynamic ws, dynamic rng) Selection()
    {
        dynamic app = ActiveApp() ?? throw new CgrError("no Excel running");
        dynamic sel = app.Selection;
        dynamic ws;
        try { ws = sel.Worksheet; } catch (Exception) { throw new CgrError("the selection isn't a range of cells"); }
        if (sel.Areas.Count > 1) sel = sel.Areas.Item(1);
        return (new Xl { App = app, Book = ws.Parent }, ws, sel);
    }

    public dynamic Sheet(string name)
    {
        try { return Book.Worksheets[name]; }
        catch (COMException) { throw new CgrError($"no worksheet named '{name}' in {Book.Name}"); }
    }

    // "Sheet!A1:B2", "'Sheet name'!A1", "A1" (active sheet) or a defined name.
    public (dynamic ws, dynamic rng) Resolve(string reference)
    {
        string sheet = null, addr = reference;
        int bang = reference.LastIndexOf('!');
        if (bang >= 0)
        {
            sheet = reference.Substring(0, bang);
            addr = reference.Substring(bang + 1);
            if (sheet.Length > 1 && sheet.StartsWith("'") && sheet.EndsWith("'")) sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
        }
        dynamic ws = sheet != null ? Sheet(sheet) : Book.ActiveSheet;
        dynamic rng = null;
        try { rng = ws.Range[addr]; }
        catch (COMException)
        {
            if (sheet == null)
                try { rng = Book.Names.Item(addr).RefersToRange; ws = rng.Worksheet; } catch (COMException) { }
        }
        if (rng == null) throw new CgrError($"can't resolve '{reference}'");
        if (rng.Areas.Count > 1)
        {
            Console.Error.WriteLine($"cgr: '{reference}' has several areas; using the first");
            rng = rng.Areas.Item(1);
        }
        return (ws, rng);
    }

    public static string AddrOf(dynamic rng) =>
        Cells.Addr((int)rng.Row, (int)rng.Column, (int)rng.Row + (int)rng.Rows.Count - 1, (int)rng.Column + (int)rng.Columns.Count - 1);

    public static string FullOf(dynamic rng) => Cells.Sheet((string)rng.Worksheet.Name) + "!" + AddrOf(rng);

    // The part of rng that has anything in it (whole columns stay cheap).
    public dynamic Clip(dynamic ws, dynamic rng) => App.Intersect(rng, ws.UsedRange);

    public Block Used(string sheet)
    {
        if (!used.TryGetValue(sheet, out var b))
        {
            dynamic ws = Sheet(sheet);
            used[sheet] = b = Read(ws.UsedRange);
        }
        return b;
    }

    public List<dynamic> Worksheets()
    {
        var list = new List<dynamic>();
        foreach (dynamic ws in Book.Worksheets) list.Add(ws);
        return list;
    }

    static bool? hasFormula2;   // Formula2 is Excel 365's dynamic-array-aware formula property

    public static object GetFormulas(dynamic rng, bool r1c1 = false)
    {
        if (hasFormula2 != false)
        {
            try { object x = r1c1 ? rng.Formula2R1C1 : rng.Formula2; hasFormula2 = true; return x; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { hasFormula2 = false; }
            catch (COMException) when (hasFormula2 == null) { hasFormula2 = false; }
        }
        return r1c1 ? rng.FormulaR1C1 : rng.Formula;
    }

    public static void SetFormulas(dynamic rng, object value)
    {
        if (hasFormula2 != false)
        {
            try { rng.Formula2 = value; hasFormula2 = true; return; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { hasFormula2 = false; }
        }
        rng.Formula = value;
    }

    public static Block Read(dynamic rng, bool r1c1 = true)
    {
        int rows = rng.Rows.Count, cols = rng.Columns.Count;
        var b = new Block { Ws = rng.Worksheet, Sheet = rng.Worksheet.Name, Row = rng.Row, Col = rng.Column, Rows = rows, Cols = cols };
        b.F = Cells.Grid(GetFormulas(rng), rows, cols);
        b.V = Cells.Grid(rng.Value2, rows, cols);
        if (r1c1) b.R = Cells.Grid(GetFormulas(rng, true), rows, cols);
        return b;
    }

    // Reads a range-level property, splitting the range until each piece has a single value.
    // Excel returns null for a property that differs across the range, so this costs one
    // call per uniform run rather than one per cell.
    public static void Runs(dynamic ws, int r1, int c1, int r2, int c2, Func<dynamic, object> get, List<Run> acc)
    {
        object v = get(ws.Range[Cells.Addr(r1, c1, r2, c2)]);
        bool mixed = v == null || v is DBNull;
        if (!mixed || (r1 == r2 && c1 == c2))
        {
            acc.Add(new Run { R1 = r1, C1 = c1, R2 = r2, C2 = c2, Value = v is DBNull ? null : v });
            return;
        }
        if (r2 - r1 >= c2 - c1)
        {
            int mid = (r1 + r2) / 2;
            Runs(ws, r1, c1, mid, c2, get, acc);
            Runs(ws, mid + 1, c1, r2, c2, get, acc);
        }
        else
        {
            int mid = (c1 + c2) / 2;
            Runs(ws, r1, c1, r2, mid, get, acc);
            Runs(ws, r1, mid + 1, r2, c2, get, acc);
        }
    }

    // Joins neighbouring runs with equal values that together form a rectangle
    // (undoing the fragmentation that the halving in Runs leaves behind).
    public static List<Run> Coalesce(List<Run> runs)
    {
        var list = new List<Run>(runs);
        for (bool merged = true; merged;)
        {
            merged = false;
            for (int p = 0; p < list.Count && !merged; p++)
                for (int q = 0; q < list.Count && !merged; q++)
                {
                    Run a = list[p], b = list[q];
                    if (p == q || !Equals(a.Value, b.Value)) continue;
                    bool below = a.C1 == b.C1 && a.C2 == b.C2 && b.R1 == a.R2 + 1;
                    bool right = a.R1 == b.R1 && a.R2 == b.R2 && b.C1 == a.C2 + 1;
                    if (!below && !right) continue;
                    list[p] = new Run { R1 = a.R1, C1 = a.C1, R2 = b.R2, C2 = b.C2, Value = a.Value };
                    list.RemoveAt(q);
                    merged = true;
                }
        }
        return list;
    }

    public static string[,] NumberFormats(Block b)
    {
        var fmts = new string[b.Rows, b.Cols];
        var runs = new List<Run>();
        for (int j = 0; j < b.Cols; j++)
            Runs(b.Ws, b.Row, b.Col + j, b.Row2, b.Col + j, (Func<dynamic, object>)(r => r.NumberFormat), runs);
        foreach (var run in runs)
            for (int r = run.R1; r <= run.R2; r++)
                for (int c = run.C1; c <= run.C2; c++)
                    fmts[r - b.Row, c - b.Col] = run.Value as string;
        return fmts;
    }

    public const int xlCalculationManual = -4135;

    public void RecalcIfManual()
    {
        if ((int)App.Calculation == xlCalculationManual) App.Calculate();
    }
}
