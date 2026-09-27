# comgrid

`cgr` reads and edits the workbooks open in your running Excel over COM, and is
built to be driven by a coding agent: you talk about a spreadsheet, the agent looks at
it and changes it, and you watch it happen in Excel.

- **Live by default.** It attaches to the Excel you already have open, so it sees
  unsaved edits and there are no file locks. Writes land in the workbook on your screen.
- **Compact output.** Formulas are compared by what they reference, so a column filled
  down 50 rows prints as one line. The 283x35 model sheet this was developed against
  prints in 80 lines instead of about 9,000.
- **Fast.** Ranges are read in bulk (one COM call per property per block), not cell by
  cell. Reading every cell of that workbook one at a time took minutes; in bulk it takes about 0.4 s.
- **Undoable.** Excel clears its own undo stack whenever a program writes a cell, so cgr
  journals what each write replaced, and `cgr undo` restores it.

## Build

Windows, .NET Framework 4.8 (ships with Windows 10/11), and the Roslyn compiler from
Visual Studio or the free Build Tools. No SDK, no NuGet:

    build.bat        ->  bin\cgr.exe  (~100 KB)

## Use

    cgr books                                 workbooks open in any Excel instance
    cgr outline                               sheets, cell islands, headers, row labels, names, charts
    cgr dump "'Annual model'!B14:AI65"        collapsed formulas with values and number formats
    cgr dump Comparison!B6:K11 --grid         values as displayed, tab-separated
    cgr sel                                   whatever is selected in Excel right now
    cgr find IRR                              formulas/values containing text (numbers match by value)
    cgr trace Comparison!J7 --depth 2         what a formula reads
    cgr trace Assumptions!C12 --dependents    what reads a cell

    cgr set Assumptions!C12 0.05              value or =formula (fills relative refs across a range)
    cgr set Sheet!B5 - < block.tsv            a block of values/formulas from stdin
    cgr fmt Sheet!B6:K6 bold=1 fill=#233D59   formatting; `cgr fmt RANGE` alone shows it
    cgr insert Sheet!5:7 / delete Sheet!C:D / addsheet Name
    cgr undo [N] / cgr log

    cgr whatif Assumptions!C12=0.02,0.03,0.04 -- Comparison!J7:J11   sweep an input, inputs restored

`-b BOOK` picks a workbook by name, unique substring or path (default: `$CGR_BOOK`, then
Excel's active workbook). `cgr open FILE --isolated` opens a file in a separate hidden
Excel instance: useful for experiments that shouldn't touch your session.

A dump looks like this:

    'Annual model'!B14:AI116  (103x34)
    B15:B65    "San Conrado #8"  ×51
    C15:C65    0, 1, … 50  (step 1, ×51)
    D15:D65    =$C$6*(1+Assumptions!$C$12)^C15  ×51  → 789000 … 3458901.84876  [\$#,##0;(\$#,##0);"–"]
    J16:J65    =$C$6*Properties!L6*(1+Assumptions!$C$11)^(C16-1)  ×50  → 2761.5 … 9260.0738824
    ...

## How it works

- **Attaching.** `GetActiveObject("Excel.Application")` for the active instance, plus
  the running object table, where every open workbook file registers a moniker; that is
  how workbooks in other Excel instances are found.
- **Busy Excel.** While you are typing in a cell, Excel rejects COM calls. An
  `IOleMessageFilter` retries them for a few seconds before giving up with a clear message.
- **Late binding.** Everything goes through `dynamic` / `IDispatch`, so no Office interop
  assemblies or type libraries are needed. That matters: on the development machine, the
  typed interop that PowerShell uses fails with `TYPE_E_CANTLOADLIBRARY`.
- **Collapsing formulas.** Each formula is split into a skeleton (text outside references)
  and the rectangles it references. Going down a column, two cells belong to the same run
  if their skeletons match and every reference either stays put or moves with the cell,
  consistently along the run. That groups `=A1*$B$1` filled down, and also generated
  formulas that write a fixed reference without `$`.
- **Formats in few calls.** Range properties return null when the range is mixed, so cgr
  halves ranges until each piece is uniform: one call per run of equal formatting, not per cell.
- **Journal.** `%LOCALAPPDATA%\comgrid\journal\<workbook>-<hash>.jsonl`. Content writes
  record the formulas before and after; `undo` refuses if the cells changed since the
  write, unless `--force`. Deleting rows can only be undone for contents: formats, and
  references the delete turned into `#REF!`, stay as they are.

## Excel bugs found along the way

On Excel 16.0.20326 (Microsoft 365, click-to-run), reading `.Rows`, `.Columns` or
`.Cells` on the Range returned by a chart's `ChartObject.TopLeftCell` (or
`BottomRightCell`) **crashes Excel** with an access violation. Scalar properties such as
`.Row` and `.Address` are fine. cgr only reads `.Row` / `.Column` there.
