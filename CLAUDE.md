# comgrid

`cgr` (C#, .NET Framework 4.8, late-bound COM) reads and edits workbooks open in Excel.
See README.md for commands and design.

- Build: `cmd /c build.bat` from PowerShell (Git Bash's `cmd //c` loses the cwd). Output `bin\cgr.exe`.
- Source is `src\*.cs`: Program (CLI), Excel (attach, ROT, busy filter, bulk reads,
  format runs), Look (outline/dump/find), Refs (formula refs, shapes, trace), Edit
  (writes, journal, undo, whatif), Cells (addresses, value and number-format rendering).
- **`dynamic` gotcha:** passing a `dynamic` argument makes the whole call late-bound, and
  lambdas can't be passed then (CS1977). Type the variables (`List<dynamic> x = ...`) or
  cast the lambda to `Func<dynamic, object>`.
- **Test new COM calls in an isolated instance first**: `cgr open copy.xlsx --isolated`,
  then `cgr -b copy ...`; kill leftover hidden instances afterwards. Some innocent-looking
  COM reads crash Excel outright (see README, "Excel bugs"). `CGR_DEBUG=1` prints step
  markers to stderr to find which call did it.
- Ranges from shapes/charts (`TopLeftCell`, `BottomRightCell`): only read scalar properties.
- After `insert`/`delete`, the Range object moved or died; take addresses before the operation.
