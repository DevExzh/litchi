# Default-budget XLS scanner fixtures

Generate from `test-data/poi/test-data/spreadsheet/Simple.xls` using:

```
xls0686-generate test-data/poi/test-data/spreadsheet/Simple.xls OUTPUT_DIRECTORY
```

The original first-sheet cell remains at (0,0). Insert 70,000 or 100,000
NUMBER frames in row-major order, two columns per row, starting at row 1.
The resulting first sheets have 70,001 or 100,001 stored occurrences.
DIMENSIONS is extended and later BoundSheet offsets are shifted. The output
contains only a freshly wrapped Workbook stream; this is a deterministic
scanner fixture, not a claim of native Office preservation or compatibility.
Helpers derive from the committed XLS query-cache integration fixture helpers.
The packet retains template/output hashes, two-run determinism and independent
full-visitor counts/digests. No timestamps or random data are introduced here.
