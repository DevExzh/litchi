# XLSX row-start index

Status: accepted for the measured narrow-range query

The immutable XLSX worksheet store now records row-start offsets over its
existing row-major cells. Range selection can begin at the requested row
instead of scanning every preceding row. It does not add cells, change
normalization, or alter transaction/readback semantics.

Matched ABBA results for `xlsx_narrow_column_range_scan` report a p50
geometric-mean change of **-80.499%** and mean geometric-mean change of
**-79.962%**. The full-scan guardrail is effectively neutral (mean **+0.03%**)
and first-cell lookup is **-1.31%** mean. The implementation adds 17 heap
allocations and **+0.25%** RSS in the profile; those costs are retained because
the query reduction is material. Raw paired samples are
[`before A`](../results/abba-xlsx-range-before-a.json),
[`after A`](../results/abba-xlsx-range-after-a.json),
[`before B`](../results/abba-xlsx-range-before-b.json), and
[`after B`](../results/abba-xlsx-range-after-b.json).

This is not evidence for XLSX CRUD generally: broad edits, patching, merge and
split operations, real-workbook diversity, and cold-source behavior remain
outside the measured scenario.
