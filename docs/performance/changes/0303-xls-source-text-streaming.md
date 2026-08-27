# Change 0303: source-backed XLS text projection

Status: implemented

performance_claim: none

On supported filesystem platforms (`unix`, `windows`), the XLS facade routes
text extraction through the source owner. Each ordinary worksheet is scanned to
EOF, retained in a bounded worksheet-local sparse coordinate map, and emitted as
dense row objects only after the scan completes. This removes full
archive/eager semantic-workbook materialization from that facade text path while
preserving the eager projection's row, column, duplicate, formula, formatting,
and terminal-newline semantics.

The source path is bounded by explicit retained-cell and decoded-text-byte
limits, plus the existing scan and text-output limits. Allocation failures,
resource limits, framing and decoded supported-family errors, cancellation, and
stale sources remain typed. `write_text_to*` reports exact completed-row and
accepted-byte progress; `text()` discards that report and returns only a typed
error when conversion fails.

Rows are paragraph objects, so `slide_separator` is ignored. Dense tab-only rows
are non-empty objects. The source scan has no unbounded SST cache. The new
retained-cell and retained-text-byte setters are infallible builders; zero is
rejected by `from_*_with_limits` validation.

This note makes no claim about latency, RSS, allocation count, I/O volume, or
general memory usage. The returned `String` still retains the complete output;
the bounded source projection applies while constructing it.
