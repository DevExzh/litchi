# Change 0275: source-backed XLS selective read owner

Date: 2026-08-25

Status: production correctness and benchmark-boundary closure;
`performance_claim: none`

## Production boundary

`litchi-xls` now owns an immutable `SourceBackedWorkbook` over a caller-owned
`ReadAt` and `SharedOleFile`. BIFF8 open traverses CFB metadata and the complete
Workbook-global substream to validate BOF/BoundSheet topology, codepage, SST,
and formatting state. It does not read any worksheet payload during open,
worksheet-name enumeration, or worksheet-count queries. A selected-cell query
reads only the selected worksheet and scans it through EOF so duplicate scalar
and cached-formula records retain the eager owner's last-record semantics.

The selected-cell path covers Number, RK/MulRK, Label/LabelSst, BoolErr,
Blank/MulBlank, cached numeric/date formula values, and cached formula strings
continued across records. It is a read-only semantic owner: no source-backed
edit, patch, save, or publication API is introduced. Full facade text remains
an explicit compatibility boundary that copies the complete source under the
default 128 MiB materialization ceiling and invokes the existing eager parser.
Metadata uses the shared CFB property-set reader without triggering that eager
text materialization.

On Unix and Windows, `litchi::sheet::Workbook::open(Path)` retains this owner
for BIFF8 XLS. Byte-backed construction remains eager. Valid BIFF2-BIFF5 input
uses a dedicated `UnsupportedBiffVersion` signal and the established eager
path. OLE host precedence remains Word, then PowerPoint, then Excel; malformed
BIFF, FILEPASS, I/O, stale-source, allocation, and resource errors do not fall
back. An open file descriptor stays pinned across atomic pathname replacement,
while a new open sees the replacement.

Finite input/global/SST/sheet/worksheet-scan/materialization limits, typed
allocation failures, per-record size bounds, source-version fences, and
cooperative `ExecutionContext` cancellation remain explicit. FILEPASS is a
typed unsupported-encryption error and inert macro/opaque streams are never
executed. Columns outside the BIFF8 visible range return `None` without a
worksheet read.

## Matched benchmark boundary

Five opt-in selectors raise the registry from 393 to 398 names; the default
remains 36 cases / 198 records:

- `xls_source_backed_open`, paired with existing `xls_semantic_open`
- `xls_eager_open_list_worksheets`
- `xls_source_backed_open_list_worksheets`
- `xls_eager_open_one_cell`
- `xls_source_backed_open_one_cell`

Every eager/source interval includes a fresh open plus the named operation.
Hashing, semantic parity, corpus validation, source-version comparison, and
logical-range evidence construction occur after elapsed time. Pure dispatch
tests prevent duplicate eager-open execution and reject shape, payload, or
writer-shape overrides.

The fixed `litchi-xls-comments-opaque-heavy-v1` corpus is 16,995,840 bytes with
archive SHA-256
`6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53`.
Its 80,946-byte Workbook stream has SHA-256
`c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041`.
It has two worksheets and eight 2 MiB opaque regular streams. The harness
classifies actual CFB structural sectors, Workbook globals, selected and
unselected worksheets, opaque regular streams, and root mini-stream payloads
separately. Any opaque-stream overlap fails the locality gate.

## Descriptive release smoke and next hotspot

A dirty-current-worktree release smoke used one warmup and five samples. It is
prioritization evidence only. Eager/source p50 values were respectively:

- open: `513426 / 5992675 ns`
- open plus list: `508470 / 6092613 ns`
- open plus one cell: `504264 / 6780949 ns`

Source open/list recorded 401 logical `ReadAt` calls and 138,459 bytes per
sample. One-cell recorded 429 calls and 138,593 bytes. The one-cell breakdown
was 136,704 CFB-structural bytes, 1,755 Workbook-global bytes, 134 selected-
worksheet bytes, and zero unselected-worksheet or opaque-payload bytes. Open
and list recorded zero worksheet bytes. `complete_archive_materialized_bytes`
was zero; this metric intentionally excludes parser-owned Workbook globals and
SST. The in-memory harness source itself retains the complete caller-owned
archive, unlike the facade's filesystem `FileSource`.

These are logical overlap counters, not physical filesystem I/O. The smoke is
too small, dirty, and unpaired for latency acceptance, and it is materially
slower than eager open. The next measured hotspot is fine-grained CFB/global
request and freshness-fence overhead, followed by possible deferral of
parser-owned global/SST materialization. Any change must retain malformed-tail,
duplicate-record, source-fence, cancellation, and resource-limit behavior.

## Verification and exclusions

Focused verification passed 8 `FileSource` tests, 22 source-backed XLS
integration tests, 11 XLS facade tests, 20 mixed XLS/XLSX facade tests, and
feature checks for `xls+ods`, `xls+docx`, and `xls+pptx+xlsb`. The two focused
harness tests, harness library check, formatting, and independent production
and evidence reviews also passed.

No ABBA package or strict claim-registry entry is added. No broad XLS latency,
throughput, physical-I/O, cold-cache, allocation/RSS, peak-memory, producer,
edit/save, encrypted-file, or all-BIFF-version claim follows.
