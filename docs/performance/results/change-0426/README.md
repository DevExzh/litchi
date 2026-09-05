# 0426: facade correctness and BIFF8 formula ancillary data

The baseline is `f22917bbe`. Change 0425 reproduced six facade failures at
`340cc91ae` with the same dependency lockfile; its retained receipts are the
baseline evidence. This batch first corrects five stale fixtures/assertions and
then fixes the remaining XLS reader failure. It makes no performance claim.

The five test corrections follow content-derived ODT ownership, the existing
archive-wide ZIP layout proof, and a valid minimal OPC relationship graph.
Malformed inert OOXML-looking extras do not change a valid ODT's owner, while
valid OOXML/ODF polyglots still enforce DOCX input limits. XLSX tests now corrupt
an unselected member's payload instead of its framing metadata. The XLSB
wrong-format test now supplies a complete OPC root relationship and retains its
typed refusal. The initial full facade rerun passes those five cases and leaves
only the XLS formula failure.

`formula-census.py` independently reads the checked-in CFB Workbook stream and
finds 1,416 Formula records. Five contain a standard ten-byte `PtgExtraMem`
suffix. The prior Formula codec required the declared token stream to end at
the record boundary, rejecting valid `CellParsedFormula.rgcb`. `scope.json`
binds the fixture/specification evidence and normative references. The
[design review](checks/xls-ancillary-design-review.md) maps the source and edit
paths; the implementation and verification evidence distinguish exact byte
retention from formula evaluation or support for every formula grammar.

The bounded ancillary owner validates nonempty Mem/Array tails, retains the
original cell and tokens, and shares an immutable source buffer through formula
metadata. Writer publication requires both the original coordinates and the
actual emitted token bytes. Structural coordinate changes and canonical formula
resources refuse tails they cannot represent safely. Cache/style changes retain
the bytes. Empty tails preserve the existing opaque-token contract; this is
not a new complete RPN validator or certification of missing ancillary data.
Elf/revision forms remain outside the supported ancillary scope.

The native workbook supplies the five-record census, eager/source-backed reads
and exact owned no-op coverage. Permitted numeric publication, cache/style
edits and inverse restoration use a separate macro-free synthetic workbook
containing a native token/tail pair and an unrelated opaque CFB stream. The
native file's macros correctly refuse source-backed numeric publication; its
style append hits an existing XFCRC mismatch. Those failed fixture attempts are
retained and neither production restriction is weakened. Structural refusal is
asserted at staging, where the API performs it. The initial compile attempts
also retain a trivial-cast lint and an ambiguous integer-literal error, both
subsequently corrected.

The final complete XLS suite passes 1,341 tests with one ignored doctest. All
461 facade tests pass with 11 ignored, closing the six failures reproduced in
0425. The 12 new public integration tests are included in the XLS total. All
eight dependent XLS/CFB harness tests pass. Strict XLS Clippy and warning-denied
XLS/facade rustdoc pass. The facade strict rerun retains the same 18 prior
findings. No lint allowance is added; ODF layout, harness lint and native-resave
lockfile debt from 0425 remain open. The first complete XLS compile also found
four redundant test qualifications; subsequent strict attempts found one
unnecessary closure and a constant-width test iterator. These are corrected
and their failed receipts are retained.

All Cargo/build/test/profiling workloads are serialized by root with Rust
1.98.1, four build jobs, and one test thread. `check.py` retains exact commands,
source hashes (including new Rust files and the ignored workspace lock), raw
output, and terminal outcomes. The initial fixture-only receipt predates the
driver's improved all-target test-summary counting; later receipts also count
passes inside failing targets. No failed attempt is deleted.

The type-layout probe records these static sizes from a retained pre-batch
library and the freshly compiled candidate, with library/source/binary hashes:

| Rust type | Before | After |
| --- | ---: | ---: |
| Formula metadata | 24 B | 32 B |
| CellRecord | 80 B | 88 B |
| Semantic Cell | 152 B | 160 B |

The eight-byte field cost also applies when metadata has no ancillary owner,
including ordinary semantic cells. The nonempty owner duplicates the existing
Formula token bytes in its combined token/tail buffer to prove later writer
identity; cloning metadata shares that owner. This correctness tradeoff is
explicitly retained. It is not an allocation, live-memory or RSS measurement,
and no timing or memory reduction is inferred. Native application resaves,
full-workspace tests, broad performance recapture and the outstanding global
strict-gate debt are not implied. The complete non-iWork performance goal
remains active.

Crate boundaries, the 15-category/32-selector CRUD index, all nine strict claim
replays, and formatting pass. The verification scope follows the XLS production
owner, facade consumers and dependent harness; the complete 45-package run from
0425 is historical evidence and is not represented as a fresh run here.

Run `python3 -B docs/performance/results/change-0426/verify.py` to replay all
retained command outcomes, source consistency, raw/compressed log hashes, static
layout observations and the inventory without Cargo. Final replay passes all
22 command receipts and 68 inventoried files; 24 logs are compressed losslessly
with original and stored hashes. When the native fixture
is present, the verifier also reruns the independent Python CFB census. Use
`planned-checks.json` and `check.py` for fresh serialized build/test commands
with new receipt tags; existing receipts are deliberately not overwritten.
