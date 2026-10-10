# 0175 — XLSB package creation and canonical builder

## API and behavior

`litchi_xlsb::Builder` (also exported from the focused `writer` module) replaces
`WorkbookWriter`. `Builder::build()` returns a validated `Package`; serialization
and filesystem publication belong to that package. The builder's old `save`
entry point is removed without an alias. `Package::new()` replaces `create()`
and authors the deterministic minimal workbook containing one visible Sheet1.
An empty builder refuses without inventing a sheet; a caller can add a sheet and
retry. Building validates the authored workbook and every ordinary worksheet
before returning a package.

A repeated-build regression exposed accumulated shared-string occurrence
counts in the former writer. Each build now starts fresh shared-string staging,
restores the prior scratch state on failure, and recomputes the current string
closure. Repeated text builds are byte-identical; replacing the last text cell
removes the now-unneeded SST while preserving previously returned packages.

The unified `create(FileFormat::Xlsb)` route adopts the native package as an
eager workbook. Its created snapshot supports existing reads, cell writes,
sheet rename, publication and reopening. Template, Strict and initial metadata
requests are explicit typed refusals and capability gaps. Disabled XLSB builds
retain their format refusal.

Native source examples/tests, public examples, facade fixtures and benchmark
generators migrate to `Builder::build()` followed by package publication.
Historical evidence remains unchanged. The glossary profiles forbid the old
writer name, `Package::create` and builder publication methods.

## Validation and evidence

The native suite passed 1,138 tests/doctests; the combined
`docx,xlsx,pptx,xlsb,encryption` facade suite passed 1,197. Native and combined
facade all-target Clippy, native rustdoc JSON and facade documentation passed
with warnings denied. The full XLSB glossary profile passed all 31 name bans.
The crate-boundary audit passed for 66 packages and 268 internal dependency
declarations, retaining the existing 11 explicit debt items.

All 26 migrated public XLSB examples, the facade library and XLSB integration
test target compile. The first `litchi --features xlsb --all-targets` attempt
also selected the unrelated `core_props_office` example, whose imports require
DOCX/PPTX/XLSX despite lacking an example feature gate; verification therefore
explicitly selected the affected XLSB consumers. Minimal facade compilation
and all seven DOCX creation tests with XLSB disabled passed. The regenerated
OfficeCLI capability report passes deterministic checking and its generator's
14 unit tests passed.

The existing paired XLSB cell-value and sheet-name harnesses compile after
fixture migration, as do `xlsb_crud` and `xlsb_synthetic_fixture` in the broader
performance tool. No change was made to their timed operations.

[First](results/0175-xlsb-create-first.json) and
[repeat](results/0175-xlsb-create-repeat.json) measurements each contain nine
ABBA rounds with 18 samples per leg, pinned to CPU 0 with no concurrent Cargo,
test or boundary-audit job. Both routes expose Sheet1 immediately and produce
identical 5,155-byte ZIP output. Both failure matrices are empty.

| Run | Native median (µs) | Unified median (µs) | Unified/native |
|---|---:|---:|---:|
| First | 20.790 | 22.430 | 1.079 |
| Repeat | 20.681 | 22.150 | 1.071 |

The native leg measures `Package::new().into_workbook()`; the unified leg
measures `create(FileFormat::Xlsb)`. Both include initial package creation and
acquisition of a ready workbook projection. Serialization, equivalence checks
and owner destruction are outside the timer. These small creation measurements
make no open, write, corpus or memory acceptance claim. Raw samples, source-file
hash manifests and binary/build provenance are in each JSON. Reproduce with
`tools/perf-baseline/unified-xlsb-create/README.md` and `--rounds 9`.

Builds used offline resolution and external targets. Workspace and broader
performance-tool builds used copied lockfiles selecting cached `cc` 1.6.0;
both inherited repository lockfiles remain unchanged. The new standalone
harness has its own committed source lockfile. The branch already contained
the rewritten base. Temporary implementation worktrees and branches were
removed after integration; scratch builds, copied locks, patches and logs were
removed after recording the reduced evidence. No binary or large evidence
artifact is committed.

## Remaining scope

Initial unified metadata, template and Strict creation are still explicit
follow-ups. XLSB formula authoring, formatting, structural facade CRUD and
durable cell history remain open. Other-format and Python requirements of the
full goal remain required.
