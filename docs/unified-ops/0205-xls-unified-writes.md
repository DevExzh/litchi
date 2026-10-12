# 0205 — Unified XLS scalar writes

The XLS adapter routes scalar cell `set`, explicit `clear`/null, missing-cell
insertion and worksheet-name changes through the canonical native
`Package::edit` transaction. Path, range and query targets use the shared
registry and semantic limits. The generic transaction owner retains native
patches for replay, inverse and batch attribution; callers keep immutable
input snapshots on failure. This is the first unified XLS write slice, not
completion of the all-format goal.

Scalar record-family changes retain the existing XF style and use the native
row-block rewrite where an in-place replacement is impossible. Formula cells
refuse scalar replacement and clear explicitly: setting a cached formula
result would not implement scalar replacement. Capability output lists that
gap alongside formula/style, tab-state and structural operations. Empty
property maps preserve exact bytes.

Explicitly authorized signature removal retains the native input, editing
and positional query policies through a fenced transformed-byte handoff.
Unlocked encrypted owners still refuse ordinary editing/publication. Exact
patch history charges now include both artifact buffers and forward/inverse
semantic data, with saturating arithmetic and no serialization or inverse
clone. The charge measures owned payloads; it does not claim allocator-level
peak heap measurement.

## Native string-table closure

Runtime tests exposed a pre-existing native issue: appending a replacement
string, or clearing a last string reference, could leave the shared-string
table's total and unique counts inconsistent. The native strict parser
correctly rejected the candidate. The [MS-XLS SST specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b6231b92-d32e-4626-badd-c3310a672bab)
defines the total as workbook references, so inflating it to bypass validation
would be incorrect. The native closure counts final references, prunes unused
plain entries, remaps LabelSst indexes, and publishes exact total/unique counts.
It creates an SST for the first text in a numeric-only or empty workbook and
regenerates existing ExtSST offsets after rebuilding continuation records.

Retained source strings carry raw headers, code units, rich runs and ExtRst
bytes, including producer-controlled reserved fields. A read projection's
tolerance repairs are not silently applied during publication. Wide character
continuations keep code-unit boundaries. Removing the last rich, phonetic or
otherwise unrepresented string identity is refused until the semantic inverse
can restore that metadata. Explicit resource removals verify identity and keep
the existing suffix-only rule so untouched LabelSst indexes cannot silently
refer to another string. These are native dependency gaps, also reported by
the facade's capabilities.

## Verification and measurement

The all-feature core/XLS/facade coverage passes 3,595 tests and doctests, with
one ignored. The XLS capability golden was regenerated, followed by a passing
31-test all-feature capability/write rerun. Twenty-one external write tests
cover source/eager owners, type changes and style retention, first text,
null/clear/no-ops, formula refusals, ranges/queries, history replay/inverse,
batch rollback/attribution, budgets, signed/encrypted policies and stale-source
publication. Native tests additionally exercise packed numeric cells, atomic
resource staging, exact and semantic inverse, shared-string compaction,
wide continuations, ExtSST offsets and opaque rich/phonetic preservation.

Warning-denied all-target Clippy and rustdoc pass for the three packages.
The minimal facade build, scoped formatting, XLS glossary audit (24 banned
names, one transaction family), and crate-boundary checker (66 packages,
268 internal dependencies, 11 existing debt items) pass. The generated
OfficeCLI delta and all 15 generator tests pass. Repository lockfiles remain
unchanged; offline validation uses an external lockfile selecting cached
cc 1.6.0 and an external build target.

Reproduce the broad suite with `cargo test -p litchi-core -p litchi-xls
-p litchi-ops --all-features`. No Python/CLI acceptance closure is claimed.

Five release samples compare one native `Package::edit`/`set_value`/`commit`
against one unified `File::set`, changing A1 from 1 to -1 in eager numeric
workbooks. Median microseconds:

| Cells | Native | Unified | Unified/native |
|---:|---:|---:|---:|
| 1 | 52.680 | 53.000 | 1.0061× |
| 100 | 110.551 | 110.520 | 0.9997× |
| 1,000 | 731.903 | 717.894 | 0.9809× |
| 10,000 | 6784.880 | 6738.061 | 0.9931× |

Every output byte, full checksum and source-immutability check agrees. These
samples are within the 10% facade overhead target for this scenario. Corpus
generation, opening, output serialization and correctness checks are excluded;
native edit/commit publication work is included. The small differences below
1.0× are measurement noise, not a claimed speedup. This does not qualify text
table rebuilds, source-backed writes, memory peaks, batched scaling or the full
performance program. The paired harness is
`crates/litchi-ops/examples/xls_write_bench.rs`.

The [measurement JSON](results/0205-xls-write.json) contains all samples.
Source commit: `f5f4dd74abd28c407ec6ac4c1494e684c7df1ab1`. Build with
`cargo build --release -p litchi-ops --features xls --example xls_write_bench`
and run with `LITCHI_XLS_BENCH_SAMPLES=5 target/release/examples/xls_write_bench`.
The measured binary SHA-256 is
`f47d754db2ae61aa64130b8d3c625587dfe51c2afeb82e62b12992b4fb6d0add`;
no binary is committed.

## Source-test correction (0206)

Record 0206 found that the common write-test helper opened both branches
from bytes, so its `Auto` branch was eager for CFB. The dedicated filesystem
freshness/publication tests did use sources, but the broad source/eager claim
above was too broad. The helper now opens real temporary filesystem sources
and retains their lifetime; record 0206 reports the corrected validation.

## Remaining scope

XLS structural verbs, formula/style authoring, security publication,
transaction coalescing and the 1,000-operation batch performance gate remain
open. Source-backed and axis query streaming from record 0204 also remain
open, as do the remaining formats, Python/CLI and full acceptance gates.
The rewritten base remains an ancestor at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; no additional rebase was needed.
