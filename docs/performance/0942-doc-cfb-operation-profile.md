# 0942 — native CPU attribution inside the DOC edit lifecycle

This batch adds a profiling harness and measures the restored DOC repeated-edit
path after the rejected 0920–0922 candidates. It changes no production code and
makes no latency, memory-saving, or adoption claim. Unlike the earlier
whole-process profiles, the new callchain marker identifies samples inside the
measured lifecycle while retaining the unassigned remainder.

## Method and validation

The standalone [harness](../../tools/perf-cfb-profile/README.md) derives from
0917. Its non-inlined `profiled_measured_case` wrapper surrounds source cloning,
bounded snapshot opening, one Edit with repeated paragraph-zero replacements,
commit, output copying, and local destruction. Preparation, semantic oracles,
per-sample validation, and serialization remain outside the wrapper. This is a
DOC lifecycle containing CFB work, not a standalone CFB benchmark.

Production source is based on `73559c87c37898b675100b02985181a7764c66d2`.
The immutable `source-v2` export and `profile-v3` build use Rust 1.95.0,
`--offline --locked -j2 --release`, `RUSTFLAGS=-Cforce-frame-pointers=yes`,
`CARGO_PROFILE_RELEASE_DEBUG=2`, and `CARGO_INCREMENTAL=0`. The executable is
126,326,672 bytes, SHA-256
`92f85ce6019405fa1b0d97ec796947cd7ec240d2b8143f9cf06356d97409e70e`.
Static inspection finds the marker at ELF address `0xdf130`, its direct
non-tail call to `run_case` at `0xe3d50`, frame-pointer prologues, and source
line information. The wrapper still moves its large result with a 0x910-byte
`memcpy`; that harness overhead remains profile-visible.

The frozen [plan](results/change-0942/profile-plan-v5.json) uses CPU 14,
`cycles:u`, 997 Hz, frame-pointer callchains, zero warmups, and 200 samples per
child. One FloatingPictures/one-edit preflight precedes sixteen formal children:
two fixtures × 1/2/8/32 edits × two repetitions, forward then reverse cell order.
All children run serially, with terminal receipts before parsing and no retries.
The [manifest](results/change-0942/profile-manifest-v5.json) binds source,
binary, fixtures, validator, tools, environment, and commands.

The preflight passes with 214 marker samples. All sixteen formal children pass
the existing 0917 correctness validator, including route equality, preservation,
patch inversion, conflict refusal, no-op sharing, and atomic edit controls.
Formatting and warning-denied Clippy pass for the harness. Offline reduction
rehashes the retained artifacts and reproduces all sixteen summaries.

Before Office capture, a non-Office format check established the actual
`perf script` header layout and independently counted 294 samples and
1,205,516,268 period units. Export uses `--no-inline --show-lost-events`.
Parser controls reject malformed headers, zero periods, wrong events, and
markers from another DSO. Truncated marker stacks remain observed ancestry but
do not count as fully resolved. Superseded build and preparation evidence is
retained; it is not a discarded Office cohort.

## Observed marker subset

The formal captures contain 18,209 samples, of which 5,866 contain the bound
marker. No lost records or explicit truncation are reported. This does not
prove complete unwinding: some frames are unresolved and system libraries may
omit frame pointers. A missing marker therefore leaves operation membership
unknown. Marker period coverage ranges from 10.3% to 54.5% of each process;
these percentages are not the fraction of elapsed time spent editing.

The table shows each repetition separately. Owner percentages use sampled
period weights within that child's marker subset. Each sample belongs to the
nearest observed physical project frame when walking from leaf toward marker.
Inlining and missing frames prevent interpreting these as exact source-module
costs. OLE-common, core, and other complete the disjoint buckets in the
[machine-readable reduction](results/change-0942/formal-reduction-v5.json).

| Fixture, edits | Marker samples, R1 / R2 | CFB weight, R1 / R2 | DOC weight, R1 / R2 |
| --- | ---: | ---: | ---: |
| NoHeadFoot, 1 | 16 / 29 | 49.6% / 41.5% | 31.4% / 38.1% |
| NoHeadFoot, 2 | 33 / 30 | 42.4% / 36.5% | 33.4% / 40.2% |
| NoHeadFoot, 8 | 73 / 68 | 46.6% / 48.3% | 35.4% / 41.4% |
| NoHeadFoot, 32 | 208 / 208 | 58.7% / 56.8% | 27.3% / 29.2% |
| FloatingPictures, 1 | 198 / 227 | 18.7% / 19.8% | 73.2% / 71.3% |
| FloatingPictures, 2 | 254 / 257 | 21.3% / 19.5% | 68.5% / 70.8% |
| FloatingPictures, 8 | 526 / 524 | 31.6% / 33.5% | 53.8% / 54.9% |
| FloatingPictures, 32 | 1,623 / 1,592 | 36.8% / 38.5% | 48.6% / 48.1% |

Short NoHeadFoot cells are too sparse for detailed function rankings. The
larger cells consistently expose copying, allocation, and DOC parsing work:
FloatingPictures/32 has `memmove` self weight of 26.3% / 25.7% and
`parse_sprms` self weight of 5.7% / 5.7%. These are measured-subset observations,
not removable-work estimates or speedup bounds. Inclusive function totals
overlap and must not be added to the exclusive owner or leaf totals.

## Follow-up and custody

The profile supports separate investigation of repeated CFB reconstruction and
DOC parsing/copying; it does not justify reviving the rejected reserve,
source-binding, or borrowed-FIB candidates. Any production candidate still
needs a prospective normal-release comparison and all preservation gates.

Caller analysis narrows that investigation: `OleWriter::write_to` appears in
281 / 277 marker stacks for FloatingPictures/32, while copying stacks name
`RevisionEditor::replace_plain_text`, OLE `Package::capture`, and CFB
`open_stream` and `write_to` as immediate callers. Source review should first
trace repeated package capture/reopening and layout planning while preserving
stream validation and source identity. These caller counts are inclusive and
overlap; they differ from exclusive sampled-leaf self costs above.

The native harness preserves the 0917 JSON schema, including its historical
`timing_claim: true` field. For this instrumented build the outer protocol
explicitly overrides that interpretation: raw elapsed arrays are retained only
as diagnostics and are not timing qualification. RSS describes the entire
profile process; there is no allocator comparison or memory improvement claim.

Raw captures, source manifests, build logs, and binaries remain outside Git
under `/home/zhuhe/litchi-performance-0942`. The verified archive is
`/home/zhuhe/litchi-performance-0942-archive/litchi-performance-0942-v1.tar.zst`,
176,574,454 bytes, SHA-256
`bc03acd40a3a1992df8b27047904039139db72f5e46abdc28586d222ed609301`.
Extraction verifies all 21,823 selected files and 845,267,181 logical bytes.
The [custody summary](results/change-0942/custody-summary-v1.json) records the
per-file manifest identity, regeneration commands, build command and environment,
source/binary hashes, and cleanup receipt. The committed
[archive script](results/change-0942/archive-v1.py) generates that manifest and
archive with `--go --archive`, verifies with `--go --verify`, and permits
cleanup only after verification; its exact original invocation is in the summary.

[Independent review](results/change-0942/independent-review-v1.json) checks all
seventeen children and replays a relocated raw packet. A second
[independent counter](results/change-0942/independent-count-v1.py) reproduces
sample weights and owner buckets without importing the capture parser, and
checks binary receipt and terminal-command equality. The relocated copy,
verification extraction, and build caches are removed; both bound executables
and the source/evidence closures are retained outside Git.
