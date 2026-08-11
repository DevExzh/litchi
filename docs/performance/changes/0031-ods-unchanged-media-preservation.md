# ODS unchanged-media publication and comparison

Date: 2026-08-11

Production base: `78ac94b2113b67db3b59090dd6b09407b0a3a140`

Scope: compact `content.xml` replacements in `litchi-odf-common` and unified
ODS changed-member comparison in `litchi-ods`. OLE2, OOXML, RTF and the other
ODF family production crates are unchanged. iWork/IWA was explicitly excluded
while its crates are changing independently.

## Profile gate and hypothesis

The existing generated ODS semantic corpus had only three archive members and
could not attribute unchanged media work. A new opt-in
`ods_media_one_edit_save` case builds a deterministic medium spreadsheet plus
eight 2 MiB incompressible opaque resources under `Pictures/`. The timed path
opens a public unified document snapshot, edits the middle cell, commits and
materializes the output. Outside timing it reopens the complete grid and
verifies every resource path, manifest media type and exact payload.

The source audit found that one cell edit inflated each unchanged media member
seven times: once to rebuild/recompress it for publication, then twice in each
of three source/target comparisons used by staging, authored-part validation
and durable patch construction. Raw-copying unchanged members alone was
latency-neutral in an eight-sample pilot because the six comparison inflations
remained. The accepted implementation combines raw publication with a
conservative physical-identity hint so those comparisons can avoid
decompression.

## Change and fallbacks

The shared ODF package owner now supports a checked compact `content.xml`
replacement that regenerates only that ZIP member and raw-copies every other
validated member. It retains source physical order, central-directory order,
comments, extras, descriptors and compression representation. Exact semantic
no-ops return the accepted source bytes unchanged.

Raw preservation is deliberately best-effort. Encryption, signatures,
size-bearing content manifest entries, ZIP64, prefixed/multi-disk/ambiguous or
otherwise unsupported ZIP layouts, unsupported compression, duplicate or
unsafe paths, and non-splice whole-content changes use the established logical
rebuild before publication. Signed fallback retains the existing signature-
stripping behavior.

The same common owner can identify physically identical members without
inflating their bodies. Both raw local-member spans must be byte-identical;
their complete central records must also match except for the local-header
offset that changes when an earlier member changes length. ODS uses this only
when `META-INF/manifest.xml` is itself physically identical, so payload and
media-type identity are both proved. Every unproved member uses the former
logical decompression and `File { bytes, media_type }` comparison. Recompressed
but logically identical members therefore remain unchanged, while manifest
media-type or payload changes still produce the same deterministic effect.

No public ODS transaction or patch API, operation order, exact-source lineage,
inverse, security policy, compactness validation, package reopen, snapshot
parse, typed grid/resource readback, dependency edge, unsafe code, runtime,
lock or global state changed. Two explicit byte-slice annotations in the flat
ODS validator also remove test-build type ambiguity without changing accepted
entities or text rules.

## Corpus and experiment

The media-rich archive has 2,048 cells, eight 2 MiB resources, 11 ZIP members,
16,887,808 logical payload bytes and 16,790,689 archive bytes. Its SHA-256 is
`46b7f61cb74639115f6d120dc6498b97d6b310d51c78c4fb85ac60d6fc758b14`.
The existing medium no-media guard has 2,048 cells, three members, a 6,983-byte
archive and SHA-256
`070a56361bfb2cc69815abe13b12c52db8cc01f8ba3bfde669c4adef986918c6`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The formal order was before A, after A, after B,
before B with three warmups and 30 samples per case per leg, pooling 60 samples
per state.

The frozen common-harness baseline executable SHA-256 is
`1904875ba051b00a22c77e8dba14f0cf6ba4b29b7c8c15de386ecf8b143b1add`;
its `.text` SHA-256 is
`391033ff7d4353bad46fd2676d7e1278936f22603f2f9fedbbac2d2c930f4bab`.
The accepted executable SHA-256 is
`dae054e512a045216c754821c93599f570a1fe54d13d586bdbdfb73405dfe6d5`;
its `.text` SHA-256 is
`dbff56d4765d750a44aa6eaafd5f286b34522aa85344e2d34a020fe1c8edbd41`.
After the test-only raw-selection assertion was strengthened, the final source
rebuild produced ELF SHA-256
`38536012fac010c736a547ce1bed5b79e34570a0fdddbca1ac89855373b9ff68`
with the same `.text` SHA-256, so the formally measured release code is exact.

## Formal latency result

| Workload | Before p50 | After p50 | p50 delta | Mean delta | p95 delta | Approximate 95% interval for mean delta |
|---|---:|---:|---:|---:|---:|---:|
| Media-rich one-cell edit/save | 325.902 ms | 310.472 ms | **-4.73%** | **-5.73%** | **-7.65%** | [-7.13%, -4.34%] |
| Existing medium no-media edit/save | 20.818 ms | 20.657 ms | -0.77% | -0.81% | -4.59% | [-1.51%, -0.11%] |

Both accepted legs are below both baseline legs for the media-rich mean. The
existing no-media case does not regress, so the result is accepted as material
for the newly attributed unchanged-media workload rather than generalized to
all ODF edits.

## Memory and counter attribution

Matched one-sample Heaptrack processes include deterministic corpus generation
and final verification:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 387,014 | 387,426 | +412 (+0.11%) |
| Temporary allocations | 89,726 | 89,774 | +48 (+0.05%) |
| Peak heap | 156.53 MiB | 142.79 MiB | **-8.78%** |
| Heaptrack RSS | 165.83 MiB | 165.71 MiB | -0.07% (flat) |
| Leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

Uninstrumented GNU Time reports maximum RSS of 158,320 KiB before and
157,188 KiB after (-0.71%), with zero major faults. A matched `perf stat`
process reports task clock -3.54%, cycles -1.54%, instructions -0.63%,
branches -0.05% and cache references -5.92%. Branch misses (+0.42%) and cache
misses (+1.29%) move slightly in the wrong direction and are disclosed rather
than treated as a locality improvement.

## Correctness and verification

- focused raw-preservation tests prove exact untouched local spans and central
  records for mimetype, manifest, a 4 MiB opaque resource and an unknown member;
- exact no-op, complete logical reopen and signed fallback/signature stripping
  are checked;
- ODS physical-diff tests prove recompression falls back to logical equality,
  while payload and manifest media-type changes remain effects;
- the media-rich harness test proves deterministic corpus bytes and exact
  resource preservation after a real public one-cell transaction;
- all-feature/all-target `litchi-odf-common` tests pass: 261 tests; its
  warning-denied all-target Clippy gate passes;
- all-feature/all-target `litchi-ods` tests pass: 242 tests; production-library
  warning-denied Clippy passes;
- all 24 harness tests and warning-denied all-target harness Clippy pass;
- formatting, `git diff --check`, JSON/hash checks, release corpus identity and
  staged-scope checks are commit gates.

The broader ODS all-target warning-denied Clippy command still reports the six
known unrelated test/module lints recorded with change 0027. Harness builds
still print the two existing ODF GenericArray deprecation warnings. Neither is
presented as a passing gate.

Raw ABBA, Heaptrack, GNU Time, `perf stat` and harness reports are under
`docs/performance/results/`; their digests are in
`ods-media-preservation-sha256.txt`.

## Next non-iWork work

1. ODF: keep structural edits, other family publications, real-producer media,
   signatures/encryption and source-backed reads as separate measured paths.
2. OOXML: profile a broader semantic-planning/emission boundary or the
   x14ac/dyDescent scan before another production change.
3. RTF: attribute formatted/media and broader real-producer edits beyond the
   current capability-bounded variants.
4. OLE2: continue shared CFB payload/final-reader attribution without reviving
   the rejected XLS terminal-render handoff.

iWork remains deferred while the `iwa-*` crates are modified independently.
