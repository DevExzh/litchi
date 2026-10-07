# DOC SPRM census implementation plan (0966)

Status: diagnostic source staged and build/unit controls pass; no formal
workload run. This plan is
for the 0966 reservation/reparse census only; it makes no timing, allocation,
RSS, speedup, or production-adoption claim.

## Source and authority

The source basis is the immutable external export
`/home/zhuhe/litchi-performance-0966/source-head-v1`, taken from current HEAD
`a499d34f5119a034bc92abde1e1dfcf529159b8d`. It contains the workspace crates,
the retained `tools/perf-cfb-profile` harness, the generated current lockfile,
and the two existing DOC fixtures. The repository checkout was not edited.
The 0942 executable is not reused: it predates this census hook and its build
target contains only the old profiling binary. A fresh target directory is
required for 0966.

Before staging the patch, the full ADR index and all 32 accepted records were
read (51,050 lines). Their hashes are retained in
`adr-sha256-v1.txt`. The census follows the records that matter directly:

- ADR 0001/0002 keep diagnostics below the ordinary semantic facade and inside
  the owning DOC crate.
- ADR 0003 requires the immutable snapshot/edit/commit lifecycle and exact
  correctness boundaries to remain unchanged.
- ADR 0005 requires representative measured evidence, explicit input/output
  boundaries, finite resource behavior, and no hidden runtime state.
- ADR 0006 preserves malformed/unknown content and typed failures.
- ADR 0007 keeps legacy DOC semantics separate from other Office formats.
- ADR 0008 requires production source restoration and verification gates before
  a support or optimization claim.
- ADR 0031's opt-in execution-context rules do not authorize hidden workers;
  this census is single-threaded and does not add concurrency.
- ADR 0032 permits only bounded derived values under an accepted memo record;
  this patch retains no derived value in production state and is not a memo.

## External patch

The only diagnostic production-source copy is
`source-head-v1/crates/litchi-doc/src/sprm.rs`. Under the existing
`performance-diagnostics` feature it adds a `sprm::census` module and one
recording call immediately after the unchanged `parse_sprms_two_byte` result.
The ordinary parser grammar, allocation request, result type, error order, and
returned bytes are unchanged.

The hook is enabled only inside `profiled_measured_case` through a calling
thread-local guard. Preparation calls and post-lifecycle semantic oracles run
with the guard disabled. The guard contains fixed arrays and scalar counters;
it never stores an input slice, opcode, byte offset, parsed `Vec`, source hash,
caller label, or per-call record. A dropped guard disables recording.

Each bucket records:

- parser calls, successful results, errors, empty/non-empty input calls, and
  total input bytes;
- minimum/maximum input length plus a ten-bin fixed histogram (`0`, `1`,
  `2-3`, `4-7`, `8-15`, `16-31`, `32-63`, `64-127`, `128-255`, `256+`);
- total parsed SPRM count plus the same bounded histogram of successful output
  counts;
- total/minimum/maximum `Vec<Sprm>` capacity;
- `output_capacity_total * size_of::<Sprm>()`, labeled explicitly as
  `logical_vec_capacity_bytes`, the **logical Vec-capacity model**, never
  allocator bytes;
- five fixed parser-error classes: opcode, length, zero-length, overflow, and
  operand.

All increments saturate and set a `saturated` flag. The independent checker
rejects any saturated report. This makes the counters bounded without hiding a
counter overflow.

## Lifecycle phase boundaries

The external harness marks five fixed lifecycle buckets before calling the
unchanged public operations:

1. `snapshot_open`: `Snapshot::open_bounded` and its strict/public validation;
2. `edit_open`: `snapshot.edit()`;
3. `staged_replacement`: every repeated paragraph-zero replacement, aggregated
   over the selected edit count;
4. `final_commit`: `edit.commit()` and candidate reopen/validation;
5. `output_copy`: copying committed snapshot bytes into the result `Vec`.

An `other` bucket is retained as a fail-closed witness. The independent checker
requires `other.calls == 0` for this lifecycle and requires every scalar and
histogram counter in the five phase buckets to sum exactly to `total`. This is
an accounting check, not a timing partition. Existing preparation, route
oracles, hashes, serialization, and final public readback are outside the guard
and therefore cannot enter the census.

## Matrix and child protocol

The selected operation is the retained 0942 public DOC lifecycle over:

- `NoHeadFoot.doc` (26,112 bytes,
  `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48`);
- `FloatingPictures.doc` (335,360 bytes,
  `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`);
- edit counts `1`, `2`, `8`, and `32`;
- retention routes `zero`, `default_8mib`, and `release_8mib`.

This is 24 cells. Each cell runs in two fresh child processes, for 48 reports;
each child uses `--warmup 0 --samples 1`. Children run serially, with no retry
or replacement. The harness preparation executes all three retention routes
outside the census so the existing cross-route, patch, no-op, limit, stream,
directory, semantic, and public-reader gates remain active. The selected route
is the only operation inside the census guard. The report sets
`timing_claim: false` and emits no timing comparison.

## Build route and provenance

After root authorizes the build, use the external source root and a new target:

```text
cd /home/zhuhe/litchi-performance-0966/source-head-v1
RUSTFLAGS= CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=2 \
  cargo +1.95.0 build --offline --locked -j2 --release \
  --bin doc_sprm_census_0966 \
  --manifest-path tools/perf-cfb-profile/Cargo.toml \
  --target-dir /home/zhuhe/litchi-performance-0966/target-0966-census
```

The build receipt must bind the current HEAD, source-export manifest hashes,
external patch hashes, Rust toolchain, command, Cargo lock hash, executable
path/size/SHA-256/build ID, and the exact fixture hashes. It must not write a
binary, target directory, or evidence file into the repository, and it must not
mutate `/home/zhuhe/litchi-performance-0942` or its archived target.

The binary is diagnostic-only. Its source feature is the already accepted
runtime-neutral diagnostics feature; no new production dependency or public
ordinary CRUD API is introduced in the repository. The external report schema
is `litchi.performance.0966.doc_sprm_census.v1`.

The source export was supplemented outside the repository with the fixture
overlay recorded in `fixture-overlay-v1.json` (`test-data/images`,
`test-data/poi`, `test-data/libreoffice-core`, and `test-data/office-interop`).
The overlay is required for the external unit suite's existing fixture tests;
it is not part of the formal census workload.

## Independent checks

`check_sprm_census_0966.py` is staged beside this plan. It does not open a DOC
or invoke the executable. For each JSON report it verifies:

- schema, source identity, 24-cell key, one sample/no warmup, and
  `timing_claim == false`;
- the exact 20 existing correctness gate keys, each with the literal boolean
  value `true`, plus sample/top/reference output hash, byte-count, and
  capacity parity;
- call/success/error, empty/non-empty, input histogram, output histogram, and
  error-kind conservation, with fixed histogram lengths and integer counters;
- the logical capacity lower-bound equation using the reported
  `sprm_size_bytes`;
- all phase-to-total scalar and histogram conservation and no unclassified
  calls;
- exactly two reports per cell and byte-for-byte equality of the normalized
  census/output/correctness identities across the fresh-child repeats.

Static source review separately checks that the only hook is the one wrapper
recording `grpprl.len()` and the parse result, and that the fixed census types
contain no input-retaining field. Aggregate repeat equality cannot identify
callers or prove a missed hook, so no caller attribution or savings estimate
may be derived from these reports. A future implementation candidate must use
the census only to choose the next semantic/performance experiment, then pass
the existing differential, malformed-input, preservation, and normal-release
qualification gates.

The external diagnostic build passed `cargo check` and the release build with
the command above. The focused SPRM parser tests, census guard/reset/error
controls, harness compile-only test, and the complete `litchi-doc` diagnostic
feature suite passed after the fixture overlay (`1130` tests passed, `2`
ignored). No executable was invoked to produce a formal report.
