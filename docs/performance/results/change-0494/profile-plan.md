# DOCX opened-document edit profiling plan (0494)

This plan profiles the existing source-backed DOCX one-edit/save route while
the provider matrix is being extended. It is an attribution plan, not an
optimization conclusion. The profile must establish which work is serial and
which work is merely visible because the process is small before any change is
proposed.

The route under test opens the fixed 0188 media-rich DOCX, replaces one existing
main-story paragraph, commits the edit, publishes one sequential output, and
drops the package. The corpus is 16,793,036 archive bytes, 20 ZIP members, 200
paragraphs, and eight 2 MiB media members. Its archive SHA-256 is
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.
The profile is scoped to this corpus and this one-edit operation.

## Provider cells

The current `docx-edit-provider` selector accepts the following stable
provider names and controls. The emitted report records the selected provider
and its exact source-read counters; the profiler binds the complete argv in its
receipt.

| Cell | Required selector | Source behavior | What it proves |
| --- | --- | --- | --- |
| owned exact | `--provider owned` | caller-owned immutable bytes with exact positional reads | in-memory provider baseline |
| file warm | `--provider file` | `FileSource` opened on the staged corpus after staging | recent-file observation; no cold-cache claim |
| bounded short read | `--provider short --short-range 4096` | caller `ReadAt` returns bounded short ranges while preserving the same source bytes | sensitivity to short-read fragmentation; no network claim |

The provider factory, corpus generation, file staging, source hash, output path
setup, and trace-buffer reservation are setup. They must be outside the
operation clock reported by the Rust selector. This unmanaged route does not
construct a managed `ExecutionContext`; no managed-context setup claim belongs
in this profile.
The process-level profiler necessarily sees setup as well; this is why the
operation-local report is authoritative for lifecycle latency and the hardware
counter profile is labelled whole-child.

The measured lifecycle is one fresh opened source-backed package, one existing
paragraph replacement, one commit, one sequential sink publication, and package
and document destruction. The output remains live until after the lifecycle
clock. Output hashing, per-row semantic reopen, the preflight patch
replay/inverse/stale/foreign-source checks, media preservation checks, and
report serialization are outside that clock. The per-row commit identity check
and source-version check remain part of the operation's correctness envelope.
A row must report the boundary explicitly rather than folding these checks into
the edit/save latency.

The managed read-ahead policy is not a hidden fourth provider. The current DOCX
transaction path has a typed refusal when an owned edit snapshot would escape a
managed reservation. Until a separately reviewed bounded edit snapshot exists,
the profiler records that refusal as an unavailable/typed outcome and does not
turn it into a zero-time or copied-bytes result.

## Operation-local counters

The Rust report should retain per-sample counters with `available`, `value`, and
`scope` (or an equivalent typed unavailable record). Counters are requested to
be operation-local, so they can identify work that `perf stat` cannot attribute
to a library phase:

* phase elapsed times for mandatory open/catalog, main-document load and XML
  snapshot, edit planning, commit/patch construction, publication, and package
  and document drop;
* source `ReadAt` calls, requested bytes, returned bytes, short reads, maximum
  request, and the offset/request/returned range distribution;
* source-version callback/check count and the version identity before and after
  the operation;
* main-part decompressed bytes, changed XML bytes, raw unchanged bytes copied,
  changed-member compressed bytes, output bytes, sink write calls, and write-size
  distribution;
* source-artifact fingerprint/hash calls and bytes, if the selected path asks
  for one; the report must distinguish a fingerprint from a source-version
  check and must not infer one from the other;
* cache loads/hits/failures, retained bytes and entries, budget input/memory/
  object deltas, and allocation metrics when the allocator executable is used;
* post-clock reopen/readback status, output SHA-256, unchanged member/media
  hashes, one-operation commit diagnostics, and the preflight patch
  replay/inverse/stale/foreign oracle results.

A counter is not required to exist in the implementation merely because this
plan names it. If a path does not perform that operation, the report should
say `not_applicable`; if instrumentation cannot observe it, it should say
`unavailable` with a reason. It must never publish a guessed zero. In particular,
raw-copy bytes, fingerprint bytes, and reopened bytes are different quantities.

The current `docx_edit_provider_v1` report already exposes lifecycle latency,
output bytes and sink writes, materialization count, source version before and
after, provider-specific logical/physical range counters, short/empty reads,
cache materialization evidence, and the preflight/output oracles. Owned has no
read wrapper, so its logical and physical read counters are `unavailable` by
design; file-warm has logical `FileSource` counters and no separate physical
counter; short-read has logical adapter counters, underlying physical counters,
and range-adapter counters. The current report does not expose phase spans,
raw-copy versus changed-member bytes, decompression/recompression bytes,
fingerprint-call bytes, cache hit/failure/retained-byte observations, or
managed budget deltas. Those fields are therefore `unavailable` or
`not_applicable` for the initial profile unless the provider coder adds
operation-local instrumentation without changing the clock. The whole-child
`perf`/`strace` values must not be used to fill those missing fields.

The profiler delegates report acceptance to the canonical
`measure.validate_report` implementation with `role="normal"` and the
corresponding canonical arm (`owned`, `file-warm`, or `short`). That validator
owns the schema, exact route scopes, corpus and output identity, limits, row
invariants, provider controls, preflight oracles, binary identity, source
revision, and sample/warmup checks. The profile helper retains only the
validator's returned identity and hashes; it does not maintain a second report
schema oracle.

Before starting a child, the helper calls canonical `measure.load_builds()` and
`measure._load_protocol(builds)`. The retained summary carries both build
receipt hashes, each retained executable and gate receipt binding, the shared
source manifest, the protocol hash and its build/source binding, and the
authenticated normal revision. This makes a profile result a single binding
of retained build, source, gate, and frozen protocol custody. The exact source
archive identity remains 16,793,036 bytes, 20 members, SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.

The report's `limits` object is a description of the current unmanaged route,
not a measured budget. It contains the full `ReadLimitsRecord`, cache maximum
bytes and entries, `resource_budget.managed: false` with all observed budget
delta fields absent, `max_tracked_ranges: 131072`, and
`sink_max_write_bytes: 1048576`. The canonical validator verifies that shape
and the profiler retains the complete raw report. It does not turn those
bounds into input, memory, object, or work consumption.

## Expected serial work to test

The following work is expected from the current call graph and is a set of
profiling hypotheses. It is not an optimization assumption; the local counters,
sampled stacks, and syscall trace must confirm or disprove each item.

1. `SourceBackedPackage::from_read_at` performs mandatory ZIP/OPC indexing,
   content-type and relationship/catalog validation, and source identity setup.
   This work is needed before the selector can resolve the main Part.
2. `edit_document` obtains the main Part and materializes its decoded payload,
   checks the DOCX content type and execution budget, scans the XML, and creates
   the immutable edit snapshot. This is separate from opening and must be
   counted separately even when a cache makes a later read a hit.
3. The semantic replacement and commit validate the selected paragraph and
   construct the changed XML/patch. Patch or snapshot digest work belongs to the
   commit phase when the implementation performs it.
4. Publication enters the package's exact-publication path, which invokes the
   read-ahead fence when a package was opened with an explicit read-ahead
   policy. This selector uses the unmanaged compatibility constructor, so its
   normal route has no managed read-ahead window to disable. It reads the
   original main Part for the exact no-op/changed decision, validates the
   physical overlay topology, and prepares the preservation traversal.
5. The preservation writer walks the source ZIP framing, emits the changed
   member, recompresses its replacement as required, and raw-copies every
   untouched member to the sequential sink. Its source reads, decompression,
   compression, output writes, and source/version fences must remain separate in
   the report.
6. Any complete-artifact fingerprint or final source check is a distinct serial
   operation. The profiler must measure its bytes and calls when present rather
   than attributing them to raw copying or the semantic edit.
7. Package/document drop releases source caches and budget reservations. A drop
   cost is part of the lifecycle only if the selector's operation clock includes
   it; post-clock reopen and full preservation oracles are reported separately.

The large media members should be unchanged and raw-copied in this scenario.
That is a correctness expectation to verify from the report and output oracle,
not permission to assume that no media bytes were read or written. The output
sink may receive all unchanged media bytes even when no decompression occurs.

## Exact profile commands

Build and source custody are owned by the 0494 build protocol. After the final
normal and allocator executables have been retained and the protocol is
frozen, run the following from the repository root. `profile.py` authenticates
both build receipts, their gate receipts, the shared source manifest, and the
protocol before accepting the normal executable. It pins the process to
advisory CPU 2 and writes an immutable subdirectory for each provider and
tool. The profiler uses five measured iterations and two warmups for tool
stability; formal latency evidence uses the separate 30-sample matrix.

```sh
python3 -B docs/performance/results/change-0494/profile.py \
  --build-record docs/performance/results/change-0494/build-normal.json \
  --provider owned --provider file-warm --provider short-read \
  --samples 5 --warmup 2 --cpu 2 \
  --record-owned --record-samples 100 --record-warmup 3 --record-frequency 199 \
  --output-dir docs/performance/results/change-0494/profiling-r1
```

Run that command through the existing source-custody gate so the whole profile
batch holds the advisory CPU lock and retains terminal stdout/stderr:

```sh
python3 -B docs/performance/results/change-0494/gate.py profile-r1 \
  python3 -B docs/performance/results/change-0494/profile.py \
  --build-record docs/performance/results/change-0494/build-normal.json \
  --provider owned --provider file-warm --provider short-read \
  --samples 5 --warmup 2 --cpu 2 \
  --record-owned --record-samples 100 --record-warmup 3 --record-frequency 199 \
  --output-dir docs/performance/results/change-0494/profiling-r1
```

The gate wrapper is mandatory for the retained run; `profile.py` deliberately
does not take a second lock. Each observer receives a private `TMPDIR` below
`TEMP/managed/profile-r1-<label>/<label>/tmp`. After its process group
terminates, the helper calls canonical `measure._cleanup_private`, writes its
exact `docx-edit-provider-private-cleanup-v1` receipt into the provider output,
and fails closed if any managed scratch remains. The path check rejects
symlinks and scratch outside `TEMP/managed`.

The command expands to one `perf stat` and one `strace -f -c` child per
provider. These are the exact command shapes, with `BINARY`, `REV`, `OUT`, and
`CPU` resolved by the script from the authenticated build record:

```sh
/usr/bin/taskset -c CPU /usr/bin/perf stat --no-big-num -x, \
  -e task-clock,cycles,instructions,branches,branch-misses,\
L1-dcache-loads,L1-dcache-load-misses,LLC-loads,LLC-load-misses,\
minor-faults,major-faults,context-switches,cpu-migrations,page-faults \
  -o OUT/perf-stat.csv -- \
  BINARY docx-edit-provider --provider owned \
  --samples 5 --warmup 2 --source-revision REV \
  --output OUT/perf-report.json

/usr/bin/taskset -c CPU /usr/bin/strace -f -c -o OUT/strace-summary.txt -- \
  BINARY docx-edit-provider --provider owned \
  --samples 5 --warmup 2 --source-revision REV \
  --output OUT/strace-report.json
```

The retained helper also takes one owned-source call-stack profile under the
same CPU lock. It uses 100 workload samples after three warmups, a 199 Hz
`cycles:u` event, and DWARF callchains. This is attribution evidence and is
deliberately limited to the owned cell so it does not become a second formal
matrix:

```sh
/usr/bin/taskset -c CPU /usr/bin/perf record -F 199 -e cycles:u \
  --call-graph dwarf -o OUT/owned/record/perf.data -- \
  BINARY docx-edit-provider --provider owned --samples 100 --warmup 3 \
  --source-revision REV --output OUT/owned/record/perf-record-report.json
/usr/bin/taskset -c CPU /usr/bin/perf script --header --demangle \
  -F comm,pid,tid,cpu,time,period,event,ip,sym,dso \
  -i OUT/owned/record/perf.data > OUT/owned/record/perf-script.txt
```

The retained `profile-r1` record completed successfully and contains the
workload's `perf.data`; its first script export failed because `cycles:u`
records did not carry a CPU attribute while the field list requested `cpu`.
`recover_profile.py` reprocesses that existing file and omits only the
unsupported field. It authenticates the original profile summary, retained
`perf.data`, canonical report, record terminal, original profile gate, helper
archive, build pair, and protocol before writing a separate recovery
directory. It never launches the DOCX workload or edits `profiling-r1`.
If the checkout has advanced since the original workload, the recovery records
the current source manifest and its mismatch with the retained workload source
instead of presenting the re-export as a new source-bound measurement. The
recovery receipt records current source manifests before and after the
postprocessing and reports any concurrent change; the original build-bound
source remains the authenticated workload source. A gate source-change flag is
therefore an observation of concurrent checkout activity, not evidence that
the retained binary or `perf.data` was rerun.

Run the recovery under the same gate lock after the helper source is retained:

```sh
python3 -B docs/performance/results/change-0494/gate.py profile-recovery-r1 \
  python3 -B docs/performance/results/change-0494/recover_profile.py \
  --profile-dir docs/performance/results/change-0494/profiling-r1 \
  --build-record docs/performance/results/change-0494/build-normal.json \
  --gate-receipt docs/performance/results/change-0494/validation/profile-r1.json \
  --helper-archive docs/performance/results/change-0494/profiling-r1-helper-sources \
  --cpu 2 \
  --output-dir docs/performance/results/change-0494/profiling-recovery-r1
```

The recovery command is exactly:

```sh
/usr/bin/taskset -c CPU /usr/bin/perf script --header --demangle \
  -F comm,pid,tid,time,period,event,ip,sym,dso \
  -i profiling-r1/owned/record/perf.data
```

The resulting folded stacks and stack summary are a re-export of the retained
sample stream. Their periods remain statistical weights and cannot be counted
as a second workload or a fresh timing sample. Existing `strace` reports in
`profiling-r1` are already available and retained; the recovery does not rerun
or reinterpret those valid observer results.

The first derived observer summary had two parser defects: it treated perf CSV
field 3 (running time in nanoseconds) as the running percentage, and it only
accepted strace rows that included an explicit errors field. The immutable raw
files are retained, so `recover_observer_counters.py` derives a separate
corrected receipt without launching a provider or an observer:

```sh
python3 -B docs/performance/results/change-0494/gate.py profile-counter-recovery-r5 \
  python3 -B docs/performance/results/change-0494/recover_observer_counters.py \
  --profile-dir docs/performance/results/change-0494/profiling-r1 \
  --recovery-summary docs/performance/results/change-0494/profiling-recovery-r3/recovery-summary.json \
  --recovery-script docs/performance/results/change-0494/profiling-recovery-r3/perf-script.txt \
  --build-record docs/performance/results/change-0494/build-normal.json \
  --gate-receipt docs/performance/results/change-0494/validation/profile-r1.json \
  --helper-archive docs/performance/results/change-0494/profiling-r1-helper-sources \
  --recovery-helper-archive docs/performance/results/change-0494/profiling-recovery-r3-helper-sources \
  --output-dir docs/performance/results/change-0494/profiling-counter-recovery-r5
```

The receipt preserves each provider's raw and old parsed values, applies the
perf field contract `value,unit,event,running_time_ns,running_percent`, accepts
the strace contract with an implicit zero errors column, and rejects unexplained
rows or totals that do not conserve. The corrected whole-child call totals are
13,443 (owned), 127,427 (file-warm), and 2,332,877 (short-read), with one
reported error in each. The corrected cycles running percentages are 83.00% for
all three providers; the running nanoseconds remain separate values. Raw zero
L1 counters are retained without a claim that cache misses are proven to be
zero.

The same receipt reclassifies the retained r3 perf-script text. It binds the
3,239,852-byte export and its 1,382 samples to the authenticated `perf.data`
and report. The corrected stream has 24,761,902,814 total period,
15,644,544,271 period under bare `run_sample`, and 492,494,338 period in the
`publish_docx_source_edit` ancestor subset. Because the export contains bare
symbols, these phase labels carry `bare_run_sample_ambiguous`; they are
statistical callchain weights, not exact phase costs. The stack and observer
sections remain whole-child evidence and cannot replace operation-local report
counters.

The final correction gate receipt records `source_unchanged: true`. An earlier
retry observed a concurrent edit to `crates/litchi-odg/src/package/snapshot.rs`;
that retry is retained separately as custody history. The final correction
observed an unchanged source manifest, launched no executable, and remains
bound to the retained build source, report, `perf.data`, and archived capture
helper.

`profile.py` retains both process-group started/terminal receipts, stdout and
stderr, the `perf.data` hash, the raw script export, deterministic folded
stacks, and `perf-stack-summary.json`. Its stack summary filters sampled frames
containing `docx_edit_provider::run_sample`: a
`publish_docx_source_edit` ancestor is the operation-body candidate,
`verify_docx_source_edit_output` or `sha256_hex` is the post-clock output
oracle, and other `run_sample` frames are setup/teardown. `prepare` frames are
preflight; remaining frames are process setup or unclassified. The classifier
uses sample periods as statistical weights, not exact phase costs. If DWARF
stacks or the PMU are unavailable, the retained result says `unavailable`; a
nonzero target workload remains `failed` and is never relabelled as an
unavailable profiler. For `perf stat`, only perf event-setup diagnostics such
as `No permission to enable`, event syntax/open failures, or unsupported PMU
events trigger the reduced core-event retry. A generic target `permission
denied` or other application error remains `failed`.

The current profiler applies the same distinction to strace: process launch
failure and explicit ptrace/strace setup or unsupported-syscall diagnostics are
`unavailable`, while a traced target's ordinary nonzero exit, including target
permission errors, is `failed`. The retained r1 workload remains bound to the
archived profiler helper in `profiling-r1-helper-sources`; the future strace
classification is recorded as a helper change and was not applied retroactively
to those captures.

For the file-warm label, replace `--provider owned` with `--provider file`. For
the short-read label, use `--provider short --short-range 4096`. The two tools run
separate fresh processes; their process elapsed time is not compared with each
other because `perf` and `strace` add different observer overhead. Use the
report's operation-local elapsed vectors for latency, and use the tool output
for counters and syscall evidence.

`profile.py` must refuse an existing output directory or artifact, retain the
resolved argv, cwd, environment subset, build-record hash, binary hash, source
revision, tool versions, exit status, and hashes of every output. A nonzero
target exit, missing report, changed source manifest, or malformed report is a
failed profile. An unavailable profiler is a typed `unavailable` result and is
not a failed DOCX run; the target still runs once without that observer only if
the command explicitly records that fallback and does not merge its timing with
observed profiles.

## Counter interpretation

`perf stat` is whole-child evidence. It includes corpus generation, file staging
when the selected provider stages internally, process startup, Rust allocator
initialization, report JSON serialization, and all warmup and measured rows.
It does not identify which phase generated a cycle, cache miss, or branch miss.
The report phase counters are the only basis for an operation-level attribution.

`strace -f -c` is also whole-child evidence. Retain calls, errors, total time,
and the named syscall rows (`openat`, `read`, `pread64`, `write`, `pwrite64`,
`lseek`, `mmap`, `munmap`, `fstat`, `fdatasync`, `fsync`, and `rename` when
present). It does not report bytes transferred, decompressed bytes, or source
range sizes. Source `ReadAt` counters and sink counters provide those quantities.

The profile summary may derive IPC as `instructions / cycles`, branch-miss rate
as `branch-misses / branches`, and cache-miss rates only when both numerator and
denominator are numeric and supported. Unsupported, not-counted, or permission-
denied events remain `unavailable` with raw text retained. They are never
converted to zero and never silently dropped from the event inventory.

## Confounding and scope

* `owned` measures an in-memory source and is not a disk result.
* `file-warm` uses a recently staged regular file and is not a cold-cache result.
  It must not be described as filesystem latency or remote I/O.
* `short-read` is a caller adapter experiment. It does not model a network,
  bandwidth, packet loss, or server scheduling.
* The fixed corpus is synthetic and does not establish Word, LibreOffice, or
  another producer/version result.
* Whole-child counters include setup and serialization. The Rust operation
  report must keep setup and oracle work outside its lifecycle clock.
* A profile is descriptive until paired with the formal provider matrix,
  correctness/oracle verification, and an independent review. It cannot by
  itself justify an optimization, a 10x claim, or a general CRUD conclusion.

The initial tool probe is retained in
[`profile-availability.json`](profile-availability.json). It used only
`perf stat ... -- true` and `strace -c -f -- true`; those successful probes say
nothing about the DOCX workload. Later profiles must retain their own event
availability and raw outputs because PMU permissions and syscall support can
change between runs.
