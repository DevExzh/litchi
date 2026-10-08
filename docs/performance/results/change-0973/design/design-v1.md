# 0973 — preregistered Float/e1 instruction-variation diagnostic

**Status:** design only; no build, Cargo invocation, Office workload, or new
profile capture has started. This is a diagnostic plan for the retained 0967
baseline executable. It does not qualify a latency result or authorize a
production change.

The purpose is to test whether the process-level instruction variation observed
in the retained Float/e1 analysis is concentrated in the profiled marker
lifecycle and its decoded owners. The analysis uses all 12 children, with no
fast/slow selection, no replacement child, and no causal or adoption claim.

## Accepted constraints and custody

The current repository is HEAD `52b225d88e260b5430fdba570ee9d3dc479e41de`.
The current GOAL and accepted ADR set were revalidated before this design:

- GOAL SHA-256: `bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`.
- ADR README SHA-256: `2af0e5814378c2765c12213e495d02c2895a27f142bc618ca1f3037824bbe028`.
- The 32-entry revalidation receipt is
  `/home/zhuhe/litchi-performance-0973/receipts/adr-goal-revalidation-0973-v1.json`,
  SHA-256 `80dc8af647aa4992f2281262e7be5bde40ac192aa1d392b0fc5d6525052d7d90`.
  It records no source, build, or workload mutation.

The executable is the unchanged retained 0967 historical baseline:

- `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile`
- 47,909,224 bytes, SHA-256
  `0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb`
- Build ID `e5c457384eeebd057932c444efda77618f137399`.

The fixed fixture is `FloatingPictures.doc`, one edit, default retention:

- `/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc`
- 335,360 bytes, SHA-256
  `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`.

The harness is retained for custody and CLI identity only; it is not rebuilt:

- `tools/perf-cfb-profile/src/lib.rs` SHA-256
  `b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc`.
- `Cargo.toml` SHA-256
  `41dee12e4eb7d1e3853a33a7743afdf6b5d6519d5e7456e61ce4978aa9177945`.
- Source manifest SHA-256 `f1818af5c306fb6ea01b59ca8df29710e773386b246071202bd4aa896f4e8e2b`;
  source-file manifest SHA-256
  `99e65d43e0feed89133eec60b095130d68058a071a02e8a15b05dea01209bb6c`.

The ASLR decoder must use the sealed v6 0972 result as its regression oracle:

- final 0972 diagnosis SHA-256
  `80c54eeb6116b585b9088399c14eef2583b73202faa4a364f0983968753eff9c`;
- v6 author seal SHA-256
  `a8be68fb63e7460d839ca23ab2457b2415762da76d8a92d26e8640281a7bde1c`;
- corrected raw scan v2 SHA-256
  `6ef76bc01cc7e298c924eda95cfd29dc0cc070ac3e3de4cbe9d87a7cf1b93a9b`.

These constraints implement the accepted preservation, bounded-resource,
explicit-execution, ownership, verification, and evidence rules in ADRs 0001,
0002, 0003, 0005, 0006, 0008, 0024, and 0031. No production code, public API,
source manifest, or binary is changed for this design.

## Fixed capture matrix

The prospective matrix is one cell, `docfloat/e1/default`, with 12 fresh child
processes launched serially in order 01 through 12 on CPU 31. Each child uses
five warmups followed by 80 measured samples. The harness marker
`doc_retention_probe_0942_cfb_profile::profiled_measured_case` is called 85
times per child: the five warmups and 80 measured calls are one combined marker
lifecycle for perf attribution. The harness p50 and p95 are computed from the
80 measured `elapsed_ns` rows only.

The operation-local harness scope is recorded exactly from the retained report:
source clone/open, one edit with the requested replacement, retention handling,
commit, output `Vec` copy, and local lifecycle destruction. Source-file I/O,
correctness checks, hashing, JSON serialization, launcher work, and perf
bookkeeping are outside that `Instant` scope. Perf counters cover the launched
child process tree, so marker periods and p50 are deliberately separate
observations.

The fixed record command shape is:

```text
taskset --cpu-list 31 /usr/bin/perf record \
  -e instructions:u -c 1000000 --call-graph dwarf,32768 \
  --period --timestamp --sample-cpu --buildid-all --no-buildid-cache \
  --user-callchains -o <child>.perf.data -- \
  /usr/bin/python3.14 /home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py \
  --metadata <child>.launch.json --requested-affinity "[31]" -- \
  <retained-binary> --case docfloat --input <FloatingPictures.doc> \
  --edits 1 --retention default --warmup 5 --samples 80
```

The runner records the exact argv, environment (`LC_ALL=C`, `LANG=C`,
`TZ=UTC`), resolved executable paths and hashes, launcher PID/exec PID, actual
CPU mask, terminal status, report, and all raw capture paths. It performs no
global affinity, governor, cgroup, cache, or perf setting changes.

## Decoder staging gate

`staging_protocol` owns the external ASLR generalization before root GO. It may
inspect or compile an external decoder helper, but it must not build the
retained Rust target or launch the harness. The helper must satisfy all of the
following for every child:

1. Read the exact target `PERF_RECORD_MMAP2` rows from `perf script -D`, require
   the target path and retained Build ID, and identify the offset-zero target
   map. Derive that map's load bias from the child data; no runtime address,
   PID, or ASLR base may be hardcoded.
2. Parse the retained ELF `PT_LOAD` headers. Derive the executable `PF_X`
   interval from `load_bias + p_vaddr .. load_bias + p_vaddr + p_memsz`, and
   derive symbol runtime addresses from the same child load bias.
3. Derive the perf comparator start from the executable mapping as
   `exec_map_start - page_down(MMAP2.pgoff)`. Require page alignment,
   consistent mapped load segments, and an unambiguous relation to the
   offset-zero load bias. Missing, aliased, or inconsistent mappings reject the
   child's decoded result.
4. Apply the v6 correction only to the exact retained DSO. The interposer uses
   the six-argument `dwfl_report_elf` ABI, forwards `fd` and `add_p_vaddr`,
   changes the target base to the child-derived ELF bias, and changes the
   returned module start only for a direct `/usr/bin/perf` caller to the
   child-derived comparator start. It captures the caller PC at the interposer
   and restores `errno`; non-target DSOs and Build ID aliases are untouched.
5. Run stock, no-op, and corrected offline replays on the sealed 0972 capture.
   Stock/no-op output and sample headers must remain byte-identical, the raw
   hash must remain unchanged, and the corrected replay must reproduce the
   sealed 0972 marker recovery (78 marker chains, 390,000,000 marker period).
   A derived relocation/ASLR synthetic control must recover the same symbols
   after translating the mapping metadata; the original sealed bytes remain
   immutable.

`staging_protocol` records the layout, ELF interval, derived bases, decoder
commands, helper/source hashes, synthetic-control result, and all unresolved
or truncated counts. `cpu_review` performs the final independent review of
that receipt, the storage arithmetic, and the no-claim boundaries. Root GO is
not valid until both reviews pass. The pre-existing 0973
`diagnostic-proposal-v1.md` is retained as staging context and is not silently
overwritten.

## Pilot and stop rules

The first child is the retained pilot. After its terminal packet and corrected
decode are complete, it must satisfy:

- at least 100 exact marker samples in the corrected target DSO;
- zero lost events;
- positive marker period share;
- report identity, fixture/output correctness, PID/CPU identity, Build ID,
  dynamic-layout receipt, and all decoder controls pass.

A first-child failure stops the run without retry and preserves the failed
artifacts. If child 02 through 12 fails the same gate or any terminal,
correctness, identity, or decoder check, the run stops immediately without a
replacement child. No partial run is reduced into an all-12 conclusion. A
successful child still retains its raw `perf.data`; no sample, warmup, owner,
or unresolved row is silently filtered.

The external output root is
`/home/zhuhe/litchi-performance-0973`. Raw captures and verbose decoder
artifacts remain outside the repository. No Office workload, native build,
Cargo command, or iWork operation is permitted in this design phase.

## Per-child measurements and owner accounting

The corrected decoder emits, for every child:

- total sample count, total `instructions:u` period, lost events, empty or
  truncated stacks, unresolved frames, CPU/PID/event identity, and Build ID;
- marker sample count and marker period across the full 85-call lifecycle;
- the 80-row harness p50/p95 and the exact list of measured `elapsed_ns` values;
- all fully qualified resolved symbols in marker-containing chains, keyed by
  DSO and symbol, with sample count and weighted event period.

For an inclusive owner `s`, the reducer adds one copy of a sample's period when
`s` appears anywhere in that sample's resolved marker chain. Repeated frames of
`s` in one sample are deduplicated; different owners intentionally overlap.
This yields `inclusive_owner_period[s]` and its descriptive share of marker
period. It is not valid to sum inclusive owners or call the result a phase
fraction. The fixed summary includes `profiled_measured_case`, `measure_one`,
`run_case`, `run`, and every other resolved target-DSO symbol; no owner is
chosen after inspecting the p50 results.

The exclusive leaf proxy is a separate map. For each marker sample, the reducer
examines the physical leaf, frame 0, before looking at its ancestors. If frame 0
resolves, its period is assigned to exactly one `(DSO, symbol)` bucket across
all DSOs, including libc allocation or memory symbols. There is no target-DSO
filter on this proxy. If frame 0 is absent, unresolved, or invalid, its period
is assigned to an explicit `unknown_leaf` bucket even when a later ancestor
resolves. Thus the sum of resolved leaf buckets plus `unknown_leaf` equals the
marker period. An optional separately named `deepest_resolved_ancestor` map may
describe the nearest resolved frame when the physical leaf is unknown; it is
never substituted for the leaf proxy. The leaf proxy is useful for comparing
one deepest observed instruction owner per sample, but it is not an exact timed
stage or source attribution. Inclusive maps and the exclusive leaf proxy are
never merged.

The reducer keeps per-child raw periods and normalized shares. Missing owners
are reported as zero observed coverage only for the descriptive association
table, with an explicit presence count; zero or absent coverage is not evidence
that the owner did no work. Unknown and truncated marker samples are retained
in separate buckets and are never reassigned.

## Preregistered analysis

After all 12 children pass, the analysis uses every child exactly once (`n=12`)
and computes p50 correlations against:

1. total process instruction period and its share;
2. marker period and marker-period share across the 85-call lifecycle;
3. each fixed inclusive-owner period/share;
4. each fixed exclusive-leaf-proxy period/share.

For every metric it reports Pearson and Spearman coefficients, the paired child
values, presence/unknown counts, and the full owner table. There is no
fast/slow cutoff, post-hoc child selection, replacement, or pooling of warmup
and measured `Instant` rows into a new timing metric. Event periods are
instruction-counter weights, not wall time.

For each child, define `D_i` as total sampled instruction period and `M_i` as
the period whose decoded ancestry contains the marker. Define `R_i = D_i - M_i`
and label it **marker ancestry not observed**. This residual can contain setup
or wrapper samples, but it can also contain work inside the marker lifecycle
whose marker ancestry was missing or truncated; it must not be interpreted as
an outside-operation bucket. The report includes the per-child conservation
check `D_i = M_i + R_i` and the descriptive variance ledger
`Var(D) = Var(M) + Var(R) + 2 Cov(M,R)`. The ledger describes observed
variation only and is not a causal work split.

The diagnostic interpretation is fixed in advance. A consistent association
between p50 and marker/owner periods narrows the variation to the observed
marker lifecycle and its decoded owner coverage. A strong total-period
association with flat marker/owner periods points toward variation in `R`, whose
marker ancestry was not observed. It may reflect process-tree or wrapper work,
or missing/truncated ancestry within the marker lifecycle, so the mechanism
remains unresolved. Weak or conflicting associations likewise leave the
mechanism unresolved. These are descriptive hypotheses for the next
investigation; none proves causality, an exact timed stage, a source owner, a
speedup, or production benefit. `stage_timed_commit` and `timed_iteration` are
not assumed to exist in the retained binary and are not invented if absent.

## Sampling overhead and storage check

The retained 0972 capture is the direct sizing reference: period 5,000,000,
16,384-byte stacks, 1,387 records, 16,624 bytes per record, and a 23,129,880
byte `perf.data`. Holding the approximately 6.935 billion instruction total
constant, period 1,000,000 projects about 6,935 records. A 32,768-byte stack
adds 16,384 bytes to the observed record, projecting about 33,008 bytes per
record and 228,910,480 bytes (about 229 MB) of raw data per child. Twelve
children project about 2.75 GB of raw `perf.data`; the 0970 retained
instruction range (about 6.99–7.50 billion) gives roughly 2.77–2.97 GB across
12 children.

The host currently has approximately 364.5 GB free on the filesystem containing
the external root. Before GO, the host receipt must be refreshed and reserve at
least 8 GB for raw captures, compact reports, decoder output, and failure
artifacts. Full verbose `perf -vvv` output is not collected for every child;
raw data plus the prescribed summaries are authoritative. A verbose diagnostic
is retained only for a decoder failure or the first-child forensic check.

The 0971 harmless capability receipt already proves `instructions:u` and
DWARF callgraphs through `dwarf,65528` on `/usr/bin/true` under perf 7.0.14,
so no new capability probe is required for this design. If host drift makes one
necessary, it is limited to the same harmless `/usr/bin/true` probe and must be
coordinated with the author before any native launch. The 1,000,000 period is
expected to create roughly five times the sampling interrupts of 0972 and the
32 KiB stack roughly doubles stack payload; this is accepted diagnostic
instrumentation overhead, not a latency comparison. Lost events and terminal
status remain hard evidence gates.

## Evidence and disposition

The existing 0973 staging files are preserved:

- `diagnostic-proposal-v1.md`, SHA-256
  `c149b0f5b95b6005b7bea0c513aad376c6a8ab32bfba7c2792a0973696c4117c`;
- `receipts/adr-goal-revalidation-0973-v1.json`, SHA-256
  `80dc8af647aa4992f2281262e7be5bde40ac192aa1d392b0fc5d6525052d7d90`.

This design is the preregistration artifact. It does not itself authorize a
workload. After staging_protocol and cpu_review seal the decoder and root GO is
issued, a future capture may run the fixed serial matrix. Any final report
must preserve failures, unresolved/truncated counts, owner overlap, and the
instrumentation boundary, and must state that the result is diagnostic only.
