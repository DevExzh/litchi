# Change 0502: ODG metadata open and semantic traversal evidence

The ODG batch adds inert page transitions, 3D shape owners, enhanced geometry,
image maps, contours, and glue points. It also adds a namespace-aware presence
gate so the enhanced-geometry and auxiliary-child validation passes are skipped
when the input contains none of their owned elements. Open also reuses parsed
content and named style definitions while resolving page transitions and
building the style inventory. This record measures these work eliminations and
the resulting package-open path; it does not claim a CRUD speedup.

## Protocol and provenance

The probe is [0502-odg-open](../probes/0502-odg-open/). It generates four
deterministic packages with `litchi-odf-common::PackageWriter`, then times
`litchi_odg::Drawing::from_bytes` followed by a typed page/shape traversal. The
package is generated and hashed outside the timed interval. Each child runs 25
warmups and 200 measured samples, pinned to CPU 2, in two serial repeats. The
before executable was built from clean `HEAD` `f3f9221171864fd48b02f7758dc2ebb1d14f9022`
in `/tmp/litchi-odg-before-0502`; the after executable was built from the
working tree containing the ODG batch, the presence gate, parsed-style reuse,
and boxed cold metadata. Both builds used
the probe's committed `Cargo.lock`, `cargo build --release --locked`, Rust
1.95.0, and Cargo 1.95.0.

The formal after timing executable has SHA-256
`fa50dd70ad275613f69bc1d4392a22e8e1756be29b9526e9b9c654dfba745380`; its
timed source manifest has SHA-256
`215ca88fd0098847bcc9d5c42262d309cae2493beedc2c67ac710b162af0342e` in
[`source-manifest-timing.sha256`](../results/change-0502/source-manifest-timing.sha256).
The final tree manifest, after probe formatting and whitespace-only source
updates, is `b23ac35c06f07096f7a703826fac112f215a7ece37b533a89c310376426a3f36`
in [`source-manifest.sha256`](../results/change-0502/source-manifest.sha256);
those post-capture edits were not retimed. The earlier pre-boxing timing
binary hash is retained in the reports as `caf9cc...`; its manifest is
[`source-manifest-pre-boxing.sha256`](../results/change-0502/source-manifest-pre-boxing.sha256)
alongside `after-pre-boxing/`. Executables were removed during scratch cleanup.
The raw machine-readable timing reports and
`/usr/bin/time -v` receipts are retained under
[`results/change-0502`](../results/change-0502/). The before and after package
hashes match for every corpus:

| Corpus | Pages × shapes/page | Input bytes | Input SHA-256 |
| --- | ---: | ---: | --- |
| `plain-small` | 4 × 64 | 2,305 | `d0586e7285ed874dc19be801fc94a5bcb43f1a195beac375ec71246276d18354` |
| `plain-large` | 32 × 256 | 49,143 | `1c8ba0a184b37a0b3cb5ed1c0a2ba987d665b25b21feffc46c3e6e6466e4bb84` |
| `metadata-small` | 4 × 32 | 2,652 | `7a11a8501defff985a9c2b4889b053e52af4b82cae42db588fbc9d6e6e9b6028` |
| `metadata-large` | 16 × 128 | 16,719 | `56ff7ee9e0a7868e79e4224395d7f9ef6ea7e5576c64e83b7a5be194a737963f` |

The host is x86-64 Ubuntu on an AMD EPYC 9R45 with 32 logical CPUs and a
7.0.0 kernel. Linux denied hardware counters (`perf_event_paranoid=4`), so
there are no cycles, instructions, cache, branch, or IPC claims. Process
maximum RSS and page-fault fields come from `/usr/bin/time -v`; they include
process setup and report serialization and are not operation-local allocation
measurements.

The capture can be replayed from either worktree with the recorded lockfile and
an isolated target directory. Build the probe with the following command, then
run the four corpora twice in each phase:

```sh
cargo build --release --locked \
  --manifest-path docs/performance/probes/0502-odg-open/Cargo.toml \
  --target-dir "$TARGET"
for repeat in r1 r2; do
  for corpus in plain-small plain-large metadata-small metadata-large; do
    LITCHI_GIT_REV=<revision> taskset -c 2 /usr/bin/time -v \
      -o "results/change-0502/<phase>/${corpus}-${repeat}.time.txt" \
      "$TARGET/release/odg-open-probe" --corpus "$corpus" \
      --warmups 25 --samples 200 \
      --output "results/change-0502/<phase>/${corpus}-${repeat}.json"
  done
done
python3 docs/performance/results/change-0502/summarize.py
```

Replace `<phase>` with `before` or `after`, and `<revision>` with the source
marker. The companion `shape-size` binary prints
`size_of::<litchi_odg::shape::Shape>()` and its alignment; those values are
retained in `layout.json` as a type-layout observation, not an allocator
measurement.

## Matched timing

Values are p50 milliseconds for clean HEAD → the final working tree, followed
by the percentage change. R1 and R2 are separate serial repeats.

| Corpus | R1 p50 before → after (change) | R2 p50 before → after (change) |
| --- | ---: | ---: |
| `plain-small` | 1.977128 → 2.093838 (+5.903%) | 1.963457 → 2.088838 (+6.386%) |
| `plain-large` | 60.441910 → 64.106069 (+6.062%) | 60.357609 → 63.948216 (+5.949%) |
| `metadata-small` | 1.251204 → 1.919377 (+53.402%) | 1.245544 → 1.933267 (+55.215%) |
| `metadata-large` | 18.046305 → 37.913401 (+110.090%) | 17.993619 → 37.991028 (+111.136%) |

The p95 changes are +5.925%/+6.493% for plain-small, +6.682%/+5.693%
for plain-large, +52.987%/+55.082% for metadata-small, and
+111.533%/+110.065% for metadata-large. The p99 changes are
+1.285%/+6.446%, +6.722%/+4.854%, +65.127%/+55.336%, and
+110.839%/+111.273%, respectively.

The retained `summary.json` also reports a deterministic 10,000-resample
bootstrap interval for each observed p50 change. It uses the probe's
upper-middle order-statistic p50 estimator for both the observed ratio and
each resample. The 95% intervals are +5.77–6.01% and +6.31–6.48% for
plain-small R1/R2, +5.95–6.18% and +5.87–6.00% for plain-large,
+53.22–53.55% and +55.02–55.39% for metadata-small, and +109.89–110.28%
and +110.92–111.35% for metadata-large. These intervals describe sampling
uncertainty among the 200 samples within each process run;
they do not estimate machine-to-machine, day-to-day, compiler, or cache-state
uncertainty. The bootstrap seed, sample count, metric, and interpretation are
machine-readable in `summary.json`.

The two work eliminations are useful but do not close the regression review.
The retained intermediate reports distinguish the unoptimized metadata
candidate from the candidate before boxing.
After cold metadata was boxed, the final clean-HEAD comparison remains
5.9–6.4% slower on plain p50 and 53.4–111.1% slower on metadata p50. The
retained metadata parser and the remaining per-shape representation cost are
therefore follow-up hotspots rather than evidence of an end-to-end improvement.

## Memory observations

Maximum resident set size in KiB for before → after is:

| Corpus | R1 | R2 |
| --- | ---: | ---: |
| `plain-small` | 6,872 → 7,352 (+6.98%) | 6,856 → 7,360 (+7.35%) |
| `plain-large` | 22,048 → 20,456 (-7.22%) | 20,756 → 22,528 (+8.54%) |
| `metadata-small` | 6,896 → 7,328 (+6.26%) | 6,892 → 7,608 (+10.39%) |
| `metadata-large` | 12,316 → 14,860 (+20.66%) | 12,372 → 14,604 (+18.04%) |

These are whole-child observations including setup and serialization. They do
not isolate operation-local allocations or resident memory.

The layout probe reports `Shape` at 792 bytes before the metadata batch,
936 bytes with inline optional metadata, and 800 bytes after grouping cold
metadata behind one optional box. Alignment remains 8 bytes. The final
per-shape increase over the old baseline is 8 bytes (+1.01%); the plain-large
corpus therefore retains 65,536 additional struct bytes instead of 1,179,648.
This arithmetic excludes allocator headers and child allocations.

A separate [heaptrack receipt](../results/change-0502/heaptrack-plain-large.json)
compares the pre-boxing and boxed candidates on the same plain-large corpus,
with CPU 2 affinity, 25 warmups, and 200 samples. Heaptrack 1.5.0 reports
174,853,447 allocation calls in each whole child and peak heap display values
of 13.69M → 12.57M (about -8.2%). Instrumented p50 differs by -0.71%; it is not
an uninstrumented latency result. Reported RSS includes profiler overhead and
increased from 23.29M to 24.10M. The representation is retained for its measured
heap reduction without claiming a latency or allocation-count improvement.

The compact receipt retains executable and raw-trace hashes plus exact summary
fields. Large raw heaptrack traces and scratch binaries were removed; the
receipt cannot substitute for inspecting those traces. Reproduce the bounded
comparison by building each candidate and running:

```sh
taskset -c 2 heaptrack -o /tmp/odg-heaptrack "$ODG_PROBE_BINARY" \
  --corpus plain-large --warmups 25 --samples 200 \
  --output /tmp/odg-heaptrack-probe.json
heaptrack_print /tmp/odg-heaptrack.<pid>.gz
```

Use the actual trace filename printed by heaptrack (compression suffixes may
vary). This whole-child capture includes generation, warmups, serialization,
process setup, and profiler overhead; it is not an operation-local allocation
measurement. It predates final formatting and inverse-only changes, separately
identified from the formal timing source manifest.

## Interpretation and remaining scope

The gate removes two unconditional full XML passes for inputs without the new
enhanced or auxiliary owners. Parsed-style reuse removes the duplicate style
definition scans in the open path. Their measured benefits are retained as
necessary enablers, with no broad performance claim. The remaining plain
overhead is consistent with the larger per-shape retained model and other new
metadata plumbing; the metadata lanes additionally exercise the new validation
and retained values. Any further change should capture a new before/after pair
with the source manifest recorded.

This probe does not cover filesystem cold/warm behavior, range sources, edits,
commit and save, durable patches, untouched-media copy-through, malformed or
adversarial inputs, native producer corpora, concurrency, or scaling. The
metadata candidate intentionally exposes more typed state than clean HEAD, so
its per-run semantic checksum is checked for repeat stability but is not
expected to equal the before checksum. The wider non-iWork performance goal in
[`docs/GOAL.md`](../../GOAL.md) remains open.
