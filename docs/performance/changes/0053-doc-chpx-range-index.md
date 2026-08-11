# Change 0053: native DOC CHPX range index

Date: 2026-08-11

Production control: `fcc89cbeaab2c38cd11d9bb5cfb634f9f2e340f2`

Scope: private native DOC character-run lookup only. iWork/IWA crates were
explicitly excluded.

## Hypothesis and change

The post-0051 profile attributed 7.30-7.86% of complete-process self cycles to
`ParagraphExtractor::extract_runs`. Annotated profiles placed about 80-82% of
that frame in `ChpBinTable::runs_in_range`, which linearly filtered every CHPX
run for every paragraph. The 512-paragraph generated document therefore did
quadratic overlap tests during paragraph enumeration and during the same
complete semantic verification outside the timer.

`ChpBinTable::runs_in_range` now binary-searches the first run whose end is
strictly after the query start, then scans only matching runs until starts
reach the query end:

```rust
let first = runs.partition_point(|run| run.end_cp <= start_cp);
runs[first..]
    .iter()
    .take_while(|run| run.start_cp < end_cp)
```

This changes each lookup from `O(number of runs)` to
`O(log(number of runs) + matches)`. Parsing already sorts by `(start_cp,
end_cp)`, removes contained runs, clamps partial overlaps, and rejects empty
runs. Consequently retained starts are ordered and retained ends are strictly
increasing, making both predicates monotonic. The returned references and
their order are unchanged, including the established behavior for empty and
reversed queries and `u32` boundaries.

There is no new storage, allocation, public API, cache, runtime, lock,
dependency, unsafe code, parser leniency, output change, or validation
handoff. Formatting objects, direct `grpprl`, style cascading, fields,
pictures, comments, glossary parsing, exact patches, and both final readbacks
retain their former owners.

## Matched latency evidence

The frozen release binaries have SHA-256:

- control: `14223b6a37388507a26b0c03d11f4d2605f2c3658321aee6d2583644bc6fb1bd`;
- candidate: `172d31d20fc7205c97dc78b6377c1e4368ba48f797610b67b7f70ff452b8c438`.

Both use the unchanged standalone harness, release profile, Rust 1.95.0,
Linux 6.8.0-101-generic, the Rust system allocator, and CPU 11 pinned with
`taskset`. The fixed large DOC is 97,792 bytes with 512 paragraphs and SHA-256
`3d96764fe48e213b972ff5921df183dab9e8bfc8c8e751bcf3bf20190de4fec6`.
Its 81,920-byte `WordDocument` stream has SHA-256
`33e6cd70a45181c28d4a3e7bfa4e7817bd82d7b2e89e39437a589243abdc38eb`.

The primary `doc_semantic_list_paragraphs` measurement used 50 warmups and 500
samples in each of five control/candidate and five candidate/control pairs.
Pooling raw samples gives 5,000 observations per state while balancing binary
order. Corpus construction and complete semantic verification remain outside
timing.

| Large DOC paragraph list | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 454.100 us | 358.414 us | **-21.07%** |
| mean | 464.207 us | 367.062 us | **-20.93%** |
| p95 | 533.146 us | 426.523 us | **-20.00%** |
| p99 | 667.327 us | 515.229 us | **-22.79%** |

The approximate independent-sample 95% interval for the mean delta is
`[-21.29%, -20.57%]`. All ten matched pair comparisons improve: p50 deltas
range from -23.82% to -19.12%, and mean deltas range from -24.82% to -18.11%.
The reports also contain 5,000 tiny observations per state; their 1.963 ->
2.003 us p50 movement (+40 ns, +2.04%) is retained as a neutral smoke result.

Large guard reports use two forward and two reverse pairs with 30 warmups and
300 samples per leg, or 1,200 pooled observations per state.

| Guard | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| open | 358.949 us | 328.265 us | -8.55% | -8.04% | -12.43% |
| one paragraph | 471.689 us | 371.788 us | **-21.18%** | -20.93% | -23.61% |
| full text | 0.921 us | 0.921 us | 0.00% | -3.26% | -3.64% |
| exact no-op edit/save | 224.528 us | 225.821 us | +0.58% | +0.49% | +2.47% |
| one edit/save | 921.907 us | 908.299 us | -1.48% | -1.20% | +0.66% |

The no-op p99 moved +3.10% on a 270 us tail and the changed edit/save p99 moved
+1.73%; both p50/mean results and their p95 guard remain within the 3% gate.

## Attribution and resources

Matched 3,000-sample `perf record` processes captured 8,393 control and 7,534
candidate samples. No lost-sample warning was reported. Kernel symbols are
restricted on this host, but userspace DOC frames resolve:

| Self-cycle frame | Before | After |
|---|---:|---:|
| `ParagraphExtractor::extract_runs` | 7.56% | 1.23% |

The surviving frame includes matching-run property cascading; the full-vector
range scan is no longer present.

Matched whole-process `perf stat` A/B/B/A used 50 warmups and 3,000 measured
lists per leg. Pooling the two legs per state gives:

| Counter | Delta |
|---|---:|
| task clock | -9.12% |
| cycles | -8.57% |
| instructions | -8.92% |
| branches | -10.89% |
| branch misses | -11.50% |
| cache references | -37.26% |
| cache misses | -4.02% |
| page faults | unchanged |
| CPU migrations | 0 -> 0 |

These process counters include corpus setup and the post-timer verifier, so
their reductions are intentionally smaller than the timed paragraph-list
result.

Heaptrack used two warmups and 20 measured lists per state:

| Whole-process metric | Before | After |
|---|---:|---:|
| allocation calls | 748,147 | 748,147 |
| temporary allocations | 355,439 | 355,439 |
| peak heap | 6.63 MB | 6.63 MB |
| Heaptrack RSS | 18.49 MB | 18.53 MB |
| leaked bytes | 544 B | 544 B |

Uninstrumented GNU Time A/B/B/A reports the same 30,848-30,976 KiB maximum-RSS
range in both states and zero major faults. The change adds no allocation and
does not alter the memory envelope.

Raw ABBA reports, profile summaries, counter reports, RSS reports, and binary
provenance are indexed by
[`doc-chpx-range-sha256.txt`](../results/doc-chpx-range-sha256.txt).

## Correctness and quality gates

New differential tests compare the indexed iterator with the former scalar
predicate over empty and singleton tables, adjacent runs, gaps, all/no
matches, exact starts/ends, empty and reversed ranges, and zero/`u32::MAX`
boundaries. A separate test proves exact reference identity and ordering.
Existing parser normalization tests and complete DOC suites retain coverage of
overlapping physical CHPX entries, mixed compressed/Unicode pieces,
formatting/styles, fields/pictures, comments, tables, revisions, glossary,
malformed FKP/BTE input, exact patch/inverse, preservation, and Word/
LibreOffice fixtures.

Verification completed:

- `litchi-doc --all-targets --all-features`: 961 unit tests passed, two
  fixture-dependent tests remained ignored, and every integration/example
  target passed;
- warning-denied all-target/all-feature DOC Clippy passed, including the
  previously requested deprecation cleanup already committed in `1194fbc7f`;
- warning-denied DOC rustdoc passed after three inherited public-to-private
  intra-doc links were converted to plain private-module references;
- all 32 standalone harness tests and warning-denied all-target Clippy passed;
- the DOC libFuzzer target compiles;
- formatting and `git diff --check` pass.

## Remaining work

This index removes repeated CHPX full-vector scans. It does not alter CFB
publication, PAPX/CHPX parsing, character-property cascading, edit
serialization, security policy, or the mandatory strict owner and independent
public-reader reopens. Further OLE2 work requires fresh attribution to a
distinct owner. The measured RTF parser-block reservation and ODS durable-blob
ownership opportunities remain separate tranches; OOXML multi-Part source
publication also remains a wider design. iWork/IWA stays excluded while its
crates are modified independently.
