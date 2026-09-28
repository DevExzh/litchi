# 0817 protocol review

This packet is a baseline-only measurement of the unchanged ordinary-save
harness at revision `953866d382c8e8248de2463ce317279fdd77d5b4`. It does not
make a production optimization, adoption, speedup, regression, or historical
timing claim.

## Corpus and custody

The three caller-named real inputs are fixed by path, size, and SHA-256:

| Format | Input | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| DOCX | `test-data/ooxml/docx/documentProperties.docx` | 23,503 | `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5` |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | 8,435 | `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4` |
| PPTX | `test-data/ooxml/pptx/shapes.pptx` | 68,822 | `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` |

Each input is classified once by its package main part: `word/document.xml`,
`xl/workbook.xml`, or `ppt/presentation.xml`. Enforce the 32 MiB input bound
and bounded ZIP/XML parsing. Reject missing, duplicate, ambiguous, malformed,
encrypted, unsafe, or traversal-bearing entries. Record the member inventory,
compression methods, sizes, CRCs, XML hashes, relationship/content-type graph,
edit target, and source hash.

Use the existing tracked `tools/perf-baseline/Cargo.lock` graph. Record its
exact external differences from the workspace lock and do not update either
lock. Production and tracked harness sources must remain unchanged at the
frozen base. Do not infer a Microsoft producer from package metadata: DOCX and
PPTX producer identity is unavailable; the XLSX provenance is retained as
tracked test-data provenance.

## Admission and output evidence

Run the exporter for three generated and three real corpus cases before any
qualification or timed selector capture. Each case has five policy outputs:
`default`, `full`, `file-only`, `no-sync`, and `stream`; the 30 outputs are
untimed controls. The `default` and `full` filesystem policies retain their
documented durability semantics. Reduced-durability and stream artifacts do
not substitute for the native default save.

The exporter’s reopen/hash/source checks are necessary but not sufficient. Run
`artifact_audit.py` as an independent ZIP/XML oracle and require its report to
have the expected schema, manifest identity, six cases, and `ok: true` before
qualification or native capture. For every output, verify:

- the ZIP opens under finite limits, has the expected format/main part, has no
  duplicate or unsafe member, and leaves the source byte-identical;
- the member inventory and expected edit closure are preserved;
- every XML member parses, and canonical XML is unchanged outside the edit
  closure;
- every untouched binary member is byte-identical after decompression;
- relationship, content-type, and package graphs remain valid; and
- source/output member hashes, changed-member set, policy identity, and reopen
  result are retained.

The independent semantic oracle is format-specific:

- DOCX output reopens, all source paragraph semantics remain, and exactly one
  final paragraph contains `litchi-perf-0638-ordinary-save`.
- XLSX sheet names/count, stored cell semantics, styles, shared strings,
  auto-filter, extension data, and worksheet structure remain unchanged except
  first-sheet `A1`, which becomes the marker.
- PPTX slide order/count and all shape text remain unchanged except the first
  admitted slide-0 shape, which contains the marker; media and unrelated parts
  remain unchanged.

The audit must bind the real case’s staged archive byte-for-byte to its
caller-named path. A typed edit refusal remains a refusal with its output and
reason retained; deterministic source equality is not a successful edit. The
qualification driver must link each selector to the audit case by format,
origin, source hash, edit target, edit outcome, and published output identity.
It must not silently substitute another corpus or continue an unadmitted
selector under the expected 12-selector denominator.

## Frozen timing matrix

The target matrix is 12 selectors: three real inputs × four phases. The phases
are independent observations and are not additive:

- `lifecycle`: open + semantic edit + save; destination preparation, readback,
  digest, and cleanup are outside the timer;
- `edit`: open is outside; semantic edit and commit only;
- `atomic_publish`: open/edit are outside; default full-durability save only,
  including temporary-file write, `sync_all`, rename, and parent-directory
  synchronization;
- `counting_publish`: sequential sink accounting for DOCX and XLSX; for PPTX,
  `Package::to_bytes` materializes the complete package and then performs the
  bounded sink operation. This is materialization-plus-sink evidence, not a
  streaming PPTX serializer or bounded-memory streaming claim.

Do not place semantic oracles inside a measured interval. Validate the output
outside the interval against the frozen artifact identity, retain failed
samples, and prove source preservation outside the interval.

### Native lane

Build and run `litchi-perf-baseline` with no Cargo features. The native binary
must not enable `allocator-metrics` or `ordinary-save-process-metrics`.

Use six counterbalanced blocks in order
`forward, reverse, forward, reverse, reverse, forward`, with 30 measured
samples and three warmups per selector, on CPUs 12–19. The target is 72
reports and 2,160 measured samples. Native elapsed values are the only latency
values used for the baseline timing table.

Every native sample is associated with the frozen selector, source hash, edit
outcome, and output reference. A failed operation or output mismatch remains
in the packet and prevents that selector from being reported as admitted.

### Observer and qualification lanes

Build `litchi-perf-baseline-alloc` with exactly the combined features
`allocator-metrics,ordinary-save-process-metrics`. The observer is diagnostic:
its elapsed values are instrumented and are never pooled with native latency.
Use two blocks in `forward, reverse` order, three samples, and no warmup per
selector: 24 reports and 72 samples.

Allocation regions bracket only the timed operation before owner destruction;
retained live bytes are not a leak measurement. Procfs windows include probe
activity and retain their empty controls without subtraction. CPU ticks may
quantize short operations to zero, and read/write counters are not physical
device attribution. The external time wrapper’s whole-child RSS includes
setup and corpus qualification and is not an operation-local allocation peak.

Run qualification with the same observer binary, one forward sample and no
warmup per selector, after artifact audit and before native capture: 12 reports
and 12 samples. Qualification must prove the selector is admitted, maps to the
audited source/output identity, and has a stable edit outcome before the native
lane starts.

The expected packet totals are therefore:

| Lane | Reports | Samples |
| --- | ---: | ---: |
| Native | 72 | 2,160 |
| Observer | 24 | 72 |
| Qualification | 12 | 12 |
| **Total** | **108** | **2,244** |

If any real selector has a typed refusal, preserve its qualification and audit
evidence and mark it non-admitted. Do not replace it with a generated case,
source-equality pseudo-success, or a different real fixture. The 108-report
target is reached only when all 12 selectors pass admission.

## Build and driver checks

The quality driver runs the six gates on the exact harness lock graph:

1. `cargo fmt --check`;
2. offline locked all-feature/all-target `cargo check`;
3. offline locked all-feature tests;
4. offline locked all-target Clippy with warnings denied;
5. offline locked all-feature rustdoc with warnings denied; and
6. the crate-boundary checker.

The build driver must preserve failed attempts and use an owned target. It
builds the three binaries with the profile in `plan.json`: opt-level 3, thin
LTO, one codegen unit, debug info 1, nonincremental, unwind panic, and two
Cargo jobs. It records the actual profile environment used for each child,
toolchain identity, command, feature list, binary hash, and log. In particular,
the native row must show `features: []`; the observer row must show exactly the
two diagnostic features above. A feature-enabled native executable is a hard
identity failure, not a comparable timing result.

Every driver rechecks production/harness source hashes, both lock identities,
the 35 architecture inputs, corpus/provenance, host/cgroup descriptors, packet
and driver hashes, and unrelated workspace files before and after each child.
The capture driver must require the successful independent artifact-audit
report and qualification identity before opening native handles. Readers run
only after capture handles terminate. Owned target and scratch directories are
cleaned after validation/review and before final sealing; the final seal records
the cleanup witness. Unrelated workspace files remain intact.

## Statistics and claim boundary

Report nearest-rank p50, p95, p99, and mean within each process, retaining all
block values and failures. Use the median of six process p50 values only as a
descriptive summary. Flag max/min block metric above 1.05 and median p99/p50
above 1.05. If bootstrapping is reported, use 10,000 resamples, seed `817817`,
sorted endpoints 250 and 9749, and 95% confidence.

Record available operation, output, write, sink, allocation, procfs, and RSS
vectors, and label unavailable metrics explicitly. The packet is scoped to
warm filesystem/provider caches and full durability. It makes no cold-cache,
network, physical-device, decompression-cause, scaling, optimization, or
historical before/after claim.

## Historical novelty and remaining gap

`documentProperties.docx` was used in earlier container/read and raw-source
work, but not this real-file OOXML `Package::open` → semantic edit →
`Package::save` route. `dateAutofilter.xlsx` has provenance and filter-related
coverage, but not real-file ordinary-save timing; its shared strings,
auto-filter, and MCE/extLst content add preservation coverage. `shapes.pptx`
has extensive prior low-level ZIP/OPC/layout coverage, but not this opened
presentation shape edit plus ordinary save. These are new route/case evidence,
not wholly new package fixtures.

The strongest remaining corpus gap is a producer-identified, larger,
multi-part, media-rich ordinary-save corpus: large or sparse/dense XLSX
workbooks, media-heavy decks, and a DOCX with an admitted semantic edit plus
unknown-part/relationship preservation. Current files are small, and DOCX/PPTX
producer identity is unavailable. Native Office round-trip and physical
cold/range-source evidence also remain outside this packet.

## Reviewed test-only recovery

The first complete all-feature test command retains 640 passed tests and one
ignored test across 26 successful suites, then fails only the final XLSX
planning-allocation integration test on an obsolete instrumentation label.
All frozen inputs for that attempt are archived before this amendment.
The next quality attempt proves production/runtime/other-test source equality,
reuses those successful suites, and executes the corrected integration under
all features and allocator-only plus the not-yet-run all-feature doctests.
Formatting, all-target checking, warning-denied Clippy/rustdoc and boundary
checks run freshly. This is explicit test-result reuse, not a claim that the
original full command passed. The quality target is reused as a Cargo cache;
commands and environment record its exact path.
