# 0815 — private PPTX borrowed event-arm binding workflow

This packet freezes a fresh performance trial for the reviewed
`borrowed-event-arm-bindings` candidate. The before leg starts at
`55bb2ead3498043dd22b53555507b24402486248`, the current production baseline,
and the production source census contains 9,196 tracked files. The only
permitted production change is `crates/litchi-pptx/src/notes/codec.rs`.

The candidate archive is root-owned and reviewed separately. It keeps the
buffered scanner, refusal corpus, and `inspect_element_oracle` unchanged. It
changes only the `Start` and `Empty` event bindings in `scan_processed_xml`:
the two payloads are bound by `ref`, and the four corresponding borrowed
arguments lose an unnecessary `&`. The event bodies, ordering, parser/error
handling, resolver behavior, tests, dependencies, and public API remain out
of scope. No result or adoption claim exists at packet freeze.

The independent probe is a byte-for-byte copy of the final 0813 probe source.
It retains the `namespace_uri_probe::capture_region_0793` owner,
`litchi-perf-0780-static-mce-capabilities` marker, repaired `valid-4attr`
fixture, and all 36 tests. Root `Cargo.lock` and `rustfmt.toml` are copied into
`inputs/` before the first build and their hashes are checked by every driver.
The packet also binds the six probe files, 35 architecture inputs, three
unrelated workspace files, toolchain observation, driver sources, and
candidate archive before the first build.

The baseline production quality gates are reused from sealed 0813 after
because all 9,196 tracked production file hashes match that source. This
reuse executes no Cargo command and retains the sealed six-gate result: 1,241
passed tests, zero failed, three ignored, and 85 suites. Both 0815 probe legs
still run fresh formatting, release all-feature tests, and release Clippy with
warnings denied. The candidate is applied only after the before build, fresh
probe gates, and all eighteen before-only qualification reports pass the
independent raw-oracle review.

The public PPTX matrix has eighteen rows: six shapes (`tiny`, `medium`,
`large`, `vendor`, `unicode-vendor`, and `valid-4attr`) crossed with `capture`,
`commit`, and `lifecycle`. Native timing uses six alternating paired blocks in
the order `AB BA AB BA BA AB`, CPU 12, three warmups, and thirty samples. This
produces 216 native reports and 6,480 measured samples. Allocation timing uses
two paired blocks in the order `AB BA`, three samples, and no warmup: 72 reports
and 216 samples. Before-only qualification adds 18 reports and 18 samples.

Four owner-scoped Callgrind publications cover large capture only: two ordered
before/after pairs, one sample, and no warmup. They use the exact
`namespace_uri_probe::capture_region_0793` owner, collect only `Ir`, and make
no latency, RSS, phase-fraction, or speedup claim. The frozen total is 310
reports and 6,718 samples.

The adoption policy requires at least one capture or lifecycle row to improve
by at least 3 percent, with the 10,000-resample bootstrap high endpoint below
1.00. Any of the eighteen rows with a ratio above 1.05 and bootstrap low
endpoint above 1.00 is a veto. Each paired allocation block must also avoid
increases in allocation calls, allocated bytes, net live bytes, and
peak-above-entry bytes. The frozen bootstrap seed is 815815. Allocation
reductions alone cannot satisfy the public workflow benefit gate.

The root agent runs these drivers serially and keeps workload execution,
Cargo, rustfmt, profiling, and binary tools under root custody:

```text
python3 -B docs/performance/results/change-0815/quality_reuse.py
python3 -B docs/performance/results/change-0815/build.py before
python3 -B docs/performance/results/change-0815/codegen.py before
python3 -B docs/performance/results/change-0815/probe_quality.py before
python3 -B docs/performance/results/change-0815/capture.py qualification
# The independent reader writes qualification-audit.json after raw review.
python3 -B docs/performance/results/change-0815/apply_candidate.py
python3 -B docs/performance/results/change-0815/quality.py after
python3 -B docs/performance/results/change-0815/build.py after
python3 -B docs/performance/results/change-0815/codegen.py after
python3 -B docs/performance/results/change-0815/codegen_analysis.py --write
python3 -B docs/performance/results/change-0815/probe_quality.py after
# Root reviews codegen-before/codegen-after and writes codegen-gate.json.
python3 -B docs/performance/results/change-0815/capture.py native
python3 -B docs/performance/results/change-0815/capture.py allocation
python3 -B docs/performance/results/change-0815/profile.py
```

The after leg retains the before binaries and each build records exact source,
probe, root-input, command, and binary identities. Independent readers consume
immutable receipts only after all captures terminate. If the adoption policy
rejects the candidate, the root agent runs `restore_candidate.py` and verifies
the exact before source manifest. The target and six copied binaries are
removed only after source disposition, raw analysis, and cleanup audits pass.

`decision.json` and `disposition.json` own the retention decision separately
from immutable numerical results. `chronology.json` records artifact timestamp
observations for replay after checkout; workload times remain driver receipts.
This packet makes no universal speedup, tail-latency, RSS, cold-cache, or
cross-format claim. The broader project goal remains active and iWork remains
out of scope.

## Final result

Rejected: no eligible capture/lifecycle benefit, with a 5.299% large-lifecycle
p50 regression triggering the frozen veto. All 144 paired allocation metrics
are equal. The static arm copies disappear, but native workflow performance
does not improve. The baseline source is restored exactly; see `decision.json`,
`disposition.json`, the independent reviews, and the top-level 0815 report.
