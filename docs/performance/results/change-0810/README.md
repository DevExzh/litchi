# 0810 — private PPTX direct event handling workflow

This packet freezes a fresh performance trial for the reviewed
`direct-reader-namespace-resolver` candidate. The before leg starts at
`3677e31be5c9d5582a1f6d531ebb4d54db5a0acc`, the current repaired baseline, and
the production source census contains 9,196 tracked files. The only permitted
production change is `crates/litchi-pptx/src/notes/codec.rs`; the candidate is
applied only after the before build, probe quality gates, and all eighteen
before-only qualification reports pass independent raw-oracle review.

The candidate archive is root-owned and reviewed separately. It keeps the
buffered `NsReader` scanner and `inspect_element_oracle` unchanged as the
differential oracle. The candidate uses direct `Reader` events with the public
`NamespaceResolver`, pushes namespace scope before node and depth checks, and
pops after `Empty` and `End` before the next read. No new dependency or public
API change is in scope.

The independent probe is a byte-for-byte copy of the final 0806 probe source.
It retains the `namespace_uri_probe::capture_region_0793` owner,
`litchi-perf-0780-static-mce-capabilities` marker, repaired `valid-4attr`
fixture, and all 36 tests. Root `Cargo.lock` and `rustfmt.toml` are copied into
`inputs/` before the first build and their hashes are checked by every driver.
The packet binds these copies, the probe files, the driver sources, and the
architecture inputs before any measurement begins.

The public PPTX matrix has eighteen rows: six shapes (`tiny`, `medium`,
`large`, `vendor`, `unicode-vendor`, and `valid-4attr`) crossed with `capture`,
`commit`, and `lifecycle`. Native timing uses six alternating paired blocks in
the order `AB BA AB BA BA AB`, CPU 12, three warmups, and thirty samples. This
produces 216 native reports and 6,480 measured samples. Allocation timing uses
two paired blocks in the order `AB BA`, three samples, and no warmup: 72 reports
and 216 samples. Before-only qualification adds 18 reports and 18 samples.

Four owner-scoped Callgrind publications cover large capture only: two ordered
before/after pairs, one sample, and no warmup. They use the exact
`namespace_uri_probe::capture_region_0793` owner, collect only `Ir`, and must
produce one positive numbered publication per process. Callgrind is attribution
evidence; it makes no latency, RSS, phase-fraction, or speedup claim. There is
no cross-format, heaptrack, or perf lane. The frozen total is 310 reports and
6,718 samples.

Adoption requires at least one capture or lifecycle row to improve by at least
3 percent, with the 10,000-resample bootstrap high endpoint below 1.00. Any of
the eighteen rows with a ratio above 1.05 and bootstrap low endpoint above 1.00
is a veto. Each paired allocation block must also avoid increases in allocation
calls, allocated bytes, net live bytes, and peak-above-entry bytes. Allocation
reductions alone cannot satisfy the public workflow benefit gate. The frozen
bootstrap seed is 810810.

The root agent runs these drivers serially and keeps workload execution,
Cargo, rustfmt, and profiling under root custody:

```text
python3 -B docs/performance/results/change-0810/build.py before
python3 -B docs/performance/results/change-0810/probe_quality.py before
python3 -B docs/performance/results/change-0810/capture.py qualification
# The independent reader writes qualification-audit.json after raw review.
python3 -B docs/performance/results/change-0810/apply_candidate.py
python3 -B docs/performance/results/change-0810/quality.py after
python3 -B docs/performance/results/change-0810/build.py after
python3 -B docs/performance/results/change-0810/probe_quality.py after
python3 -B docs/performance/results/change-0810/capture.py native
python3 -B docs/performance/results/change-0810/capture.py allocation
python3 -B docs/performance/results/change-0810/profile.py
```

The after leg retains the before binaries and each build records exact source,
probe, root-input, command, and binary identities. Independent readers consume
immutable receipts only after all captures terminate. If the adoption policy
rejects the candidate, the root agent runs `restore_candidate.py` and verifies
the exact before source manifest. The target and six copied binaries are
removed only after source disposition, raw analysis, and cleanup audits pass.

The completed trial retains the one-file candidate. Four capture rows meet the
frozen benefit threshold (3.803–4.975%); there are no latency/resource vetoes,
and all 144 allocation comparisons are equal. All six production quality gates
and both 36-test probe lanes pass. All 310 reports and 6,718 measured samples
pass exact output and semantic checks. The independent raw audit agrees with
the main reader. Three isolated p99 increases remain visible diagnostics.

`decision.json` and `disposition.json` own the retention decision separately
from immutable numerical results. `chronology.json` records artifact timestamp
observations for replay after checkout; workload times remain driver receipts.
`cleanup.json` witnesses removal of the owned target and its six verified
binaries. `validate.py --final` joins all readers and custody checks; `seal.py`
checks the exact staged or committed batch.

See the [full report](../../0810-pptx-direct-event-workflow.md),
[source review](source-review.md), [profile review](profile-review.md), and
[independent result review](results-review.md). The broad project goal remains
active. No universal speedup, tail-latency, RSS, or cold-cache claim is made.
