# 0808 — private PPTX notes/codec workflow qualification

**Terminal status: deferred before performance trials.** Only the eighteen
baseline qualification reports were captured. Candidate formatting, checking,
and 1,241 tests pass; warning-denied Clippy fails at three pre-existing test
expressions, reproduced by a restored-baseline control. Production is restored
exactly, and the owned target and three baseline binaries have been removed.
The following frozen full-trial protocol records the original plan; the
early-stop section below records what actually ran.

This packet freezes a fresh experiment for one private PPTX candidate. The
candidate archive is supplied by the review agent and is limited to
`crates/litchi-pptx/src/notes/codec.rs`. The before leg starts at
`d28e3dc702d84f8752a2e329c2d3fc15f5c94f06`; its full tracked production source
manifest contains 9,196 files. No other production path may change. The
candidate is applied only after the before build, probe quality gates, and
all eighteen before qualification reports pass the independent raw-oracle
audit. The candidate then passes all six production quality gates before
paired trials.

The independent probe is an exact byte copy of the final 0806 probe source.
It retains the `namespace_uri_probe::capture_region_0793` owner and the
`litchi-perf-0780-static-mce-capabilities` marker. Its `valid-4attr` shape and
36 tests are therefore part of the frozen probe contract. The copied files and
hashes are recorded in `inheritance.json`; the build driver binds every copied
file before compiling.

The public PPTX matrix has eighteen rows: six shapes (`tiny`, `medium`,
`large`, `vendor`, `unicode-vendor`, and `valid-4attr`) crossed with `capture`,
`commit`, and `lifecycle`. Native timing uses six alternating paired blocks in
the order `AB BA AB BA BA AB`, CPU 12, three warmups, and thirty samples. This
produces 216 native reports and 6,480 measured samples. Allocation timing uses
two paired blocks, three samples, and no warmup: 72 reports and 216 samples.
The before-only qualification adds 18 reports and 18 samples before the
candidate is applied.

Four additional owner-scoped Callgrind publications cover large capture only:
two ordered before/after pairs, one sample, and no warmup. They use the exact
`namespace_uri_probe::capture_region_0793` owner and one positive numbered
publication per process. Callgrind `Ir` is retained as an attribution
diagnostic; it makes no latency, RSS, phase-fraction, or speedup claim. There
is no cross-format lane, heaptrack lane, or perf lane in this private PPTX
experiment. The frozen total is 310 reports and 6,718 samples.

Adoption requires at least one capture or lifecycle row to improve by at least
3%, with the 10,000-resample bootstrap high endpoint below 1.00. Any of the
eighteen rows with a ratio above 1.05 and bootstrap low endpoint above 1.00 is
a veto. Each paired allocation block must also avoid increases in allocation
calls, allocated bytes, net live bytes, and peak-above-entry bytes. The frozen
bootstrap seed is 808080; allocation-only reductions cannot satisfy the
workflow benefit gate.

The root agent runs the drivers serially after reviewing this packet:

```text
python3 -B docs/performance/results/change-0808/build.py before
python3 -B docs/performance/results/change-0808/probe_quality.py before
python3 -B docs/performance/results/change-0808/capture.py qualification
# apply_candidate.py additionally requires qualification-audit.json from the
# independent raw-oracle reader, with accepted_before_application=true.
python3 -B docs/performance/results/change-0808/apply_candidate.py
python3 -B docs/performance/results/change-0808/quality.py after
python3 -B docs/performance/results/change-0808/build.py after
python3 -B docs/performance/results/change-0808/probe_quality.py after
python3 -B docs/performance/results/change-0808/capture.py native
python3 -B docs/performance/results/change-0808/capture.py allocation
python3 -B docs/performance/results/change-0808/profile.py
```

`build.py` produces three independently identified binaries per leg: native,
allocation, and profile. `quality.py` retains six production gates for the
PPTX package set, including all-features/all-targets Clippy and the crate
boundary check. `probe_quality.py` retains format, 36-test, and warning-denied
Clippy gates for both legs. The application and restoration scripts are
root-only and fail closed on any source path outside the one-file allowlist.

Independent analyzers and validators consume the immutable receipts after
capture. Their ownership is separate from the workload drivers. Temporary targets and
copied binaries are removed only after those readers verify their identities.

## Early stop before the after release measurement build

This trial terminated before the after release measurement build and before every native,
allocation, and Callgrind lane. The after production quality run executed
exactly four gates: format, check, and tests passed; warning-denied all-targets
Clippy failed at the three pre-existing `err-expect` sites in
`crates/litchi-pptx/src/opened/tests.rs` (464, 538, and 557). The retained
before qualification has 18 reports and 18 samples, and the before probe
quality receipt records three passing gates with 36 tests.

The restored-baseline control reproduced the same three `err-expect` diagnostics
at those locations, so this packet does not claim a six-gate production-quality
pass.

After root records the rejected disposition and restores the exact
`build-before/source.json` source manifest, run the separate early-stop cleanup
driver:

```text
python3 -B docs/performance/results/change-0808/cleanup_early_stop.py
```

It verifies the restoration, qualification and failure receipts, checks the
three before binary identities, and removes only the owned
`/home/zhuhe/code/litchi-target-0808` tree. It writes
`early-stop-cleanup.json`; the candidate archive, application witness,
qualification reports, probe receipts, and failed quality logs remain. This
early-stop record contains no comparative timing or adoption claim and is
separate from the full planned protocol.

Replay the actual terminal record after cleanup:

```text
python3 -B docs/performance/results/change-0808/analysis.py --qualification --check
python3 -B docs/performance/results/change-0808/validate_early_stop.py --check --require-cleanup
python3 -B docs/performance/results/change-0808/seal_packet.py --check-head
```

The five unexecuted full-trial reader drafts were removed after the stop.
The retained terminal reader validates all frozen inputs, accepted qualification,
partial quality, the baseline control, exact restoration, and cleanup. The seal
covers the packet and six report/index documents; production has no retained diff.
