# 0959 semantic owner ranking packet

This is a bounded 0959 evidence packet for the sealed 0958 heaptrack traces,
the integrated Number layout probe, and the Arc candidate's pre-measurement
quality/source custody. The ranking helper and layout probe launched no
workload. The corrected v8 native and allocator lanes have now completed and
the bounded comparison summary is included below. The Arc candidate is
rejected by the frozen gates; production remains unchanged.

## Capture protocol and epoch status

The frozen prospective protocol is
[`protocol/protocol-0959-v4.md`](protocol/protocol-0959-v4.md), with the
marker-free selected-cell control documented separately in
[`review/query-control-protocol-v1.md`](review/query-control-protocol-v1.md).
The operative plan is [`plan-final-v8.json`](plan-final-v8.json), SHA-256
`9964f2f3a86dd467de6d7840807400aa7468a48f7d993b50ea296c6e0909abf8`, and its
capture epoch is `v2-corrected-output-ownership`. The v7 capture helper and
v7 static controls are retained at
[`tools/capture_0959_v7.py`](tools/capture_0959_v7.py) and
[`tools/controls_0959_v7.py`](tools/controls_0959_v7.py). Their SHA-256 values
are `0c3e298c1e9db9d6e40a313d82f9c75a27ffbd3a2d42d7bd612b73078ec5c10e` and
`c969edaaf2d2202382411f27b531ac6187bdf2f188e5a6a8ef78430de8df9328`.

The static controls passed in
[`receipts/controls-0959-v7.json`](receipts/controls-0959-v7.json), and the
second root authorization is
[`receipts/root-go-v2.json`](receipts/root-go-v2.json). The v1 native epoch is
preserved under `quality-failures/native-capture-v1/`: its first child failed
before fixture preparation because the old helper precreated the output
directory, producing zero samples. The correction is recorded in
[`review/capture-output-directory-correction-v1.json`](review/capture-output-directory-correction-v1.json);
the v1 epoch was not retried or replaced. At the initial packet refresh, native
v2 was running and allocator capture remained pending; both lanes subsequently
terminated successfully. Raw capture trees and the full comparison remain
external; bounded terminal manifests and the case-level summary are retained
in this packet.

## Final comparison and decision

The full reduced comparison remains external because it is 1,447,293 bytes:
`/home/zhuhe/litchi-performance-0959/analysis/comparison-v3.json`, SHA-256
`9dbaefcd25b348c6b3f3c23ff4ed94b6e160fa7c95a5e37667a7efafa80242eb`.
The bounded projection is
[`analysis/comparison-summary-v3.json`](analysis/comparison-summary-v3.json),
497,234 bytes, SHA-256
`43a1c6eb5de972e46c1a07ca50161d7bde43ec38c7be01102a15480229f593db`.
It retains all case-level native and allocator metrics, aggregate confidence
intervals, paired/block flags, gate outcomes, and the external replay command;
verbose child rows remain external.

The sealed packets contain 72 native children and 18 allocator children. The
native medium-one-edit total ratio is 1.088054 at p50 and 1.101930 at p95,
with safety confidence bounds above 1.05. Allocated-byte and region-peak
ratios do not pass the allocator improvement gate, despite lower allocation
call ratios. The root decision is
[`review/root-decision-v1.json`](review/root-decision-v1.json), SHA-256
`1a2ae65919ec36f6a4c6294cdbdca8716442be2b56bfd891356d8f854d6c9534`:
reject the Arc candidate and keep production unchanged. The selected-cell
query control was not run and cannot change this decision.

The sealed-capture audit and independent reduction audit both passed. They
recomputed 63 aggregate metrics and found no mismatch:
[`review/sealed-capture-audit-v1.json`](review/sealed-capture-audit-v1.json),
SHA-256 `6b10171134d77a11a77ca001b0bc67feb120f76d6f0d89e99452cc78a094816f`,
and
[`review/final-reduction-independent-review-v1.json`](review/final-reduction-independent-review-v1.json),
SHA-256 `cbe22d6bf6fdfc786840d345db53937c0b9524d1a6a682483add22526ca4fdb0`.

The first reducer attempt failed before reduction and is preserved in
[`review/root-reducer-v2-terminal-v1.json`](review/root-reducer-v2-terminal-v1.json);
the corrected final-plan adapter and reducer correction receipt are retained
alongside it. Cleanup custody is recorded in
[`receipts/cleanup-v2.json`](receipts/cleanup-v2.json).

The in-repository result is
[`analysis/owner-ranking-summary-v3.json`](analysis/owner-ranking-summary-v3.json),
SHA-256 `f50bb1f2ce39f50d5488d9c8a6a95855fcc98dfc526bbd447e58c871b0f359b4`.
It is a 176,734-byte bounded projection of the full 623,430-byte report,
which remains external and is bound by
[`review/owner-ranking-artifact-descriptor-v1.json`](review/owner-ranking-artifact-descriptor-v1.json).
The summary reports 13,032,599 whole-process calls and 3,406,101,409
requested bytes. The exact fully qualified stage/timed selection contains
2,081,214 calls and 274,320,408 requested bytes; 6 unknown rows and the
unqualified remainder stay explicit.

The ranking uses the first meaningful `litchi_xlsx` frame above allocator and
runtime frames as a nearest semantic owner. It is not an inclusive call-tree
sum. Exact fully qualified `stage_timed_commit` and `timed_iteration`
callsites select rows; `MultiSourceEdit::commit` is reported as a semantic
owner/tag but is never used as a selection surrogate. Growth, merge, numeric,
duplicate-checking, XML, copy/clone, and realloc labels overlap and are
descriptive rather than additive. The result provides prioritization evidence
only: it makes no timing, phase-only allocation, speedup, or adoption claim.

The helper and final receipt are:

* [`tools/rank_0959_v1.py`](tools/rank_0959_v1.py), SHA-256
  `3daeaac5269257d60afc1089434c7eaf3d130d85ff6e62cb8bf2194adc0be2be`.
* [`review/semantic-owner-ranking-receipt-v3.json`](review/semantic-owner-ranking-receipt-v3.json),
  SHA-256 `c942c41a1b119240142259bbfdf5fc7d0a0764d62268e70f9c87e172b4c4eb79`.
* [`root-start-v1.json`](root-start-v1.json), SHA-256
  `2f48c05e0aa0027ea8be494f29bc1917ab5d192975bb0c3cab43d60d79856cc4`.

The 0958 correction records why the independent ranking was needed:

* [`review/0958-surrogate-selection-review-v2.json`](review/0958-surrogate-selection-review-v2.json),
  SHA-256 `084427c576d884be708618188a45cc11f75d490de363d6f79bc9100e9b372792`,
  corrects the v1 trace-count arithmetic and finds that the 0958 resolver's
  frame-first scan can label exact selected stage/timed stacks as the
  commit-core surrogate. The selected totals and raw captures remain
  unchanged; the resolution labels require correction. The v1 receipt is
  retained as superseded history.
* [`review/callsite-selection-correction-v1.json`](review/callsite-selection-correction-v1.json),
  SHA-256 `d942ba8967430fa2f22fc98b60f6192497219eebf9245e82251d8cb4836f4812`,
  records the v2 token-substring error and the v3 fully qualified benchmark
  callsite requirement.
* [`review/selection-scope-correction-v1.json`](review/selection-scope-correction-v1.json),
  SHA-256 `48d5a1150d5a4ca903d1d8251370c676be79c5b8b7d0886524ec0d8e80e5dc07`,
  preserves the earlier whole-process size-bucket selection defect.

The ranking was replayed from fresh external output with:

```sh
python3 -B /home/zhuhe/litchi-performance-0959/tools/rank_0959_v1.py \
  --manifest /home/zhuhe/litchi-performance-0958/analysis-input-manifest-v1.json \
  --reference-analysis /home/zhuhe/litchi-performance-0958/analysis-v1/attribute-v3.json \
  --output /home/zhuhe/litchi-performance-0959/analysis/owner-ranking-v3.json
```

The 0958 input manifest, decoder traces, and full 0959 report remain external;
the descriptor records their paths, sizes, hashes, and fresh-output policy.
The packet contains no source tree, raw capture, full report, executable, or
other file at or above 1 MiB.

## Integrated layout probe

The corrected integrated probe passed with no workload and no production
candidate. [`receipts/layout-probe-result-v2.json`](receipts/layout-probe-result-v2.json)
has SHA-256 `601dd859bec08538cf74d0364cb73e5e69416017f6d0ea65aec9a1cc37a51411`.
It reports `Number` 16 bytes/alignment 8, `Value` 24/8, `Cell` 72/8,
`Stored` 184/8, `Formula` 72/8, and `Content` 72/8, with corresponding
`Option` sizes unchanged. The rough alternative probe rows are not integrated
production layouts. The v1 `-D unused-qualifications` failure remains under
`quality-failures/` and is preserved by the layout summary.

The probe patch and runner are
[`review/cell-layout-probe-v2.patch`](review/cell-layout-probe-v2.patch) and
[`tooling/run-layout-probe-0959-v2.py`](tooling/run-layout-probe-0959-v2.py).
The layout summary also binds its terminal, stdout/stderr, lock cleanup, and
source descriptors.

## Arc candidate custody and quality

The frozen candidate patch is
[`receipts/candidate-arc-v1.patch`](receipts/candidate-arc-v1.patch), 6,581
bytes, SHA-256 `8df3336d9964f47a5656cc3a093615bdb113315448c43a00f66101f9f4f85b68`.
The candidate freeze and baseline-comparison freeze both pass:
[`receipts/root-candidate-freeze-v1.json`](receipts/root-candidate-freeze-v1.json)
and [`receipts/baseline-comparison-freeze-v1.json`](receipts/baseline-comparison-freeze-v1.json).
The three production delta paths are `cell.rs`, `raw/worksheet/selected.rs`,
and `raw/worksheet/semantic.rs`; source trees, manifests, and binaries remain
external. The source custody descriptor is
[`receipts/source-baseline-v2.json`](receipts/source-baseline-v2.json), SHA-256
`a9bfd96b76a8754b75d0828f5cb701ab031312f131903a1a06bef030360bfd31`.

Focused candidate checks passed in the v2 and v3 receipts. The full XLSX
tests, Clippy, and scoped rustfmt receipts all contain return code 0; their
argv, environment, source start/end hashes, and bounded logs are retained
under `receipts/` and `logs/`. The final v2 native and allocator build
adapters passed and bind external binaries by SHA-256. The earlier v1 build
failure is preserved under `quality-failures/`; its stderr records the
`-C embed-bitcode=no` and `-C lto` conflict. It is historical and does not
represent the v2 build result.

The Number design and semantic review are
[`review/number-storage-design-v1.md`](review/number-storage-design-v1.md) and
[`review/number-representation-review-v2.md`](review/number-representation-review-v2.md).
They require exact lexical preservation, parser/fallback/error ordering,
representation-independent equality/hash, source-backed clone checks, and
end-to-end output/inverse validation. Focused quality does not establish an
allocation saving or authorize adoption. The earlier bound plan remains
[`plan-v5-bound.json`](plan-v5-bound.json) with status `prospective`; it is a
historical pre-capture snapshot. The authorized v8 plan above supersedes it for
the completed native and allocator epochs. The adoption gates reject the Arc
candidate as recorded above.

The current bounded packet census is recorded in
[`review/packet-provenance-v4.json`](review/packet-provenance-v4.json). It
describes the final bounded packet while excluding the provenance snapshots
themselves from the census. The v1-v3 provenance files remain as historical
snapshots; v4 records the completed independent audits.
