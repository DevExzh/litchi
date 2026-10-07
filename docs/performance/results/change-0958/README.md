# 0958 bounded heaptrack attribution packet

This packet records the completed 0958 XLSX heaptrack capture and its
offline-only allocation attribution. Six children were admitted and captured;
the first-child symbol feasibility check ran after that child's terminal and
before the remaining children, while bulk decoding and attribution began only
after capture freeze. The frozen capture manifest records 57 files,
92,466,528 bytes, return code 0, and the six plan-v5 child IDs. Raw heaptrack
profiles, decoder exports, and source/build trees remain external because
repository evidence files must stay below 1 MiB.

The authoritative result is
[`analysis-v1/attribute-v3.json`](analysis-v1/attribute-v3.json), SHA-256
`9fb5ae3b1b484e2fc8294c66d4712d8f818c1aed60a974fbe51bcd03173cf983`.
It reports 13,032,599 whole-process allocation calls and 3,406,101,409
requested bytes across six traces. The admitted owner subset contains
2,081,214 calls and 274,320,408 requested bytes. Every run conserves raw
decoder totals, histogram totals, selected rows, unqualified rows, and the
exclusive leaf partition. The bounded compact projection is
[`analysis-v1/summary-compact-v1.json`](analysis-v1/summary-compact-v1.json),
SHA-256 `60651f3f1495bd167b4a38d042b6d5201527ceae10403a0c6a7610e14ce61197`.

The owner result is a whole-process attribution subset. A row is selected only
when `timed_iteration` and the stage owner are resolved and no unknown-frame
marker is present. `stage_timed_commit` is the exact stage owner when present;
`MultiSourceEdit::commit` is retained as a separately labelled commit-core
surrogate when the stage is inlined or unresolved. These mixed resolution
buckets are disjoint within the selected rows, but they are not a complete
phase allocation measurement. Unknown and unqualified rows remain visible.
Source tags such as number, stored, clone, merge, copy, and realloc overlap;
only `exclusive_leaf_partition` is additive. The packet makes no timing,
physical I/O, production-adoption, or phase-only allocation claim.

The source-bound numeric basis is
[`review/number-ownership-trace-v2.md`](review/number-ownership-trace-v2.md),
SHA-256 `3de25114247d354cf6b512db9599aac5ce921149a06fb13351dd048382cd75d2`.
The independent OPC audit is
[`review/opc-attribute-singleton-audit.md`](review/opc-attribute-singleton-audit.md).
The independent six-run review is
[`review/independent-analysis-review-v1.md`](review/independent-analysis-review-v1.md),
SHA-256 `8f4ae832e1bf211aee7c29db14ca0915735d869e45252504ae6fdba3741f9ec1`.

The final packet bindings are:

* [`plan-v5.json`](plan-v5.json), SHA-256
  `5bd826e1f93913b5be2fc5878be0c5270e323789222fffcf0e97121dcbc37695`.
* [`tooling/capture_0958_v5.py`](tooling/capture_0958_v5.py), SHA-256
  `15c09dffefed8bf6084ec0c54ddf7d19f2f69a3c70f4e2bf41a7442a40000f6c`.
* [`capture-manifest-v1.json`](capture-manifest-v1.json), SHA-256
  `bd30035f2c7926ec579fea58ecd1e5c5fe34da6b2ae129473bee033630e5c465`.
* [`analysis-input-manifest-v1.json`](analysis-input-manifest-v1.json), SHA-256
  `0dc7a9a4757d73a05d9fc289581a8ec6add494565f039b94e7addf8d18629495`.
* [`tools/attribute_0958_v1.py`](tools/attribute_0958_v1.py), SHA-256
  `1bb7584076762eefadc409016abf50684157e9350a76386255c5763e47d7cffd`.
* [`tooling/attribute-contract-v2.json`](tooling/attribute-contract-v2.json),
  SHA-256 `9818ca39217c7fd363f5abe04e5a2feecde1450bfc8e4596c0e9fc362ef546e5`.
* [`review/attribute-tooling-receipt-v2.json`](review/attribute-tooling-receipt-v2.json)
  and [`review/attribute-controls-receipt-v2.json`](review/attribute-controls-receipt-v2.json),
  which bind the final analyzer and its positive/fail-closed controls.
The independent replay receipt is [`review/root-replay-receipt-v1.json`](review/root-replay-receipt-v1.json), SHA-256 `4d83458f8185309788c3d46660f404f89a5562b95093765700d572eac83f9099`; it records exact field equality to the authoritative output except the generated timestamp.
The final review receipt is [`review/attribute-final-review-receipt-v1.json`](review/attribute-final-review-receipt-v1.json), SHA-256 `a2a9f122ae17cb33c29703e588efd149d6b71f0a5eb533c27d9bfe2b22e5d63e`.
The compact external-artifact binding is [`review/attribute-artifact-manifest-v1.json`](review/attribute-artifact-manifest-v1.json), SHA-256 `e132f847eea3fa1d48a67b8dcbba7b3710bbf17e7d2a1b2dde12f346453588de`.
The cleanup receipt is [`review/cleanup-v1.json`](review/cleanup-v1.json), SHA-256 `1ee437f0afcc523c5952c84bfe23dd2c5e3abaeab56f22c37ab97e4ec18dc3d4`; it records only disposable 0958 scratch/cache removal.
Root’s bounded final validation is [`review/root-final-validation-v1.json`](review/root-final-validation-v1.json), SHA-256 `416a587c0b0eedded55453c2fc2fecace3b091229582731b8f86d3d5a4487124`.
The packet-wide descriptor is [`review/final-packet-provenance-v1.json`](review/final-packet-provenance-v1.json); it records the bounded file census, final bindings, external artifact roots, replay commands, and historical analysis failures.

`plan-v5.json` preserves its pre-capture `prospective_no_capture` status as a
historical plan field. The capture manifest and capture-freeze receipt are the
authoritative post-capture custody records.

The capture and analysis can be regenerated only into fresh external output
directories. The executed command forms were:

```sh
python3 -B /home/zhuhe/litchi-performance-0958/tooling/capture_0958_v5.py \
  --capture --go
python3 -B /home/zhuhe/litchi-performance-0949/tools/decode_heaptrack_0949_v1.py \
  --profile h00-medium-percent-r1=/home/zhuhe/litchi-performance-0958/capture-v1/children/h00-medium-percent-r1/heaptrack.data.gz.zst \
  --profile h01-dense-percent-r1=/home/zhuhe/litchi-performance-0958/capture-v1/children/h01-dense-percent-r1/heaptrack.data.gz.zst \
  --profile h02-medium-one-edit-r1=/home/zhuhe/litchi-performance-0958/capture-v1/children/h02-medium-one-edit-r1/heaptrack.data.gz.zst \
  --profile h03-medium-one-edit-r2=/home/zhuhe/litchi-performance-0958/capture-v1/children/h03-medium-one-edit-r2/heaptrack.data.gz.zst \
  --profile h04-dense-percent-r2=/home/zhuhe/litchi-performance-0958/capture-v1/children/h04-dense-percent-r2/heaptrack.data.gz.zst \
  --profile h05-medium-percent-r2=/home/zhuhe/litchi-performance-0958/capture-v1/children/h05-medium-percent-r2/heaptrack.data.gz.zst \
  --output-root /home/zhuhe/litchi-performance-0958/raw-decoder-v1
python3 -B /home/zhuhe/litchi-performance-0958/tools/export_heaptrack_0958_v1.py
python3 -B /home/zhuhe/litchi-performance-0958/tools/attribute_0958_v1.py \
  --manifest /home/zhuhe/litchi-performance-0958/analysis-input-manifest-v1.json \
  --output /home/zhuhe/litchi-performance-0958/analysis-v1/attribute-v3.json
python3 -B /home/zhuhe/litchi-performance-0958/tools/attribute_0958_v1.py --self-test
python3 -B /home/zhuhe/litchi-performance-0958/review/mock-controls-v5.py \
  > /home/zhuhe/litchi-performance-0958/review/mock-controls-v5.stdout.json
```

The capture, decoder, exporter, and analyzer use sealed input descriptors and
refuse existing output paths. The commands above name the historical sealed
roots; a replay must substitute a fresh external root for every output and
must not rerun a workload against an existing capture root. A relocated replay therefore needs a versioned adapter
that rewrites only approved historical path fields and uses a fresh output
root; it must not fall back to the original packet. The external input manifest
binds each profile, report, decoder summary, size aggregate, trace aggregate,
histogram, and folded export by path, size, mode, and SHA-256. The six raw
capture profiles and derived exports are external under
`/home/zhuhe/litchi-performance-0958`; their names, sizes, hashes, and
regeneration descriptors are retained in the manifests above.

The packet retains the earlier v1 control/receipt files as historical context.
The first attribution output failed because its folded parser accepted only
tab-separated rows; v2 parsed the data but used an overbroad generic clone
classification. Both remain external in the analysis history and are excluded
from the authoritative interpretation. The root-start and source/build
provenance files bind the retained 0957 executable and source closure; no
benchmark binary or large source manifest is copied here.

This packet is evidence custody and attribution only. The main 0958 report and
any production decision are maintained separately by the root agent.
