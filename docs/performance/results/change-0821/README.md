# 0821 — real-file save durability attribution

This packet resumes the unexecuted 0820 real-file durability matrix on the
repaired base `8312aaa29b59f73d2a7409cb501828989320a8c9`. It covers the same
three checked-in DOCX, XLSX, and PPTX inputs, two save phases, and four
durability policies. The plan contains 24 selectors and admits 216 reports
and 4,488 samples after the independent artifact and qualification gates.

The production and runtime harness source bytes are unchanged. The packet
reuses the committed 0820 repair quality result through
`quality.py`: it verifies all six passing gate result descriptors and logs,
the focused result, the 641-pass test summary, the original frozen-input
witness, and every payload in the 0820 seal. It then emits a fresh 0821
quality receipt bound to the current revision. This adapter executes no Cargo
command and does not claim a second test run.

Root owns release builds, artifact export, admission, qualification, native
and observer captures, offline replay, cleanup, and sealing. All outputs must
remain tied to the packet source, corpus hashes, policy route, and explicit
filesystem root. Native and observer latencies are kept separate; policy
ratios are configuration attribution only. No historical timing comparison or
production optimization claim is admitted.

Before measurement, run:

```sh
python3 -B docs/performance/results/change-0821/quality.py
python3 -B docs/performance/results/change-0821/build.py
```

Then run the root-owned artifact, admission, capture, replay, cleanup, and
seal drivers in their documented order. The release target is
`/home/zhuhe/code/litchi-target-0821` and the owned filesystem root is
`/home/zhuhe/code/litchi-fs-0821`; both are refused if stale or foreign.

The prior 0820 packet remains immutable. iWork is outside this packet.

## Retained result and replay

The completed packet retains 216 reports / 4,488 samples: 24 qualification,
144 native, and 48 observer reports. All six full/default paired intervals
include 1.0. Weaker durability configurations are faster on this host but do
not authorize weakening the default. See [the report](../../0821-real-file-save-durability.md).

Artifact execution order is export, `preservation.py`, then `admission.py
artifacts`. Root initially omitted the middle step; `execution-errors.json`
and `admission-0` retain that parent failure and successful child audit. The
second admission attempt passed on the same exports after preservation.
No timing started before artifact and qualification admission passed.

Offline replay after capture completion:

```sh
python3 -B docs/performance/results/change-0821/analyze.py --check
python3 -B docs/performance/results/change-0821/root_audit.py --check
python3 -B docs/performance/results/change-0821/validate.py --final
python3 -B docs/performance/results/change-0821/seal.py --check-head
```

Reader and validator initial schema failures are retained alongside corrected
replay evidence. `cleanup.json` records owned target/scratch removal; `seal.json`
binds the final packet and six performance documents.
