# 0711 DOCX `alt::scan` paired pilot packet

This packet records a rejected measured pilot for a private ownership change
inside `litchi_docx::alt::codec::scan`. The candidate borrowed slice-backed XML
events and the namespace resolver within each existing scan iteration. It was
not retained: the final checkout and `source-final.json` are byte-for-byte the
baseline, and the 0710 custom-properties preservation fix remains in place.
`performance_claim: none`.
The report is [`docs/performance/change-0711.md`](../../change-0711.md).

## Result

The frozen native gate required at least 3% improvement in both edit p50 and
edit mean for both corpora in both paired windows. Lifecycle p50 and mean had
to stay within 3% regression, and allocator request counts and requested bytes
had to stay within 3% regression. The packet has 32 isolated children: 16
native children at 100 samples/10 warmups and 16 allocator children at three
samples/no warmup. Output parity passes and 47 of 48 hard gates pass. The sole
failure is `pair-2/numbered-list/edit/p50/improvement`: 2.888465327% versus
the required 3%. The authoritative analysis is
[`analysis.json`](analysis.json); the disposition is recorded in
[`pilot-gate.json`](pilot-gate.json) and [`disposition.json`](disposition.json).

No candidate full quality suite, cfg(test) run, RSS capture, or Callgrind run
was executed at any point in this pilot; the native rejection stopped those
conditional follow-ups. The packet makes no hardware, RSS,
cold-cache, throughput, scaling, native Office, producer compatibility, or
broader DOCX performance claim.

## Capture contract

The two corpora are the generated medium harness DOCX and the admitted
`NumberedList.docx` fixture. The stages are `baseline-A1`, `candidate-B1`,
`candidate-B2`, and `baseline-A2`; B2 and A2 reverse corpus and phase order.
Each child is pinned to CPU 12 and records source, binary, fixture, argv, and
constraints custody. The native and allocator results use the same two phases:
`edit` and `lifecycle`. Their clocks are independent and are not additive.

The capture and verification tools are:

| Artifact | Purpose |
| --- | --- |
| [`plan.json`](plan.json) | Frozen corpora, stages, sample counts, thresholds, and limitations |
| [`capture.py`](capture.py) | One isolated child per corpus/phase/stage/lane with overwrite refusal |
| [`analyze.py`](analyze.py) | Custody, strict decoded-manifest validation, parity, statistics, and gates |
| [`analysis.json`](analysis.json) | Recomputed raw statistics, paired rows, flags, and rejection |
| [`pilot-gate.json`](pilot-gate.json) | Explicit native-gate disposition |
| [`disposition.json`](disposition.json) | Rejection, restored-source, and unexecuted-follow-up record |
| [`oracle/README.md`](oracle/README.md) | 49-case public scanner differential oracle |
| [`oracle/baseline/report.json`](oracle/baseline/report.json) and [`oracle/candidate/report.json`](oracle/candidate/report.json) | Byte-identical 49-case reports: 28 successes and 21 errors each |
| [`source-baseline.json`](source-baseline.json), [`source-candidate.json`](source-candidate.json), [`source-final.json`](source-final.json) | Exact source custody for baseline, candidate, and restored checkout |
| [`candidate.patch`](candidate.patch) | Retained candidate-only source patch |
| [`edit-attribution.json`](edit-attribution.json) | Historical profile selection evidence, explicitly diagnostic only |

The strict oracle includes the existing DOCTYPE event acceptance case. Its
acceptance is retained as an event-parity observation; it is not a generic XML
security claim and was not changed into a refusal.

## Reproduction and cleanup

From the repository root, the packet scripts are run with Python's bytecode
cache disabled. The coordinator builds the two source-bound binary stages and
invokes each capture stage serially. After all receipts are present, the
analyzer verifies the live binaries or exact post-cleanup binary witnesses and
writes a fresh output path. Owned scratch cleanup is complete. The packet's
final seal retains exact hashes for the reports, binaries, source manifests,
and six owned cleanup witnesses.

The allocator's `net_live` and `peak_above_start` values are per-operation
diagnostics. They are never summed across phases, repeats, or processes.
Likewise, the four lifecycle/edit distributions are separate clocks and do not
form a decomposition of a lifecycle sample.

The next experiment must preserve the scanner's ranges, unknown markup,
namespace and markup-compatibility choices, relationship checks, typed
refusals, limits, and allocation-failure behavior. The rejected pilot is
evidence for that bounded technical follow-up; the candidate remains
unshipped.

## Terminal verification

All six repository evidence gates pass. Four in-memory analyzer corruption
checks reject incorrect statistics, command arguments, decoded manifest bytes,
and deterministic output. The unmodified analyzer replay is byte-identical
after cleanup. Three owned scratch roots were removed; six exact binary
identities remain in `cleanup.json`. The terminal documentation check and
packet audit are reproducible with `final-gate.py` and `audit.py`; the artifact
manifest binds the final evidence files.
