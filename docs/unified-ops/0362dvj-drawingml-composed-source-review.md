# 0362dvj — composed DrawingML source and consumer review

The coordinator integrates the shared DrawingML owner contract and all five
raw-codec lanes together. Every lane has completed independent source review;
this record accepts the composed source and consumer migration, while native
performance qualification remains required.

| Lane | Reviewed leaf |
| --- | --- |
| Theme and Ink | `28402d25172b9bde944c6005bbfbdf6b8fcdf46b` |
| Chart | `f9173086353407a21851ea004fc8bb25ee60205d` |
| Diagram | `0f914eb8a08c8119302cf725a2f77687ab888e21` |
| Geometry, color, transform and formatcode2 | `e5d96d490ac57fb25c3921c5b04009d33ce62fac` |
| Text and media | `fa4f065ce3c22e90127e1e7754ba6623771b391c` |

The four native owners use the canonical edit/commit/package-replay shape and
shared Commit accessors. Raw codecs require explicit policy arguments and
preserve typed limit errors. The migration includes format hosts, tests,
examples, glossary registrations and live standalone consumers. Raw-only
codecs and scanners do not acquire an invented owner lifecycle.

## Composition repairs

The sole textual conflict is PPTX shape text-body ingress. The resolution
keeps the fragment lane's `text_body_limits_for` helper, which selects the
scene's input budget, and the text lane's new `text_reader::read` name. It
does not substitute the output budget for input admission.

The fresh composed check identifies ten newer DOCX projected-Ink consumers
that postdate the leaf base: four production calls and six tests. All now
pass the required `shared::Limits::DEFAULT` argument. This preserves the
historical fixed payload-scanner policy. Independent review confirms that
host per-payload and aggregate byte limits are checked before these calls;
story node/depth limits still apply to the story scanner, and the content
profile independently checks source, XML and projection bounds. The default
argument does not widen retained host admission.

Theme replacement returns are closed through the opaque-parent follow-up:
complete supported scheme state, bounded script counts, decoded namespace
identity, bounded duplicate tracking and modeled conformance checks preserve
foreign opaque descendants exactly. Diagram paired roots and graph ordering,
chart nested-label admission, and contextual SVG namespace/byte fidelity also
retain their reviewed fixes.

## Validation

All checks use the current root lock
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`, one Cargo
job, no incremental compilation, a 20 GB memory cap, and an external target.
The tested source index is `37dceb650c5c699a81229ac1dee6fbd98478b597`.

- DrawingML plus DOCX, PPTX, spreadsheet-drawing, XLSX, XLSB and OPS:
  all-features/all-targets locked offline check passes (309.729 seconds).
- DrawingML all-features/all-targets locked offline tests: 485 pass across
  22 suites, with no failures or ignored tests (51.976 seconds including build).
- DOCX projected-Ink profile tests: 15 pass, including opaque ancestry,
  effective defaults and retained payload allocation.
- PPTX asymmetric subordinate-reader input-policy regression: one passes.

The target ends at 5,997,810,347 bytes. Source index and root lock remain
unchanged through each passing command. Exact commands and external log
identities are retained in [`results/0362dvj/`](results/0362dvj/).

Two failed attempts are preserved. The first reused an older leaf target and
loaded stale ZIP/signing metadata despite current source signatures. Clearing
the 68 workspace packages' cached outputs, while retaining registry caches,
resolved those mismatches. The second exposed the four DOCX production calls
described above; their migration also updates the six associated tests. Neither
failure is relabeled as a passing check. Future reuse of a target across
different worktrees must invalidate workspace-package outputs first.

The earlier fragment checks remain diagnostic unlocked/6d evidence; this
record supplies the required assembled 4100 validation. The standalone SVG,
SmartArt and theme harness receipts retain their own source and lock pins.
Fresh all-workspace inventory, the full gate, corpus replay and native
qualification are still pending. Owner-transition measurements must use
`945eb1d6b9` as the pre-owner baseline, with separate raw-codec baselines.
Only redaction audit 22 remains deferred.
