# 0362ul165a — native protocol admission repairs

The DOC/Ink v5 and W8 File v4 packets now pass independent source review.
They are preparation for the required native work in 165, not build,
archive, fidelity, allocation or performance results. Frozen C and pending P
remain at `33a0a63d3553712c1276da592f7846a0c7dcabd6` while full65 runs.

## DOC admission returns

Review returns three successive defects before execution:

- v2 leaves two executable source-HEAD checks at retired 472b despite
  rebinding paths. v3 corrects them and tests the rendered source guard.
- Root then returns v3 because supplied output hashes can replace the four
  fixed fidelity expectations. A source-only reproduction proves that four
  arbitrary valid hashes pass admission and become the rendered expectations.
  Fresh output observations must match the original oracles, not select them.
- v4 retains the original oracles but its valid `bind` entry point recurses
  through the nested renderer substitutions. Tests of individual renderers
  had missed this public-entry-point failure.

v5 imports only the immutable v2 implementation, captures its original
renderer and restores it after `bind` and CLI `main`. It applies both C2a
HEAD checks and the four original 0359qha output expectations. Five source-only
tests exercise positive full bind and CLI wiring, renderer restoration and
fail-before-write behavior. The inherited source guard also rejects retired
HEAD and extra baseline edits. Synthetic admission fixtures do not stand in
for actual source, binary or native qualification.

Independent review verifies the eight packet identities, exact source and
three locks, unchanged original output oracles and retained path, binary,
freshness, overlay, sampling, profiling, timeout and quiet-host guards.
Earlier packets and returns remain intact; v3's earlier READY is superseded
by the root fidelity return.

## W8 binding returns and evidence recovery

Preparing a successor accidentally overwrites three files in the initial W8
packet cited by 165. Root stops further mutation. The successor is preserved
separately, and the original README, projection delta and manifest are
recovered from their exact retained tool-call patches. Root independently
checks all seven original seals and the exact eight-file set. The restored
manifest is again 1,678 bytes with SHA-256
`1bc410531bb48ed51c8166fb156288bf28d56d07c9a33c1f4bd0f581098c6299`.
No executed build, archive or measurement depended on the overwritten files.

The preserved v2 successor is returned for seven stale absolute paths after
its move. v3 corrects the paths, but review catches a stale pair-tool digest
in the quiet runner. v4 updates the full path/hash chain and tests the actual
source-check path through `root_paths`, `verify_digest`, source-preflight
policy identity, `load_reviewed_policy()` and fixed 109-route validation.

All 15 focused source-only tests pass. The three source transitions, twelve
dependency edges, three C2a locks, 109-route order, lifecycle limits, output
caps and fidelity checks remain intact. Real postgate artifacts remain absent;
synthetic complete-binding tests and stale/altered negatives make no native
acceptance claim.

## External evidence

Paths below are relative to `(deleted-coord-artifact)`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `doc-durable-c2a-native-prep-20261026-v3-root-oracle-return.json` | 1,514 | `b9c133586b445da0cfdb4075438771edf787783ef3d078e0b9154ca7ffa527d7` |
| `doc-durable-c2a-native-prep-20261026-v4-independent-review-return-20261026.json` | 3,684 | `cd7350480c2db8ea6ac063d3b7a33126737f3792e09993e8994b305361ee7903` |
| `doc-durable-c2a-native-prep-20261026-v5-independent-review-20261026/review-receipt.json` | 6,250 | `13730a38f39924a55cab10c5d12adc6a4591f6a1ad0f62be6e0cb9922817599e` |
| `w8-file-native-C2a-33a0a63d35-20261026-recovery/recovery-receipt.json` | 2,320 | `df212c4f1b2c22f31ca29e7c3ac096152d3ae91c8a4dd72131fc169c15eb66b7` |
| `w8-file-native-C2a-33a0a63d35-20261026-v3-independent-review-20261026/review-receipt.json` | 4,094 | `73a5dcaa562d1da5425afeba6d3ccd847ab51373f623e0030ea14499b2c60919` |
| `w8-file-native-C2a-33a0a63d35-20261026-v4-independent-review-20261026/review-receipt.json` | 5,442 | `9657918eed085bf1eba81bace04dffbdf157295513b533aeaf6daacd5ee3e0f2` |

The source-only packet manifests are DOC/Ink v5
`a6f158f621227fae484425798fa25d4bbb2af5a32cca1a4d57e7fb2d031fa599`
and W8 v4
`931da66e1e353533c070528b2bae8ad6d57d086a72cb09034be5390988a121a3`.

## ODF source leaf and remaining work

The ODF raw-codec probe, paired runner, preparation and tests are committed
on isolated `uo2/odf-raw-native-final` at
`48d68bcc29fb56afd3c20c54f6e518374a5e3d46`. Seven Python checks and rustfmt
pass; independent review is still open. There has been no Cargo build or
native execution, and the leaf is not merged or accepted as qualification.

Full65 remains live in the Ops integration chunks after 1,022 library tests
pass. All 25 qualification rows, final census, corpus postcheck and remaining
control cleanup stay required. Only memo 22 is deferred. Nothing is pushed.
