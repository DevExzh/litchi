# 0362ul16c — final W4b/W5 source preparation

This is the final source-only preparation packet for W4b XLSB and W5 legacy
native qualification at the frozen candidate `c796d0acc24c7c6e6fd08980b9e31af6bb86b18b`.
It makes no timing, allocation, correctness, or acceptance claim. The reduced
receipt is [results/0362ul16c/source-preparation.json](results/0362ul16c/source-preparation.json).

## Source pair

The fresh external source roots are:

```text
W4b before: /home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final/sources/w4b-before
W4b after:  /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-c796d0acc2/source
W5 before:  /home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final/sources/w5-before
W5 after:   /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-c796d0acc2/source-gate
```

W4b before is commit `be644568846ad2fcf0622386f2f2438d0eac21c0` with the
required synthetic pre-W4b tree `684a028b2ae30a3a110b87bf6743fc598c6691a2`;
the tree is not represented as a commit. W5 before is commit
`7afba5f146aa1a16eba73ebed2144dd041b33fb5`. Both candidate endpoints are
clean at the selected c796 ref and have tree
`6fe4f6f68046d2ee99dc7ca982004c406e484336`.

The current preparer applied only the reviewed compatibility overlays:

- W4b translates the three historical selector resolutions and the commit
  field in `xlsb_w4b_probe.rs`, and adds the recorded `Cell` trait import to
  the historical CRUD consumer.
- W5 translates the historical File options entry point and commit accessor,
  copies the recorded route module, and keeps both later DOC publication
  controls.

The baseline dirty paths are exactly those overlays. They are intentional and
are allow-listed by the manifests; no candidate source path is dirty.

## Locks, manifests and source identity

The source-preparation receipt is external at
`/home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final/receipts/source-preparation-c796.json`
(6,532 bytes, SHA-256
`dcd01ffdb8d1947e9e5f57e94d48b46b1aeae8372b633af379599a4110d88639`). It
was produced by the current preparer, not by treating the W4b tree object as a
commit. The static source-identity receipt is external at
`/home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final/receipts/manifest-source-validation-c796.json`
(33,283 bytes, SHA-256
`1d47716fdba5029a0e791bf66abf7fb771a27daf7f13821ceb43c6e9f8c9141d`). It
passes `load_manifest`, the workload checks, Cargo argv normalization,
source/target/artifact containment, validator bindings, and `_source_identity`
for both legs of both lanes.

The synthetic baseline commits do not track a root `Cargo.lock`. The exact
generated pre-wave root lock (`139,442` bytes, SHA-256
`cbe70791291d4a0830a8749985aa717230abf74d61c108fb284d25990c9ad681`) was
reused byte-for-byte from the previously verified baseline endpoint. The
tracked perf-baseline lock is `57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f`
(`57,028` bytes). The static runner check hashes these locks rather than
silently generating or accepting a different resolution.

The W4b fixture is the tracked
`test-data/poi/test-data/spreadsheet/testVarious.xlsb`; its absolute path and
hash are included in the static receipt. The preparer, runner and both current
manifests are likewise pinned there.

## Workloads and deferred execution

The W4b manifest binds the probe, sheet-dispatch and CRUD programs, three
release Cargo builds per leg, 31 rounds and exact
`before, after, after, before` order. The W5 manifest binds the source and
ordinary probes, one release build per leg, 31 rounds, the recorded DOC batch
round count of 7, and the same ABBA order. The correctness analyzers are the
tracked scripts from the selected candidate sources.

The protocol command passed all 41 tests; its external log is
`/home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final/receipts/protocol-tests.log`
(332 bytes, SHA-256
`744067b4ee12d09a8353ced5d9c89d19608ab311e78548ad08b07d2ce503c708`). No
Cargo build, executable probe/correctness run, or quiet-host ABBA run was
started because the native heavy grant was not available. The exact deferred
commands are listed in the reduced receipt and are ready for a separately
quiesced slot with fresh external targets and output directories. No binary or
large evidence file is committed.
