# Worktree cleanup — 2026-10-23

Two inactive merged leaves were removed with `git worktree remove` without
`--force`. Their branch refs remain available for reproducible source recovery.

- `uo2-xlsb-provenance-review`: HEAD `7b3fd27e2ec9f6be3103807932828f183981b4fb`, ancestor of integration `a814baa1de`; 37,389 bytes reclaimed. It had no ignored or untracked files.
- `uo2-iwa-detection-api`: HEAD `7f3aef217eb9090199b8389a8d70bca9633e4ee3`, ancestor of integration `a814baa1de`; 341,906,569 bytes reclaimed. Eleven generated ignored files (461,243 bytes) were moved first to quarantine, with relative paths and SHA-256 records.

Both trees were clean, had no targets, protected measurement binaries, or active
process/file-descriptor users. Their branch refs were checked after removal;
the full receipt records exact `git worktree add` commands for recreation.

The DOCX section/redaction leaves and the old W8 DOCX/RTF leaves contain ignored
performance captures, native bundles, and portable `.zst` evidence. The replay
error audit leaf has tracked changes. Those trees were retained. Full inventory,
quarantine manifests, hashes, and before/after worktree lists are in
`.coord/worktree-cleanup-2026-10-23/` outside the checkout.

## Second cleanup wave

Four additional inactive source-only leaves were removed on 2026-10-23 after
their commits were confirmed as ancestors of integration `0ef732727d`:

- `uo2/0362rla-opc-policy-sparse` at `a79627ee3593708d8b630acd35fd185d573f0dbf`:
  58,274,082 bytes.
- `uo2/consumer-v8-repair` at
  `65e8f94f74bc0dcaef8c0f1bda240d67ac6db6a4`: 6,258,992,707 bytes.
- `uo2/0362ug-odf-errors` at
  `3b47fca40e5327c808bf5d04e51ee34837866f1f`: 73,609,004 bytes.
- `uo2/odf-consumer-repair` at
  `702accfa6243ad96ef0585087f3f523ea5e73f74`: 93,648,857 bytes.

All four trees had clean tracked state and zero untracked or ignored files.
Target, binary, capture, native, and portable-evidence directory checks found
nothing local to those leaves, and the process/file-descriptor audit found no
active users after the audit commands completed. No files were quarantined.
The consumer repair checkout contained shared tracked `docs/performance/results`
content; it was not changed or removed from Git, and its branch ref remains
available. The four branch refs remain after removal.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362rla-opc-policy-sparse uo2/0362rla-opc-policy-sparse
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-consumer-v8-repair uo2/consumer-v8-repair
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362ug-odf-errors uo2/0362ug-odf-errors
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-odf-consumer-repair uo2/odf-consumer-repair
```

The second wave reclaimed 6,484,524,650 bytes. It used `git worktree remove`
without `--force`; no branch or tracked object was deleted.

## Third cleanup wave

Five further inactive source-only leaves were removed on 2026-10-23 after
their commits were confirmed as ancestors of integration
`18983596a3a1b0f2ab937f4fb1f7d9ef29f18f4c`:

- `uo2/consumer-v9-repair` at
  `59b9a87de1aa8162b1ca9fe8fdc5cca21ebb314b`: 124,336,611 bytes.
- `uo2/xlsx-current-consumers` at
  `1469c148f71cdd7dd494731ac378d60897c98c05`: 6,258,986,092 bytes.
- `uo2/opc-required-policy` at
  `d6ff788d3fe1d8a139b5e26d55102e274b1cd890`: 6,259,140,205 bytes.
- `uo2/0362ui-pptx-consumers` at
  `c077e9f8d19be821cd682d03d550b4d579f7b10a`: 64,327,811 bytes.
- `uo2/0362rnb-opc-recipient-policy` at
  `752a5fa9d1d0ba00ff5e2adb2144e0d46d192988`: 36,456,058 bytes.

The apparent checkout total was 12,743,246,777 bytes. The required-policy
checkout included four generated ignored Python bytecode files totaling
131,671 bytes; those were moved before removal, so the retained source/tree
total was 12,743,115,106 bytes. The quarantine is outside the checkout at
`.coord/worktree-cleanup-2026-10-23/quarantine/third-wave/`, with relative
paths, sizes, and SHA-256 values in `manifest.json`:

- `tools/__pycache__/api_debt_inventory.cpython-314.pyc`, 17,557 bytes,
  `cabb601951fd23533aee89139df55a11d29dd5cf28cd29960aabedad36437522`.
- `tools/__pycache__/check_api_glossary.cpython-314.pyc`, 44,684 bytes,
  `e9255bc1b6625cb53fffb246f8b8b8b3d8e8e09ca882b758648bedc0f82ec0ff`.
- `tools/__pycache__/rustdoc_return_types.cpython-314.pyc`, 15,603 bytes,
  `293c01eeb71adb89ed39789d0f27c469db1996fbc210dee64da3f9d63a0fcb90`.
- `tools/__pycache__/test_check_api_glossary.cpython-314.pyc`, 53,827 bytes,
  `ddff51a86d068e6ef5df9fc0f89356258bf442c9c173b7f8d24a5725afb1a1aa`.

All five trees had clean tracked state. Target, binary, capture, native,
portable-evidence, and other protected measurement checks found no local
artifacts, and the process cwd/file-descriptor audit found no active users.
Shared tracked `docs/performance/results` and `docs/report` evidence was not
changed or removed from Git; all five branch refs remain at the hashes above.
No active/source leaf was touched. Removal used `git worktree remove` without
`--force`.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-consumer-v9-repair uo2/consumer-v9-repair
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-xlsx-current-consumers uo2/xlsx-current-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362rd-substrate-scope uo2/opc-required-policy
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362ui-pptx-consumers uo2/0362ui-pptx-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362rnb-opc-recipient-policy uo2/0362rnb-opc-recipient-policy
```

The third wave reclaimed 12,743,246,777 checkout bytes while retaining all
source refs, tracked evidence, and quarantined generated files.

## Fourth cleanup wave

Six further inactive source-only leaves were removed on 2026-10-23 after
their commits were confirmed as ancestors of the read-only integration
checkout at `42b1c5a3c22deae4f8ca3b1d5a850ed5a13c1176`:

- `uo2/opc-stream-compile` at
  `3ced8395b6282e16c87d276614a3cacd94e56848`: 5,959,079,822 checkout bytes;
  46,418 ignored generated bytes.
- `uo2/docx-restore-capability` at
  `7753b2c0a2a713bb253f257c6b7e58edd864524a`: 6,260,593,652 bytes.
- `uo2/opc-external-consumers` at
  `1367a79d0547850df5555c3660ef829b53c53ca6`: 6,260,366,922 bytes.
- `uo2/workspace-opc-consumers` at
  `b485c0cec8e45d9912dd59319dafffb52cefcd53`: 6,260,590,215 bytes.
- `uo2/xls-final-followups` at
  `91a578db22a154046e616d5c2827c0a77e94bb08`: 6,259,355,983 checkout bytes;
  139,571 ignored generated bytes.
- `uo2/odraw-final-contract` at
  `841e7307983dd4fdaf90f61a2563d16d0e3d6438`: 6,259,502,437 checkout bytes;
  271,234 ignored generated bytes.

The apparent checkout total was 37,259,489,031 bytes. The quarantined
generated files totaled 457,223 bytes, leaving 37,259,031,808 bytes of source
trees. The full quarantine manifest, including relative paths, sizes, and
SHA-256 values, is outside the checkout at
`.coord/worktree-cleanup-2026-10-23/quarantine/fourth-wave/manifest.json`.
It contains the generated `Cargo.lock` files from the OPC stream, XLS, and
ODraw leaves plus four ODraw `tools/__pycache__` files.

All six trees had clean tracked state. Targeted checks found no untracked or
ignored target, binary, capture, native-bundle, portable-bundle, or other
protected measurement artifact. The shared tracked `docs/performance/results`
and `docs/report` native/capture evidence was not changed or removed from Git;
the six branch refs remain at the hashes above. Process cwd and file-descriptor
checks found no active users. The old signed and DOCX helper leaves remained
active and were not touched. No Cargo build or other build command ran.
Removal used `git worktree remove` without `--force`.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-opc-stream-compile uo2/opc-stream-compile
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-docx-restore-capability uo2/docx-restore-capability
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-opc-external-consumers uo2/opc-external-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-workspace-opc-consumers uo2/workspace-opc-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-xls-final-followups uo2/xls-final-followups
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-odraw-final-contract uo2/odraw-final-contract
```

The fourth wave reclaimed 37,259,489,031 checkout bytes while retaining all
source refs, tracked evidence, and quarantined generated files.

## Fifth cleanup wave

Three inactive source-only leaves were removed on 2026-10-23 after their
commits were confirmed as ancestors of the read-only assembly checkout at
`ff86f4a1c07f05a98a8dddf272daef56942656ea`:

- `uo2/opc-stream-runtime` at
  `9acb6932f6a0eed0c3b20421077a574745d89beb`: 6,260,627,662 bytes.
- `uo2/opc-token-runtime` at
  `f5d3387b3f9d236156f98c5b144660d6521452bf`: 6,260,628,960 bytes.
- `uo2/bench-matrix-assembly` at
  `c3a22a3a5f8eb668bb928c24d3c6bafdca756320`: 6,261,096,033 checkout bytes;
  323,671 ignored generated bytes.

The apparent checkout total was 18,782,352,655 bytes. The two generated bench
Python bytecode files were quarantined outside the checkout at
`.coord/worktree-cleanup-2026-10-23/quarantine/fifth-wave/`; their relative
paths, sizes, and SHA-256 values are in `manifest.json`. The retained source
tree total was 18,782,028,984 bytes.

All three trees had clean tracked state. Targeted checks found no local
untracked or ignored targets, caches, binaries, capture/native bundles,
portable bundles, or other protected measurement artifacts. Shared tracked
measurement references under `docs/performance/results` and `docs/report`
were not changed or removed from Git, and all three branch refs remain at the
hashes above. Process cwd and file-descriptor checks found no active users.
The signed replay, W8 DOCX, contract-inventory, boundary, XLS F3, XLSX final,
workspace-v16, and ZIP-runtime leaves remained protected and untouched. Trunk
`e0d46c99509073af99f2ce5119bf9e38db59717c` and the assembly checkout were not
modified. No Cargo or build command ran. Removal used `git worktree remove`
without `--force`.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-opc-stream-runtime uo2/opc-stream-runtime
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-opc-token-runtime uo2/opc-token-runtime
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-bench-matrix-assembly uo2/bench-matrix-assembly
```

The fifth wave reclaimed 18,782,352,655 checkout bytes while retaining all
source refs, tracked measurement evidence, and quarantined generated files.

## Sixth cleanup wave

Six completed source leaves were removed on 2026-10-23 after their commits
were confirmed as ancestors of assembly
`75fb774d28bc2b68274eb1abfb88fee71040c2f0`:

- `uo2/w8-docx-document` at
  `ff220b3b4ab4ea16e28cf9dce1b727f21dcba633`: 6,258,085,963 checkout bytes;
  230,568 ignored generated bytes.
- `uo2/workspace-v16-consumers` at
  `27d9541e46e36299776092f3bbb740803877860f`: 6,260,784,184 bytes.
- `uo2/zip-runtime-repair` at
  `30b9024f9018fa06534d4d73a241f0df4871ccfb`: 6,260,637,682 bytes.
- `uo2/boundary-assembly` at
  `4a0ca19bfd919b01748b84c985f8782886f1c44a`: 265,672,510 checkout bytes;
  6,153,449 ignored generated bytes.
- `uo2/final-contract-inventory` at
  `de105e8ccdca617c8c985c7de861e59ef6999bdc`: 6,260,701,614 checkout bytes;
  60,291 ignored generated bytes.
- The detached `uo2-xlsx-final-qualification` checkout at source commit
  `ff86f4a1c07f05a98a8dddf272daef56942656ea`: 6,260,782,279 bytes. It had no
  branch ref; the full source commit is retained in the receipt and exact
  recreation command below.

The apparent checkout total was 31,566,664,232 bytes. The eight generated
ignored files totaled 6,444,308 bytes and were quarantined outside the
checkout at `.coord/worktree-cleanup-2026-10-23/quarantine/sixth-wave/`, with
relative paths, sizes, and SHA-256 values in `manifest.json`. The retained
source tree total was 31,560,219,924 bytes.

The DOCX raw leaf referenced the approved external target
`/home/zhuhe/code/litchi-worktrees/uo2-docx-p0/target`; its measured 2,564,794,152
bytes were preserved and never touched. No removed worktree contained a local
target, cache, binary, capture/native bundle, portable bundle, or other
protected measurement artifact after quarantine. Shared tracked measurement
evidence under `docs/performance/results` and `docs/report` was not changed or
removed from Git. All five branch refs remain at the hashes above, and the
detached XLSX source commit remains recoverable by full hash. Process cwd and
file-descriptor checks found no active users. Signed lanes, ODF raw, XLS F3,
root, integration, and other unmerged leaves remained protected and untouched.
Trunk `ac16fb586ef3db067e08ff649bd6baf8cc253bdf` was not modified. No Cargo or
build command ran. Removal used `git worktree remove` without `--force`.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-w8-docx-document uo2/w8-docx-document
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-workspace-v16-consumers uo2/workspace-v16-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-zip-runtime-repair uo2/zip-runtime-repair
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-boundary-assembly uo2/boundary-assembly
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-final-contract-inventory uo2/final-contract-inventory
git worktree add --detach /home/zhuhe/code/litchi-worktrees/uo2-xlsx-final-qualification ff86f4a1c07f05a98a8dddf272daef56942656ea
```

The sixth wave reclaimed 31,566,664,232 checkout bytes while retaining the
external DOCX target, all source refs, tracked evidence, and quarantined
generated files.

## Seventh cleanup wave

Two completed inactive source leaves were removed on 2026-10-23 after their
commits were confirmed as ancestors of integration
`355cfe414118f0a76dc907a33893e938bfb0bebe`:

- `uo2/workspace-v21-consumers` at
  `ddb398129cf08f6d864e6092c73224ae2c5df045`: 6,261,086,693 bytes.
- `uo2/clippy-v2-consumers` at
  `a1697ddbdb1027f43021b0bbf3f1948d2ca69f45`: 6,261,100,997 bytes.

The apparent checkout total was 12,522,187,690 bytes. Both trees had clean
tracked state and zero untracked or ignored files, so no quarantine was
needed. Targeted checks found no local targets, caches, protected executables,
or measurement artifacts. Shared tracked `docs/performance/results` and
`docs/report` evidence, including receipt JSON under a tracked `binaries`
directory, was not changed or removed from Git. Process cwd and
file-descriptor checks found no active users. Dirty perf-v23, old ODF, budget,
XLS F3, active review/qualification, target, and executable leaves were not
touched. Removal used `git worktree remove` without `--force`.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-workspace-v21-consumers uo2/workspace-v21-consumers
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-clippy-v2-consumers uo2/clippy-v2-consumers
```

The seventh wave reclaimed 12,522,187,690 checkout bytes while retaining both
branch refs, tracked evidence, and all protected worktrees.

## Main target cache quarantine

On 2026-10-23 the idle main target at
`/home/zhuhe/code/litchi-worktrees/unified-ops-api/target` was reduced
reversibly. The integration target was `eb718eece7`; no Cargo, rustc, rustdoc,
or file-descriptor user was present. Before quarantine it measured
19,050,144,236 apparent bytes and 19,605,524,480 allocated bytes. Afterward it
measured 18,609,608,396 apparent bytes and 19,113,074,688 allocated bytes,
reclaiming 440,535,840 apparent and 492,449,792 allocated bytes.

The per-file manifest sums are 754,466,808 apparent bytes and 793,866,240
4-KiB-rounded bytes. Those are record sums, not reclaim totals: 10,940 records
collapse to 10,745 hard-linked inodes. The quarantined payload therefore has
440,535,840 apparent bytes under `du`; its 492,662,784 allocated bytes differ
slightly from the target delta because of directory accounting.

The external quarantine is
`/home/zhuhe/code/litchi-worktrees/.coord/target-cache-cleanup-2026-10-23/wave2/`.
Its pre-move manifest has 3,705 item roots and 10,940 file records; SHA-256:
`5d3b995657c7c477472c9475d30a62608b0b37ffa3cec2c1f64b7efea76e88e2`.
`postcheck.json` reports all 3,705 source paths absent and destination hashes
verified. The exact restore command is:

```sh
python3 /home/zhuhe/code/litchi-worktrees/.coord/target-cache-cleanup-2026-10-23/wave2/restore.py
```

The moved generated cache was limited to:

- 361 old `target/debug/build` directories (645,498,373 apparent bytes) that
  no dep-info file references, plus 361 matching old fingerprint directories;
  the closure guard retained all 32 build-output names referenced by dep-info.
- The two known stale unit-test groups `e5a12b34de2106e0` and
  `4040652213167156`, including sidecars and fingerprints (101,933,591
  apparent bytes); the two main files are ELF unit-test executables, not
  release or measurement binaries.
- 2,977 fingerprint-only directories (6,788,484 apparent bytes) with no
  matching output hash anywhere outside `.fingerprint`.

The rejected first selection was fully restored before the corrected wave2.
The external receipt
`/home/zhuhe/code/litchi-worktrees/.coord/target-cache-cleanup-2026-10-23/wave1/restoration.json`
cites the wave1 manifest SHA-256
`e0997eb7f0ba67b91e95e3831cd15fd5c3a169fcec62f24f13657973f195629b`, the
successful `restored 23027 items` output, and wave1's verification of all
92,487 file hashes. Wave2 later re-quarantined 3,703 overlapping roots and
10,938 of those file records, plus two corrected sidecars; the wave1 payload
is empty.

The compiled historical `.rmeta`/`.rlib` groups whose dep-info names other
worktrees were not moved: Cargo crate-hash dependency closure could not be
proven without a build. The all-feature root
`liblitchi_ops-b00e406e2533c222.rmeta`, format-only root
`litchi_ops-572a08bc21edd59c`, every artifact at or after the 14:30 UTC cutoff,
all release/measurement examples, and the isolated wsdoc target were retained.
Moved executable files consisted of generated Cargo build-script outputs under
unreferenced `target/debug/build` directories and exactly the two stale unit-test
ELF executables named above; no release or measurement binary was touched. No
tracked file or PROGRESS file was changed.

## Eighth cleanup wave

Three clean inactive source leaves were removed on 2026-10-23. Their branch
refs remain available for source recreation:

- `uo2/0362us-native-budget` at
  `78ad0abc0ea4ec642eb0da27b0fab693048ea45b`: 6,261,575,540 checkout bytes;
  its tip is an ancestor of `uo2/budget-validation` and has no commits unique
  against that validation branch.
- `uo2/budget-v3-native-errors` at
  `4cd986a67c0581234c4c3f6888a2acca41c6aed8`: 6,261,550,620 checkout bytes;
  its tip is an ancestor of `uo2/budget-validation`.
- `uo2/clippy-v3-consumers` at
  `ed359c95c19e02aafdb208dc4e68383caf98e005`: 6,261,108,633 checkout bytes;
  its tip is an ancestor of integration `eb718eece7`.

All three trees had clean tracked state and no untracked or ignored files.
Targeted checks found no target, binary, capture, native, portable, or other
protected measurement artifact. Process cwd and file-descriptor audits found
no active users. Active budget-validation and other budget leaves were not
touched. Removal used `git worktree remove` without `--force`; no Cargo or
build command ran. The wave reclaimed 18,784,234,793 checkout bytes.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362us-native-budget uo2/0362us-native-budget
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-budget-v3-native-errors uo2/budget-v3-native-errors
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-clippy-v3-consumers uo2/clippy-v3-consumers
```

## Ninth cleanup wave

Two inactive source leaves were removed on 2026-10-23 after their exact tips
were confirmed as ancestors of diagnostic `af09b4fcf92a54df5ae23d211a94f682357887d3`:

- `uo2/xlsx-feature-budget` at
  `750d91838625ec7b4139e430bdf3bb2978131b2b`: 6,257,520,792 apparent
  checkout bytes.
- `uo2/docx-snapshot-budget` at
  `90641c24e8937f5913800b9373eb1a0cf7910578`: 6,261,107,545 apparent
  checkout bytes.

The apparent checkout total was 12,518,628,337 bytes. The XLSX leaf had two
ignored generated files totaling 151,969 bytes (`Cargo.lock` and one Python
bytecode file). They were moved before removal, with relative paths, sizes and
SHA-256 values recorded in the external manifest
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/ninth-wave/manifest.json`.
Its manifest SHA-256 is
`1f2234070cc2467329771120b83125f21c10788dc77192ae296b4a3f25aa861c`.
The retained source-tree total was 12,518,476,368 bytes. The reversible restore
command is:

```sh
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/ninth-wave/restore.py
```

Both removed leaves were clean after quarantine, had no targets, binaries,
captures, native bundles, portable bundles, or protected measurement artifacts,
and had zero active process cwd/file-descriptor users. Removal used
`git worktree remove` without `--force`; the main integration tip
`3cad59be707c7431bc1a8c115e1ff534a048dc57` and diagnostic tip were not
modified. The branch refs remain at the hashes above and no tracked evidence
or PROGRESS file was changed. No Cargo or build command ran.

Exact recreation commands:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-xlsx-feature-budget uo2/xlsx-feature-budget
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-docx-snapshot-budget uo2/docx-snapshot-budget
```

The raw-helper review leaf `uo2/w4-raw-capability-review` at
`81b4bd433b00ad11887681f4113aeaf7b74bb1db` was retained: it has a tracked
change to `tools/perf-baseline/src/docx_bounded_tail_append_compare.rs`, an
ignored `target/`, and an active descriptor observed during the audit. The
signed XLSB leaf `uo2/signed-xlsb-budget` at
`6521e6b07a683e1585cad7dfd76d707474d90b49` was clean and merged into
diagnostic, but its ignored `target/` was protected and no target quarantine
was authorized. The Numbers limit payload leaf `uo2/iwa-limit-payload` at
`734ffb83ddee85a6707c3c81e8f4e27a1d69ba77` is not an ancestor of diagnostic
(it is an ancestor of integration), and its ignored `target/` was likewise
retained. Active signed-native/cell, XLSX metadata-admission, PPTX dead-helper,
DOC qualification, budget-validation, main/trunk, build-target and all
measurement leaves were untouched.

The ninth wave reclaimed 12,518,628,337 apparent checkout bytes while
retaining 151,969 generated bytes in the external quarantine and preserving
all source refs and protected build/measurement artifacts.

## Tenth cleanup wave

The three requested leaves were checked against main integration
`f994ff2029af03be19b066d10ae279889a3eb74e` and frozen diagnostic
`2cadd934b7d2ed5f6411d6fddf165bb8b16bc3ae`; all three exact tips were
ancestors of diagnostic. Their tracked trees were clean, and the only ignored
entry initially reported in each was the generated `Cargo.lock`. No local
target, binary, capture, native, portable, or other protected measurement
directory was present at the top level.

The clean inactive metadata leaf was removed without `--force`:

- `uo2/xlsx-metadata-admission` at
  `c67826a6bf53495e37e9da5fbb6cd8508c3de66f`: 6,262,042,222 apparent
  checkout bytes.

Its ignored `Cargo.lock` was moved before removal. It is 139,592 bytes with
SHA-256
`1d44ce42d8a66dd3967e80524d9cd830e15a2697646c43d84492628692361dc0`.
The external manifest is
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/tenth-wave/manifest.json`
with SHA-256
`da37f6d387c26c9bc561b1a32305ad9a16755f4fe4d4e3be841acf00de8db44a`.
The exact restore command is:

```sh
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/tenth-wave/restore.py
```

The consumer-facade and rich-transfer leaves were retained despite clean
tracked state and merged ancestry because the process/file-descriptor audit
found active users:

- `uo2/consumer-v9-facade-batch` at
  `6995c497f5570c92fe2f6cc5b7e9eef1b6095a48` (6,262,079,889 apparent
  bytes) had PID 647768 holding a descriptor under
  `docs/performance/results/change-0806/profile/complete.json`.
- `uo2/xlsx-rich-transfer` at
  `044c37c51936ec4ae1a08efc0e2192b605b72368` (6,262,105,244 apparent
  bytes) had active shell/git processes with cwd in the worktree, including
  PIDs 3645907, 3655429, and 3663013, plus descriptors held by the ongoing
  repository scan process.

The apparent total for all three candidates was 18,786,227,355 bytes. This
wave reclaimed 6,262,042,222 apparent checkout bytes; the other 12,524,185,133
bytes remain protected. No active process, external target, tracked
measurement evidence, dirty leaf, signed/replay/PPTX/main/budget/Ink/ODP/IWA
lane, or PROGRESS file was touched. No Cargo or build command ran.

All three branch refs remain available. Exact recreation commands are:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-xlsx-metadata-admission uo2/xlsx-metadata-admission
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-consumer-v9-facade-batch uo2/consumer-v9-facade-batch
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-xlsx-rich-transfer uo2/xlsx-rich-transfer
```

## Eleventh cleanup wave — 2026-10-23

The coordinator verified the external receipt and artifact identities before
recording this completed cleanup on the retained integration branch.

### Removed cleanup receipt worktree

The inactive, tracked-clean `uo2/rtf-refusal-perf` worktree was verified at
`3baa225cf8428b585e0f0a4b6113ac00a34a93bd`, an ancestor of the supplied main
integration target `87ce73a7363c1e9f1d4587c616bcc01d2bf08256`. It had no
process cwd or file-descriptor users. Its only ignored entry was `Cargo.lock`:
139,553 bytes, SHA-256
`925d25101fd62063b9bc6f5f95e504eb0c9fe492167de3504345eef8d11f067f`.

The lock was moved to
`quarantine/eleventh-wave/cleanup-lock/Cargo.lock`, with metadata in
`cleanup-lock-manifest.json` and restore helper
`restore-cleanup-lock.py`. The worktree was then removed with
`git worktree remove` without `--force`. Its branch ref remains intact.
The removed checkout path represented 6,257,821,238 apparent bytes; the lock
remains in the external quarantine and is not counted as filesystem space
freed.

Recreate and restore:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-rtf-refusal-perf uo2/rtf-refusal-perf
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/eleventh-wave/restore-cleanup-lock.py
```

### Relocated abandoned ODF partial

`/home/zhuhe/code/litchi-worktrees/uo2-odf-facade-runtime-v12-partial` was
confirmed to be an unregistered duplicate checkout. Its `.git` file points to
the valid registration
`/home/zhuhe/code/litchi/.git/worktrees/uo2-odf-facade-runtime-v12`, whose
real registered leaf remains at
`/home/zhuhe/code/litchi-worktrees/uo2-odf-facade-runtime-v12` on branch
`uo2/odf-facade-runtime-v12`, HEAD
`dbcb23941b0f31909c7ded27269fd14c0c9aa5f2`.

The partial had no active process cwd/file-descriptor users. Index-only audits
found zero nonignored untracked and zero ignored entries. Its pre-move
manifest contains 48,383 regular files, 1,060,162,316 file bytes, and SHA-256
`1a2c893250dd9a2c633171ab22014fb1b5a3983a665b9973d771bbddbc2e3645` for
`partial-manifest.json`. No top-level target, binary, capture, native, or
portable directory was present; tracked performance evidence inside the
partial is preserved in the quarantine and remains recoverable from the
branch objects.

The entire directory was moved, not deleted, to
`quarantine/eleventh-wave/partial/`. This relocation is reversible and is not
claimed as filesystem space freed. Restore with:

```sh
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/eleventh-wave/restore.py
```

The restore helper requires the original partial path to be absent and leaves
the valid registered leaf untouched. The external manifest records the
exact source pointer, branch, HEAD, paths, modes, sizes, and SHA-256 values.

### Retained signed replay leaf

`/home/zhuhe/code/litchi-worktrees/uo2-signed-replay` at
`06afe43fa08a4bddbdab0745b54042f6106dde8e` was assessed read-only and
retained. It is 10,441,206,655 apparent bytes, has no tracked/untracked
changes or active process users, and contains protected ignored
`.coord/signed-native-repair/target` evidence plus generated lock/bytecode.
No target or measurement artifact was moved.

External receipt: `/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/eleventh-wave/receipt.md` (3,292 bytes, SHA-256
`ea5e131486b2a04a9c3b2088c778941ec82740ed83cc8cab0346945bd6ec5240`). The complete external partial manifest is
9,604,583 bytes. To regenerate its original-path manifest, first restore the
partial with `python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/eleventh-wave/restore.py`,
then run `python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/eleventh-wave/build-manifest.py`.
No complete manifest or target binary is committed.

## Twelfth cleanup: four merged source leaves

The following tracked-clean, inactive leaves were verified as ancestors of
integration `d20e6ffbaa6d5907db0e31bf4c03732df7937511` (which contains
diagnostic `f25b55ca77f2fdf262975dce8c1f7338406bc65f`). Each had zero active
process cwd/file-descriptor users and no top-level target, binary, capture,
native, portable, or protected measurement directory:

- `uo2/consumer-v9-facade-batch` at `6995c497f5570c92fe2f6cc5b7e9eef1b6095a48`:
  6,262,079,889 apparent bytes.
- `uo2/coordinator-durable-ink` at `f97f0eef7a3622f68cd972f2b183c721dfedcf23`:
  6,255,818,948 apparent bytes.
- `uo2/coordinator-return-contract` at `0dfb00b5902b84179485f9756b7454efaab6920d`:
  6,262,741,276 apparent bytes.
- `uo2/coordinator-w8-part-a` at `fd5f513df62ee8a824d3bb9a6fed4b857514b5dd`:
  6,256,383,347 apparent bytes.

All four were removed with `git worktree remove` without `--force`. Their
branch refs remain. The apparent checkout total removed was 25,037,023,460
bytes. Twenty-five ignored generated lock/bytecode files totaling 8,173,100
bytes were moved, not deleted, to the external quarantine. The manifest with
relative paths, sizes, and SHA-256 values is
`quarantine/twelfth-wave/manifest.json`, SHA-256
`0b2f36891d121215a0cc7efdc903595b5df6de4886f0c971839c00aa27fdca91`.

Recreate the branches and restore those generated files with:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-consumer-v9-facade-batch uo2/consumer-v9-facade-batch
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-coordinator-durable-ink uo2/coordinator-durable-ink
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-coordinator-return-contract uo2/coordinator-return-contract
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-coordinator-w8-part-a uo2/coordinator-w8-part-a
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twelfth-wave/restore.py
```

## Leaves retained

The coordinator constructors leaf has an ignored 24 GB target and an active
measurement descriptor (PID 3385965). The coordinator W4 leaf has active
measurement descriptors (PID 647768). W7 DOCX and XLSX rich-transfer leaves
have an active evidence descriptor (PID 1547392). DOC toggle has a protected
target. DOCX ink-perf and OPC stream-restore retain tracked performance
evidence. These were not touched.

All explicitly active/protected lanes were retained: trunk, coordinator
metadata integration, budget validation, signed PPTX/XLSB/replay, final
contract inventory, DOCX runtime, XLSX normalization, IWA qualification,
legacy safe-runner, W8 readiness, and all `.coord` baseline/build/evidence
worktrees. Targets, native binaries, external evidence, dirty leaves, and
measurement artifacts were not moved.

## ODF and legacy incomplete checkouts

The registered `uo2/odf-facade-runtime-v12` leaf at
`dbcb23941b0f31909c7ded27269fd14c0c9aa5f2` remains untouched. It is not a
sparse checkout: 460,304 tracked paths are absent from its worktree. No
discard or force removal is safe. A future cleanup must first restore the
missing tracked paths from `HEAD` (or create and verify a separate complete
checkout), then re-audit users and protected evidence before any removal.
The empty `uo2/legacy-replay-errors` worktree was likewise retained because
its tracked paths are absent and its state is ambiguous.

The coordinator verified all 25 quarantine file hashes and sizes, retained branch refs and removed path absence. The external receipt is /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twelfth-wave/receipt.md (3560 bytes, SHA-256 `657e21e9403c8fd83befe96ba4f0078173f1e38301329fcc743a0d17c600c96c`). The manifest is 4,803 bytes. No build or measurement artifact was removed.

## Thirteenth cleanup: five merged source leaves

Against integration `5dae0abf78aa325a6ce395355108624ed5e66e1d`, the five
approved leaves were rechecked immediately before removal. All had clean
tracked/index state, zero nonignored untracked files, zero active process
cwd/file-descriptor users at the final guard, and no protected top-level target,
binary, capture, native, portable, or measurement directory:

- `uo2/0360adc-raw-aggregate-sparse` at
  `f3588167a1ed7739bfd4b7a476d0a7b4a1778e28`: 71,865,303 apparent bytes.
- `uo2/0362re-docx-section` at
  `4004efa04f8dd8621ab66c5914844a7617656fcb`: 6,257,806,042 bytes.
- `uo2/0362rg-docx-redact-sanitize` at
  `52d9edc77cbcc605ffd3f94434b637dc8f90644e`: 6,257,835,178 bytes.
- `uo2/0362ria-error-classification` at
  `f716e4b329dbabd4863cd6f8d8a19002596c15d5`: 185,226,211 bytes.
- `uo2/w8-docx-paragraph-copy` at
  `e1af87ed91d358488122877095d7111cab7553b0`: 6,256,072,690 bytes.

Each worktree was removed with `git worktree remove` without `--force`; every
branch ref remains. The measured apparent checkout total is exactly
19,028,805,424 bytes. The earlier candidate note said 18,029,455,424 due to
an arithmetic error; there was no filesystem measurement discrepancy.

Ten ignored generated files totaling 739,055 bytes were moved, not deleted,
to `quarantine/thirteenth-wave/ignored/`. Their relative paths, sizes, and
SHA-256 values are in `manifest.json`, SHA-256
`5f23fb94827c4b60a928690eab204b02928acb27658f2015128a3d3736aede53`.

Exact recreation and restore:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0360adc-raw-aggregate-sparse uo2/0360adc-raw-aggregate-sparse
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362re-docx-section uo2/0362re-docx-section
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362rg-docx-redact-sanitize uo2/0362rg-docx-redact-sanitize
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-0362ri-error-classification uo2/0362ria-error-classification
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-w8-docx-paragraph-copy uo2/w8-docx-paragraph-copy
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/thirteenth-wave/restore.py
```

The coordinator independently verified all ten quarantine sizes and hashes,
retained branch identities, source ancestry and removed path absence. The
external receipt is `/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/thirteenth-wave/receipt.md` (2317 bytes, SHA-256
`0601cab17193857fdef67156c76ab7e67301b94b2ef9501c28a259f0e1f323bf`). The manifest is
2791 bytes. No target, native binary or measurement evidence was removed.

## Fourteenth cleanup: reviewed ODG source leaf

After the ODG merge, `uo2/docx-runtime` at
`0dca7284f9eb826d2f4361af1f3fef047d924811` was freshly checked as clean,
inactive and ancestral to integration `8779d059ad`. It had no nonignored
untracked files or live process cwd/file-descriptor users. Removal used
`git worktree remove` without force and retained the branch. The checkout
represented 6,263,368,185 apparent bytes.

Three ignored generated files totaling 205,668 bytes were hash-quarantined
under `.coord/worktree-cleanup-2026-10-23/quarantine/fourteenth-wave/`.
The external receipt is 1,608 bytes, SHA-256
`9c8594161122e80f4d50dfd442e75fbf7c7ab570b2b0bb0df3155f62f42f69cb`.
The 1,213-byte manifest has SHA-256
`bf72491ca41d33f6c218f109d2d6512efa2086ab728de2b2124bc0132e3b561a`;
the 1,513-byte restore helper has SHA-256
`a13f236862a31bc758ef0a60117ef68a967975a55072ff255568c5f8c3233619`.
The coordinator independently verified all retained file identities, branch
head, ancestry, removed path absence and protected external root existence.

Recreate and restore with:

```sh
git worktree add /home/zhuhe/code/litchi-worktrees/uo2-docx-runtime uo2/docx-runtime
python3 /home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/fourteenth-wave/restore.py
```

External `.coord/docx-runtime` and all `.coord/odg-joined*` targets and evidence
remain untouched. Active sparse correction leaves and protected measurement
worktrees were excluded from this cleanup.

## Fifteenth cleanup, 2026-10-24

After reviewed source integration at `3676e21cfa`, the clean, inactive DOC
lazy-entry worktree was removed without force. Branch `uo2/doc-lazy-entry`
remains at `2a5a7c3facebf8b7d7598170953353e83caee5e8`. Its apparent checkout
size was 314,138,078 bytes. The ignored 139,591-byte Cargo.lock was moved to
external quarantine; SHA-256 is
`76340b069df633ff1c63c98e93e8b10c1e7711164ca2515f490872142ab46c72`.
The external DOC target and logs remain unchanged. The coordinator verified
ancestry, retained ref, removed path, receipt identities and quarantined bytes.

The receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/fifteenth-wave/receipt.md`
(SHA-256 `9daa8a99f905b8d4372146d3038f4d8755c58e167d68e1023eb13d8a5589480c`).
The adjacent manifest has SHA-256
`6793d485b5dfacc3f0e184a83cc7509d1d722a746785f053181af5a6634a6772`;
`restore.py` recreates the ignored file after restoring the retained branch.

ODB/ODM was rechecked as clean and merged, but a live search still held its
evidence files. That worktree was left intact pending the search owner's
confirmation and a fresh handle audit. No measurement artifacts, targets,
branch refs or incomplete source worktrees were deleted.

## Sixteenth cleanup, 2026-10-24

After the owners stopped their stale searches, a fresh handle audit found no
users of the clean, merged ODB/ODM leaf. It was removed without force; branch
`uo2/odb-odm-lazy-entry` remains at
`dba6b055603ae7ff0b7b488ea41510a8f1a164ab`. The apparent checkout size was
6,263,385,091 bytes. Three ignored files totaling 205,680 bytes are preserved
in external quarantine with their hashes and restore instructions. The
coordinator independently verified the files, receipt, ancestry, retained ref
and removed-path absence. The external target/log root remains unchanged.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/sixteenth-wave/receipt.md`
(SHA-256 `540b65f860ad6764acfb4fd70b6b50bf0f5632a8f775d2f1c4484baf040f1c9e`).
The adjacent manifest has SHA-256
`4c2d6760f6722120286db58fd290a250d3a2d74ec02de3f5ddf067c3985ae0f5`.
`restore.py` restores the three ignored files after worktree recreation.

## Seventeenth cleanup, 2026-10-24

A fresh process audit allowed removal of the clean, merged ODS lazy-entry
worktree without force. Branch `uo2/ods-lazy-entry` remains at
`daf6a85bd40b6d7e714091f691cc7c85c4327808`, ancestral to integration
`a3eb017da2`. The apparent checkout size was 440,249,548 bytes. Its sole
ignored file, the 139,591-byte Cargo.lock, is hash-quarantined with SHA-256
`76340b069df633ff1c63c98e93e8b10c1e7711164ca2515f490872142ab46c72`.
The coordinator independently verified its bytes, retained branch, ancestry,
removed-path absence and the protected external target/log root's existence.
That 3,025,233,257-byte external root remains untouched.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/seventeenth-wave/receipt.md`
(1,535 bytes, SHA-256
`1e6bb0a6aacfe283a36fbc55abdd6316bc5f8ff93b9db3f8a48eb77ed0ad61c5`).
The adjacent 622-byte manifest has SHA-256
`41d818e1e8929cb44d19ca64a97ef649fd085c0932920ad91e2b15e43cd1854c`;
the 1,419-byte restore helper has SHA-256
`c38ee4c6558f7a6872a076a246912061a64125730aa9dbf6e4bdca3d24b6d068`.
Recreate with `git worktree add` using the retained branch, then run that
`restore.py`. XLS and PPTX leaves remain untouched because foreign searches
still hold their files. No build target or measurement artifact was removed.

## Eighteenth cleanup, 2026-10-24

After the foreign PPTX search exited, a fresh audit found no process users of
the clean, merged PPTX lazy-entry leaf. It was removed without force; branch
`uo2/pptx-lazy-entry` remains at
`e59c72184dde29c2820c5d42a01d22de3ed532ea`, ancestral to integration
`d0e1782f9b`. The apparent checkout size was 362,854,381 bytes. Its ignored
139,592-byte Cargo.lock is quarantined with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.
The coordinator verified the quarantine, ancestry, retained ref and removed
path. The external PPTX target/log root remains at its protected location;
its recorded apparent size is 4,888,455,639 bytes. No target or measurement
artifact was deleted. The foreign-held XLS worktree remains untouched.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/eighteenth-wave/receipt.md`
(1,714 bytes, SHA-256
`3915081cbca99113e2168143409e845a63fca85e5cee34966c40add0f3246bcf`).
The adjacent 624-byte manifest has SHA-256
`1d19a0a91f68870a0b01ea4bb14e668dcc63640da0436a7c98c07cbe7153a1cf`;
the 1,419-byte restore helper has SHA-256
`a1f109fcc2acd13afa949ac1e682cfd2a89b30337956f185d0c52d4ce980469b`.
Recreate the worktree from the retained branch, then run that `restore.py`.


## Nineteenth cleanup, 2026-10-24

After iWork integration at `9ea602e0e1` and XLSB integration at `0d35a83a31`,
fresh ancestry, cleanliness and process audits permitted removal of three
completed iWork leaves without force. Their branches remain:

| Leaf | Retained tip | Apparent checkout bytes |
| --- | --- | ---: |
| Keynote | `c7e49c5b35b4d9483bb29ff73a40a791edef5416` | 411,791,360 |
| Numbers | `ef0954e86ea25957e92529925907a4fb14ffdc98` | 412,499,968 |
| Pages | `4441c86bbeb6ab11202ba2156cf73d18a8a44323` | 412,762,112 |

Total apparent checkout size was 1,237,053,440 bytes. Four ignored lockfiles
(total 455,063 bytes) were hash-quarantined with restore instructions. The
coordinator independently verified all four file identities, retained refs,
ancestry, removed paths and protected external-root existence. External
Keynote, Numbers and Pages targets/logs/evidence remain untouched. XLSB was
left intact because foreign searches held source descendants; its lockfile
was not moved. No target or measurement artifact was removed.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/nineteenth-wave/receipt.md`
(3,049 bytes, SHA-256
`a707b90ee0efd1df5862eafec108ebdea143cd5a4c8358a969a9f28c669b2c3c`).
The adjacent 2,435-byte manifest has SHA-256
`b0ae741e22054766a4438229347803b367d8e9c4172c41f7eba6e86513fc0fd0`;
the 1,637-byte restore helper has SHA-256
`b835c5ffeadcb9745344b3a608fd54729a305d7804c489c0d2cb16709b46deb7`.
Recreate each worktree from its retained branch, then run that `restore.py`.

## Twentieth cleanup, 2026-10-24

After the foreign searches exited, fresh process and ancestry checks against
integration `e862b653d6` permitted removal of the clean XLS and XLSB leaves
without force. Their branches remain at `8c2b79f7c5309636b990d3502f85037c4a493345`
and `16ba8da04f470caf53d39192cdd090c0929452e7`. The removed checkouts occupied
411,664,384 and 411,815,936 apparent bytes, totaling 823,480,320 bytes.

Both ignored Cargo.lock files are quarantined separately, each 139,592 bytes
with SHA-256 `6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.
The coordinator independently verified the quarantined files, retained refs,
ancestry, removed paths and protected external-root existence. External targets,
logs and measurement artifacts remain untouched. The active XLS property,
XLSX, PPT and DOCX leaves were not modified.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twentieth-wave/receipt.md`
(1,954 bytes, SHA-256
`a3a736b24399784af229a4d4b2f4bc57f26bddff93071e84204644d3bac89df3`).
The adjacent 1,180-byte manifest has SHA-256
`97353d945b51411bac8fa6278c37cb67c3e7fd72644d4230a79c6e28eb63e3c0`;
the 1,636-byte restore helper has SHA-256
`6bfa5ece6b163cd6ebc8ac41b33d3a80af226b90e6280bbb5ac22fa0ccfb3525`.
Recreate each worktree from its retained branch, then run that `restore.py`.


## Twenty-first cleanup, 2026-10-24

After the XLSX composition at `3d129677ce`, fresh ancestry, cleanliness and
process checks permitted removal of the Pivot, Graphs and Objects leaves
without force. Their retained tips are `83a1b6a0901c066d46b334a4f8aeb01803c8a1ce`,
`672b7498e8881262ded29010fc95864b7cda63a9` and
`2ea78edd48e99795e85db85e8632264acc4379eb`. Apparent checkout sizes were
411,680,768, 411,684,864 and 411,668,480 bytes, totaling 1,235,034,112 bytes.
No ignored or untracked files required quarantine; no external target root
existed for these three leaves. The coordinator independently verified all
receipt hashes, retained refs, ancestry and removed paths.

The core XLSX leaf remained intact because live review jobs held its source
files. Its ignored local target and lockfile, and the separate external target
and logs, were not moved or deleted during this cleanup.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twenty-first-wave/receipt.md`
(2,478 bytes, SHA-256
`281713d334ff9fe58667af7a246ded708117ad4b264f35964a68f963e67b8a8d`).
The adjacent 1,885-byte manifest has SHA-256
`80c14d48d5b031b1cfd4dcc3d4ec379bd0f3a63874659419681426bbf1570893`;
the 1,639-byte restore helper has SHA-256
`520f6e6dd99dcae416323df4d8eec0ebb7bf2f9b0d884c13193428035549381d`.
Recreate each removed worktree from its retained branch.


## Twenty-second cleanup, 2026-10-24

After the XLSX reviewer stopped its own stale Git jobs, fresh process and
ancestry checks against `5ce9829784` permitted removal of the XLSX core and
DOCX lazy-entry leaves without force. Their retained branch tips are
`7201866b326fb097af6b10ab29f3720c7341bfca` and
`3a6c56958c7d42cdcf9bc7374c1a00bb50e83ec5`. Source checkouts after artifact
relocation occupied 414,289,920 and 411,680,768 apparent bytes, so removal
reclaimed 825,970,688 apparent bytes.

The ignored local XLSX target and lockfile, plus three DOCX lockfiles, were
moved into external quarantine. The coordinator independently verified all
2,937 file sizes, SHA-256 values and modes, along with retained branch refs,
ancestry, removed source paths and protected external-root existence. Summed
regular-file sizes are 1,696,689,675 bytes; hardlinks make the quarantine's
allocated size 1,575,936,000 bytes. These relocated artifacts are preserved,
not counted as reclaimed space. External lane targets and evidence remain
untouched.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twenty-second-wave/receipt.md`
(2,849 bytes, SHA-256
`84590894cc3ca47da9d7a6466cb54d7fdf2bb4139378ca04e5fa976648e31b8c`).
The external 909,967-byte manifest has SHA-256
`4eac8d8d22cd74c81b8f1f3d2a279667793995a3ccaab6fb055809b70091df4b`.
The 8,089-byte mover has SHA-256
`d8b8466bfb044ebda5ccfe997d9f9a7fefb58d5ff109473f2ab3445a41dc4f6d`;
the 2,066-byte restore helper has SHA-256
`eb1e8d0b053c5ddd8e4a3dfbfd4f2a60d91e3a7f3c64832a599827f289515558`.
Recreate both worktrees from their retained branches, then run `restore.py`.


## Twenty-third cleanup, 2026-10-24

Fresh checks against `c1eda2fdf6` permitted removal of the completed PPT
and PPTX refusal-kind leaves without force. Their retained refs remain at
`70a9b5762258a69cec53c04b1b7d30b5e8286703` and
`dd39bbe82eff43b63d6f129a86a20027fe88496e`. The pre-relocation checkout
sizes were 412,717,056 and 7,551,844,352 apparent bytes. Eight ignored
generated files totaling 761,901 bytes were separately quarantined; their
806,912 allocated bytes are preserved and not claimed as reclaimed space.
The coordinator independently verified every quarantined file, retained ref,
ancestry, removed path and protected external-root existence. No external
target/evidence or active XLS-property leaf was modified.

Receipt:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twenty-third-wave/receipt.md`
(2,310 bytes, SHA-256
`c146b2403d9956d40226dc824c7457cea9f13db1337402b9f1f9055e2e4f3ca0`).
The adjacent 2,096-byte manifest has SHA-256
`99005a6ee2747022374a71510cd4ae25dd86d8b62ed34914623c09fe75e2510e`;
the 1,639-byte restore helper has SHA-256
`86e0fa2c95a4902ddc2707578922ab5214e06ebb17b2ead763fe3776d4230300`.
Recreate the leaves from their retained branches, then run `restore.py`.


## Twenty-fourth cleanup, 2026-10-24

Fresh checks against `f64213a09b` permitted removal of four merged leaves:
W8 footnote limits, DOCX raw, DOCX SVG and DOCX variables. Their retained tips
are `660205a29cdc384646dbeadad1a34e8ca1a89cf2`,
`b96137acb4d0db0405f99f846e185022c3534bab`,
`135a7b6e635c4ac679b19897e49ea416cc78efa3` and
`7dae046af8de8426fd78beded6bdbc8da312c774`. Removal used no force. The
pre-relocation checkout total was 32,152,076,288 apparent bytes. The
coordinator independently verified all 5,266 preserved entries by size,
SHA-256 and mode, plus branch tips, ancestry, removed paths and retained
external roots. The 2,028,289,001-byte regular-file sum includes hardlinks;
quarantine occupies 1,974,628,352 allocated bytes. These artifacts remain
preserved and are not claimed as reclaimed space.

Story-text remains because foreign source handles appeared during final
checks. ODraw remains because its tip is not an integration ancestor.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twenty-fourth-wave/`.
Its final receipt is 4,041 bytes, SHA-256
`bbbff01f5383a711426f67f47d89c8e6daabf8565859d119cdd57d0db45aff82`.
The 1,463,756-byte `manifest.json` remains outside Git, SHA-256
`8387b2ab354bd763d223ccf4b551a0daa8b92354bf0d3cf360d1f3c3f321b49f`.
The generating command was `python3` with the adjacent `move_w8.py`
(3,317 bytes, SHA-256
`4cb3f314c264daa3167cbc6535f9247d97199aafcc3c5c433954d885bedd7f1f`),
a wrapper around the retained cleanup22 mover. Regeneration requires
recreating the retained branches and restoring their ignored files first,
using the exact commands in the receipt. The adjacent `restore.py` is
2,066 bytes, SHA-256
`68c4d97a3148e3799e390b6036b72da7dd8426b5e16987caf039dbf0b23b33d7`.
The corrected wave label and receipt digest typo preserve earlier metadata
copies alongside the final files; no payload changed.


## Twenty-fifth cleanup, 2026-10-24

After review and merge at `2dfc4f134d`, the XLS property/feature leaf was
removed without force. Its retained branch tip is
`9d38e2c4f28e28be4532823e66a13de5637bb916`; the source checkout measured
7,551,512,576 apparent bytes before relocation. Three ignored files totaling
203,491 bytes were quarantined; their 225,280 allocated bytes are preserved,
not reclaimed. The external target and evidence root remains intact.

The coordinator found and returned a quarantine-layout mismatch before
acceptance: the payloads were directly under `ignored/`, while the manifest
and restore helper expected `ignored/xls-property/`. The mover owner corrected
the layout and the coordinator independently reverified every size/hash,
retained ref, ancestry, removed path and external-root existence.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twenty-fifth-wave/`.
The final receipt is 3,323 bytes, SHA-256
`6d628526eff3d42fb4ab457c8ca21c977cf2b7f26629a848a7ff5ca7b892ed9c`.
The 973-byte manifest has SHA-256
`5692ff8c58a16cbd0f712ed169bf2c21c51d2136301b5522a5af5d3b91dc9ad2`;
the 1,639-byte restore helper has SHA-256
`5a78eef2210c991dafcfed01ece0b829c195a261c8d296c0ade4643c64226004`.
Recreate the leaf from its retained branch, then run the adjacent `restore.py`.


## Twenty-sixth cleanup, 2026-10-24

The final IWA fixture-binding and W8 probe-binding leaves were removed without
force after fresh checks against frozen `b6732d02b2`. Their retained tips are
`4dae391375523d08672f2115a8011fb5628d2cfa` and
`d714aee19d052da58d732b0483b06b60d504f9a5`. Pre-relocation checkouts measured
7,552,303,104 and 8,093,696 apparent bytes. Seven ignored bytecode files were
hash-quarantined under the manifest-prescribed leaf prefixes; the coordinator
verified all payload sizes/hashes, refs, ancestry and removed source paths.
Their 454,656 allocated bytes remain preserved, not reclaimed. External
protocol logs, native endpoints and all targets remain untouched.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/twenty-sixth-wave/`.
The 2,177-byte receipt has SHA-256
`a2fe626527f62fcd40ec635369c4d515f0c5f38ae2c12155dbb26e5b4e84362a`.
The 2,038-byte manifest has SHA-256
`3a3946537cd75e25a0a53c0ccbcf5394a4b8a58c03b91aa57c5fc147f9968b15`;
the 1,639-byte restore helper has SHA-256
`f5b3ad7169441de455310d76374f0c8bce50b1d07299b6368fd16686a6f8050d`.
Recreate the leaves from their retained branches, then run the adjacent
`restore.py`.


## Twenty-seventh cleanup, 2026-10-24

The reviewed W4b/W5 historical-probe leaf was removed without force after
merge at `f6fe430a4d19d94aa360c0e9ed03ebd757265225`. Its retained branch
`uo2/w4b-baseline-compat` points to
`ccc536d5242cefcc6f58df200342c38faa5ec91e`. The first attempt correctly
left the leaf intact while two review searches held it; their owner stopped
only those searches before the final process checks and removal.

The sparse checkout measured 16,750,285 apparent bytes before removal. Four
ignored bytecode files totaling 185,161 bytes remain hash-quarantined. The
coordinator verified every payload's size, hash and mode, the retained ref,
merge ancestry, removed path and all four protected native source endpoints.
Native baseline/candidate endpoints, external targets and logs remain intact.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/`.
The successful `wave27-w4b-success.json` is 2,889 bytes, SHA-256
`1e69594e4eb3aa1a6a99ce06464a2cf15624f898cb9851a9d493dc92a30d7db2`.
The earlier 3,137-byte blocked receipt is retained, SHA-256
`6b13e5584001370a7bd4227bd1598d20a03062fb6635a7ebde0f7aeeb6551f51`.
Under `quarantine/wave27-w4b/`, the 1,210-byte manifest has SHA-256
`4b387d4d6b2d193b3e04c29e42f9994ff2d126bf42d5869f90d01e9f8a260bc1`;
the 1,489-byte restore helper has SHA-256
`cb07b3abea6858b661984627c73c01f8e3cf69289836308d95168b05c7542536`.
Recreate the leaf from its retained branch, then run that `restore.py`.


## Twenty-eighth cleanup, 2026-10-24

The reviewed inventory and encrypted-DOCX leaves were removed without force
after their merges into `dd78c02a23a202b6fa25562872a58113c4b20cb8`.
Their retained branches point to `668249cebf45f58ca6c460e36499561093cc4976`
and `696352d066ee6eb92cdcd0be941ca9102d041570`. The checkouts measured
7,552,352,256 and 7,552,401,408 apparent bytes before removal. Nine ignored
files totaling 322,218 bytes remain in quarantine; those preserved bytes
are not reclaimed. All external targets, locked validation logs and native
endpoints remain intact. No process held either leaf and no signals were sent.

The coordinator independently verified every payload's size, SHA-256 and
mode, both retained refs and merge ancestry, the removed paths, external
evidence identities and the restore helper's expected directory layout.
The 6,584-byte receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/wave28-cleanup-success.json`,
SHA-256 `30e7c38ce3ea3507843a64cbbf4fe28a85229efa432f8169df591d432b64d4d8`.
Under the adjacent `quarantine/wave28/`, the 2,416-byte manifest has SHA-256
`cfca857cf0bb33f1457c4ecb382c51836906d65a88c8d501501fb144f8d0b630`;
the 1,845-byte restore helper has SHA-256
`787069c930c1604a4aa7a9e426d02d1eca6ec213a1d65bab028f6e5c652b9c39`.
Recreate both leaves from their retained branches before running `restore.py`.


## Twenty-ninth cleanup, 2026-10-24

The reviewed short Unix socket fixture leaf was removed without force after
merge at `250769e1dcb692994a382829c238bf23bda5035d`. Its retained branch
`uo2/short-unix-socket-fixtures` points to
`35f47e34dce1c57695db71e3709c01ec29cc170c`. The checkout measured
7,552,471,040 apparent bytes before removal. Its 139,592-byte ignored lock
remains hash-quarantined. No process held the leaf; no signals were sent.
The DOCX cache and all validation evidence remain intact.

The coordinator independently verified the retained ref and ancestry,
removed path, payload size/hash/mode, external logs and restoration layout.
The 4,543-byte receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/wave29-cleanup-success.json`,
SHA-256 `323693fe374143b443f4fe15fb13df1d04b35f1e3e8939700a9711d563df6078`.
Under `quarantine/wave29/`, the 788-byte manifest has SHA-256
`682fc30452608a7d3d952437387c56f86ebc94c88c414068ef38e22470da8169`;
the 1,755-byte restore helper has SHA-256
`99ebea657b9386ccac7b02dccb0d3f240c08877aeea59fb644e8d7851a5f98e9`.
Recreate the leaf from its retained branch before running `restore.py`.
The receipt also retains a read-only target inspection; no build artifacts
were deleted by this cleanup.

## Build-cache headroom after cleanup twenty-nine

With no active holder of the coordinator target, 126 regenerable ELF
executables under `budget-validation/target/debug/examples` were removed.
Each was an exact SHA-256 duplicate of an executable already recorded as
pruned by a completed safe-chunked check. Before unlinking, the cleanup
rechecked each file's type, executable mode, size, hash, device, inode and
link count against its manifest and matching completed receipt. It removed
3,712,220,600 bytes. All three benchmark/profile executables and 5,610
non-executable files remained; the immediate and final holder scans were empty.

The external directory is
`/home/zhuhe/code/litchi-worktrees/.coord/target-cache-cleanup-2026-10-23/wave3-examples/`.
Its 3,605,760-byte `manifest.json` remains outside Git, SHA-256
`7a25d4cff6bd9a42f77daa62587e52bf9a7f78d958ecfa7a49219c46c2bb0138`.
The 45,319-byte `deletion-receipt.json` has SHA-256
`0088c05ac29ad9aefa361a5a7cc8b04668fecea64bf0faf346042dfe71452c02`.
The coordinator verified the receipt identities, candidate totals and
preservation/holder observations. Build commands in the retained gate logs
regenerate these duplicate outputs. No source, benchmark binary, native
endpoint or validation evidence was removed. The target's allocated size
after cleanup was 14,922,784,768 bytes; subsequent checks may grow it again.

## Thirtieth cleanup, 2026-10-24

The merged upstream-guard, evidence-catalog and corpus-report leaves were
removed without force. Their retained tips are `00b2831d067908aeb93211d1e82ff1380ed47efd`,
`1d2f6cfcad3a3b269a9ed4c639e088e61fd6a8c1` and
`2510bf5b40d2db17e774cf68734f7e08d91700ec`. No live holder was found and
no signal was sent. The cleanup preserved 14,437 ignored records, including
the corpus leaf's unexpected local target. Its 20,343,611,392 allocated
quarantine bytes were preserved at this checkpoint, not reclaimed. The
later audited cache prune below supersedes only that cache-retention claim.
That local target exceeds
the validation lane's 20-billion-byte ceiling; the corpus owner must rerun
the affected checks with explicit bounded external-target provenance.

Independent verification passed every payload's type, size, hash and mode,
all 696 directory modes, exact restoration layout, all 94 hardlink groups,
retained refs/ancestry and removed paths. The verification's final historical
external-target size comparison returned exit 1 because the approved corpus
rerun was already rebuilding that target. The receipt preserves this result
as superseded by authorized reuse; it does not claim that a live build cache
still equals its earlier snapshot. The cleanup's own before/after snapshots
remain historical evidence.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave30/`.
The 6,608,213-byte `manifest.json` remains outside Git, SHA-256
`bcbd9d94c404131cee9e4c60d680a43b25ad306593be4e7d0693f821efb821ba`.
The 4,344-byte cleanup receipt has SHA-256
`0c3bba4101354e978cf0b94e35c83e54528d99faea08c56c98fdaa3c8e0b664b`;
the 4,219-byte restore helper has SHA-256
`534829b9c03ee4892ae27d4b6f68d530c8d89bc4d2f310e27bfdf4cefe90aa84`.
The generating command was `python3` with the adjacent `prepare_cleanup.py`.
Regeneration requires recreating the three retained branches, restoring their
ignored files with `restore.py`, and directing the helper's `QUARANTINE`,
`MANIFEST` and `RECEIPT` paths to a fresh external directory. Do not reuse the
archived oversized target for validation.

The 5,063-byte independent receipt is
`.coord/worktree-cleanup-2026-10-23/wave30-independent-verification.json`,
SHA-256 `b504842a0b015fbc9fc8e8e5360adcd258258617b1a3ca6c4a58af7048128a0d`.
It binds the durable verification log, original exit 1 and active worker
observation, alongside the completed payload verification.

## Thirty-first cleanup, 2026-10-24

The merged W8 policy and XLS catalog leaves were removed without force.
Their retained tips are `e12b57d6246ee825fc2879b8b71bc7b57546204d` and
`3ddd72e97e46219d9e75c4a8b5cf7925ac86eeeb`. Eight ignored records totaling
451,585 bytes remain quarantined. The coordinator verified every payload's
size/hash/mode, directory layout, retained ref, ancestry and removed path.
All external native endpoints, XLS targets and logs remain intact. No live
holder was found and no signal was sent; the active corpus rerun was untouched.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave31/`.
The 9,601-byte manifest has SHA-256
`7b39a5f45db462a0a73b59b8b9b3b0ca956a1678f715518aaed0aa8190bf4269`;
the 4,269-byte receipt has SHA-256
`4c066c0ae8ce3ced70a8ee168be4bdb455c6e08fe17a66b8924cf0a955ae15f5`;
the 3,978-byte restore helper has SHA-256
`9fd26f2ba511a08d61666459616b24641f3bd2dafada88c45b5882e6de977a7f`.
Recreate the two retained branch worktrees before running `restore.py`.


## Thirty-second cleanup, 2026-10-24

The merged gate-registration leaf was removed without force. Its retained
branch tip is `8a795fb7692198fe1ae50e44e4bcafa461479456`, merged at
`8220f809b928fac57ec6048f28079cc45c421adf`. It had no ignored files. The
coordinator independently verified the retained ref, ancestry, absent path
and empty payload container. The active and prepared gate runners remain
intact; no holder was found or signal sent.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave32/`.
The 3,208-byte manifest has SHA-256
`dd48c9878503bb108a17e002a1444459e0ed7f3fca6b0f76b6b82cc577a50eea`;
the 3,185-byte receipt has SHA-256
`5c301768005e32314ca47717f0694da656544c0488e59dd6dbfc5d27f1c69288`;
the 3,961-byte restore helper has SHA-256
`ec17d03d7e052188affc75ba6b60dc71aacd7617deb49ee0d5bdcb9d3fa30f8b`.

## Thirty-third cleanup, 2026-10-24

The reviewed boundary and corrected corpus-receipt leaves were removed
without force after their merges. Their retained tips are
`5a8cfa93c13352808baa825742a17e50f68696a4` and
`36cf4e3dadfb8e29f78df50d9d07bd00ede10c15`, merged at `b7968d4814` and
`885cdc82af`. Three ignored records totaling 6,309,808 bytes remain
quarantined: two Python caches and the corpus lock. The coordinator verified
every payload's size, hash, mode and directory layout, retained refs,
ancestry and removed paths. External corpus caches and receipts, both gate
runners, the earlier oversized-target quarantine and active OPC/ODP repair
leaves remain intact. No live holder was found and no signal was sent.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave33/`.
The 12,211-byte manifest has SHA-256
`e9dfafa2dad22e01e421ee46a5a5418b86f765928c2fe036af5fe6592f2bf63b`;
the 18,748-byte receipt has SHA-256
`2f2183e0cce371e47ab124d17bd327fd292d19371b6c46252b40f20861c90e73`;
the 3,961-byte restore helper has SHA-256
`19983979ab7d1780fd09aee8ed4ebc91e2d794bca89d0c312eef76c8986c970a`.
Recreate the two retained branch worktrees before running `restore.py`.


## Thirty-fourth cleanup, 2026-10-24

The merged ODP doctest and OPC retention leaves were removed without force.
Their retained tips are `21c47647afec3f8c3278f427f8eb2d68beb83735` and
`42e3f2314c944cb3e5aca00725bd732c4cb537fb`, merged at `26ab39411e` and
`49c25ef67f`. Both ignored lockfiles remain quarantined: 139,591 bytes each,
SHA-256 `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The coordinator verified both payloads, modes, refs, ancestry and removed
paths. External targets, raw validation logs and native artifacts remain
intact. No source holder was found and no signal was sent.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave34/`.
The 13,435-byte manifest has SHA-256
`1491b4cba410e2dcecff97cb809e26bfcb7c38f0c05de7c88bad1d0ec314a61a`.
The final 21,562-byte `receipt.json` has SHA-256
`d83420ee926d36df16f81980db5bf88e85127de97b4361ecde2a138c9d6fd50c`.
Per-leaf receipts are `final-odp-doctests-receipt.json` (21,042 bytes,
`13f1c98d234ac1b991041ca94ef0d516cd729e05911dc7e87a050d3279aaffe8`)
and `final-opc-retention-receipt.json` (21,045 bytes,
`4c637b32091a4a190314c6a6acd744e8647841d69d8fdb5e5dc712fd9ad8e86f`).
The 4,007-byte restore helper has SHA-256
`c591173d61b7543a7d7490d60292c39a5a76932b18ffacb08c3565367325f8a6`.
Recreate the retained branch worktrees before running `restore.py`.


## Regenerable corpus cache after cleanup thirty-four

The rejected corpus target was audited before pruning. Exactly 13,893
regenerable Cargo-cache and log-identified test-executable paths were removed;
unknown executables, native outputs, generated protocol/source files, locks
and all raw receipts remain. The 20,731,702,487 summed path bytes represented
13,799 distinct inodes and 20,158,658,579 unique file bytes. Actual allocated
storage reclaimed was 20,181,409,792 bytes (39,416,816 blocks of 512 bytes).
No holder was found, and every identity/hash was checked before removal.

Independent verification confirms every authorized removal, all 536 retained
corpus files and all 692 remaining directory modes/layouts, plus the eight
other wave-thirty leaf payloads. The 173 directories in the initial filtered
restore manifest are the retained files' required parents; it does not claim
to restore discarded cache contents. The complete replacement restore route
covers all three leaves: 544 retained files and 177 required directories.
All-leaf and individual-leaf dry-runs pass. Original manifests, receipts and
helpers remain immutable historical evidence.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/target-cache-cleanup-2026-10-23/wave4-corpus-cache-prune-20261024/`.
The 6,416,031-byte `deletion-receipt.json` remains outside Git, SHA-256
`db262ebefc292b9ccacc3bfaf7e7950b0f47902ca8983abbd348e4bd13d07162`.
It was generated by `python3 prune_cache.py`, followed by
`python3 finalize_prune.py`: the first command completed the authorized
unlinks but its postcheck incorrectly included the removed target files in
the protected set. The finalizer rechecked the exact remaining set and every
preserved hash/mode with the correct exclusion, without another unlink.
The 5,293-byte independent receipt has SHA-256
`cdb7fb8df0f9f25c41ad77c72c2cf1b81547192e4c698ae18b92a92074038daf`.

The definitive complete restore manifest is
`wave30-post-prune-preserved-manifest.json` (347,178 bytes,
`9b9afa820a1ea367f724b5d07dea3997c1542786e494da3412bdbbe6d4f55758`).
Use `python3 wave30-post-prune-restore.py --dry-run` to verify all retained
payloads; its `--leaf` selection restores one recreated source worktree.
The 7,118-byte helper has SHA-256
`2c1c0ab83db471610c0b44bb321d167fd0ca08791cafc343286c4ed95c4758ba`.
Its independent 2,671-byte verification receipt has SHA-256
`c7df77c9624287030cdae10d43d366cea64fe6b5bdc0aca5e6d3b231d941c026`.
The adjacent README records the functional rebuild command with the retained
source/lock, 20-GiB memory cap, one job, disabled incremental/debug output and
offline locked resolution. Rebuilding does not promise byte-identical old
cache output; the original rejected run did not record all build settings.

## Additional gate-target headroom

A separate exact-match cleanup removes 144 stale example executables from
the inactive budget target, totaling 1,554,094,416 bytes. Each was a regular
ELF file with one link and an exact name/size/hash match to an example
already pruned by a successful recorded Cargo command. Failed chunks never
supply deletion evidence. All 13 benchmark/profile twins and 6,228 metadata
files remain; the coordinator verified their hashes and modes independently.
The budget target falls from 19,005,423,149 to 17,451,328,733 bytes. The
active documentation job uses a separate target and is untouched.

External directory:
`/home/zhuhe/code/litchi-worktrees/.coord/target-cache-cleanup-2026-10-24/wave4-examples-sibling-plan/`.
The immutable 695,485-byte plan has SHA-256
`97f493ec904350761ce1c87ddf81e0d5d7ea7df164cbaca1256885ea3440c9db`.
The 3,827,904-byte `deletion-receipt-completed.json` has SHA-256
`0d5f96928ccdb1e2afeb717291b8154d6114e039b898d0ce5acaaea7c09db6f2`;
the 4,043,596-byte `preservation-inventory-completed.json` has SHA-256
`5aa7233848f58514ed79f424ae5a31d9326ef07c331c033fd7c486d80363d7ed`.
Both remain outside Git. The initial preflight stopped without deletion on
an inherited environment-only target value; exact argv/open-handle checks
confirmed that the live documentation child used only its isolated target.
Every later per-unlink holder scan was clear.

The original commands ran inline. Equivalent retained scripts and the
reproduction preconditions are under `reproduction/`; run
`python3 reproduction/verify_plan_and_receipt.py` for read-only verification
and `python3 reproduction/delete_sibling_candidates.py --dry-run` for a
non-mutating replay check. Recreating the candidates requires the successful
build commands recorded there, followed by a fresh identity audit: new
inodes and event times will differ. All source receipts bind the binaries to
`f6fe430a4d19d94aa360c0e9ed03ebd757265225`; `49c25ef67f` is only the later
checkout at deletion time. The corrected reproduction README is 4,277 bytes,
SHA-256 `d138191537a51d4349e382fba5aeea50bccc40a359efece46b7e769a2afcdad7`.

## Thirty-fifth cleanup: reviewed OPC protocol leaf

The coordinator merged the reviewed OPC read-ahead protocol tip
`80983c725d9b803f066545e05f8463a25de4120c` at `945eb1d6b9`, then removed
only `uo2-final-opc-read-ahead-perf` with non-forced worktree removal.
There were zero ignored or untracked payloads. The branch remains at its
exact tip; ancestry, the absent worktree and the zero-payload manifest verify.
The pending W8 leaf, frozen source and external targets were untouched.

Evidence remains under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave35/`:

- `manifest.json`: 11,174 bytes, SHA-256
  `8d3b95bdcf8801bad2f4c55c883d389e5ea777cae31f2dec8b7957361cb5ec98`.
- `receipt.json`: 20,058 bytes, SHA-256
  `5ee533109c90daf7eb5ec399c6af9b0ed4e7112870a73a90448013503d558ca5`.
- `restore.py`: 4,865 bytes, SHA-256
  `8e0a25698e0425d7d89f849ae9ba107d4ba687b5e7607059059fee62a4e6494e`.

This helper requires the worktree to be recreated before invocation; it has
no `--dry-run` option. No payload restoration is needed for this empty leaf.
No cache artifacts or measurements were deleted by this cleanup.

## Thirty-sixth cleanup: reviewed inventory and W8 leaves

After merges `54398757b8` and `fae96321e3`, the coordinator removed
`uo2-w8-final-rebind` at `e814e4512f7414d3e3136699b8a98e37e09710a5`
and `uo2-inventory-delegated-globs` at
`2099c66520cace074d788e1e3b5ac5b12ade1587`. Both branch refs remain.
The 15 ignored files, totaling 598,662 bytes, were quarantined; the
coordinator independently verified every hash and mode. Removal was
non-forced. The first preflight stopped on a lingering W8 audit pipeline;
its owner identified and terminated only that owned process group before
cleanup rechecked the exact leaf paths. Active repair leaves and targets
were untouched.

Evidence directory:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave36/`.
The 20,692-byte manifest has SHA-256
`34e8f18483af99f0319f4ee3c210035ce6c244ff758ed74fab6836682f60a2c0`;
the 21,242-byte receipt has SHA-256
`7be76f199d9c4488966445a455ed9102847c718397b9defeccf9ceb11834354e`.
The 4,849-byte restore helper has SHA-256
`92737064528012e0d8404846876fbf149f5590f3d2da1e00a86d773b5c1c6fe9`.
Its CLI is `python3 restore.py [leaf]`, after recreating the selected
worktree; it has no `--dry-run` option.

## Thirty-seventh cleanup: Markdown, IWA aliases and font/cache limits

At staging `e35cede32ae730d351ce9d29d243c35edd321ddf`, three reviewed
leaf tips were already merged and their tracked worktrees were clean:

| Removed worktree | Retained branch tip | Retained files | Bytes |
| --- | --- | ---: | ---: |
| `uo2-final-markdown-reader` | `533902609a6c974c387975775241c0133f1f0817` | 3 | 205686 |
| `uo2-iwa-protos-alias-repair` | `122aa7cc33276e3bc039e2ed665948cb3bcd27ac` | 4 | 6332669 |
| `uo2-font-cache-limits` | `a03d518eca5e0b8bfa7993be614d58aa8c03514b` | 1 | 139591 |

All eight ignored payloads (6,677,946 bytes) were quarantined before
non-forced removal. The coordinator independently verified each payload's
size, hash and mode, all branch tips and merge ancestry, and the removed
paths. The cleanup worker exercised the restore helper in three temporary
sparse worktrees, verified the restored files, and removed those temporary
worktrees non-forced. Exact-path holder checks were clear after the owners
terminated their obsolete audit pipelines; cleanup sent no signals.

Evidence remains outside Git under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave37/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 29811 | `d32dfbef509e14dc9180451eb4319507ad4a3a87737cf74f1d956d70c16f0c4a` |
| `receipt.json` | 44388 | `f23d47af5d6290d46b8e54060624b6304fd1fa142008c0f10b0a78aaa77c1083` |
| `prepare_cleanup.py` | 22347 | `87f2034ee26e27a50bd052881411a4334327b686ebc848fda7fa58e29e4d97b6` |
| `restore.py` | 4828 | `31b1c136ceb87ef5f01f5d0042f9b02b65fb216fd5b0bd8e6092aee4649d6b04` |

The restore helper takes an optional leaf name after its worktree has been
recreated. External targets, receipts, frozen source, active repair leaves
and the trunk's untracked constructor probe were preserved.

## Thirty-eighth cleanup: five reviewed API repair leaves

At staging `2b837312365d3b8d2d9e16d6f02e00822e734712`, five clean,
merged leaves were removed non-forced. Their branch refs remain at the
reviewed tips:

| Removed worktree | Retained branch tip | Retained files | Bytes |
| --- | --- | ---: | ---: |
| `uo2-sign-limits` | `4e2d49dd4339204c169c63cd44251dcdf58f49f5` | 4112 | 2945482862 |
| `uo2-exhaustive-errors` | `c4740038e6d15c60c2d457135662c6bee170bd12` | 1 | 139592 |
| `uo2-api-result-contract-repair` | `435e1547bd8fef14996e938f7e926fb69f63c56a` | 6 | 231998 |
| `uo2-inventory-semantic-values` | `817134864abc3d1db0f488357abc0cc77594aa83` | 2 | 35724 |
| `uo2-spreadsheet-chart-builder` | `341164e0bf38f5d5d90a8ad2dc46ea62a9211ba2` | 6 | 327708 |

All 4,127 ignored payloads (2,946,217,884 bytes) remain in quarantine,
including the signature leaf's build target. This removes redundant source
worktrees; it does not claim those retained bytes as reclaimed disk space.
The coordinator independently verifies every payload's size, mode and
SHA-256, branch tips, ancestry at removal, and the absence of the old paths.
The cleanup worker exercises restoration in five temporary sparse worktrees
and removes those copies non-forced. An obsolete audit process group was
identified and stopped by its owner before removal; cleanup sends no signals.

Evidence remains outside Git under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave38/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 1870340 | `ef17f5ad6dfca16ca16855873638c08f3604577936279cb01d73681e25544ea8` |
| `receipt.json` | 59382 | `f544befb1715094ad1776e809bcb12a7b0ae4027b83c0032256a83d25ec53adb` |
| `prepare_cleanup.py` | 6268 | `bd6e312c39d96516b12f4b8ed82fc5c1e16334518b24ab8ea0168c0f01ef6013` |
| `restore.py` | 4828 | `1b1c91898119bd4b05e39cf1e7c0356d2f430d9398f02f21066bf302a3cc7a31` |

The manifest exceeds 1 MiB and is deliberately external. Recreate a selected
worktree at its retained branch, then run `python3 restore.py <leaf>` from
this evidence directory to restore its ignored payloads. The retained
`prepare_cleanup.py` records how to regenerate the manifest from restored
leaves; a new cleanup requires a fresh holder and identity audit. Active
CFB, DrawingML, ZIP and legacy raw leaves, external targets, frozen source
and the trunk's constructor probe are preserved.

## Thirty-ninth cleanup: raw legacy, ZIP and exact source roles

At staging `2021a41c4cb217844131892138a7d0452a81d6c1`, three clean merged
leaves were removed non-forced; their branch refs remain:

| Removed worktree | Retained branch tip | Retained files | Bytes |
| --- | --- | ---: | ---: |
| `uo2-legacy-raw-boundary` | `6a55d251d29607c8c2020c9046821dc66ee3c395` | 1 | 139606 |
| `uo2-zip-explicit-limits` | `0e3cfc0e1cb0efe8c371623e42c347aa90640461` | 1 | 139591 |
| `uo2-final-contract-inventory-readiness` | `22aed11db8db678cbcac68725ec52a9f0083f884` | 9 | 287389 |

All eleven payloads (566,586 apparent bytes; 618,496 allocated bytes) remain
in quarantine. Restoration is exercised in temporary sparse worktrees with
hash and mode checks. The coordinator independently verifies every payload,
branch tip and merge ancestry, removed paths and the protected 15-byte trunk
constructor probe. Cleanup sends no signals and changes no external target,
active source or native endpoint. The concurrent DrawingML directory mtime
change is recorded separately from cleanup. Exact physical disk reclamation
was not measured; retained payload bytes are not described as reclaimed.

Evidence remains outside Git under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave39/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 62625 | `a27af2789018991e7fc2a7eb9b2234a76c2c772ebf31c40b94f0d07c757439d0` |
| `receipt.json` | 109941 | `54becaf7a51d26c6b7202856c620ff8ae51769d4d1a07de71ced862e948297d9` |
| `prepare_cleanup.py` | 6928 | `f7b529737ad45c387f8419d4fb6328cc1033751e0c9cd4ee0ab38ca4b9e3c1e1` |
| `restore.py` | 4828 | `d7077650b69a82b252c044489872a63889d4a862a31cfd6c16aae8c40f796758` |

The independent coordinator receipt is `root-verification.json`, SHA-256
`78a792acf6c99df2cae5aa4ba823add8ae381db4adc4969ea65a1a21a5acdd8e`.
Recreate a selected worktree at its retained branch, then run
`python3 restore.py <leaf>` in this evidence directory. A later cleanup must
repeat holder and identity checks.

## Fortieth and forty-first cleanup: XLSX protocol and CFB

Two clean merged leaves were removed non-forced. Their branch refs remain:

| Wave | Removed worktree | Retained branch tip | Files | Bytes |
| --- | --- | --- | ---: | ---: |
| 40 | `uo2-xlsx-raw-native-protocol` | `5e6555d8406709382e26ecbcdd6cb18ded449e31` | 8 | 277300 |
| 41 | `uo2-cfb-explicit-limits` | `bf7f48741bc02d1f5e736a7a732c25808d28af36` | 6 | 327678 |

All fourteen ignored payloads (604,978 apparent bytes; 679,936 allocated
bytes) remain in quarantine. Restoration verifies every hash and mode.
The coordinator independently verifies payload bytes, modes, retained branch
tips, merge ancestry, removed paths and the protected trunk constructor probe.
External targets and active source endpoints remain untouched. Source
pre-removal allocated sizes are accounting only; exact physical reclamation
was not measured and retained payloads are not claimed as reclaimed.

Wave forty removed the XLSX leaf and then recorded an immutable engine
`aborted` status because staging advanced concurrently from `dcc7cfbdef` to
`4624f4bea6` before its final HEAD check. The separate verification preserves
that failure and confirms the old HEAD and leaf remain ancestors, the path
is absent, and every retained payload and reference is correct. It does not
rewrite the original receipt. The CFB leaf was initially held by an obsolete
audit pipeline; its author identified and terminated only that owned process
group. Fresh holder checks were empty before wave forty-one, which completed
with staging held at `4624f4bea6`. Cleanup itself sent no signals.

Evidence directories are outside Git under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `wave40/manifest.json` | 61161 | `8b0dac0f87a8275cbc20d913466775260716f76bf94be47971ebda35ed547509` |
| `wave40/receipt.json` | 56738 | `dab8ea11a0084dcb06250e61c53af0c9655e613b2018480b4cddd5e31423cce5` |
| `wave40/prepare_cleanup.py` | 6465 | `afd51d155bca8c6065f9e30242d6534c36d9585789f26596110d51a23fd16a9e` |
| `wave40/restore.py` | 4804 | `59db544ba29628c0f0bdc1f05a7373aa50c8d3672b92ebd5ea9f06834e61f9a0` |
| `wave40/coordinator-verification.json` | 2660 | `0bc7a661a7f4187ca25d5201b17fff15601d41ad3d366ed4cfdb49733194a837` |
| `wave40/root-verification.json` | 715 | `6b19841e129bf0cac384b154c6684da2a068a159150a10c4433a7d49057a66a8` |
| `wave41/manifest.json` | 60017 | `fb70656c83192506ba267ab204df059133cd8d0b68b846ca9d47c771c5180dad` |
| `wave41/receipt.json` | 112375 | `6bc232904e703055f9cd6445cb7fd07823ba49eb4430b693f7e25e120d3c2837` |
| `wave41/prepare_cleanup.py` | 6395 | `c811ff3c732566c953e7435aed8d3669e52c99b1d65997697d58478612bc4fde` |
| `wave41/restore.py` | 4804 | `55b4cec2d79a818ba81b60b70d332f4dafbdb93c67dfca195ed37d1bcf5ea8d7` |
| `wave41/coordinator-verification.json` | 2652 | `9fb2aa41b05ee512c7bc1f7896b0d3096a8bc35d883a9aabdf42305f7d206fb7` |
| `wave41/root-verification.json` | 715 | `0b7c0a49db089cec9667ee29f220366803f770edda981dd0d1cd2aa77570284a` |

Recreate the selected worktree at its retained branch, then run
`python3 restore.py <leaf>` in its evidence directory. The recorded preparation
helper regenerates a manifest from restored leaves; any later removal must
repeat holder and identity checks.

## Forty-second cleanup: merged legacy native protocol

At staging `778d7a5e6cc1b28838c7462049cfdc5a5288debe`, the clean merged
`uo2-legacy-native-protocol` worktree was removed non-forced. Its branch
remains at `e8d51bb07a00dd0b9feeb6ed5e134c521c89be4d`. All nine ignored
payloads (356,146 apparent bytes; 397,312 allocated bytes) remain in quarantine.
Restoration passed every hash/mode check. The coordinator independently
verified every payload, retained ref, merge ancestry, absent leaf path and
unchanged protected trunk probe. Physical reclaimed bytes were not measured.

A fresh preflight first stopped on an obsolete Git audit process group. Its
author confirmed ownership and terminated only that group before holder
checks were repeated. Cleanup sent no signals. The endpoint audit found no
active native preparation or target depending on this leaf. All 54 protected
external roots, 121 untouched worktrees and prior cleanup receipts remain
unchanged. Staging HEAD was held throughout this cleanup.

Evidence remains outside Git under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave42/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 63078 | `11849a10de6f8cbf5232526d012d926ce83505cb11b31f54ea69e38e425436f2` |
| `receipt.json` | 114821 | `a0f3d6214d515ddcaff89cc69d2dd915030d9264e85b46e39ced77fba3ff7c34` |
| `prepare_cleanup.py` | 6615 | `d2a86c85984de528dbebc3c95a6eb9c19b4595a91242d5f11d1102187f31efe4` |
| `restore.py` | 4804 | `e04fed70c54421538cf379e78ee1f62ebbf73d7319e140be40d328e71852fa75` |
| `endpoint-audit.json` | 1630 | `290fbd12c986a1ea387fe672c368ea36c56f732651674583432135ca7bd2a87b` |
| `coordinator-verification.json` | 3069 | `0169c4778a64acccc6c626e7df09939210923d3de884b55c28ca1a0361314cbc` |
| `root-verification.json` | 715 | `c42bfb94f53baa15d75b7463040a2c2fff5b2565767abca6323a91a30acd606c` |

Recreate the leaf at its retained branch, then run
`python3 restore.py legacy-native-protocol` in this evidence directory.
Any further removal requires a fresh holder and endpoint audit.
