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
