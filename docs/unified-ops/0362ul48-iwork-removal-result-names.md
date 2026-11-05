# 0362ul48 — iWork removal-result names

Date: 2026-10-25. This source packet removes the format prefixes from the 17
public iWork removal-result types and migrates their declarations, constructors,
return types, re-exports, tests and documentation together. The source commit
is `375f55e8dfcea343836179166ca059da93e747ee` on
`uo2/iwa-removal-result-names-0362ul48`; the source remained clean throughout
validation. The change is limited to the 17 symbol renames plus re-export line
wrapping. There are no compatibility aliases or behavior changes.

The canonical names are `RemovedTextBox`, `RemovedSlideChart`,
`RemovedSlideImage`, `RemovedSlideShape`, `RemovedSlideTable`,
`RemovedSheetAudio`, `RemovedSheetChart`, `RemovedSheetImage`,
`RemovedSheetMovie`, `RemovedSheetShape`, `RemovedAudio`, `RemovedBodyShape`,
`RemovedBodyChart`, `RemovedImage` and `RemovedMovie` in their respective
Keynote, Numbers and Pages modules. The glossary contains exact bans for all
17 old names and the corresponding canonical exports.

## Pinned validation

The runner was
`/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py`, SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.
The exact root lock is 139,606 bytes, SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.
The runner also observed the unchanged auxiliary lock hashes
`51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`
(`tools/native-resave/Cargo.lock`) and
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`
(`tools/perf-baseline/Cargo.lock`).

The tracked root configuration was materialized before the corrected run and
verified byte-for-byte against `HEAD`:

```text
rust-toolchain.toml  e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310
clippy.toml          f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c
rustfmt.toml         5b08a5db07e6acaee62319a09314b9302c8cbc0f8e6b88b5ee8695f6303383d3
.cargo/config.toml   50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d
```

`rustup show active-toolchain` selected
`1.95.0-x86_64-unknown-linux-gnu`, overridden by the tracked
`rust-toolchain.toml`. The corrected Cargo was `cargo 1.95.0
(f2d3ce0bd 2026-03-21)`, rustc was `rustc 1.95.0 (59807616e 2026-04-14)`,
and clippy was `clippy 0.1.95 (59807616e1 2026-04-14)`. A live compiler process
was independently observed at
`/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/rustc`.

The corrected all-feature/all-target clippy rows used `isolated=False` so they
could run alongside the recorded ordinary builders, while retaining the
20-GB target guard, one Cargo job and offline/locked mode:

```text
cargo clippy --locked -q -p litchi-iwa --all-features --all-targets -- -D warnings
receipt: /home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025j/state/104-iwa-clippy-all-targets-pinned-attempt-1792888372194569321-341904.log
receipt SHA-256: e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855

cargo clippy --locked -q -p litchi-ops -p litchi --all-features --all-targets -- -D warnings
receipt: /home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025k/state/105-iwa-consumer-clippy-all-targets-pinned-attempt-1792888850858631436-584060.log
receipt SHA-256: e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855
```

Both corrected rows passed. The iWork row reached a 369,331,620-byte target;
the consumer row reached 1,026,071,993 bytes. The source commit and all three
recorded Cargo lock hashes were unchanged after each row.

The selected `run_step` fields captured in the terminal are persisted at
`docs/unified-ops/results/0362ul48/pinned-clippy-receipts.json`. That file
records the exact external log and Cargo-target TSV paths, exit/status values,
source and lock identities, target guard measurements and provenance. It does
not reconstruct the unavailable full `preflight` member or claim that a
standalone `state.json` exists.

The source/API checks also passed:

```text
python3 -m unittest tools.test_check_iwork_leaf_public_api
  16 tests passed
python3 -m unittest tools.test_check_api_glossary
  54 tests passed
```

The consumer check passed in the earlier hydrated run and is retained as a
separate receipt. The iWork check was also successful after fixture hydration.
Both check rows predate reintroducing the tracked toolchain file, so they are
retained setup-complete diagnostics rather than the final pinned Rust evidence;
the corrected clippy rows above are the final pinned all-target compilation
evidence:

```text
cargo check --locked -p litchi-ops -p litchi --all-features --all-targets
receipt: /home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025i/state/102-iwa-consumer-check-all-targets-attempt-1792888219123124194-278223.log
receipt SHA-256: 883fa55f4eb83c790ea52708e82f2e97b105343e402a6a17443f5aa6b139b6b0

cargo check --locked -p litchi-iwa --all-features --all-targets
receipt: /home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025f/state/100-iwa-check-all-targets-attempt-1792887473862444114-4151089.log
receipt SHA-256: f927d0a1979617ea1f3b22119d4450989c93fc5684923d9444a90f2de46e743a
```

## Preserved setup and wrong-toolchain receipts

The first attempts deliberately remain outside Git and are not relabeled as
source failures. Before the exact workspace target closure was hydrated,
Cargo metadata stopped on missing tracked workspace targets. The preserved
receipts are:

```text
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025/state/100-iwa-check-all-targets-attempt-1792886935334652141-3881408.log
  SHA-256 12e0d5c45a4d6af791a4e9e0cc07ef56e0fbee8a7815ddf151b88894709036c
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025b/state/100-iwa-check-all-targets-attempt-1792887019394640529-3926057.log
  SHA-256 f09dfcbac29ce724afff0176cd4b0d2d239b6fc3742da82dbaf73f440785b379
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025c/state/100-iwa-check-all-targets-attempt-1792887144681610240-3972200.log
  SHA-256 c851c22a7b9f2b7c030a544236d0ba9a070a1d84463c43b894e6f886d3f10090
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025d/state/100-iwa-check-all-targets-attempt-1792887259851451371-4028014.log
  SHA-256 4af59655918b27fbcf35db3b7b0c7fb4f60cf2e8214cae171a82a6592511078d
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025e/state/100-iwa-check-all-targets-attempt-1792887366779185930-4093141.log
  SHA-256 ef2e0cd1db724fba8bc9839b7d9b606341c90e52d8b957b57d8cb3b17c0d9720
```

After the fixture closure was hydrated, the initial stable-toolchain clippy
rows were run before `rust-toolchain.toml` was materialized. They are preserved
as wrong-toolchain diagnostics, not accepted project debt:

```text
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025f/state/101-iwa-clippy-all-targets-attempt-1792887594705473653-21093.log
  SHA-256 126f3ebd2e9c524ab40d98ba4e396b39f4fb22353ea757bf624ee06fb413f9e3
  diagnostic: stable clippy's chunks_exact_to_as_chunks deny in litchi-iwa-protos
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025i/state/103-iwa-consumer-clippy-all-targets-attempt-1792888277972113893-307087.log
  SHA-256 85d3a4295b6bb3205daf40950bf06cd7f28ac8d3dcca3ad35b21aa88a2463790
  diagnostic: stable clippy's chunks_exact_to_as_chunks denies in litchi-ole-common
```

There was also one pinned-toolchain consumer attempt before the sparse pattern
was repaired. It failed at test compilation because
`test-data/iwork/pages/body-table-catalog-native.pages` was absent from the
sparse checkout; it was not a source diagnostic. Its structured selected
receipt is the second row in
`results/0362ul48/pinned-clippy-receipts.json`, with log SHA-256
`340f94a4d3a361cd317d7825812b8f3862fb8580a48b8b79c6e31fe73a2f36ef`.

The failed sparse consumer-check receipts remain preserved as well; they were
missing-fixture setup failures, not rename diagnostics:

```text
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025/state/102-iwa-consumer-check-all-targets-attempt-1792887656195548536-46750.log
  SHA-256 561330d7ebebb3ea78243ab965fa1e5695c740856b10763fb3fa34fb9a1cad56
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025g/state/102-iwa-consumer-check-all-targets-attempt-1792887916525820540-144992.log
  SHA-256 2d05689ce3278df627be5641cceff966e7f01f3f1c74cb698f4ffb4c3ffd011d
/home/zhuhe/.cache/litchi-api-contract/iwa-names-375f55e8df-validation-20261025h/state/102-iwa-consumer-check-all-targets-attempt-1792888184172891374-261151.log
  SHA-256 b2b374f8d0758243316897d6cf7073d71c5a25bd06cbc2f3ef65c436df0c7319
```

All required fixtures were then represented in the sparse specification and
the final pinned clippy rows passed. No native run, ABBA measurement, or
rustdoc run was performed in this packet; no PROGRESS change is included.

## Coordinator integration review

The final source at `0524a814aee7ae24a6f07a879ecd69dd0285e641` differs
from the pinned, validated `375f55e8dfcea343836179166ca059da93e747ee`
only in this record and the selected receipt JSON. Independent review verifies
the seventeen canonical result-name migrations are identifier-only changes
with re-export formatting, preserving behavior. All three selected log and
Cargo-target trace digests, source identities, unchanged lock identities, and
the runner digest match the retained external artifacts. The missing-fixture
failure remains recorded; the later pinned consumer Clippy invocation passes.
The incomplete preflight capture is explicitly disclosed and is not recreated.

The initial composed check exposed a matrix-coverage defect: its record-name
pattern does not recognize numbered follow-ups such as `0362ul48`. Adding a
classification therefore produced a false stale-classification failure (one
of 82 tests). That attempted classification is removed and a separate tooling
repair must recognize and audit every affected directory. This selected receipt
contains validation evidence only. The naming merge does not claim the fresh
integrated gate or final native performance acceptance.
