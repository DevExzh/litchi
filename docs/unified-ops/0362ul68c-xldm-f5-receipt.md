# 0362ul68c — XLDM scoped F5 receipt

Status: scoped F5 validation passed under the approved runner. The run was
offline and locked, with one build job, one test thread, no incremental or
debug artifacts, an 18,000,000,000-byte preflight target guard, a
19,000,000,000-byte live target guard, and a 20,000,000,000-byte hard target
guard. No native qualification,
rustdoc, broad workspace gate, benchmark, or performance run was performed.

Runner:

`/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py`

Runner SHA-256:

`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`

The tracked checkout was fully hydrated before the run (460,846 tracked
paths, zero missing paths). The source lockfile hashes were byte-exact:

* `Cargo.lock` — `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`
* `tools/perf-baseline/Cargo.lock` — `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`
* `tools/native-resave/Cargo.lock` — `51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`

The complete XLDM/XLSX portion passed as follows:

* source `87e241633e6ab9a823080f0c1276442f31bae027`: XLDM all-target tests,
  XLDM clippy, and XLSX all-target tests — PASS;
* source `c5424e23e48ba4e0ced4134de22ab56162555c71`: XLSB all-target tests,
  the affected XLSX/XLSB/ooxml-common check, spreadsheet clippy, feature-
  scoped `litchi-ops` check, feature-scoped `litchi` check, and both matching
  clippy runs — PASS.

The earlier accepted rows remain source-equivalent through the final source.
`git diff --name-status 87e241633e c5424e23e4` contains only
`crates/litchi-xlsb/src/data_model/proof.rs` and
`crates/litchi-xlsb/src/data_model/transaction.rs`; the XLDM and XLSX trees
are identical (`git diff --quiet 87e241633e c5424e23e4 --
crates/litchi-xldm crates/litchi-xlsx`). Thus the XLDM tests, XLDM clippy, and
XLSX all-target tests accepted at `87e241633e` remain valid for `c5424e23e4`.

Accepted artifact manifests (the row logs are also listed so every accepted
row has an exact path, size, and SHA-256):

| accepted rows | manifest.json (bytes; SHA-256) | rows.json (bytes; SHA-256) |
| --- | --- | --- |
| `xldm-tests`, `xldm-clippy`, `xlsx-tests` | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-87e241633e/manifest.json` (1,891; `3ddc0312c1f8a30e42a7f78a6e04858bec8585a682266a734391c27cf768b7d3`) | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-87e241633e/rows.json` (34,570; `579380ac78b983d526ad14331852e51ca0115d03583a2ea04370a24e43d526e6`) |
| `xlsb-tests`, `xlsx-xlsb-affected-check`, `spreadsheet-clippy`, `ops-xlsx-xlsb-check`, `facade-xlsx-xlsb-check`, `ops-xlsx-xlsb-clippy`, `facade-xlsx-xlsb-clippy` | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/manifest.json` (1,930; `49b2c6336cbfe60822e7546896980fc9f60b93b5d779611d6cf706e6793f5354`) | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/rows.json` (41,107; `b8f7a6c6fe88d115637826ac524c924d182cf201f067ffbf58183595b84fd67e`) |

The accepted row logs are:

* `xldm-tests`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-87e241633e/00-xldm-tests-attempt-f5.log` (11,749; `cda27b3512cf4096af8c9b7e8d0dfce3b0dad2d13c58023a6b88b88238d7f6e9`)
* `xldm-clippy`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-87e241633e/01-xldm-clippy-attempt-f5.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
* `xlsx-tests`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-87e241633e/02-xlsx-tests-attempt-f5.log` (260,918; `f13c9e3598a9b486eefda0db4ed49fb1de3a40f102ea1242f7f5cf5e1620ec24`)
* `xlsb-tests`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/00-xlsb-tests-attempt-f5.log` (122,716; `b137c015a0f203451ef839350cc37101de5cae61f54ca415c7c8719279296d52`)
* `xlsx-xlsb-affected-check`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/01-xlsx-xlsb-affected-check-attempt-f5.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
* `spreadsheet-clippy`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/02-spreadsheet-clippy-attempt-f5.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
* `ops-xlsx-xlsb-check`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/03-ops-xlsx-xlsb-check-attempt-f5.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
* `facade-xlsx-xlsb-check`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/04-facade-xlsx-xlsb-check-attempt-f5.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
* `ops-xlsx-xlsb-clippy`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/05-ops-xlsx-xlsb-clippy-attempt-f5.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)
* `facade-xlsx-xlsb-clippy`: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4/06-facade-xlsx-xlsb-clippy-attempt-f5.log` (0; `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`)

Failed first-attempt state references are preserved for review. Each state
contains its own `manifest.json` and `rows.json`; the listed log is the first
failure that stopped that attempt:

* `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-1c44a0ea11` — `xldm-tests` log, 16,251 bytes, `3438b55f299b0e2c8be9c742580a4572aa723e1711137d7737f30f25fd4e2bce`.
* `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-d87ff2e3d2` — `xldm-tests` log, 1,352 bytes, `99cdf346ecb892bf8a46c8a039eb8056c22d57570605d0104f2eb79cd0d3c6e5`.
* `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-28186f9389` — `xldm-clippy` log, 3,366 bytes, `f0e467cd8098a01aec7c756221309eebad41cfe955c0b638cbd5d703d9ac6f04`.
* `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-4c6d319a08` — `xlsx-tests` log, 11,255 bytes, `7b3e91b0692bb783b346b88028599993866cb846787a5efa80c890c2229c3c3b`.
* `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-b9fb6e31d2` — `xlsx-tests` log, 7,047 bytes, `9efd139b4f540fe686a5ce6d1b5206e594fc75fdfb010ba1efc1e6624885f1cf`.
* `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-2728041fba` — `spreadsheet-clippy` log, 1,564 bytes, `5ba99bcaa58039ce3815a6767a79298a14d057b55dd085d6e698d8765fb79ce1`.
* `/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-eeb484e6b0` — `xlsb-tests` log, 1,701 bytes, `eb728332b8ee992baeb24ebbbef8a8e86e83641ef0fb1b23cd7aed3d4fb339de`.

The final continuation state is
`/home/zhuhe/code/litchi-worktrees/.coord/0362ul68-f5-xldm-c5424e23e4`.
Its final runner-recorded target size was 3,957,146,335 bytes; its largest
recorded live target sample was the same and stayed below the
19,000,000,000-byte guard. The earlier first-failure states remain beside it
and retain their logs and digests; each compile/test failure was fixed in a
separate source commit before the next receipt.
