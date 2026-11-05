# 0362ul26b — compile-only cohort packet

Date: 2026-10-24. This source-only packet prepares the first executable
cohort for the approved generic compile-only overlap protocol. It does not
launch Cargo, native probes, controls, profilers, ABBA, or a qualification
measurement. The optional DrawingML lane is not included. The OLE-owner and
ODF paired runners remain globally quiet.

## Frozen inputs and fresh packet

The accepted production candidate is `(deleted-coord-artifact)`,
HEAD `472b6705475b073e7fe08f6e11ee17fb1e29f838`, tree
`ebc1cee05f124c0c63977aff07a23288c19e7d00`. The five lanes and their
reviewed baseline endpoints are:

| Lane | Baseline HEAD/tree | Candidate | Fresh target/output/TMP slug |
| --- | --- | --- | --- |
| `w4b-xlsb` | `be644568846ad2fcf0622386f2f2438d0eac21c0` / `684a028b2ae30a3a110b87bf6743fc598c6691a2` | production `472b` | `w4b-xlsb` |
| `w4-xlsx-raw` | `0befb112d96d2409e6f40c9abfeb4813665905ad` / `57f0cfa412c67007f006791410d935fc83aca724` | production `472b` | `w4-xlsx-raw` |
| `legacy-rtf` | `86d708878fa537eab42fd4bd5ffbc0f79c606191` / `da5c4ea2b3cef6e32222a3c1565167e8f5e5a89b` | production `472b` | `rtf` |
| `legacy-ole-properties` | `604822b44107999aaddb0bffb96b705e9464f546` / `152ff0abf3a49bb957eeae5ebae7fb2b27156223` | reviewed source-candidate `472b` | `ole-properties` |
| `legacy-xls-properties` | `509bf28e7531ad75f278eaec3a732714a243915a` / `cd282aea544138c152923a3c4c7ba3c7a10bc6b4` | production `472b` | `xls-properties` |

The old source files were not edited. The W4b and XLSX values below are the
tracked templates at the runner source commit (W4b
`32374ee41072e86a7bee5f1a7fd13e172c74d85a17283a61e9b06b12ac244e8d`, 4,958
bytes; XLSX raw
`5b15481cbc139772d3c9ed03bc386e6e4f1dfa4a62a3af327fe2319baabbef83`, 6,140
bytes). The RTF/OLE/XLS values are explicitly the final 472b external packet
manifests from record 0362ul8f, not the historical c796 manifests: RTF
`6cdbd2158793d867218a9a5c85847b0140804c5c28dde6ef577125a17c5d667b` (4,838
bytes), OLE
`6ef65bfa960f4869db7a511bd0d0b55604e2001df4f767c564792e7ee00349be` (6,238
bytes), and XLS
`3495667fdad6f8d9a02929357c376df4574e2448ac2b080316a4d7964f6da2bd` (4,816
bytes). The corresponding historical c796 artifacts are retained and are
`57eff528690834135e74edb897dba55a0373e3a38a2dc3b5c245bedc7b275356` (RTF,
4,660 bytes), `a8a4f34647165fc627ae7aacdc87db846dc7b28b4e0ecfca63c361797522fb8c`
(OLE, 5,637 bytes), and
`36dcc6296d7434b1729c4dbb4bdfde5425539bc7703c9a09f09501715d684ff3` (XLS,
4,693 bytes). These labels preserve both historical and final provenance;
neither class was overwritten.

Fresh rendered copies and receipts are outside Git at:

`(deleted-coord-artifact)`

The generated manifest identities are:

| Lane | Manifest bytes | Manifest SHA-256 |
| --- | ---: | --- |
| `w4b-xlsb` | 7,201 | `c7113cd22bcbc2bccfc5d6daf1d9f64b6dc8b79c5ec880891fe7892bb343e5bd` |
| `w4-xlsx-raw` | 7,524 | `5f0a89072c29570037f6d895f34b345239d0910c9c94338a3faeeb813ab0e4de` |
| `legacy-rtf` | 4,849 | `b0bd81ad05195e5fe52a6f7f9fbce488d2867c95ba6ed40f8c950335a2c3d888` |
| `legacy-ole-properties` | 6,326 | `89978640faee9a90e940888979d76c988c4cdee1522a4cbe941ec08727e6bb2d` |
| `legacy-xls-properties` | 4,904 | `3809bbcd46a76a74375b39ad37405225a02432bceb6314d2dd0b84e11ef77d5e` |

Every lane has distinct `outputs/<slug>`, `targets/<slug>/before|after`, and
`tmp/<slug>` paths. At receipt time all lane-specific paths were absent; only
the packet base directories existed. The lease directory was present and
empty. No old output, target, or TMP path was reused.

The external receipt is 33,078 bytes,
SHA-256 `460ef0f380828d20ef74059489d32e0dfad08d68cf670c76f13b5d87ec343e8c`:

`(deleted-coord-artifact)/packet-receipt.json`

## Permit and phase binding

The immutable permit is `litchi.native-build-permit.v1`, grant
`cfb-compile-cohort-20261024-v1`, job group `cfb-compile-cohort-20261024`,
`max_jobs: 4`, allowed command class `cargo-build`, and expiry
`2026-10-26T20:24:45Z`. It admits exactly the five manifest hashes above and
uses the external lease directory in the packet. It is 4,911 bytes,
SHA-256 `1151ea2453a18db4a850d9d00b5cc9ec46834db1dde8d173c674ab351092ff44`.

The runner added only its fixed `--offline --locked` flags before launch. The
following hashes are SHA-256 of the canonical phase record containing the
normalized argv, cwd, target, and source endpoint; the complete records are
in `packet-receipt.json`.

| Lane | Phase hash (phase ID: hash) |
| --- | --- |
| `legacy-ole-properties` | `build-after-ole_properties_probe`: `a35e9e5e15e895b355c201481d74abd76f2f7929da0eb3da7c1502543746d9c7`; `build-before-ole_properties_probe`: `5eb00189710e585440d02e82b83dca8479fd17ab9fe24812da4ce06135b7900c` |
| `legacy-rtf` | `build-after-rtf_refusal_contract`: `8a1bd5e7541c5f1cf8706320c633899ee419d70909f8b9a6bc3ffe40b65951b1`; `build-before-rtf_refusal_contract`: `a9273c8969dcc2fbeaaa55b6422430a4c50defec56f2207a7ee19f7ac261069a` |
| `legacy-xls-properties` | `build-after-xls_metadata_bench`: `83ab157315cac5d2c4a9987e7680f25b3eb2791a84f8024d75c86c2f8d28ec33`; `build-before-xls_metadata_bench`: `7aef885b6da52ef8fa583fa19f01d06c826d47c553563ef40d578892710b6ae8` |
| `w4-xlsx-raw` | `build-after-xlsx_raw_identity`: `ffa08b7bd658ff41d67d0f8537fdd9548062f7d0bdec4fd2a933e7be1bdefc5e`; `build-before-xlsx_raw_identity`: `bff56709fbb0194aa7ce3ba2e4832b8fd194bd3d64a24b4d6aafeda67cd4d209` |
| `w4b-xlsb` | `build-after-xlsb_w4b_probe`: `009a0507f1c8b9c1a220809998f0bcbb8a9b455cf3e168e29ae984776dbb87c4`; `build-after-sheet_dispatch_bench`: `592f3a47f4bbd34f02eb43125dd8415aa68232b1c049c5492e2d678707243e3c`; `build-after-xlsb_crud`: `496d5dec72795cfb8ae219db58b8134abf83cae019e1d1be7cf9f104b9952928`; `build-before-xlsb_w4b_probe`: `41f0785247f44600bfd798a68ca28531b34cf511105984e21170bf1ae2a15d0a`; `build-before-sheet_dispatch_bench`: `18712acf1f4643d9043f023292173899172d0083fe6ae6d64ffd7266c48ed2ab`; `build-before-xlsb_crud`: `1c4f2981a27d3a687a2d262c1f83f8ac605361d53f45adb40991e10c9dcb8476` |

When the coordinator opens the quiet build window, each admitted lane can be
started with the generic runner, for example:

```sh
unset LITCHI_CAPPED LITCHI_BLESS
python3 tools/native_qualification_runner.py \
  --manifest (deleted-coord-artifact) \
  --phase build
```

The same command applies to each of the other four rendered manifests. The
coordinator must launch no more than four at once and must not add a lane or
change a manifest, source, target, output, TMP, phase, or permit. The runner
still owns `CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, `LITCHI_MEM_MAX=20G`,
target/output freshness, path/config validation, and the 20,000,000,000-byte
target limit. A permit does not admit controls, validators, ABBA, rustdoc,
profilers, or any non-`cargo build` phase; those remain globally exclusive
and must be run later in the quiet window.

## Source-only checks

The current generic runner (`9a1c5a6042b7413e0e63979848b963eacfc6229f`) is
195,892 bytes, SHA-256
`23e1c8384723d203b685d742ce788a2c35158081d310c82c37faa7ef6277338b`.
Its focused test source is 80,795 bytes, SHA-256
`4e947762d0236cf58b684cef17c1c9374d2d43ef90ad529cddfd571ae885940a`.
The legacy preparer, legacy validator, XLSX preparer, and XLSX validator are
respectively pinned to
`3aad0d288aa653643d2efa2782b8ac5d96ad82fbdc22a07d57cb9fa1b47871bf2`,
`9f2d2f7a4eac2eb51b8918b3e8e8d0670cec6d5c20ad25b30bf4010f80593dc7`,
`76e7fa1c90a2dff8b34324bfb40bc83fedc9500f67c7949ca142002cde9502a4`, and
`8d65c18d2efd00a0f26ffd182089fd1a18e12630d36b9d4254a7d8ca7aab8b03`.

The accepted W4b/W5 source-preparation receipt is 6,404 bytes,
`664fac47612eacc9c21caf5099785b8486e050b9dc4644ce55dc682c31d9f0b2`; its
manifest-source validation receipt is 36,381 bytes,
`1029b6e4f7c59ca72a0806b3f537cc2ce1f87f4e6e47c06ffd57adbbadf5e2f7`.
The fresh XLSX preparation receipt is 10,340 bytes,
`4dafe3c38ca84f66d704819bec90c45ef438a3a51d35e5ccef635d89f355337e`, and
the three fresh legacy receipts are:

| Lane | Bytes | SHA-256 |
| --- | ---: | --- |
| RTF | 7,392 | `ea5245b73cce463fcd214edb4f0a4cb8660e152c541483bca147c690f6c1aa5e` |
| OLE properties | 7,814 | `c3e201ebd873a205891bf62e0480ed4991b6e9e38ed9d8427341b5a31846ef51` |
| XLS properties | 7,726 | `65aeb76da428f226d9b0885ff7d0768130484415527f1e5a2feb499f97b27c6d` |

The current runner loaded all five rendered manifests, accepted all five
permit bindings, and resolved every normalized build phase without starting
Cargo. The current runner then accepted every generated legacy preparation
receipt and source identity. The static-validation log is 12,536 bytes,
SHA-256 `ec31904c377457479bd969866d2a90fadb54bc7d56a5a210a88423db491a0f1b`;
the legacy identity log is 2,217 bytes,
`b89c9459f7f7f338e5548be992cad14d6fc65e836fdb3d6c5a234fcf8e021530`.

The focused source-only command passed 78 tests (exit 0); its external log is
333 bytes, SHA-256
`c45f27167413fe23c2c7a554da1ac5986a546c45762071e4328c5924ee8d9e64`.
The source-preparation commands exited 0. Their generated receipts are the
authoritative outputs; their stdout/stderr was not captured. No Cargo build,
correctness probe, timing, allocation capture, ABBA result, RSS result, or
performance claim is made by this record.

## Follow-up v2 — explicit TMPDIR repair

Independent review found that the v1 W4b and XLSX rendered manifests did not
carry an explicit `TMPDIR`, although fresh TMP directories had been reserved.
The v1 packet above remains immutable. This v2 packet preserves the same
`litchi.native-build-permit.v1` protocol and source endpoints, adds a distinct
explicit TMPDIR to both manifests, and regenerates every manifest-bound
receipt and permit. It is outside Git at:

`(deleted-coord-artifact)`

The v2 manifest identities and TMPDIR bindings are:

| Lane | Manifest bytes | Manifest SHA-256 | TMPDIR |
| --- | ---: | --- | --- |
| `w4b-xlsb` | 7,333 | `ba50bf05b13f177640f36234d4bd42692708c1dd36edad1bec1d3e514c80858f` | `packet-v2/tmp/w4b-xlsb` |
| `w4-xlsx-raw` | 7,697 | `14bbbe0886b8ac8989dd0c37691b31575ce84645330afb65668ad0717c452e24` | `packet-v2/tmp/w4-xlsx-raw` |
| `legacy-rtf` | 4,849 | `d87e9bc74dddda54111756ad6a96d5a0ebdbf580c3c89bd310fe76efda7e67d9` | `packet-v2/tmp/rtf` |
| `legacy-ole-properties` | 6,326 | `f1980403282d90a4f9438455b3d789d5f7ba8b0bd765d66ad827b6f7811f5ae9` | `packet-v2/tmp/ole-properties` |
| `legacy-xls-properties` | 4,904 | `2acb6b0caa1830cab16c63ec7494e37920553e510e0ef8ac5b97c1bacd9ca447` | `packet-v2/tmp/xls-properties` |

All v2 lane-specific output, target, and TMP paths were absent at receipt time
and are pairwise disjoint. The v2 permit is 4,809 bytes,
SHA-256 `2b97f236d018905091bf2624d7f5e941c038839715b8fd2e992e99290371ee4f`,
grant `cfb-compile-cohort-20261024-v2`, `max_jobs: 4`, and expires at
`2026-10-26T20:47:20Z`. The v2 packet receipt is 36,078 bytes,
SHA-256 `7d15632e350641bd3041ce2c567adb143b0f0943daedf1bca24671882f7e1ac0`.

The regenerated preparation receipts are:

| Receipt | Bytes | SHA-256 |
| --- | ---: | --- |
| XLSX raw | 10,340 | `4dafe3c38ca84f66d704819bec90c45ef438a3a51d35e5ccef635d89f355337e` |
| RTF | 7,392 | `831bf4ca78b6991ecdf0f7026bad60c6af74f2d0f51a784af2b74aa83b4a2bfb` |
| OLE properties | 7,814 | `b4f4d9f54969da168f56c33b711407db2fcbc506b4603471e538482375313799` |
| XLS properties | 7,726 | `f82c01160f28554ba1d5f12064b62398a5f455bba2226f93d29a096f76808f1b` |
| accepted W4b/W5 source preparation, copied byte-for-byte | 6,404 | `664fac47612eacc9c21caf5099785b8486e050b9dc4644ce55dc682c31d9f0b2` |
| accepted W4b/W5 manifest-source validation, copied byte-for-byte | 36,381 | `1029b6e4f7c59ca72a0806b3f537cc2ce1f87f4e6e47c06ffd57adbbadf5e2f7` |

The v2 static validation passed manifest loading, all five permit bindings,
all 14 normalized build phases, and TMP/output/target disjointness. Its log
is 13,094 bytes, SHA-256
`c1fcf2b84ddce5997a7be7dec0d9043f625b2c9695e69cbbe82ae3005da9ba28`.
The current runner accepted all six regenerated legacy source legs; its log
is 2,217 bytes, SHA-256
`b89c9459f7f7f338e5548be992cad14d6fc65e836fdb3d6c5a234fcf8e021530`.
The repeated Python-only suite passed 78 tests (exit 0); its 333-byte log is
SHA-256 `cf6c0b8e2d82d0c387e7465b033d113bf0a4692339fe2c1ed79b9d5a25bc8fdd`.
The XLSX preparation stdout/stderr log is empty (0 bytes,
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`).

The exact v2 phase hashes (canonical normalized argv, cwd, target, and source)
are retained in the packet receipt:

| Lane | Phase hashes |
| --- | --- |
| `legacy-ole-properties` | after `d8881c09a50324b79ff492b0749203ec8cabe6874c061a70f9db0d3cb9efece0`; before `b2def306ef9d938e286dc7923b43f91d251482acb62198a5182929888b4ecd8a` |
| `legacy-rtf` | after `aba9ce6095d3f02ab7b535be8677e74a97fcf6bc4d9df22a11d6bcd5b9898cb9`; before `db98efd28040d166361b3dd59d1e0349c4e1a2b367e467949df8e8c3b1ed2f1c` |
| `legacy-xls-properties` | after `1d6e02a60f957a217d58c1c0f499a7618665481ac7a7957009466c9c991cb9de`; before `97a86d5fd3d3ec97de2720e2a44daa88f1bfb1e6b6c7ad7929cfd62c5b59ddc3` |
| `w4-xlsx-raw` | after `9ef2771a6dc353f341a46d47c17af34a49e49f406df711b344d18eda07035d67`; before `409abdaf9357862b825bdab7ae84afb8529268edb595dad27b9ba7f05f3e9179` |
| `w4b-xlsb` | after sheet `ed2f07124c1580677243360db823023b095a55553d9b07192e88bfbe437b9c09`, after CRUD `acab7259936646f29b0df58cd42bfc8e1df9eb7b7973d0c3df364e0c1ab30bc6`, after probe `1bb18f293198a2f35e8ee893c7c8c3bca62b7d7d312d1e8d1794d1ee53173f45`; before sheet `8f79f20de0008f98af76656b22763bdc26d3da4b6fc54f21b0190f4611bf3d63`, before CRUD `45f937e8a39b5fe05f8d3a7a9f29be017fcf8dffe646719c3ac27cc098e6d3df`, before probe `f1ccf7bf6e3052e6da1c49a4f2f834e653bbb012efb12f619a65a617d86302e0` |

The v2 launch command remains build-only and is not authorized while the
known unleased full-gate builder is active:

```sh
unset LITCHI_CAPPED LITCHI_BLESS
python3 tools/native_qualification_runner.py \
  --manifest (deleted-coord-artifact) \
  --phase build
```

No v2 Cargo/native command was launched. Controls, validators, ABBA, rustdoc,
profilers, and all measurements remain globally exclusive.
