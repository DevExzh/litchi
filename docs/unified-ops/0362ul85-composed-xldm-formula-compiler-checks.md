# 0362ul85 — Composed XLDM and Formula compiler checks

Date: 2026-10-25. Status: all nine requested compiler rows pass on frozen
source `f14235da957f81aaff5316c18f7992b656619264`. These are compiler checks,
not native execution, timing measurements, a full gate or final acceptance.

The detached source is
`(deleted-coord-artifact)`.
Root verified its exact commit, detached/clean state, selected tracked compile
inputs, approved locks and runner before the accepted run. The capture includes
13,384 tracked compile paths and 1,964 test-data files. No missing compile input
was reported. Existing compiled artifacts were reused with incremental compilation
disabled; no native/performance freshness claim is made.

The first f14235 attempt passes the four spreadsheet/ops/facade rows, then
stops because its DOCX command requested a nonexistent `formula` feature.
The coordinator confirmed the feature declaration and corrected the command
to DOCX `--all-features --all-targets`. A fresh continuation passes that row
and the remaining four. The failed attempt remains failed and retained. An
earlier interrupted capture on `11d14a661d` is separate historical evidence
and contributes no pass to this result.

All commands use offline/locked Cargo, one build job, one test thread, no
debug artifacts, no incremental compilation and the capped runner. The live
target ceiling is 19,000,000,000 bytes and the hard ceiling is 20,000,000,000
bytes. The maximum sampled target in the continuation is 3,687,596,149 bytes;
its terminal target is 3,687,603,132 bytes. Every terminal ownership sample is
empty. The source and all three locks are unchanged.

Approved runner SHA-256:
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.

| Lock | SHA-256 |
| --- | --- |
| `Cargo.lock` | `efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811` |
| `tools/perf-baseline/Cargo.lock` | `de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea` |
| `tools/native-resave/Cargo.lock` | `989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae` |

The accepted rows are:

| Row | Seconds | Log SHA-256 |
| --- | ---: | --- |
| `spreadsheet-check` | 80.003 | `12e8d538489cefbd2d33e102033c5388bad8c4f030e48000a00294ab63c0c990` |
| `spreadsheet-clippy` | 110.011 | `6d15a20b216be9aa49a56506f8755b294311da6c785ce47c5d2da65210ee565b` |
| `ops-xlsx-xlsb-check` | 55.679 | `06e719d7fe8f23f7faa67a96ee275584635ab82c558d432ec440ba807da60be4` |
| `facade-xlsx-xlsb-check` | 42.463 | `937b08d731c18c8d92b2bec69a66c720047c56bf24f2cbaed77842b31fbe22ad` |
| `formula-doc-docx-check` | 54.609 | `6dae037d1149acbb8b90266e8257883b232e2b392a4418ef63020fb562dea4ae` |
| `ops-facade-doc-formula-check` | 43.250 | `54b4338a03258dbf4eede7baaf70ba132909e63132c28d3ceeea03ad962b28bf` |
| `ops-facade-docx-formula-check` | 43.970 | `d24a1f6d0d326d03744ceee897a633a3dba634db1886ff75e48ac771fb90c620` |
| `perf-baseline-check` | 147.496 | `bd461c53030a583cb4b3f889f71368b152241fcdebdde14bb32eaacd9c505a3f` |
| `native-resave-check` | 3.406 | `a0f70ec6abfc9d11684db99d418bc6c2334d607d86ddad0d6dae88529b3d2312` |

The exact command arrays, source fences, lock identities, log sizes/digests,
Cargo target traces and terminal ownership samples are bound by the following
external receipts. Paths are relative to
`(deleted-coord-artifact)`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `f5-compiler-validation/manifest.json` | 2,292 | `8994e88d3f9767a4c0e5286cb42f8bea9c5706bc09d811ec9c55a00c64a804d3` |
| `f5-compiler-validation/rows.json` | 37,670 | `dc90345d09dd1d2e45271d8f7ec29bf3b1cacd19d595246f7f34c7ffe0178458` |
| `f5-compiler-validation/result.json` | 39,334 | `a6d7d71280cf25b91d0c8e1764ac3a4dce92dc47777640ef5391281baccf67ea` |
| `f5-compiler-continuation/manifest.json` | 2,310 | `79305d1fc09959aa34559b82d7785d5d60e786625a25996eacab57e54e1bcb6d` |
| `f5-compiler-continuation/rows.json` | 26,894 | `430264e7a3a5c7c25f625e1c4b7184af2d310c0c0e054170b389b677c48229a3` |
| `f5-compiler-continuation/result.json` | 28,564 | `584d7acb25a7ac5bfcf699ccebd3152c4639a47f8e3514fcebd5d2f37b2441a8` |

Root independently verified all accepted rows and the failed command-spec
row against the actual logs and traces, source/lock identity, target guards
and empty terminal ownership. No failure is relabelled. The remaining PPTX
engineering, final API census, complete gate and native/performance qualification
remain required. The build slot is released.
