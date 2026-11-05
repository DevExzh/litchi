# XLS F3 candidate v2 slot-3 build

The granted build-only phase succeeded from the frozen clean candidate
endpoint at `472b6705475b073e7fe08f6e11ee17fb1e29f838` (tree
`ebc1cee05f124c0c63977aff07a23288c19e7d00`). It used CPU 30, the candidate
worktree as cwd, the 1.95.0 toolchain, offline/locked Cargo, JOBS=1, debug=0,
20G memory cap, the coordinator shim prefix, and fresh roots:

- target: `(deleted-coord-artifact)`;
- TMPDIR: `(deleted-coord-artifact)`;
- PID/SID/PGID: `1518284/1518284/1518284`;
- build log: `xls-f3-candidate-472b-repair-v2-slot3-build.log`, SHA-256
  `44964368d787e248306ab199720f4dfc75305cc9b2426198e4e0e355cfa7aab7`.

The binary is
`.../targets/xls-f3-candidate-472b-repair-v2-slot3/release/examples/xls_f3_qualification`,
7,463,128 bytes, SHA-256
`fff038b1320e20666f4e9cb49d7a907350dd9ecabce50eabfdecef100b0a7e54`.
Cargo ended with `Finished release profile [optimized] target(s) in 3m 11s`.

The persistent target guard used PID 1530055, tolerated disappearing files,
recorded a 353,713,181-byte peak under the 20,000,000,000-byte limit, and
ended cleanly. An initial inline monitor recorded one sample, then detached;
the gap and restart are recorded in the receipt, with no continuous-monitor
claim. The old failed target/log and successful baseline binary were not
touched. Qualification/ABBA did not run.

Full external receipt, including structured argv/env and the later driver argv
bound to this slot-3 binary:

`(deleted-coord-artifact)/sheet-native-build-receipt-472b-f3-candidate-v2-slot3.json`

SHA-256 `47abc06d0cd2463d1653d1afc4a923ee3323e2fd5b3243c431f9484ef38f8f33`.
