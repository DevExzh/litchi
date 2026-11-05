# XLS F3 role-split paired handoff

The historical baseline blocker is resolved at the protocol layer without
touching either frozen source endpoint. The new external driver commit is
`d447c8c0b9efe2d4394c47419cb7768ec63834fc` on
`uo2/xls-f3-driver-role-split-v2`; it binds control to the historical
`0b8f68d7e9d3c97cfa62cedceacdeb2ad5f4ba76` consumer adapter and candidate to
the exact `472b6705475b073e7fe08f6e11ee17fb1e29f838` tree. Non-diagnostic
defaults enforce both refs and calculate independent probe, overlay and
source manifests.

The existing baseline binary is reusable: its endpoint identity is
`f40dfbc6ec0dadd20700c7d58facfb6ead264e02463cfb596999e31737fd5633`, its
successful build log is `4e50d9b7d0862913eaf8c9032f3ef4d3a86ef8fdf1abf24dec4c0543a2d72e03`,
and its binary is SHA-256
`9d39aa79e808096cd0487d88e8eddd277182ff5ab91a501d8e9caf6793d3be2c`. These
match the preserved build receipt, source endpoint, 0b8f probe overlay, F3
lock and fixture; no rebuild was performed.

Workload equivalence was checked against the actual baseline/candidate F3
files. `xls_f3_corpus.rs` differs only by two required `OleFileLimits`
arguments and normalizes to `a2229868d7ee46f26cae212cc0164a3e4882d0ad85a3b2bffe366abdfbc2e8b9`.
`xls_f3_qualification.rs` differs only by five required limits and making
`--implementation` explicit; because the runner supplies that flag on every
leg, its normalized token hash is
`cf117411c1f5fc8dd5ecc533978342f2734131f72fcbf8bafcdc2ceb0c291529`.
The F3 package library and both examples are byte-identical.

Combined packet:

- `(deleted-coord-artifact)/sheet-native-readiness-472b-f3-role-split.json`
  SHA-256 `51258c4b0e296af95cb675621f91a3f370cfbafb5c410880b1c7dc5bae205df5`.

The candidate build has a fresh target/TMP root in that packet; the paired
driver command reuses the exact successful baseline binary and points at the
fresh candidate artifact. Source identity, role-split options validation and
driver compilation passed. No Cargo/native/qualification run or acceptance
claim is included; root controls the build grant and final review.
