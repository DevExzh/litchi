# XLS/F3 and XLSX native execution handoff

This is a source-only handoff. It makes no acceptance claim and ran no Cargo,
native, timing, allocation, profiling, or qualification command. The complete,
hash-bound external packet is:

`/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/sheet-native-execution-handoff-472b-slot3-cfb-v2.json`

Packet SHA-256:
`a203bfac3cf3f4013be3e1e6853d74dadf8b475868251e9c7ac03de8de4558fd`.

The reviewed F3 role-split driver is commit `d447c8c0b9efe2d4394c47419cb7768ec63834fc`,
SHA-256 `419dfe0d55a0f8ba3cef3ad5ee247917da46a4b0c4d7efb3916a377fa771284b`.
Its later qualification command is bound to the actual slot-3 candidate binary
(`fff038b1320e20666f4e9cb49d7a907350dd9ecabce50eabfdecef100b0a7e54`) and
reuses the original baseline binary
(`9d39aa79e808096cd0487d88e8eddd277182ff5ab91a501d8e9caf6793d3be2c`).
The old unsuffixed and superseded candidate paths are forbidden.

The XLSX source-family driver reuses the proved baseline/candidate timing and
compact-merge binaries. It records all six eager/source families, five ABBA
blocks, two warmups, 500 ms target, 150 ms minimum, and compact merge/unmerge,
with CPU 30 and fresh output/TMP roots.

The raw XLSX preparation receipt, reviewed CFB cohort-v2 manifest/permit,
runner `23e1c838`, preparer and identity validator are all hash-bound in the
packet. The previous generic-runner binding is preserved only as superseded
provenance. Its required phases are
the locked before/after builds, then 31-round before/after/after/before ABBA,
both `raw-alloc` and `raw-native` routes, four cases, five iterations and
two warmups. The raw lane is the only concrete blocker: its before/after
binaries have not been proved in this handoff and require a quiet build slot
before ABBA. All lanes use the recorded 20 GB runner/target contract; F3
children additionally use a 2 GiB address-space limit.
