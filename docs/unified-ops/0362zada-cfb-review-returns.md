# 0362zada — CFB retained-policy and composition review returns

Date: 2026-10-24. Parent: [0362zad](0362zad-cfb-explicit-limits.md).

The CFB wave was returned before integration for four lost-policy paths.
Facade eager open/unlock now derives its CFB profile from the admitted
`OpenOptions` ceiling; source-layout adoption requires and retains the
caller's `OleFileLimits`; sequential validation derives its physical limits
from the planned output; overlay composition and inverse replay retain
`SharedOleFileLimits`. Tight-directory and allocation regressions exercise
expansions that a default profile would admit. The larger V4 output and
eager-ingress regressions prevent the inverse error of restoring a smaller
default. Exact leaf commands, source patch and immutable log hashes are in
0362zad. Independent review accepts leaf `bf7f48741bc02d1f5e736a7a732c25808d28af36`.

The coordinator composes that leaf with staging `2021a41c4cb217844131892138a7d0452a81d6c1`.
The XLS validation conflict keeps both its explicit semantic `SourceLimits`
and the canonical `OleFile::open(reader, cfb_limits)` call. The glossary
conflict keeps the completed ZIP rows and adds CFB; ZIP is already merged.

The first composed workspace check failed because the legacy raw repair's
VBA conversion was added independently of the error-extensibility wave.
The DOC core projection and File classifier now each have a future-variant
fallback. All known variant mappings remain unchanged; the DOC fallback is
fixed and content-free, and the File fallback is `CorruptFile`. Independent
review accepts resolved index tree `de139650d80d513854c81a60df649fb7b26dd883`.
No production CFB implementation changed during coordinator conflict resolution.

## Composed validation

The initial failing check is preserved, not replaced by the rerun:
`workspace-check-2021a-bf7.log`, 13,943 bytes, SHA-256
`e16a4924a35feb66af31c48ee8228ad62b02b18ac66ad09a60602a7bb2787ea0`.
It ran on index tree `115e8f2579d548eca447095b51f5b89ce08665a0` and exited 101.
Both runs use root lock SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`
and the capped Cargo shim with `LITCHI_MEM_MAX=20G`, one build job,
`CARGO_INCREMENTAL=0`, disabled debug info and offline resolution. The
external target is `.coord/target-cfb-explicit-limits`.

Command:

```sh
cargo check --workspace --offline --locked --all-features --all-targets --message-format=short
```


The next composed run passed the VBA projection but exposed the chart-builder
migration's redundant `litchi_xlsb::builder` qualification in an integration
test whose `builder` module was already imported. That test now uses the
existing module import. Its failed log, `workspace-check-vba-fallbacks.log`,
is 8,803 bytes, SHA-256
`40cdc4a5e38c2ba538b36558b589dc21b1e6d2667f3613c135d662ac16899333`.

The corrected assembled index tree
`57abd651f19bc0ee91f24e1f60d2850a41eb7f6c` passes the complete workspace
command above: exit 0, 285.068 seconds wall time. The immutable log is
`workspace-check-vba-chart.log`, 7,279 bytes, SHA-256
`536bdfa7b6ffca2e96b17bdfdb4ff9eaff95c2db2ab2d38cc0fcaca865dd8919`.
The lock remains byte-identical after every run. Only this record and reduced
JSON receipts are added after the validated source tree.

The three reduced receipts are committed under
[`results/0362zada/`](results/0362zada/); full logs remain under
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/cfb-composed/`.

The four composed glossary/inventory/disposition modules passed 91 tests
in 1.305 seconds before these two compile-only fallback repairs. Changed-file
rustfmt and diff checks pass. Final capture, the full gate, corpus replay and
native qualification on the assembled source remain required. This record
accepts the source repairs and their stated validation, not final qualification.
Only redaction audit 22 remains deferred. Nothing is pushed.
