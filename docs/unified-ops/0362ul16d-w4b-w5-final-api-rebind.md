# 0362ul16d — rebind W4b/W5 preparation to final 472b production

This packet rebinds the reviewed W4b/W5 source preparation to the final
production source `472b6705475b073e7fe08f6e11ee17fb1e29f838`, while leaving
the c796 packet untouched. It is source preparation only: no Cargo build,
executable correctness run, timing, allocation, or acceptance claim is made.
The reduced receipt is
[results/0362ul16d/source-rebind.json](results/0362ul16d/source-rebind.json).

## Source and baseline identities

The final production checkout is clean and has tree
`ebc1cee05f124c0c63977aff07a23288c19e7d00`. Two clean read-only candidate
endpoints at that same ref are used because the preparer requires pairwise
disjoint source roots:

```text
W4b after: /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
W5 after:  /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source-w5
```

The fresh baseline endpoints are:

```text
W4b before: /home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final-472b/sources/w4b-before
W5 before:  /home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final-472b/sources/w5-before
```

W4b before remains commit `be644568846ad2fcf0622386f2f2438d0eac21c0` with
synthetic tree `684a028b2ae30a3a110b87bf6743fc598c6691a2`; the tree is not a
commit. W5 before remains `7afba5f146aa1a16eba73ebed2144dd041b33fb5`.
Only the recorded baseline overlay paths are dirty.

## The final W5 probe delta

The W4b probe, CRUD consumer, ordinary W5 probe, route module, analyzers,
locks and manifests are byte-identical to the c796 packet. The final
`legacy_w5_source_probe.rs` is not equivalent to the c796 file: it adds 18
insertions (2,382-byte binary diff), growing from 19,083 bytes and SHA-256
`83b5eb8f…` to 19,983 bytes and SHA-256
`b808e4dc…`. The additions are an optional
`LITCHI_PROBE_OUTPUT_DIR` sink that writes each selected scenario's output to
an external directory. The default timed routes and stdout report remain.

The preparer retained that sink in the adapted historical baseline probe
(20,753 bytes, SHA-256 `224d2cd0…`) and removed only the candidate-only API
calls required by the historical source. The current `legacy-w5.json` does
not set `LITCHI_PROBE_OUTPUT_DIR`; this record therefore makes no claim that
the c796 and 472b probes are behaviorally identical, and the final run must
use the reviewed manifest/environment exactly as recorded.

## Locks and validation

The source-preparation receipt is external at
`/home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final-472b/receipts/source-preparation-472b.json`
(6,404 bytes, SHA-256
`664fac47612eacc9c21caf5099785b8486e050b9dc4644ce55dc682c31d9f0b2`). The
static validation receipt is external at
`/home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final-472b/receipts/manifest-source-validation-472b.json`
(36,381 bytes, SHA-256
`1029b6e4f7c59ca72a0806b3f537cc2ce1f87f4e6e47c06ffd57adbbadf5e2f7`). It
passes manifest/workload validation, normalized Cargo argv checks,
source/target/artifact containment, analyzer bindings, fixture identity and
runner `_source_identity` for both legs of both lanes.

The historical baseline root lock is the reused generated lock
`cbe70791291d4a0830a8749985aa717230abf74d61c108fb284d25990c9ad681`
(`139,442` bytes), with perf lock
`57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f`
(`57,028` bytes). The final production root/perf locks are
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`
(`139,606` bytes) and
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`
(`57,227` bytes). All eight lane lock identities are bound in the reduced
receipt.

The focused protocol command passed all 41 tests; its external log is
`/home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final-472b/receipts/protocol-tests-472b.log`
(332 bytes, SHA-256
`a54eaa63cff67d4408aa370450c3249bc90b3db8eba9e47e389684484c1ce022`).

The earlier c796 packet remains preserved at
`docs/unified-ops/results/0362ul16c/source-preparation.json` in commit
`a1ffe12a048109bf23487e76a1b11523786327cc`; no c796 receipt or source
endpoint was modified. The build, analyzer/correctness and quiet ABBA
commands remain deferred until a separately authorized native slot.
