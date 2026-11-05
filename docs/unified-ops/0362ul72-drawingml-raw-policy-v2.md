# 0362ul72 — DrawingML raw qualification policy v2

Date: 2026-10-25. This is a source-only protocol implementation. It does not
run Cargo, rustdoc, a native probe, or a timing workload. The existing v1
receipts and records retain their original meaning and are not relabeled by
this change.

## Fixed historical library

Version one copied the candidate performance library and maintained a list of
reverse consumer substitutions for each later API merge. Version two removes
that maintenance. The historical endpoint receives the exact Git blob at the
fixed, trusted source ref below:

```text
ref:       472b6705475b073e7fe08f6e11ee17fb1e29f838
path:      tools/perf-baseline/src/lib.rs
blob SHA-1: fbc55c5dcdb7cf86e5377e4f8012954a7fb4e572
SHA-256:   280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a
bytes:     3,088,216
```

The preparer verifies the fixed commit, tree blob, Git blob id, byte count,
and SHA-256 before writing the historical file. No ref, blob id, or historical
hash is accepted from the command line. The candidate retains its current
library and the same probe, fixture, validator, and route files are copied
byte-for-byte into the historical worktree; the library itself is never
rewritten.

## Role-separated dependency identities

The v2 policy binds each role's lock independently:

| Role | Baseline SHA-256 / bytes | Candidate SHA-256 / bytes |
| --- | --- | --- |
| workspace `Cargo.lock` | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` / 139,606 | `efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811` / 139,189 |
| `tools/perf-baseline/Cargo.lock` | `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c` / 57,227 | `de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea` / 56,941 |
| `tools/native-resave/Cargo.lock` | `51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378` / 56,975 | `989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae` / 56,689 |

The dependency policy records the three reviewed root transitions:

```text
allsorts 0.17.0: registry -> vendor/allsorts
flate2   1.1.10: registry -> vendor/flate2
zlib-rs  0.6.8:  registry -> vendor/zlib-rs
```

The candidate validator checks the three manifests, the exact role lock
bytes, and every tracked blob below all three vendor roots. A sparse checkout,
missing tracked vendor file, symlink, unreviewed vendor root, or modified
tracked blob fails closed. The vendor-tree digest and manifest hashes are
part of the checked-in candidate binding, not user-supplied receipt fields.

## Incomplete candidate binding

The final candidate commit was not frozen when this leaf was prepared.
`CANDIDATE_BINDING` is therefore intentionally `None`, and the v2 `prepare`
and receipt-validation entry points fail before touching a source or receipt.
The coordinator must bind the exact candidate commit, manifest identities, and
vendor-tree digest in a later source-freeze change. Supplying a different
candidate hash or an external JSON policy cannot activate the protocol.

The focused source-only suite covers fixed historical-blob materialization,
role lock separation, the three dependency transitions, sparse/mutated vendor
trees, manifest pending status, and fail-closed candidate preparation:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_drawingml_raw_policy_v2
```

It passes 9 tests on this leaf. No performance, allocator, RSS, fidelity, or
native acceptance claim is made here.
