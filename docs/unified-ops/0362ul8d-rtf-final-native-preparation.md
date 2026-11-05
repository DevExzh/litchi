# 0362ul8d — RTF final native source preparation

Date: 2026-10-24.  This is a source/preparation receipt for the final RTF
native qualification.  It does not claim a Cargo build, native execution,
timing result, or acceptance.

## Source pair and reviewed adapter

The production API source fence is
`69194102b6b702b720a73f074e3cefeb2f69526a`.  The preparation pin repair is
`82798234ede3d77bc2a96a36d2da5d0c9aafaec2`; it changes only the RTF
preparer/validator/test provenance pins, so the candidate API tree remains the
production 691 tree.  The fresh external source roots are:

```text
baseline: (deleted-coord-artifact)
candidate: (deleted-coord-artifact)
```

The baseline is the frozen `86d708878fa537eab42fd4bd5ffbc0f79c606191`
checkout plus the reviewed consumer-only adapter.  Its only dirty paths are
`tools/perf-baseline/Cargo.toml` and the added
`tools/perf-baseline/examples/refusal_contract_bench.rs`; the candidate is
clean at `82798234ede3d77bc2a96a36d2da5d0c9aafaec2` (the production API files
are unchanged from 691).

| identity | baseline | candidate |
| --- | --- | --- |
| `tools/perf-baseline/Cargo.lock` | `57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f` | `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c` |
| `tools/perf-baseline/Cargo.toml` | `af876fbf715edb643d2c7abb8643efb252413c640824890e90b1849207889676` | `61e103809753ffc528b2cd2c2e00a1525d40fc4e6360668ad982bec22a60bb73` |
| `refusal_contract_bench.rs` | `0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81` | `0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81` |

The live current workspace example remains separately bound to
`crates/litchi-rtf/examples/refusal_contract_bench.rs`, SHA-256
`1726acb58953ae3f57d9b74207f75678a3bcc7ac4c993443aac5583d3943f791`, and
the external root lock
`(deleted-coord-artifact)/4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1.Cargo.lock`,
SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.

The old staging pins (`6dbfc0…` lock and `60cde3…` manifest) and all old
endpoints/evidence remain historical in 0362ul8b and the external
`rtf-probe-compile` directory.  They are not silently reused for this final
source pair.

## Fresh driver manifest and preparation receipt

The immutable external manifest and source-only receipt are:

```text
manifest: (deleted-coord-artifact)/legacy-rtf-final.json
manifest SHA-256: bcdf11642a7d7eba139781ff9ac113013f78b23cc0cdba2e154a6d47c91a26f8 (4649 bytes)
receipt: (deleted-coord-artifact)/source-preparation.json
receipt SHA-256: 0058a7049b09903134b2f9653421217c9d737670f76fc38a6320e77ae3bbc8cd (7504 bytes)
The superseded first receipt, generated before the candidate revision was
bound to the pin-repair commit, remains at
`driver/source-preparation-before-revision-fix.json`, SHA-256
`5349e7ff15b8a2f437fba218870b463f254a7615cbe6f0cd5f3eddf3f8550441`.
```

The receipt binds the candidate-side preparer
`prepare_legacy_native_inputs.py` (SHA-256
`cc6ef9b6ed02cbbd070c94cdf2c4c9c868aa7f55f6cefc1d461495b43e96dd88`), the
validator (SHA-256
`0ba2f93e5163bd1120985a5adf453dd179b9f2bacd3bb818d1ff8ff05819e92a`), and
the runner (SHA-256
`07fb47403d2fb27a560ba7afba9f59cefbce0b561bb6cc38161ab33b2ba3b62a`).  The
recorded source identity digests are:

```text
baseline tree: da5c4ea2b3cef6e32222a3c1565167e8f5e5a89b
baseline source-tree: 5089da679079ee37dc9964669e9ac726483216dd87f3faac74b260466be99551
baseline identity: ef4918a814bf428fdd1f4b0cfe6cf77c30306c4e7000b6b7e77d69f71ba608b0
candidate tree: 327fc63edbcd90f8a516b0b09e4f8880906532a
candidate source-tree: ae242f9af29d6f467123891c8938f327892a6c121ed565f9182817773e6894f8
candidate identity: 4280454216291cab4b6aba7f18db333f2d51353cec0aedac8807c14fe548a2b1
```

This lane has no external fixture files; its fixture identity is the explicit
empty list.  The workload remains 31 rounds, five ABBA blocks, exact order
`before, after, after, before`, and all eight reviewed routes.  The source
adapter still preserves the native error-variant distinction, fixed
content-free refusal checks, operation-only timed replay design, semantic
replacement/remainder/reopen/inverse/no-op/tail checks, and source/target/output
separation.

## Source-only commands

The historical receipt was produced with the following relative-path command:
it is retained as executed evidence, but is superseded by the manifest
identity repair in [0362ul8e](0362ul8e-legacy-manifest-identity.md), whose
preparer rejects relative manifest paths:

```text
python3 source-candidate/tools/native_qualification/prepare_legacy_native_inputs.py --lane legacy-rtf --manifest driver/legacy-rtf-final.json --before source-baseline --after source-candidate --output driver/source-preparation.json
```

For a future re-preparation, first materialize the source endpoints and final
manifest under a new packet root.  Use the repaired qualification checkout
(the checkout containing `03eec6dbc651ec2d84d87b86c0f69245c38a58b4`) and a new
external root; do not point a new execution at the historical 691 packet:

```sh
QUALIFICATION_SOURCE=/home/zhuhe/code/litchi-worktrees/uo2-legacy-manifest-identity
RTF_PREPARATION_ROOT=(deleted-coord-artifact)
test ! -e "$RTF_PREPARATION_ROOT/driver/source-preparation.json"
python3 "$QUALIFICATION_SOURCE/tools/native_qualification/prepare_legacy_native_inputs.py" \
  --lane legacy-rtf \
  --manifest "$RTF_PREPARATION_ROOT/driver/legacy-rtf-final.json" \
  --before "$RTF_PREPARATION_ROOT/source-baseline" \
  --after "$RTF_PREPARATION_ROOT/source-candidate" \
  --output "$RTF_PREPARATION_ROOT/driver/source-preparation.json"
```

The explicit manifest binds candidate revision
`82798234ede3d77bc2a96a36d2da5d0c9aafaec2`, loaded successfully, and a build-free runner identity
round-trip matched both receipt identities.  The focused protocol command
completed with 44 tests passing:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_legacy_native_qualification tools.test_native_qualification_runner tools.test_no_3rdparty_references
Ran 44 tests
OK
```

No Cargo command, native execution, timing, allocation claim, or performance
result was run in this preparation packet.  The paired build/correctness and
quiet ABBA qualification remain the next authorized slot's work.
