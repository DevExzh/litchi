# 0920 final source review

This record is a read-only review of the isolated 0920 CFB owned-output
candidate. It records source identity and the selected validation receipts; it
does not authorize production adoption or a performance capture.

## Read scope

The review read `docs/adr/README.md`, all non-iWork material in ADR 0001
through ADR 0027, and ADR 0030 through ADR 0032 in bounded, untruncated
reads. It read ADR 0008 through its non-iWork migration, verification, and
quality-gate material. The user-authorized iWork exclusions were ADR 0028,
ADR 0029, and the iWork amendment block of ADR 0024. No source reconstruction
was performed for the historical read-order issue.

The candidate source tree is `/tmp/litchi-performance-0920/source`; the
baseline tree is `/tmp/litchi-performance-0920-baseline/source`. The bound
candidate source receipt reports 942 files versus 938 baseline files, no
missing files, and no unexpected mismatches. Candidate-only files are the
owned-output test and the three 0920 packet metadata files. The reviewed
production changes are limited to:

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| `crates/litchi-cfb/src/writer/core.rs` | 124854 | `4c4b3c38533694a92e9a2d2dc4b119ef427f6ed5633f110e2d30b3fb74b4d368` |
| `crates/litchi-cfb/src/writer/layout.rs` | 96624 | `c8d33cfab4a4c38d05395dfe665553eea88f364e4a7a5486c6b8e64730037233` |
| `crates/litchi-ole-common/src/object/codec.rs` | 52878 | `24db2bec57e06627f31c2c68ea9f414b01fa95dfb3c8bbd6673b7d9045c47bd9` |
| `crates/litchi-cfb/tests/owned_output.rs` | 2516 | `e1c2b74059d7efd9d9854fb2b2b987a758dc9c356a8f73f4429a4880ef024127` |

The candidate patch is 9,665 bytes with SHA-256
`0441e3d739155ab4bbe05cf2dde02c837d10c48b3162352f6b3327055ab929e3`.
The candidate source receipt is 173,028 bytes with SHA-256
`b0fd362805cd6ba1abba0ff0a43ee52df871568183f6aac4802eadb641b31a4e`.
It binds base revision
`bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`, baseline archive
`869aa001cc81972c3881f972ea6d34fe47bccf2bc2dc2c90c3303662eb6a6d3f`, and
candidate source manifest
`c4e30450dd4e0f5f185c6b535da790f375063ab07c996fa168947d6564deda24`.

## Source conclusion

The CFB change is internally coherent. `write_to_vec` reserves the checked
accepted-reuse-plan length before emission, maps reservation failure to the
typed CFB allocation error, verifies the emitted length, and publishes the
sector-layout report only after verification. The common OLE codec delegates
its layout render through this owned path. Existing fallback and copy-through
routes remain separate. The focused owned-output tests cover generic output,
accepted source-layout reuse, and a reuse decline with byte equality.

No source defect requiring reconstruction was found. This is a conditional
source-level approval for comparative qualification after the custody and
freeze gates below. It is not a claim that the whole workspace is covered by
the selected tests.

## Build, graph, and binary receipts

The candidate build command was:

```text
taskset -c 14,15 cargo build --locked --offline -j2 --release --manifest-path /tmp/litchi-performance-0920/source/docs/performance/results/change-0917/probe/Cargo.toml --bins
```

It exited 0. The final candidate build receipt is 3,193 bytes with SHA-256
`03708890258129c636c380e9604154b22bb9dbb367058dd4c2e6435105411371`; its
start receipt is 1,611 bytes with SHA-256
`5ac421f85bba578498e56cf450a278e5f278e7553478d264df1ba2c9cdb1c6ed`. The
receipt binds probe lock
`abd9f47a5aaa973a760cd61392de79c2b6143026c300dc1329831652aa54c06c`, the
candidate patch above, and the candidate manifest above. The standalone
normalized probe graph has 144 equal nodes on both sides and SHA-256
`437357b509d1b8f377043c4de6510302f10dfa1e6a5fc38a3662d7d54df0c6c9`.

All binaries remain outside Git. The final candidate and baseline artifact
identities are:

| Lane | Candidate bytes / SHA-256 | Baseline bytes / SHA-256 |
| --- | --- | --- |
| Native | 48,098,944 / `4901f974ba32f8453399c4ddff31e9f08702fe2d03124152d8113427f54dece3` | 54,795,488 / `9f3c2bb2c589ba8e7a47b99c051d4a4910e41002d28f1caa1d1edc036a7515e5` |
| Allocator | 48,103,456 / `2f17dd59305e6395cc53c7c8cf4f9a33be78f2a74544d0946ebcecb196b270e9` | 54,800,008 / `20b96ebbdf611c22b7114b88580044a707d7fce2e4c8811c8084679feb1060dd` |

The baseline source receipt is 118,656 bytes with SHA-256
`4b7967901c42fccd89dcb4893f9147eaaad0f66106ef7be6ea4ac4112b074867`; its
build receipt is 2,603 bytes with SHA-256
`eb052a09753ef8c8b941c61df288107229735ec2ab81efcf2e081c89a5f03c47`.

The root parity receipt reports 157 matching fingerprint records for compiler,
profile, features, declared features, rustflags, config, and compile kind; it
also reports equal toolchains and environments apart from target directories.
The receipt is useful evidence, but the freeze path must bind the full required
compiler/profile/flags/allocator/normalized-dependency descriptor before
capture.

## Selected validation receipts

These are scoped receipts produced in isolated external quality workspaces;
the review did not rerun them:

| Receipt command | Exit and reported totals |
| --- | --- |
| `cargo test --offline -j1 -p litchi-cfb --features write` | 554 passed, 0 failed, 2 ignored |
| `cargo test --offline -j1 -p litchi-ole-common` | 285 passed, 0 failed, 2 ignored |
| `cargo test --offline -j1 -p litchi-doc` | 1361 passed, 0 failed, 13 ignored |
| Focused final CFB tests, `final-29` through `final-36` | all recorded exit 0; includes owned output, emission order, layout smoke/metadata, common check, formatting, and owned-library selections |

The selected quality-summary receipt is 13,573 bytes with SHA-256
`ba08867169de85359c614e5f25c4a377d80af176db061de634c126076db150ae`. Its
source-hash receipt binds the four candidate files above. The candidate
functional qualification reports 48/48 children passing with zero warmups and
one sample, and unchanged source-manifest identity. Its raw candidate result
is 394,714 bytes with SHA-256
`0ad6f64b946fbbf19a6c142d2db20d47c786c66a4a2df0c4cb06402c73f30e3f`; the
baseline raw result is 394,117 bytes with SHA-256
`b33110b585fdb65a0b24ea31f6cf834a113687939f5d2e349b38f60750e42ba5`. These
are functional and identity receipts, not performance statistics or full
workspace coverage.

## Freeze disposition and next gate

The protocol remains `prospective`; `capacity-controls.json` reports seven
negative controls rejected, the legitimate smaller-capacity control accepted,
53,713 bytes for the planned capacity summary, and zero performance children.
No comparison manifest or performance child has been frozen or run.

The protocol requires review disposition, capacity-control custody, and build
parity before freeze. The current Python runner has the Git-revision validator,
but its freeze/load path does not yet consume every control/reviewer receipt or
the complete build-parity descriptor. Root can enforce those checks through a
pre-freeze review/receipt gate or tooling changes; no ADR requires that Python
be the sole enforcement mechanism. The review record itself was added after
the candidate-source receipt and must be hash-bound into the final packet if
the root chooses to freeze this candidate.

Only after the comparative adoption gates and applicable quality checks pass
may root apply this exact validated candidate to the production checkout. No source rewrite is required merely to repair chronology.
Until then, reuse the exact external source/build artifacts and do not claim a
comparative performance result or production adoption.
