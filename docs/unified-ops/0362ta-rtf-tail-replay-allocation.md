# 0362ta — RTF tail replay span storage

This slice removes the second retained copy of a logical tail patch's
inserted bytes.  A native `tail_append::Patch` now records the checked
`root_close` offset and inserted length, borrowing the bytes from its retained
before/after package when it is converted to the durable wire patch.  The
durable patch still owns its `Box<[u8]>`, so its encoding, bounds, inverse and
source-digest checks are unchanged.  The inverse swaps the retained package
handles and copies only the span metadata.

The existing exact-source logical and durable round-trip test passed unchanged;
it covers forward, inverse, deterministic wire bytes and a foreign-source
refusal.  The coordinator's separate RTF inverse diagnostic correction is not
duplicated here.

## Allocation evidence

The matched temporary probe used a one-megabyte appended paragraph, the same
limits and fixture for the pre-change `cbf4e30276` source and this candidate.
It checked authored output and durable JSON bytes as well as changed, inverse
and no-op replay.  The source diff used by the candidate is kept outside the
repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-alloc-candidate.diff`
(3,616 bytes, SHA-256
`2c0fd086b4469cad64e1c139021d99bfe8505de9ab962b0697a32daa5b106e11`).

| route | baseline allocations / bytes | candidate allocations / bytes |
|---|---:|---:|
| authored commit | 22 / 9,451,376 | 22 / 9,451,376 |
| changed replay | 1 / 1,048,606 | 0 / 0 |
| inverse replay | 1 / 1,048,606 | 0 / 0 |
| exact no-op replay | 0 / 0 | 0 / 0 |

The authored output was 1,048,625 bytes with SHA-256
`85ab6f00ad070dc60eddc632118bb067b9dfe2969212946212ba072ef8c9ab3d` in both
runs.  The durable JSON was 2,097,519 bytes with SHA-256
`1cd14e2103724a4f217c46ece09dbe847f2edf7b87a39edc03a4c60438e6763a` in both
runs.  The exact-byte check was `cmp -s` between the baseline and candidate
authored files and between the baseline and candidate durable files; both
returned exit 0.  The files are
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-baseline-large.authored.bin`,
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-candidate-large.authored.bin`,
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-baseline-large.durable.json`
and
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-candidate-large.durable.json`.
Raw pre-change and candidate logs, source trees, binaries and these exact
matching files remain outside the repository:

* baseline log: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-alloc-baseline-large-v2.txt`, 3,624 bytes, SHA-256
  `b0928d517c55f0a33fe84fcc05b2a8dfaa666aafee8addddf9e9f9e93fac2311`;
* candidate log: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-alloc-candidate-large-v2.txt`, 3,583 bytes, SHA-256
  `8b5157df1dceb70415c05c80d5392551f92040b04a0a8745b3c734d7cd121295`;
* candidate binary: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-alloc-candidate-large-v2.bin`, 8,881,200 bytes, SHA-256
  `5f7e81108de9be25b7d5559a8f070b6591ccb6c21665808f12d2551d89cbb363`;
* baseline binary: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-alloc-baseline-large-v2.bin`, 8,879,808 bytes, SHA-256
  `d957b53eec59ad072e5cde3f41ca48957dad5f31db77574342444ed5a64bc171`.

The command for both source snapshots was:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/w8-file/env.sh
CARGO_TARGET_DIR=<separate external target>
CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2
cargo run --offline --locked -p litchi-rtf --example tail_replay_alloc_probe
```

No timing or binary is part of this commit.

## Validation

With the capped environment, these commands passed:

```text
cargo test --offline --locked -p litchi-rtf --test logical_tail_append --test source_backed_tail_append -- --nocapture
cargo test --offline --locked -p litchi-rtf --all-targets
cargo clippy --offline --locked -p litchi-rtf --all-targets --all-features -- -D warnings
cargo fmt --all -- --check
```

The all-target test log is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-alloc-tests.log` (162,109
bytes, SHA-256
`9dd1c49e7921603652e43067c6197fc7f7b86df4e5ff142fb75eb853227ac367`).
The all-feature RTF Clippy log is
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-clippy.log` (1,467
bytes, SHA-256
`15488d3ebbd22f13b5b8abdc7e33069db8e82ea884e805ac63127a65af2a6e94`).
Fresh hidden-item RTF rustdoc JSON was generated with warnings denied and is
outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-litchi_rtf-rustdoc.json`
(49,776,372 bytes, SHA-256
`20374a4543f8ff82711bd4d595eb4ee391fd8f5adccc56fddbe4730a21f73542`).

