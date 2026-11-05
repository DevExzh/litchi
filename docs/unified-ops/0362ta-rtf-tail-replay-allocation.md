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


## Reproducible replacement receipt

The original temporary probe source was not retained, so its first receipt is
historical only. It is superseded by the following source-preserved rerun;
these are the numbers to use for the allocation conclusion. The new probe is
kept outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-probe.rs`
(6,682 bytes, SHA-256
`28be655f8e8c17b64e79eb3badf2671926103db755c809aaa9256c188bd628b7`). The
same source bytes were copied into both source snapshots (the copied files
also hash to that value, and `cmp -s` exited 0):

- baseline source snapshot `cbf4e30276d9d9566db6e0e2c141486ab8144fdc` at
  `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-alloc-baseline-src`;
- candidate source snapshot `28d4262acc` at the RTF replay worktree.

The probe accepts the appended text length as its only argument. It runs the
same four routes for a zero-byte append and a 1,048,576-byte append: authored
commit, changed replay, inverse replay, and exact no-op replay. It writes
output and deterministic durable JSON outside the checkout and checks the
forward and inverse bytes before printing the counters. The exact commands
were, once per source snapshot:

```text
cp /home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-probe.rs crates/litchi-rtf/examples/tail_replay_alloc_probe.rs
source /home/zhuhe/code/litchi-worktrees/.coord/w8-file/env.sh
CARGO_TARGET_DIR=<separate external target> CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2
RTF_PROBE_OUT_PREFIX=<external prefix>/small cargo run --offline --locked -p litchi-rtf --example tail_replay_alloc_probe -- 0
RTF_PROBE_OUT_PREFIX=<external prefix>/large cargo run --offline --locked -p litchi-rtf --example tail_replay_alloc_probe -- 1048576
```

Protected debug probe binaries were copied out of the generated targets:

- baseline: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-baseline-v1.bin`, 10,118,256 bytes, SHA-256 `25ad1cbda41ed02c9562a1af1661a9513fd4dcf2f9b5dc5c6a0e51acf96f6ae6`;
- candidate: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-candidate-v1.bin`, 10,119,480 bytes, SHA-256 `8b8578214ff543ba83c2508c55ce5b70f309ee88e4f481dae4e4d8fe76e6173b`.

The raw logs are retained outside the repository:

- baseline zero-byte: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-baseline-small.log`, 2,697 bytes, SHA-256 `1a52ea111dc93d38306c7c07cf5b58a9464d359c27eaa5b771214a4f881db202`;
- baseline 1 MiB: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-baseline-large.log`, 944 bytes, SHA-256 `eed5490ca5adf12db8057d7a3612c87a7af61e85297cb906345955b83339d495`;
- candidate zero-byte: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-candidate-small2.log`, 1,016 bytes, SHA-256 `b7f583be3262876936416b53a11ed67abf9ccfc7429e716e38773e1eb7fb79b2`;
- candidate 1 MiB: `/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-candidate-large2.log`, 933 bytes, SHA-256 `3e620d773b9c96ea7649f4ebd3868cf178c7f5992e4bee686df44c4551982d31`.

The new allocation results are:

| payload | route | baseline allocations / bytes | candidate allocations / bytes |
|---:|---|---:|---:|
| 0 | authored commit | 25 / 10,442 | 25 / 10,442 |
| 0 | changed replay | 1 / 30 | 0 / 0 |
| 0 | inverse replay | 1 / 30 | 0 / 0 |
| 0 | exact no-op replay | 0 / 0 | 0 / 0 |
| 1 MiB | authored commit | 29 / 10,501,324 | 29 / 10,501,324 |
| 1 MiB | changed replay | 1 / 1,048,606 | 0 / 0 |
| 1 MiB | inverse replay | 1 / 1,048,606 | 0 / 0 |
| 1 MiB | exact no-op replay | 0 / 0 | 0 / 0 |

Both snapshots produced byte-identical outputs. For the zero-byte route, the
authored output is 55 bytes (SHA-256
`1a097e857d51ddc8c81d461ee80cb2074ea3151f8da086dac008382721efd159`) and the
durable JSON is 362 bytes (SHA-256
`c8b3c7c3cc045895f85779a4edc780e9d07038222558e3e19d0dc2a930944f2f`). For the
1 MiB route, the authored output is 1,048,631 bytes (SHA-256
`fa8e9ed5f9b6e13647ec8347578d60c5251470078b6700403a1a49d4a5235080`) and the
durable JSON is 2,097,519 bytes (SHA-256
`bab6dbd2d804284ff155dea982dc005eca0ba292b7851aa1dc372da56b2f34ec`). Each
baseline/candidate pair was compared by exact SHA-256 equality.

The replacement preserves the earlier conclusion: retaining the inserted span
removes the replay allocation, including the 1 MiB case, without moving work
into authored commit or changing authored/durable bytes. No timing claim is
made by this allocation receipt.

The replacement output files are under
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-tail-replay-baseline/` and
`.../w8-rtf-tail-replay-candidate/` with suffixes
`small.authored.bin`, `small.durable.json`, `large.authored.bin`, and
`large.durable.json` (candidate uses `small2`/`large2` prefixes). `cmp -s`
returned 0 for each corresponding baseline/candidate authored and durable
pair.
