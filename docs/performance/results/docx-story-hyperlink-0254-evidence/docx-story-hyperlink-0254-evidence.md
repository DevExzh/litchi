# DOCX story-hyperlink planning: controlled release ABBA evidence

This evidence compares the pre-index control (`ecbc8cdb5` plus the committed
benchmark harness `9f2350e39` cherry-picked as `dbf33bdc0`) with candidate
`9f2350e39`.  It is limited to the committed opt-in case
`docx_story_hyperlink_plan` and its deterministic `49-stories-1152-links`
corpus.

## Reproduction

The control and candidate were built once in fresh worktrees under
`/home/zhuhe/CodeProjects/litchi-worktrees`, sharing
`docx-story-hyperlink-0254-target`:

```text
taskset -c 0 env CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/litchi-worktrees/docx-story-hyperlink-0254-target CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=1 cargo build --manifest-path tools/perf-baseline/Cargo.toml --release --bin litchi-perf-baseline
```

The control worktree was created at `ecbc8cdb5`, then ran
`git cherry-pick 9f2350e39`, producing `dbf33bdc0`.  The candidate worktree was
created at `9f2350e39`.  Both worktrees reported clean.  The copied immutable
executables were verified before every leg:

| implementation | revision | bytes | mode | SHA-256 |
| --- | --- | ---: | --- | --- |
| control | `dbf33bdc0fa381c531fb8a2eda92f388e9a20a49` | 45,459,504 | `0555` | `ec2d0e73f5f449315197e664412375f04f27737954571cbb1db37fbdfee4fd3d` |
| candidate | `9f2350e393da9a8e3ce98892cf2abc32dda9a56d` | 45,523,880 | `0555` | `1a4e2bdc0ff1720173054ed5f61fbfa06b4e239416086a175e8a6f8c2b4ff67b` |

Each leg used this command shape, with the corresponding immutable executable
and report path:

```text
taskset -c 0 env CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/litchi-worktrees/docx-story-hyperlink-0254-target CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=1 <immutable-binary> --case docx_story_hyperlink_plan --warmup 20 --samples 500 --json <report>
```

The four legs were run in order A1, B1, B2, A2.  Rust 1.95.0, one permitted
CPU (`0`), one Cargo build job, and `CARGO_INCREMENTAL=0` were used.  Free
space stayed above the 10 GiB stop threshold: 25,783,263,232 bytes before the
control build, 21,870,080,000 after it, 21,881,217,024 before the candidate
build, and 23,984,619,520 after it.

## Results

All values are nanoseconds for eight repeated planning calls per retained
sample; each leg has 500 retained samples after 20 warmups.

| leg | implementation | p50 | mean | p95 | p99 |
| --- | --- | ---: | ---: | ---: | ---: |
| A1 | control | 764,218 | 774,770.294 | 817,544 | 1,028,599 |
| B1 | candidate | 66,008 | 69,743.670 | 83,072 | 107,655 |
| B2 | candidate | 64,767 | 67,583.676 | 82,081 | 104,441 |
| A2 | control | 770,355 | 779,340.766 | 848,527 | 887,040 |

Candidate reduction was positive in both directions for every reported
statistic: A1→B1 was 89.53–91.36%, and A2→B2 was 88.23–91.59%. Same-path
drift stayed within the default ceilings (p50/mean 5%, p95 10%, p99 15%) for
both implementations. The gates therefore accept all four statistics, but
the only supported claim is the scoped prepared-snapshot planning result;
this is not an end-to-end DOCX, I/O, allocation, RSS, or general speedup
claim.

The corpus identity and output SHA-256 were equal across all legs.  The
strict generic `tools/perf_abba_summary.py` could not be used: the current
checkout version requires a `binary_identity` field absent from these
committed-harness reports, the older control version lacks the fixed custom
corpus exception, and the DOCX source object includes `plan_ns`, which is
measured data rather than source identity.  `docx-story-hyperlink-0254-summary.json`
is consequently a fail-closed custom recomputation from the retained raw
samples using the harness's Rust p50/nearest-rank/Welford rules.

`source_immutability_verified` is retained as a harness flag, but is not
treated as independent proof because the harness checks its own generated
corpus/source path.  The independently retained corpus, output, binary, and
raw-report hashes are the stronger identity evidence.

## Artifacts

- [`docx-story-hyperlink-0254-summary.json`](docx-story-hyperlink-0254-summary.json)
  and its `.zst` form contain the custom ABBA summary and gates.
- `docx-story-hyperlink-0254-a1.json`, `b1.json`, `b2.json`, and `a2.json`
  are the untouched raw harness reports; each also has a `.zst` copy.
- The two `.bin` files are the mode-0555 executables used for the legs.
- [`docx-story-hyperlink-0254-manifest.json`](docx-story-hyperlink-0254-manifest.json)
  records every artifact's byte count, mode, and SHA-256.

The dedicated worktrees and 1.4 GiB Cargo target were removed after these
artifacts were preserved; the repository's pre-existing unrelated dirty files
were not touched.
