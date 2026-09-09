# 0488 disk cleanup

The user requested removal of unused intermediate files and worktrees, with
`/home/zhuhe/code/litchi-spec-gaps` explicitly protected because other agents
are working there. That workspace and associated spec-gap build directories
are excluded. Other projects and active work are also excluded.

## Rebuildable outputs

[The cache receipt](build-cache-cleanup.json) records removal of the main
workspace's `target` and `tools/perf-baseline/target`: 503,605,403,648 allocated
bytes (469.02 GiB). These contained Cargo compiler objects, incremental state,
example/test executables, release duplicates, rustdoc output, and sanitizer
build cache. No tracked files were present. Immediately before each removal,
`/proc` checks found no process using those roots as a working directory,
executable, open file, or memory mapping. No process was stopped.

The retained benchmark/fuzz executables in `~/.cache/litchi-goal-*`, source
manifests, captures, and reports remain. The [0487 seal verification after
cleanup](0487-seal-verification-after-cleanup.json) passes for all 144 formal
processes, 4,320 samples, 12 diagnostic children, and two fuzz campaigns.
Earlier raw Cargo output paths are rebuildable locations, not permanent binary
custody. Rebuilding is required before running workspace tests again.

## Temporary worktrees

The [read-only audit](worktrees-audit.json) identified clean, inactive temporary
checkouts whose commits remain reachable from existing branches. All eleven
selected worktrees were removed with `git worktree remove` without force;
[per-worktree receipts](worktree-removals.json) record their heads, retaining
refs, status, process checks, and command results. No branch was deleted.
These removals reclaimed another 12.25 GiB, including the old
DOCX streaming candidate's ignored build output. Dirty, untracked, active,
unreachable detached, and potentially spec-gap-associated trees remain.

Total allocated storage removed was 481.26 GiB across
the root filesystem and `/tmp`. [The final summary](summary.json) records both
filesystems and the unchanged hashes of unrelated local files. Root filesystem
free space is about 875 GiB; other sessions continue
working, so subsequent filesystem usage can change independently.

## Ongoing cleanup policy

Future performance batches should retain source revisions/manifests, recipes,
reports, and only the executables actually bound by retained evidence. Failed
scratch builds and replaced build outputs should be removed after jobs finish.
Shared Cargo caches are disposable and should be bounded by periodic cleanup;
retaining them indefinitely accumulated hundreds of GiB. Worktree removal must
preserve dirty files and every commit not otherwise reachable from a branch.
The full non-iWork performance goal remains active.
