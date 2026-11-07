# Agent instructions

## Documentation and commit messages

- **`FEATURE_MATRIX.md` files are feature checklists only.** Keep feature names
  and support/read/write status. Never add implementation details, API inventories,
  test details, validation results, progress notes, or verification narratives.
- **Never commit verification-process documents.** Do not add test reports,
  command logs, validation receipts, evidence summaries, or procedural write-ups.
  Keep verification artifacts and working notes outside the repository.
- **Briefly describe changes in commit messages.** Do not create verification
  documents or append implementation/test narratives to feature matrices.

## Repository size restrictions

GitHub blocks any file larger than 100 MiB, warns above 50 MiB, and rejects
pushes larger than 2 GB. Measurement evidence committed under
`docs/performance/results/` pushed this repository past those limits, so on
2026-10-05 the history of `feat/office-format-completeness` (including commits
already on GitHub) was rewritten to remove every evidence file of 1 MiB or
more. These rules keep the repository within the limits. They apply to every
branch, worktree and agent.

### Never commit

- **Any file of 1 MiB or more under `docs/performance/results/` or
  `docs/report/`.** This covers every kind of evidence: raw captures
  (`capture*/` directories, `perf.data`, heaptrack output, stack and unwind
  dumps), analysis and qualification JSON, comparison and report JSON, command
  logs, and source tarballs.
- **Benchmark or harness binaries, of any size.** Do not commit anything under
  a `binaries/` directory, or any other compiled executable. Record the source
  commit, the build command and the binary's SHA-256 outside the repository instead.

### Keeping large evidence

Keep verification artifacts and working notes outside the repository regardless
of size, for example in the session scratchpad or another directory outside the
checkout. Do not commit verification summaries or receipts; briefly describe
changes in commit messages.

### Check before every commit

Both commands must print nothing; otherwise unstage the files they list.

```sh
git diff --cached --no-renames --diff-filter=AM --raw --no-abbrev -- docs/performance/results docs/report |
  awk -F'\t' '{split($1, m, " "); print m[4], $2}' |
  git cat-file --batch-check='%(objectsize) %(rest)' |
  awk '$1 >= 1048576'
git diff --cached --name-only --diff-filter=AM | grep '/binaries/'
```

### After the 2026-10-05 rewrite

- Every commit from 2026-08-11 onward has a new hash. Change records still cite
  the old hashes; `docs/performance/history-rewrite-2026-10-05.commit-map`
  maps each old hash to its new one.
- Never merge or push a branch that still contains pre-rewrite commits; that
  would bring the purged files back. Rebase it onto the rewritten history
  first.
- Older checkouts may still hold the purged files on disk as untracked files.
  Do not add them back. `.gitignore` lists the purged command logs by name,
  because its evidence-log exception would otherwise re-include them.

## Test data: never reference 3rdparty/

`3rdparty/` holds untracked upstream checkouts (LibreOffice, Apache POI,
OfficeCLI and others). It is absent from worktrees, CI and other checkouts.

- **Never reference a file under `3rdparty/`** from code, tests, benches,
  examples, fuzz targets or tool scripts. This covers `include_bytes!`,
  `include_str!`, path joins and `../3rdparty/...` relative paths.
- **Copy the needed fixture into the tracked `test-data/` tree** and reference
  it there. Next to it, record where it came from: the upstream project, its
  path in that project and its license.
- The repository size rules above still apply. Keep fixtures small, and never
  commit a binary executable.
