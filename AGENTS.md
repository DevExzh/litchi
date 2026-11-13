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
pushes larger than 2 GB. Measurement evidence committed to this repository
once pushed it past those limits. The history of
`feat/office-format-completeness` was rewritten on 2026-10-05 to purge the
large files, and the evidence trees `docs/performance/`, `docs/report/` and
`docs/unified-ops/` were then deleted. These rules apply to every branch,
worktree and agent.

### Never commit

- **The evidence trees.** Do not re-create `docs/performance/`, `docs/report/`
  or `docs/unified-ops/`, and do not link to them or read from them.
- **Verification or evidence documents, anywhere.** This covers change
  records, progress logs, test reports, receipts, command logs, and analysis,
  qualification, comparison or report JSON.
- **Raw captures.** This covers `capture*/` directories, `perf.data`,
  heaptrack output, stack and unwind dumps, and source tarballs.
- **Benchmark or harness binaries, of any size.** Do not commit anything under
  a `binaries/` directory, or any other compiled executable.
- **Any other new file of 1 MiB or more** outside `test-data/`. Keep fixtures
  under `test-data/` small as well.

Keep verification artifacts and working notes outside the repository, for
example in the session scratchpad or another directory outside the checkout.
Briefly describe changes in commit messages.

### Check before every commit

All three commands must print nothing; otherwise unstage the files they list.

```sh
git diff --cached --name-only --diff-filter=AM -- docs/performance docs/report docs/unified-ops
git diff --cached --name-only --diff-filter=AM | grep '/binaries/'
git diff --cached --no-renames --diff-filter=A --raw --no-abbrev -- . ':!test-data' |
  awk -F'\t' '{split($1, m, " "); print m[4], $2}' |
  git cat-file --batch-check='%(objectsize) %(rest)' |
  awk '$1 >= 1048576'
```

### History before the 2026-10-05 rewrite

- Every commit from 2026-08-11 onward has a new hash since the rewrite.
- Never merge or push a branch that still contains pre-rewrite commits; that
  would bring the purged files back. Rebase it onto the rewritten history
  first.
- Older checkouts may still hold purged or deleted evidence files on disk as
  untracked files. Do not add them back.

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
