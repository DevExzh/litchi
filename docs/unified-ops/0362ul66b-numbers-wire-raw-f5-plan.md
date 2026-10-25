# 0362ul66b — Numbers-wire scoped F5 handoff

Status: source is ready; F5 is planned but not started.  This receipt is for
the coordinator's first free heavy-build slot.  It does not claim a build,
test, clippy, rustdoc, native, or performance result.

## Source and lock hydration

The source under test is commit `2be299386f` on
`uo2/numbers-wire-raw-0362ul66`, the independent-review return for 0362ul66.
The three approved lock contents are identified by their SHA-256 values:

| lock | approved SHA-256 | use in this handoff |
| --- | --- | --- |
| workspace `Cargo.lock` | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` | required by the workspace checks |
| `tools/perf-baseline/Cargo.lock` | `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c` | hydrated and verified; no perf-baseline build is in this scoped run |
| `tools/native-resave/Cargo.lock` | `51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378` | hydrated and verified; no native-resave build is in this scoped run |

Hydration copies those exact approved files into the ignored workspace root
and the two tool directories, then runs `sha256sum` over all three.  The
approved source files are copied from an existing checkout; `cargo
generate-lockfile` and every unlocked Cargo invocation are forbidden.  The
lock copies are coordination inputs, not source changes and are not staged.
Before each build row, the runner must re-check all three hashes and fail
closed on a mismatch.

The runner is the approved bounded Cargo shim from record 0362ul65:

```text
/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/bin/cargo
sha256 495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d
```

Use a fresh target outside the worktree, with one build job, one test thread,
no incremental artifacts, debug information disabled, and the existing 20 GiB
memory / 18 GiB start / 19 GiB live / 20 GiB hard target guards.  Keep the
target and all logs under `.coord`; none are source evidence for this commit.

## Scoped build and test matrix

The actual slot runs only the Numbers-wire leaf and its affected iWork
consumers.  Every Cargo command is `--offline --locked`, uses the hydrated
workspace lock, and is run serially:

1. `litchi-numbers-wire`: all library and integration targets, then strict
   clippy (`--all-targets --all-features -- -D warnings`).
2. `litchi-numbers`, `litchi-pages`, `litchi-keynote`, and `litchi-iwa`: first
   `cargo check --all-targets --all-features` (the feature set is empty for
   `litchi-iwa`), then the matching all-target strict-clippy row.  This checks
   every migrated consumer without running native-resave or the unified
   facade.
3. Run the non-native focused integration tests for Numbers table cells,
   table data/list and tile readers; Pages table-cell/table-merge readers;
   Keynote table readers; and the IWA Numbers/Pages/Keynote handoff tests.
   The raw leaf test suite includes the independent input/item ceilings,
   exact and one-under admission, packed-item accounting, formula-owner
   admission, output preflight, and unknown-tail fidelity tests.

Fuzz manifests are compile-only rows, never fuzz executions.  Run a fuzz
`cargo check --locked` row only when that manifest already has an approved
lock.  The current four manifests have no checked-in lock, so they must be
reported as *not run* rather than allowing Cargo to create one.  No native
tool, OfficeCLI, `tools/native-resave`, `litchi-ops`, `litchi`, or full-
workspace row belongs in this F5 receipt.

## API and source checks

The slot closes with the staged source checks already used for 0362ul66:

- `git diff --check` and the repository size/binary guards;
- JSON parsing of `tools/api_glossary.json` and the glossary unit tests;
- the API glossary gate with the `numbers-wire` profile and the affected
  `numbers-minimal`, `numbers-full`, `pages`, `keynote-minimal`,
  `keynote-full`, and `iwa` profiles;
- the required-signature rows for every `BncCell`, `BncCellView`, pre-BNC,
  cell-value, table-reader, and table-writer entry point; and
- a tracked-consumer grep proving that no one-argument BNC parse/encode or
  decode call and no removed `*_with_limit` spelling remains.

The glossary check is run only against the complete hydrated workspace (the
sparse source checkout cannot build the facade rustdoc profile).  A successful
source check is not promoted to a Cargo result until every scoped row above
has a receipt tied to commit `2be299386f` and the three lock hashes.
