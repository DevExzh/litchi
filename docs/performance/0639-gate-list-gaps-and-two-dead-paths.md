# 0639: the two suites no per-crate gate could see now have modes CI runs, the CRUD demo's PPTX update takes the route that works, and a dead whole-input slurp is gone

Status: retained, gate and correctness work. `performance_claim: none` — this
record carries gate coverage, one example fix and one deletion, not a
claim-registry entry. **No production read path changed**: the only file under
`crates/` that is not an example is `crates/litchi/src/sheet/workbook_types.rs`,
and the 66 lines removed from it had no caller.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Three items the first wave left behind, taken together because each is small
and none is a measurement:

- **(a)** Change [0619](0619-harness-xls-lifecycle-assertion.md) found a test in
  `tools/perf-baseline` red for 54 records, and change
  [0629](0629-facade-docx-budget-test-bisect.md) found a facade test red for 94
  commits. Both went unobserved for the same structural reason: no per-package
  selection in the standing gate list reaches either suite. 0629's own
  `GOAL_AUDIT` paragraph says so — "A feature-bearing `cargo test -p litchi`
  belongs in the standing gate list."
- **(b)** Change [0607](0607-pptx-authored-slide-regeneration-design.md) recorded
  as *Defect observed (report only)* that
  `crates/litchi/examples/office_crud_demo.rs` performs its PPTX UPDATE through
  `Package::open` then `presentation_mut()`, which its own census shows refuses
  on all 78 corpus decks and on litchi's own output.
- **(c)** Change [0587](0587-remaining-opportunity-survey.md)'s CORE section
  records `refine_workbook_format` as "a public, dead, whole-input slurp of any
  `Read` — a latent GOAL-hypothesis-1 ingress if it is ever wired".

## What was changed

### (a) Two single-command gate modes, wired into the job CI already runs

`tools/non_iwork_gate.py` gains two modes, `facade-format-tests` and
`harness-tests`, and `.github/workflows/rust-ci.yml`'s existing
`non-iwork-release-gate` job gains one step for each, placed after
`print`/`verify` and **before** the 45-root `check` sweep so a stale
expectation fails in minutes rather than after the long modes.

| Mode | Command (argv, no shell) | What it reaches that nothing else did |
| --- | --- | --- |
| `facade-format-tests` | `cargo test --package litchi --no-default-features --features docx,pptx,xls,xlsx --lib --tests --no-fail-fast -- --test-threads=1` | The facade's feature-gated tests. `litchi`'s `default = []`, so `cargo test -p litchi` compiles almost none of them. |
| `harness-tests` | `cargo test --manifest-path tools/perf-baseline/Cargo.toml --lib --tests --no-fail-fast` | The performance harness's own suite. `tools/perf-baseline` is a separate Cargo project with its own `[workspace]` table, so no `--workspace`, `--package` or exclusion selection in this gate reaches it. |

`docx,pptx,xls,xlsx` is the smallest feature set that compiles the DOCX, XLSX,
PPTX and XLS facade tests. Each leaf pulls its own substrate through Cargo's
feature graph — `docx = ["dep:litchi-docx", "opc"]`,
`pptx = [..., "opc", "ooxml-common", "drawingml"]`,
`xlsx = [..., "opc", "ooxml-common", "drawingml", "sheet"]`,
`xls = [..., "cfb", "ole", "sheet"]` — so naming the substrate features
separately would add nothing. Of the 2,331 `feature = "…"` `cfg` sites under
`crates/litchi/src` and `crates/litchi/tests`, these four account for **998**
(docx 417, pptx 247, xlsx 235, xls 99); every other name is a format or
capability outside the brief's four.

`--no-fail-fast` is deliberate on both: Cargo stops after the first failing
target by default, which silently skips every later test binary. The
`cargo test` job in the same workflow already carries that flag and a comment
saying a facade integration failure had been invisible for exactly that reason.

The two modes differ in one flag. `facade-format-tests` passes
`-- --test-threads=1`, as `lib-tests` and `doc-tests` do; `harness-tests` does
not. The workspace modes serialize because one invocation over 45 roots keeps
every test binary alive until the final link, and a standalone project has no
such fan-out; `perf-baseline.yml` has always run this suite at Cargo's default
parallelism; and the measurement below puts the cost of serializing it at
1,869.73 s against 581.14 s for the same 531 passing tests. Build parallelism
stays bounded by the gate's `CARGO_BUILD_JOBS` invariant either way.

Six unit tests in `tools/test_non_iwork_gate.py` (51 → 57) pin the two modes:
the exact argv of each, that both are deterministic, that the four features are
sorted, are declared by `crates/litchi/Cargo.toml` and are all in the plan's
**safe** facade closure, that `default = []` still holds, that the harness
manifest path is relative and present, that `harness-tests` raises `GateError`
when the manifest is absent, that the harness manifest declares no iWork
dependency (the condition under which its suite belongs in *this* gate), and
that `rust-ci.yml` invokes both modes with a `--record-file` before it invokes
`check`.

### (b) The demo's PPTX update takes the opened-package route

`crates/litchi/examples/office_crud_demo.rs` replaces
`PptxPackage::open(..)` + `presentation_mut()` + `add_slide()` with
`opened_presentation_transaction()` + `set_shape_text` + `add_text_box` +
`commit()` + `apply_opened_presentation_commit()`, then saves and reads the
result back so the demonstration proves the edit landed. A comment says why
`presentation_mut` is not the route for an opened deck. `Transaction` has no
`add_slide`, so the step demonstrates the two edits the route does own —
retitling the closing slide and appending a closing note to it — instead of
adding a fourth slide.

### (c) `refine_workbook_format` deleted

66 lines leave `crates/litchi/src/sheet/workbook_types.rs`: the 48-line
function, its 15-line unit test `refinement_restores_a_nonzero_cursor`, the
two-line `use litchi_core::FileFormat;` import that only it used, and one blank
line `rustfmt` then reclaimed. Nothing else in the file moved.

## Why it is sound

### The two modes are additions, not redefinitions

Neither mode changes an existing mode's argv, the workspace plan, the facade
feature closure, the dependency-tree checks, the execution recorder, the
environment invariants or the report schema. `command_specs` returns early for
each new mode before any bulk-root or facade selection is built;
`test_non_lib_modes_are_per_root_deterministic_and_use_exact_flags` and
`test_lib_tests_serializes_all_bulk_roots_before_facade` still pass unchanged,
which is the statement that `check`, `clippy`, `doc`, `lib-tests`, `doc-tests`
and `deprecated` are byte-identical in their generated commands.

`harness-tests` runs a manifest outside the workspace, which is the point: that
is the slice no selection reaches. It stays inside this gate's remit only while
the harness itself is non-iWork, so a unit test parses
`tools/perf-baseline/Cargo.toml`'s `[dependencies]` and fails if any name is an
iWork package or carries the `litchi-iwa` prefix. The manifest path is relative
and commands run with the repository root as their working directory
(`_execute_command` passes `cwd=ROOT`), so no checkout location is encoded. A
missing manifest raises `GateError` rather than handing Cargo a path that would
fail obscurely.

### The demo's new route is the one the package supports, and it is a delta

`presentation_mut()` refuses an opened package by construction:
`mutable_pres` is `Some(..)` only in `Package::new()`, and every byte-reading
ingress funnels through `from_opc_package_with_provenance`, which sets it to
`None` (change 0607, §1: 78 of 78 corpus decks and 2 of 2 litchi-authored decks
refused). `opened_presentation_transaction()` composes an immutable snapshot,
stages the edits, and `apply_opened_presentation_commit` publishes one patch
that adds, removes or re-blobs exactly the parts the patch names and leaves
every other part's name, content type, relationships and blob `Arc` untouched.
The member census below shows that on this deck: one part of 43 differs.

Error identity is untouched. The refusal `presentation_mut` raised is still
raised for the same reason by the same code; the example simply no longer asks
for it. No library file changed for this item.

### Deleting the helper removes a latent ingress and no behaviour

`refine_workbook_format` read its whole input into a `Vec<u8>`
(`reader.read_to_end(&mut data)?`) with no bound at all, then ran three
detectors over the buffer. That is the shape `docs/GOAL.md` hypothesis 1 names
and ADR 0003's bounded-source direction argues against, and it is why change
0587 flagged it. It had **no caller**: a workspace-wide grep over `*.rs` finds
the definition and one test invocation, nothing else, and the compiler agreed —
the item carried `#[allow(dead_code, reason = "kept as a detection helper; …")]`,
which only silences a lint that fires when nothing reaches an item.

It was also never public. `crates/litchi/src/sheet/mod.rs` declares
`mod workbook_types;` with no visibility modifier and re-exports nothing from
it, so no path outside the crate could name the function; its `pub` was
crate-internal spelling, not API. Deleting it therefore removes no documented
item, which the rustdoc index comparison below confirms exactly.

The sibling `detect_workbook_format_from_signature` is deliberately left alone.
It is dead too, but it is *bounded* — it reads at most eight bytes, classifies
the signature, restores the caller's cursor and never buffers the input — so it
is not the ingress 0587 named, and touching it is outside this change's scope.

### Contracts untouched

No typed error, limit, refusal point, validation order, output byte, public API,
fixture, corpus, selector, JSON schema or existing gate function is reachable
from any of the three items. No `unsafe`, no dependency, no Rayon pool, no
ambient I/O and no archive type, lock or executor crosses a boundary here.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
cargo 1.95.0. Base `c7326f680`, branch
`perf/0639-gate-list-gaps-and-two-dead-paths`, worktree
`/home/zhuhe/code/litchi-worktrees/0639`. **No performance measurement was
taken and no A/A floor applies**: there is no latency, throughput, allocation
or peak-RSS figure of any library operation anywhere in this record. The brief
asks for counts confirming that no production read path changed, and that is
what follows. Two wall times appear, both of a *test suite* and both only to
justify one gate flag; they are labelled where they occur and nothing else in
this record depends on them. The host carried eight concurrent agents
throughout, which affects those two wall times and nothing else.

### The two new modes, at the base and on this branch (measured)

Each leg is one Cargo invocation with the gate's own environment
(`CARGO_TARGET_DIR=target/non-iwork-gate`, `CARGO_BUILD_JOBS=1`,
`CARGO_INCREMENTAL=0`, `CARGO_PROFILE_DEV_DEBUG=0`,
`CARGO_PROFILE_TEST_DEBUG=0`). The base has no gate modes, so the base leg runs
the identical argv directly; that argv is the one `command_specs` emits, pinned
by the unit tests.

| Mode | Leg | Targets | passed / failed / ignored | Exit |
| --- | --- | ---: | ---: | ---: |
| `facade-format-tests` | base `c7326f680` | 25 | **231 / 0 / 0** | 0 |
| `facade-format-tests` | this branch | 25 | **230 / 0 / 0** | 0 |
| `harness-tests` | base `c7326f680` | 18 | **531 / 0 / 1** | 0 |
| `harness-tests` | this branch | 18 | **531 / 0 / 1** | 0 |

**Both modes are green at the base**, so neither reveals a pre-existing failure
to attribute: the two suites changes 0619 and 0629 repaired are still repaired,
54 and 94 records later, and this change adds no red. The facade delta is
exactly **−1**, and it is the one test deleted with `refine_workbook_format`;
nothing else moved. The harness totals are identical in both legs, which is the
statement that deleting the helper and rewriting the example reach nothing the
harness measures.

For contrast, the run this gap allows today, on the same branch:

| Command | Result |
| --- | --- |
| `cargo test --locked --offline -p litchi` (default features) | exit 0, 20 targets, **0 passed, 0 failed, 0 ignored** — `default = []` compiles none of them |
| `python3 tools/non_iwork_gate.py facade-format-tests` | exit 0, 25 targets, **230 passed** |

**Test-thread policy** (measured, and the reason the two modes differ).
`facade-format-tests` serializes like every other test mode in this gate; its
whole suite executes in 1.49 s, so serialization costs nothing. `harness-tests`
does not, and the figure is why: the same 531 passing tests over the same 18
targets took **1,869.73 s** of test execution serialized and **581.14 s** at
Cargo's default parallelism — 3.22×, for a project with none of the 45-root
link fan-out the bulk modes serialize against, and whose own workflow has always
run it unserialized. Both figures are summed `finished in …` per target, so
neither includes build time. Build parallelism stays bounded by the gate's
`CARGO_BUILD_JOBS` invariant either way. These two numbers are wall times and
are reported only to justify a flag; they are not a performance result, they
were taken on a host running eight agents, and nothing in this record depends on
them being repeatable.

### The demo, before and after (measured)

`cargo build --locked --offline --example office_crud_demo --features docx,pptx,xlsx,ooxml-common`,
run in an empty directory so every output file is this run's.

| Leg | Exit | Last line |
| --- | ---: | --- |
| Base `c7326f680` | **1** | `Error: UnsafeEdit { operation: "presentation_mut", reason: "the lossless facade cannot hydrate an opened package into the mutable writer" }` |
| This branch | **0** | `✓ All operations completed successfully!` |

The base leg writes five of the six files the demo promises and never reaches
`demo_presentation_updated.pptx`; DOCX and XLSX complete, PPTX completes CREATE
and READ and dies at UPDATE. That is change 0607's prediction reproduced
verbatim, 32 records later. The branch leg writes all six and reports
`Closing slide now has 4 shapes`, read back from the saved file.

**Member census of the update** (`demo/member-delta.txt`), comparing
`demo_presentation.pptx` with `demo_presentation_updated.pptx`:

| | |
| --- | ---: |
| Members before / after | 43 / 43 |
| Member order identical | yes |
| Members byte-identical | **42** |
| Members changed | **1** (`ppt/slides/slide3.xml`, 1,584 → 2,154 bytes) |
| Members added / removed | 0 / 0 |

Both edited strings are present in the one changed part. This is the per-part
delta change 0607 described for `opened::patch::apply`, shown on a real save.

### The deletion (measured)

| | |
| --- | ---: |
| Lines removed from `crates/litchi/src/sheet/workbook_types.rs` | **66** (48 function, 15 test, 2 import, 1 blank) |
| Lines added there | 0 |
| Callers of `refine_workbook_format` outside its own test, workspace-wide | **0** |
| Unbounded `read_to_end` sites removed from the facade | **1** |
| Public rustdoc item pages of `litchi` before / after | **8,086 / 8,086**, diff empty |

### No production read path changed (measured)

| | |
| --- | ---: |
| Files under `crates/` changed | **2**: one example, one source file |
| Production functions added, moved or rewritten | **0** |
| Production functions deleted | **1**, with zero callers |
| Harness test verdicts differing between the legs | **0** of 531 |
| Facade test verdicts differing between the legs | **1** of 231, the deleted test |
| Public rustdoc item pages differing | **0** of 8,086 |
| Bytes of any produced package differing, other than the demo's own edit | **0** (42 of 43 members byte-identical; the 43rd is the edit) |

The example and the deleted helper are on no measured path, which is why this
record takes no timing: `tools/perf-baseline` has no selector that runs
`office_crud_demo`, and `refine_workbook_format` had no caller to run. The
harness's own 531 verdicts, identical in both legs, are the check that this is
true rather than assumed.

### Evidence tiers

**Measured**: both gate legs' per-target totals and exit statuses on the base
and on this branch; the 51 → 57 unit-test count; both demo runs, their exit
statuses and their output-file inventories; every number in the 43-member
census; the 0-caller, 66-line and `read_to_end` counts; the 8,086-page rustdoc
comparison and its sha256s; and the two test-suite wall times that justify the
`harness-tests` thread flag. **Modelled**: nothing. **Unknown**: whether either
newly reachable suite is adequate; whether the four format leaves stay the
smallest sufficient set as the facade's `cfg` sites change; and whether
`detect_workbook_format_from_signature` should exist at all.

## Correctness evidence

`crates/litchi` is the only crate touched — one example and one source file —
so no other crate's `clippy`, `test` or `doc` gate is in scope. Tails are in
[`results/change-0639/gates.txt`](results/change-0639/gates.txt).

| Gate | Result |
| --- | --- |
| `cargo fmt --all --check` | exit 0, no diff |
| `cargo clippy --locked --offline -p litchi --all-targets` (default features) | exit 0; three pre-existing `dead_code` warnings in `tests/unexpected_format.rs` — the set change 0629 recorded citing change 0621 |
| `cargo clippy --locked --offline -p litchi --all-targets --no-default-features --features docx,pptx,xls,xlsx` | exit 0; nine pre-existing warnings in `detection_smart/detected.rs` and `document/doc.rs`, no new one. **Neither run mentions `office_crud_demo.rs` or `workbook_types.rs`**, and this feature set builds the example, so the rewritten step is linted. |
| `cargo doc --locked --offline -p litchi --no-deps` | exit 0, no rustdoc warning |
| `cargo doc --locked --offline -p litchi --no-deps --no-default-features --features docx,pptx,xls,xlsx` | exit 0 |
| `cargo test --locked --offline -p litchi` (default features) | exit 0, 20 targets, **0 passed, 0 failed, 0 ignored** — the gap itself, reproduced |
| `python3 -m unittest tools.test_non_iwork_gate` | exit 0, **57 tests OK** (51 at the base) |
| `python3 tools/non_iwork_gate.py print` | exit 0; 64 workspace packages, 18 excluded, 45 bulk, 35 safe facade features, 10 unsafe — identical to the base |
| `python3 tools/non_iwork_gate.py verify` | exit 0; `verified_bulk_tree_roots=45`, `verified_facade_safe_trees=35`, `verified_facade_combined_tree=1` |
| `python3 tools/non_iwork_gate.py facade-format-tests` | exit 0; 25 targets, **230 passed, 0 failed** |
| `python3 tools/non_iwork_gate.py harness-tests` | exit 0; 18 targets, **531 passed, 0 failed, 1 ignored** |
| The same two argv on the untouched base tree | exit 0 and exit 0; **231 / 0 / 0** and **531 / 0 / 1** |
| `python3 -m unittest tools.test_check_example_targets` | exit 0, 12 tests OK |
| `python3 tools/check_example_targets.py` | **exit 1, pre-existing** — see below |

### The one failing gate is pre-existing and is iWork's

`tools/check_example_targets.py` reports four cross-package duplicate example
target names and exits 1. Every duplicate is between `litchi-keynote`,
`litchi-numbers` and `litchi-pages`: `edit_chart_arrangement`,
`edit_image_adjustments`, `edit_table_cell_number_format` and `save_package`.
Those are iWork crates this change never touches and the brief excludes. It was
reproduced on the untouched shared before checkout at
`/home/zhuhe/code/litchi-worktrees/before-c7326f680` with the identical four
findings and the identical exit status, so it predates this branch. The example
this change edits, `office_crud_demo`, is not among them and its target name is
unchanged; the gate's own unit tests pass.

### Deterministic, correctness-bearing checks

- **The demo runs end to end**, exit 0, on a fixture it authors itself, and the
  saved package is read back inside the same run — `pres.slides()`,
  `slide.shape_count()` — so the demonstration fails loudly if the edit did not
  land. The base binary, built with the identical command, exits 1.
- **The member census** is a sha256 per member of both packages: 42 of 43
  byte-identical, order identical, nothing added or removed, both edited strings
  in the one changed part.
- **The public API comparison** builds the rustdoc of `litchi` with the four
  format features on the base and on this branch and diffs the generated item
  pages: **8,086 both sides, empty diff**.
- **Six new unit tests** pin the two modes' argv and invariants, including a
  fail-closed check that `harness-tests` raises `GateError` when its manifest is
  absent and a check that the harness manifest declares no iWork dependency.
- **The harness's own 531 verdicts are identical in both legs**, which is the
  check that no production read path moved.

## Validation preserved

Nothing about validation changed, because no validating code changed. The
XLS/XLSX/XLSB/ODS/Numbers signature classifier
`detect_workbook_format_from_signature` is byte-identical and still refuses a
truncated OLE2 prefix, still returns `Error::NotOfficeFile` for an unknown
signature and still restores the caller's cursor on both the success and the
failure path; its four unit tests are untouched and still run. The PPTX
package's refusals are byte-identical: `presentation_mut` still raises
`UnsafeEdit` for an opened package, `opened_presentation` still refuses a
package carrying a mutable model, and `apply_opened_presentation_commit` still
refuses a stale or foreign candidate graph. The release gate's existing modes
generate byte-identical commands, and its topology, feature-closure and
dependency-tree checks are unchanged, so nothing it refused before is admitted
now.

## Limitations

- **Nothing here makes any library operation faster**, and nothing here is a
  claim. No speedup, regression, allocation, peak-RSS, cold-cache,
  physical-I/O, range-source, concurrency or cross-platform result is stated,
  and no claim-registry entry is added. The one ratio in this record, 3.22×, is
  between two ways of running a *test suite* and is used only to pick a flag.
- **The two wall times are single runs on a busy host.** They were taken while
  eight agents shared the machine, with no warm-up, no repetition and no floor,
  because their only job is to separate 31 minutes from 10. Treat them as an
  order of magnitude, not as a measurement.
- **The gate job gets longer.** `facade-format-tests` rebuilds `litchi` under a
  feature set the later modes do not share, and `harness-tests` builds a
  project outside the workspace, both under the job's `CARGO_BUILD_JOBS=1`
  invariant. This record does not measure the added CI minutes on a GitHub
  runner; the local figures above are from a 32-core host.
- **One standing gate fails, and it is not this change's.**
  `tools/check_example_targets.py` exits 1 on four cross-package duplicate
  example target names among `litchi-keynote`, `litchi-numbers` and
  `litchi-pages`, reproduced identically on the untouched before checkout. It
  is left open: every crate it names is iWork, which the brief excludes.
- **The new modes close a reachability gap, not a correctness one.** They make
  two suites runnable from the standing gate list and executed by CI; they do
  not assert anything about what those suites cover. Neither suite gained a
  test here.
- **CI runs these on `main` only.** Both workflows in this repository trigger on
  `push`/`pull_request` against `main`. The performance program commits to
  `feat/office-format-completeness`, so a change that never reaches a PR is
  still gated only by what its author runs locally. This record does not widen
  the trigger; doing so would change what CI spends on every branch and is a
  decision for the workflow's owner.
- **The harness suite is not otherwise strengthened.** `perf-baseline.yml`'s
  `smoke` job already runs `cargo test --locked --manifest-path "$PERF_MANIFEST"`
  plus a `security_corpus` ignored-test run, a release `xlsx_filesystem` leg and
  an `allocator-metrics` bin run. `harness-tests` deliberately duplicates only
  the plain suite, in a workflow whose path filter is broader, and adds none of
  the others.
- **The demo is not a test.** `office_crud_demo` is an example; nothing runs it
  in CI, and this change does not add a runner for it. The evidence is two
  manual runs and one member census, not a regression fence.
- **The demo's UPDATE no longer adds a slide.** Adding a slide to an *opened*
  deck is not something `opened::Transaction` exposes, and change 0607 states
  that giving an opened deck a route into the mutable writer "would be the wrong
  thing to build". The demonstration therefore shows the edits the supported
  route owns. Whether the opened transaction should gain slide insertion is a
  design question this record does not open.
- **`detect_workbook_format_from_signature` is still dead.** It is bounded and
  out of this change's scope. Whether the facade wants a signature classifier at
  all is a separate question; the `#[allow(dead_code)]` on it is untouched, as
  is its inaccurate "public" wording.
- **`crates/litchi` does not inherit the workspace lint table.** Its
  `Cargo.toml` has no `[lints]` section, so `unused`, `unreachable_pub` and the
  clippy set are warnings there rather than denials. That is why a dead `pub fn`
  in a private module survived 0587's survey, and it is not changed here.

## Retained evidence

[`results/change-0639/`](results/change-0639/README.md) — the two demo
transcripts, the member census, the rustdoc public-item comparison, every gate
tail, `decision.json` and `log-sections.md`.
