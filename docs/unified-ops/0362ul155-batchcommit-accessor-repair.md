# 0362ul155 — canonical BatchCommit accessors and the Keynote alias removal

This follow-up records the contract-accessor repair committed at
`29107fe739f3ea729b643cf4406b3a8729d9864e`, on top of the frozen unified
source used by the contract-accessors leaf.  The preceding implementation
commit had the complete source migration but no 0362ul155 record; this file
is the required provenance and validation boundary.  Nothing in this record
weakens the 0351 accessor or immutable-owner contract.

## Canonical surface

`litchi_ops::BatchCommit` now exposes the same consuming and borrowing value
shape as the shared `Commit<T>`:

```text
file(), into_file(), patch(), diagnostics(), value(), into_value(), into_parts()
```

The old Rust members `results()` and `into_results()` are deleted, with no
alias or deprecation shim.  The implementation documentation, the facade's
batch test, litchi-cli, the pyo3 Rust backing path, ops tests, fuzz harness,
and all standalone unified batch harnesses use `value()`/`into_value()`.  The
Python `Commit.results` property remains intentionally unchanged as a
semantic Python property over the renamed Rust accessor; this is not a Rust
alias.

The deprecated `litchi_keynote::Show::slide` compatibility member is deleted.
No Keynote production or test consumer used that member.  The minimal and
full Keynote glossary profiles now carry an exact `litchi_keynote::Show::slide`
ban so the removed alias cannot return.

## Consumer and documentation closure

The migration is in one implementation commit across 56 tracked files:

- `litchi-ops` implementation, every affected format integration test, fuzz
  target and unified facade test;
- litchi-cli rendering documentation and batch execution;
- pyo3-litchi's Rust extraction path, while preserving the Python property;
- the five standalone harnesses that directly consumed the old Rust member:
  `unified-batch-closure`, `unified-docx-batch`,
  `unified-pptx-chunk-publication`, `unified-pptx-command-batch` and
  `unified-xlsx-batch`;
- the four other harnesses linked by the W8 record
  (`unified-docx-text-fanout`, `unified-xlsb-cell-value`,
  `unified-xlsb-remove` and `unified-xlsb-sheet-batch`) were also scanned and
  had no obsolete BatchCommit accessor use;
- `docs/API_GLOSSARY.md`, the 0351 rule-7 wording, the W8 contract record,
  the unified README detection wording, and `tools/api_glossary.json`.

The glossary gate registers the three unified batch result methods
(`File::batch`, `File::batch_with`, `File::batch_with_options`) as returning
`BatchCommit` and bans both deleted members at the exact BatchCommit scope.
Detection prose now distinguishes the required raw `&SourceLimits` policy
from an owner's `OpenOptions` bridge.  W8's decision no longer describes a
`results()`/`into_results()` exception, and its RTF section records the
current typed `Refused(Refusal)` shape.  Existing `results()` methods in the
ODF/ODT leaf APIs and their performance consumers are unrelated APIs and are
unchanged.

## Source-only evidence

The following checks passed against the implementation commit before this
record was added:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -B -m tools.test_check_api_glossary
Ran 57 tests ... OK

PYTHONDONTWRITEBYTECODE=1 python3 -B -m tools.test_api_debt_inventory
Ran 13 tests ... OK

PYTHONDONTWRITEBYTECODE=1 python3 -B -m json.tool tools/api_glossary.json
git diff --check
```

An active-surface scan checked 239 relevant Rust files, including all
unified ops tests/fuzz, facade/CLI/Python backing code and the W8-linked
standalone harnesses.  A subsequent repository-wide scan covered 12,855
tracked Rust files and 5,355 tracked contract/documentation files: no
`BatchCommit` use of either deleted Rust spelling remains.  The only
remaining `.results()` calls are leaf ODF/ODT result APIs (plus their
matching perf code), and the only current glossary hits are the intentional
exact-ban replacement explanations.  The Keynote source has no
`pub fn slide(...)` declaration.

The staged artifact guards also passed: no staged evidence object at or above
1 MiB and no staged `binaries/` path.  The worktree is clean at the committed
tip.

## Validation boundary

No Cargo, rustc, rustdoc, native, benchmark, or full workspace gate command
was run in this leaf.  The coordinator's separate full-65 C1 run currently
has a terminal CRUD PPT hyperlink-add failure; that failure is unrelated to
this accessor migration for purposes of this record, remains unproven as a
source-wide result, and is not claimed as a pass here.  The coordinator must
run the source-bound compiler, focused tests, final glossary/inventory gate,
native fidelity and quiet ABBA measurements after integrating this follow-up.
Only the redaction audit remains an allowed deferral; this record does not
defer any API work.
