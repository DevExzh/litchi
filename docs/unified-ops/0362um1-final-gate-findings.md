# 0362um1 — final gate findings and required repairs

The full gate on `f6fe430a4d19d94aa360c0e9ed03ebd757265225` passed
formatting, combined strict consumers and every per-format strict profile:
25 completed steps. Its operations tests exposed four failures before the
runner's session ended. The coordinator confirmed the session was missing,
both runner processes had exited and no Cargo process still used the gate
checkout or target. The operations step did not finalize a result row.
This is an interrupted, failing attempt, not a full gate pass.

[The retained receipt](results/0362um1/interrupted-full-gate.json) binds the
plan/state, all 25 completed steps and the five existing operations-test
chunk logs. It preserves the partial results without assigning a successful
status to the unfinished step.

| Finding | Evidence | Required action |
| --- | --- | --- |
| XLS nested read fence | `nested_fenced_reads_share_the_outer_check`: two source-version observations where the test expected one | Establish whether both are required for source consistency; retain correct fencing and verify its cost. Do not remove a required fence merely to satisfy a count. |
| Corpus fidelity report | 1,020 files opened; triage and unexpected-fidelity assertions passed, then the TSV golden differed | Generate the report in a separate checkout and inspect every changed row before accepting an intended report update. |
| CRUD report | Zero unreached claimed cells; the checklist golden differed | Review every changed outcome and error classification; do not bless away capability loss. |
| Unlocked encrypted DOCX | The edit closure ran where the test required an early publication-policy refusal | Restore the required refusal without losing lazy edit entry, encryption protection or signed/no-op fidelity. |

Focused repair leaves own these investigations. The frozen gate checkout
remains unchanged. In particular, blessed report generation must never use
a test executable compiled for that checkout: its manifest path is embedded
in the executable.

Separate selected diagnostic runs retain the remaining observations without
overwriting the interrupted attempt. [Their reduced receipt](results/0362um1/selected-diagnostics.json)
binds all completed steps, external state/plan/log identities, source/locks,
runner policy and individual test chunks. The coordinator verified the
recorded log sizes and hashes.

- The facade suite completed with 208 passed and one failed test. The failure
  happened at Unix socket fixture creation: the gate's long `TMPDIR` exceeded
  `SUN_LEN`, before the library opened the input. Reviewed correction 0362un1
  uses a short temporary parent for all three socket fixtures. Each focused
  test passes with a deliberately 167-character `TMPDIR`; production behavior
  and assertions are unchanged.
- Core tests pass: 332 tests and doctests, with no failures or ignored cases.
  The iWork leaf API, non-iWork boundary and static corpus disposition checks
  also pass on the frozen source.
- The tooling suite runs 1,360 tests and returns four failures: the upstream
  reference scanner matches ODG path-rejection literals, two iWork table-merge
  topology rules still expect eager constructors, and diagnostic evidence is
  missing from the benchmark index. The rules and evidence classifications
  require repair; the failures are not waived.

The final observed target size is 18,446,285,179 bytes, below the 20-GB cap.
Every completed step retains stable post-exit observations with no remaining
owned builders. These selected results do not establish full acceptance.
Once the required corrections are reviewed and composed, a fresh complete
gate is required on the new source. This record uses suffix `0362um1` to
distinguish it from the pre-existing 0362um OPC consumer record.

The final inventory review also found that its owner-name heuristic omitted
policy-declared immutable owners with other names. That is an inventory
coverage defect requiring repair and regression tests before the all-68
capture can establish closure. The new disposition validator provides
structural checks only; source evidence and coordinator review must still
justify every row.

For W8 performance, the official before leg remains pristine f2 with its
reviewed probe/consumer adapter and historical dependency resolution. Copying
the candidate's manifests onto f2 would change the baseline build graph and
could hide a cost introduced by the API work. That proposed overlay is not
accepted as the official baseline. A separate explicit comparison policy is
being implemented and reviewed to bind the real original/final projections,
match external dependencies and build settings, and allow only the exact
reviewed local graph changes. The existing equal-projection mode remains
unchanged; no native result is accepted through a bypass.

These are required engineering corrections. No performance regression,
capability loss or API-contract debt is accepted as an exception. Only the
comprehensive redaction audit in memo item 22 remains deferred.
