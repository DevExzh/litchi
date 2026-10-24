# 0362uma — final gate findings and required repairs

The full gate on `f6fe430a4d19d94aa360c0e9ed03ebd757265225` passed
formatting, combined strict consumers and every per-format strict profile:
25 completed steps. Its operations tests exposed four failures before the
runner's session ended. The coordinator confirmed the session was missing,
both runner processes had exited and no Cargo process still used the gate
checkout or target. The operations step did not finalize a result row.
This is an interrupted, failing attempt, not a full gate pass.

[The retained receipt](results/0362uma/interrupted-full-gate.json) binds the
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
overwriting the interrupted attempt. [Their reduced receipt](results/0362uma/selected-diagnostics.json)
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
gate is required on the new source. This record uses suffix `0362uma` to
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
reviewed under record 0362yh to bind the real original/final projections,
match external dependencies and build settings, and allow only the exact
reviewed local graph changes. The existing equal-projection mode remains
unchanged; no native result is accepted through a bypass.

A subsequent native diagnostic selection passes strict all-target/all-feature
Clippy for all 19 native packages and ODS. Its completed 201 test/build chunks
expose one real retention failure and twelve stale doctest consumers. The OPC source
artifact retains its read-ahead allocation after the package is dropped
(`retained_source_artifact_does_not_pin_the_read_ahead_window`: 1,603 bytes
instead of zero). Four OPC doctests use missing or external-invalid limit
paths; eight ODP doctests use the previous projection/error spellings.
Dedicated repairs preserve the retention assertion and every example.

The completed failing chunks remain at
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/remaining-leaves-f6fe430a4d/leaf-tests/`:

| Log | Bytes | SHA-256 |
| --- | ---: | --- |
| `0033.log` | 135574 | `9393f3c06794f68ae1028e6c956568f0f088e7e22843e844039fc46377b69f24` |
| `0036.log` | 123357 | `8c8f6398f51e95481570cc4f3e28214a0aee8c2099fb317f74188af611e815de` |
| `0083.log` | 150365 | `7c6b3149b81dca4c00724fb3f768fd1f124b58b716e00b948cefbe7c73bb1538` |

After the reviewed boundary and corpus-receipt merges, all 14 final tooling
modules pass together: 1,413 tests in 45.816 seconds at `885cdc82af`.
The external command/source-bound log is
`.coord/final-acceptance-2026-10-24/composed-tooling-885cdc82af.log`,
634,360 bytes, SHA-256
`b54f9fbdc25056c61317af08bde26517fc0d9d7246fda8c3db5d762b55b3f2e8`.
This is a composed tooling pass, not the complete source gate.

[The completed native diagnostic receipt](results/0362uma/native-diagnostics.json)
binds all 201 chunks, the three finalized gate rows and all external
plan/state/log hashes. Cargo test-result summaries total 21,589 passed,
13 failed and 83 ignored cases. The only failed chunks are 33, 36 and 83,
covering the retention check and twelve doctests above. The final target is
19,005,423,149 bytes; stable post-exit samples show no remaining builders.
The process exits 1, and all later steps in this selection remain unrun.

Reviewed repairs 0362umb and 0362umc are now composed at
`49c25ef67f62a8bb943e7c6a0bf464e49ae87d07`. The coordinator advanced the
clean diagnostic checkout to that source with its lock unchanged, and adopted
the independently reviewed 0362uva runner from its separate directory.
The isolated `wsdoc` selection passes with warnings denied and no competing
Cargo builds in 1,057.472 seconds. Its empty command log, stable source/locks,
process isolation and 4,504,969,096-byte final target are bound by the
[reduced receipt](results/0362uma/isolated-wsdoc.json). The prepared runner
retains all 53 steps; this selection completes only workspace documentation.
The f6fe diagnostic results are not relabeled as acceptance on the new freeze.

All 68 fresh library rustdoc captures subsequently passed on that same source,
under `/home/zhuhe/.cache/litchi-api-contract/49c25ef67f62a8bb943e7c6a0bf464e49ae87d07/`.
They use all features, include hidden public items, and record every package,
source/lock/toolchain identity, JSON hash and resource log. The coordinator
verified all 68 receipts and artifact hashes: 1,094,716,714 JSON bytes,
maximum per-command RSS 2,997,526,528 bytes, and maximum target size
5,644,910,073 bytes. The raw JSON remains external; the
[reduced capture receipt](results/0362uma/all-68-inventory-capture.json)
binds each artifact and command log, the source tools and regeneration command.

The final inventory generator then failed before emitting findings:
`ValueError: glob target requires module rustdoc data: litchi_core::detection`.
The raw `run.json` remains `incomplete`, the process exits 2, and the TSV
contains only its header. This is a required public re-export coverage repair,
not zero API debt or an accepted exception. The failed run and every successful
capture remain unchanged for validating the repair. Semantic dispositions and
final inventory closure remain outstanding; the full gate launch waits for
the resolver repair and review.

These are required engineering corrections. No performance regression,
capability loss or API-contract debt is accepted as an exception. Only the
comprehensive redaction audit in memo item 22 remains deferred.

The failed aggregate capture remains immutable. Subsequent reviewed tooling
repairs (0362ul1 and 0362ul3) enable complete diagnostic enumeration, with
foreign-definition coverage and nominal identities checked before suppressing
duplicate boundary findings. Source review then identifies further required
engineering, recorded in 0362ul5. The full gate must therefore bind the next
assembled source rather than treating the historical 49c capture as closure.
