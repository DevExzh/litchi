# 0889 MCE control qualification

The MCE differential harness now accounts for its input population and records
structured codec and stream outcomes. This closes two gaps identified in
[0888](0888-executable-residency-and-mce-controls.md): unreadable inputs could
silently disappear, and the old timing digests omitted observable fields.
This batch changes diagnostic code only. It adopts no production optimization
and makes no latency, allocation, or RSS claim.

## Input and outcome contracts

The v2 differential report records package and member identities, selected and
read member counts, and input failures. A failed read makes qualification
incomplete. The independent Python ZIP reader checks the real-fixture population
without using the Rust harness's archive reader or traversal implementation.
Archive discovery retains content identities rather than all archive payloads;
processing reads one package at a time and refuses a changed identity. The
result maps and input inventory still grow with the qualified population, so
this is not a constant-memory claim for the whole diagnostic.

Codec records bind the input, capabilities, limits, output size and SHA-256,
all report counters, and typed refusals. Stream records bind those inputs to
separate raw and semantic callback counts and framed transcript digests,
complete stream reports, and primary and retained prior errors. Framing uses
length-prefixed fields, preserving qualified names, expanded names, attribute
order, raw lexical values, decoded values, and event variants. The existing
timing wrappers are not promoted to full correctness oracles.

Codec and stream outcomes are compared within their own contracts. They are
not assumed equivalent: the stream rejects some malformed structures that the
codec's no-MCE passthrough does not validate. Small tests assert independently
specified outputs, counters, transcript hashes, and refusal behavior; repeated
broad-corpus equality is an additional determinism check.
The generated corpus is not an independently implemented full MCE reference.

## Real input population

| Package | XML, relationship, and VML members |
| --- | ---: |
| `StructuredRefs-lots-with-lookups.xlsx` | 36 |
| `drawing.docx` | 11 |
| `NumberedList.docx` | 20 |

The [independent census](results/change-0889/independent-input-census.json)
binds all three package hashes and all 67 selected member hashes. It includes
the real worksheet, document, and styles parts proposed in 0888 and the other
selected members of those same packages. It is not an all-OOXML-corpus claim.

The [external verifier](results/change-0889/verify_oracles.py) checks exact
outcome-key coverage, input identities, capability and limit settings, report
fields, complete-oracle flags, and repeat equality. A missing-outcome negative
check demonstrates that matching aggregate counts or repeated incomplete
outputs cannot satisfy the population check.

## Validation and evidence

The corrected source passes **24 focused tests**, formatting, warning-denied
Clippy for the binary and tests, warning-denied rustdoc, and a diagnostic build
with Rust 1.95.0, locked/offline dependencies, two build jobs, no incremental
compilation, and no debug information. These are development-profile
correctness checks, not release performance measurements. No production,
dependency, or public API changes are made. Existing `forbid(unsafe_code)` and
library resource policies remain intact.

The crate-boundary check passes for 65 workspace packages and 244 internal
dependency declarations, with 11 existing explicit debt items. Its first-run
result is reused after checking that its production and manifest inputs are
unchanged; the diagnostic source corrections do not change that graph.

Each of two fresh diagnostic processes reads all 67 real members and evaluates
20,000 generated seeds plus 20,000 aliasing seeds. These produce 13,811 and
19,418 distinct generated input hashes respectively; seed count is not a
claim of unique documents. Each report has **200,201 structured MCE records**
and 200,405 legacy results, including publication-audit and styles outcomes.
Codec and stream refusals remain in the comparison rather than being dropped:

| Outcome per process | Records |
| --- | ---: |
| Codec success | 51,330 |
| Codec refusal | 48,804 |
| Stream success | 51,934 |
| Stream refusal | 48,133 |

The ten additional runtime controls cover zero/one/four/five attributes,
1,024/1,025 and 4,096/4,097 attributes with namespace declarations included,
and bare/understood-extension directive selection. The focused unit tests
also cover namespace rebinding versus new prefixes; directive tokens; depth;
input/output bytes; stream event count, event bytes, raw attribute bytes,
expanded-name bytes, and context bytes; raw lexical spelling; opaque
extensions; and primary/prior error precedence. Exact and one-over checks use
small configured bounds where appropriate. Deterministic allocator-failure
injection is not provided by this public diagnostic API, and the streaming
API emits callbacks rather than a materialized output buffer.

The two full JSON reports are byte-identical at 226,180,145 bytes each, SHA-256
`53f4e5b747b9c019aae3922e7df1ef7f98bbbb4a047f74b467443e8d49eefe57`.
Deterministic gzip copies retain those bytes after verified decompression.
[Independent verification](results/change-0889/oracle-verification.json)
passes exact population, configuration, report, and control checks, including
negative checks that remove an outcome or the entire control panel.

Initial validation evidence is retained:

- Attempt 1 found an incorrect opaque-extension fixture expectation and two
  Clippy issues. Registering opaque descendants does not override
  `mc:Ignorable` for the extension itself; the corrected test explicitly
  checks both ignorable and understood cases. Its source overlay and logs
  remain available.
- Attempt 2 passed the corrected Rust gates, then stopped on a Git-status
  whitespace bug in the evidence runner before any corpus execution.
- Attempt 3 passed the gates and both broad runs. Its missing-root negative
  passed. The runner then tried to hash a deliberately unreadable copy after
  removing its permissions and stopped before launching that negative child.
  Completed corpus runs are retained, not repeated to replace that attempt.

The [aggregate qualification receipt](results/change-0889/qualification-receipt.json)
passes using attempt 3 for the broad runs and missing-root refusal, attempt 5
for the unreadable-package refusal, and attempt 7 for the corrupt-member
refusal. Attempt 4 stopped at a target-marker mismatch before starting a child;
its launch failure is retained separately. Attempt 5's first corruption probe
changed a DEFLATE final-block flag but decoded to identical bytes. The
[independent mutation review](results/change-0889/negative-mutation-review.json)
records that result and verifies that the replacement reserved block type
fails decompression. Attempt 6's extra unreadable and corrupt runs both passed
and remain recorded; attempt 7 supplies the final corrupt-only receipt.

The task-owned validation target and temporary archive copies are removed.
The [cleanup receipt](results/change-0889/attempts/validation-7/target-cleanup.json)
retains the target inventory and binary identity. The separate
[negative-input cleanup receipt](results/change-0889/attempts/validation-5/negative-input-cleanup.json)
binds the two fixture copies left by attempt 5 before removal. No worktrees were created
for this batch; unrelated worktrees and the shared diagnostic build cache are
preserved. Post-cleanup [oracle verification](results/change-0889/post-cleanup-oracle-verification.json)
and [packet verification](results/change-0889/post-cleanup-packet-verification.json)
replay retained evidence without rebuilding or repeating the corpus workload.

From the repository root, replay the evidence with Python 3:

```sh
python3 -B docs/performance/results/change-0889/input_census.py
python3 -B docs/performance/results/change-0889/verify_oracles.py --first docs/performance/results/change-0889/attempts/validation-3/differential-first.json.gz --second docs/performance/results/change-0889/attempts/validation-3/differential-second.json.gz --output /tmp/0889-oracle-replay.json
python3 -B docs/performance/results/change-0889/verify_packet.py --check-current-source --require-cleanup --output /tmp/0889-packet-replay.json
```

Output paths must not already exist. Source verification requires the recorded
Git base and unchanged diagnostic source; receipts retain the original Linux
workspace paths, toolchain, and execution environment. The source overlays
reconstruct the inputs independently of the removed target.

The proposed repeated-name-expansion MCE optimization remains a hypothesis.
It needs matched release measurements and these controls before adoption.
The rejected 0887 candidate stays rejected.
