# 0362ul165s — Diagnostic runner and metadata qualification review

The remaining twenty diagnostic checks use a fresh source at
`7c043ea168204035c8550d51d0f490ec271ce2f8`. They cover full-gate indices
38–52 and runtime-supplement indices 7–11. No earlier result is reused, and
this suffix cannot establish final all-65 acceptance. Row 37 is recorded
as skipped; its OfficeCLI preservation barrier is neither opened nor
claimed satisfied.

## Diagnostic runner correction

Root and independent review return the first handoff before execution.
It authenticates the current runner but obtains the executing module from
the historical rebinder. That historical runner omits the new ODF consumer
lock and the current tooling-test target-guard behavior. Its post-step
source check would therefore reject the current four-lock source identity.

The v3 successor loads and executes the current runner, asserting its module
path and SHA-256. A substitution control rejects the historical module.
Source controls and dry-check pass; the earlier packet remains preserved.
The current runner is `1d537811540f9ecac0c57171266d431f0bf5ca35f63e6d58a6e5cb8da7ff475b`
and its helper is `67298a85d45eb57621ad8465b8a67b3be0c933efa55b6a16b4929bd75a1ca126`.

Independent review accepts the corrected execution path. The diagnostic
starts in `litchi-final-gate-current-remaining20-v3.service`, with a 20 GiB
memory limit, no swap and one Cargo job. Root verifies those limits. The
Python binding build at row 38 passes; 300 of 301 tests pass. The failure
expects an empty signed-file edit to refuse, but receives a no-op commit.
That individual-edit expectation is under contract review separately from
the owner's confirmed immediate-refusal policy for signed batches. Source
and all four locks remain unchanged, no resource or leftover-process error
occurs, and the target is about 2.96 GB. The diagnostic continues to the
feature matrix; there is no final-gate acceptance yet.

## DOCX metadata contract and review

The forward-core oracle follows the already confirmed contract in
[ADR0015](../adr/0015-lossless-core-properties-crud.md#consequences).
Semantic core edits retain modeled lexical values and canonicalize XML;
comments, processing instructions, formatting and prefix spelling need
not survive that rewrite. This is not a new fidelity exception or a
production defect. Unknown content and tolerant repairs remain subject
to the existing refusal rules.

When the actual patch leaves core unchanged, app/custom/no-op routes must
preserve source core bytes exactly. Current inverse replay must restore
the original bytes. The probe must determine this policy from the actual
core transition, rather than merely its requested scenario.

Independent review of native leaf `9c918bd24a` closes the first seven
findings but returns three further evidence-integrity defects: a later
child can alter earlier command logs or event metadata; the standalone
validator accepts unrelated files under `commands/`; and the validator
publishes success before the final runner receipt exists. Successor
`b8798fa4ab` closes those three findings, with 58 passing source tests.
Review returns one further CLI defect: validating completed output writes
a new pending summary, making a repeated validation fail. Successor
`bdb6343c3f` makes completed validation read-only and passes independent
review and 59 source tests. The leaf remains unmerged pending compiler
validation. Passing source tests are not native qualification evidence.

Four bounded release builds of the ordinary and allocator probes are
prepared against immutable `9c918bd24a` and historical `37db93de89`.
The historical overlay changes only five benchmark consumer files;
production code, allocation counters and dependency lock remain unchanged.
Independent review accepts the v2 build handoff after it fixes rejection
of its own receipt. These builds are queued behind the diagnostic; neither
probe execution nor timing has run.

## Branch disposition

The read-only census observes 77 heads and 66 worktrees, including the new
detached diagnostic source. All 70 excluded refs retain their recorded
heads. The DOCX metadata native leaf is the only unmerged active unified
branch. Five other unified branch controls are already merged. Root then
retires the completed Q472a diagnostic worktree and branch, preserving its
139,207-byte ignored lock and all external logs and receipts. Its exact
commit remains in trunk history. The census becomes 76 heads and 65
worktrees; cumulative retirement reaches 283 worktrees and 445 refs. Four
merged controls still supply pending protocol preparation. A reference in
a closed historical receipt alone does not require keeping a checkout.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `final-gate-current-remaining20-diagnostic-20261026-v2-independent-review-20261026/review-receipt.json` | `8f8f5986420d91ab6d10624755ec3dc7d0790d2136dfe293949b7e9b43b95fbd` |
| `final-gate-current-remaining20-diagnostic-20261026-v3/plan.json` | `5a9ed5f86646fc3e1cdfbd8e0fbe509f2a39d655418ffda599260f7c909b37dc` |
| `final-gate-current-remaining20-diagnostic-20261026-v3/source-only-receipt.json` | `db6061f919aaa5879ba516691be19c5126801f31f324646d2a11a7989c476091` |
| `final-gate-current-remaining20-diagnostic-20261026-v3-independent-review-20261026/review-receipt.json` | `0b6f7ff961f2760d312253ba581d854747c54465ac6b2892445803a8aad19774` |
| `final-gate-current-remaining20-diagnostic-20261026-v3/launch.json` | `3b75c6adda7a3c8b7fc786ba59b096de756abbb3172c772bb253a692c44ce598` |
| `remaining20-python-signed-noop-failure-20261026/receipt.json` | `fe9d957df5f7057b234a2f4cb7584dba601de632c1e9496fe4f69d57e5c86d62` |
| `docx-metadata-core-contract-root-closure-20261026/receipt.json` | `7c95fe2c0b1c1dee975c1235fe1dcd55d36ccf185cfde836a0e8868284ef30f2` |
| `docx-metadata-native-9c918-independent-review-20261026-fresh/review-receipt.json` | `bf4371c29066f6543d74f561cedd85af384ec971fb7d903937b9f549172193e1` |
| `docx-metadata-native-9c918-independent-review-20261026-fresh/final-receipt-order-addendum.json` | `0e7eb349825d122051890e9b2b3c1c10292387db88a9399c72b68d12844a2d87` |
| `docx-metadata-native-b8798-independent-review-20261026-fresh/review-receipt.json` | `f0f36ab70a4862bb3522ab94911d9707875eeafce4f715e069ed6dbeda5ddbf3` |
| `docx-metadata-native-bdb634-independent-review-20261026-fresh/review-receipt.json` | `9374e881ba34ddfd5331a10884c99c9ddbcef446154a3beaa773096c2fc7e919` |
| `docx-metadata-native-release-compile-9c918-v2-independent-review-20261026/review-receipt.json` | `e89e92933633e57b223e1d8cd4ef0e88ce73ac80bc5f27674b8e5642f1ce122b` |
| `branch-worktree-census-current-7c-20261026/receipt.json` | `2a84c4470d3cbb613b3c040dec32412648512a9223db190512f4b107c646ba24` |
| `retired-q472a-20261026/receipt.json` | `3f565da63089c0b8308cfcb34c95ffb22dfba633f4388fe95608d5897f3fd5e0` |

Actual final-candidate rebinding, all 29 native qualification groups and
the final all-65 gate remain required. Nothing is pushed.
