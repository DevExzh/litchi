# Value-inspection performance capture history

The owned performance evidence retains two complete frozen-source pairs and the failed setup/partial attempts that explain their provenance. The existing `performance-report.md` and `.json` remain the earlier complete report; `performance-report-latest.md` and `.json` are derived from the root-owned current result directories.

## Capture history

| session | disposition | retained evidence |
| --- | --- | --- |
| `21287` | Setup stopped before timing because `/tmp` could not create the baseline worktree on the full tmpfs. | Preflight-only diagnostic directories remain retained. |
| `71321` | Partial run: baseline completed; candidate stopped at `numbervalue-invalid-separator.evaluate.sample-01` because the old scalar checksum path rejected a 1×1 error array. | 3,000 successful rows plus one empty failed stdout receipt remain retained. |
| `85483` | Complete frozen-source pair. | Diagnostic baseline `063738Z` and candidate `063808Z`. |
| `87626` | Root duplicate complete pair; current retained result set. | `baseline-*` and `candidate-final`. |

## Independent audit and disposition

`../../root-performance-audit.json` (SHA-256 `02f78b8127b754f18e52ee9824c2d12361b03a8d6af589f3c55d78458de49d13`) independently verified 6,900 complete samples across the two complete pairs and found unchanged accounting fields in all 56 matched groups per pair. The audit retains **2 earlier-complete review triggers** and **8 latest-complete review triggers**. Candidate source/profile/lock and candidate binary hashes match across the complete pairs; baseline rebuild hashes are recorded in the JSON history.

The earlier complete pair retains the SUMIFS sparse-reference parse latency trigger at **+7.7169%**, with bootstrap interval **[−1.9621%, +10.9134%]**, and one **+172 KiB** RSS trigger. The latest pair has no >5% latency trigger and eight RSS triggers from **+168 KiB to +212 KiB**. These observations remain review triggers; no RSS cause is inferred from unchanged allocator counters.

Accounting, source/profile stability, typed outputs, read bounds, cancellation totals, and cleanup receipts are PASS. Candidate cancellation receipts contain one successful read across four sticky-cancel repeats; the normalized integer field is therefore zero by floor division.

## Retention

The refreshed `retained-files.json` covers every file under this `results` directory except the manifest itself. It supersedes the prior 12954-entry manifest and includes both complete pairs, the partial/failed raw receipts, current raw receipts, and the derived history/latest reports. The final manifest hash is reported with the handoff because the manifest intentionally excludes itself.
