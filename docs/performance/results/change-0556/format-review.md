# 0556 XLSX provenance merge format review

`status: post-format source audit`

`performance_claim: none`

`production_change_by_reviewer: none`

This review binds the final formatted candidate after the candidate-01 format
failure was corrected. It audits source identity and the format-only delta; it
does not change production or capture performance. The base revision is
`3210f07ac7a9daa2686e781645f6539a156a26ce`.

## Exact artifacts

The pre-format candidate review remains preserved at
[`candidate-attempts/before-format/candidate-review.md`](candidate-attempts/before-format/candidate-review.md),
SHA-256 `eb441d2a3c3db1e5876e4461f3acb553d6d5ca3af1c13e0d43ee4318003da2eb`.
The source proof is [`source-proof.md`](source-proof.md), SHA-256
`ccf77475d133b2833b823e93c4e162a122a3e9fcb208c6d1ca17660817d39dbe`.

| artifact | SHA-256 |
| --- | --- |
| baseline `baseline-cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| pre-format candidate source | `fc3ede23ee3ff9b7c68b0f11d73bc95ff168105a44a8ad409f7116f92c4b196e` |
| final candidate source `candidate/cell.rs` | `f7fa882fcf4ccd4909950660f0b614760c7b96eaccf44341129506db3f00e68c` |
| final candidate patch `candidate/candidate.patch` | `d8f509ecbeaac1b2346f490fe22c0e73a5c3074c7ed82418bc88c198642eba8c` |
| current production source (restored baseline) | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| format correction receipt | `461dd8e762fe813c5bb01da784719ddff3ff205c71bb5ebc8bcd05fc07b3f36e` |

The regenerated final patch passes `git apply --check` against the restored
baseline source. The final candidate source and patch are the exact bytes
used by the corrected candidate quality runs.

## Format-only proof

The pre-format and final candidate sources have the same
`#[cfg(test)]\nmod tests` marker at byte offset 33,489. Every byte before that
marker is identical; the shared prefix is 33,489 bytes with SHA-256
`3b45ff0c32ea0d4e3f5ad648caa1689a7d75271bcfb3b226e73acd13e2f06c35`. This
covers the production implementation, including the linear merge and the
explicit parsed-iterator drop at the end of the walk (`candidate/cell.rs:879-881`),
as well as the unchanged test-only `Store::get` method.

The only source delta is inside the test module: 90 inserted and 68 removed
lines, all from line wrapping and formatter trailing commas. Running
`rustfmt --edition 2024 --emit stdout` with each file fed through stdin
produced byte-identical output (55,341 bytes; SHA-256
`f7fa882fcf4ccd4909950660f0b614760c7b96eaccf44341129506db3f00e68c`). This
confirms that the post-format source is the formatter's canonical form of the
pre-format source. `cargo fmt --all` completed with exit code 0 in
`candidate-attempts/before-format/format-apply.json`, and the corrected
candidate-03 format check also exited 0.

The explicit iterator drop remains relevant to resource review. The parsed
`IntoIter` necessarily overlaps the merged cell vector during the interleaved
walk; the post-loop drop releases its backing allocation before merge-range
allocation and index rebuilding. Formatting does not change that lifetime
boundary, and this review makes no peak-memory or allocation claim.

## Quality-attempt disposition

Candidate-01 is a failed aggregate because its format check ran before the
correction; its XLSX tests and warning-denied Clippy rows passed, but that run
is not the final formatted quality receipt. Candidate-02 is intentionally
inadmissible: it started while the formatting command was still live, its
source guard rejected the changed manifest, and no result or quality receipt
was published. Its `aborted.json` (SHA-256
`98b4a35982a0d011d4eb8b1e36aa07bf9d40012ddab90fd2e00d8e077e1cb526`) is
retained as the exclusion record.

The corrected candidate-03 targeted run passed XLSX tests, warning-denied
Clippy, and format. The candidate-04 command run passed workspace check,
XLSX documentation, and no-default-features check. These are quality
receipts only (candidate-03 result SHA-256
`b04688082fd8d656c4cb2a7eb4f9f4374b558675c11fc589b21cfb1a0132c780`,
candidate-04 result SHA-256
`941457ef21907d7d5d86264433ca340cd599cfc7c10cc14a7d054d57296b4199`); no
performance capture or adoption decision is represented here. Root must keep
the candidate source and patch hashes fixed for any subsequent measurement or
terminal decision.

## Verifier sanity check

`python3 -B docs/performance/results/change-0556/verify.py` passed every
pre-final assertion, including baseline restoration, final candidate source
and patch binding, candidate-01's pre-format manifest, candidate-02's
inadmissible abort record, candidate-03 and candidate-04 receipts, and the
format-correction prefix check. It then stopped at the expected missing
`final-checks/result.json`, because the final decision is still pending. No
misleading assertion or quality overclaim was found in that pre-final path.

## Verdict

Format review passes. The final candidate's production prefix is byte-identical
to the pre-format candidate, the test-only suffix is formatter canonical, and
the final patch applies cleanly to baseline. The existing source-proof and
candidate-review conclusions remain valid; correctness, allocation, and
performance decisions still belong to the root campaign.
