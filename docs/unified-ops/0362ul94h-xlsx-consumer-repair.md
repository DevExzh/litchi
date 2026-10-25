# 0362ul94h — XLSX consumer error-contract repair

This source-only follow-up repairs the XLSX diagnostic from the actual OOXML
consumer check at `5ec78a48589c7ab4396ee90cb1f6d41cf7f1aaa6`.  The retained
compiler log is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/ooxml-combined-0362ul94-f5-consumers-5ec78a4858/02-check-leaves-attempt-actual-f5-02.log`;
it is 33,227 bytes with SHA-256
`bde73361831156dabd2d8b7032b40ef472506856aa4dda8be40781425eaba220`.

## Repair

The semantic transaction's `join_normalization` match now handles the two
crate-owned semantic failures introduced by the shared OOXML error contract:
`SourceChanged` and `Internal`.  Both are explicitly classified as outside
the source-equivalence taxonomy (`JoinNormalization::Other`), while
`JoinError::normalization_source()` and its standard error source continue to
retain the original typed leaf.  This prevents either failure from being
collapsed into malformed input, a refusal, or publication.  The existing
unified classifier continues to map direct `SourceChanged` to `io_error` and
direct `Internal` to `internal_error`.

A focused unit test exercises both new arms, including their fixed leaf
messages.  No wildcard arm or compatibility alias was added, and no edit,
publication, fidelity, or performance behavior changed.

## Validation status

No Cargo, build, rustdoc, runtime, or performance command was run for this
follow-up.  The changed source has been checked with `git diff --check`; the
coordinator must rerun the frozen OOXML consumer gate on the composed source.

## Review return

The initial repair covered the XLSX leaf and its retained normalization source;
it did not yet cover the unified `litchi_xlsx::Error::Join` wrapper. Therefore
the classifier statement above applies to direct leaf errors only. The
follow-up in record 0362ul94ha adds that chained classifier and its actual
wrapped-join conversion regression; the two records must be composed together.
