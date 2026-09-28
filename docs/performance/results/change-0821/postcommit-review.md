# 0821 post-commit replay repair review

Status: **bounded custody and sealing fix approved for the normal follow-up
commit**. This review is static; it ran no Cargo command, workload, reader,
validator, or seal command.

Commit `e540859d7a` recorded the completed 0821 packet, but its
`seal.py --check-head` replay failed because `analyze.py` required the live
source revision to equal the measured base `8312aaa29b`. The retained
`postcommit-replay-failure.json` records that failure, and
`seal-at-e540859d7a.json` preserves the original seal. The failure occurred
after measurement and cleanup; it did not change a report, binary, frozen
input, or derived result.

The bounded reader fix replaces HEAD equality with a
`git merge-base --is-ancestor` check. The measured base must therefore remain
in the current history. The existing source witness remains exact: the live
9,197 production-file and 87 tool-file hash maps equal the retained 0820
repair source, the 0821 build source, and the `8312aaa29b` tree. The
quality-reuse and build descriptors still identify the measured base, so an
evidence-only commit cannot silently substitute source bytes.

The sealing fix compares the staged index or `HEAD` against the packet's
origin base, rather than comparing only the most recent commit. It retains
the exact expected path/hash map and staged/committed byte checks, so the
normal follow-up commit covers the complete aggregate packet change set and
the archived failure evidence.

Static byte checks against `e540859d7a` found no change to `analysis.json`,
`analysis.md`, `native.csv`, `observer.json`, `root-audit.json`, `validate.py`,
or `build.json`. The current production and tool bytes also match the 8312
tree, and 8312 is an ancestor of the current `HEAD`. No production or runtime
harness source, capture receipt, raw report, or measurement was altered.

The fix is limited to post-commit reader custody and aggregate seal path
selection. It is ready for root's normal follow-up commit and final index/
HEAD seal checks.
