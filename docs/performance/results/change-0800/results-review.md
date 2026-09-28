# 0800 final results review

This is a bounded correctness review of the finished 0800 packet. I reconciled
the main report with the source review, correction manifests, controlled edge
output, final quality receipts, test summary, and retained failure archive. I
did not run Cargo, tests, a replay, a native capture, or a profiler.

## Defect and correction evidence

The correction manifest records six intentional source changes: the five
checked-attribute helper copies and the shared OPC test file. The helper diff is
limited to `name_at`: it retains the first non-whitespace byte in the raw key
and starts delimiter search at the following byte. The duplicate-remapping
guard still covers only `UnquotedValue`, `ExpectedValue`, and `ExpectedQuote`;
`ExpectedEq` remains a lexical error. No API, dependency, iterator layout,
ownership, or duplicate-state strategy changed.

The final controlled edge executable contains the original and corrected
helpers and compares both with quick-xml. At accepted prefixes 1 and 31, all
three report `Duplicated(10, 2)` and `Duplicated(292, 2)` respectively. At
prefixes 32, 33, and 34, the original reports `ExpectedQuote(319, 34)`,
`ExpectedQuote(329, 34)`, and `ExpectedQuote(339, 34)`, while quick-xml and the
correction report `Duplicated(302, 2)`, `Duplicated(312, 2)`, and
`Duplicated(322, 2)`. The run asserts these relationships; it does not merely
print them.

This establishes error priority and exact byte positions for the affected
handoff. The correction still obtains the lexical error by asking the
unchecked iterator for the item, then remaps it, so it does not claim to avoid
scanning a malformed value. The helper also mirrors quick-xml's raw byte
grammar and does not validate XML `Name` syntax. A leading `=` is part of the
key, and the existing XML whitespace set remains space, tab, carriage return,
and line feed.

## Semantic test and quality gates

The added shared regression test covers six accepted-prefix counts (1, 31, 32,
33, 34, and 64), five raw key forms (`=`, `=n0`, `=long_name`, `ordinary`, and
`é`), three whitespace separators, and six malformed or valid value forms.
That is 540 duplicate-before-value cases and 270 lexical/nonduplicate controls
per helper copy. It compares full yielded/error sequences with quick-xml,
including exact positions, then checks cloning immediately before the error and
repeated fused `None` results. Existing random, differential, borrowed-value,
canonical-copy, and bounded-comparison tests remain in the final run.

All four final quality gates exit successfully: formatting for the six changed
files, the locked controlled edge oracle, default-feature tests for the five
affected crates, and all-target Clippy with warnings denied. The retained test
summary records 1,538 passed, zero failed, three ignored, zero filtered, across
67 suites. The three ignored tests are explicitly scoped: one needs an external
decoded LibreOffice corpus, one needs an independently generated ZIP64 corpus,
and one is an explicit checked-in asset regeneration test. None is an ignored
XML-attribute correctness test.

The first quality attempt is correctly retained as a setup failure. Its single
failure was the new matrix's `==n0` premise: quick-xml treats the second `=` as
the delimiter, so that spelling does not form the intended valid prefix. The
final matrix uses `=long_name`; the failed and final manifests show identical
parser-helper hashes, with the correction confined to the test fixture and
diagnostic assertion text. This does not hide a production failure or rerun a
captured performance case.

## Scope and disposition

The correction fixes an existing late bounded-handoff parity defect. It keeps
the ordered-map fallback and its value-scan/remapping behavior intact, and it
does not validate XML names or make a performance claim. The source-only
no-replay candidate is based on the pre-correction commit, remains unbuilt and
unmeasured, and must be rebased and pass fresh semantic gates before any later
experiment. The 0799 rejection and the no-adoption policy remain in force.

The packet's custody records preserve the before/after sources, controlled and
exploratory binary identities, failed quality attempt, final gates, and the
deferred candidate. Cleanup records removal of the owned target while retaining
those witnesses. I found no correctness or results blocker in this packet.
