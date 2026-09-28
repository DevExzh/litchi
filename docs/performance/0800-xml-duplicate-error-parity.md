# 0800 — duplicate-error parity at the bounded XML handoff

The five checked-attribute helper copies now preserve quick-xml's duplicate error
for unusual keys beginning with `=` after the 32-name handoff. This is a
correctness repair, with no performance improvement claim.

The previous batch identified replay as a cost to remove. While preparing a
no-replay candidate, source review found an existing error-ordering discrepancy.
The [packet](results/change-0800/README.md) retains the source-only candidate but
defers its build and measurement until it is rebased on this corrected baseline.
The performance experiment's unused draft harness was removed; no candidate
native or counter captures were run.

## Defect and correction

The checked helper promises quick-xml's items and first error, with the same byte
positions. quick-xml checks the first 32 names; the bounded helper then uses an
ordered map. In that late path, the unchecked iterator parses an item first. If
it returns a value error, `duplicate_before_value` recovers the key and translates
the error to `Duplicated` when the key was already seen.

quick-xml consumes the first non-whitespace key byte before seeking the next `=`
or XML whitespace. The recovery helper searched from that first byte. Thus
`=n0="first"` has the raw key `=n0` in quick-xml, but error recovery treated a
later occurrence as an empty key and failed to find it in the map. This layer
does not validate XML Names; it must preserve the lexical behavior of its
quick-xml dependency, including these unusual prefixes.

The retained probe builds `e`, a sequence of distinct `=nI="I"` attributes, and
the malformed duplicate `=n0="unterminated`. Its observations are:

| Successful prefix | Original helper | quick-xml and corrected helper |
|---:|---|---|
| 1 | `Duplicated(10, 2)` | `Duplicated(10, 2)` |
| 31 | `Duplicated(292, 2)` | `Duplicated(292, 2)` |
| 32 | `ExpectedQuote(319, 34)` | `Duplicated(302, 2)` |
| 33 | `ExpectedQuote(329, 34)` | `Duplicated(312, 2)` |
| 34 | `ExpectedQuote(339, 34)` | `Duplicated(322, 2)` |

The correction searches for the delimiter starting one byte after the first
non-whitespace byte, while retaining that byte in the returned name. It changes
the canonical OPC helper and the OLE-common, signing, XLDM and XML-minifier copies.
The shared test is the only other production file changed. No public API,
dependency, iterator layout, ownership, or duplicate-state strategy changes.

The error-remapping guard still handles only `UnquotedValue`, `ExpectedValue`,
and `ExpectedQuote`. Missing equals signs retain `ExpectedEq`. Valid parsing,
borrowed values, XML whitespace, fused exhaustion, and the bounded ordered-map
algorithm are unchanged. The existing late path still scans the value before
remapping its error; this fix corrects error priority and makes no claim to
remove that scan.

## Verification and scope

The controlled probe compiles the exact original and corrected helpers into one
locked executable, checks both against quick-xml, and asserts the original
handoff discrepancies. It is a correctness probe, not a timing experiment.

The new shared regression test exercises accepted prefixes 1, 31, 32, 33, 34,
and 64; five raw key forms; and three whitespace separators. It compares 540
duplicate cases and 270 lexical/nonduplicate controls with quick-xml per helper
copy. Duplicate cases cover valid, unquoted, missing and unterminated values,
exact positions, cloning immediately before the error, and repeated terminal
`None`. Existing differential/random, canonical-copy, borrowed-value and bounded
comparison tests remain in force.

The first full test attempt failed because a new test assumed `==n0` was a
complete accepted key. quick-xml treats its second `=` as the delimiter and
rejects the first value, so that premise was wrong. The final matrix uses
`=long_name` for that valid-prefix case. The failed source and logs are retained;
the production parser correction was unchanged.

All four final gates pass: formatting for the six changed files; the locked
controlled edge probe; full default-feature tests for the five affected production
crates; and all-target Clippy with warnings denied. The test log records 1,538
passed, zero failed, three ignored, and zero filtered tests across 67 suites.
The ignored tests require an externally supplied decoded LibreOffice corpus,
an independently generated ZIP64 corpus, or explicit checked-in asset regeneration.
This is not a full-workspace or all-features verification claim.

The adopted change is limited to the five key-recovery helpers and their shared
regression test. The no-replay candidate remains unbuilt, unmeasured and unadopted;
its archived baseline predates this repair. The next performance experiment must
rebase it and repeat its semantic gates before using any measurements. The 0799
rejection remains in force. No speedup, resource reduction, or performance/CRUD
coverage is claimed from this correctness packet.

The owned build target was removed after the gates passed (3,711,347,610 logical
bytes). Both exploratory and controlled executable identities remain in the
cleanup witness. All other production files, the 35 architecture inputs,
unrelated workspace files, and existing worktrees are preserved. The correction,
failed and final checks, reviews, and deferred source archive are sealed together.
