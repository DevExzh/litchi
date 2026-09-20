# Byte-position text integration

Status: implementation underway; no support or validation claim yet.
Baseline is the committed ordinary-text batch recorded in baseline.json.

Root checked the local ODF 1.4 Part 4 section 6.7: exactly FINDB, LEFTB,
LENB, MIDB, REPLACEB, RIGHTB and SEARCHB are specified. The standard explicitly
leaves byte units implementation-dependent. This batch selects semantic UTF-8,
backward snapping of interior starts and complete-scalar clipping as documented
in contract.md. It does not emulate a native DBCS code page.

The existing text dispatcher already controls scalar evaluation, matrix
argument scheduling, streaming reference selection and Empty-to-Text coercion.
Extend that classifier and argument map, keeping the ordinary demand cache
conservative. Add a focused bytes module and reuse shared text conversion,
source-error ordering, output leases and Unicode SEARCH matching. Do not
materialize ranges or decoded-character vectors. A UTF-8 boundary adjustment
needs at most three continuation-byte steps; prefix/suffix clipping can use
borrowed source boundaries. SEARCHB may reuse the folded matcher and map the
matched original scalar position back to a UTF-8 offset with charged work.

Validation must include all seven names in scalar/value Scalar/Matrix modes,
every UTF-8 width and interior start, empty/end sentinels, overflow/domain
coercions, Missing/Complex/error precedence, reference-list refusal, projected
branches, borrowed-text ownership, output/storage/work limits, cancellation
and source fences. Independent expected values must come from the declared
UTF-8 rules, with native code-page differences disclosed separately.

Before commit, run isolated locked gates and a matched performance profile
against the baseline. Retain source/fixture/harness hashes, check the timed
consumer as well as preflight, and assert cancellation repeats explicitly.
Historical captures remain evidence; temporary builds and worktrees are removed.

## Integration progress (not a freeze receipt)

The seven names now route through the existing text classifier, argument
metadata and scalar kernel bridge. The independent integer/string UTF-8 oracle
passes all 1,376 observations after root integration. Root replaced decoded
prefix/suffix clipping scans with at most three UTF-8 boundary adjustments;
input work charging remains in place. Identical integer conversion and owned/
borrowed slice helpers are shared with the ordinary text module through
module-local visibility. SEARCHB's scalar-to-byte result mapping uses a checked
single scan. Temporary debugging output has been removed.

Root library clippy passes with warnings denied. These focused observations
are provisional: semantic/resource review, native divergence evidence, complete
isolated integration gates and matched performance results are still required
before the batch can be committed as complete. The gate runner and isolated
checkout are prepared; the retained dependency lock is copied unchanged from
the preceding text batch. The ambient workspace lock is not replaced.
