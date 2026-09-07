# 0459 — reject lookup reorder and resolve phase profiles

The ordinary ODP attribute cache was tested with local-name comparison before
namespace comparison. The 24-report, 720-sample A1/B1/B2/A2 matrix fails the
predeclared 5% medium/large normal p50 gate; allocation metrics are identical.
R1 normal deltas range from -0.994% to +0.078%. R2 normal-medium p50 rises
13.327%, and allocator-large rises 43.039%; all six adverse timing flags remain.
The production line is reverted exactly. No latency improvement is claimed.

A separate unchanged-source diagnostic build with frame pointers and unwind
tables resolves 94.62% / 91.60% of sampled periods to API markers. Candidate
readback accounts for roughly 59–60% of commit samples; transaction staging
metadata and source-fragment scanning account for roughly 64% and 32–34% of
transaction samples. Family reopen is a low-impact target. These inclusive
estimates overlap and include warmups.

The tested candidate passes 368 ODP tests, including differential lookup and
native preservation coverage. Final owner Clippy, docs and formatting pass;
the broad formatting attempt retains an unrelated existing Keynote difference.
The [bundle](../results/change-0459/README.md) records frozen inputs, profiles,
counters, summaries, rejection, portable verification and owned staging cleanup.
The next work is shared staging tokenization with preserved validation/error
order, alongside candidate slide parsing. The full non-iWork goal stays open.
