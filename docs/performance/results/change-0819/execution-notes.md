# 0819 execution notes

This batch starts at `0c784f7aecda2a38bbe6cc98a684ce68da1f0dc8` after the
0818 DOCX relationship-preservation repair. The preceding 0817 attempt never
reached timing because independent artifact admission failed; its evidence
remains unchanged. The 0818 evidence was untimed and does not supply this
batch’s release binary or output admission.

Before execution, root checked the commit history, working tree, 9,197 tracked
production files, 87 tracked harness files, corpus identities, both Cargo locks,
and all 35 normative document hashes. Three unrelated workspace files remain
outside this batch. Static review corrected a stale source census and kept
full `cargo test --all-features` without `--all-targets`, preserving doctests.
No failed execution was discarded by those preflight edits.

The settled protocol distinguishes the plan’s nearest-rank process quantiles
from the raw harness’s integer-midpoint p50. Native and instrumented observer
latencies have separate roles. No production or harness source changes and
no historical speedup comparison are part of this baseline.

Root owns and serializes all Cargo, exporter, and benchmark execution.
Subagents prepare and review scripts; result readers run only after every
capture handle is terminal. Every capture retains commands, source witnesses,
input identities, logs, RSS, raw samples, and hashes.

All six quality gates completed successfully; the full harness test log has
641 passed, one ignored, zero failed across 28 summaries. The three serial
release builds completed. Fresh six-case/thirty-policy artifact audit and ZIP
preservation admission passed, followed by all 12 qualification reports, 72
native reports, and 24 observer reports. All root execution handles were
terminal before offline replay was delegated.

The first two offline reader attempts failed on reader-schema assumptions:
audit output identity and harness configuration. Their logs were retained;
reader-only corrections admitted the same unmodified raw records on attempt 2.
The reader's write/check/validate and root's independent validate passed.
Root independently recomputed all twelve native nearest-rank p50/p95/p99 and
bootstrap intervals, replayed the artifact auditor and ZIP preservation checker,
and confirmed the 0817 and 0818 packet payloads remain unchanged. Resource
summary medians use only the six observer samples per case, not qualification.

Independent results review passed. Owned target/scratch cleanup verified all
three stable binaries before removal and deleted 7,012 files totaling
7,452,891,288 logical bytes. Post-cleanup validation found and corrected the
reader's dictionary-representation sorting assumption; the failed attempt is
retained and the same evidence passes with binaries absent.
