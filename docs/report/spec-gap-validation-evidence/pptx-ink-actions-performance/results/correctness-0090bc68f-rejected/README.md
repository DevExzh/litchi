# Rejected 0090 PPTX correctness capture

This bundle preserves the fresh 0090 release preflight/build/host/matrix
receipts, the release regression stdout/stderr, and the six separately
authorized individual-lane diagnostics. It is rejected as a correctness
matrix result.

The retained-baseline boundary fix passed its Rust regression and removed the
prior systematic metadata failures. The matrix then failed six lanes:
package_read_large_distinct, package_read_near_distinct,
package_apply_medium_distinct, package_inverse_small_shared,
package_inverse_case_equivalent, and stale_content_type. The fail-closed
operator therefore executed zero individual lane commands. The six diagnostic
lanes were run afterward only under explicit follow-up authorization against
the immutable 0090 binary; they are diagnostics, not admitted matrix evidence.

Diagnostic predicates and proposed minimal fixes are recorded in
retention-manifest.json. The distinct lanes failed preservation_ok because the
receipt reduced per-target pointer identity to shared when repeated inbound
edges existed. The inverse lanes failed retained_baseline_balance_ok by 22
bytes because inverse application used the retained package. The stale content
type lane matched Error::ContentType and preserved source/manifest bytes, but
baseline_reopenable did not accept the intentional typed read refusal.

The retained copy contains 38 capture files,
2 regression files, and
19 diagnostic files:
59 files and 4970923
bytes. The exact source paths, retained paths, sizes, and SHA-256 values are
in retention-manifest.json.

The captured release binary had SHA-256
54463a17367957e8ee6c7a9acfcce27943ba4b480b8cc1196e3da4ae652c004f.
Its original path and both target inventories remain in the manifest. No /usr/bin/time,
timed lane, native PowerPoint acceptance, or speedup claim was produced.

Root audited all 59 raw files against Git bytes committed in `2e6140b69`,
verified the binary hash and both target inventories, then removed the owned
external receipts and targets (709,872,828 target bytes). The clean source
worktree remains. A post-fix capture requires fresh targets and results.
