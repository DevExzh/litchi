# Source, correctness and measurement review

The root works locally after earlier delegated agents reached usage limits.
The baseline is a4a6f7df437123d166d8b901ada573a44b7aadd9 (0441). The accepted ADR
tree remains c950b6c8be822561b498d7bbe87c460873dcbf49 from the prior complete
read. The scope is seven ODP source files listed in source-files.json.

Settings, declarations and page metadata retain their own constructors,
byte gates, state transitions, namespace resolution, lazy attribute handling,
collection bounds and final validation. Individual public parse functions now
drive those states. The shared private staging driver borrows each event once
and steps all successful states through one NsReader with trim_text(false).
Settings errors return immediately. Declaration/page errors are deferred;
page stepping stops after declaration failure. XML read failures still belong
to settings. EOF finalization uses explicit locals in historical priority order.
Lower-priority work can occur before a later higher-priority failure; existing
finite input/model bounds remain, but invalid-input performance is unmeasured.

The mutable constructor takes the grouped outputs. Semantic slides, source
fragments, origin mapping, exact preserved markup, detached editing, publication
and readback remain on their existing paths. No public API, dependency, unsafe
code, executor, ambient I/O or cache lifetime changes were introduced.

Each old complete parse body is retained byte-for-byte behind cfg(test),
with unchanged helpers. The verifier independently checks these bodies against
baseline source. Three tests perform 353 differential input comparisons and
two explicit competing-error assertions. Coverage includes namespace rebinding,
default namespaces, entity/CDATA handling, unknown markup, malformed XML and
attributes, finish-time errors, 8 MiB exact/+1 XML and 65,537-page bounds.
All 355 ODP tests pass with zero ignored. The full harness passes 368 tests
with one existing ignored test. Strict owner Clippy passes all targets with
warnings denied and without dependency linting.

The existing ODP suite includes LibreOffice/odfpy settings parse-save-reopen
fixtures and a real-producer tdf169979.odp transaction, text-box edit, commit
and exact preservation check for untouched members, manifest, config, OLE and
compressed spans. No fresh LibreOffice process or native visual rendering was
run. Available ODF fuzz targets cover detection and ODT; the facade fuzz targets
are iWork-only, and ODP has no fuzz target. Those unrelated targets were not
run or presented as ODP coverage. Differential adversarial cases are not
coverage-guided fuzzing. No new unsafe or concurrent code warrants Miri/loom.

The initial prototype cloned events and passed three focused tests. Before
candidate measurement, it was tightened to borrowed events and the full ODP
suite passed. Initial recipe/template files remain under draft-history; only
the final candidate is used by after-build and all B measurements. Original
parser reference-body hashes are retained separately. No measured attempt was
excluded or substituted. All root CPU jobs were serialized, and source switches
occurred only after terminal jobs. A recovered harness handle was audited by
its live process IDs and final receipt before another CPU job started.

The predeclared normal p50 gate passes. Medium/large gains are 9.830–13.307%
in both repeats; tiny gains are 6.923–7.020%. The allocator/peak gate fails.
All 60 allocator observations per role/shape agree, saving only 38 calls and
11,902 requested bytes per size. Peak and retained bytes are unchanged.
All five repeat flags remain visible. The candidate large p99 repeat rises
7.091%, while paired R2 p99 falls 2.268%. No general tail-speedup or
instrumented-timing claim follows. Profiles include setup, warmups and oracle
work, and retain 13 before/14 after addr2line warnings in each conversion.
Zero L1 counters support no cache-miss claim.

Broader workspace/ODF feature checks, warning-denied rustdoc, explicit-file
formatting, boundary checks and final portable proofs are bound by
checks-policy.json. Completed receipts, not this checklist, are authoritative.
The original report oracle, protocol and capture/profile drivers remain frozen.
The full non-iWork goal, registry and representative coverage scope remain open.

The non-iWork workspace all-targets/no-default-features check with litchi/odf,
ODF all-targets/all-features check, warning-denied owner rustdoc and explicit
seven-file Rust 1.98.1 formatting check all passed. Workspace checks reported
no warnings. Descriptive p50 bootstrap intervals do not overlap in any matched
normal pair; temporal grouping remains a limitation of this ABBA experiment.

The final boundary audit passed for 64 packages/240 internal dependency
declarations with the same 14 explicit migration debt items. Final source
review confirms the checkout matches the measured candidate, all four retained
executables match build hashes, the accepted ADR tree and prior sealed bundle
are unchanged, and user-owned GOAL.md retains its pinned digest.

Portable copied controls passed before and after cleanup. All 11 precleanup
and 12 postcleanup independently corrupted copies were rejected after inventory
refresh, including the cleanup GOAL-digest mutation. Cleanup removed only the
four bound executables under /tmp/litchi-goal-0442-binaries (1,830,317,656 bytes),
preserving both build-cache directory identities and user-owned GOAL.md. The
probe copies cleaned themselves up. Compression and inventory are resealed
before final read-only verification. No tests, checks or measured attempts failed.
