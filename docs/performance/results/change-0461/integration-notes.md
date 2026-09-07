# Integration and measurement scope

The baseline is revision `05f432d48`, including the accepted 0460 shared staging
scan. Exact compiled source manifests and retained executable bindings identify
the measured epochs. Source text artifacts and `source-delta.json` bind the
complete before/after code without introducing additional Rust build inputs.
The hypothesis is independent of the rejected 0459 predicate reorder.

Only ODP attribute-cache matching/decoding changes. The namespace-first predicate,
cache/source order, iterator advancement, lazy normalization, malformed/duplicate
attribute reachability and error strings are preserved. Cached invalid values
remain cached; a newly reached invalid matched value advances the iterator but
returns before appending that attribute, as in the baseline. Drawing-attribute
harvesting and namespace lifetimes are unchanged. No new index, allocation,
public API, dependency, clock, thread, unsafe code or validation bypass is added.

The new focused test uses the unchanged direct raw-attribute reader as an
independent value/error oracle. `Parser::get_attr` wraps the optimized cache and
would not be independent; that test-reference issue was corrected before gates.
Existing malformed, duplicate, namespace-shadowing, foreign/unknown namespace,
first-match, normalized-value and harvest-order corpora remain applicable. Full
owner tests also exercise native fixture XML, no-ops, patch replay/inverse and
preservation. These are library-level native checks, not Office GUI roundtrips.
ODP has no dedicated fuzz target; the unchanged ODF detection and ODT targets
do not exercise this private cache. No unrelated fuzz coverage is claimed.

Normal and allocator binaries use identical release settings across variants.
The 24-report / 720-sample A1/B1/B2/A2 matrix is the acceptance boundary; four
separate large normal phase reports (120 operations) and two 100-operation
whole-process counter runs are mechanism diagnostics. All workloads run serially
on CPU 2. Setup, oracle preflight, row validation and reporting lie outside the
lifecycle clock; process RSS and perf counters include process-lifetime work.
Normal and allocator timings are separate. Quantiles use midpoint p50 and
nearest-rank p95/p99; independent bootstrap intervals retain seeds and cannot
remove run-order effects or correct for multiple comparisons.

The frozen host record reuses the 0460 host schema with the current revision
and recording time. Its protocol hash was bound before baseline capture.
An assembly check can observe an in-progress source tree while inspecting the
retained baseline executable: its executable binding, not the ambient source
manifest in the command wrapper, identifies the inspected code. Both assembly
outputs record the exact objdump argv and executable SHA-256.

Applicable final gates are release all-feature owner and harness tests,
warning-denied all-target owner Clippy and rustdoc, scoped formatting, crate
boundaries, bundle verification, fresh-copy replay and summary-tamper rejection.
No performance claim extends to cold caches, range providers, bounded-worker
scaling, other CRUD operations or native application behavior. Registry counts
remain 439 selectors / 36 defaults. The full non-iWork goal remains open.

Final disposition: the candidate passes 372 ODP and 387 harness tests (one
ignored). It is rejected by the frozen practical latency gate and the sole
Rust file is restored exactly to `05f432d48`. Restored-source strict Clippy,
rustdoc, scoped formatting and boundaries pass, as do precleanup, fresh-copy
replay and summary-tamper rejection. Cleanup removes four retained executables
totaling 233,048,120 bytes. No new production optimization is retained.
