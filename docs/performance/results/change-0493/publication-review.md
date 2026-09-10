# Change 0493 publication seam review

This is a read-only audit of the current source-backed OPC integration.  The
line numbers below refer to the checkout observed on 2026-09-10 and are useful
anchors; function names are the stable references.  The audit covers
`crates/litchi-opc/src/source_backed.rs`, its `read_ahead`, `splice`, and
`artifact_restore` children, the DOCX forwarding facade, ADRs 0005, 0006,
0008, 0010, and 0011, and
`docs/performance/results/change-0492/production-next.md`.

The required rule is narrow: an operation that reads through
`IndexedArchive<SourceReader>` for exact preservation, materialization, or
physical authorization must call
`SourceBackedPackage::disable_read_ahead_for_publication()` before its first
such read.  The transition is package-wide and permanent.  It drains admitted
forward reads, invalidates and releases the window, and leaves every later
`SourceReader` request on the exact path.  Metadata-only catalog calls do not
need the transition.  Ordinary semantic `PartView` reads may remain forward
because they return decoded meaning rather than an exact ZIP traversal.

## Current seam coverage

The current worktree already has the following calls.  The first archive use
column names the use that would be unsafe if the call were moved below it.

| Operation and location | First archive-backed use | Audit result |
| --- | --- | --- |
| `into_opc_package_inner` (6293), `to_opc_package_inner` (6345), through `materialize_opc_package_with_accounting` (6376; call at 6381) | `archive.read_session()` and `read_part_with_session` at 6387–6393 | Covered.  Both materializer inner methods call the transition first (6297 and 6349), before source/context or managed checks; the common materializer repeats the idempotent guard at 6381 before opening its read session.  The repeated call protects the helper if a future caller reaches it directly. |
| `write_topology_to_stream` (6436; call at 6441) | Destination comparison (`read_part`) at 6550, followed by relationship/content-types reads at 6981 and 7370 | Covered.  The call is before the empty-plan branch, so the exact no-op also releases a window.  That is safe, although the plan permits an empty operation to go directly to `write_exact_source`. |
| `source_xml_part` (5826; call at 5827), reached by `PartView::source_xml` (3481) | `read_part` at 5847 | Covered.  The returned `SourceXmlPart` is a source-preserving token and must contain an exact decoded source payload. |
| `authorize_precompressed_inner` (6086; call at 6091), reached by `PartView::authorize_precompressed` (3494) and `data_and_authorize_precompressed` (3513) | `archive.metadata_for` at 6114–6116, then verified decoded/precompressed reads | Covered.  This path authorizes a physical compressed transfer and therefore cannot use a forward window. |
| `write_single_part_overlay_to_stream` (7919; call at 7930), reached by the three one-part wrappers (7851, 7869, 7895) | Selected payload read at 7943–7944 | Covered.  The call precedes validation metadata and the exact/no-op comparison. |
| `write_part_overlay_with_external_relationship_removals_to_stream` (8002; call at 8009) | Relationship `read_entry` at 8066 | Covered.  This branch bypasses the one-part helper when removal IDs are present and is an easy path to miss. |
| `write_part_overlays_with_external_relationship_removals_to_stream` (8106; call at 8111) | Relationship `read_entry` at 8266 | Covered.  The call precedes sorting, relationship lookup, and all selected member reads, including the mixed empty/non-empty removal cases. |
| `write_part_overlays_impl` (8418; call at 8428), reached by replacement and deletion wrappers (8343, 8362, 8394, 8409) | Validation session at 8514 and selected payload reads at 8519–8520 | Covered.  The call is before the empty-plan branch and before selected payload comparison. |
| `write_changed_overlays_with_appended_inner` (9410; call at 9418), reached by all changed-overlay preservation helpers | `archive.preservation_index_with_limits` at 9431 | Covered and valuable as a common defense.  It runs before `monitor_publication`, scratch setup, and the preservation index. |
| `prepare_source_part_splice_with_replay_handle` (1139; call at 1146) | Target metadata at 1165–1168 | Covered for replay preparation. |
| `prepare_source_part_splice_inner` (1285; call at 1294), reached by `prepare_source_part_splice_with_fragment` (1072) and `prepare_source_part_splice` (1101) | Target metadata at 1313–1317 | Covered for decoded and preallocated fragment preparation.  `allocate_source_part_splice_fragment` (1000) only allocates memory and has no archive read. |

The three public `SourcePartSplicePlan::write_to_stream*` entry points (491,
497, and 510) converge on `write_to_stream_inner` (578).  The inner function
now calls the transition directly at `splice.rs:584`, before its
source/context fences, no-op branch, preservation index (around 660), or
verified archive reader (around 707).  The replay and decoded/preallocated
plan constructors also transition at `splice.rs:1146` and `1294`.  This direct
writer guard closes the previous constructor-only seam and remains harmless
for an exact no-op: the no-op copies a `SourceArtifact` directly, but a
publication still permanently releases any package-owned window.

`SourcePartSplicePublication::write_inverse_to_stream` (1498; guard at 1503)
and `SourceBackedPackage::restore_source_artifact_to_stream` in
`artifact_restore.rs` (42; guard at 49) read only `SourceSnapshot` bytes through
the exact-artifact helpers.  They do not read `IndexedArchive`, so they have no
strict “before archive read” gap.  The broader package-publication policy is
now implemented as well: both transition the current package before source
freshness checks or the first source copy, including exact no-ops, and
`exact_restore_noop_permanently_releases_the_opt_in_window` covers the restore
case.  `SourceArtifact` itself remains independent because it intentionally
owns no package cache or read-ahead state.

## Deliberate exemptions

Do not add the transition to every `PartView` method.  `PartView::data`,
`data_with_observer`, `data_with_accounting`, the verified decoded-reader
methods, and `stream_to` are semantic payload operations and are explicitly
allowed to use forward mode by the 0492 plan.  Their low-level helpers
(`read_part`, `with_verified_decoded_reader`, and `stream_part_to`) are shared
by both semantic and publication callers; a blanket transition there would
silently disable the intended warm semantic path.

`rels`, `iter_parts`, `physical_member_names`, `has_encrypted_entries`,
`part`, `main_document_part`, `non_part_members`, cache diagnostics, and
`validate_topology_source_boundary` inspect already-indexed metadata.  They do
not perform a source-provider payload read.  Private metadata helpers such as
`has_signature_infrastructure`, `build_physical_member_lookup`,
`source_entry_id_case_insensitive`, `validate_*_limits`, and
`read_content_types_xml` are currently reachable only below the guarded
topology or overlay paths.  The common high-level calls above must remain in
place if that call graph changes.

`SourceArtifact::fingerprint`, `write_to_stream`, and
`write_to_stream_with_accounting` (around 2748–2851) are exact direct snapshot
operations.  They intentionally do not carry the archive-owned read-ahead
`Arc`; this is the ownership boundary required by 0492 and ADR 0005.

## Concurrency and lifetime critique

The settled `read_ahead.rs` design uses a short `Mutex<ModeState>` admission
gate with an `Option<ThreadId>` forward owner and transition owner.  A
`ForwardReadLease` serializes the one retained-buffer operation, while the
buffer and reservation move out of the state lock before the provider and
source-version fences.  `FillLease` restores the buffer on unwind, and the
lock helpers recover poisoned state so publication can still fail closed and
release memory.  A same-thread recursive read or transition receives a typed
refusal.  `begin_transition` sets the exact mode and monotonic admission flag
before waiting for the admitted owner, then `take_window` drops the resources
before the caller enters ZIP callbacks.  No hidden thread-local ownership list
is used.
`disable_for_exact_publication` takes the mode admission lock, sets `Exact`,
takes the window and reservation under the state lock, drops those values after
the state lock, and releases the mode lock before the caller enters ZIP
preservation.  The mode-to-state lock order is consistent in read, transition,
and diagnostics paths.  The exact branch drops its mode guard before calling
the provider.  These points match the intended monotonic publication design.

The gated-fill and racing-reader test
(`exact_transition_waits_for_a_gated_fill_and_future_reads_are_exact`,
`read_ahead.rs:1417`) prove that a transition waits for one forward operation
already admitted by the mode gate, that readers racing after the monotonic
fence take the exact path, and that no forward read starts after the fence.
The diagnostics, version-callback, panic, and poison tests cover the callback
and unwind cases that were previously open.  The adapter checks
source/context after closing admission and again after taking the exact state;
the package transition helper now enters that admission gate before invoking
its source/context fences.  A post-admission failure leaves the mode exact and
propagates before any archive callback.  Source and context fences remain
outside adapter mutexes, and the buffer and reservation are dropped before
caller-owned ZIP callbacks.

The shared `Arc<ArchiveReadAhead>` on every `SourceReader` clone is the right
ownership shape: the `IndexedArchive` reader and package handle observe one
mode, and `SourceArtifact` clones observe none.  The physical input helper
also commits accepted bytes rather than requested window capacity, which is
consistent with the budget contract.

The adapter-level owner check is before its first provider/version fence: the
version-callback test confirms that a recursive `read_at` or transition is
refused while a forward lease is active.  `SourceSnapshot::version()` here is
the captured, non-callback value; only `ensure_current()` calls the provider.
At the package seam, `disable_read_ahead_for_publication()` now directly
invokes `ArchiveReadAhead::disable_for_exact_publication()`.  The materializer
inner methods, the splice writer, and inverse/restore paths call that helper
first, before their own source/context checks.  The integration test
`package_publication_callback_refuses_before_recursive_source_version_fence`
(`tests/source_read_ahead.rs:481`) confirms that a provider callback attempting
`to_opc_package()` receives the typed refusal before causing another source
version callback.  This closes the prior callback re-entry and lock-order
blocker; adapter mutexes are still not held during provider callbacks.

## Tests to extend

The adapter tests in `crates/litchi-opc/src/source_backed/read_ahead.rs` cover
policy bounds, a hit and crossing prefix, managed input/memory release,
empty/EOF behavior, an unmonitored source change, a blocked fill followed by
transition, readers racing after the transition gate closes
(`:1417–1488`), provider diagnostics during a fill, version-callback
recursion, provider panic cleanup, and poisoned-state recovery.  The provider
panic and reentrancy tests exercise the adapter directly, while the package
callback test below proves the package helper's ordering.

`crates/litchi-opc/tests/source_read_ahead.rs` now exercises exact/default
policy, managed accounting, exact artifact save, borrowed materialization,
the package callback ordering (`:481–510`), source-XML capture diagnostics
(`:513–537`), exact restore release (`:540–575`), one-part and batch
relationship removal, batch replacement/deletion, typed and precompressed
topology additions, expected-artifact splice preview, managed
changed-overlay publication (`:1070–1147`), artifact lifetime, and
stale-source refusal.  The managed changed-overlay test includes a physical
call log and first-write memory assertion.  The precompressed topology test
necessarily authorizes its token first, which already enters exact mode, and
the splice test prepares its plan first, which also transitions; neither
independently proves the later topology or writer guard.  Source-XML now
proves that diagnostics stop changing and the retained window is released,
but it still lacks an independent physical range call log.  These are
validation-coverage caveats rather than known unguarded seams.

Existing source-backed tests around topology/relationship publication
(approximately 12159–13480), overlays and managed output
(approximately 16626–18182), and splice replay in
`source_backed/splice.rs` (approximately 3620–4583) provide fixtures and typed
failure assertions to extend.  The materialization test
`publication_disables_read_ahead_before_output_and_keeps_future_reads_exact`
proves one exact transition and post-transition semantic read; the additional
branch tests establish output preservation, while seam-specific diagnostics or
physical-call assertions are still needed to prove each transition itself.

The DOCX facade's policy constructors and diagnostics forward correctly to OPC
(source-backed facade around lines 499–528 and 808–809).  The new
`crates/litchi-docx/tests/source_read_policy.rs` smoke case warms a managed
semantic read, prepares a tail-append publication, verifies the transition
released the window, and reopens the result.  It compares semantic physical
call counts with an exact control; a lower-level range call log would still
make the exactness assertion stronger.

## Review conclusion

Current `SourceBackedPackage` coverage is complete for the listed materializer,
overlay, topology, preservation, source-XML, precompressed-auth, and direct
splice-writer paths, with the relationship-removal branches explicitly
guarded.  The direct splice writer guard is present at `splice.rs:584`, and
replay/decoded preparation transition at `splice.rs:1146` and `1294`.
Inverse/restore now also apply the chosen package-wide release policy first
(`splice.rs:1503`, `artifact_restore.rs:49`).  The lease rewrite is
structurally settled: it removes the old `RwLock`, keeps ownership bounded to
one `ThreadId`, and releases the moved window before publication ZIP work.  No
known unguarded archive-publication seam or transition-order blocker remains.
The remaining caveat is validation breadth: source-XML and some branch tests
prove diagnostics/output but do not independently record every physical range
call after their own transition.  The legacy constructors and
`SourceReadPolicy::default()` were audited and remain exact; no unwanted
default opt-in was found.

## Queued-reader cancellation assessment

The original lease implementation had a managed-waiter gap: when a forward
fill owned the admission slot, `ArchiveReadAhead::admit_forward()` used an
untimed condition-variable wait before the reader reached its normal
`check_before_read()`.  An already-cancelled or subsequently-cancelled
managed reader could therefore remain blocked until the provider returned.
That did not allow stale bytes to publish, but it violated the cancellation
responsiveness established by the source-cache waiter contract.

The current implementation passes the snapshot context into admission.  A
managed waiter checks it and polls with a 10 ms `wait_timeout` while waiting
for either a transition owner or a forward owner (`read_ahead.rs:721–780`).
The normal pre-provider check remains after admission, preserving the source
version callback re-entry guard.  The focused tests
`managed_waiter_polls_cancellation_while_fill_remains_blocked` and
`already_cancelled_managed_waiter_does_not_wait_for_a_blocked_fill`
(`read_ahead.rs:1515–1626`) verify subsequent and already-requested
cancellation before the gated provider is released.

The cooperative limitation applies to the provider call already in progress
and to the transition drain: a synchronous provider or cleanup path cannot be
asynchronously interrupted.  It does not justify holding a queued managed
reader hostage to that call.  The queued reader now returns typed
`OpcError::Cancelled`; unmanaged readers have no cancellation context and
retain their untimed wait behavior.
