# 0704 design: bounded retained slide MCE projections

Status: production design review. This document proposes the private candidate
for the measured 0703 seam. It makes no 0704 production or performance claim;
the only source change in this batch is this design record.

The 0703 trace makes the target specific. On the real 13-slide package, a
changed one-slide commit repeated 17 of the 18 successful default-profile MCE
projections from capture; the two-slide workflow repeated 16. The complete
capture's outputs accounted for 543,070 bytes of observed `Vec` capacity,
including the presentation output; the slide-only total was 538,356 bytes.
These are logical capacity totals from the diagnostic, not RSS or allocator
measurements.
The candidate below retains those owned slide projections across the immutable
capture snapshot and a changed transaction. It deliberately leaves the small
presentation/catalog seam, presentation notes scans, and the `from_part` path
out of scope.

## Contract that governs the candidate

ADR 0005's accepted amendments at lines 2291–2325 permit a bounded memo only
when it is keyed by an allocation the owner already holds, keeps a strong
reference that closes the ABA hole, names allocations in that owner's package,
is rebuilt or projected rather than mutated, and treats a miss as ordinary
recomputation. The same record requires fallible reservation and a debug/test
value check.

The final amendment at lines 2327–2373 classifies bytes held by a returned
value as retention. Its ceiling must be a member of the operation's finite
policy, be intersected when policies meet, and be observable and releasable.
The retained allocation must already have been made. Exceeding a ceiling for a
pure optimization silently falls back to recomputation; it cannot become a
new refusal. The opened PPTX path has no execution context, so its `Limits`
member is the interim authority until that path gains an execution context
under ADR 0031.

The existing cross-slide candidate retention is the local precedent. It keeps
the exact allocation taken by owned package ingress in an `Arc<Vec<u8>>`,
charges a finite `Limits` ceiling, exposes a length accessor, releases with a
value-preserving `&mut self` method, intersects source and destination limits,
and rebuilds when the candidate is too large. Its `debug_assert!` reserializes
the candidate in test/debug builds. The MCE memo should follow those ownership
and observability rules while remaining separate from the serialized-archive
retention policy.

## Private value and key

Add a private immutable memo to `opened::Snapshot`:

```text
Snapshot {
    ...
    retained_mce: Option<Arc<RetainedMce>>,
}

RetainedMce {
    entries: Vec<RetainedMceEntry>,  // sorted by (raw_ptr, raw_len)
    charged_bytes: usize,
}

RetainedMceEntry {
    key: (raw_ptr: usize, raw_len: usize),
    raw: Arc<Vec<u8>>,
    processed: Arc<Vec<u8>>,
}
```

`RetainedMce` is private and is constructed only by the context-aware
`from_part_with_name` route, which calls the exact default
`process_ooxml` (`Capabilities::default()` and `Limits::default()`) profile.
There is no persisted memo format or cross-profile constructor, so a manual
profile-version field would add no protection in this binary. If a future
custom profile route is introduced, it must use a separate scoped memo type or
constructor rather than widening this one. No entry from a custom
`process_markup_compatibility` call, semantic text path, font path, notes path,
or `SlidePart::from_part` is admitted. A key is only a lookup index; a hit
additionally checks that the current source pointer and length match the
entry's retained `raw` allocation. The retained raw `Arc` prevents
allocator-address ABA. Length is part of the key, including for an empty
payload. The cached output limit is still checked on every hit, so a hit never
turns a current limit check into a memoized assertion.

The output is retained only when the successful MCE result is `Cow::Owned`.
The owned `Vec<u8>` is moved into `Arc<Vec<u8>>`; it is never cloned. A
`Cow::Borrowed` result is used for this capture and released as today. That
keeps marker-free captures at their existing allocation shape and avoids
retaining a fact whose only value is that a scan found no marker. The 0703
primary target is the owned real-deck projection. No MCE error, root error,
name error, or parsed refusal is stored as an entry.

The raw payload is not copied. The snapshot already owns the package payload;
the `raw` strong reference is an identity witness and another owner of the
same allocation. The added retention charge covers the transformed output and
memo metadata, not a second charge for raw bytes already charged by package
ownership. An entry can be published only when the raw `Arc` is proven to be
an allocation in the snapshot's own package; otherwise the candidate is
dropped and the ordinary path remains authoritative.

## Capture path and exact source observations

The private `SlidePart::from_part_with_name` route should accept a
capture-local `CaptureMce` context. Keep the existing method as a thin
uncached wrapper for other callers if that makes the source census smaller;
only the opened capture calls the context-aware private variant.

For each slide, the context-aware route keeps the established sequence:

1. Validate the slide content type.
2. Call `part.blob()` for the existing 64 MiB raw-size check.
3. Call `part.blob()` again and keep that second exact slice as the source
   witness used by MCE and `SlideRootProof`.
4. Before MCE, search the parent/new memo by that slice's pointer and length.
   On an exact witness hit, take an `Arc` clone of the stored processed output;
   on a miss, run `process_ooxml` over the second observation.
5. Recheck the default MCE input/output bounds on the hit. An entry that does
   not satisfy the current default bounds is treated as a miss, never as a
   limit error.
6. Run `root_name_from_xml`, `c_sld_name_from_xml`, and the optional
   `root_conformance_from_processed` proof over the selected bytes in their
   current order. The proof's raw witness remains the current second
   `blob()` slice, not the memo's owned output.
7. Admit a fresh owned result only after the root/name route has succeeded.
   The notes proof is still recomputed and is never memoized. An invalid proof
   still disables later proof collection exactly as it does today.

The context must not call `blob_arc()` as a new side effect for every slide.
The existing `PartDigests` pass already performs the alias check and retains
verified raw owners while computing the complete revision. The context can
hold a borrowed `(ptr, len)` witness plus a pending owned output during
semantic capture. At finalization, it asks the completed `PartDigests` for the
raw `Arc` for each key. This has three advantages:

* it adds no extra foreign `Part::blob_arc()` call or third `blob()` call;
* a foreign part whose `blob_arc()` does not alias its visible `blob()` is
  automatically omitted by the existing `memo_key` gate; and
* every published entry's `raw` owner comes directly from the current
  package's already-computed digest memo.

On a parent lookup, the retained raw `Arc` itself proves that the pointer has
not been recycled. If the current source slice does not match its pointer and
length, the lookup misses. If final `PartDigests` has no matching owner, the
provisional output is not published in the new memo. Thus a foreign or
detached owner can only take the uncached route and can never make a memo outlive
the package it describes.

The context's pending entries are private and capture-local. A parent hit
returns a cloned processed `Arc` to a small projection guard; a fresh owned
`Vec` stays in that guard until root/name/proof checks finish. The guard exposes
only `&[u8]` to the parsers, so the rest of the slide code does not know whether
the bytes were fresh, borrowed, or reused. This avoids an `Arc<Vec<u8>>` to
`Vec<u8>` copy on a hit and avoids moving a fresh vector into an `Arc` until
admission is known.

## Finalization, projection, and commit handoff

After the ordinary capture validation has completed, and after the existing
complete-package revision/digest pass has produced `PartDigests`, finalize the
context as follows:

1. Carry parent entries only when their `(ptr, len)` occurs in the current
   digest memo and the current digest memo's raw owner is the same allocation.
   This drops entries whose allocation disappeared; an allocation from a
   removed slide that remains owned by another current part may remain, since
   the transformation is independent of part identity.
2. Admit each parent hit or fresh owned result only when its checked
   provisional visit charge fits the running capture ceiling and its pending
   vector reservation succeeds. Repeated references are charged individually
   during capture, even when they will later share or deduplicate an output;
   this is a conservative admission gate. Fresh results that cannot be
   admitted stay ordinary owned results, and parent hits that cannot be
   provisionally retained still return their already-validated processed Arc.
3. Sort pending visits by `(raw_ptr, raw_len)`, deduplicate equal keys, ask the
   completed digest memo for each current raw owner, and publish one immutable
   `Arc<RetainedMce>` with the unique final table's actual capacity charge. If
   any cache-only vector reservation or ownership projection fails, drop the
   cache candidate and publish the snapshot with `retained_mce: None`. The
   semantic capture result and its existing typed errors do not change.

The changed transaction passes the source snapshot's private memo only to the
new revision-and-digests capture helper. The existing no-op branch returns the
source snapshot before recapture, so it does not need an MCE lookup. The helper
keeps all ordinary capture checks and only uses the memo for
`from_part_with_name` slide projections.

Publication has two cases. If the candidate is complete-package equal to the
committed snapshot, `Snapshot::rebound_to` projects the committed MCE memo onto
the candidate's `PartDigests` raw owners. A current-slide membership filter is
not required: default `process_ooxml(raw)` is a pure function of the raw bytes
and profile, independent of part URI and content type, while every future
slide hit still reruns content-type, root, name, and notes-proof validation.
An entry that remains owned by an unrelated current part is therefore still a
valid bounded memo entry; it is observable and releasable, and can become
useful again if that allocation is later reached as a slide. Equal bytes in
newly allocated foreign or rebuilt parts do not hit: content equality is
insufficient for this identity-bound memo. If the candidate drifts, the
existing uncached or payload-digest-parent capture path remains in force. A
changed slide's new raw allocation misses; unchanged slide allocations can
carry their processed `Arc`s into the published snapshot. The digest-owner
projection still drops an entry when the allocation is no longer present in
the package at all.

`Snapshot::clone` clones the outer memo `Arc`, just as it clones the package
`Arc`. `release_retained_mce(&mut self)` sets only that snapshot's option to
`None`; it leaves slides, revision, limits, and package semantics unchanged.
Another snapshot or transaction clone may still hold the shared memo, so the
method's documentation must say that it releases this value's reference and
physical bytes are freed only after the last owner drops them. The public
observation should be `retained_mce_bytes() -> usize`, reporting zero when the
memo is absent and otherwise its aggregate logical charge; it never exposes
XML contents. `Debug` reports only the presence/count/charge.

Expose the same value-preserving release operation at each owner boundary:
`Snapshot::release_retained_mce(&mut self)`,
`Transaction::release_retained_mce(&mut self)`, and
`Commit::release_retained_mce(&mut self)`. The latter two delegate to their
owned snapshot/source. Their `retained_mce_bytes()` accessors return zero
after release. A transaction release drops only that transaction's source
reference; a commit release leaves its patch and candidate snapshot valid but
forces later publication to recapture any reusable slide projection.

The package facade does not need a second ambient MCE cache. A caller that
keeps a snapshot or transaction keeps the memo through the authorized
operation. Mutations that do not publish a snapshot never adopt a memo. This
matches ADR 0005's release rule and avoids retaining bytes that a mutable
facade's current graph no longer owns.

## Limits and resource accounting

Keep `Limits::new` source-compatible for its existing six nonzero fields and
add a private-field/public-accessor builder, for example:

```text
Limits::default().with_max_retained_mce_bytes(1 * 1024 * 1024)
Limits::max_retained_mce_bytes()
```

`with_max_retained_mce_bytes(self, usize) -> Self` is additive. Zero is the
explicit optional-retention disable value and does not invalidate the policy;
the original six `Limits::new` arguments remain nonzero and continue to reject
zero exactly as before. The default proposed for 0704 is **1 MiB**. It is
above the 0703 real-deck slide-only logical capacity total (538,356 bytes)
while providing a much smaller bound than the 512 MiB MCE output ceiling.
This is a retention ceiling, not a new MCE processing limit.

When two `Limits` policies meet, `max_retained_mce_bytes` is intersected with
the other six members. Update both private `intersect_limits` helpers (patch
and cross-slide planning), even though cross-slide code does not consume this
memo. A future execution-context implementation should move this charge to
the context rather than create a second budget system.

The charge is a checked logical upper bound for the memory the published memo
keeps. Count the final entry-vector allocation once, rather than adding its
element metadata once per entry and then adding the vector metadata again:

```text
actual entries Vec capacity * size_of::<RetainedMceEntry>()
+ sum over entries (
      processed.capacity()
    + conservative Arc<Vec<u8>> owner/header allowance
  )
+ Arc<RetainedMce> owner/header allowance
```

The implementation should use `checked_add`/`checked_mul` for every term and
reserve the entry vector fallibly before publishing. The exact `Vec::capacity`
is charged, not `processed.len()`, because unused capacity remains resident.
The fixed `Arc` allowances are deliberately conservative logical accounting;
Rust does not expose a portable exact allocator-overhead query.
`retained_mce_bytes()` reports the same charge. This is an explicit resource
accounting unit, not a claim that it equals RSS or allocator bookkeeping.

Pending capture candidates are transient operation storage. The owned output
vector is moved into its final `Arc` without copying, and final entries move
the pending Arc handles; the pending vector and final vector are both allowed
to coexist briefly while finalization is building the table. They must not
both be added to `charged_bytes`: that accessor reports the published logical
table once, using unique entries and the final vector's actual capacity. The
capture path separately keeps a checked `pending_charge` for every admitted
visit, including duplicate references, so provisional output and entry work
cannot grow without the retention ceiling. The pending/candidate vector
metadata is transient workspace and is fallibly reserved; it is excluded from
the accessor and this logical retained-table budget is not a peak RSS claim.

The raw owner `Arc<Vec<u8>>` contributes an entry pointer and the identity
witness but no raw payload charge: the snapshot's package already holds that
payload allocation. The output `Arc<Vec<u8>>` is a new owner around the already
made MCE vector, so its wrapper allocation and the vector capacity are charged.
Parent outputs shared into a new snapshot are counted once in that new memo's
logical table even though the payload allocation is physically shared. This
conservative per-value accounting keeps each value's ceiling independent and
does not pretend to provide a process-global byte counter.

Retention admission is optional. If the ceiling would be exceeded, if checked
arithmetic overflows, if an entry/table `try_reserve` fails, or if projection
cannot prove package ownership, the route silently falls back to the ordinary
MCE result and leaves the value's memo empty or partial. These cache-only
failures must never become `Error::Limit`, `Error::Allocation`, or a changed
semantic refusal. As with the existing `Arc::new` patterns in this crate,
global allocator failure while constructing an admitted `Arc` is not a
recoverable typed error; the design's fallible guarantee covers all explicit
reservations and admission arithmetic, not language-level OOM recovery.

## Invariants and refusal preservation

The implementation and tests must keep these facts true:

* raw input size is checked before lookup and cached output size is checked
  before use;
* content-type, root-name, producer-name, notes-root proof, first-name-error,
  relationship, duplicate-identity, slide-count, and notes inventory checks
  run in their current order;
* an MCE error is returned exactly as on a cold capture and is never cached;
* a name or notes-proof outcome is recomputed from the reused output and is
  never the memoized value;
* `SlideRootProof` continues to borrow the current raw source. Only the
  processed-XML projection is dropped before fallback name allocation; the
  proof remains alive until notes consumes it, as in the current
  implementation;
* output bytes and all snapshot semantic values are byte/value identical to an
  uncached capture, including `Cow` ownership where that ownership is visible
  to the existing private route;
* no public XML, global map, process-wide state, lock, executor, or cross-format
  cache is introduced; and
* release, a cache miss, an admission refusal, and an allocation fallback do
  not alter the snapshot revision, patch bytes, error type, or published
  package bytes.

The debug/test path should re-run `process_ooxml` for every memo hit and
`debug_assert_eq!` the fresh output against the retained `Arc` bytes. The
release path uses the retained bytes directly. This mirrors the existing
`PartDigests` and cross-slide retention proof style.

## Gate list before retaining a production candidate

The coder should not retain the candidate until all of the following are
available in the 0704 packet:

1. A source census and focused oracle prove the real 0703 capture, one-edit,
   two-edit, and no-op semantic results match byte-for-byte with memo enabled,
   disabled, released, and over-ceiling. The primary measurement must show
   changed-commit slide MCE calls avoided; the catalog seam is not a substitute.
2. Exact default-profile output SHA-256, length, capacity, `Cow` ownership, and
   root/name/notes-proof comparisons cover the existing marked,
   marker-free, malformed, over-input, over-output, late-name, and notes
   refusal fixtures. Cache hits must rerun the proof scans, and no refusal may
   be served from a stored error. Add strict or other profile cases only when
   an existing PPTX fixture or the implementation's actual call path makes
   them relevant.
3. A foreign `Part` whose `blob_arc()` copies or mismatches its visible
   `blob()` proves that no entry is retained and that the cold result and
   error order are unchanged. A part replacement and an allocator-address ABA
   adversary prove that the strong raw `Arc` closes the key-reuse hole.
4. Snapshot clone, `release_retained_mce`, transaction clone, changed-slide
   invalidation, unchanged-slide carry, candidate `rebound_to`, equal-content
   fresh-allocation rebind, and package-facade mutation-release tests prove
   the memo never outlives the package allocation it describes.
5. Exact-limit, one-under, one-over, zero/opt-out, metadata-only, capacity
   greater-than-length, checked-overflow, and fallible-reserve cases prove the
   aggregate ceiling and silent fallback. The observed accessor charge must
   equal the documented logical charge, with no claim of exact RSS.
6. `Limits` intersections in same-package patch merge and cross-package
   planning are tested; the tighter policy must win without changing any
   patch or refusal result.
7. A fresh release benchmark over the representative real and generated PPTX
   sources, including the existing marker-free, marked, malformed, and refusal
   cases that exercise this route, measures capture, changed commit, apply,
   allocations, peak resident memory, and output identity. The decision must
   separate slide MCE savings from presentation/notes work and report the
   retention memory cost and fallback rate. Unrelated-format benchmarks are
   unnecessary unless the implementation crosses a format boundary.
8. Existing format gates, focused opened/notes/MCE tests, all-feature tests,
   source-boundary and non-iWork evidence checks, strict report classification,
   cleanup, and final artifact hashes pass after the candidate is chosen.

Until these gates pass, this record is an implementable design, not a claim
that 0704 should be retained. The list is completion evidence for the coder
and reviewer, not an additional user-approval barrier.
