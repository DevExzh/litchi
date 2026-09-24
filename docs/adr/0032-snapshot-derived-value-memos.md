# ADR 0032: Snapshot memos of small derived values of payload bytes

- Status: Accepted (2026-09-24, by the owner's decision recorded in
  [change 0758](../performance/0758-owner-decisions-2026-09-24.md); proposed
  2026-09-23)
- Date: 2026-09-23
- Supersedes: nothing. Amends: the 2026-09-16 amendment of
  [ADR 0005](0005-io-memory-and-performance.md) titled "retained per-part digest
  memos on an opened-presentation snapshot and its facade", by widening the one
  noun it admits. Every other sentence of ADR 0005, and every other accepted
  record, is untouched.
- Raised by: [change 0743](../performance/0743-pptx-semantic-text-and-edit-path.md),
  whose commit `99ce9c5e34` implemented such a memo, measured it, and was
  reverted by `07207d2aef` because no accepted record admits it.

## Context

Every opened-presentation capture validates the notes graph, and the notes
graph requires each slide's root conformance. That classification is
established by the complete notes-graph scan of the slide (`scan_processed_xml`
under the Transitional and then the Strict profile): a whole-document validating
pass whose only result the capture keeps is `Option<Conformance>`. On the
harness's 100-slide × 100-text-box deck the scan is most of a capture, and the
capture is a fixed cost of every opened-presentation transaction.

`Transaction::commit` then captures the staged package again, from scratch. The
staged package shares the payload allocation of every slide the transaction did
not rewrite, so for a one-slide edit the commit rescans 99 slides whose bytes
are, by allocation identity, the bytes the source snapshot's capture already
classified. ADR 0003 asks a commit to validate the changed dependency closure;
repeating the classification of unchanged bytes is not part of that obligation.

The 2026-09-16 amendment of ADR 0005 already lets an opened-presentation
snapshot keep a memo over bytes its package holds, under a list of conditions,
for exactly one kind of value: "a bounded table of **digests** over bytes the owner
already holds". A notes-root classification is not a digest. Change 0743's
commit `99ce9c5e34` kept one anyway, as a `SlideRootMemo` on the snapshot, and
did not meet one condition: when the table could not be reserved it became
empty instead of raising the typed resource error the amendment requires. The
program's charter forbids implementing behaviour an accepted record does not
admit, so the commit was reverted and this record asks for the admission
explicitly.

What the memo was worth, measured by change 0743 with the harness's own public
calls on the 100 × 100 deck (a probe built at every commit, median of three
processes, core 8):

| after | one-edit commit | one-edit edit/save total | medium one-edit total |
| --- | ---: | ---: | ---: |
| the borrowed notes scan (`8039963c33`) | 21.45 ms | 46.44 ms | 1.550 ms |
| plus the memo (`99ce9c5e34`) | 2.13 ms | 27.11 ms | 1.346 ms |

The commit fell from 21.4 to 2.1 ms (−90%) and the whole one-edit edit/save
from 46.4 to 27.1 ms (−42%), with the published archive, the patch and the
revision byte-identical, every refusal identical over the repository's PPTX
corpus, and no `Part::blob` or `blob_arc` observation added. Without the memo
the one-edit commit remains a second complete capture.

## Decision

### 1. What may be memoized

The amendment's noun widens from *digests* to **small derived values of
payload bytes**: a value `v = f(b)` where

- `f` is a deterministic, total function of the payload bytes `b` alone and of
  compile-time constants (a processing profile, a root name, fixed ceilings). It
  may not read the part's name, content type or relationships, another part, a
  caller-chosen limit, a clock or any other ambient input, unless that input is
  part of the key;
- `v` has a fixed, small size, stated in the implementing record and bounded by
  a constant (a digest, a classification, a verdict, a count); a memo whose
  value grows with the payload is a cache of parsed values and stays governed by
  ADR 0005's eviction rules;
- a failure of `f` is never memoized: only values `f` returned are stored, so a
  refusal is always recomputed from the bytes.

A digest is one such value, so every memo the amendment admits today stays
admitted unchanged.

### 2. Under exactly the amendment's conditions

Every condition of the 2026-09-16 amendment applies unchanged: the memo is keyed
on the identity of an allocation the owner holds and retains a strong reference
to it, so a hit proves byte identity; every entry names an allocation the owning
snapshot's own package holds, which the implementation proves through that
snapshot's own part-digest memo; a memo carried onto another package is
projected onto that package's allocations, never inherited; a facade that
retains one adopts the memo of each snapshot it publishes and releases it at
every mutation that produces none; a miss is an ordinary recomputation, so no
value, refusal, limit or published byte depends on which entries it holds; it
is rebuilt or projected, never mutated in place; it holds no limit and raises no
error the unmemoized path would not raise (the reservation error of section 3 is
of the same kind as the operation's other reservation errors); its resident cost
is bounded by the part count and stated in the implementing record; and every
reused value is re-derived by a `debug_assert!` in test and debug builds.

### 3. Reservation failure is a typed resource error

Building the memo charges its table through fallible reservation, and
exhaustion there is the typed `Error::Allocation` the operation reports for its
other reservations — never an abort and never a silent empty table. Projecting
an existing memo onto a rebound package follows the part-digest memo it sits
beside: that projection already degrades to an empty memo when it cannot
reserve, because a rebind is infallible and an empty memo is only a miss. The
two memos therefore never disagree about what a rebind carries. If the reviewer
prefers the projection to be fallible as well, `Snapshot::rebound_to` and its
callers must become fallible for both memos together; this record does not
propose that.

### 4. What re-applies

With this record accepted, change 0743's commit `99ce9c5e34` re-applies
unchanged except for the rule in section 3: `SlideRootMemo::from_records`
returns `Result<Self>` and its reservation failure becomes
`Error::Allocation { resource: "opened-presentation slide-root memo", .. }`,
propagated by the capture that builds it; `SlideRootMemo::project` keeps the
empty-memo degradation the part-digest projection has. Its resident cost is one
32-byte entry per captured slide (allocation key, `Arc`, classification) plus a
table header, at most 4,096 entries under the default `Limits::max_parts`
(128 KiB); on the 100-slide deck 3,200 bytes of entries and a 40-byte header.

## Consequences

- A one-slide commit on a large deck stops rescanning every untouched slide; the
  commit becomes proportional to the rewritten slides plus the capture's
  per-slide relationship, root, name and identity checks.
- A capture gains one more reservation that can fail with `Error::Allocation`.
  No other refusal, limit or published byte moves.
- The snapshot holds one small table more. It pins no payload its package does
  not already own, and a clone shares it.
- Future memos of this shape (a per-payload verdict, a per-payload count) need
  no new record, only an implementing record that proves the conditions.
- The widening does not admit memos of parsed trees, processed XML or any value
  whose size grows with the payload; those remain caches under ADR 0005, or
  retention under its second 2026-09-16 amendment.

## Verification

Acceptance should require, at minimum:

1. The ten focused tests `99ce9c5e34` carried: a one-slide commit rescans
   exactly one slide; memo-assisted and cold captures agree on value and on
   refusal (CDATA, unbound prefix, first, middle and last slide); equal bytes in
   a new allocation miss; a foreign part whose `Arc` does not alias its payload
   is never memoized or pinned; a lookup adds no payload observation over a
   plain recapture; a published snapshot keeps entries only for its own
   allocations; a notes-owning deck reuses proofs and keeps its notes graph; and
   memo-assisted and plain recaptures agree over every repository PPTX fixture.
2. A test that a refused reservation while building the memo returns
   `Error::Allocation` and leaves no partial snapshot.
3. The existing 0704 MCE-retention tests, including the observation-count test,
   passing with the memo in place.
4. A re-measurement of the harness's `pptx_semantic_one_edit_save` and
   `pptx_semantic_opened_transaction_phases` against the head without the memo,
   both legs built by one command.

Nothing in this record may be implemented before a human accepts it, and an
implementation must be measured before any of it is claimed.
