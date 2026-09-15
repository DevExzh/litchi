# Log sections for change 0600

Four paragraphs for the coordinator to merge, one per log, in the style of each
log's newest section.

## For `HOTSPOTS.md`

## 0600 — one fewer observation per cold Part read, a bounded monitored-read scope, and an allocation-free member lookup

Items ZIP-3 and ZIP-4 of the [0587 queue](0587-remaining-opportunity-survey.md)
are implemented. ZIP-4 is the larger half: `lookup_member_name` allocated and
re-normalized a `String` on every `contains`, `metadata`, `read` and session
admission, at least twice per structural member, and a name that is already
canonical — no `:`, no `\`, no empty, `.` or `..` segment, no trailing `/` — is
now decided by one byte pass and looked up as a borrowed `&str`.
`normalize_str_fallibly` falls 71%, the `rfind(':')` 23%, and open-plus-one-part
instructions fall **3.45%, 2.58% and 2.56%** on a 132-member XLSX, a 48-member
PPTX and a 10-member DOCX, with allocations per open down 6.32%, 5.93% and 6.59%.
ZIP-3's per-site analysis reached only one of the five observations a cold
`part().data()` takes — the pre-publication one, a strict subset of the fence
that follows the provisional publication and can roll it back — so the count is
5 → 4, not the 5 → 2 the survey hoped for, and the record says which four sites
survive and why. The sticky `monitor_reads` flag is the item with the real
reach: one `stream_to` used to convert every later positional read on that
package into two `version()` calls for the package's lifetime, and eight cold
Part reads after a stream cost **102 observations against 40 on an identical
package that had not streamed**; a counted RAII scope over the fifteen sites
that set the flag brings both to 32. ZIP-2 (the span model) and ZIP-5 (the
accessor 0577 is blocked on) remain the next items on this path.
[Change and limitations](0600-opc-cold-read-observations-and-name-lookup.md);
[evidence](results/change-0600/README.md).

## For `GOAL_AUDIT.md`

## 0600 — one fewer observation per cold Part read, a bounded monitored-read scope, and an allocation-free member lookup

`docs/GOAL.md` puts unnecessary work ahead of I/O and unnecessary allocation
ahead of layout and algorithms, and this change takes only those two steps: not
one positional request, not one requested byte and not one output byte moves on
any measured phase, on any of three real fixtures, in either read order. The
"exact no-ops stay exact" rule is what bounds the design — the observation
removed is the only one of five that a strict-subset argument reaches, and the
relocation onto the failed-publication branch keeps change 0317's precedence
rather than trading it for a count. The `monitor_reads` bound was allowed only
after checking that change 0327's publication protocol never reads the flag: its
fences are direct `ensure_current()` calls, so bounding the flag cannot weaken
them, and what the flag does guarantee — a per-chunk fence *during* an
incremental copy — is now scoped to exactly the operation that needs it. The
audit's standing caveat that instructions rank work rather than latency applies
to the 3.45% figure and is stated in the record; `performance_claim: none` and
no claim-registry entry. The pre-existing non-determinism found on the way — the
relationship iteration order of `Relationships::iter()` varies per process
because it walks a `HashMap` — is reported, not fixed.
[Change and limitations](0600-opc-cold-read-observations-and-name-lookup.md);
[evidence](results/change-0600/README.md).

## For `REPORT.md`

## 0600 — one fewer observation per cold Part read, a bounded monitored-read scope, and an allocation-free member lookup

Three mechanisms. `lookup_member_name` now returns a `Cow` and borrows the
caller's `&str` whenever the name is already its own lookup key, which it is for
every member name a real OOXML package carries; `walk_relationship_graph` builds
two copies of each relationship target name instead of three; and
`SourceSnapshot::monitor_publication` returns a `#[must_use]` counted RAII scope
instead of latching an `AtomicBool` that nothing ever cleared. Measured on three
real fixtures: open-plus-one-part instructions −3.45%/−2.58%/−2.56%, allocations
per open −6.32%/−5.93%/−6.59%, observations per cold Part read 5 → 4, and the
cold reads that follow a `stream_to` 102 → 32, 88 → 32 and 78 → 24 — exactly the
counts an identical package that never streamed pays. Correctness is a
byte-identical corpus digest over **179 packages and 3,450 lines** covering every
Part's `data()`, its `stream_to()`, a second `data()` taken after that stream,
every relationship edge and three refusal probes per package; an exhaustive
lookup-equivalence test over every string of length ≤ 4 in `{a, / , \, ., :}`;
and a corpus lookup test over **4,215 members × 8 spellings**. Two of the new
invariant tests were confirmed to fail against the pre-change behaviour. Gates
are clean on both crates with 1,271 tests passing, and the four dependent format
crates pass unchanged. No speedup claim follows; `performance_claim: none`.
[Change and limitations](0600-opc-cold-read-observations-and-name-lookup.md);
[evidence](results/change-0600/README.md).

## For `ADR_COMPLIANCE.md`

## Change 0600 compliance update

Change 0600 removes one source observation from the cold OPC Part load, bounds
the lifetime of the `monitor_reads` flag, and gives `lookup_member_name` a
borrow-only path for canonical names. ADR 0005 is satisfied on its own terms: no
`ReadLimits`, `SourceCacheLimits` or ZIP limit value, check or ordering moved,
and read-order independence is preserved because the lookup is a pure function
of the name and the catalog verdicts are proven identical member by member over
179 packages in both physical read orders. Change 0317's error precedence —
source-version failure, then execution failure, then the mapped ZIP error — is
preserved by relocating the removed observation onto the
`publish_pending_with_observer` failure branch, after the flight has been
completed as failed so no flight leaks and no waiter hangs. Change 0327 is
untouched: its provisional-publication protocol fences with direct
`ensure_current()` calls and never consults `monitor_reads`, and the fence that
decides a publication — the one taken outside the cache locks, which can roll
the publication back — is the one that remains. The monitored scope counts
rather than latches, so a concurrent operation, a nested scope and a scope on a
clone of the same snapshot cannot end one another's monitoring; while any scope
is open the fencing is exactly what the latch provided. ADR 0006 is untouched:
no output byte, no typed refusal and no preservation path changed, and the only
refusal that disappears is an `ErrorKind::Allocation` for an allocation that no
longer happens. No new `unsafe`, no weakened defence, no hidden global pool, no
ambient I/O, and no archive, lock or executor type reaches a public signature —
`MonitoredReads` and `MonitoredReadDepth` are module-private and
`LookupMemberName` was already private. See
[Change 0600](0600-opc-cold-read-observations-and-name-lookup.md);
`performance_claim: none`.
