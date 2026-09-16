# Log paragraphs for change 0646

Four paragraphs, one each for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and
`ADR_COMPLIANCE.md`, in the style of their newest sections. The coordinator
merges these; this change does not edit those files.

## For `HOTSPOTS.md`

**Queue item 8 is measured, and the cross-package copy's remaining hotspot is
not where callgrind says it is.** 0598 left the second candidate build in place
and priced its deflate at 5.96 G `Ir` per lifecycle — 22% of a media-rich
lifecycle's 26.8 G — with SHA-256 at 72%. Natively the ranking inverts: a
scratch implementation that retains the planned archive and reuses it removes
exactly one of the two serializations, worth **11.43% of lifecycle instructions**
but **72.10% of the apply phase's wall clock** (339.4/337.9 ms before,
94.4/94.5 ms after, pair spreads 0.46% and 0.12%) and **23.10% of whole-child
native cycles** against a 0.77% A/A floor. The gap is SHA-NI: valgrind masks the
CPUID bit, so every instruction share this area's records attribute to hashing is
about five times its native cycle share, and deflate of the copied closure is
correspondingly under-ranked. **Any further ranking of this path should be read
off `perf stat`, not off the `Ir` tables.** Everything 0598 left in place stays
in place — the eight `package_fingerprint` calls, the four
`physical_package_fingerprint` calls and the eight captures per lifecycle are
identical in both legs — so what the second serialization was hiding is now
visible: **planning is the largest remaining native term** at a p50 of 311.5 ms
against apply's 94.4 ms, and it still builds, deflates and reopens a whole
candidate. The next item in this
area is still 0587's PPTX-1(c) — per-part digests with a redefined
complete-package revision — which needs a magic bump for `LPRM0001` and
`LPCP0002` and has needed a frozen design record since change 0590.

## For `GOAL_AUDIT.md`

Change 0646 is a case where the size gate is met by a wide margin and the change
is still not made, and the reason is worth stating plainly: **the saving is
large, the mechanism is sound, and the blocker is that ADR 0005's budget does not
exist in this path.** ADR 0005 says "Every operation charges a hierarchical
resource budget supplied by an execution context." `litchi-core` provides
`Budget`, `Resource::Memory` and three finite profiles; the PPTX
opened-presentation path uses none of them. Its `opened::Limits` bounds the size
of an operation's *output* (`max_patch_bytes`), not live memory — with one
exception, `max_history_bytes`, which `History::push` does enforce on retained
undo patches with a typed `Error::Limit` and eviction. That exception is the
precedent and the shape to copy, and it is also the measure of how far the crate
is from ADR 0005: one enforced live-memory ceiling, in one place, reached by no
budget.

The second lesson is about *where* an optimization's cost lands. Retention costs
nothing at planning — the allocator probe shows the plan's live bytes rise by the
archive plus a 40-byte `Arc` header and by nothing else, and fall back exactly
when the plan is dropped — so it is not "extra work", it is a decision not to
free. But it moves a bounded transient inside one call into unbounded state on a
public value the caller holds. **ADR 0005's "cache behavior is semantically
invisible" clause is the wrong licence for that**, and the record says so:
change 0598 used that clause correctly for a 32-byte memo on an immutable
snapshot; a ceiling-bounded 128 MiB hold on a caller-owned plan is a retention
policy, not a cache, and a retention policy belongs under the budget. The gate
this yields is not a number. It is: *an optimization that changes how long memory
lives must be opt-in until something is charging for it.*

## For `REPORT.md`

**0646 — the cross-package copy's second deflate can be removed by retaining the
planned candidate archive, and the price is one whole serialized package held in
the plan.** Design, retained; **no production code changed**;
`performance_claim: none`. Scratch implementation measured and kept as a patch in
the packet. Counts (callgrind isolation pairs, CPU 21): `bounded_package_bytes`
2 → 1 call per lifecycle, `PackageWriter::write_to_stream` 7 → 6,
`PreservationIndex::write_to` 2 → 1, `zlib_rs::deflate::deflate` 1,112 → 556
calls and 5.956 G → 2.978 G `Ir` (**−50.00%**), with every hash, capture, reopen
and the whole replan unchanged; lifecycle `Ir` 26.80 G → 23.74 G (−11.43%) on
media-rich and −2.10% on plain. `sha2::sha256::compress256` is unchanged to five
significant figures, which is the design working: the archive digest is
recomputed over the retained bytes rather than carried, so the hash is a wash and
what is removed is the serialization and its deflate. Paired ABBA timing, 30
samples per leg: `pptx_cross_copy_media_rich` 659.1 → 412.5 ms (**−37.42%**, A/A
floor −0.15%, B/B +0.22%), its apply phase **339.4 → 94.4 ms (−72.10%)** with
pair spreads of 0.46% and 0.12% and its plan phase flat at −0.64%;
`pptx_cross_copy_media_rich_lifecycle` −35.61% against a −3.44% A/A floor; the
plain selectors −2.99% and −3.46% against floors of −1.34% and +0.31%. Native
`perf stat` whole child: cycles **−23.10%** (A/A −0.77%, B/B −0.18%). One phase
is bimodal at ±4× and was chased rather than reported as a result:
`publication_ns` on the media-rich lifecycle reads, per leg, 1.536 / 1.529 /
1.527 / **6.545** ms — the outlier is a *before* leg, and in an earlier window on
a superseded revision it was an *after* leg — while the same phase on the
non-lifecycle selector is 6.51 / 6.30 / 6.15 / 6.50 ms in the same processes.
Publication executes no changed code. Memory: the plan's live bytes rise by exactly the archive plus 40 bytes
(31,583 on plain, 33,599,911 on media-rich) and return to parity when the plan is
dropped; lifecycle allocated bytes −19.98%, region peak +12.18% on media-rich and
**−5.45% on plain**, because on a small candidate the removed writer's growth
buffer outweighs what is retained. Verdicts unchanged: 871 `litchi-pptx` tests
pass on the untouched base and all 871 pass on the retained path (874 with the
three binding-proof tests the patch adds), with a `debug_assert` re-serializing
the candidate on every reuse. Fourteen admission gates are defined; **G2 blocks
implementation**, because ADR 0005's hierarchical budget does not exist in this
path and the implementable alternative — a sixth member of `opened::Limits` — is
a breaking public API change.

## For `ADR_COMPLIANCE.md`

**ADR 0005 (bounded memory; scratch storage as an explicit capability).** The
record states the memory question exactly rather than asserting compliance: a
plan retaining the candidate archive owns up to `max_patch_bytes` more, unshared,
for its whole lifetime — 256 MiB per plan at the default limits, the entire
`Resource::Memory` limit of `Profile::Server` — and **nothing in the PPTX
opened-presentation path charges a budget**. The "cache behavior is semantically
invisible" clause is explicitly declined as the licence, and the design is opt-in
with `Rebuild` as the default, a ceiling that belongs in `opened::Limits` beside
the already-enforced `max_history_bytes`, and a typed
`Error::Limit { resource: "cross-slide retained candidate archive bytes", limit }`
for the strict route. One shape gap is named rather than papered over:
`litchi_pptx::Error::Limit` carries `resource` and `limit` but no observed value,
which ADR 0005 requires of a limit error. **The scratch clause is honoured by
never spilling**: a candidate archive is document content, so above the ceiling
the design rebuilds in memory. The record also records why a caller-supplied
scratch provider would need a *different* fallback rule from the one ADR 0005
states — absence must mean rebuild, not a typed resource error, because litchi is
never obliged to hold the candidate — and leaves that route unspecified.
**ADR 0003 and change 0454.** No revision value, proof format or durable encoding
changes: `LPCP0002` and its six 32-byte revisions are untouched, `apply_patch`
retains nothing, and the plan stays an immutable value whose equality does not
become a byte comparison. **ADR 0011 and the 2026-08-21 OPC exact-source
amendment.** The retained archive is the exact source the candidate reopen
already authorized. The one place the design brushes ADR 0011 is the copy at
application: `OpcPackage::from_vec*` takes an owned `Vec`, so the retained `Arc`
is cloned once, and removing that needs a shared-bytes ingress in `litchi-opc` —
an ownership decision the record declines to make. **One proof is replaced by an
argument, and it is named.** A fresh `to_stream` of the in-memory candidate is no
longer compared against the retained bytes; its warrant is the four staleness
proofs plus determinism, re-derived by a `debug_assert` in debug and test builds
and backed in release by the recomputed graph and archive revisions. A release
test proves a substituted archive is still refused with the destination left
byte-identical.
