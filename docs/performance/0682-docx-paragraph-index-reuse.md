# 0682 — DOCX paragraph-index reuse across document views

Status: retained and implemented. `performance_claim: none` — the measurements
below are scoped evidence, not registered claims or coverage promotions.

The repeated-view profile in [0680](0680-docx-paragraph-memo-design.md)
showed that every fresh `Package::document()` paragraph query rebuilt the same
bounded range index. Physical OPC payload caching already avoided repeated
decompression. This change retains one successful semantic index in the DOCX
package so another document view can reuse those ranges.

## Ownership and validity

The cache is private to `litchi-docx`; it does not add a physical-package cache
or change an ordinary CRUD signature. Eager packages key a memo by a weak
reference to the exact raw main-part allocation and the semantic route. A hit
requires upgrading that reference and comparing allocation identity. Resolving
the current main relationship still precedes this lookup. A replaced part,
retargeted relationship, or changed serialization therefore cannot reuse ranges
from a different allocation; restoring an old allocation during save rollback
is safe. The weak key does not itself retain raw document XML.

Source-backed packages key their owner-local memo by the captured source
identity/revision and semantic route. Managed documents retain `PartData` and
the existing conservative index admission. The cache never detaches their
payload into an uncharged `Arc<Vec<u8>>`. Eager and unmanaged source-backed
views retain lazy first-query construction and the existing MCE visibility
pass. Managed misses perform the bounded validation and eager index scan under
the established admissions; successful hits share that completed result.

One mutex serializes first construction. Only complete successful indexes are
retained; failures and unsupported index construction do not become permanent
shared negative entries. Per-view query behavior and fallback parsing remain
unchanged. Source and execution checks still fence managed publication and
each ordinary source-backed document call.

## Retention and budget scope

The logical cache ceiling is 8 MiB for one range array plus 128 modeled bytes
of entry overhead. It is not a measured allocator ceiling or a global bound
on every live document. Old views may pin an evicted memo. The managed index
reservation follows the shared memo until its last owner drops, including when
the package has already been dropped.

The conservative managed admission remains `24 * (xml_len / 4 + 1) + 1024`
memory bytes and `xml_len / 4 + 4` objects. Those reservations include a build
envelope and must not be confused with the smaller logical retained weight.
Clean-cache reclamation releases only an unpinned memo; active views continue
to hold their reservations.

Other package operations reclaim a clean index before their own admissions.
The payload-only pressure retry is restricted to the OPC path whose memory
and object admission precedes work and I/O. Source-XML operations are not
retried wholesale: their later admission can follow cumulative work charges.
Cancellation or stale-source failure drops the tentative document before
attempting clean-cache reclamation.
Delayed tail-append plans retain a reference to their parent DOCX package so
publication can reclaim a clean memo warmed after preparation. Forward,
inverse and source-fingerprint entry points use the same cleanup rule.

## Evidence and limitations

Public regressions in
[`paragraph_index_reuse.rs`](../../crates/litchi-docx/tests/paragraph_index_reuse.rs)
exercise repeated counts and text, BOM/MCE visibility, mutable serialization,
raw main-part replacement, root relationship retargeting, failed-save rollback
and retry, source changes, cancellation, and managed pressure reclamation.
Internal tests distinguish semantic cache hits/builds from OPC payload hits.

The retained probe, raw measurements, reproduction commands and integration
logs live in [results/change-0682](results/change-0682/README.md). The baseline
is revision `35686023471e066d54a26737316e0d648f787f91`; candidate source hashes
bind the measurements and checks to the implemented files.

Package opening and its new base cache allocation are outside the measured
loops. The first index allocation now contains the shared memo as well. The fresh-view loops
include view construction and queries; they are not whole-package latency
measurements or zero-allocation claims.

Eight CPU-12-pinned ABBA repetitions per case compare five fresh views/counts
on one already-open package, including the first index build. The table gives
A1→B1 p50 and the two paired timing changes (negative means lower latency).
The largest absolute baseline A/A p50 movement for these fresh-view cases
was 2.19%; A1/A2 drift within the final ABBA window reached 4.18%. All 114
matrix cases have matching result digests and dimensions;
source payload cold-load/hit counters are unchanged.

| Route | Corpus | p50 milliseconds, before → after | Timing change, B1/A1 and B2/A2 | Allocated bytes, before → after |
| --- | --- | ---: | ---: | ---: |
| Eager | 200 generated paragraphs | 0.345 → 0.135 | −61.01% / −61.01% | 44,225 → 19,085 |
| Eager | 10,000 generated paragraphs | 16.404 → 5.998 | −63.44% / −63.33% | 1,726,465 → 355,533 |
| Source-backed | 200 generated paragraphs | 0.375 → 0.153 | −59.19% / −58.45% | 135,915 → 110,775 |
| Source-backed | 10,000 generated paragraphs | 16.519 → 6.146 | −62.80% / −62.62% | 2,333,781 → 962,849 |
| Managed | 200 generated paragraphs | 0.840 → 0.186 | −77.86% / −78.06% | 1,782,905 → 430,565 |
| Managed | 10,000 generated paragraphs | 39.433 → 8.108 | −79.44% / −79.75% | 83,311,771 → 17,148,839 |
| Eager | `ComplexNumberedLists.docx` | 0.357 → 0.296 | −17.27% / −20.24% | 287,650 → 266,798 |
| Source-backed | `ComplexNumberedLists.docx` | 0.366 → 0.303 | −17.32% / −15.82% | 374,437 → 353,585 |

The managed real-fixture route retains its existing MCE refusal and is not a
successful timing row. There is no tail-latency claim from eight samples.

The initial 25-query managed same-view control took only 250–280 ns and had an
inconsistent 12% movement in one leg. Longer 100,000-query windows exposed a
2.95–6.26% regression in the two archived candidates. Their disassembly shows
an extra dependent load through `Arc<Memo> → Arc<Index>`. The final memo owns
the index directly, keeping sharing and reservations on `Arc<Memo>` while
removing the inner allocation and load. The final longer control changes
−1.32% to −0.04%, with zero allocation calls in either version; no control
speedup is claimed. Its baseline A/A floor is 0.48%. Eager/source first-index
allocation counts now match the baseline, with eight extra requested bytes.

After five managed views have been dropped, charged memory rises from 10,113
to 71,833 bytes for 200 paragraphs and from 500,113 to 3,501,833 bytes for
10,000 paragraphs. The differences are exactly the retained conservative
index reservations, 61,720 and 3,001,720 bytes. These are budget gauges, not
allocator-live-byte or RSS measurements. The logical memo weights are only
1,728 and 80,128 bytes; pressure reclamation and last-owner release are tested.

Final validation passes formatting, all-feature/all-target compilation,
Clippy with warnings denied, 1,523 DOCX tests (31 existing ignored examples),
104 facade tests with DOCX/ODT features, and rustdoc with warnings denied.
Seven evidence/dependency checks also pass, including 50 claim-gate tests.
The retained logs include the initial qualification/import failures and the
successful final rerun.

This record covers paragraph queries on repeated document views. It does not
establish full CRUD coverage, physical cold-cache behavior, remote-source
performance, RSS savings, or program-level completion. The broader
[goal audit](GOAL_AUDIT.md) remains open. iWork is outside this batch.
