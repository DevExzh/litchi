# 0680: bounded DOCX paragraph-index memo design

Status: design and evidence only. performance_claim: none. This record does
not change production code, the public API, limits, refusal order, or package
bytes. It closes the measurement and ownership prerequisite left by [0670](0670-docx-parser-residues.md)
and gives the next implementation a bounded cache shape. OLE2 and OOXML remain
the active priority; iWork is excluded.

The relevant accepted records were read before this design: the ADR index and
rules in [docs/adr/README.md](../adr/README.md), ADR 0001 (strict public
layers), ADR 0002 (crate ownership), ADR 0003 (immutable snapshots), ADR 0005
(weighted clean-value caches and hierarchical budgets), ADR 0006 (preservation
and refusal order), ADR 0008 (evidence gates), ADR 0010/0011 (archive and OPC
ownership), ADR 0024 (current topology), and ADR 0030/0031 (lazy part decode
and execution budgets). No proposed ADR is used as authority.

## Current implementation and measured residue

ParagraphIndex is an Arc<[ParagraphRange]>; each range is two u32 values, so
its retained range storage is eight bytes per paragraph. The index stores
offsets and does not own XML
([document_part.rs:35-73](../../crates/litchi-docx/src/parts/document_part.rs#L35)).
Its scanner is bounded at 1,000,000 ranges and uses fallible vector growth.
Each eager Package::document() currently calls DocumentPart::from_part, which
runs the visible-XML pass and creates a fresh OnceLock
([model.rs:565-575](../../crates/litchi-docx/src/package/model.rs#L565)).
The first paragraph query fills that cell; a new document view therefore gets a
new index cell
([document_part.rs:577-603](../../crates/litchi-docx/src/parts/document_part.rs#L577)).

The source-backed package has the same per-view cell. Its physical OPC cache
returns the main payload from one cold load and later hits, but
source_backed::Package::document() still creates a new document and repeats the
index path
([source_backed.rs:629-684](../../crates/litchi-docx/src/source_backed.rs#L629)).
A managed source view currently reserves DocumentIndexAdmission, parser
workspace, objects, depth, and cumulative work before its eager scan. The
admission price is conservative:

~~~
ranges = xml_len / 4 + 1
memory = ranges * 24 + 1024
objects = ranges + 3
~~~

For the probe's 10,000-paragraph XML (xml_len = 500,113), this is 125,029
range slots, 3,001,720 bytes, and 125,032 objects. For its 200-paragraph XML
(xml_len = 10,113), it is 2,529 slots, 61,720 bytes, and 2,532 objects.
Those are admission ceilings, not measurements of the retained Arc or a claim
about peak process memory. The current implementation is visible at
[source_backed.rs:2863-2885](../../crates/litchi-docx/src/source_backed.rs#L2863).

The retained source payload has a different owner. SourceBackedPackage and its
PartCache own physical compressed-payload reads and weighted byte eviction;
DOCX semantic state does not belong in that OPC cache. A managed PartData
handle stays attached to its physical reservation. The new memo must therefore
be owned by each DOCX Package, with an index-only value and its own typed
admission. It must not detach a managed payload or add a reverse
litchi-opc -> litchi-docx dependency.

The probe in [results/change-0680/probe/](results/change-0680/probe/) creates
the package before the measured loop, snapshots the counting allocator
immediately before and after that loop, asserts paragraph_count() == expected
inside every counted iteration, and prints source-cache diagnostics only after
the sampled delta. Its paired callgrind totals subtract the reps = 0 process
baseline from reps = 1 and reps = 5; the raw generated profiles were used to
make the readable summary in
[results/change-0680/measurements/callgrind-summary.txt](results/change-0680/measurements/callgrind-summary.txt)
and then removed; the small allocator/callgrind diagnostic .log files remain
alongside the probe source.

The isolated index work is stable across the two routes and scales with the
generated document:

| fixture | route | fresh document() | fresh document() + paragraph_count() | one view, first count | one view, later counts |
| ---: | --- | ---: | ---: | ---: | ---: |
| 200 paragraphs | eager | 266,385 Ir; 2,558 B | 1,173,009 Ir; 8,845 B | 906,807 Ir; 6,287 B | 0 new B; callgrind noise |
| 200 paragraphs | source-backed | 409,608 Ir; 94,248 B | 1,316,290 Ir; 100,535 B | 906,754 Ir; 6,287 B | 0 new B; callgrind noise |
| 10,000 paragraphs | eager | 12,026,610 Ir; 2,558 B | 56,894,106 Ir; 345,293 B | 44,867,354 Ir; 342,735 B | 0 new B; callgrind noise |
| 10,000 paragraphs | source-backed | 15,196,785 Ir; 609,874 B | 60,080,202 Ir; 952,609 B | 44,883,540 Ir; 342,735 B | 0 new B; callgrind noise |

Ir is callgrind's retired-instruction total after the paired subtraction; B
is the counting allocator's allocated-byte delta. The source-backed
10,000-paragraph document() includes its one cold physical payload load. Its
five-view run reports one cold load and four cache hits; the semantic index
allocation is still 342,735 bytes per fresh counted view. The 200-paragraph
source run reports the same one-cold/four-hit pattern and 6,287 bytes per
fresh counted view. These figures are operation-shaped evidence for repeated
work, not latency, throughput, RSS, or speedup results.

## Bounded implementation to take next

Add a private DocumentIndexCache in litchi-docx and put one cache handle on
the owning eager Package and one on the owning source-backed Package. The
cache is a single-current-generation weighted clean-value cache. It has no
process-global state and no public cache type. Its entry is conceptually:

~~~
DocumentIndexMemo {
    generation: GenerationKey,
    value: Option<Arc<ParagraphIndex>>,
    index_admission: Option<Arc<DocumentIndexAdmission>>,
    weight_bytes: u64,
}
~~~

The ParagraphIndex remains XML-free. The view continues to own the visible XML
or managed PartData; the memo keeps only ranges and the admission that prices
those ranges. The generation key is package-owned and cannot be recycled: an
eager package increments it at every successful replacement of the
main-document blob, while a source-backed package keys it to the immutable
main-part/source-version identity. The semantic route (managed raw bytes or
unmanaged visible bytes) is part of the key. This prevents an old range array
from being applied to a new XML generation without hashing or copying the
document.

The cache's initial policy should be explicit and private: one current entry,
weighted by checked range_count * size_of::<ParagraphRange>() plus a fixed
memo/Arc envelope, with a clean-retention ceiling derived from the existing
1,000,000-range bound (an initial eight-MiB ceiling is the rounded policy
value). The exact fixed envelope must be measured with the implementation and
recorded before the code gate. The ceiling applies to clean values retained by
the cache map. A view that has pinned a value is reported separately as active
weight and remains charged by its managed admission; the cache cannot pretend
that an active value disappeared merely because its map entry was trimmed. A
value above the clean ceiling is still returned to the active view when its
ordinary parser/admission succeeds, but it bypasses cache publication and is
recomputed by a later view. The ceiling therefore controls retention only; it
cannot turn a valid query into a limit refusal. This eight-MiB value is a
logical clean-index limit, not the managed reservation limit. Diagnostics and
admission must separately carry the full conservative `DocumentIndexAdmission`
charge (based on XML length), plus pinned old values and build workspace; the
hierarchical budget remains the limit for those live reservations.

Cache entries are clean when only the cache owns their Arc. A view that has
received the memo pins it, so eviction removes only the cache's reference and
the active view continues to hold the index and its admission. Replacing a
generation drops the old clean entry; an active old view keeps its reservation
until its last owner goes away. This is the weighted clean-value rule required
by ADR 0005. A bare package-level OnceLock<Option<Arc<ParagraphIndex>>> would
be a permanent owner and is not sufficient. A OnceLock may remain inside one
cache build/flight if the surrounding entry is evictable.

The cache needs an explicit pressure seam; generation replacement alone must
not be the only way to trim a clean value. Implement
DocumentIndexCache::trim_clean(required_weight) as a private owner-facing
operation. It removes clean least-recently-used entries until the requested
weight fits the clean ceiling, increments an eviction counter, and reports
whether each removed entry was clean or still pinned by an active view. The
managed admission helper calls this seam before its first Memory/Objects
reservation and retries once after a ResourceLimit. The owner also calls it
when a new generation is installed and from the existing DOCX semantic
admission/pressure path. An unmanaged owner may call the same seam when it
explicitly releases clean semantic state; there is no ambient process-wide
pressure callback. If every candidate is pinned, trim stops and the new
managed reservation sees the still-live old charge. A private diagnostic
snapshot records hits, misses, builds, clean evictions, pinned-entry drops,
oversized bypasses, admission failures, clean weight, and active pinned
weight, so tests can distinguish map retention from caller-owned values.

The build path is deliberately split by current ownership:

1. On a cache hit, check cancellation/execution and source freshness, clone
   the memo, and do not admit parser workspace or consume parser work. The
   physical source cache remains responsible for its own hit and reservation.
2. On a miss, establish the generation key and perform the current visibility
   and source checks. Trim clean entries before admitting the build. Managed
   source builds reserve the existing DocumentIndexAdmission before
   ParagraphIndex::from_xml, then retain that typed admission in the memo. The
   existing parser workspace, depth, and cumulative work charges remain around
   the actual scan. Unmanaged builds keep their current no-admission behavior.
   If an active old memo remains pinned while a new generation builds, its
   reservation stays in the same hierarchical budget; the new admission must
   fit alongside it or return the existing typed resource error before parser
   work starts. Dropping an old clean map entry never releases an active
   handle's reservation.
3. Publish only a complete Some(Arc<ParagraphIndex>) for the still-current
   generation. A cancelled, stale, allocation-failed, or over-bound build
   publishes no partial value. Preserve the current per-view fallback cell for
   a swallowed None; do not retain a negative cache value until a separate
   measurement proves that it is useful and harmless.
4. Recheck cancellation/source version before returning the view. An edit or
   source change cannot make a memo hit win over the existing freshness fence.

The eager package needs one invalidation seam for all main-document blob
replacements. Existing direct set_blob sites in package document, story,
data-store, and raw-edit paths should route through that private seam before
the memo is enabled. A no-op replacement keeps the same generation only when
the existing package byte-identity contract keeps the same blob allocation;
otherwise it creates a new generation. Source-backed packages have no mutable
main-document publication in this view; their source-version check is the
invalidation fence.

## Gates for the production change

The next code change is ready only when the focused suite proves all of these
properties:

* two fresh views on one unchanged package share one successful memo, and a
  concurrent first query has one published value;
* paragraph_count, paragraphs, paragraph, and source paragraph_text return the
  same values as the current streaming fallback, including an assertion
  against the expected paragraph count;
* replacing the main XML cannot reuse old ranges, and old active snapshots
  remain byte and semantically stable;
* a cache hit does not detach a managed PartData, and a managed admission is
  charged once for the retained memo, released after cache eviction and the
  last active view, with no residual memory or object charge;
* a pinned entry can be removed from the clean cache without freeing the
  active value, while an oversized value bypasses retention and later
  recomputes with the same result;
* explicit trim/pressure tests observe clean eviction, pinned weight that
  remains charged during a second-generation build, and the semantic cache
  counters separately from the OPC payload-cache counters;
* malformed XML, range overflow, cancellation, source change, and budget
  refusal retain the current error identity and no partial memo becomes
  visible; and
* source diagnostics still show one physical cold load followed by hits for
  repeated views, while semantic hit/miss counters distinguish the new DOCX
  cache from the OPC payload cache.

The implementation should first move the existing conservative
DocumentIndexAdmission ownership into this cache. Repricing its 24 bytes/range
envelope or changing parser-work accounting is a separate budget change: the
probe's 342,735-byte isolated index path and the 3,001,720-byte managed
admission for the 10,000-paragraph fixture are not the same quantity, and
neither justifies a refusal-boundary change by itself.

No claim is registered here. The retained source, commands, paired totals, and
source audit are in [results/change-0680/README.md](results/change-0680/README.md).
