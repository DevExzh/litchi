# Remaining measurement and lifecycle work

The full non-iWork goal remains open. Correcting an allocator observer is a
prerequisite for trustworthy measurements, not completion of the resource or
CRUD program.

1. Add a separate operation-region high-water tracker. Initialize it from
   live bytes at region entry and update it with post-allocation values;
   never reset the global peak. Initialization, cross-thread updates and
   finish need an explicit synchronization contract. A single-worker scope
   must not silently claim arbitrary concurrent accuracy. Exercise a prior
   global peak larger than the region peak, realloc growth/shrink, freeing
   pre-existing allocations and overflow/unavailable states. Label `after -
   before` as signed net live change, not operation-owned retention.
2. Add explicit retained-object/drop boundaries and near-output-limit media
   cases. Account for raw archives, the first plan's patch payloads, application
   candidate payloads and metadata. Global high-water and process RSS remain
   different quantities from an operation-local peak or bounded aggregate
   memory guarantee.
3. Establish a matched source-backed media-rich PPTX lifecycle. The existing
   `presentation/source_cross_copy.rs` admits the generated slide's eight
   relationship-free 2 MiB image leaves. The current source-backed harness is
   plain-only and measures planning plus publication with opening excluded.
   A distinct lifecycle selector must include source/destination opening and
   use the owned corpus's exact input archives and slide positions.

For item 3, the source-backed oracle must verify nine added OPC parts and ten
ZIP members: the slide, its relationships member and eight images. Check
collision remapping, content types, image bytes, leaf relationships, copied
slide semantics and untouched destination raw records. New media ZIP framing
need not match the owned writer because target names are remapped; keep a
separate deterministic physical output digest for each implementation. Do not
substitute the current plain phase selector for a lifecycle comparison.

Retain typed refusals for unknown non-Part members, encrypted entries,
signatures/macros/protection, unsupported or external image relationships,
non-leaf media, layout/dialect mismatch, stale revisions, malformed sizes,
cancellation and resource limits. The admission conclusion comes from source
review and existing multiple-image/collision tests, not a new workload result.

Broader source-backed OPC/XLSX read/edit lifecycles, physical cold/range access,
native-producer breadth, managed-cache contention and explicit scaling remain
required. Existing source-backed capability is not broad performance evidence.
