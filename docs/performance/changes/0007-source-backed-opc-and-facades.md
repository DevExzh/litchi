# Source-backed OPC and additive format facades

Status: implemented selective-open stage; performance claims limited

`SourceBackedPackage` opens an immutable positional OPC source with source
versions, finite `SourceCacheLimits`, weighted LRU eviction, per-entry
single-flight loading, and content-free cache diagnostics. DOCX, XLSX, and
PPTX expose additive source-backed facades; ordinary package/CRUD APIs retain
their existing semantic shape and do not expose archive handles or physical
identifiers.

Cache bytes are bounded by `SourceCacheLimits`, but are **not yet charged to the
hierarchical `Budget`**. That integration gap remains explicit. Raw ZIP
preservation is implemented and tested at the soapberry layer, but OPC
integration and performance measurement are still pending.

The EOCD terminal-probe ABBA evidence shows structural-open source reads reduced
by **73.6% to 98.5%** and ordinary-payload overlap reduced to zero. It does not
support a latency claim: later `EntryId` and cache-diagnostics changes confound
the comparison and some cells exceed the 5% variance threshold. See
[`EOCD before A`](../results/abba-eocd-before-a.json),
[`after A`](../results/abba-eocd-after-a.json),
[`before B`](../results/abba-eocd-before-b.json),
[`after B`](../results/abba-eocd-after-b.json), and the
[`source-versus-eager record`](../results/stage3-source-vs-eager-many-small.json).

The committed positional XLSX source record is
[`xlsx-source-positional.json`](../results/xlsx-source-positional.json) (CPU2,
warm-up 5, n=30). Open p50 is 33.881 us (tiny), 56.493 us (medium), and
139.897 us (dense); list-after-open is 90 ns, 125 ns, and 981 ns respectively
with zero timed source reads. First-cell p50 is 76.685 us, 1.095 ms, and
67.854 ms; narrow-column p50 is 75.925 us, 1.079 ms, and 68.558 ms. The last
two operations overlap exactly the selected worksheet member physically;
unselected worksheet read-call counts are zero. These are physical-overlap
counts, not claims about logical materialization or complete CRUD cost.
