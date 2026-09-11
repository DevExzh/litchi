You are the principal Rust performance engineer responsible for an end-to-end
performance program for Litchi, a lossless Office document read/write library.

Repository:
https://github.com/DevExzh/litchi

Target branch:
feat/office-format-completeness

The document docs/CRUD_Scenario_Checklist.md is the required scenario taxonomy.

Your job is not to produce a speculative optimization report. Explore the
repository, establish reproducible performance evidence, implement the highest-
impact optimizations, validate them, and leave the repository with benchmarks,
tests, measurements, and documentation that make the results independently
reproducible.

========================================================================
MISSION
========================================================================

Improve the comprehensive performance of common Office CRUD workflows as close
as practical to the hardware limits, with an ambitious program-level goal of an
order-of-magnitude improvement in the most important currently bottlenecked
end-to-end scenarios.

Optimize all relevant dimensions together:

- wall-clock latency and tail latency;
- single-thread throughput and instructions per byte/object;
- cache locality and branch-prediction friendliness;
- allocation count, allocation size, peak live bytes, and peak RSS;
- memory copies, memory-bandwidth pressure, and working-set size;
- instruction-level parallelism;
- safe SIMD where it is demonstrably useful;
- scalable parallelism under Amdahl's law;
- lock contention and synchronization overhead;
- disk access patterns, positional and sequential reads, and syscall count;
- caller-supplied remote/range sources without adding ambient networking;
- sequential non-seek output;
- zero-copy or single-copy ownership where physically possible;
- compressed-entry or stream passthrough for unchanged content;
- cold-cache, warm-cache, and concurrent workloads.

Treat “10x” as a program goal for the largest end-to-end bottlenecks, not as a
license to claim that every operation can or did improve by 10x. Every claim
must be scoped to a named scenario, corpus, machine, build, and metric.

Use this optimization order:

1. Eliminate unnecessary work.
2. Eliminate unnecessary I/O, decompression, parsing, allocation, copying,
   validation, and recompression.
3. Improve data layout and cache behavior.
4. Improve algorithms and incremental indexes.
5. Introduce bounded explicit parallelism.
6. Apply SIMD or lower-level instruction tuning only to measured hot loops.

Do not begin with broad SIMD rewrites or lock-free data structures.

========================================================================
NON-NEGOTIABLE ADR AND SEMANTIC CONSTRAINTS
========================================================================

Before editing code, read docs/adr/README.md and every accepted ADR. Pay
particular attention to:

- ADR 0001: priorities and API layers;
- ADR 0002: crate topology and dependency direction;
- ADR 0003: immutable snapshots, edits, commits, patches, and conflicts;
- ADR 0005: I/O, memory, caching, execution, and performance evidence;
- ADR 0006: preservation, validation, security, and compatibility;
- ADR 0008: migration and verification state;
- ADR 0010 and ADR 0011: archive and physical-package ownership;
- ADR 0024: current workspace topology.

Accepted ADRs are hard constraints. Do not silently violate or reinterpret
them. If a desirable optimization would conflict with an accepted ADR, do not
implement the conflicting behavior. Record the conflict and, only when truly
necessary, draft a separate proposed ADR for human review. Continue with all
compatible work.

The following rules are mandatory:

1. Correctness, lossless preservation, bounded resource use, and safety take
   precedence over speed.

2. Preserve is the default behavior. Untouched package members, streams,
   unknown markup, unknown binary fields, namespace choices, lexical forms,
   relationships, ordering, compression metadata, timestamps, extensions,
   macros, media, and other unsupported content must survive according to the
   existing preservation contract.

3. Never trade a typed refusal for a partial or guessed edit. Dependency-closure
   edits must remain atomic and fail closed.

4. Keep immutable cheap-to-share Send + Sync snapshots, selector-first public
   APIs, isolated Edit objects, atomic Commit publication, source-checked
   reversible patches, deterministic conflicts, and no last-writer-wins
   behavior.

5. Exact semantic no-ops must remain exact no-ops and should share the original
   source allocations and bytes wherever the current contract requires this.

6. Validation must not mutate. Normal save must not silently repair or
   normalize. Determinism and security boundaries must remain intact.

7. Do not expose archive implementation types, physical IDs, raw locks,
   Arc<RwLock<_>>, executors, runtime handles, or unsafe storage through ordinary
   public CRUD APIs.

8. Do not add ambient filesystem, network, clock, process, random, Tokio, or
   Rayon behavior where the ADRs require explicit providers or execution
   contexts.

9. Do not use a hidden global Rayon pool. CPU parallelism must be opt-in and
   controlled by an explicit execution context with thread, memory, I/O,
   cancellation, and task-granularity budgets.

10. Do not weaken existing forbid(unsafe_code) or deny(unsafe_code) policies.
    Prefer safe Rust. Any proposed unsafe optimization must be isolated in an
    ADR-permitted low-level owner, have a scalar safe fallback, contain explicit
    safety invariants, and show material measured benefit. Do not add it merely
    because it appears faster.

11. Preserve crate ownership and dependency-direction checks. A performance
    shortcut must not reintroduce an archive implementation dependency into a
    facade or format crate that does not own that grammar.

12. Preserve source limits, hierarchical budgets, cancellation behavior, typed
    errors, encryption and signature boundaries, and malicious-input defenses.
    A faster implementation must not weaken zip-bomb, XML, graph, allocation,
    or integer-overflow protections.

13. Do not make unsupported performance claims in comments, documentation,
    commit messages, or reports.

========================================================================
REPOSITORY EXPLORATION
========================================================================

Map the complete data path before optimizing it:

source/path/bytes/range provider
    -> format detection
    -> ZIP/OPC, CFB, ODF, or iWork container index
    -> content types, relationships, directory, or component graph
    -> mandatory catalog
    -> lazy semantic part/stream parsing
    -> queries and traversal
    -> Edit planning
    -> dependency-closure validation
    -> Commit and Patch construction
    -> publication validation
    -> save/repackage/atomic replacement

Inspect at least:

- litchi-core source, budget, patch, and selector infrastructure;
- soapberry-zip archive indexing, ReaderAt, decompression, caching, and writer;
- litchi-opc physical package, PackageReader, OpcPackage, Part, and writer paths;
- litchi-cfb read, sector, FAT, MiniFAT, directory, stream, and writer paths;
- litchi-ooxml-common and shared DrawingML owners;
- the DOCX, XLSX, PPTX, DOC, XLS, PPT, and XLSB vertical paths;
- shared detection and prepared-source paths;
- existing fuzz targets, fixtures, native Office evidence, and CI checks;
- current allocation helpers, caches, Arc ownership, arenas, SmallVec usage,
  hash maps, ordered maps, trait objects, parser scratch buffers, and serializers.

For every important path, document:

- which bytes are read;
- which bytes are decompressed;
- which bytes are copied;
- which data is retained;
- which values allocate;
- which structures are rebuilt;
- which validation is repeated;
- which locks are acquired;
- which work is parallel;
- which work is serial;
- which work is proportional to the whole document rather than the touched
  dependency closure.

Do not assume a comment saying “lazy”, “zero-copy”, “parallel”, or “SIMD” proves
an end-to-end property.

========================================================================
PHASE 1: BUILD A REPRODUCIBLE PERFORMANCE BASELINE
========================================================================

Before making invasive changes, add a benchmark and profiling facility that is
separate from production runtime dependencies and respects crate boundaries.

Use Criterion, Divan, a purpose-built process benchmark harness, or an
equivalent stable solution. Use external tools where available, including:

- perf stat and perf record on Linux;
- cargo-flamegraph, samply, or equivalent sampling profilers;
- DHAT, heaptrack, allocator instrumentation, or equivalent allocation tools;
- strace or equivalent syscall tracing;
- platform-appropriate counters on macOS and Windows.

Provide graceful fallbacks when a tool is unavailable.

Record at least:

- p50, p95, and p99 elapsed time where meaningful;
- throughput in source bytes, logical bytes, cells, rows, paragraphs, slides,
  parts, or operations per second;
- cycles, instructions, IPC, branches, and branch misses;
- L1 data misses, LLC misses, and page faults where available;
- allocations, allocated bytes, peak live bytes, and peak RSS;
- bytes read from the source;
- range-read count and request-size distribution;
- bytes decompressed and recompressed;
- bytes copied in memory where practical;
- output bytes and write-call size distribution;
- lock wait time or a reliable contention proxy;
- CPU utilization and scaling efficiency at 1, 2, 4, 8, and available-core
  thread counts;
- serial fraction and an Amdahl-law model for parallel scenarios.

Benchmark both cold and warm states. Isolate process-level benchmarks when
allocator state, filesystem cache, or RSS matters. Record CPU model, core
count, memory, storage, operating system, Rust toolchain, compiler flags,
allocator, and relevant environment settings.

Use sufficient warm-up and samples. Report uncertainty or confidence intervals.
Do not present a single noisy run as evidence.

Create deterministic corpus manifests and hashes. Include representative:

- tiny, small, medium, large, and very large documents;
- sparse and dense spreadsheets;
- many-small-part and few-large-part packages;
- text-heavy and media-heavy documents;
- highly compressible and incompressible payloads;
- documents with many relationships and deep dependency closures;
- unknown vendor extensions and unsupported content;
- encrypted, signed, protected, macro-enabled, and external-link-bearing files;
- malformed and adversarial files that remain within configured limits;
- files from multiple real-world producers where licensing permits;
- synthetic files generated deterministically for scaling studies.

Do not commit confidential documents. Prefer reproducible generators and
content-addressed public fixtures.

========================================================================
CRUD SCENARIO MATRIX
========================================================================

Use CRUD_Scenario_Checklist.md as the required coverage map. Include benchmarks
and correctness checks for all relevant categories:

1. Content reading and extraction.
2. Structural queries and document analysis.
3. Conversion and export.
4. Creation from scratch.
5. Template filling and targeted replacement.
6. Append-only and incremental generation.
7. Structural editing.
8. Content deletion and document sanitization.
9. Cross-document copying and assembly.
10. Merging and splitting.
11. Comparison, patching, and three-way merge.
12. Validation, repair, and normalization.
13. Dynamic content calculation and refresh.
14. Security, protection, encryption, and signing.
15. Low-level package, Part, Relationship, stream, and extension operations.

At minimum, implement end-to-end benchmarks for:

- open and identify format;
- open and list worksheets/slides/sections without reading their payloads;
- query one property or one named object;
- read one cell, paragraph, slide, image, or low-level Part;
- scan all stored cells, paragraphs, or slides;
- full text extraction;
- semantic conversion to a sequential sink;
- create a small document;
- create or append a very large row/paragraph/slide stream;
- exact no-op edit and commit;
- edit one cell, title, paragraph, property, or relationship and save;
- update approximately 1% of logical objects and save;
- bulk update all matching objects;
- clear versus remove versus hide versus detach versus garbage collection;
- sanitization and irreversible redaction;
- copy one worksheet, slide, section, chart, or media-bearing object with its
  complete dependency closure;
- merge and split;
- produce, serialize, apply, invert, and merge patches;
- validate without mutation;
- repair through an explicit plan;
- preserve an unknown extension while editing understood content;
- copy or replace one low-level Part while retaining every untouched member.

Distinguish the four append semantics described by the checklist:

- streaming creation from scratch;
- logical append to an existing structure;
- adding a new package Part;
- arbitrary modification followed by repackaging.

Do not combine them into one misleading “append” number.

Run important scenarios against:

- owned bytes;
- borrowed bytes;
- filesystem-backed positional input;
- an instrumented caller-supplied ReadAt source;
- a simulated high-latency range source with configurable latency, bandwidth,
  request overhead, and maximum range size;
- sequential non-seek output;
- filesystem atomic save.

Do not add an ambient HTTP client to a core or format crate. Remote behavior
must be exercised through an explicit caller-provided source/provider.

========================================================================
INITIAL HYPOTHESES TO VERIFY, NOT ASSUME
========================================================================

Profile these likely bottlenecks and either confirm or disprove them:

1. Path and generic-reader OPC input may be fully slurped into Vec<u8> before
   parsing instead of retaining a positional source.

2. Ordinary OpcPackage opening may decompress and retain every admitted Part,
   even for selective structural queries.

3. LazyArchiveReader::read may clone cached decompressed buffers; its single
   RwLock-protected string-keyed map may cause contention, duplicate miss work,
   excessive key allocations, and unbounded retention.

4. Bulk ZIP decompression may use the hidden global Rayon pool rather than an
   ADR-compliant execution context.

5. Stored ZIP entries may be copied even when a borrowed source range is stable.

6. Saving may sort, validate, serialize, and recompress every Part rather than
   raw-copying unchanged compressed entries and regenerating only dirty closure
   members.

7. OpcPackage may pay for duplicated indexes, Box<dyn Part> indirection,
   source-XML maps, repeated PackURI/String allocation, and linear
   case-insensitive fallback scans.

8. XLSX selective sheet loading may be undermined by eager OPC materialization.

9. XLSX small edits may clone package metadata, scan broad graph regions,
   recreate catalogs, and reopen more data than the changed dependency closure
   requires.

10. Per-cell Arc<str>, Box<str>, enums, Options, and AoS layout may dominate
    allocation count, cache footprint, or memory bandwidth in dense and large
    sparse sheets.

11. CFB may use shared Read + Seek, allocate and zero a sector Vec for frequent
    reads, materialize complete MiniFAT/ministream buffers, and copy streams
    unnecessarily.

12. Detection may repeat container indexing unless an opaque prepared source is
    reused.

These are hypotheses. Do not change an architecture solely because it appears
suboptimal in source code. Require profiles and scenario measurements.

========================================================================
OPTIMIZATION WORKSTREAM A: SOURCE AND I/O
========================================================================

Converge container readers on the ADR-defined immutable positional-source
model.

- Reuse or adapt litchi_core::ReadAt rather than maintaining incompatible
  positional abstractions when ownership permits.
- Retain the source and its identity/version in snapshots.
- For filesystem input, use positional reads instead of shared seeking.
- Consider mmap only in an ADR-permitted low-level owner and only after proving
  it improves representative workloads without weakening portability, limits,
  or safety.
- For remote/range sources, read the ZIP tail and central directory through
  bounded coalesced ranges, then fetch member ranges on demand.
- Add bounded read-ahead and range coalescing based on measured access patterns.
- Avoid request amplification from many tiny remote reads.
- Reuse bounded scratch buffers through an explicit operation/execution context.
- Use sequential reads when traversal is naturally sequential and positional
  reads when independent ranges are needed.
- Measure syscall count and read-size distribution rather than assuming larger
  buffers are always better.
- Consider read_vectored/write_vectored when profiles show a useful scatter/
  gather pattern.
- Preserve non-seek sequential sinks and typed partial-output failures.
- Keep atomic filesystem replacement and durability behavior unchanged.

For stored ZIP entries, provide borrowed or shared range-backed access when the
source lifetime permits it.

For Deflate entries, do not call decompression “zero-copy.” Instead, minimize
the path to one bounded materialization, stream decompressed bytes directly into
a consuming parser/sink where possible, and avoid retaining them when the
scenario does not require retention.

========================================================================
OPTIMIZATION WORKSTREAM B: CONTAINER LAZINESS AND COPY-ON-WRITE
========================================================================

Redesign internal package state, without leaking physical implementation types,
so that opening a document does only mandatory work.

A successful open should generally retain:

- the stable source;
- the validated ZIP/CFB directory or equivalent physical index;
- mandatory content-type, relationship, root, workbook, presentation, or
  document catalogs;
- lightweight descriptors for unloaded Parts or streams;
- explicit limits, source version, and cache policy.

Do not decompress and retain all semantic payloads merely to construct a
snapshot.

Implement or improve:

- per-Part/stream lazy materialization;
- a weighted cache with explicit byte/object budgets;
- clean-entry eviction;
- pinning while borrowed views or edits require an entry;
- dirty entries that cannot be evicted until committed or discarded;
- per-entry single-flight initialization so concurrent misses do not duplicate
  decompression;
- sharding or immutable indexes when contention measurements justify them;
- compact internal IDs or interned names while preserving public semantic names;
- a canonical case-insensitive lookup index where required by the format,
  rather than a repeated full scan;
- stable source ranges and compressed-entry descriptors;
- explicit cache diagnostics useful to benchmarks.

For save and commit:

- copy untouched compressed ZIP payloads without decompression/recompression
  when the preservation contract and ZIP framing permit it;
- preserve required local and central records, flags, methods, CRCs, sizes,
  timestamps, order, extra fields, comments, descriptors, and ZIP64 details;
- reserialize and recompress only dirty members;
- copy large unchanged binary media directly from source to output;
- regenerate mandatory manifests only when their semantic state changed;
- keep deterministic output rules and exact-source patch authorization;
- support sequential, non-seek output.

Treat direct compressed-entry passthrough as a correctness-sensitive container
feature, not a byte concatenation shortcut. Cover data descriptors, ZIP64,
stored versus Deflate entries, duplicate/ambiguous names, malformed metadata,
and sink failures.

For CFB, investigate sector- or stream-level copy-on-write and raw copying of
unchanged physical regions. Implement it only where directory, FAT, MiniFAT,
allocation, and preservation invariants can be proven.

========================================================================
OPTIMIZATION WORKSTREAM C: MEMORY LAYOUT AND ALLOCATIONS
========================================================================

Use allocation profiles and cache counters to optimize high-frequency objects.

Candidates include:

- splitting hot and cold fields;
- compact immutable arrays instead of pointer-rich graphs;
- internal tagged enums instead of trait objects in measured hot loops;
- small inline collections for genuinely small distributions;
- arenas or region allocation tied to a parse/snapshot lifetime;
- string/URI interning scoped to one snapshot;
- source-owner plus byte-span representations;
- shared backing buffers with compact offsets;
- compact lexical-number and text stores;
- structure-of-arrays or hybrid layouts for large homogeneous cell/record sets;
- row/block indexes over a compact cell store;
- persistent/COW collections for snapshots and edit plans;
- reusing buffers through explicit bounded contexts;
- eliminating temporary Strings, Vecs, maps, and formatted error messages from
  success hot paths;
- avoiding repeated Arc control blocks for millions of tiny values;
- fallible exact or well-estimated reservations.

Preserve exact lexical forms and unknown bytes. Do not replace lossless values
with normalized numeric or text representations merely to save memory.

Do not introduce a faster non-DoS-resistant hash strategy for untrusted keys
without proving that existing bounds and collision behavior maintain the
security contract.

Measure type sizes, cache-line occupancy, allocations per logical object, and
random versus sequential access before and after layout changes.

========================================================================
OPTIMIZATION WORKSTREAM D: PARSING, INDEXING, AND SERIALIZATION
========================================================================

Optimize parsers by reducing passes and ownership conversions:

- fuse validation, indexing, extent calculation, and semantic extraction when
  doing so preserves error ordering and proof obligations;
- retain byte spans into stable owners for borrowed views;
- parse only accessed Parts and requested semantic domains;
- cache validated indexes by source version;
- avoid reparsing unchanged XML or binary streams during independent reads;
- compare byte names and attributes before UTF-8/String conversion;
- use memchr/memmem and compact token dispatch where profiles show benefit;
- keep common branches predictable and move uncommon validation/error work out
  of the hot success path without skipping it;
- reduce redundant namespace, relationship, and reference resolution;
- build indexes in source order when that improves traversal locality;
- avoid intermediate DOMs when a streaming/indexed representation is enough.

Optimize writers by:

- writing directly to caller sinks;
- reserving bounded output buffers based on measured size models;
- using itoa/ryu or direct lexical retention rather than temporary formatting;
- batching small writes without creating archive-sized buffers;
- preserving unchanged byte spans;
- generating only affected XML/binary regions where safe;
- avoiding repeated sort/materialize cycles by maintaining deterministic
  ordered plans.

Do not cache error results in a way that prevents recovery after cancellation,
budget changes, or a distinct source version.

========================================================================
OPTIMIZATION WORKSTREAM E: EDITS, PATCHES, AND DEPENDENCY CLOSURES
========================================================================

Make tiny updates proportional to their actual semantic and physical closure.

Investigate:

- precomputed incoming/outgoing relationship indexes;
- format-owned dependency-closure indexes;
- effect/read/write-set indexes reused by join, commit, patch, and validation;
- incremental invalidation of derived catalogs;
- persistent maps, compact sorted vectors, or Arc-backed slices for edit state;
- moving rather than cloning edit payloads;
- O(1) or near-O(1) exact no-op detection;
- changed-Part plans that avoid cloning the entire package graph;
- incremental verification whose proof is equivalent to the required full
  invariant;
- lazy reopening of a candidate while still satisfying ADR publication rules;
- source-version and exact-byte guards shared rather than copied;
- bounded lazy inverse material for reversible patches;
- explicit history weight based on retained bytes, not merely operation count.

Do not skip a final reopen or semantic readback when an accepted ADR requires
it. Instead, make that reopen cheap through shared source state, mandatory-only
opening, lazy payloads, and incremental indexes.

For cross-document copying, compute the complete dependency closure once.
Deduplicate media, styles, themes, fonts, and other resources only when exact or
semantically permitted equivalence is proven. Keep collision resolution
explicit.

========================================================================
OPTIMIZATION WORKSTREAM F: PARALLELISM, ILP, AND CONTENTION
========================================================================

Add or complete an explicit execution context that can control:

- maximum worker count;
- CPU task budget;
- memory and decompression-in-flight budget;
- I/O concurrency;
- cancellation;
- task size thresholds;
- optional caller-provided executor or scoped worker facility.

Do not expose a runtime through ordinary document APIs.

Construct a dependency DAG for work that is genuinely independent, such as:

- separate ZIP entries or CFB streams;
- independent worksheets, slides, stories, or sections;
- disjoint edit subplans;
- independent validation domains;
- compression of changed output members.

Use bounded pipelines where I/O, decompression, parsing, semantic processing,
and serialization can overlap. Apply backpressure so parallel decompression
cannot exceed the memory budget.

Avoid parallelizing tiny Parts, short rows, or small documents. Derive and
benchmark thresholds from size and cost estimates.

Measure:

- serial fraction;
- task scheduling overhead;
- lock wait;
- cache coherence traffic;
- bandwidth saturation;
- useful CPU utilization;
- scaling efficiency.

Use Amdahl's law explicitly. A graph showing eight busy threads is not evidence
of speedup.

Improve instruction-level parallelism inside serial hot loops where profiles
justify it, for example through independent accumulators, batched validation,
software pipelining, or reduced dependency chains.

Use lock-free structures only after proving locks are a material bottleneck and
that the lock-free design improves end-to-end results. Prefer immutable state,
per-entry OnceLock/single-flight state, sharding, and short critical sections
when they are simpler and faster.

Avoid saturating memory bandwidth with redundant parallel copies or
decompression. Tune working-set size and task granularity before increasing
thread count.

========================================================================
OPTIMIZATION WORKSTREAM G: SIMD
========================================================================

Apply SIMD only to loops that remain significant after higher-level work
elimination.

Potential candidates, subject to measurement, include:

- XML delimiter/name/attribute scanning;
- UTF-8 and UTF-16 validation or conversion;
- escaping and unescaping;
- CRC and checksums;
- byte classification and delimiter search;
- BIFF and protobuf field scanning;
- varint or fixed-width record classification;
- large numeric lexical scans;
- bulk copy, compare, or equality checks where the standard library is not
  already optimal.

Requirements:

- runtime feature detection performed outside the inner loop;
- scalar fallback;
- x86_64 and aarch64 coverage where practical;
- exact differential tests against the scalar implementation;
- malformed-input tests;
- minimum-size thresholds;
- isolated, audited ownership;
- measured end-to-end benefit, not only a synthetic microbenchmark.

Do not hand-vectorize code already compiled to equivalent or faster machine
code. Inspect generated assembly or compiler reports when necessary.

========================================================================
LEGACY CFB-SPECIFIC WORK
========================================================================

For litchi-cfb and DOC/XLS/PPT paths, investigate:

- adopting immutable ReadAt instead of shared Read + Seek;
- positional concurrent reads without a cursor lock;
- reusable bounded sector buffers;
- avoiding a new zero-filled Vec for every sector read;
- reading contiguous sector runs directly into the final destination;
- separating mandatory FAT/directory metadata from lazily loaded streams;
- compact sector-chain indexes;
- loading only referenced MiniFAT/ministream ranges where feasible;
- file-backed or source-range stream views;
- bounded spooling for huge streams;
- unchanged-stream or unchanged-sector copy-through on save;
- parallel independent stream processing under an explicit execution context.

Preserve all CFB validation, ownership, cycle, overlap, FAT, MiniFAT, directory,
and truncation checks.

========================================================================
DECISION AND REGRESSION RULES
========================================================================

For every optimization:

1. Capture before measurements.
2. State the hypothesis and expected mechanism.
3. Implement the smallest coherent change.
4. Run correctness, preservation, adversarial, and performance tests.
5. Capture after measurements using the same setup.
6. Keep the change only if the result is statistically and practically useful,
   or if it is a necessary measured enabler with a clearly documented follow-up.
7. Revert speculative complexity that does not improve representative workloads.

As an initial gate, flag any scenario with approximately more than a 5%
latency/throughput regression, 5% peak-RSS regression, or material loss of
scaling. Treat this as a review trigger rather than hiding the regression in a
geometric mean. Tighten thresholds after measuring normal benchmark noise.

Prefer improvements that help multiple CRUD scenarios and shared substrates.
Do not optimize one microbenchmark by damaging common workloads.

Track geometric means only across explicitly named, sensibly normalized
scenario groups. Also report every individual result.

========================================================================
CORRECTNESS AND SAFETY VERIFICATION
========================================================================

After each coherent workstream, run all applicable existing gates, including:

- cargo fmt;
- cargo check for the workspace and feature combinations;
- warning-denied Clippy/lint policy;
- rustdoc with warnings denied;
- crate-boundary and dependency-direction checks;
- unit, integration, doc, compile-pass, and compile-fail tests;
- existing fuzz targets;
- relevant native Office fixture checks.

Add tests for:

- byte and semantic preservation of untouched content;
- unknown extension pass-through;
- exact no-op byte/source sharing;
- reversible patch application and inverse restoration;
- source mismatch conflicts;
- deterministic output;
- resource limits and allocation failures;
- cancellation;
- concurrent access and joins;
- cache eviction and pinning;
- sink failure after partial output;
- remote/range short reads and source-version changes;
- ZIP64, data descriptors, stored and Deflate entries;
- malformed ZIP, XML, CFB, BIFF, and relationship graphs;
- signed, encrypted, protected, and macro-enabled files;
- cross-platform behavior.

Use Miri, sanitizers, loom-style concurrency testing, or equivalent tools where
they are applicable and supported, especially for new low-level ownership or
concurrency code.

========================================================================
DELIVERABLES
========================================================================

Produce and maintain:

1. A benchmark harness and deterministic corpus/generator manifest.

2. docs/performance/BASELINE.md containing:
   - hardware and software environment;
   - corpus descriptions and hashes;
   - commands;
   - baseline metrics;
   - flamegraph/profile summaries;
   - allocation and I/O summaries;
   - Amdahl and scaling analyses.

3. docs/performance/HOTSPOTS.md containing:
   - end-to-end path maps;
   - confirmed bottlenecks;
   - disproven hypotheses;
   - ranked opportunities by expected total CRUD impact, risk, and ADR
     compatibility.

4. An ADR-compliance matrix for every architectural optimization.

5. Production code changes with focused tests.

6. Machine-readable before/after results, preferably JSON or CSV, including
   benchmark version, git revision, corpus hash, environment, and uncertainty.

7. A concise per-change performance record explaining:
   - what work was removed or accelerated;
   - which scenarios improved;
   - which scenarios regressed;
   - memory impact;
   - concurrency impact;
   - remaining limitations.

8. A lightweight, stable performance smoke check suitable for CI, plus a
   fuller manually triggered or scheduled benchmark workflow. Do not make noisy
   cloud-hosted microbenchmarks a hard merge gate until variance is understood.

9. A final performance report containing:
   - before/after tables;
   - geometric means and individual results;
   - peak RSS and allocation changes;
   - copied/decompressed/recompressed byte changes;
   - cold/warm and local/range-source results;
   - scaling curves;
   - remaining serial fractions;
   - remaining top bottlenecks;
   - precise, non-marketing claims.

========================================================================
DEFINITION OF DONE
========================================================================

The work is complete only when:

- the important CRUD scenarios have reproducible baselines;
- the largest measured bottlenecks have been addressed in Amdahl/ROI order;
- selective reads perform work proportional to mandatory metadata plus accessed
  content rather than total uncompressed document size where the format permits;
- targeted updates avoid parsing and recompressing unrelated content where the
  preservation and container contracts permit;
- unchanged large media and package members can flow from source to output
  without unnecessary decompression or logical-byte copies;
- streaming creation and append scenarios have bounded memory proportional to
  an explicit window rather than total output size;
- parallel paths use explicit bounded execution and demonstrate real scaling;
- hot data layouts and caches are justified by allocation and hardware-counter
  evidence;
- SIMD is limited to proven hot loops with fallbacks and differential tests;
- all accepted ADR, correctness, preservation, security, determinism, and
  boundary requirements still pass;
- no performance claim lacks reproducible evidence.

Start by reading and profiling. Do not ask the user to select a hotspot before
you have measured the repository. Preserve unrelated existing work, do not
discard local changes, and do not rewrite large subsystems merely to make the
code look more “performance-oriented.”