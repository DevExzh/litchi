# ODS bounded fresh row creation: declared design

The existing ODS `Builder` retains the worksheet model during authoring, the complete
`content.xml` string, and the returned ZIP byte vector. Batch 0433 measures that
path before production changes and compares it with a sequential scalar row
creation capability. This is fresh creation; logical append to an existing
worksheet, addition of a package Part, and arbitrary editing/repackaging remain
separate requirements.

The proposed ownership split keeps row/cell grammar in `litchi-ods` and XML audit,
manifest bookkeeping, and ZIP publication in `litchi-odf-common`. The existing
opaque-reader API must continue rejecting XML. A separate generated-XML path
accepts only a checked nested-element envelope and independently checked complete
child elements. It retains one bounded fragment buffer and feeds audited bytes
to the existing ZIP reader publication path. It never accepts an unaudited raw
stream as XML.

Required composition checks are declaration placement, one root per fragment,
matched envelope tags, no outside character data, globally aggregated XML bytes,
events, attributes and text, and insertion-depth plus fragment-depth limits.
Fragment EOF events are synthetic and do not increase the assembled document's
event count. Buffer errors must latch even if a producer ignores its returned
error. A producer cannot indicate EOF with a nonempty fragment. The suffix and
ZIP finalization are reachable only after successful producer EOF.

The ODS surface uses fixed namespace bindings, one ordinary `Sheet1`, and ordered
scalar cells. Rows may contain any count and ordering of the supported scalar
values within the explicit limits, including empty rows and an empty sheet.
The four-column Number/Text/Boolean/Empty shape belongs only to the benchmark. It must validate finite numbers and XML characters, preserve text
including carriage returns and whitespace, and charge explicit execution work,
objects, output and the retained XML window. Caller-owned input/iterator storage
and ZIP/auditor implementation allocation are distinct from the configured row
window; allocator observations will measure the complete operation region.

All output is caller-provided, sequential and non-seeking. Cancellation and sink
failures must retain accepted-byte progress and a typed cause. Fresh streaming
publication cannot roll back a sink prefix; failure must never finalize a package
that silently omits rejected input.

The comparison covers the same generated rows and scalar values across the
buffered and streaming paths. ZIP or XML lexical identity between implementations
is not assumed. Each implementation must have deterministic artifact hashes,
exact expected members/mimetype/manifest, and complete independent reopen/readback.
Normal binaries supply latency evidence; allocator binaries supply wrapped
requested-allocation observations. Process RSS includes setup and verification.
Before/after buffered controls distinguish shared-code effects from the new API.

This design is a coverage-driven bounded-memory enabler, not a claim that ODS is
the program's largest remaining measured latency bottleneck. The broader goal
and all other creation/append obligations remain active.

Source inspection found that the existing package author's default audit profile
permits 250,000 aggregate XML attributes. The initially proposed 131,072-row
matched corpus would exceed that limit. The matched large ODS corpus is therefore
32,768 rows, alongside 64 and 8,192 rows. This changes only the new ODS corpus;
existing XLSX shapes and production safety limits remain intact. The formal
protocol is frozen after confirming the supported shapes with pilots.

Independent review refined the first-row sequencing before implementation.
Envelope and limit preparation consume no producer input and occur before the
ZIP entry operation. The first row callback runs only after ZIP admission and
local-header publication. The generated reader audits the complete row before
returning any of that fragment's bytes. A first- or later-row failure poisons
the underlying ZIP writer; an incomplete header prefix is allowed under the
caller-owned sequential sink contract. The ODS operation drops the package on
every error. This avoids consuming a one-shot row source before transport
admission and needs no new ZIP preflight API or duplicate failed-state flag.
