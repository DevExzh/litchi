# DOCX bounded tail append: architectural obligations

The 30 previously read ADR/README files remain byte-identical;
`adr-refresh.json` records the current check. The required scenario is logical
append to an existing DOCX main story (goal category 6). This does not combine
fresh streaming creation, adding a package Part, or general repackaging into
one append claim. The shared OPC insertion implemented in 0482 supplies the
physical publication owner; DOCX must establish its semantic closure here.

| Authority | Required proof |
| --- | --- |
| ADR 0001, 0002, 0024 | Public values name plain paragraphs and tail placement. DOCX owns Word grammar and topology; OPC owns package publication; ZIP owns physical preservation. No archive implementation types enter the semantic API. |
| ADR 0003 | The original package remains immutable. Preparation proves the source and candidate before output. The plan carries source-bound scalar proofs and bounded authored material. Exact no-ops retain original bytes; immediate inverse authenticates the candidate before restoring the retained original. Durable recipes and multi-edit composition require their own evidence. |
| ADR 0005 | Complete main-document and candidate buffers and per-paragraph range indexes must be absent from the new route. Token admission precedes parser growth. Source, fragment, settings, parser, depth, event, paragraph, work, output, and execution budgets remain explicit. Package metadata and ZIP codec storage are separate costs; operation allocation measurements do not establish constant process RSS. |
| ADR 0006 | Strict and Transitional dialects must match package relationships. Changed signed, macro-enabled, external, protected, or unsupported dependent graphs fail closed. Final section properties are opaque retained bytes. No implicit normalization, repair, filesystem, network, or worker behavior is introduced. |
| ADR 0008 | Focused semantic, preservation, limits, source change, cancellation, sink failure, and inverse tests accompany applicable compiler, lint, documentation, boundary, and fuzz gates. Synthetic tests and internal readback do not establish native Office save/reopen acceptance. |
| ADR 0010, 0011 | The format proves raw decoded insertion offset and source/candidate digests through bounded readers. OPC verifies those proofs again and delegates measure/emit replay to ZIP, preserving every unselected physical member under the existing contract. |

The performance comparison must use the same deterministic source and append
the same plain text through the existing materialized control and new bounded
route. Normal and allocator timing remain separate. Timed regions include
opening the source adapter/package, preparation/editing, publication, sink
digest finalization, and owner destruction. Independent source/output semantic
and physical oracles remain outside timing. Process RSS includes those setup
and verification costs, warmups, report writing, and teardown.

Source-size scaling with one finite paragraph cannot prove output-size-
independent generation of a very large authored paragraph stream. That
remaining requirement needs a replayable producer or explicit caller-owned
replay source shared across validation, publication, and durable patch use.

The authored fragment now has an OPC-owned preallocation handoff. The package
reserves the requested buffer capacity before allocation, refuses excess
reported capacity, initializes in bounded cancellable work chunks, and exposes
only a fixed-length mutable slice. Preparation accepts the owner only from
the exact allocating package instance and transfers its reservation into the
plan. This closes the gap between format-owned encoding and package-owned
publication without retaining two charges for the same buffer. The existing
`Arc<Vec<u8>>` insertion route remains available with its existing admission
contract. These reservations describe requested storage, not allocator
bookkeeping or physical process RSS.

The shared XML auditor bounds duplicate-attribute scratch by the smaller of
the aggregate attribute ceiling and the bounded token window. The aggregate
counter and accepted XML policy remain unchanged. This is a source-derived
workspace bound; timing and allocator effects require the retained comparison
before any performance claim.

DOCX now names the archive-free `xml-minifier` audit profile directly when
deriving the OPC splice's finite XML limits. The dependency registry records
this edge; it introduces no ZIP implementation dependency and leaves package
publication in OPC. The initial boundary check correctly rejected the edge
before that registry entry was added.
