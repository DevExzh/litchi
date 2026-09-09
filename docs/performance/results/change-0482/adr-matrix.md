# Bounded XML audit and decoded OPC splice: architectural obligations

The 30 previously read ADR/README files are byte-identical to the prior batch;
`adr-refresh.json` records the fresh check. This batch implements shared XML
validation and physical publication needed by the existing-document append work in
`../change-0481/window-contract.md`. It does not certify the DOCX transaction.

| Authority | Obligation and verification |
| --- | --- |
| ADR 0001, 0002, 0024 | XML auditing stays in `xml-minifier`. The diagnostic harness alone gains a direct dependency. No physical ZIP types enter DOCX. The boundary gate verifies repository ownership. |
| ADR 0003 | The reader is an auditor and never mutates input. OPC insertion plans authenticate the original decoded member, fragment, candidate and source version. Successful publication records the exact candidate archive fingerprint; inverse output authenticates that archive before restoring the retained original source. Exact insertion no-ops copy the original archive. DOCX transactions and durable serialized patches remain subsequent work. |
| ADR 0005 | Callers supply the reader and finite total/token/depth/event/text/attribute limits. Token admission must precede proportional parser growth; retained parser state must have an explicit finite envelope. Source or caller buffering is counted separately from auditor storage. No implicit storage, networking, workers or runtime is introduced. |
| ADR 0006 | Existing authored compactness rules, including inherited `xml:space`, lexical tag whitespace and DTD refusal, must remain enforced. Differential tests compare slice and reader acceptance and successful reports. Streaming may report an earlier encountered defect before an unseen suffix's UTF-8 or total-byte defect; this is documented rather than presented as identical global error priority. |
| ADR 0008 | Focused malformed, limit, I/O and chunk-boundary tests, producer XML differential checks and ASan/libFuzzer runs accompany compiler, lint, rustdoc and boundary gates. OPC tests cover opaque member preservation, source conflicts, exact no-ops, inverse restoration, cancellation, quotas and partial output. These are substrate checks; native Office append and durable serialized inverse publication remain unproven. |
| ADR 0010, 0011 | OPC owns decoded source/fragment/candidate proof checks, XML publication policy and source-bound replay. ZIP owns physical preservation and private-layout memory bounds. No physical ZIP types enter format-facing CRUD APIs. Changed signed/encrypted input is refused; exact no-op copies preserve the source artifact. |

Managed inverse output uses the current candidate's execution context when
present, otherwise the retained original context. I/O, workspace and output
are charged to that selected context once; source freshness and cancellation
remain checked for both retained sources. Fragment hashing and replay consume
bounded work chunks. ZIP decoder, replay and preservation ownership are reserved
separately from XML audit and adapter windows. Their scalar bounds cover the
locked backend and explicitly named owned storage, not process RSS or arbitrary
caller source/sink allocations.

The benchmark compares two routes from the same deterministic reader:
materialize then use the existing slice auditor, or use the new reader auditor.
Both run in the same build. This comparison can establish measured primitive
cost and storage scaling for the named corpus. It cannot establish reduced
DOCX append latency, constant package memory, cold-cache behavior, remote
scaling or a program-wide speedup. Normal and allocator timings remain separate.
