# Profile r2 — timed-work attribution (base 009d515bef)

Written by the coordinator from the profile analyst's final report (the analyst's
harness could not write report files). Supporting data: `folded/` (gzipped
period-weighted stacks and analysis scripts that reproduce every figure),
`docx.summary.txt`, `xlsx.summary.txt`, `pptx.summary.txt`, `*-run.json`,
`stat-*.txt`. The raw `perf script` texts were deleted after folding.

## Method

Period-weighted perf samples restricted to stacks whose chain contains the timed
call (fresh writers: `write_fresh_{doc,ppt,xls}`; DOCX semantic: `text()` /
`edit_document` / `replace_paragraph_text` / `publish_document_edit` / `write_plain`
directly under `run_semantic_docx`; streaming: `write_pptx_stream`,
`write_streaming_xlsx`). Untimed `verify_semantic_docx` and opens are excluded
(verification alone was 56–77% of the DOCX-semantic profiles). DOCX streaming is from
`docx.summary.txt` (sample count, not period). Milliseconds = share × base p50
(streaming p50 from `base-streaming.json`: DOCX 92.1, XLSX 163.2, PPTX 186.0 ms);
profiled runs were 10–80% slower than base, so ms estimates are ±20%.

## Per-case attribution (% of timed cycles; K = kernel page faults)

| Case (p50 ms) | Components | Largest avoidable cost |
|---|---|---|
| docx_semantic_full_text (3.16) | namespace binding 44 (48 incl.); quick-xml 32; text extraction 19; UTF-8/memcpy/memcmp 3.5 | Linear prefix scan per element at `litchi-ooxml-common/src/binding_tracker.rs:318`, called from `litchi-docx/src/paragraph/codec/text.rs:106-119`; ~1.4 ms |
| docx_semantic_noop_edit_save (3.84) | `edit_document` scan 98 (model build 51, quick-xml 28, resolve_event 15); save 0.1 | Eager full-layout scan a no-op commit never uses: `litchi-docx/src/document/transaction.rs:897` → `scan_document_with_context` (:5729); ~3.7 ms |
| docx_semantic_one_edit_save (8.69) | scan 42.5; XML audit #1 22.0; XML audit #2 21.7; deflate+CRC of document.xml 10.9; misc 3 | The same 500 KB audited twice by `verify_source`: `transaction.rs:4707` (source gate) and `litchi-opc/src/pkgwriter.rs:228`; 1.9 ms each |
| doc_fresh_write_to (5.9) | per-character UTF-16 ~60 (12.7 of it `extend_from_slice` reserve checks); K 29; memcpy ~6 | Three per-character passes per run in `litchi-doc/src/writer/core/package/package.rs` (`encode_utf16().count()` :187 via codec.rs:4; offset loop to :243; per-unit `extend_from_slice` :246-247); ~3 ms |
| ppt_fresh_write_to (4.9) | UTF-16 count 27; memcpy 31; K 34; alloc 3 | `litchi-ppt/src/writer/escher/codec/text.rs:45` recounts UTF-16 units even on the ASCII branch (:27); ~1.3 ms, byte-identical fix |
| xls_fresh_write_to (3.9) | SipHash 45; K 37; memcpy 9; memcmp 4 | 32.7 KB shared strings hashed in `contains_key` and again in `insert`, map not reserved, strings cloned twice (`litchi-xls/src/writer/core/package.rs:127-130`), hashed/compared again at `stream/semantic.rs:34/43`; ~1.7 ms |
| docx_streaming_create (92.1) | budget 29–31; crc32 21.7 incl. (all scalar path); escaping 28 incl.; deflate 15.5; `encoded_text_size` preflight 12 | Per-operation `consume` CAS up the ancestor chain plus `Arc<Node>` clone/drop (`litchi-core/src/budget.rs:237`); ~27 ms |
| xlsx_streaming_create (163.2) | deflate 73 (76 incl.; `longest_match` 53); budget 13.7; memcmp 2.9; crc32 2.7; escaping/formatting 3.7 | Budget accounting in `write_row`; ~22 ms |
| pptx_streaming_create (186.0) | deflate 63 (Huffman build at block flush 42 incl., `build_tree` 25); memset 10; budget 5.8; SipHash 3.5; crc32 3.5 | Fixed per-member deflate cost over 16,421 small members: full state reset `soapberry-zip/src/writer.rs:2490` (6.4%, ~12 ms) plus a dynamic Huffman tree per member |

## Ranked opportunities

1. Budget leases/cheaper budget primitive (streaming DOCX 30%, XLSX 13.7%, PPTX 5.8%; ~60 ms total) — in progress as record 0752.
2. Coalesce tiny writes before CRC32/deflate in the streaming part writer (~18–20 ms on DOCX) — in progress as record 0752.
3. PPTX streaming per-member deflate fixed cost (12 ms reset, up to ~55 ms with tree building) — output bytes change if the strategy changes.
4. Lazy or memoized DOCX `edit_document` layout scan (~7 ms across no-op and one-edit) — the scan is also a structural admission check.
5. ASCII fast paths for UTF-16 in the DOC/PPT fresh writers (~4.3 ms, byte-identical).
6. Audit each DOCX publish once (1.9 ms per duplicate audit) — the writer's audit is the last line of defence; a proof must bind the exact published bytes.

Next in line: XLS fresh-writer SipHash (~1.7 ms), DOCX namespace-prefix cache for full text (~1.4 ms), fresh-writer page faults (12–17 MB touched for 4–5 MB output), `opc_mutated_save` incompressible deflate (60 ms; unprofiled).
