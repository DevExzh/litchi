# Change 0126: ordinary-root DOCX source-path evidence

Status: harnessed for correctness and logical-range evidence; no release
performance acceptance claim.

This change adds eight opt-in filesystem selectors for the ordinary unified
DOCX root:

```text
docx_file_eager_open             docx_file_source_open
docx_file_eager_paragraph_count  docx_file_source_paragraph_count
docx_file_eager_list_paragraphs  docx_file_source_list_paragraphs
docx_file_eager_full_text        docx_file_source_full_text
```

The fixed corpus is the existing
`build_docx_source_edit_corpus` output: 200 paragraphs, eight deterministic
incompressible 2 MiB media Parts, generator
`litchi-docx-source-edit-media-v1`, and archive SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.
The corpus builder and its bytes are unchanged.

## Timed scope

The eager open control times `fs::read` followed by
`litchi::Document::from_bytes`. The source control times
`litchi::Document::open(path)`, which adopts the immutable filesystem source.
For the six query controls, eager/source roots are prepared before the timer;
the measured interval contains only the named root query. Verification,
archive parsing, semantic digesting, source hashing, and the independent
source replay are outside the timer. Every sample retains the existing fresh
child-process and warm/cold-requested harness protocol.

## Independent source replay

Each source selector receives an untimed typed
`litchi_docx::source_backed::Package` replay over an instrumented positional
source. The replay records total and phase-specific calls, returned bytes,
request sizes, compressed-range coverage, and source-package successful-load
materializations. Opening the source-backed package has zero overlap with the
compressed main-document, media, unselected ordinary-Part, or core-properties
ranges. A semantic query first prepares the source-backed document; that
preparation completely covers the compressed main-document range, and the
query then reads no main-document, media, unselected, or core-properties range.
Each replay emits an explicit classification and fails the harness if the
classification is not met.

## Untimed correctness and preservation gates

The child verifies complete eager/source semantic parity for paragraph count,
paragraph text, full text, tables, ordered elements, and metadata. It also
checks exact source length and SHA-256 plus logical OPC part topology, package
and part relationships, content types, every Part blob hash, every media
payload hash, and source immutability. These gates do not claim byte-level ZIP
framing or member-order equivalence beyond the exact pinned source hash. Eager
controls intentionally have no source replay and mark
their generic logical-read counter scope as not applicable. No selected
paragraph claim is made.

The selectors are opt-in and raise the selectable `Case` matrix from 245 to
253 while preserving the default 36 cases / 198 records.

## Claim boundary

This is correctness and logical compressed-range evidence only. It does not
claim latency improvement, throughput, physical filesystem I/O, decompressed
bytes, allocations, RSS or peak memory, cold-cache behavior, ABBA/release
acceptance, broad security coverage, or Markdown performance. A release
latency or resource claim would require a separately frozen matched ABBA run
with the required controls and profilers.
