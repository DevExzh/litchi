# 0838 — DOCX durable inverse physical restoration diagnostic

Current production code reproduces the physical restoration limitation that
blocked the fresh-compression candidate in 0837. A serialized paragraph copy
or removal inverse restores every decompressed member exactly, but does not
necessarily restore the original compressed main-document bytes. Retained-source
undo restores the entire original archive in every measured diagnostic case.
No production code, wire format, contract, or existing test changed.

## Controlled comparison

The standalone probe is built against committed source
`4435bae2c9727f52d8bf1af75a4b788d2fccd5c9`. Its only input is the four-member
DOCX retained from 0837, with two paragraphs (`one`, `two`) and an unused
32 KiB binary part. It retains that input and creates two controls using current
public APIs: `OpcPackage::from_bytes` plus `PackageWriter::to_bytes`, and
`StreamingArchiveWriter::start_entry` for every member with Deflate. All three
inputs have identical member names and decompressed bytes. The current owned
staged control is also byte-identical to the retained source: these are three
source paths yielding two physical encodings, not three distinct encodings.

Each of three fresh serial processes copies paragraph 0 before position 2 and
removes paragraph 0 on all three sources. Every case exercises forward
publication, serialized inverse replay, retained-publication inverse, canonical
wire re-encoding, and rejection of an archive changed outside the main document.
The rejected 0837 OPC candidate is absent from this build.

| Source variant | Copy durable exact | Removal durable exact | Main compressed bytes, source → inverse |
| --- | --- | --- | --- |
| Retained 0837 staged source | No | No | 135 → 141 |
| Current ordinary borrowed regeneration | Yes | Yes | 141 → 141 |
| Current owned staged writer | No | No | 135 → 141 |

The table holds in all three processes. Across all 18 cases:

- Forward paragraph sequences are correct; every durable inverse restores all
  decompressed member bytes exactly.
- Retained-publication inverse restores the complete source archive exactly.
- Changing a non-main part makes durable replay return `StaleSource` without
  writing to its sink.
- Every durable wire re-encodes canonically after decoding.
- Independent Python ZIP/XML replay validates 63 archives and 18 patch wires.
  All unselected local-header and compressed-payload spans remain identical
  through forward and durable inverse publication. Only `word/document.xml`
  has different compressed payload bytes in the non-exact durable cases.
- All retained source, output, and wire bytes are identical across processes.

These are correctness diagnostics, with zero timed operations. The input is one
small synthetic DOCX, not an external-producer corpus or a universal inverse
qualification. The independent reader checks ZIP/XML content and wire magic,
version, length, and hash; canonical wire re-encoding and stale-source refusal
are assertions executed by the Rust probe. The stale-source experiment changes
a non-main payload; it does not isolate a logically identical archive with a
different physical encoding.

## Contract and optimization implications

In `paragraph_copy.rs`, copy/removal durable wire v1 stores XML snapshots,
operation state and limits, and an expected whole-artifact fingerprint. On
publication, that fingerprint is replaced with the emitted archive fingerprint.
It authenticates the input to inverse replay; it does not encode the original
compressed stream. Replay regenerates the edited main member through the current
encoder. The retained-publication inverse instead retains a `SourceArtifact`
and copies that original archive after checking the published input.

The public retained inverse explicitly promises exact original-package
restoration. General semantic reversibility and whole-artifact input checking
do not by themselves imply arbitrary ZIP-byte reconstruction from the durable
wire. Existing copy and removal tests nevertheless require serialized inverse
whole-archive equality for ordinary freshly authored fixtures. Those tests stay
unchanged, and the 0837 candidate stays rejected: changing fresh OPC output
would regress a currently passing composition even though a similar limitation
already exists for other physical sources.

Future compression-policy work needs a deliberate physical restoration design
and compatibility qualification before timing. Original payload/provenance or
an authenticated original-artifact provider are possible design inputs; a hash
alone cannot reconstruct arbitrary original compressed bytes. This batch does
not select or implement such a design. Byte-transparent optimizations remain
available independently. OLE2/OOXML performance work remains active; iWork is
excluded.

## Validation and evidence

Fresh probe formatting, all-target compilation, strict Clippy, rustdoc, and
release build pass. Initial Clippy rejected one unused import; that source and
receipt are retained, followed by the passing `probe-v2-*` checks. The production
source inventory is identical to 0837's final validated inventory, so the
previous affected-owner quality gates (6,847 passing tests, 38 ignored) are
explicitly reused rather than claimed as rerun. Probe dependencies are pinned
to versions/checksums already in the workspace lockfile.

The [packet](results/change-0838/README.md) contains the frozen plan and source,
build/admission bindings, all process receipts, all physical outputs, independent
reader, closure, cleanup, and seal. Run its `audit.py` to reproduce the retained
artifact analysis without compiling Rust. No latency, allocation, CPU, RSS,
cold-cache, or speedup claim follows.
