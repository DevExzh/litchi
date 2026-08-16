# Change 0164: RTF ordinary paragraph split/adjacent-merge evidence

Date: 2026-08-17

Status: opt-in correctness, phase, and sequential-sink evidence only. No
latency, speedup, allocation/RSS, transaction-memory, physical-I/O, cold-cache,
source-backed, real-producer, or general rich-RTF claim is accepted.

## Scope

The standalone harness adds two selectors for the committed native RTF
ordinary-body APIs:

- `rtf_semantic_split_paragraph_save`
- `rtf_semantic_merge_paragraph_save`

They raise the selectable matrix from 309 to 311 names while leaving the
historical default at 36 cases / 198 records. Each selector emits one record
for each of the tiny, medium, and large shapes. The selectors use the existing
`Edit::split_paragraph` and `Edit::merge_paragraphs` operations and do not add
a source-backed or eager/source comparison.

## Corpus and exact closure

Both operations reuse the exact generated plain lifecycle corpus rather than
the richer semantic read/edit corpus. The shapes contain 24, 200, and 10,000
ordinary paragraphs and 1,304, 10,808, and 540,008 source bytes respectively.
The tiny source identity is SHA-256
`73641cf09b630632deabce8585c67f395a6bd3ac01eedcca6a8b7224ef00d252`.

The target is deterministic: the middle paragraph is split at an interior
ASCII/UTF-8 boundary, and the merge joins that paragraph with its immediate
successor. Split inserts exactly the canonical five-byte `\\par ` boundary;
merge removes the authenticated exact adjacent boundary, so the expected
output is source length plus five or minus five bytes. Expected bytes are built
by an independent raw splice and are not obtained from the candidate snapshot.
The report records the before/after paragraph counts, selected positions and
split offset.

The edit closure is deliberately narrow: an uncompressed, ASCII, root-level,
contiguous ordinary body whose text maps one-to-one to source bytes and whose
boundaries are exact `\\par` controls. Compressed/LZFu and non-ASCII transport,
unknown or opaque syntax, nested groups, non-paragraph controls, tables,
fields, drawings, objects, pictures, shapes, form fields, bookmarks,
revisions, annotations, notes, math/custom XML, protection ranges, editable
regions, body-story events, external/transformation/mail-merge metadata, and
protected documents remain typed refusals. Formatting, encoded legacy text,
binary payloads, structural rich content, signatures and broad producer
compatibility are outside this proof.

## Timing and sink boundary

Each retained sample reports separate `open_ns`, `stage_ns`, `commit_ns`,
`publication_ns`, and complete `lifecycle_ns` vectors. The lifecycle interval
covers one `Document::from_bytes`, one checked split or adjacent-merge staging
call, `Edit::commit`, and public `snapshot().write_to` publication. Corpus
construction, independent expected-output splicing, reopen/readback, patch
encoding, refusal probes, and all correctness gates remain outside the timed
interval.

Publication uses the fixed 16-KiB `WindowedHashingSink`. It retains zero output
bytes, hashes the complete accepted stream, bounds each logical write window,
and records accepted bytes, write calls, and largest write. The window bounds
the publication sink only; it is not a transaction, candidate-snapshot,
allocator, process-memory, or physical-I/O bound. Digest finalization is after
the lifecycle timer stops.

## Untimed acceptance gates

Each selector publishes its evidence under the `SourceSummary`
`rtf_paragraph_split_merge` key and requires the following gates before
retained samples are accepted:

- complete semantic reopen and the exact split/merge paragraph projection;
- independent raw splice equality and unchanged surrounding source bytes;
- an empty-edit exact no-op and source immutability check;
- volatile patch forward application and exact inverse restoration;
- deterministic durable JSON serialization, durable forward application, and
  durable inverse restoration;
- stale- and foreign-source durable refusal;
- forged result-artifact refusal (`forged_result_artifact_refusal_verified`);
- the existing bounded invalid-selector/offset, unsupported-source, protected,
  and finite-limit refusal matrix;
- intentional partial-output and zero-progress sink failures; and
- deterministic source/output SHA-256 and per-sample publication digest checks.

The selector summary exposes both the combined `refusal_verified` gate and
`forged_result_artifact_refusal_verified`. The native RTF focused split/merge
tests remain the authority for exact boundary-byte restoration and
forged-boundary precondition refusal.

## Reproduction

Focused harness gate:

```sh
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  rtf_paragraph_split_merge_selectors_are_opt_in_bounded_and_gate_complete \
  -- --nocapture
```

All-shape debug smoke:

```sh
cargo run --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 \
  --semantic-shape tiny,medium,large \
  --rtf-variant plain \
  --case rtf_semantic_split_paragraph_save,rtf_semantic_merge_paragraph_save \
  --json target/perf/rtf-paragraph-split-merge-0164-smoke.json
```

The smoke is correctness and schema evidence only. A future performance
conclusion requires clean release binaries, CPU-pinned balanced ABBA samples,
identical corpus/output hashes, and separately justified resource and I/O
measurements.

## Remaining gaps

Rich formatting, non-ASCII/code-page and compressed changed publication,
tables, fields, drawings, objects, pictures, shapes, headers/footers,
annotations, revisions, protected/signed documents, malformed/security-heavy
corpora, insertion or cross-document composition, real Word/LibreOffice
producers, source-backed RTF, allocation/peak-memory/RSS, physical/cold I/O,
and general RTF merge/split behavior remain open.
