# Evidence: change 0650, the DOCX editor's byte-order-mark offset skew

Change record:
[`0650-docx-editor-byte-order-mark-admission.md`](../../0650-docx-editor-byte-order-mark-admission.md).

Disposition: retained, a correctness fix. `performance_claim: none`. Three
source files and one test file under `crates/litchi-docx/src/writer/doc/`
changed; the counts here are deterministic admission censuses and constructed
witnesses, reported as evidence and not registered as claims. No timing was run.

## Contents

| Path | What it is |
| --- | --- |
| `probe/src/main.rs` | The admission probe. Ten aspects per DOCX-family fixture, all through documented entry points: `Package::open`; `Package::document()` with `paragraph_count()`/`text()`; whether the part blob and the reader's normalized view still carry a mark; `Document::sections()`; `Package::document_mut()`; `save` with no edit; `document_mut()` then `save` with no edit; the managed `edit_document` + `insert_paragraph` + `publish_document_edit` + `save` route; `document_mut().add_paragraph_with_text` + `save`; and reopening the edited artifact. Each save aspect opens the package fresh and reports the published SHA-256 and byte count, or the verbatim typed refusal. Two extra modes: `fromxml` runs `MutableDocument::from_xml` on a raw XML file, and `partxml` writes `/word/document.xml` exactly as the opened package holds it. |
| `probe/Cargo.toml.template` | The probe manifest. `REPLACE_WITH_CHECKOUT` is substituted with each leg's tree, so both legs are the same probe source against different crates. No lockfile: both legs resolve the same versions through the same path dependencies. |
| `census/census-before.txt` | The probe's `census` over `test-data/` on the base checkout `c7326f680`. 63 fixtures × 10 aspects = 630 lines. |
| `census/census-after.txt` | The same over the same corpus on this branch. |
| `census/census-diff.txt` | `diff` of the two. **Three changed lines**, all on `alt-chunk-header.docx`, all refusal text; every published SHA-256, paragraph count and text length is identical. |
| `census/variants-before.txt` | The probe's `census` over the four constructed packages on the base checkout. 4 × 10 = 40 lines. |
| `census/variants-after.txt` | The same on this branch. |
| `census/variants-diff.txt` | `diff` of the two: six changed lines, all on the two byte-order-marked members. |
| `census/bom-census.txt` | Whether each of the 63 fixtures' `word/document.xml` starts with a UTF-8 byte order mark, with its length. Exactly one does. |
| `witness/alt-chunk-header-document.xml` | `word/document.xml` of `alt-chunk-header.docx`, byte-identical to the zip member and to what `Package::open` hands the editor (sha256 `f573a236ac24f491ef2728208b8fdb7d63c70013f57aeae95dd737ae6ad840ee`, 3,476 bytes, mark included). |
| `witness/w1-paragraph.xml` … `w5-altchunk-only.xml` | Five unmarked constructed bodies: one paragraph; paragraph + `altChunk`; paragraph + `sectPr` + `altChunk`; paragraph + `altChunk` + `sectPr`; `altChunk` alone. |
| `witness/bom-w1-paragraph.xml`, `bom-w2-altchunk.xml`, `bom-w4-altchunk-then-sect.xml`, `bom-w5-altchunk-only.xml` | The same bodies with a three-byte mark prepended and nothing else changed. `bom-w1-paragraph.xml` is the minimal witness: one paragraph, refused before, accepted after. |
| `witness/b1-nobom.xml` | The fixture's own part with the mark removed and nothing else changed — the bisection step that separates the two triggers. |
| `witness/b3-noaltchunk.xml`, `b4-noignorable.xml`, `b5-nosectpr.xml` | The other three bisection steps: the fixture without its `altChunk` line, without its `mc:Ignorable` attribute, and without its `sectPr` block. |
| `witness/witness-outcomes.txt` | `MutableDocument::from_xml` on all fourteen witnesses, both legs, verbatim. |
| `witness/build-variants.py` | Rebuilds the four constructed packages from two tracked fixtures. Each pair differs only in the mark; `a0/a1` come from a main document with no markup-compatibility markup, `b0/b1` from one that has it. |
| `witness/bom-presence.py` | Regenerates `census/bom-census.txt` from a `test-data/` root. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `follow-ups.md` | The four questions this change witnessed and did not resolve, each with its site, its witness and what a resolution would have to decide. |
| `gates.txt` | The tail of every gate. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

- Base: `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` (change 0630), branch
  `feat/office-format-completeness`.
- Branch: `perf/0650-docx-document-mut-admission`, which carries this packet in
  its single commit.
- Two legs, both built `--release`:
  - **before** — the shared read-only checkout
    `/home/zhuhe/code/litchi-worktrees/before-c7326f680`, with
    `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0650-probe-before`.
  - **after** — this branch's worktree
    `/home/zhuhe/code/litchi-worktrees/0650`, with
    `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0650-probe-after`.
- Probe binaries, SHA-256:
  - before `cec0c18cb7c8c9c0811993d571824d6e0d067be05034e7e01f93d612f880dbc7`
  - after `80f8a7c3ebbbae82c1086c2948e68705d572762e9086c3b479bef1adbb69a740`
- Both censuses were run against the **before** checkout's `test-data/`, so the
  corpus bytes are identical for both legs.
- Fixtures cited, SHA-256:
  - `alt-chunk-header.docx` `ee33b932a6c31f430bc8a6e450592970419b0bcc155e5c9884065602b9c733b8`
  - `ooxml/docx/documentProperties.docx` `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5`
  - `ooxml/docx/Hyperlink.docx` `ebd1f0ae1a9d6ef35c4b301209b135a0bd2f8dfde9c44270bbbea1cd19b3fd66`
  - `libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx` `6e71a4d554ea90df1ddcc9b6fbbcc677a94c6617cb35f9e5b7d4ccab46f1a79e`
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0
  (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd 2026-03-21), quick-xml 0.41.0.
  Every measured process pinned with `taskset -c 18`. The host was shared with
  seven other agents; no timing was taken, so contention cannot affect any
  figure in this packet.

## Reproducing

```sh
# one leg
cp -r probe /tmp/0650-probe && cd /tmp/0650-probe
sed 's|REPLACE_WITH_CHECKOUT|<checkout>|' Cargo.toml.template > Cargo.toml
CARGO_TARGET_DIR=<target> cargo build --release

# the corpus census
target/release/docx-editor-admission-probe census <checkout>/test-data <scratch>

# the constructed variants
python3 witness/build-variants.py <checkout>/test-data <variants>
target/release/docx-editor-admission-probe census <variants> <scratch>

# one witness
target/release/docx-editor-admission-probe fromxml witness/bom-w1-paragraph.xml
```

The census writes one temporary artifact per save aspect into `<scratch>` and
deletes it as soon as it has been digested.
