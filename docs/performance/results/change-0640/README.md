# Evidence: change 0640, the seven refused `.doc` fixtures of change 0587's defect 4

Change record: [`0640-doc-field-table-flag-refusals.md`](../../0640-doc-field-table-flag-refusals.md).

Disposition: retained, correctness fix. `performance_claim: none`. Two of the
seven refusals are removed because the bits they check are redundant with a
structure the reader already validates; five are confirmed correct and
unchanged. Everything in this packet is deterministic: structural decodes of
fixture bytes, a corpus differential, a syscall census and callgrind isolation
pairs. No timing leg was run, because nothing is claimed that would need one.

## Contents

| Path | What it is |
| --- | --- |
| `probe/main.rs`, `probe/Cargo.toml` | The measurement driver. `digest-doc` is change 0596's differential oracle **verbatim**, so the digests here are directly comparable with `results/change-0596/differential/`. `fld-dump` decodes every `Plcfld` straight from the Table stream, independently of `litchi-doc`'s field parser, and annotates each end marker's `fNested`/`fHasSep` against the structure; `stsh-dump` lists every style through a tolerant parse and marks duplicates; `open-leniency` reports the strict and lenient open outcome; `provenance` prints the FIB identification; `text-dump` / `text-dump-lenient` print `Document::text()`; `profile` loops the eager open for callgrind. The `Cargo.toml` names the repository crates; each leg was built from a copy whose path dependencies pointed at that leg's checkout. |
| `structure/fld-watermark.txt` | `fld-dump` of `ole/doc/watermark.doc`: the `TOC` field with its five nested `HYPERLINK` fields, every marker's CP, `Fld` bytes, decoded flags, the agree/disagree verdict per end marker, and each field's instruction text. This is the evidence that the five disagreeing markers are ordinary nested hyperlinks in a table of contents. |
| `structure/fld-poi-test.txt` | The same for `poi/test-data/document/test.doc`: one ` SEQ CHAPTER \h \r 1` field, two markers, no separator, `grffldEnd = 0x80`. |
| `structure/fld-corpus-scan.txt` | `fld-dump` run over all 57 `.doc` fixtures under `test-data/`, reduced to one line per file: end-marker count, `fNested` disagreements, `fHasSep` disagreements. The population table in the record is this file. Five files report `LOADERR` because this probe's own loader cannot reach their Table stream: `redline-1.doc` (invalid byte order), `cfb-v3-uninitialized-size-high-word.doc` and `TestMickey.doc` (Word 95), `word6-no-table-stream.doc` and `TestBug52117.doc` (Word 6.0). The three password-protected fixtures do load but carry an empty `fcPlcfFldMom`, so they contribute **zero** markers and cannot inject ciphertext into the population: the totals are the same with and without them. |
| `structure/stsh-{duplicate-style-names,footnote,lists-margins,picture,pictures_escher}.txt` | Every style's `istd`, `sti`, kind, name, aliases, base and next style for the five duplicate-style fixtures, read through `Leniency::TolerateStylesheetDefects` so the duplicate does not hide the rest of the sheet, with the strict parse's exact error and the `ToleranceReport` at the top. |
| `structure/stsh-{watermark,poi-test}.txt` | The same for the two field-table fixtures. Both stylesheets parse strictly; these are retained for the producer evidence (`WW8Num*z*`, `Internet Link`, `Absatz-Standardschriftart` beside `Default Paragraph Font`). |
| `structure/provenance.txt` | The producer evidence: the FIB identification for all seven fixtures (probe mode `provenance`), their OLE property sets and `\x01CompObj` user type read with the python `olefile` module, and the `WW-` / `WW8Num` stylesheet signature counts, with a reading of what each does and does not establish. |
| `structure/open-leniency.txt` | The strict and lenient open outcome for all seven fixtures, on the **unmodified base commit**. This is the answer to change 0587's open question. |
| `differential/digest-before.tsv`, `digest-after.tsv` | One line per `.doc` fixture under `test-data/` (57 files): open outcome or the exact `Debug` text of the refusal, `Document::text()` length and hash, paragraph count and a hash over every paragraph's `Debug` form (which carries its resolved properties, its runs and each run's character properties and revision marks), `paragraph_count()` from the separate public query, the section count and a hash of the section table, a hash of `get_all_subdoc_ranges()`, and the length and hash of `FileInformationBlock::raw_data()`. |
| `differential/digest-diff.txt` | The whole before/after difference after the absolute `test-data` prefix is normalized: two changed records, both on fixtures that were already refused and still are. 42 admitted and 15 refused on each leg; every one of the 42 is identical in every column. |
| `counts/counts-summary.txt` | The syscall census and the callgrind isolation pairs, with the reason both instruction deltas are read as code-layout drift. |
| `counts/strace-{before,after}-{utas,floating}.txt` | `strace -c` for one eager open on each fixture and leg. Identical censuses. |
| `counts/binaries.sha256` | The two staged probe binaries the counts used. Both were copied out of their Cargo target directories before measurement (change 0627). |
| `text/{fixture}-litchi-lenient.txt` | `Document::text()` for each of the five fixtures leniency admits, read through `TolerateStylesheetDefects`. |
| `text/{fixture}-libreoffice.txt` | LibreOffice 26.2.5.2's `txt:Text (encoded):UTF8` export of the same five files — the independent reading. |
| `text/plausibility.txt` | The word-sequence comparison of the two, with each difference read. Two fixtures are word-for-word identical; the other three differ only by generated list labels, field instructions, and note stories. |
| `follow-ups.md` | The two refusals this change unmasked, decoded from the bytes, with a **provisional** reading of each and the reason neither was acted on. Nothing there is authorized by this change. |
| `gates.txt` | Tails of every gate run in the candidate worktree. |
| `log-sections.md` | The four log paragraphs for the coordinator to merge into `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`. |
| `decision.json` | The `litchi-perf-change-decision` record. |

## Provenance

- Base commit: `c7326f680`
  (`docs(perf): close the first wave on the 0587 queue and refresh it (0630)`).
- Candidate branch: `perf/0640-doc-facade-field-table-refusals`.
- Control leg built from the shared read-only checkout at the base commit
  (`/home/zhuhe/code/litchi-worktrees/before-c7326f680`) with its own
  `CARGO_TARGET_DIR`; candidate leg built from the candidate worktree. Both
  `--release`.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws;
  rustc 1.95.0; valgrind 3.26.0; LibreOffice 26.2.5.2 620(Build:2).
  Every measured process pinned to CPU 16 with `taskset`, `RAYON_NUM_THREADS=1`
  for the callgrind runs. Seven other agents were building and measuring on the
  same host throughout.
- The probe source retained here names `/home/zhuhe/code/litchi/crates/...` in
  its `Cargo.toml`, as change 0596's did; each measured leg used a copy of that
  file with the prefix rewritten to that leg's checkout.

Binary sha256 (`counts/binaries.sha256` is the authoritative copy):

| binary | leg |
| --- | --- |
| `doc0640-before` | control, built against `litchi-worktrees/before-c7326f680` |
| `doc0640-after` | candidate, built against `litchi-worktrees/0640` |

Fixtures, all already in the tree:

| fixture | bytes | role |
| --- | ---: | --- |
| `test-data/ole/doc/watermark.doc` | 20,480 | `fNested` witness |
| `test-data/poi/test-data/document/test.doc` | 32,768 | `fHasSep` witness |
| `test-data/ole/doc/duplicate-style-names.doc` | 64,512 | duplicate-style fixture |
| `test-data/ole/doc/footnote.doc` | 9,728 | duplicate-style fixture |
| `test-data/ole/doc/lists-margins.doc` | 10,752 | duplicate-style fixture |
| `test-data/ole/doc/picture.doc` | 1,448,448 | duplicate-style fixture |
| `test-data/ole/doc/pictures_escher.doc` | 120,320 | duplicate-style fixture |
| `test-data/poi/test-data/document/au.edu.utas.www___data_assets_word_doc_0003_154335_International-Travel-Approval-Request-Form.doc` | 86,528 | counts, 67 field end markers |
| `test-data/ole/doc/FloatingPictures.doc` | 335,360 | counts, 9 field end markers |
| all 57 `.doc` files under `test-data/` | — | corpus scan and differential |

## Cross-check worth knowing

`differential/digest-before.tsv` is **byte-identical**, after path
normalization, to `results/change-0596/differential/digest-after.tsv`. That
checks two things at once: this packet's oracle is change 0596's oracle, and
nothing landed between change 0596 and this base commit that moved any DOC
observation in the digest.
