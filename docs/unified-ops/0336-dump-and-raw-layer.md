# 0336 — `dump` and the raw layer: parts, raw, add-part, raw-set, validate (audit WP5)

- Status: implemented on `uo2/dump-raw`, cut from `feat/unified-ops-api` at
  `0056b9cccd`; trunk merged at `3537339d66` (records 0334, 0335) and at
  `f2d89af7c2` (records 0333, 0337). Both merges conflicted only in the
  generated reports (and once in a corpus tool test); the reports were
  regenerated, every golden re-blessed to no change, and the corpus
  re-recorded to no change.
- Date: 2026-10-17
- Input: the definition-of-done audit of 2026-10-17 (§1 item 5, WP5); the
  goal's "OfficeCLI-compatible `apply`/`batch`/`dump`"; design §12 (raw
  escape layer); record 0015 W10, D19 and Q14.
- Scope: `crates/litchi-ops` (new `raw` module, `batch/dump.rs`, the
  capability verbs, one `Source` variant in `error.rs`, one dependency),
  `crates/litchi-opc` (`validation::validate_parts`), `crates/litchi-cli`
  (five verbs), `crates/pyo3-litchi` (six methods and stubs), one facade
  line (`raw` in the `litchi_ops` re-export list), one small validator fix
  in `crates/litchi-docx/src/validation.rs` (§4.2), the corpus tools and
  their generated reports.
- `performance_claim`: §6. Release builds, pinned with `taskset -c 7`,
  three processes, ABBA within each; evidence in
  [results/0336/](results/0336/).

## 1. Summary

| Item | Before | After |
|---|---|---|
| `File::dump` | absent | OfficeCLI v2 batch for DOCX, ODT, RTF, DOC, XLSX, XLSB, ODS, XLS, PPTX; typed refusal elsewhere |
| `File::{parts, raw}` | absent | OOXML parts by name or OfficeCLI alias, `.rels` members; packaged ODF members |
| `File::add_part` | absent | OOXML: OfficeCLI's `chart`/`header`/`footer` or any payload, with content type and relationship |
| `File::raw_set` | absent (record 0015: "needs an XPath engine") | OOXML XML parts: an XPath location-path subset, spliced, audited before publication |
| `File::validate` | absent | every format but iWork: leaf validators plus OPC part-graph checks, typed findings |
| CLI | — | `dump`, `raw`, `raw-set`, `add-part`, `validate` |
| Python | — | `File.dump`, `parts`, `raw`, `add_part`, `raw_set`, `validate`; stubs synced |
| `capabilities()` verbs | 18 | 23 (`dump`, `raw`, `validate`, `add-part`, `raw-set`) |
| OfficeCLI corpus: passed / gap / bug (against the trunk at `f2d89af7c2`'s parent) | 4,255 / 537 / 0 | **4,393 / 399 / 0** |
| corpus G1 (missing verbs) | 138 lines | **0** |
| v1 subset: passed / open | 2,184 / 105 | **2,241 / 48**; the 57 "planned WP5" lines are closed |

`raw_set` is implemented, not deferred: it writes through the audited OPC
publication path the brief names (§3.3).

## 2. `File::dump`

### 2.1 Shape

`File::dump() -> batch::Dump`. `Dump::items()` is a list of `BatchItem`s
whose first item is `{"command":"meta","dumpVersion":2}`; serialized, a
dump is the JSON array `File::batch`, `litchi batch` and Python's
`File.batch` read. Text follows OfficeCLI's v2 newline semantics: a soft
break inside a run is `\u{b}` (OfficeCLI's `\v`), and a `\n` in a shape's
`text` separates paragraphs.

The dump walks the node view and, per family, emits what rebuilds it on
`create(format)`:

- **Word processing** (DOCX, ODT, RTF, DOC): the styles the created document
  lacks (`add /styles`), then every body paragraph (`add` with the
  properties `add` takes, then a `set` with the remaining writable ones),
  its runs, hyperlinks, breaks and tabs, and tables (`add` with `rows`/
  `cols`, then each cell's paragraphs into the created cells; a short row
  `remove`s its extra created cells). DOC adds a paragraph with its text
  and formats its first run, because DOC has no run `add`; RTF keeps no
  paragraph formatting `set` after an `add` (its created body has no
  literal source to splice into).
- **Workbooks** (XLSX, XLSB, ODS, XLS): the first sheet is renamed, the
  others added, hidden sheets hidden last; cells are `set` by A1 path with
  their typed value (number, boolean, text) or formula and their writable
  formatting; row layout is `set` on `row[N]`. XLS sets formatting apart
  from the value, as its writer requires.
- **Presentations** (PPTX): slides are added by layout name; each source
  placeholder is matched to the created slide's placeholder of the same
  `phType` and `set`, created placeholders the source lacks are removed,
  other shapes are added and their text set; notes are set.

Every property is filtered through the registry: a dump carries only keys
the element's `add` (or `set`) accepts, and drops a read form the write
coercion refuses (a colour with its theme annotation; PPTX scheme colours,
which the PPTX writer refuses). Nothing is dropped silently: each skipped
node is an `unsupported_element` warning and each dropped property an
`unsupported_property` warning naming the property, in `Dump::warnings()`.
The messages name only the element kind and positional path. Native
identities are not carried: list numbering (`numId`) is dropped and
reported (ADR 0034), and a style reference to a style the replay target
will not have is dropped and reported.

Formats without a plan (PPT, ODP, OTH, ODM, ODG, ODC, ODI, ODB, ODF
formula, iWork) refuse with `unsupported_capability`, naming the format.

### 2.2 Round-trip evidence

`crates/litchi-ops/tests/dump_round_trip.rs`: `dump`, replay with `batch`
onto `create(format)`, and require the same text view **and the same dump
of the replay** (the dump is a fixed point, so every element and property
it carries survived). Workbook text views ignore formula cells (a replayed
formula has no cached value; the write verbs never invent one) and row
records without cells.

- `corpus_round_trips`: 13 corpus files: DOCX ×4 (`table_footnotes`,
  `Hyperlink`, `table-alignment`, `headerFooter`), XLSX ×3 (`SimpleStrict`
  with formulas, `InlineStrings`, `Formatting`), XLSB ×2, ODS, RTF ×2, XLS.
- `built_documents_round_trip` (DOCX, ODT, RTF, DOC),
  `built_workbooks_round_trip` (XLSX, XLSB, ODS, XLS),
  `built_presentation_round_trips` (three layouts, placeholders, shapes).
- `unplanned_formats_refuse` (ODP, PPT, ODG).
- `survey` (ignored; release): every corpus file up to 512 KiB.

Survey (release build, every corpus file up to 512 KiB that opens; the
classes are exclusive):

| Format | Files | Exact | Lossy (reported) | Replay refused | Text differs | Fixed point differs | Dump refused |
|---|---:|---:|---:|---:|---:|---:|---:|
| docx | 72 | 60 | 9 | 0 | 0 | 2 | 1 |
| xlsx | 176 | 154 | 0 | 0 | 22 | 0 | 0 |
| xlsb | 19 | 10 | 0 | 3 | 1 | 5 | 0 |
| ods | 13 | 12 | 0 | 0 | 1 | 0 | 0 |
| pptx | 80 | 55 | 24 | 1 | 0 | 0 | 0 |
| odt | 24 | 15 | 9 | 0 | 0 | 0 | 0 |
| rtf | 163 | 149 | 4 | 10 | 0 | 0 | 0 |
| doc | 44 | 11 | 0 | 33 | 0 | 0 | 0 |
| xls | 113 | 60 | 0 | 22 | 3 | 25 | 3 |

- DOCX "fixed point differs" (2, one file twice in the corpus): a style
  definition whose `font` and `font.ascii` read back merged.
- XLSX "text differs" (22): mostly numbers stored with more digits than
  their IEEE value needs (`42796.284856015031` reads back as the shortest
  form `42796.28485601503`, the same double) and date cells written back as
  their text.
- PPTX "lossy" (24): pictures, charts, tables, SmartArt and groups.
- The DOC, XLS, XLSB and RTF refusals are leaf write gaps (§7 F7). The
  three "dump refused" XLS files and one DOCX file do not open for edit.

"Exact" is the fixed point with an equal text view and no warning. "Lossy"
is a text difference the warnings account for (pictures, charts, fields,
content controls). The remaining classes are §7's follow-ups.

## 3. The raw layer

### 3.1 Part addressing and reads

`File::parts()` lists the package parts in name order; `File::raw(part)`
returns a `raw::Part` (name, content type, bytes). Both read the package
the handle *publishes* (`File::to_bytes`), so edits are visible. A part is
named by its OPC part name or by OfficeCLI's aliases:

- DOCX: `/document`, `/styles`, `/numbering`, `/settings`, `/fontTable`,
  `/theme`, `/footnotes`, `/endnotes`, `/comments`, `/webSettings`,
  `/header[N]`, `/footer[N]`, `/chart[N]`;
- PPTX: `/presentation`, `/slide[N]`, `/slideMaster[N]`, `/slideLayout[N]`
  (counted across masters), `/theme`, `/notesMaster`, `/noteSlide[N]`,
  `/slide[N]/<kind>[K]`;
- XLSX: `/workbook`, `/styles`, `/sharedStrings`, `/theme`, a sheet by name
  (caseless) or `/sheet[N]`, `/<sheet>/drawing`, `/<sheet>/chart[K]` (charts
  of the sheet's drawing), `/<sheet>/<kind>[K]`; XLSB: `/workbook`.
- `.rels` members (`/_rels/.rels`, `/word/_rels/document.xml.rels`),
  serialized from the parsed relationships.
- Packaged ODF: members by name and `/content`, `/styles`, `/meta`,
  `/settings`, `/manifest`, `/mimetype`. Flat ODF has no parts.

Indexed relationships (`[N]`) are ordered by where the source part's XML
first names their ids, which is document order (slide ids, sheets,
header references).

### 3.2 `add_part`

`File::add_part(parent, raw::NewPart) -> raw::AddedPart { file, part,
relationship_id }`. `NewPart::of_type("chart" | "header" | "footer")` adds
OfficeCLI's empty parts under the next free numbered name;
`NewPart::new(content_type, relationship_type, bytes).named(..)` adds any
payload. `parent` `/` relates from the package; an XLSX sheet's chart is
related from the sheet's drawing, as OfficeCLI does. XML payloads go
through `OpcPackage::try_add_owned_xml_part_bytes`, which audits them
before they gain source provenance; a taken name is `file_exists`.

### 3.3 `raw_set`

`File::raw_set(part, xpath, raw::Action, xml) -> raw::Edited { file,
matched }`, OfficeCLI's `raw-set`:

- The XPath is a location path: `/` and `//` steps of prefixed names or
  `*`, each with `[n]`, `[last()]` or `[@name='value']` predicates.
  Prefixes resolve through OfficeCLI's twelve (`w`, `r`, `a`, `p`, `x`,
  `wp`, `mc`, `c`, `xdr`, `wps`, `wp14`, `v`) overridden by the part
  root's declarations; anything else is `invalid_xpath`, and no match is
  `zero_matches`.
- Actions: `append`, `prepend`, `insertbefore`, `insertafter`, `replace`,
  `remove`, `setattr` (and OfficeCLI's `before`/`after`/`delete`).
- The fragment is parsed, its inter-element whitespace dropped (leaf text
  kept), and each top-level element declares the namespaces it uses that
  the insertion scope lacks. Unprefixed elements take the part root's
  default namespace, as OfficeCLI does.
- The part is **spliced**: every byte outside the edited ranges is kept.
  Matches nested in one another are refused for structural actions; the
  root cannot be removed or given siblings.
- Publication: `OpcPackage::try_replace_owned_xml_part_bytes`, the audited
  path records 0313 and 0317 use (`xml-minifier` verification before the
  payload gains provenance). An unchanged result returns the same handle.

The 48 corpus `raw-set` lines (PPTX decks built from raw DrawingML, DOCX
text boxes, an XLSX bubble chart) now pass.

### 3.4 `validate`

`File::validate() -> raw::Validation { format, report }` over
`litchi_core::ValidateReport`: DOCX (`litchi_docx::validate_read_at`),
PPTX (`litchi_pptx::validate_source`), XLSX/XLSB (`litchi_opc` ingress),
each merged with the new `litchi_opc::validate_parts` (every internal
relationship names a held part; every XML part is one well-formed
element); XLS, PPT, DOC (compound-file and record validators); RTF (syntax
and active-surface inventory); packaged ODF (`validate_package`). iWork
refuses (`unsupported_capability`; no report-producing validator exists).
`is_valid()` means no error or fatal finding. It is structural, not
OfficeCLI's Open XML SDK schema validation. A validator that cannot finish
is `raw::Error`, a typed `Source` of the unified error.

### 3.5 Writes and snapshots

The raw writes do not mutate their input. They publish the package with the
change and open the result as a new handle (`open_bytes`), so the new
handle has default `OpenOptions` and no origin path; the CLI saves it to
the input path (or `--out`). They are not steps of a `transaction::Edit`
and carry no `Patch` (§7 F2).

## 4. Surfaces

### 4.1 Rust, CLI, Python

- Rust: `litchi_ops::raw::{Part, PartInfo, NewPart, AddedPart, Edited,
  Action, Validation, Error}`, `litchi_ops::batch::Dump`; the facade gains
  `litchi::raw` (one token in the existing re-export list).
- `capabilities::Verb::{Dump, Raw, Validate, AddPart, RawSet}`; `add-part`
  and `raw-set` are edits (refused on signed and locked handles). Supported
  per §2/§3; refusals are `unsupported_capability` with follow-ups. All 23
  capability goldens re-blessed (verbs only); registry goldens unchanged.
- CLI: `dump <file> [/] [--format batch] [--out F]` (the array on stdout,
  warnings on stderr, so `dump | batch --input -` works), `raw <file>
  [<part>] [--out F]` (no part lists the parts), `raw-set <file> <part>
  --xpath X --action A [--xml X]`, `add-part <file> <parent> --type T`
  (OfficeCLI's `Created T part: relId=... path=...` line), `validate
  <file>` (exit 1 on an error finding, findings on stderr).
- Python: `File.dump(strict=False)`, `parts()`, `raw(part) -> bytes`,
  `add_part(...)` and `raw_set(...)` (Commits whose `value` is `{"part",
  "relationshipId"}` or `{"matched"}`), `validate()`; stub types
  `_PartInfo`, `_Validation`, `_ValidationIssue`; `test_raw_dump.py`.

### 4.2 A validator false error, fixed

The corpus's `word/charts.sh` `validate` reported
`docx.main_document.relationship_target_content_type`: the DOCX validator
required `application/vnd.openxmlformats-officedocument.package` for every
`package` relationship, but a chart's embedded workbook is typed with its
own format (`spreadsheetml.sheet`), as Word writes it. The validator now
accepts the OOXML document types there (`litchi-docx/src/validation.rs`,
one predicate; `raw::tests::docx_with_a_chart_validates`).

## 5. OfficeCLI corpus and dispositions

Re-recorded after each trunk merge. "Before" is the trunk's committed
report at the time; the first row pair is against the base trunk
(`0056b9cccd`), the second against the trunk of the final merge (records
0333 and 0337 had moved other classes).

| | Lines | Passed | Warned | Gap | Delta | Cascade | Bug |
|---|---:|---:|---:|---:|---:|---:|---:|
| base trunk | 6,431 | 3,948 | 1,178 | 663 | 201 | 441 | 0 |
| + this record | 6,431 | 4,086 | 1,178 | 525 | 201 | 441 | 0 |
| final trunk | 6,431 | 4,255 | 1,287 | 537 | 201 | 151 | 0 |
| + this record | 6,431 | **4,393** | 1,287 | **399** | 201 | 151 | 0 |

- G1 "missing verbs": 138 → 0 (`validate` 88, `raw-set` 48, `add-part` 1,
  `dump` 1). G1 still classifies any OfficeCLI verb the CLI lacks.
- v1 subset (3,443 lines): against the base trunk passed 1,877 → 1,934,
  open 625 → 568; against the final trunk passed 2,184 → 2,241, accounted
  1,154 → 1,154, open 105 → 48. The "planned WP5" blocker (57 lines) is
  gone.
- `tools/officecli_dispositions.json`: the stale `raw-layer-verbs` entry
  and the `corpus:G1|dump`/`corpus:G1|*` tiers are removed; `raw-layer`
  (`delta:*:raw`) becomes `out-of-scope` (ref 0336): the raw layer is a set
  of commands, not an element. The capability delta's raw rows move from
  planned to out-of-scope. Both `--check` runs are green.

## 6. Measurements

`examples/dump_bench.rs`: each input opened once; the baseline walks the
whole node tree reading every node's `format` and text (what `dump`
reads); the dump leg is `File::dump`. Medians of 10 rounds, three
processes:

| Input | Items | Traversal (µs) | Dump (µs) | Ratio |
|---|---:|---:|---:|---:|
| DOCX, 400 paragraphs (`drawing.docx`) | 1,487 | 11,649–11,702 | 16,798–17,031 | 1.44–1.46 |
| XLSX, 75,770 cells (`no_drawing_patriarch.xlsx`) | 74,512 | 237,780–239,246 | 365,738–369,126 | 1.53–1.54 |
| PPTX, 200 generated slides | 1,001 | 3,351–3,446 | 7,142–7,983 | 2.13–2.32 |

- A first version looked each element and property up in the registry per
  node: 3.37× on the XLSX and 1.61× on the DOCX. Looking each element
  kind's writable properties up once fixed it (`509a56301d`).
- The PPTX fixed cost is the replay-target probe: `dump` creates one blank
  deck and adds one slide per distinct layout to learn which placeholders
  the layout instantiates (callgrind: 58 M of 130 M instructions per dump
  on this input).
- There is no budget for `dump` in the goal; this is the recorded baseline.
- Finding for the PPTX owner (not this record's code): building the
  200-slide input through `batch` costs about 90 G instructions, about 82%
  of it in `opened::Edit::insert_slide` capture and commit, so replaying a
  large PPTX dump is superlinear.

## 7. Gaps and follow-ups

- F1 `dump` of a subtree (`dump <file> /body/p[3]`, OfficeCLI's subtree
  paths) is refused (`unsupported_path`).
- F2 `raw_set`/`add_part` are not `transaction::Edit` steps: no `Patch`,
  no inverse, no place in a batch. The batch channel still refuses `raw`,
  `raw-set`, `add-part`, `validate` and `dump` items (pending), and the new
  handle drops the source `OpenOptions`. A `Leaf` variant for a republished
  package is the follow-up.
- F3 Raw writes for ODF (through the ODF package publisher); raw reads of
  legacy compound-file streams.
- F4 `validate` is structural: no OOXML schema validation (the Open XML SDK
  schema OfficeCLI runs), no iWork validator.
- F5 Dump plans for PPT and ODP (and the ODF object formats): their write
  verbs cannot yet rebuild slides and placeholders on a created file.
- F6 Dump content not rebuilt (reported as warnings): pictures, charts,
  SmartArt, comments, fields and content controls, headers/footers and
  notes stories, list numbering (native ids), DOCX style definitions the
  created document already has (`set` cannot clear the properties a source
  style leaves unset), PPTX placeholder geometry (the PPTX writer sets
  geometry only on a complete local transform), PPTX scheme colours.
- F7 Leaf replay refusals the survey found (owners: XLS, XLSB, DOC, RTF,
  PPTX): XLS formulas the BIFF8 encoder cannot tokenize and fills outside
  the palette; XLSB formulas re-read with an extra pair of parentheses and
  cell font sets that do not read back; DOC paragraphs on fast-saved or
  complex files; RTF paragraph sets on a created body; XLS row records.
- F8 A large DOCX dump replayed as one atomic batch can exceed the default
  `max_snapshot_bytes` (each step retains the main part); replay in chunks
  or raise the limit.

## 8. Gates

Run in full before the final merge and again after it (the clippy feature
subsets after it: none and `pptx`). Every exit status is 0.

| Gate | Command | Exit |
|---|---|---:|
| fmt | `cargo fmt --all --check` | 0 |
| clippy | `cargo clippy -p litchi-opc -p litchi-docx -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| clippy subsets | `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F ∈ none, docx, xlsx, pptx, xlsb, doc, xls, ppt, rtf, odt, ods, odp, odg, pages | 0 each |
| rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-opc -p litchi-ops -p litchi-cli -p litchi-docx -p litchi` | 0 |
| tests | `cargo test -p litchi-opc --all-features`; `-p litchi-docx --all-features --lib --test validation`; `-p litchi-cli --all-features`; `-p litchi-ops --all-features` (`--lib`, `--doc`, then each `--test`, 160 targets after the final merge); `-p litchi --all-features` | 0 each |
| goldens | `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every feature but `encryption` | re-blessed (verbs only) |
| corpus | `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py --check` | 0 |
| delta | `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI --check` | 0 |
| boundaries | `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0 |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| tool units | `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate` | 0 |
| Python | `crates/pyo3-litchi/tests/run_tests.sh` (292 tests after the final merge) | 0 |
| harness | `cd tools/perf-baseline && cargo check --locked --all-targets` (lock gains `quick-xml` under `litchi-ops`) | 0 |

## 9. Not claimed

- OOXML schema validity of anything `raw_set` or `add_part` writes: the
  audit proves well-formed, namespace-consistent XML, not schema order
  (an `append` after `w:sectPr` is accepted, as in OfficeCLI).
- That every dump replays: §2.2's survey and §7 F7 list the formats and
  inputs whose leaf writers refuse items; `--best-effort` replays the rest.
- Byte identity of a replay with its source: a dump rebuilds the readable
  tree, not the package.
