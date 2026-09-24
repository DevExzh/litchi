# Change 0765 evidence

[Record](../../0765-bom-prefixed-part-editing.md). Base `1d1044e3ac`; production
commit `2a2c4a9791`; test follow-up `67edd9b7ed`. Gates ran on `67edd9b7ed`.

| path | what |
| --- | --- |
| `sites/sites-at-base.txt` | every `buffer_position`/`error_position` grep hit in the surveyed crates at base (430 lines) |
| `sites/survey-summary.md` | the four read-only surveys condensed: verdicts per group, silent-corruption findings, disposition |
| `sites/sites-after.tsv` | every `buffer_position()` read at `2a2c4a9791` with its disposition: `origin` (converted through `ReaderOrigin`, or the local rule in litchi-xldm/litchi-formula), `intentional` (11 reads, each explained), `test`, `comment` |
| `census/bom_census.py`, `census/census.tsv`, `census/census-summary.txt` | members of the repository's OOXML fixtures that begin with a UTF-8 mark: 17 members in 3 of 336 packages |
| `base-failures/run1-compact-first.txt` | the new differential suites run on the base code (compact variant first): OPC 2/3, PPTX 6/8, DOCX 5/6, XLSX 3/5 fail |
| `base-failures/run2-indented-first.txt` | the same suites with the indented variant first, plus the new DrawingML theme test (fails at base) |
| `base-failures/theme-silent-corruption.txt` | the base's published theme bytes for a marked, indented theme: `</a:clrScheme>me>`, `</a:fontScheme>me>` and three lost indentation bytes before each scheme, well-formed and saved |
| `marks.txt` | per differential scenario, which marked members kept or dropped the mark (dropped only on parts regenerated from a model) |
| `abba/run_abba.sh`, `abba/run_confirm.sh` | ABBA drivers: A = base harness `058b58ba…`, B = branch harness `71453ca1…`, core 20, order A1 B1 B2 A2 A3 B3 B4 A4, `perf stat -x, -e instructions:u,cycles:u` per process |
| `abba/raw/*.json`, `abba/raw/*.perf` | the 24 primary reports and counters (`xlsx_first_cell` dense-wide 60 samples, `pptx_semantic_one_edit_save` large 40, `docx_semantic_one_edit_save` large 100) |
| `abba/confirm/*` | the PPTX confirmation run (80 samples per process) |
| `abba/analyze.py`, `abba/analysis.json`, `abba/confirm-analysis.json` | per-process p50/p95/mean, per-arm median of process p50s, paired ratios, bootstrap 95% CI (20,000 resamples, seed 765), drift, instruction and cycle ratios, output and corpus identity, >5% flags (none) |
| `binaries.sha256` | the two harness binaries (both legs built by the identical command from equal-length worktree paths) |
| `run_gates.sh`, `gates.txt` | gate commands and exit codes on `67edd9b7ed` (non-incremental builds) |
| `cleanup.json` | what was removed after the evidence was copied |

Reproduce the differential evidence with `cargo test -p litchi-opc -p litchi-xlsx
-p litchi-pptx -p litchi-docx --test byte_order_marked_parts`, `cargo test -p
litchi-drawingml --test theme` and `cargo test -p litchi-opc --test
reader_origin_contract`. For the base, copy the four `byte_order_marked_parts.rs`
files and `crates/litchi-drawingml/tests/theme.rs` into a checkout of
`1d1044e3ac`. The base runs used the test files as they were before the last
two OPC tests were added; line numbers in the base logs refer to those versions.

The timing numbers are evidence that the benign path is unaffected; no
performance claim is registered.
