# Evidence: change 0587, the whole-path opportunity survey at `2fc5fc657`

Change record: [`0587-remaining-opportunity-survey.md`](../../0587-remaining-opportunity-survey.md).

Disposition: retained. `performance_claim: none`. **No production code changed.**
This packet is the raw material of a survey: eleven per-area reports returned by
parallel survey agents, verbatim, and every scratch output those reports cite.
Nothing in it is a paired measurement and nothing in it authorizes a change.

## Contents

| Path | What it is |
| --- | --- |
| `survey/<area>.md` | The eleven reports as returned: `cfb`, `xls`, `doc-ppt`, `zip-opc-read`, `opc-save`, `xlsx`, `docx-pptx`, `xml-substrate`, `xlsb`, `core-facade-parallel`, `evidence-gaps`. Each follows one format: path map, ranked opportunities with record status and evidence tier, what is not an opportunity, measurement blockers. The record is the synthesis of these; where the two differ, the record's ranking is the coordinator's judgment and the report is the agent's. |
| `cfb/` | Byte-layout census of the OLE2 corpus (`cfb_open_model.py`, `.txt`), `strace` per-open syscall counts in `file-source` mode (`strace-*.txt`, `strace-summary.txt`), the DOC/PPT fixture list. |
| `xls/` | `rustc -Zprint-type-sizes` summary for the XLS source-backed types (`type-sizes-summary.txt`). |
| `doc-ppt/` | Fresh callgrind isolation-pair annotations for the eager DOC facade open on four fixtures, the DOC source snapshot open on two, the PPT source presentation open and the PPT text-edit snapshot open (`an-*.txt`); native `perf stat` for the same operations (`perf-stat-native.txt`); the per-fixture DOC and PPT structure surveys (`survey-*.tsv`, `ppt-tree.txt`); the throwaway driver source (`docppt-survey/`) and capture scripts. |
| `zip-opc-read/` | Logged request, observation and allocation counts for a source-backed open plus one part on three fixtures (`counts-*.txt`), the top-45 callgrind symbols of open plus one part, and the probe source (`probe/`). |
| `opc-save/` | Callgrind inclusive annotations and one caller tree for the eager XLSX open-edit-save on three fixtures (`cg-*.txt`), `strace` of the real `save(path)` route (`strace-cfs-*.txt`), and member lists before and after (`cfs-*.lst`). |
| `xlsx/` | Allocator-harness reports for the source-backed and eager XLSX commit and read scenarios (`alloc-*.json`, `alloc-summary.txt`), the type-size probe and its output (`sizes/`, `sizes-output.txt`). |
| `docx-pptx/` | Callgrind inclusive annotations for `docx_semantic_one_edit_save`, `pptx_eager_batch_edit_save` and `pptx_cross_copy_media_rich` (`incl-*.txt`), one native sample of each (`native-*.json`), and the capture logs. |
| `xml-substrate/` | `EVIDENCE.md` with every number the XML section cites; callgrind self and inclusive tables for the eager and source-backed open plus one cell on the real fixture and on its marker-stripped control (`cg-*.txt`); the control fixture (`control.xlsx`) and the script that derived it (`strip_markers.py`); the probe source (`xmlprobe/`). |
| `xlsb/` | `xlsb_crud` reports on `testVarious.xlsb` and `cond_format.xlsb`. |
| `core-facade-parallel/` | The `litchi-core` micro-benchmark source and results (`microbench/`, `microbench-results.txt`), `strace` of facade and file-source opens at 1 and 11 samples (`strace-*.txt`), and the harness's own attribution JSON for the two open modes. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad after the packet was assembled. |

## Provenance

Every measurement was taken at `2fc5fc6572ac322f4aae17f88f90033bcb119f0f` on
a clean working tree, from the shared warm workspace build (`--release
--locked`) or a scratch Cargo project with path dependencies on the crates.
Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0;
valgrind 3.26.0; `perf` available. No git worktree was created and no tracked
file was modified by any survey.

## What this packet does not establish

Every number is a single leg — a count, an instruction profile, an allocator
report or a `strace` — taken warm, unpinned unless the report says otherwise,
while other survey agents were building and measuring on the same host, with
no A/A control. Callgrind prices SHA-256 in software (valgrind masks the SHA
CPUID bit) and counts `rep movsb`/`rep stosb` per byte, so hashing and bulk-copy
shares are upper bounds; native figures are quoted beside them where taken. The
XML substrate's 12.9× is measured on one real fixture; its breadth (41 of 60
fixtures) is a count, not a size. No timing, cold-cache, physical-device,
range-source, peak-RSS, concurrency or cross-platform result is claimed.

The raw `callgrind.out` files, profiler data and every scratch build tree were
deleted after the annotations were extracted; `cleanup.json` records them.
