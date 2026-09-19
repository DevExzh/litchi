# 0681 — current evidence audit and next implementation prerequisites

Status: evidence and design work; `performance_claim: none`. No production
optimization is asserted by this record. The preceding turn made progress:
`389167b38` committed the integrated third wave, and the current checkout began
clean. OLE2 and OOXML remain active; iWork is excluded.

## Corrections checked against authoritative evidence

The 0675 integration summary described 0672's stored-cell reservation as a
warm timing regression. Recomputing all eight retained sample files confirms
that `stats.py` prints a **latency reduction**, not an increase: the before/after
p50 reductions are 3.751% and 4.134%, while the larger paired same-binary p50
reduction is 1.770%. All eight SHA-256 checks pass. The integration summary now
matches the original measurements. The scoped result remains one warm POI
worksheet, with noisy tails withheld; no new measurement or claim was added.

`CRUD_COVERAGE.md` still led with 0587's missing-selector inventory even though
later batches supplied several of those selectors. A new current-state section
links the implementations for facade DOC/PPT opens, XLS range reads, ordinary
OOXML saves, producer-shaped sources, PPTX transaction phases and XLSB structure
edit/save. The checked machine-readable coverage index is unchanged. Registry
presence does not prove a measured baseline or completion of a CRUD category.

The [evidence packet](results/change-0681/README.md) retains raw-sample hashes,
recomputed values, source registry identity and applicable gate results.

## Reviewed next steps

The parallel investigations address the first three remaining queue rows:

- [0678](0678-xls-query-cache-design.md) prices an XLS locator cache from a
  126-fixture wire census. Review identified target-specific shared-string
  validation and duplicate-cell ordering as mandatory cache-hit constraints;
  one successful query is not proof that every cell value has been decoded.
- [0679](0679-xlsx-scanner-publication-design.md) rejects early user callbacks
  behind the existing visitor contract. It identifies compact retained records
  and direct owning-result construction as compatible reductions to measure
  before considering replay with different callback semantics or explicit
  scratch. A second source read can fail after preflight, so two passes alone
  do not prove refusal-before-result.
- [0680](0680-docx-paragraph-memo-design.md) investigates the actual cost and
  ownership of repeated DOCX paragraph views, including managed retention.

These are implementation prerequisites, not three completed optimizations.
The next production batches must retain their before measurements, test the
identified refusal and budget boundaries, and keep only measured improvements
or explicitly justified enablers.

## Completion boundary

The full goal remains active. In particular, no evidence in this batch proves
bounded clean semantic caches, a reduced-retention selected XLSX scan, reusable
DOCX paragraph indexes, physical cold-cache performance, general parallel
scaling, or all required CRUD baselines. Design records establish the next
implementation and measurement steps; they are not completion evidence for
those requirements.

## Verification

All seven applicable evidence gates pass: strict and structural performance
claims, 50 claim-checker tests, report classification, CRUD coverage index,
non-iWork gate manifest and crate boundaries. Commands, revision and terminal
exit codes are retained in `results/change-0681/gates/`. The gate runner is the
existing change-0675 runner with `LITCHI_GATE_ONLY` selecting those seven gates.
The temporary gate-output directory was removed after retaining the logs.
Production test results from 0675 remain historical evidence; this batch does
not label them a fresh production test run. No production Rust file changed.

The coordinator independently reran the 0678 Python corpus probe. Both JSONL
and summary outputs are byte-identical to the retained packet; the two
temporary output files were removed. JSON parsing and targeted documentation
link checks supplement the evidence gates.

Final independent review corrected the XLS distinction between an out-of-range
SST cell-error value (which may be overwritten) and an SST read/decode failure
(which aborts). The design requires both duplicate cases. Replay cost now
includes dependency-reader work explicitly, and cache records distinguish
logical clean weight from pinned, managed and in-flight reservations.
