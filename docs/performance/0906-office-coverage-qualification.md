# 0906 — Office workload coverage qualification

The current harness has **534 selectors and 41 default selectors**. The
README's current registry count of 441 is corrected. A source-bound census
checks the `Case` enum, public selector spelling, CLI parser and default list
as a bijection. Replaying that census produces identical bytes. This is a
registry check, not proof that 534 distinct CRUD scenarios have been measured.

| Prefix | Selectors |
| --- | ---: |
| `doc_` | 14 |
| `xls_` | 40 |
| `ppt_` | 29 |
| `docx_` | 52 |
| `xlsx_` | 106 |
| `pptx_` | 66 |
| `xlsb_` | 5 |

These prefixes omit generic substrate cases and do not form a complete CRUD
taxonomy. A registered case may exercise a synthetic corpus, a bounded API
subset or a refusal. It does not establish real-producer fidelity, cold/range
input coverage, or an end-to-end speedup. Historical counts in dated harness
notes remain unchanged; this batch does not revalidate the default result-row
or corpus counts.

The [source context](results/change-0906/source-context.json) verifies that, before the help fix, the
goal, CRUD checklist, harness implementation and XML audit source match the
accepted Office branch commit `2b974e3541`. The shared target checkout remains
in a separate iWork cherry-pick. This batch does not resolve its conflicts or
claim current performance for its partially integrated sources. The 0905
numbers remain scoped to their frozen source inventory.

The final [selector census](results/change-0906/registry-final.json) and
[replay receipt](results/change-0906/registry-final-replay.json) are reproducible with:

```sh
python3 -B docs/performance/results/change-0906/registry.py --output /tmp/0906-registry-new.json
```

The helper refuses to replace an existing output. Its source grammar checks
fail if a variant, spelling, parser mapping or default entry is lost or
duplicated. It executes no benchmark or production API.

## Help coverage fixed

The retained 0905 executable omitted 11 registered selectors from `--help`,
including XLS, DOCX and PPTX validation. The original
[failed check](results/change-0906/help-check.json) is retained. The help string
now includes all selectors, and the existing registry test requires every
parser spelling to appear as a complete help token. Defaults and workload
implementations are unchanged.

[Focused verification](results/change-0906/help-regression-verification.json)
compiles the exact help function and registry test with Rust 1.95.0. The new
assertion rejects the old help and accepts the corrected help. This extraction
checks the source registry and actual rendered string; it is not a full Cargo
harness test. `rustfmt --check --edition 2024` passes. The initial driver
attempt omitted four range-default constants and failed to compile; the final
helper includes those exact source constants. Temporary build directories were
removed in both attempts.

## Next measurement

The [OOXML audit](results/change-0906/ooxml-opportunity.md) selects a fresh
paired PPTX multi-slide eager/source-backed capture first. Its timer includes
forward/inverse patch replay within the publication helper, so the result must
be labeled as that guarded workflow. Old PPTX timings cannot serve as its
current baseline. DOCX existing-section layout is the next unmeasured phase
baseline; there is no matched eager control for a speedup claim.

The [legacy audit](results/change-0906/legacy-coverage.md) distinguishes API
gaps from missing harness cases and historical evidence. XLS structural cell
operations and PPT slide operations expose APIs absent from the current legacy
benchmark matrix. It recommends a bounded DOC validated-render handoff
qualification on real fixtures as the subsequent legacy pilot. Neither audit
establishes fresh timing or comprehensive CRUD completion. Their hashes and
line references describe their recorded audit snapshots; the final census
binds the help-only harness revision.

No production library code, benchmark timing sample or profiler is changed or
executed by this batch. Only the focused Rust verification is compiled. All
owned temporary directories were removed, and no worktree was created. The
comprehensive OLE2/OOXML program remains active; ODF stays deferred and iWork
remains outside this program. Target integration remains pending while the
shared checkout has unrelated cherry-pick conflicts.
