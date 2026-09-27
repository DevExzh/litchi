# 0778 evidence packet

Final current-source captures and preservation checks pass. See [the report](../../0778-ordinary-save-durability.md) for results, workload boundaries and limitations. A preflight-discovered DOCX note-relationship defect was corrected before the admitted matrix. Default Full durability remains unchanged.

## Offline replay

From the repository root, with Python 3 and no build required:

```sh
python3 -B docs/performance/results/change-0778/validate.py
python3 -B docs/performance/results/change-0778/tables.py --check
```

The validator replays source/build/fixture custody, owner and harness gates, exported ZIP/XML semantics, bounded negative controls, child receipts, native statistics, allocation observations, counter arithmetic and syscall policy checks. After binary cleanup, exact recorded path/size/SHA witnesses replace executable availability checks. The final `seal.json` covers every packet file except itself. Original absolute capture paths are mapped to this checkout for retained artifacts; capture provenance is retained.

`analysis.json` holds all 70 native and 70 allocation groups, including edit and counting controls. Six CSVs expose all process summaries, all native metrics/RSS, all 50 spread flags, default/Full controls, allocation metrics and both counter repeats. `tables.py --check` verifies those derived tables. Raw receipts, logs and outputs remain alongside them. `trace-analysis.json` retains all 28 measured windows and policy checks. `review.md` and `measurement-review.md` delimit independent review scope; `next-step.md` is planning only.

## Capture order and scope

The coordinator ran Cargo, native children and diagnostics serially. Capture scripts refuse to overwrite lanes:

1. `docx_quality.py` and `quality.py`: five owner gates plus nine harness/build gates at final source `0b4799d649` (2,492 passed, 33 ignored).
2. `export.py`: seven source archives and 35 policy/stream outputs in `artifacts-0`.
3. `capture.py qualification`: 28 source-bound corpus identities.
4. `admit.py`: independent ZIP/XML oracle plus source/build/qualification binding.
5. `capture.py native`, then `capture.py allocation`: 280 and 140 processes.
6. `diagnostics.py trace`, then `diagnostics.py instructions`: 28 traces and 32 counter processes plus counter qualification.

Every measured save starts with an absent destination, after four default-Full setup publications. Source data is warm; the generated corpora have fixed medium size. Weaker policies have different crash/power-loss guarantees. This is neither a default optimization nor cold, remote, concurrent, existing-destination or large-corpus evidence.

## Preserved unsuccessful attempts

`preflight/` is the inventory-bound rejected preflight archive. Its early harness run passed 556 library tests then failed a binary allocator assertion under two test threads, with two consequent poisoned-lock failures. A single-thread binary retry passed but exposed two Clippy warnings. After their correction, preflight quality passed; independent artifact review then detected lost DOCX note relationships. None of that source's outputs or timings enter the final matrix.

The final packet's `quality-0` is the fresh passing harness lane. `docx-quality-0` passed the initial production fix; `docx-quality-1` rejected an invalid plain-text XML regression fixture; `docx-quality-2` passes the final valid fixture and Strict-edge replacement coverage. Historical receipt paths in `preflight/` are preserved as provenance and bound by its inventory. No failed attempt is silently replaced.
