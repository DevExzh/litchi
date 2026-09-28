# 0825 — ordinary-save comparison stopped at qualification

The planned complete ordinary-save comparison of the shipped 0824 PPTX
optimization did not reach comparative capture. Its frozen report checker
expected scalar allocation counters, but the unchanged harness emits metric
objects containing per-sample arrays. The first baseline DOCX lifecycle child
process succeeded; the wrapper rejected its report. This is a checker failure,
not a library regression or a measurement of the optimization's effect.

Production is restored exactly to `032b0e89cdeb93b6a741c1712863b6687baeb6e5`.
This batch changes only documentation and retained evidence. It makes no new
adoption/rejection decision and no lifecycle, edit, tail, allocation, RSS,
throughput, or full-save speedup claim. The scoped 0824 result remains the last
completed optimization trial; its complete-save effect remains unmeasured.

## Intended comparison and actual execution

The [frozen plan](results/change-0825/plan.json) selected the same three small
real DOCX/XLSX/PPTX files used in the admitted ordinary-save baseline. Each
would run lifecycle, edit, default full-durability path publication, and
counting publication. Both source legs would differ only in the two archived
PPTX transaction/XML files. Fresh serial release binaries were pinned to CPU
12 for export and qualification; allocator/process observations were to remain
separate from native timing. Historical artifact bytes served only as output
oracles, never as latency samples.

| Stage | Actual result |
| --- | --- |
| Prior 0824 committed seal | Passed before preparation |
| Fresh harness quality | Six gates passed; 641 tests passed, 0 failed, 1 ignored |
| PPTX quality | Exact-source 0824 before/after receipts reused; not freshly rerun |
| Frozen inputs | Full source census, harness, locks, 35 normative inputs, host/toolchain, corpora, source archives, unrelated-file hashes, drivers |
| Baseline release build | Native, artifact exporter, and observer binaries built serially |
| Baseline artifact export | Six corpora, five policies each, 37 retained files |
| Baseline independent admission | XML/OPC audit and ZIP preservation passed after reporting repair |
| Baseline qualification | First DOCX lifecycle child exited 0; capture wrapper exited 1; one report/one sample retained |
| After build and qualification | Not run |
| Native and observer comparative lanes | Not run; zero comparative reports/samples |
| Numerical analysis | Not run; no ratios, confidence intervals, or decision flags |
| Restoration and cleanup | Exact committed sources restored; both owned roots removed |

The intended 216 reports/4,488 samples in the plan are **not collected results**.
The single qualification sample is excluded from comparative evidence. The
three baseline binaries are identified by retained hashes and build receipts;
they were removed with the target root after validation.

## Failures retained without rewriting the frozen trial

The first artifact admission passed its independent XML/OPC audit and then
failed while serializing ZIP member `comment` bytes into JSON. The repair
encodes comments as hex, as already done for ZIP extra data. It preserves
exact equality checks. Original logs, child receipts, and all Python source
snapshots remain; a separate admission retry passed preservation write/check.
All thirty outputs match their admitted historical byte identities, and
untouched member metadata, compressed payloads, order, and archive comments
remain preserved.

The subsequent qualification failed in frozen
[`custody.py`](results/change-0825/custody.py): `check_alloc` requires each
allocation metric itself to be an integer. The actual schema is, for example,
`allocation_calls: {"values": [...], "status": "measured", "scope": ...}`.
The failure occurs on `docx_real_file_ordinary_save_lifecycle.allocation_calls`.
The child process receipt and raw report remain intact. No qualification
complete or admission receipt exists, and the comparative lanes never ran.
The frozen checker was not repaired in place.

Two errors in the new offline abort validator are also retained: it initially
assumed historical quality log paths used the newer descriptor-object schema,
then compared Python ZIP timestamp tuples with JSON arrays. Explicit legacy
descriptor conversion and canonical JSON comparison repaired those checks.
Neither repair changed runtime source, frozen drivers, artifacts, or raw
reports. All four failed wrapper/reader invocations bind their logs and
complete pre-repair Python snapshots in
[`reader-failures.json`](results/change-0825/reader-failures.json).

## Validation and cleanup

The authoritative
[`abort_validate.py`](results/change-0825/abort_validate.py) replays independent
artifact checks, validates quality/build/source custody and all failure
snapshots, reproduces the frozen checker's rejection of the retained report,
checks its per-sample allocation schema and byte conservation independently,
and requires the absence of comparative outputs. Its final mode also verifies
restoration and cleanup. This is failure-packet validation, not acceptance of
the planned performance trial.

Cleanup removed 7,012 files and 7,453,111,620 logical bytes from only
`/home/zhuhe/code/litchi-target-0825` and
`/home/zhuhe/code/litchi-fs-0825`. The scratch ownership marker, three binary
identities, root inventories, and absence checks are retained in
[`cleanup.json`](results/change-0825/cleanup.json). The unrelated format review,
unified API design, and matrix-analysis files retain their recorded hashes.
No iWork work is included.

Replay after the documentation commit:

```sh
python3 -B docs/performance/results/change-0825/abort_validate.py --final
python3 -B docs/performance/results/change-0825/seal.py --check-head
```

The [packet README](results/change-0825/README.md) distinguishes the authoritative
abort reader from prepared but unexecuted comparative readers. The seal binds
the packet, this report, and five indexes to the exact commit.

## Next bounded step

Before freezing another matched trial, correct the capture validator to check
metric objects, per-sample vectors, lengths, statuses/scopes, and per-sample
allocation conservation. Exercise native and observer report schemas against
existing retained reports without running or reusing their timing as new
measurements. Then freeze a fresh packet and perform the full independent
admission, qualification, and matched comparison. This failed attempt cannot
establish a complete-save benefit or regression for 0824.
