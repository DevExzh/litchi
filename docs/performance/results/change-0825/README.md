# 0825 — ordinary-save qualification stopped

This attempt did **not** complete the planned matched ordinary-save comparison.
It stopped on the first baseline qualification report because the frozen
`custody.py` checker expects scalar allocation counters, while the unchanged
harness emits metric objects containing per-sample `values` arrays. The child
process exited successfully; the capture wrapper failed and wrote no
qualification-complete receipt. No after build, native lane, observer lane,
comparative analysis, adoption decision, or performance claim exists.

The committed 0824 PPTX implementation is restored byte-for-byte. All other
production and harness sources, 35 normative inputs, locks, corpus inputs,
and three unrelated files retain their recorded hashes.

Completed evidence:

- Six fresh harness quality gates pass: 641 tests passed, zero failed, one
  ignored. Exact-source 0824 PPTX quality receipts are reused and identified
  separately; those tests were not rerun here.
- Three baseline release binaries built serially from the exact pre-0824
  transaction/XML archives; full-source receipts and binary identities remain.
- Fresh export contains six corpora and thirty policy outputs. Independent
  XML/OPC and ZIP preservation checks pass for all 37 artifact files.
- One DOCX lifecycle observer qualification report contains one sample. It is
  retained solely as failed-qualification evidence, without latency analysis.
- The two owned temporary roots were verified and removed; cleanup records
  the three built binary identities and removed file/byte inventories.

`reader-failures.json` binds every failed invocation and complete pre-repair
Python source snapshots. The first admission attempt passed its XML audit but
could not serialize ZIP member comment bytes to JSON. A separate retry used
hex encoding and passed preservation write/check. Frozen execution drivers
were never changed. Two offline abort-validator integration errors are also
retained: historical log descriptor shape and tuple/JSON-array comparison.
Both were repaired in the offline reader, followed by successful replay.

The frozen plan and [planned protocol](planned-protocol.md) describe intended,
**unexecuted** comparative stages. `analysis.py`, `raw_audit.py`, `validate.py`,
`custody_audit.py`, and `planned_cleanup.py` are unfinished-path preparations,
not successful validation evidence. Their numerical entry points were not run.
The authoritative reader for this aborted packet is `abort_validate.py`.

Replay from the repository root:

```sh
python3 -B docs/performance/results/change-0825/abort_validate.py --final
python3 -B docs/performance/results/change-0825/seal.py --check-head
```

Replay checks the original frozen checker still rejects the retained report,
replays independent artifact admission, verifies failure snapshots, quality,
build/source custody, absence of comparative outputs, restored production, and
cleanup. The final seal binds every retained packet file and the six report/
index files to the exact documentation-only commit. A future matched trial
must correct and exercise report-schema validation before freezing new drivers;
this attempt's one sample must not be reused as comparative timing evidence.
