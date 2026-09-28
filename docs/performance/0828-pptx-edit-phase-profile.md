# 0828 — PPTX phase profile stopped at symbol admission

The planned current-source PPTX phase profile produced **zero measurement
reports and zero timed samples**. Five fresh probe quality gates and both
release builds passed, but the frozen compiled-symbol admission script failed
before qualification. Production and benchmark runtime source are unchanged;
[0827](0827-ordinary-save-compaction-effect.md) remains the latest admitted
ordinary-save measurement. No new latency, CPU share or speedup is claimed.

## Failure and diagnosis

`decode.py --symbols` invoked `objdump --demangle=rust` with a mangled
`--disassemble=<symbol>` selector. The resulting output contained only ELF and
section headings, with no function instructions. The script correctly refused
to admit this output, but its selector was wrong. The command exited 1 at the
assertion checking for the selected mangled symbol in the disassembly.

The failed driver, its frozen hash, console traceback and all four partial
symbol artifacts remain in the packet. The driver was not edited after freeze.
No qualification, native or perf directory was created. The planned 23 reports
and 4,549 samples are unexecuted; the passing parity unit test is quality
verification, not a timing report.

A separately recorded post-abort static diagnostic used exact raw/demangled
`nm` rows to obtain each wrapper's address and size, then passed a bounded
address interval to `objdump`. It verified the exact demangled owner heading,
instruction bounds, frame-pointer setup and a call instruction for all four
wrappers: whole edit, initial transaction capture, text replacement, and
commit/application. All six diagnostic commands passed. This establishes a
concrete correction for the next driver; it does not replace the failed frozen
admission or authorize measurements from this batch.

## Inputs and verification

The base is `c990602492106d968898b7310bdd4bc17f3e4fcb`. Custody checks cover
9,197 production files, 87 performance-harness files, 35 normative inputs,
root/tool/probe locks and the three unrelated workspace files. The committed
0827 seal is checked against its pinned Git blobs, including historical quality
receipts; current source and lock identities match the reused evidence.

Fresh probe gates passed: formatting, release check, release tests,
warning-denied Clippy, and warning-denied rustdoc. The three tests cover CLI
bounds, exact direct/wrapped byte-and-semantic parity, and rejection of wrong
input/reference identities. Reused 0827 harness evidence records six gates,
641 passing tests, zero failures and one ignored test across 28 summaries.
Its nested PPTX quality evidence remains historical; none is described as a
fresh production test run here.

The input is the 68,822-byte checked-in `test-data/ooxml/pptx/shapes.pptx`, SHA-256
`19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`.
The independently pinned 0821 output oracle is 68,284 bytes, SHA-256
`38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`.
Archive metadata declares 48 members, 154,250 logical bytes and 60,368 compressed
payload bytes. These are archive-directory values, not measured read, inflate,
copy or allocation counters.

The AMD EPYC 9R45 host and toolchain are recorded anew. Both probe builds use
release opt-level 3, debug level 1, thin LTO, one codegen unit, two build jobs,
unwind panics and no incremental compilation. The second adds
`-C force-frame-pointers=yes`. CPU 12 was reserved in the plan for serial capture;
no capture command reached that stage.

## Retention, cleanup and next step

Every root reader attempt retains its source snapshot, command, console and
exit code. Pre-freeze and post-quality preflights passed. Offline abort
validation checks the original failure, unchanged frozen inputs, both build
receipts, the separate static diagnostic, zero report counts and cleanup.
The sole execution failure is the original compiled-symbol gate.

The owned target was removed after verifying both binaries: 3,688 files and
2,054,144,602 logical bytes. No filesystem scratch was used. The unrelated
workspace changes remain untouched.

A fresh batch must freeze the corrected address-bounded symbol check, rebuild,
and run the complete qualification/native/perf matrix with its direct, wrapped
and frame-pointer controls. Current phase costs remain unknown. The source-only
[opportunity review](results/change-0828/opportunity-review.md) records duplicate
PPTX catalog parsing and XLSX semantic readback as hypotheses, not adopted
optimizations. Complete package fingerprinting, preservation validation and
default durability remain required.

Offline replay after cleanup and commit:

```sh
python3 -B docs/performance/results/change-0828/abort_validate.py --final
python3 -B docs/performance/results/change-0828/seal.py --check-head
```

Evidence: [packet](results/change-0828/README.md),
[abort receipt](results/change-0828/abort.json),
[static diagnostic](results/change-0828/symbol-diagnostic/result.json), and
[cleanup](results/change-0828/cleanup.json).
