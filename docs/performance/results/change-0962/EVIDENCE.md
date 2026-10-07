# 0962 evidence boundaries and replay

The external root is `/home/zhuhe/litchi-performance-0962`. The committed
inventory records each retained artifact's relative filename, byte count and
SHA-256. Raw perf captures, large decoded streams and the compiled executable
stay external. Files copied here are bounded source and metadata, not a
standalone substitute for the raw evidence.

`fp-plan-v1.json` controls the corrected diagnostic captures. Build receipts
bind the retained 0957 source manifest and exact locked/offline Rust command;
`build-tools/build_diagnostic_fp_v1.py` regenerates the diagnostic executable.
The source and ordinary baseline identities are also described by change 0957.
The correct build's retained binary is independently hashed before and after
capture. Disposable target directories have been removed; their absence is
intentional and does not invalidate the retained binary descriptor.

All helpers use create-only output paths. To regenerate, provide the declared
source/toolchain/dependencies, preserve the bound input bytes and use a fresh
external directory with updated path bindings. Exact capture argv/environment
and timestamps are in each child `start.json`; use the capture helper and
frozen plan for its semantic validation. Sampling does not promise identical
raw bytes or sample weights on rerun. For exact offline replay, use the
original hash-verified raw capture and paired decoded streams. Decoder command
and terminal records distinguish actual exports from synthetic parser controls.

The baseline phase summary is a separate reduction of the existing 0961
baseline-only reports. It is not a timing measurement of the frame-pointer
build. The helper records each input report descriptor.

The initial DWARF profiles, tiny existing-binary frame-pointer probe and
unqualified build attempts are retained as diagnostic history. None supports
an accepted phase ranking. The first paired analyzer output and first review
receipt are provisional: later reductions correct labels, bind the actual
analyzer/binary, and qualify source locations. Their preservation must not be
read as final admission. Final receipt versions are identified in the change
record.

Capture/parser controls exercise synthetic corruptions of existing report
metadata. They do not add benchmark observations. CPU period weights are
sampled event weights; qualified stacks are a conservative subset of one
specific diagnostic build and do not establish elapsed phase fractions,
production timings, confidence intervals, or an optimization speedup.
