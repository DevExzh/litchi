# Bounded correctness smoke

These one-warm-up/one-sample receipts exercise the frozen InkAction draft,
source-backed scalar, no-op, insertion, repeated-scalar-write, removal, clear, move,
opaque-payload, and caller-cap paths. They are correctness and provenance
evidence only; they are not the repeated allocator/runtime profile.

`source-inputs-before.txt` and `source-inputs-after.txt` retain the committed
base revision and absolute hashes for the five InkAction source/test inputs.
The smoke runner removes its private build target on completion and leaves
these source-bound receipts available for review.

The repeated-scalar-write lane records its final value and cost only; the
public API exposes no internal coalescing diagnostic.

Caller-cap receipts retain independent pre/post source hashes and parsed-state
checks from a fresh public no-op readback; no failed commit or rollback object
is inferred for the refusal path.
