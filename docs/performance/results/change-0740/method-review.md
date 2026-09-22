# 0740 native callchain method review

status: diagnostic method review; no native attribution result and no
production change

This note records the bounded method for the unchanged 0739 PPTX
cross-slide-copy harness. It is based on the source review, the retained
qualification records, and the ordinary release binary's symbol/prologue
inspection. It does not turn a profiler observation into a removable-cost or
causal claim.

## Capture disposition

The first compressed attempt is unusable as profile evidence. Its retained
stderr reports `failed to write perf data, error: Bad address`, and replay
reports Zstd decompression corruption. The resulting `perf.data` is retained
as a failed compression control only. It must not be parsed for callchain
counts.

The uncompressed qualification records have empty `perf script` stderr and
resolve ordinary release symbols, including `prepare_cross_slide_copy_for_slides`,
`build_candidate`, `apply_plan`, and the corpus/oracle helpers. They therefore
show that the profiler can observe the binary, while the missing lifecycle
root is an unwind observability failure rather than evidence that the root was
never executing.

## Why the ordinary DWARF capture is bounded

The lifecycle function is an out-of-line symbol in the release executable.
Its disassembly first stack-probes and reserves `0xf000` bytes, then reserves
another `0xcd8` bytes, before entering the body. The planning wrapper has a
separate out-of-line symbol and a `0x5b0`-byte frame; `build_candidate` also
has a large frame. With `perf_event_max_stack=127`, the `dwarf,65528` capture
has reached the practical user-stack-byte limit available to this experiment,
but the root frame plus active callees can still exceed the captured window.

This frame-size evidence supports the stack-window explanation for the missing
root, but does not uniquely prove the cause of every failed unwind. The raw
stacks also contain unknown frames, so no claim is made that all such frames
have one mechanism. Increasing the DWARF window beyond the retained
`65528` trial is not a bounded next step.

## Diagnostic frame-pointer path

A separate binary built with
`-C force-frame-pointers=yes -C force-unwind-tables=yes` may be used only to
test callchain observability. It is not timing-equivalent to the ordinary
release binary. Its profile should use the same case arguments and oracle
projection, omit compression, and request the host's bounded frame depth:

```text
perf record -e cycles:u -F 997 --call-graph fp,127 ...
```

The raw data, script output, stderr, binary hash, source binding, and exact
invocation must be retained. If the diagnostic binary still cannot provide a
complete root-to-planner chain, the planner attribution remains unavailable;
the filter must not be widened to a generic public or dispatcher frame.

## Conservative classification

For each raw sample, sum its hardware `period` and assign it to at most one
bucket. Do not use `perf report --children` percentages, which can accumulate
children more than once. A strict planning sample must contain all of these
binary symbols in one complete chain:

```text
litchi_perf_baseline::run_pptx_cross_copy_lifecycle
litchi_pptx::opened::cross_copy_plan::plan_cross_slide_copy_for_slides
litchi_pptx::opened::cross_copy_plan::prepare_cross_slide_copy_for_slides
```

The generic public `Snapshot::plan_cross_slide_copy` symbol is optional: LTO
may inline it. The verified planning helper is out of line and its inspected
body directly calls `prepare_cross_slide_copy_for_slides`. Reject the sample
from the planning bucket if it contains either
`litchi_pptx::opened::cross_copy_plan::apply_plan` or
`litchi_pptx::package::model::Package::apply_cross_slide_copy_plan`.

Retain these mutually exclusive buckets:

* `rooted_plan_prepare`: complete root plus planning-helper plus prepare chain,
  with neither apply marker;
* `rooted_apply_excluded`: complete root plus either apply marker;
* `rooted_other`: complete root with neither the strict planning chain nor an
  apply marker; this is a mixed diagnostic bucket and includes possible
  post-clock work;
* `unrooted_setup_oracle`: no lifecycle root, including corpus construction
  and other process work;
* `ambiguous`: an unknown/truncated chain, conflicting markers, or a prepare
  frame without the verified planning caller.

The source boundary explains why the strict bucket excludes corpus setup:
`build_pptx_cross_copy_corpus` plans once before it calls
`run_pptx_cross_copy_lifecycle`. It also excludes the post-operation oracle:
the oracle runs after the lifecycle clock and does not pass through the
planning helper. Root-only totals remain unsuitable because the oracle is
still lexically inside the lifecycle function.

The only acceptable conclusion from the diagnostic profile is the sampled
cycle period observed in the strict callchain bucket for the profiled
generated corpus. It is not a fraction of ordinary release lifecycle time,
an Amdahl ceiling, a causal mechanism, or authorization for a production
optimization.
