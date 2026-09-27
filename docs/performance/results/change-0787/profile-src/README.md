# 0787 cached-part scheduling profile probe

This is an independent Callgrind probe for the `parts` route from the 0786
execution-budget baseline. It keeps the same deterministic 32-member OPC
corpus, report schema (`litchi.execution-baseline.v1`), finite execution
context, byte verification, and `--route`, `--shape`, `--workers`,
`--task-floor`, `--state`, `--samples`, `--warmup`, and `--output` options.

The profile executable is `cached-part-profile`. It accepts the common route
flag for command-line parity, then refuses `opc` and `cfb`; only
`--route parts` is a valid profile run. The normal 0786 native benchmark under
`tools/perf-execution` is unchanged and remains the source of latency evidence.

The exact profiled owner region is the public operation
`SourceBackedPackage::read_parts_ordered`, called through this probe-only
boundary:

```rust
#[inline(never)]
fn cached_part_region_0787(
    package: &SourceBackedPackage,
    uris: &[PackURI],
) -> litchi_opc::Result<litchi_opc::PartBatch>
```

The priming read calls the public method directly outside the clock. The timed
closure calls only `cached_part_region_0787`; returned `PartBatch` values stay
alive until after timing for byte verification and resource-release checks.
A post-call `black_box` return barrier prevents a tail-call boundary; its small
wrapper cost belongs to the diagnostic region. Corpus construction, package
setup, timing calls, verification, and dropping
the result are outside the Callgrind owner region. The profile build has no
source-counter feature or observer atomics, so the profiled source path is the
ordinary in-memory `ReadAt` path.

The profile driver should use one sample and no warmup per process, for
example:

```text
valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect=cached_part_profile::cached_part_region_0787 \
  --zero-before=cached_part_profile::cached_part_region_0787 \
  --dump-after=cached_part_profile::cached_part_region_0787 \
  --callgrind-out-file=profile.callgrind \
  cached-part-profile --route parts --shape large --workers 8 \
  --task-floor 0 --state primed --samples 1 --warmup 0 --output report.json
```

Callgrind instruction counts from this owner region are diagnostic evidence
for scheduling work. They are not native wall-clock measurements, do not
represent disk, network, or cold-cache behavior, and must not be combined with
the 0786 native latency samples.

The first Callgrind attempt failed before this region during rustix Linux-raw
vDSO CPU-clock initialization and collected zero instructions. The compatible
diagnostic probe returns no CPU clock value; the unchanged native benchmark
still records process CPU time. The original failed probe source is archived
in `profile-src-before-clock-fallback`.
