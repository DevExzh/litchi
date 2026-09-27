# 0784 profile qualification review

This is a read-only qualification of the 0784 capture profile. It does not
authorize a production change or make a historical-regression or speedup
claim. The profiled executable is the feature-enabled evidence probe; the
library and workspace production sources are unchanged.

## Boundary

`capture_region_0784` is compiled only with `capture-profile`, is marked
`#[inline(never)]`, and its body contains one operation:

```text
package.opened_presentation()
```

`run_capture` calls this wrapper only in `capture` mode. The `commit` and
`lifecycle` paths continue to call `Package::opened_presentation` directly.
Package construction is before the `Instant` and Callgrind collection
boundaries. Serialization, reopening, semantic verification, report writing,
and destruction of the retained package/snapshot are after them. The wrapper
returns the snapshot, so the selected region includes the complete capture
call and its required validation, while no post-capture owner drop is
attributed to it.

The symbol inventory records the exact demangled owner
`pptx_capture_probe::capture_region_0784`; the retained assembly shows a
tail jump from that symbol into
`Package::opened_presentation_with_limits`. A tail jump does not by itself
prove a profiling boundary, so the raw Callgrind dumps are the authority for
the boundary check.

## Callgrind qualification

The retained profile driver uses one fresh process per shape and one measured
capture per process, with no warmup. Its collection controls are:

```text
--collect-atstart=no
--toggle-collect=pptx_capture_probe::capture_region_0784
--zero-before=pptx_capture_probe::capture_region_0784
--dump-after=pptx_capture_probe::capture_region_0784
```

The process is pinned to CPU 12. For every one of the six profile dumps:

- the child exits successfully and the capture JSON passes the inherited
  source, output, and semantic identities;
- `.callgrind.1` exists, has `events: Ir`, has a positive summary, and says
  `Trigger: --dump-after=pptx_capture_probe::capture_region_0784`;
- the final `.callgrind` part is the expected program-termination part with
  summary zero, and no `.callgrind.2` exists;
- the raw call graph contains one positive incoming call to the wrapper and
  the wrapper's direct child is
  `litchi_pptx::package::model::Package::opened_presentation_with_limits`;
- the selected wrapper record accounts for the complete positive dump
  summary, including the `opened::model::capture_internal` subtree.

The wrapper is a two-instruction tail-jump boundary in this optimized binary.
The raw records therefore show about two self instructions on the wrapper and
the capture body as its child. The selected totals are still the intended
operation-local inclusive region. The six positive `Ir` summaries are:

| Pass | Tiny | Medium | Large |
| --- | ---: | ---: | ---: |
| 0 | 7,915,062 | 14,291,409 | 594,662,717 |
| 1 | 7,913,175 | 14,285,406 | 594,727,095 |

These are Callgrind guest instruction references under Valgrind, not native
cycles or elapsed time. Any profile lacking the triggered part, positive
summary, single incoming wrapper call, or complete selected inclusive total
must be rejected rather than inferred from the symbol inventory.

## Native control qualification

The native lane has six alternating control/profile blocks, three shapes,
three warmups, and thirty measured samples per process: 36 processes and
1,080 measured samples. The feature-enabled profile binary is compared with
the feature-off control only to quantify wrapper/code-generation perturbation;
it is not a production performance comparison. Median profile/control p50
ratios are approximately +0.883% for tiny, +1.367% for medium, and +0.966%
for large. Source bytes, output bytes, and semantic text identities match the
inherited capture oracle for every sample. Retained spread flags remain part
of the evidence, including the tiny control RSS spread and medium p99/profile
spread; they are not silently discarded.

The qualified interpretation is therefore: Callgrind captured the intended
initial-capture region, and the native controls provide perturbation context.
The packet does not establish that initial capture explains the historical
0780 regression, and it does not establish a production speedup.

## Native sampled follow-up

The follow-up `perf` lane is a separate diagnostic: two fresh large-shape
processes, 100 measured samples after three warmups, CPU 12, `cycles:u` at
499 Hz, and `dwarf,16384` call graphs. Its frozen owner is
`litchi_pptx::package::model::Package::opened_presentation_with_limits`.

The follow-up must be qualified from the decoded stack records before any
native phase fraction is reported. Count a sample only when that exact owner
appears in its resolved user stack. Inner symbols such as
`opened::model::capture_internal`, XML readers, hashing, or compression are
whole-process observations unless the owner-qualified stack proves that they
ran under the capture call. Setup, output serialization, reopening, semantic
readback, and report handling are outside the capture timer and can otherwise
contaminate a flat `perf report`.

The wrapper's optimized tail jump makes owner loss during native unwinding a
known failure mode. If the decoded records contain zero owner-qualified
samples, retain the raw data and report the qualified count as zero; do not
turn the unqualified inner-symbol histogram into a capture-phase fraction.
The Callgrind evidence above remains valid because its triggered dump directly
proves the wrapper boundary. A future native fraction would require a
non-tail-call boundary or an equivalent explicit profiling region, followed by
the same owner-stack qualification.
