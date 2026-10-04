# 0912 native SHA dispatch and snapshot stack attribution protocol review

**Decision: pass with bounded admission gates.** The proposed native capture is
appropriate for answering whether a retained 0910 process actually reaches the
SHA-256 hardware or software backend while opening the measured presentation.
It can also retain native sampled stacks for a descriptive snapshot attribution.
It must not turn sampled cycles into elapsed time, guest Ir, or a route speedup
claim.

This review is read-only. It does not build the binary, run a workload, or run
perf/GDB. The existing perf permission probe is only a preflight: its successful
`cycles:u` run with `/usr/bin/true` permits attempting a child capture but does
not admit any benchmark result.

## Frozen binary and symbols

Use the unchanged 0910 executable at
`docs/performance/results/change-0910/binaries/native`, 9,301,456 bytes,
SHA-256
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`.
Do not rebuild, relink, strip, or copy a different release into the 0912
packet. Retain the 0910 source/build custody and bind the capture command to
this path and digest.

The current ELF has one exact demangled definition for each relevant symbol in
`nm -S -C --defined-only`; the addresses are file-relative because the binary
is PIE:

| symbol | kind | address | size |
| --- | --- | ---: | ---: |
| `litchi_pptx::package::model::Package::opened_presentation_transaction` | `T` | `0x43e780` | `0x177` |
| `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration` | `t` | `0x3bd7a0` | `0x1dbf` |
| `sha2::sha256::compress256` | `T` | `0x5cb2d0` | `0x4b` |
| `sha2::sha256::x86_sha::compress` | `t` | `0x5cb380` | `0x3bc` |
| `sha2::sha256::soft::unroll::compress` | `t` | `0x5cb740` | `0x272d` |

The local `t` backend symbols are valid breakpoint and stack-resolution
targets. Bind their exact names, addresses, sizes, and the binary digest in a
receipt before reducing any profile. A `compress256` frame alone is only the
public wrapper; it does not identify the selected backend. The retained
`sha2` 0.11.0 dispatch source contains the runtime `shani_cpuid` branch; its
retained source hash is
`cb03fddc1325b60935409f097f2fab3ae51cfab15bed1254f4e93d13b28b7eff`.
The host `sha_ni` flag receipt remains host exposure only and cannot substitute
for an observation from this native process.

The release binary has `.eh_frame` and a symbol table but no debug information.
DWARF callchain unwinding is therefore a valid attempt for symbol-level
stacks, while source-line attribution is unavailable and must not be invented.

## Primary perf capture

Run four fresh serial children in the existing ABBA order, pinned to CPU 12:

```text
materialized-a: materialized 0 50
owned-a:        owned        0 50
owned-b:        owned        0 50
materialized-b: materialized 0 50
```

The intended recorder shape is one fresh output per child, with the exact
binary and artifact arguments after `--`:

```text
taskset --cpu-list 12 perf record \
  --strict-freq -e cycles:u -F 997 \
  --call-graph dwarf,16384 \
  --output=<fresh-perf.data> -- \
  docs/performance/results/change-0910/binaries/native \
  <route> 0 50 <fresh-artifact-directory>
```

`--strict-freq` is required if available in the runner. If it is omitted, the
capture must retain perf's effective-frequency/throttling diagnostics and
reject a child whose requested rate was silently changed. Keep the exact
environment, CPU placement, cwd, perf version, taskset version, command,
binary mapping/build-id, stdout, stderr, and process exit status. Retain the
binary `perf.data` files themselves, not only `perf report` text. Also retain
the unmodified `perf script`/report reductions and any lost-sample metadata.

`cycles:u` samples user execution only. The frequency and callchain depth are
collection parameters, not measurements of elapsed time. The recorder's
instrumentation overhead and the benchmark's JSON phase clocks are outside
this protocol's native claims. The four children must remain serial and must
not share a perf output path.

Each child report must independently pass the existing 0910 workload guards:
the exact route and generated corpus, `warmup_iterations=0`,
`measured_samples=50`, source identity and immutability, stable semantic
output, sink bounds, and process exit zero. The report's phase durations are
retained as guards only. Setup corpus generation is included in whole-process
perf data but is excluded from measured owner attribution by stack class.

## Stack scope and fail-closed reduction

Reduce only callchains that contain both exact frames
`litchi_pptx::package::model::Package::opened_presentation_transaction` and
`litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`, with the owner
as a descendant of the measured parent in the dynamic chain. This is the
measured opened-snapshot interval. Do not use process order, sample order,
largest count, or the presence of the `owned` command argument as a scope
authority.

Retain a separate setup class for owner stacks whose exact caller is
`litchi_perf_baseline::build_pptx_source_edit_corpus`. Setup samples are
custody diagnostics and are excluded from every native snapshot comparison.
Owner stacks with neither exact caller, a missing/truncated parent, or an
ambiguous chain are `unclassified_owner` evidence. They must be counted and
retained; they cannot be silently folded into the measured class.

For every child, the reducer must report at least:

- total sample records, user samples, callchains present, and any lost records;
- measured-owner, setup-owner, unclassified-owner, and outside-owner counts;
- stack-unwind truncation/missing-frame diagnostics, including counts affecting
  a candidate backend sample;
- exact backend observations for `x86_sha::compress` and
  `soft::unroll::compress`, separately for measured owner, setup owner, and
  outside owner; and
- the retained top exact symbol rows and raw callchains supporting each native
  attribution statement.

The parser must resolve symbols against the retained ELF symbol ranges, not
accept a substring such as `compress` or treat `compress256` as a backend.
Any lost perf record, unresolved target IP, or target callchain that cannot be
shown to contain both owner and measured-parent frames blocks a scoped
attribution for that row. A non-target truncation may remain as a disclosed
diagnostic only when all target rows are complete and the raw data is retained.

The primary native result is observational:

* at least one exact `x86_sha::compress` sample under the measured owner proves
  that the native process reached the SHA-NI backend in that interval;
* at least one exact `soft::unroll::compress` sample under the measured owner
  proves that it reached the software backend in that interval;
* seeing both in the same measured process is a mixed/ambiguous dispatch result
  requiring investigation, not evidence to average; and
* seeing neither backend is inconclusive because sampling can miss a short
  function. It is not permission to infer the software path from 0911's
  Callgrind result or the host flag.

Sample counts and target rows may describe where native samples landed inside
the exact owner stack. They must not be converted to Ir, percentages of
elapsed time, Amdahl fractions, allocation counts, or a materialized-versus-
owned speedup. The native sample is evidence of reachability and bounded stack
context, not a timing result.

## GDB fallback

Use GDB only when the primary perf packet has no backend sample in any measured
owner stack. A perf route with no target sample remains unobserved; do not call
it software-dispatched. If GDB is needed, retain the exact GDB version,
command/script, binary maps, breakpoint log, stdout/stderr, backtraces,
register/PC symbol resolution, child report, and exit status.

A breakpoint on `timed_iteration` alone is too broad: it would leave backend
breakpoints enabled through ingress, edits, commit, publication, and output
digesting. The backend-enable gate must be the exact owner entry. At each owner
entry, inspect the symbol-only backtrace and enable backend breakpoints only
when the caller is the exact `timed_iteration`; leave them disabled for the
setup caller. At owner return, disable both backend breakpoints before
continuing. The return guard must cover every successful owner return in the
child; a normal `finish`/return-address guard is acceptable only if its log
proves that the backend breakpoints were disabled after the return.

On the first hit of each backend, record the exact PC/symbol and a backtrace
showing both the owner and `timed_iteration`, then disable that backend
breakpoint to avoid turning the guarded child into a timing experiment. Let the
child run to completion and require the same source, semantic, output, sink,
and exit guards. GDB can establish backend reachability for that guarded
native interval; it supplies no timing claim and cannot replace perf's route
wide sampled attribution.

## Admission and limits

Admit a perf child only when its raw `perf.data`, command/custody, binary
identity, workload guards, loss/truncation diagnostics, and exact stack
classification all pass. Retain failed or ambiguous children and stop the
native attribution lane rather than selecting a clean-looking subset. If the
fallback is used, mark the perf packet as inconclusive for dispatch and admit
only the GDB reachability observation that passed its owner-return and workload
guards.

The 0912 result can establish which backend was observed in the native
measured owner and which exact symbols appeared in sampled owner stacks. It
cannot establish native elapsed-time attribution, a route speedup, hardware
counter totals, cold-cache or device behavior, producer-wide support, or a
production optimization. Any later payload or fingerprint change still needs
the existing semantic, source-preservation, sink, and native route gates.

