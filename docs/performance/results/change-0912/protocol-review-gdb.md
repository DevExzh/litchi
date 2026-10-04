# 0912 GDB fallback protocol review

**Decision: pass for dispatch reachability only.** The first native perf child
is rejected and its raw packet remains custody evidence: it contains six lost
chunks totaling 304 lost records, most callchains do not resolve, no exact
owner frame is admitted, and one x86 SHA frame is unscoped because its owner
frame is absent. The remaining perf children are stopped. No perf stack,
sample count, or phase timing from that lane may be reused.

This fallback can answer one bounded question: does the unchanged 0910 native
process enter the x86 SHA or software SHA backend while the measured
`opened_presentation_transaction` call is active? It cannot produce a native
snapshot profile, a native timing result, an Ir result, or a route comparison.
The GDB overhead makes every elapsed-time field in the child report
diagnostic only; discard those values.

## Scope and binary identity

Use the retained executable only:

```text
docs/performance/results/change-0910/binaries/native
SHA-256 f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d
```

Do not rebuild, relink, strip, patch, or copy the binary. Bind the GDB run to
the existing `symbol-ranges.json`, `symbols.txt`, and binary receipt. The
relevant file-relative ranges are:

| symbol | start | end |
| --- | ---: | ---: |
| `litchi_pptx::package::model::Package::opened_presentation_transaction` | `0x43e780` | `0x43e8f7` |
| `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration` | `0x3bd7a0` | `0x3bf55f` |
| `litchi_perf_baseline::build_pptx_source_edit_corpus` | `0x3a2c40` | `0x3a53a5` |
| `sha2::sha256::x86_sha::compress` | `0x5cb380` | `0x5cb73c` |
| `sha2::sha256::soft::unroll::compress` | `0x5cb740` | `0x5cdf0d` |

The PIE load base must be read from the running process mapping. Do not use
the file-relative addresses as runtime addresses. Record the mapping, the
resolved runtime address of every breakpoint, and the exact GDB version.

## Two guarded children

Run two fresh serial children, one route per process, with no warmup and one
measured operation:

```text
materialized 0 1 <fresh-materialized-artifacts>
owned        0 1 <fresh-owned-artifacts>
```

Pin each process to CPU 12 and keep the 0912 environment, cwd, source/output
paths, command receipt, GDB stdout/stderr, and target exit status. A GDB launch
may use `taskset --cpu-list 12 gdb -q -nx -batch -x <script> --args ...`; the
exact wrapper and script must be retained. Do not run a new build or a
profiler. The two routes provide coverage guards for the same binary; they are
not paired timing observations and must not be compared by their elapsed
fields.

Each target must finish with exit status zero and a completed benchmark report.
Require route identity, `warmup_iterations=0`, `measured_samples=1`, the
unchanged 0910 source and output identities, semantic output, sink bounds,
source immutability, and every existing report gate. A failed target or GDB
script is retained and rejects the fallback packet.

## ABI-scoped owner state machine

The owner-entry breakpoint is the scope authority. Resolve the runtime address
of
`litchi_pptx::package::model::Package::opened_presentation_transaction` from
the PIE mapping and the frozen ELF offset, then break at that exact address.
At entry, before the owner prologue changes the stack, read and log:

```text
entry_pc = $pc
entry_rsp = $rsp
return_pc = *(uintptr_t *)$rsp
```

On x86-64 SysV the word at entry `RSP` is the return address. Convert
`return_pc` back to a file-relative offset using the recorded PIE base and
classify it fail-closed:

* `timed_iteration` range means `measured_owner`;
* `build_pptx_source_edit_corpus` range means `setup_owner`; and
* every other range, an unreadable return word, a nested active owner, or a
  return address outside the retained binary means `unknown_owner` and rejects
  the fallback.

The existing 0911 raw owner captures establish these two dynamic callers, but
the GDB run must classify each entry from its own return address. Do not use
breakpoint order, route, or an assumed call-site offset as the classification.

At every owner entry, create a return guard at that exact `return_pc` and log
the owner state. Backend breakpoints remain disabled for `setup_owner`. For a
`measured_owner` entry only, enable the two exact backend addresses:

```text
sha2::sha256::x86_sha::compress
sha2::sha256::soft::unroll::compress
```

The return guard must disable both backend breakpoints, clear the active state,
and continue before any later caller instruction executes. It must fire once
for every owner entry. A missing return guard, an owner return through an
unobserved path, or a backend breakpoint still enabled after return rejects
the result. The expected successful child has one setup entry/return and one
measured entry/return.

This owner-entry/return state is required even when GDB prints a useful
backtrace. A breakpoint at `timed_iteration` alone is insufficient because it
also covers ingress, edits, commit, publication, and post-timer output work.

## Backend hits and backtraces

On the first hit of either backend, log the exact `$pc`, the resolved symbol
and file-relative offset, current owner state, entry/return sequence, and a
full GDB backtrace. The preferred evidence has both exact owner and
`timed_iteration` frames in the backtrace. Disable that backend breakpoint
after its first hit and continue so the child reaches its normal guards; do
not leave a breakpoint on every hash call.

If GDB's symbol-only unwinder supplies a complete backtrace, require the
measured hit to show the exact owner and measured parent and retain the text.
The release binary has no debug line information, so source lines are neither
required nor admissible. If the unwinder cannot show the frames, the ABI return
address plus the explicit active-state/return-guard log is the scope proof;
record the backtrace failure as a limitation and call the result
`abi_scoped_dispatch`, not a stack profile. Any disagreement between the
backtrace caller and ABI range is fail-closed.

Dispatch classification is:

* one or more x86 backend hits and no software backend hit in the measured
  owner: native x86 SHA backend observed;
* one or more software backend hits and no x86 backend hit: software backend
  observed;
* hits to both backends in the measured owner: mixed/ambiguous result, with no
  single-backend conclusion; and
* no backend hit in either measured owner: inconclusive, not evidence of the
  software path.

Backend hits while `setup_owner`, after the owner return, or in an unknown
state are protocol failures, not extra dispatch observations. The host
`sha_ni` flag and 0911 software Callgrind rows remain context only.

## Admission and retained evidence

Retain one immutable receipt per child containing the exact GDB command/script,
tool versions, environment, cwd, CPU placement, binary and symbol-range
digests, PIE mapping, owner entry/return log, backend first-hit logs,
backtraces, GDB stdout/stderr, target stdout/stderr, report and artifacts, and
the final target/GDB exit codes. Keep any failed child and do not select a
clean route subset.

Admit only if both children pass the workload guards, both have exactly one
setup and one measured owner return, all backend breakpoints are disabled
outside the measured owner, and the backend classification is unambiguous.
The final 0912 result may state native backend reachability in the guarded
owner interval. It must state that perf attribution failed, no native timing
or snapshot cost was measured, no sampled percentage or Ir conversion exists,
and no production optimization follows from this fallback.

