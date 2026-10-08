# 0975 bounded next-capture feasibility

Observed: 2026-10-08T05:40:35Z
Status: read-only design note; no probe, build, native workload, capture, or
production/helper edit was performed for this note.

## Evidence

The retained 0975 run stopped at the first invalid child and was not retried.

- Root terminal receipt:
  /home/zhuhe/litchi-performance-0975/receipts/root-terminal-0975-v1.json
  SHA-256
  57b2ac380aba4e23fc24638e30a63a326f6a935306cfc3c1c70b7bafc07ecfbe.
  It records child 1 passing, child 2 raw loss aggregate 184, children 3--12
  not launched, and no_retry=true.
- Runner log:
  /home/zhuhe/litchi-performance-0975/runner-0975-v1.log
  SHA-256
  79ded9e944305d9abe9a53bf8e29a8d9d7cf97594f71ccf2a2059dffc78b5983.
  Child 1 has 180 marker samples and zero raw loss; child 2 failed on raw
  loss aggregate 184.
- Child 2 raw stack scan:
  /home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-02/child-02.raw-stack.json
  SHA-256
  474083e6891eaf37569e805cbc370eb85a886b9668574a7e6ac379966009ce11.
  It records 23 PERF_RECORD_LOST records summing to 184. Its terminal
  receipt separately reports 184 PERF_RECORD_LOST_SAMPLES; those fields are
  separate protocol values and must never be added.

The actual child command is retained at
/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.command.json.
Its SHA-256 is
1999ff22736b32752ed315f4287252115b76c00f4cc437840463e7e45a3784cd.
The relevant prefix is:

    /usr/bin/taskset --cpu-list 31 /usr/bin/perf record --per-thread
    -e instructions:u -c 1000000 -m 16M --call-graph dwarf,32768 ...

The retained v5 runner source used by the captures is
/home/zhuhe/litchi-performance-0973/tools-v2/run_0973_v2.py and has SHA-256
f4f7af4ec8b8bb1818f3ce25dec046c8b8ad7c23d0a2611c76e24bf10b6aef5d.
Its run_child launches the command with preexec_fn=pin_cpu; pin_cpu calls
sched_setaffinity(0, {31}). The command then applies taskset --cpu-list 31
to the outer perf process. This binds both the recorder process and its
inherited target to CPU31.

The retained launcher
/home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py has
SHA-256
3b5b79351883a9b4efb680f47498d8e10bdf77d131a60d11c1ecb54ad95430f5.
(It is also bound in the retained plan.) It compares
sched_getaffinity(0) with the requested [31], writes the observation, and
execs the native binary; it does not set an affinity mask itself.

Installed /usr/bin/perf is version 7.0.14 and has SHA-256
174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b.
Local perf record help and the installed perf-record manual say:

- --per-thread uses per-thread mmaps and automatically disables inheritance.
- -m selects the data mmap size.
- --affinity=node|cpu sets the trace-reading thread affinity to the
  NUMA-node mask or the CPU of the mmap buffer being processed.
- -C selects monitored CPUs and creates sideband handling for all CPUs; it
  is not a dedicated recorder-placement option and was already excluded by
  the per-thread protocol.

Therefore perf record --affinity=cpu is not a reliable CPU30 offload here.
The installed help only promises the CPU of the mmap buffer being processed;
it does not promise a dedicated CPU30 reader. The retained per-thread layout
records CPU_MAP=-1, and the actual reader placement for this mode is not
verified by the available evidence. --affinity=node likewise does not select
a dedicated spare CPU. No claim is made that either policy pins the reader to
CPU31, nor about the kernel PMU overflow work or host scheduler load.

## Smallest discriminating future design

The direct hypothesis is that user-space perf consumption and file/ring
draining on CPU31 can contend with the pinned target. This is a hypothesis
from the placement, not a demonstrated cause: child 1 passed and child 2
lost records under the same protocol, so the observed run does not isolate
CPU contention.

A future isolated runner can test it while retaining the native target on
CPU31:

1. Pin only the outer perf recorder process to CPU30, which is in the
   observed allowed set 0-31. Do this explicitly in the new runner's
   Popen/outer command and bind the recorder PID's
   /proc/<pid>/status Cpus_allowed_list as evidence.
2. Add a target-side affinity setter in the command topology before the
   existing Python launcher, for example a reviewed same-PID taskset
   --cpu-list 31/exec wrapper around /usr/bin/python3.14. The existing
   launcher then continues to verify [31], and the native exec inherits
   [31]. A plain taskset --cpu-list 30 perf ... -- python ... is invalid
   for this purpose because the target would inherit CPU30 and the launcher
   would reject it.
3. Keep the native binary, fixture, native argv, instructions:u period
   1,000,000, --per-thread, -m 16M, DWARF 32,768, serial order, and
   twelve-child design otherwise fixed. This is a new runner/command
   topology and needs its own source/argv/hash bindings; no current helper
   is silently relabeled.
4. Require both affinity observations: recorder PID CPU30 and launcher/native
   metadata CPU31. Require every decoded sample CPU31, the one-PID/one-TID
   task stream, no unexpected FORK/CLONE/descendant EXIT, and the existing
   zero-loss gates. Preserve raw PERF_RECORD_LOST and terminal
   PERF_RECORD_LOST_SAMPLES as separate zero checks, require at least 100
   marker samples per child and positive marker-period coverage, stop on
   first failure, and do not retry/selectively rerun.

This tests recorder placement without claiming that CPU30 is idle or that
all recording overhead moved off CPU31. Kernel event generation and PMU
interrupt work can still occur on the target CPU, so even a loss-free result
would only show that the placement change is a useful diagnostic comparison.

## Alternatives and limits

Lowering the instruction sampling rate is a secondary mitigation, not the
first discriminating test. It would reduce sample/callchain production but
would change the observed sampling distribution and may fall below the
100-marker gate; child 1 produced only 180 marker samples at period 1M.
Any period change needs a separately sealed design and unchanged loss and
scope gates.

Making each target run longer at the same period does not directly reduce
recorder load. It increases the number of generated records and can increase
ring pressure; it also changes the workload observation window. It is not a
loss-recovery plan. A longer run would need an explicit workload/marker
contract and must not be combined silently with the CPU-placement test.

The proposed CPU split is consequently a bounded feasibility candidate, not
a qualification retry, performance result, or causal attribution.

## v2 correction: target-startup CPU ordering is unresolved

This v2 note preserves v1 unchanged and adds a required admission risk. In the
suggested CPU split, an inner target-side `taskset --cpu-list 31` would itself
start under the outer recorder's CPU30 affinity. The current perf header has
`enable_on_exec = 1`; therefore the taskset process can execute and produce
samples on CPU30 before it applies CPU31 and execs Python. The wrapper is short,
but its length does not prove that no sample can be recorded in that interval.
The existing parser requires every raw sample CPU to equal 31, so even one such
startup sample must fail the gate. The target command must also account for the
extra taskset COMM/exec transition while preserving the same-PID requirement.

This is a protocol-design issue, not a reason to reinterpret a future capture.
A future CPU-split design must either establish target CPU31 before event
measurement can begin, or define and independently validate a startup exclusion
that preserves the declared sample scope. It must record the observed recorder
and target masks and retain the all-sample CPU31 gate. No capability claim,
startup ordering guarantee, or loss explanation is made here; solving this
ordering issue requires a separately reviewed design and, if needed, an
explicitly authorized harmless control.
