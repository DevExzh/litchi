# 0974 preregistration v2: per-thread mmap diagnostic

Status: conditional preregistration for review. This file authorizes no Office
capture, build, source edit, or production change. It preserves
design-v1.md and the two failed 16M capability probes.

## Question and scope

The 0973 instruction-sampling pilot used period 1,000,000, a 32 KiB DWARF
stack, CPU 31, and the retained native launcher. Its first child produced
two raw PERF_RECORD_LOST records with values 4 and a matching terminal
PERF_RECORD_LOST_SAMPLES value 8. The 0974 v1 diagnostic added -m 16M,
but the harmless no-CPU-list probe failed ring setup under the observed
limits. The one -C 31 -m 16M probe also failed event setup: verbose perf
attempted both instructions:u and its system-wide dummy:u sideband event
and was denied with perf_event_paranoid=1.

The separate --per-thread -m 16M probe succeeded in creating perf.data and
logged one 16,781,312-byte mapping ((4096 + 1) * 4096), with one thread-map
entry and a CPU map of -1. It produced no sample because /usr/bin/true
was too short. This is setup evidence only; it does not establish sustained
capacity or a zero-loss native run.

Relative to 0973, the proposed capture has exactly two perf argv deltas:
-m 16M and --per-thread. It keeps period, event, stack size, outer
CPU-31 taskset, launcher, binary, fixture, warmups, measured iterations,
serial order, decoder, and ASLR control unchanged. --per-thread is a
scope change: the installed perf man page states that it uses per-thread
mmaps, automatically disables inheritance, and is ignored with -a or -C.

## Fixed capture plan

Run the twelve children serially in the existing order, only after a separate
review confirms the same-PID/single-thread precondition below. For child N,
use:

    /home/zhuhe/litchi-performance-0974/capture/instruction-profile-v1/child-N/

The perf portion is:

    /usr/bin/taskset --cpu-list 31 /usr/bin/perf record
      --per-thread
      -e instructions:u -c 1000000 -m 16M
      --call-graph dwarf,32768
      --period --timestamp --sample-cpu --buildid-all
      --no-buildid-cache --user-callchains
      -o /home/zhuhe/litchi-performance-0974/capture/instruction-profile-v1/child-N/child-N.perf.data
      -- /usr/bin/python3.14
      /home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py
      --metadata /home/zhuhe/litchi-performance-0974/capture/instruction-profile-v1/child-N/child-N.launch.json
      --requested-affinity '[31]' --
      /home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile
      --case docfloat
      --input /home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc
      --edits 1 --retention default --warmup 5 --samples 80

The actual receipt must preserve exact argv, environment (LANG=C,
LC_ALL=C, TZ=UTC, no LD_PRELOAD), cwd, output paths, return codes,
and before/after hashes for every bound input. The runner must record the
launcher PID/exec PID and verify that the launcher performs the documented
same-PID execve into the native binary.

Before accepting a child, audit the raw task records and process/thread
lifecycle. Require no PERF_RECORD_FORK, PERF_RECORD_EXIT for a descendant,
clone/task creation, or extra thread attributable to the native process;
also retain launcher metadata and bounded /proc/<pid>/task snapshots. If
the runtime creates a descendant or an additional thread, stop and mark the
per-thread capture unsupported rather than silently treating missing samples
as native work. The native process must remain the one profiled PID after
the launcher's exec.

A child passes only if setup and workload return successfully, the raw stream
has zero PERF_RECORD_LOST records and zero PERF_RECORD_LOST_SAMPLES values,
the task/thread audit passes, the existing stack/bounds checks pass, and at
least 100 marker samples are recovered with positive marker period share. The
two loss representations must be reported separately and must not be added.
If any child fails, preserve its artifacts and stop without replacement or
retry; make no all-12 conclusion. Even if all twelve pass, this remains a
profiled diagnostic and cannot support a speedup, latency, owner-fraction, or
production-admission claim.

Record command JSON, perf stderr/header, raw perf.data, decoded task/sample
records, process/thread audit, packet/reduction outputs, raw and trailer loss
counters, sample and marker counts, stack-cap/truncation diagnostics, and
file hashes. Keep raw and decoded evidence outside the repository.

## Buffer and format evidence

The host facts remain 32 online CPUs, CPU 31 allowed, RLIMIT_MEMLOCK 8,192
KiB, RLIMIT_STACK 8,192 KiB, perf_event_mlock_kb=516 KiB, and
perf_event_paranoid=1.

The v2 capability result is:

- capability-perthread-mmap16m-v1/capability-receipt-v1.json
  SHA-256 503cacc304daf95d8e7254feb1bf074d01cf707bfbeb6c11b049a857522705fd.
- Return code 0; perf.data size 44,336 bytes.
- Verbose setup line: mmap size 16781312B.
- Header: one instructions:u event, one thread-map entry, CPU map -1,
  sample_type=IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER,
  sample_stack_user=32768, and read_format=ID|LOST.
- Decoded metadata contained no sample or loss records because /usr/bin/true
  was too short. This is not a workload zero-loss gate.
- The raw decoder saw one PERF_RECORD_ID_INDEX, one
  PERF_RECORD_EVENT_UPDATE, one PERF_RECORD_THREAD_MAP, one
  PERF_RECORD_CPU_MAP, two PERF_RECORD_COMM, one
  PERF_RECORD_FINISHED_INIT, one PERF_RECORD_FINISHED_ROUND, and eleven
  PERF_RECORD_MMAP2.

The per-thread probe demonstrates that this installed perf mode can create
the requested 16 MiB data mapping with one tracked thread under the current
limits. It does not prove that a long native stream will not lose records or
that stack payloads will fit. Do not infer physical locked-memory headroom
from the virtual mapping log.

## Decoder and ASLR binding

Reuse the sealed 0973 ASLR translation proof without repeating relocation
testing solely because the capture directory changes:

- /home/zhuhe/litchi-performance-0973/receipts/aslr-translation-0973-final-v4.json
- SHA-256 89a48f57bf35cb556d890bcfd7476a3a6e27601fd030a643c6e3275384bfc74b

Before capture, verify that the existing protocol/reducer/decoder and
retained binary still match their existing bindings. A changed helper or
binary requires a fresh review.

## Bound evidence and inputs

| item | SHA-256 |
|---|---|
| GOAL/ADR revalidation receipt | fd68bf79fa20096c5542c3429a1266f8d94bd5fdb5907647681370ab8e02ec70 |
| authoritative 0973 final audit v2 | 977cee099b640a094425d8e7c3641cc71fe658d3c68209c9e8a4209a29e641ea |
| retained 0973 command JSON | 1d40d7d72eb8c8f56e0f6bf8edbcac96041253a8f816e909b48dc51d11982fac |
| retained native binary | 0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb |
| retained FloatingPictures.doc fixture | 41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d |
| retained launcher launch_baseline_v2.py | 3b5b79351883a9b4efb680f47498d8e10bdf77d131a60d11c1ecb54ad95430f5 |
| /usr/bin/perf 7.0.14 | 174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b |
| installed perf-record help capture | 7f9d5101c37496de4bd6c0af78869ae399eb36d9c51a7b78e3bcd4e75c3b3200 |
| installed perf-record man capture | 38ad97a73c5df417d171bf3ef58358011346f85d9ec89cac77428da70e81afb1 |
| 0973 protocol helper | 56b509fde01bbd66702ac8cd835da0d09d74d332935926078c6545093ed52bf4 |
| 0973 reducer | f33f14724f6264d87b2c3a0082c16bca711787140d06f3c7cfdd6e78b6f4c63b |
| retained host audit v1 | 37e204363bee72eb7d3080300dd32ce5ce21d53be86dee2c9e726b086f6901c5 |

The source-only CPU-map analysis and both earlier failure receipts remain
historical diagnostics; they do not authorize a fallback or a native run.

## Interpretation

A successful setup plus twelve passing children would show that the
per-thread mmap path is operational for this bounded cohort. It would not
prove that ring pressure caused the 0973 losses, and it would not establish a
profiler-free effect. Any raw loss, task/thread-scope violation, setup
failure, or marker/stack gate failure stops the cohort and supports no owner
fractions, phase attribution, speedup, or production claim.

No repository files, production sources, binaries, or iWork inputs are changed
by this preregistration.
