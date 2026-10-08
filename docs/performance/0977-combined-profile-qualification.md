# 0977 — combined profile qualification

**Status: failed qualification; child-02 variable-sample rejection is diagnosed.**
0977 is a fresh qualification packet for the retained 0975 ordinary binary
and `FloatingPictures` fixture. It is not a retry of 0975. The packet stopped
before an all-12 result, so no timing, performance, adoption, or
source-attribution claim follows; no production source or runtime code was
changed.

## Scope

The packet requests one serial cohort of twelve children for the fixed cell
`docfloat/e1/default`, with five warmups and 80 measured calls per child. The
outer recorder pins the target to CPU 31 and uses `--per-thread`,
`instructions:u`, period 1,000,000, a 16 MiB perf buffer, and 32 KiB DWARF
callchain storage. Startup and handoff samples remain in the capture, so the
packet is evaluated as a complete child lifetime rather than only the timed
calls.

The fresh 0977 bridge adapts the sealed 0976 handoff to this target. It checks
the exact parent perf process, moves every observed parent perf TID to CPU 30,
verifies the unchanged `(tid,starttime)` set, and same-PID `execve`s the
retained 0970 launcher. That launcher then same-PID execs the bound Office
binary. Bridge masks and bounded observations are recorded, while the raw
zero-loss task stream remains the lifetime authority. The Office exec chain
does not include the separate stdlib-burn control.

## Preflight requalification and controls

Preparation found that the installed `/usr/bin/perf` changed after the sealed
0976 environment. The current tool is `perf version 7.0.14`, 25,192,472 bytes,
SHA-256
`1371ae5cb9fa59632d1d0e34b877d7251df4029810ba67dc56e946545e87b531`;
the sealed 0976 tool was 25,221,960 bytes with SHA-256
`174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b`.
The package log records `linux-perf:amd64 7.0.0-34.34 -> 7.0.0-38.38` at
2026-10-08 06:46:44 UTC. Python, `taskset`, libdw, and libelf remained the
same. A final-v2 offline requalification then bound the current perf identity
and exact decoder SO and passed the four required current-perf control modes.
It establishes offline compatibility only; it does not retroactively change
the historical capture identity or establish native performance equivalence.
The frozen plan v2 now binds the current perf, helpers, decoder, and gates.

The separate harmless control completed 4,633 samples with raw and terminal
loss zero, 14 active-target observations, and 16 post-handoff observations.
It exercises the control boundary separately from the Office capture and is
not Office qualification evidence. The corrected integration-v9 terminal
binding check and the four-control suite pass, but integration-v9 is a
synthetic duplicate-packet custody check rather than a native cohort. Office
capture was launched under root GO. The package drift is environment evidence,
not a source-cause finding.

## Per-child qualification gate

A child is admissible only when all of these checks pass:

- native return code is zero and the exact report/oracle checks pass;
- the raw lost-record sum is zero and terminal `LOST_SAMPLES` is zero;
- one target PID/TID is observed on CPU 31 with the exact requested period;
- marker-period share is positive and at least 100 marker samples are present;
- the decoded packet has zero empty callchains.

A failed child preserves its raw files and stops the serial cohort. There are
no retries and no selective replacement children. A passing child alone does
not qualify the twelve-child packet.

## Office capture result

The first child passed the declared gate: native return code 0, 946 marker
samples, zero empty callchains, raw loss zero, and terminal loss zero. The
second native child also returned code 0 and reported terminal
`LOST_SAMPLES=0`, but the retained raw `STACK_USER` scanner rejected its
stream before corrected decoding and packet creation. No child-02 packet or
corrected decode exists.

The runner stopped with return code 2 after child 02. Children 03–12 were not
started, no retry or replacement child was used, and the all-12 packet was not
admitted. Offline diagnosis found 7,201 sample records: 7,200 records are
33,008 bytes and one record at sample 78 is 72 bytes. That variable record has
`callchain_nr=0`, `REGS_USER` ABI 0 (`NONE`), `STACK_USER` size 0, `misc=1`, a
kernel-shaped IP, PID/TID 3849325, CPU 31, and period 1,000,000. Its bounds and
period are valid; it is a valid variable-layout sample, not record corruption
or a bounds violation.

The fixed-offset raw scanner requires `MIN_HEADER=240` and reads `STACK_USER`
size and bytes at offsets 224 and 232, so it labels this 72-byte sample
`sample_fields_truncated`. The frozen decoder therefore correctly rejects the
child without producing corrected attribution. The final diagnostic v3
metadata binding confirms this bounded explanation; it does not provide a
child-02 decode or revive all-12 admission. Any future repair must preregister
explicit variable-layout handling and empty/kernel/startup sample semantics;
this record cannot retroactively ignore the sample. These results provide no
timing, speedup, source-owner, or adoption evidence.

A narrow quiescence caveat remains: a stale broad read-only grep traversed
`/home/zhuhe` for about 32 minutes during the historical capture. No write or
perf handle was observed, but read-side contention cannot be ruled out, so this
record makes no idle-host or contention conclusion.

## Offline reduction and custody

For an admitted child, the raw data, ELF/MMAP layout, decoder source/map/shared
object, and target binary are bound after terminal capture. The corrected final-v2
DWFL decoder from 0976 is used only by the offline decoder invocation;
`LD_PRELOAD` is absent from the native Office process. Stock and additional
decoder modes are outside this preregistered reduction path.

The custody chain is acyclic: the helper seal binds the helpers, the plan
binds that seal and its external reviews, root GO binds the exact plan, and
the capture manifest binds the plan and GO. Preparation and self-tests must
not start a build or workload, and no plan or output path may be reused. Large
raw native/decoder artifacts and compiled binaries remain external to the
repository.

## Disposition

The Office packet failed closed at child 02's decoder admission step. The raw
child artifacts and terminal records are retained with the bounded diagnostic
evidence; the incomplete packet cannot support phase attribution, a causal
owner, speedup, or adoption. Any later combined decoder/affinity change
requires its own reviewed evidence and correctness gates.
The sealed reducer was invoked once against the failed manifest. It returned
code 2 and emitted no all-12 qualification reduction.

Design source: external `/home/zhuhe/litchi-performance-0977/design-0977-v1.md`.

## Evidence and review state

| Evidence | SHA-256 |
| --- | --- |
| [child-02 raw-stack diagnostic v3](results/change-0977/receipts/child02-rawstack-diagnostic-0977-v3.json) | `dbfb076ec8010d82bb81125569dcc94907b8fcadd67035572a4e4b5532631a1b` |
| [failed-cohort independent review v3](results/change-0977/analysis/cpu-review-office-failed-cohort-0977-v3.json) | `050b5c6006122ce5ab5fb780971069196bcc03cc7a32ced93f262231d50c1e3d` |
| [temporary-cleanup supplement](results/change-0977/receipts/temporary-cleanup-0977-supplement-v1.json) | `0f8e9785bd83ad5e00fdcf10dea3b2acb0f0a693a4fe43c7fc079c5000a1667c` |

The final review confirms the bounded scanner diagnosis and failed-no-retry
disposition. It does not admit a corrected child-02 packet, an all-12 cohort,
or any performance or source-attribution claim.
