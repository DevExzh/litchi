# 0912 — native PPTX snapshot SHA dispatch

Both existing PPTX ingress routes reached the native x86 SHA-256 backend inside
the measured cold opened transaction on the unchanged 0910 executable. GDB
recorded the exact backend PC, the transaction and `timed_iteration` backtrace,
and matching owner-entry/return guards. This resolves the dispatch uncertainty
from [0911](0911-pptx-owning-snapshot-instruction-profile.md); it does not
establish native snapshot cost or adopt a production optimization.

## Native perf attempt rejected

The reviewed native sampling plan used CPU 12, `cycles:u`, 997 Hz, strict
frequency, and a 16,384-byte DWARF stack capture on the retained binary. The
first materialized child completed 50 guarded operations with the exact 0910
source/output, corpus, sink and semantic checks. Perf retained 16,920 sample
records, but 15,912 had empty decoded chains. It reported six lost chunks,
totaling 304 lost records. No exact transaction-owner frame was recovered.
The single x86 SHA frame lacked the required owner/caller context.

That child is rejected for scoped attribution, and the other three planned
perf children were not run. Raw compressed `perf.data`, loss diagnostics,
executable mappings, symbol ranges, command receipts and decoded streams are
retained. No clean-looking subset is promoted. The empty chains and loss are
observations, not a diagnosis of the unwinder or kernel. The successful
permission probe did not establish that the benchmark profile would qualify.

## Guarded debugger witnesses

The [fallback protocol](results/change-0912/protocol-review-gdb.md) runs one
fresh `materialized 0 1` child and one `owned 0 1` child, both pinned to CPU 12.
The binary is still 9,301,456 bytes, SHA-256
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`.
No executable was rebuilt or changed on disk.

At the exact transaction entry address, the debugger reads the x86-64 entry
stack return address and classifies its caller against the frozen ELF ranges.
A corpus-builder call is setup; a `timed_iteration` call is measured. Runtime
symbol addresses share the PIE base recorded in the process mapping. Backend
breakpoints are disabled during setup, enabled only for the measured owner,
and disabled before its caller resumes at the exact saved return address.
Each return also verifies that the stack pointer advanced by eight bytes.

Both witnesses have this sequence:

1. Setup owner entry and matched return, backend breakpoints disabled.
2. Measured owner entry, both backend breakpoints enabled.
3. First hit at `sha2::sha256::x86_sha::compress`; that breakpoint is disabled.
4. Matched measured owner return; both backend breakpoints are disabled.
5. Normal process exit and all benchmark guards pass.

The exact ELF virtual hardware entry, before adding the PIE load base, is
`0x5cb380`, in range `[0x5cb380, 0x5cb73c)`. These linked addresses are not
physical file byte offsets; the executable segment maps file bytes with a
`0x1000` virtual-address displacement. The witness field `file_address` stores
the linked ELF address. Both backend backtraces include
`package_fingerprint_with_memo`, `capture_internal`,
`opened_presentation_transaction`, and its measured caller. The fallback ABI
proof was available, but both runs also supplied the preferred full scoped
backtrace. Static disassembly of that same symbol contains `sha256rnds2`.
No software-backend breakpoint hit was recorded in either measured interval.
This is a bounded reachability observation, not proof about every native run
or the absence of a software backend in all execution environments.

All source/output bytes remain identical to 0910:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| source | 17,017,139 | `61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757` |
| output | 17,017,145 | `23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9` |

All debugger and perf-instrumented phase durations are excluded from latency
claims. There is no admitted native stack distribution, hardware-counter
total, elapsed-time share, Amdahl estimate, allocation result, or new speedup.
In particular, the 922,100,520 software-SHA guest instructions from 0911 cannot
be used to estimate native hashing cost on these witnessed hardware paths.

## Concrete next candidate

The [source audit](results/change-0912/payload-source-audit.md) identifies an
intermediate copy on the cold owning path. `DeferredPayload::decode` reads
through `IndexedArchive<Arc<Vec<u8>>>`, whose generic `member_source` can
allocate and zero-fill a temporary `MemberSpan`, then copy the already-retained
source bytes into it before ZIP decoding and verification.

A bounded candidate is a direct read path for this immutable owned source,
bypassing the speculative intermediate span while retaining the existing
local-header grammar, Store/Deflate decoder, CRC/data-descriptor verification,
read limit, decoded output allocation, part/aggregate limits, cached error
identity, and source provenance. Generic file, range and cancellable readers
must retain their current behavior. Avoid unsafe uninitialized buffers and do
not move inflation earlier merely to change phase accounting.

The source path makes this candidate concrete; neither the failed perf attempt
nor the debugger witnesses prove its benefit. The next implementation must pass
Store/Deflate, truncation, ZIP64, descriptor, CRC, limit, concurrent first-force,
source preservation and publication gates, then a fresh matched native
owning/materialized benchmark before adoption. Native cost attribution remains
open. No production source, accepted ADR, or public API changed in this batch.

The broader non-iWork goal remains active. ODF is deferred, iWork is excluded,
and integration remains pending the unrelated shared cherry-pick conflicts.

## Verification and cleanup

The [independent dispatch reduction](results/change-0912/dispatch-analysis.json)
validates exact commands, tool/binary custody, runtime mappings, entry/return
addresses and stack pointers, breakpoint transitions, scoped backend frames,
and semantic guards. Ten in-memory negative checks reject wrong caller/backend
addresses, missing or mismatched return guards, and reordered transitions.
The [final review](results/change-0912/result-review.md) records the result and
its limits.

[Offline replay](results/change-0912/replay.py) passes before and after cleanup.
It re-decodes all five retained perf exports byte-for-byte, reproduces the perf
rejection and both dispatch witnesses, and reruns the inherited 0910 source,
ZIP census and statistics checks. The 9,355-file frozen source and retained
binary remain unchanged, so their existing build, formatting and warning-denied
Clippy receipts are reused. No Rust source changed and no new build ran.

Both marked temporary capture directories were removed (130 logical bytes).
No worktree was created. Raw perf data and six generated source/output archives
remain explicit Gitignored local evidence with tracked identities; decoded
exports, GDB logs, helpers and receipts are committed. A fresh clone must
reconstruct the executable and regenerate omitted binary artifacts; the packet
does not claim they are embedded in Git.

```sh
python3 -B docs/performance/results/change-0912/replay.py \
  --output /tmp/new-0912-replay.json
```
