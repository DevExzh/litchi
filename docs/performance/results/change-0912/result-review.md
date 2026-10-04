# 0912 result review

**Status: pass.** This packet admits bounded native x86 SHA-256 reachability in
the measured opened-transaction owner for both PPTX routes. It does not admit a
native profile, timing result, or route comparison. This review was read-only:
it ran no build, workload, profiler, or production change.

The final report is reviewed at SHA-256
`aa7a72545a3e1d052a60435ec54e5173b67f8b5d3b1332370af6f7d63594b00e`.

The retained perf pilot remains rejected. Its strict parser reports 16,920
sample records, 15,912 empty decoded chains, six lost chunks totaling 304 lost
records, no exact transaction-owner frame, and one x86 SHA frame without the
required owner and measured-parent scope. The raw packet, loss diagnostics and
decoded exports remain retained. No perf count or sampled period is reused as a
native cost or latency claim.

The two serial GDB fallback children both pass the workload and semantic
guards on the unchanged 0910 binary (`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`).
Each records setup and measured owner entry/return pairs, classifies the caller
from its own entry-stack return address, matches each return guard with
`RSP = entry RSP + 8`, and keeps backend breakpoints disabled outside the
measured owner. Each measured interval hits
`sha2::sha256::x86_sha::compress` at linked ELF virtual address `0x5cb380`
(the witness `file_address`; corresponding file-byte offset `0x5ca380`) within
the frozen range, with a full backtrace containing the owner and
`timed_iteration`; no software-backend hit occurs. Static `sha256rnds2`
evidence is retained as supporting identity only. The witnesses establish
reachability for these two guarded children, not universal dispatch behavior.

Ten in-memory mutations fail closed: caller return address outside the frozen
setup/measured ranges, backend PC outside its range, missing or mismatched
return guard, and reordered enable/disable events. Canonical JSON and raw logs
were not modified. The independent dispatch reduction and mutation receipt
bind the commands, environment, PIE mapping, binary, symbol ranges, event
order, return guards, breakpoint state, and semantic checks.

Replay passes byte-for-byte before and after cleanup. It replays the five perf
exports, the rejected perf analysis, both GDB witness reductions, and the
reused 9,355-file native custody checks without starting a build, workload, or
profiler. Cleanup removed two marked temporary roots totaling 130 logical
bytes; no worktree or build tree was created, and retained raw perf data and
generated archives remain explicitly local evidence.

The result is limited to the generated fixture, the unchanged 0910 executable,
and the two guarded GDB children. It provides no native snapshot cost,
elapsed-time share, hardware-counter total, sampled percentage, guest-to-native
conversion, allocation/RSS result, speedup, or production optimization. The
source audit's direct owned-source read remains a candidate requiring a fresh
end-to-end measurement.

The report and reviewed evidence are bound in `result-review.json` for the
final seal.
