# 0718 DOCX serialization allocation and mapping attribution

This unchanged-source packet follows 0717's phase fault/timing association.
One newly built procfs-enabled release executable runs eight fixed children:
two corpora × two repeats × strace/DHAT. Source and support files remain exactly
0717's; no production candidate or native speedup is claimed.

The strace lane retains 100 warmups and 200 measured samples per child. Procfs
`io`/`stat`/`status` opens form 32 control pairs followed by 300 iteration pairs.
A conservative before-status-open/after-io-open bracket is aligned with each
raw acquisition-order sample delta. Attribution additionally requires the
serialization owner and run_case stack. The bracket includes probe tail work
and is not the exact timer interval. Other stacks and windows remain distinct.

DHAT runs one measured invocation without warmup. Its replacement allocator
changes mapping behavior; selected owner/caller contexts provide cumulative
allocation observations, not native page faults, RSS, complete operation
allocation coverage or additive peak measurements.

Strace's blank local Rust names are resolved offline through the captured
binary's executable ELF segment. `symbolization.json`, program headers, symbol
table and address-lookup output preserve all 188 local offset bindings. Each name
must match a containing symbol range. The trace records file offsets, so the
ELF file-offset/virtual-address distinction is required. No child is rerun.

`brk-analysis.json` separately validates requested/returned breaks and derives
within-process growth/shrink transitions from the same retained traces.

`analysis.json` validates every receipt, command, fixture and raw artifact,
reuses strict 0717/0709 phase/metric/parity validation, and retains stack groups,
marker windows and fault buckets. All raw profiles remain available. Prior
0717 harness checks and 0713 production verification are explicitly reused,
including ignored tests and the seven baseline-proven PPTX/XLSB test Clippy
exceptions. The executable build, profiler captures, corruption checks, replay,
final report gate and cleanup audit are fresh.

Replay after scratch cleanup requires no executable or profiler:

```sh
python3 -B docs/performance/results/change-0718/audit.py
python3 -B docs/performance/results/change-0718/artifact-seal.py --check
```

The protocol does not establish a faulting instruction, native timing fraction,
allocator regression, cold-cache behavior or speedup. Prior 0715 and 0618
rejections remain unchanged; all fixed children are retained.
