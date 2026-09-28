# Change 0807: current PPTX capture attribution

This packet is a completed, evidence-only follow-up to the rejected 0806
candidate. It asks whether the current production source has a stable CPU
attribution for the known PPTX capture boundary after the 0792/0800 changes.
It does not propose or adopt a production optimization, establish a speedup,
add CRUD coverage, or pool timing with an earlier packet.

The probe source, Cargo template, and lockfile are byte-for-byte reused from
0791. The native lane has six alternating control/profile blocks, three
shapes, thirty measured samples, and three warmups: 36 reports and 1,080
samples. The scoped Callgrind lane has six reports and six retained dumps. The
frame-pointer lane has two sampled perf runs with 100 probe samples each. The
three lanes therefore retain 44 reports and 1,286 measured outputs.

The owner and static marker remain `pptx_capture_probe::capture_region_0784`
and `litchi-perf-0780-static-mce-capabilities`. The frame-pointer diagnostic
continues to qualify the opened-presentation owner separately. Nested stack
counts remain inclusive diagnostics; they do not authorize phase fractions or
performance claims.

The packet compares the current source file manifest at the recorded 0807
origin revision with the file manifest captured by
`change-0806/build-before/source.json`. The file manifests must match while
the revision fields remain distinct. The 0784 Callgrind and perf readers are
loaded only through their recorded hashes, and the 0780/0784/0785 sealed
fixtures remain the semantic oracles.

Root execution is deliberately serial and staged:

```sh
python3 -B docs/performance/results/change-0807/build.py
python3 -B docs/performance/results/change-0807/build_fp.py
python3 -B docs/performance/results/change-0807/quality.py
python3 -B docs/performance/results/change-0807/capture.py
python3 -B docs/performance/results/change-0807/profile.py
python3 -B docs/performance/results/change-0807/perf_fp_capture.py
python3 -B docs/performance/results/change-0807/perf_decode.py perf-fp
python3 -B docs/performance/results/change-0807/perf_frames.py
```

The build and capture drivers are provenance drivers and refuse existing
outputs. They must be run only after the root agent reviews and freezes this
packet. Offline readers and the independent frame/scan cross-checks are
replay-only; they never execute a workload:

```sh
python3 -B docs/performance/results/change-0807/native_analysis.py --write
python3 -B docs/performance/results/change-0807/profile_analysis.py --write
python3 -B docs/performance/results/change-0807/perf_analysis.py --write
python3 -B docs/performance/results/change-0807/root_frames.py
python3 -B docs/performance/results/change-0807/root_scan_costs.py
```

The owned target is `/home/zhuhe/code/litchi-target-0807`. Cleanup must retain
binary identities before removing that target, and the final validator must
replay every retained receipt after cleanup. The final report and five indexes are sealed alongside the retained results.


All three lanes and independent replays pass. Notes scanning remains the main
observed capture path; namespace-aware event processing is the leading native
leaf (189/190 exact-owner samples). The ordinary/wrapper latency controls show
code-generation perturbation, so profile timings are not production speedups.
Each native repeat retains one unresolved interior frame; no phase fraction is
claimed. See the [report](../../0807-pptx-current-capture-profile.md).

The post-capture exploratory `event_assembly.py` retains the exact frame-pointer
symbol and disassembly; `event_sample_offsets.py --check` replays sampled leaf
IP offsets from retained frames without the removed executable. This follow-up
was not a predeclared causal experiment. It motivates a semantic review of event
handoff, not an adoption or instruction-cost claim. `event-tools.json` records
the disassembly tools after that extraction.

After all replay checks pass, `cleanup.py` verifies and removes the three copied
binaries and owned target. `seal_packet.py --write` seals payloads and documents;
`--check-index` and `--check-head` audit every staged or committed blob. The final
validator requires the cleanup witness and replays the original lanes plus the
exploratory offset census. Production source is unchanged and the broad goal
remains open.
