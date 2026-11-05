# 0362ul30 follow-up — repaired ODP handoff

Date: 2026-10-24. This source-only follow-up repairs the three execution
blockers found in review of `40165ae0af`. The original packet remains
unchanged as historical evidence; the repaired packet is a new external
directory. No Cargo build, native probe, timing, heaptrack run, qualification,
or acceptance claim is made.

The blocker review is
`/home/zhuhe/.cache/litchi-api-contract/odp-build-handoff-review-40165ae0af.json`,
6,500 bytes, SHA-256
`09ebd4f31fd554b7a5d21fc436b10cf7238b5287a1c81d9566bbaf4207c60d3c`.
It confirmed the source, lock, adapter, probe and fixture identities and
flagged: the ignored baseline `Cargo.lock` was not copied into the clone;
memory TMPDIR was not exported; and heaptrack capture/printing assumed a
`.zst` filename and omitted the histogram output argument.

## Repaired packet

The new packet is
`/home/zhuhe/code/litchi-worktrees/.coord/odp-final-build-prep-472b670547-repair-40165ae0af`.
The prior packet at
`/home/zhuhe/code/litchi-worktrees/.coord/odp-final-build-prep-472b670547`
is preserved, including its failed-review artifacts and source-only plan
receipt. Repaired packet files are JSON-valid and hashed as follows:

| File | Bytes | SHA-256 |
|---|---:|---|
| `input.json` | 1,749 | `a7f598369ee2698a8ae265c71921f112f352e08a9bcedbd44977ee31811a6aac` |
| `plan.json` | 2,058 | `dfb286badeefbd5b036730500f7da182c5ec8b9da201dcd573cfacc5cb19a342` |
| `build-handoff.json` | 14,430 | `41012547d456a841d260952b08bf76bfcd5c22976349b6624752e0bf5e9efa62` |
| `memory-protocol.json` | 10,850 | `44103b2a571a3dfb309e0aad5deed383deb63b8d3226b10a496050031f9c31cf` |

The repaired packet also passed the current runner's source-only `--plan`
validator. Its retained plan is 2,093 bytes, SHA-256
`a975c2d5598953e702d8079269dc441b09f6c4a599888cbc13ff3a1bc193c8f6`, under
`qualification-run/manifest-plan/plan.json`.

## The three repairs

1. The baseline command verifies the ignored source lock
   `04665b965a919412e29c8772956ddf6267264fa40e3e0ef7c9967fd2eb584171`
   before cloning, copies it into the disposable clone before `cargo --locked`,
   hashes it immediately before and after the build, confirms
   `git check-ignore` and the expected clean/adapter status fences, then
   removes both the lock copy and adapter in its trap. The immutable baseline
   source is never modified.
2. Both memory legs use `export TMPDIR=...`, so heaptrack and its probe child
   inherit the fresh per-leg temporary directory.
3. Each memory leg runs `heaptrack --record-only -o heaptrack-profile`,
   discovers the actual final `.gz` or `.zst` capture at bounded depth,
   writes its path and SHA-256, and invokes
   `heaptrack_print -f <capture> -H <histogram-output> > <summary>`. The
   receipt parser retains capture, histogram and summary hashes plus
   `peak_heap_bytes` and `peak_rss_bytes`.

ODP remains `build_required` and `memory separate_run_required`; the historical
+0.5 MiB peak-heap residual is not covered by latency/correctness output.
