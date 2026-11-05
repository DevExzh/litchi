# 0362ul16e — W4b/W5 build slot refused by the quiet guard

The authorized slot was configured for sequential 20 GB/job-1/offline/locked
builds using the reviewed 472b W4b/W5 packet. The first W4b build-phase
invocation was made with fresh external output and target paths. The runner
refused before spawning Cargo because unrelated Litchi Cargo/rustc processes
were active. The reduced result is
[results/0362ul16e/build-guard-refusal.json](results/0362ul16e/build-guard-refusal.json).

## Actual invocation and result

```text
python3 tools/native_qualification_runner.py \
  --manifest tools/native_qualification/manifests/w4b-xlsb.json \
  --phase build
```

Result: exit `2`; runner receipt status `failed`; `event_count=0`; no Cargo
child, build artifact, or target directory was created. The receipt is
external at
`/home/zhuhe/code/litchi-worktrees/.coord/native-w4b-w5-final-472b/run-w4b-build/build-receipt.json`
(17,167 bytes, SHA-256
`a2ede68b1779aca12e3ffa634e690cc40d19bb3fc0420d43b38942aaa77dcc91`). The
captured stderr is 12,135 bytes, SHA-256
`b50f3d35999c8c1a805fe831a1b10cbf62ccd9ffb77c96813a83f738bf960a15`; stdout
is empty. The manifest and runner hashes in the receipt match the reviewed
packet.

The guard observed active jobs including PID 700738 (`cargo build` in the
472b production source for the PPT lazy lane), PID 792388 (`cargo build` in
the IWA candidate), PID 794986 (`cargo check` in full-gate-9781), and child
PID 796580 (`rustc` compiling `litchi_ppt`). Their exact command lines and
paths are retained in the external stderr/receipt and reduced JSON.

## Scope boundary

The W5 build, correctness analyzers/preflights, executable probes and quiet
ABBA were not attempted after the refusal. The runner’s `require_quiet()` is
called before every build child as well as measurement children; no bypass or
environment override was used. A later attempt needs a genuinely quiet host,
fresh output/target directories, and the same source/lock/probe fences.
