# 0362umfc — final472b OPC read-ahead timing packet

This is a source-only execution packet.  It binds the accepted fixture-rerun
build/correctness receipt to the current fail-closed timing guard and a fresh,
future output root.  It does not run Cargo, native correctness, ABBA, or the
superseded all-in-one shell recipe.

## Packet identity

```text
packet: docs/unified-ops/results/0362umfc/opc_read_ahead_final472b_timing_packet.json
bytes: 9655
sha256: bc2d2c6b1911ce22a86536206f61ada6d5ee364aade2a4e2a653f18a19ec03d5
guard source commit: 2732b45b6126045ca603838e888d7e42fa8810f5
guard sha256: db48dde4633f98d9ae9f39706c772af79462abf4d287f09264bfd0240435e4ae
```

The accepted input receipt is the preserved external file:

```text
/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/
opc-read-ahead-final472b-prep/prefix-diagnostic-receipt-fixture-rerun.json
bytes: 8851
sha256: 1559c13d0a4cd05eef6b3defaab304a3ecfa55b4b504ce8b2d15d3ea20f2f6e3
```

It records both release binaries and all four one-sample correctness reports
as successful diagnostics, with no timing or acceptance claim.  The packet
also records the baseline/candidate revisions, trees, clean source roots,
Cargo.lock hashes, probe and fixture hashes, selected CPU-29 topology, 20-GB
address-space/target limits, and the preserved failed and corrected run roots.

## Exact execution boundary

The JSON packet’s `command_shell` is the executable command.  It uses the
accepted binaries from the corrected fixture-rerun root, `--samples 30`,
`--warmup 3`, `--poll-seconds 0.1`, zero transfer delay, separate forward and
exact policies, and the fixed `a1,b1,b2,a2` order.  Its output and temporary
paths were absent at packet creation:

```text
/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/
opc-read-ahead-final472b-timing-20261024T210504Z
/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/
opc-read-ahead-final472b-timing-tmp-20261024T210504Z
```

The quiet-host fence must be rechecked immediately before launch and during
each leg.  CPU 29 is a recorded selection, not an execution grant.  Do not
overwrite either preserved run root or invoke
`run-opc-read-ahead-final472b.sh`.
