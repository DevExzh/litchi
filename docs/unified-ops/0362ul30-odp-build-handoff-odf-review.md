# 0362ul30 follow-up — ODF durable receipt and ODP build handoff

Date: 2026-10-24. This follow-up binds the independent source review, the
fresh 18-test ODF receipt, and a corrected ODP build/memory handoff for final
candidate `472b6705475b073e7fe08f6e11ee17fb1e29f838`. It is source and packet
work only. It does not claim a Cargo build, native probe, timing, memory
measurement, qualification, or acceptance.

The reviewer receipt is
`/home/zhuhe/.cache/litchi-api-contract/odg-odf-source-review-472b-20261024.json`,
13,140 bytes, SHA-256
`dd621971189bd124858303023a0b50310e1023f534f909dd036d236d30271ba6`.
It found ODG and ODF source preparation ready for execution and retained the
ODP plan-only finding. The ODF shared-triad source suite was rerun with the
reviewer and recorded 18 tests:

```text
..................
----------------------------------------------------------------------
Ran 18 tests in 0.162s

OK
```

The durable log is
`/home/zhuhe/.cache/litchi-api-contract/odf-shared-triad-source-tests-472b-20261024.log`,
117 bytes, SHA-256
`218ffbc20075a32506a605d1a68d25378afecb27e5395dee65890b4b72ac5fd7`. Its
command was:

```text
/usr/bin/python3 -m unittest tools.test_prepare_odf_shared_triad_inputs tools.test_odf_shared_triad_runner tools.test_odf_shared_triad_paired_runner
```

## ODP handoff

The corrected, fresh packet is
`(deleted-coord-artifact)`.
Its files are small, JSON-valid and independently hashed:

| File | Bytes | SHA-256 |
|---|---:|---|
| `input.json` | 1,659 | `5992db0a0520ae3dedb47e44f5d584a8e1890c544d2611fca92db6482f6c47e6` |
| `plan.json` | 1,950 | `d213e7c0158757a1e9ff51c6649c2da8a5dafa69469bb73dbcc22f99114fcc5b` |
| `build-handoff.json` | 12,702 | `33943e0c7c0cce75acb104169a550506b95f1d21e3df869de277cae76898153d` |
| `memory-protocol.json` | 6,975 | `f929413cb6860604148c67c95c44088c0a23ddfe3b4a722e5461643727fa8180` |

The current runner's source-only `--plan` validator also passed. Its retained
plan is
`(deleted-coord-artifact)/plan.json`,
1,985 bytes, SHA-256
`3c8922eac7d6a40d424f5be109aef3d8298d4550575ce21c534ddf66ffb7d191`.

The candidate build is a `tools/perf-baseline` binary. Its exact expected
output is therefore
`.../targets/candidate/release/odp_native_qualification_probe`, not the old
`release/examples/odp_native_qualification_probe` path. The baseline is built
from a disposable clean clone of e4a with the pinned 47cc adapter and outputs
`.../targets/baseline/release/examples/odp_native_qualification_legacy_probe`.
The packet requires exactly two sequential release builds, each with one Cargo
job, fresh target/TMP paths, and a globally quiet heavy-process host. The
runner is explicitly not a builder; it is invoked only after both binaries
exist and their SHA-256 values are recorded.

The immutable source identities are bound in the packet: baseline e4a
`a1d3d5fb726673f8a162d4d9c4265f6ae79fcd4b`, candidate tree
`ebc1cee05f124c0c63977aff07a23288c19e7d00`, candidate root lock
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`,
perf-baseline lock `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`,
and current candidate probe SHA-256
`c68f790247c7f83a1acb256a57f6c3928346b10cb348d5cf30d8f89705673bef`.
The old final-472b input remains preserved and is not executable because it
points at the nonexistent candidate example path.

## Memory residual

The historical ODP tail-append residual is a +0.5 MiB peak-heap delta (5.3
MiB to 5.8 MiB), not a latency-only concern. The current ODP runner and probe
do not report allocator peak heap or peak RSS, so the corrected latency packet
does not close that residual. `memory-protocol.json` repairs the missing
qualification path with a separate, globally quiet, CPU-pinned heaptrack run
for the same fixture, `append` route and one-iteration arguments on both exact
binary identities. It requires retained `.zst` captures, `heaptrack_print -H`
text, parsed `peak_heap_bytes` and `peak_rss_bytes`, and a reviewed ratio.
That memory run has not been performed; ODP remains blocked pending build and
memory evidence.

No Cargo/native command was run in this follow-up, no immutable source or
historical packet was mutated, and no acceptance or performance claim is made.
