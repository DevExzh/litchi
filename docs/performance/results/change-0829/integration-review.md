# 0829 capture/decode integration review

Status: static review complete. I inspected the frozen `capture.py` and
`decode.py` drivers against the current `analysis.py`, `build.py`, and the
already retained quality receipt. I did not run Cargo, the probe, `perf`,
`nm`, `objdump`, a reader, Python, a formatter, or Git.

The 0829 output contracts line up at the driver boundary:

| Producer | Current output | Reader contract checked |
| --- | --- | --- |
| `build.py` | `build.v1`, two `ordinary`/`fp` artifacts, two build rows, frozen-input and source/probe descriptors | `analysis.py:606-676` accepts the build fields and explicitly permits the optional `quality` witness to be absent |
| `capture.py qualification/native` | lane-specific completion schema, ordered receipt matrix, report/RSS/log descriptors, pinned input/reference identities | `analysis.py:804-925` checks the same labels, modes, sample counts, command boundary, output oracle, and RSS receipt |
| `capture.py perf` | terminal available/unavailable completion, perf receipt(s), typed permission denial path | `analysis.py:1680-1743` branches on the terminal status; unavailable perf is required to remain typed and profile-free |
| `decode.py --symbols` | exact owner rows, numeric bounded disassembly command map, assembly/log descriptors, phase descriptors `{assembly, log}` | `analysis.py:1519-1663` checks the four owner ranges, command options/bounds, phase rows, and independent assembly bounds |
| `decode.py` | non-inline `perf script` receipts, deterministic gzip members, owner-count witness, terminal decode schema | `analysis.py:1744-1817` checks command identity, raw/frame retention, compression identity, and decoded owner evidence |

The old 0828 assumptions that would have broken this packet are absent from
the frozen driver interfaces. The owner symbols, schemas, plan base, marker,
input/reference identities, perf settings, and phase names all use 0829
values. The only retained older paths are intentional custody inputs: the
0827 quality receipts/seal and the 0828 address-bounded symbol fixture.

Two compatibility details are deliberate rather than defects. `build.py`
does not emit a top-level `quality` field; the reader explicitly accepts that
absence. The unavailable perf branch does not fabricate `reports`, `samples`,
`receipts`, or compression; the reader takes its typed-unavailable branch and
requires those profile artifacts to be absent.

No producer/reader schema blocker was found by static inspection. Actual
capture, decode, and post-capture reader outcomes remain required before this
review can establish a successful profile. In particular, the static review
does not certify that the live FP binary emits all four owner headings or that
the perf stream contains exact-owner samples.
