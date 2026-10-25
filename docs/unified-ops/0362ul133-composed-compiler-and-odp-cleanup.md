# 0362ul133 — composed compiler qualification and ODP consolidation

Frozen integration candidate `780cf4c7eebd62784670e2ab14cc49088a4bc91c`
(tree `2271acf5e4fd19e7d79d02fe25a3ba52c62569ba`) passes all eight
compiler/consumer rows: formatting of 510 changed Rust files, minimal PPTX,
strict OOXML leaves/ops/facade Clippy, DrawingML allocator performance-harness
compilation, native-resave compilation, and minimal OPC/common compilation.
Root independently verifies the terminal capture, exact commands, source,
three locks, log/trace hashes, target bounds and exited owned processes.
Final target size is 1,147,736,122 bytes. These are compilation results;
they do not establish runtime fidelity or performance acceptance.

Evidence is external under the worktree parent's
`.coord/uo2-coordinator-replay-integration-f5-compiler-780cf4c7/`:

| File | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 62,815 | `eb727e051ab56631e2641d81c0368886c36be75ca557b1a2722ed3957f322c28` |
| `result.json` | 66,070 | `5163e66978cbdef3751385fc89db31460f88621f32994ad4faed3ef0416415a8` |
| `run_f5.py` | 15,126 | `410d219dc7d14e302150fef50d44544912121aee9344ac23b8ed20bde3754eab` |

Reproduction uses that source and the three approved locks from 132 with a
fresh copy of the packet/target, followed by `python3 -B run_f5.py --go`.
The root verifier is `.coord/checkpoint96/verify_capture.py STATE`.
After successful terminal verification and a fresh root preflight, the
34-group runtime packet starts at
`.coord/ooxml-pptx-joint-c780cf4c7-f5-runtime/`. Its result remains pending.
The reviewed glossary26 and full53 packets retain isolated rustdoc rows.

## Cleanup101

Root removes two obsolete detached ODP checkouts without force:
`.coord/odp-matched-e4-source-20261024` and
`uo2-odp-heap-repair-0362yd-baseline`. Both have the canonical historical
`e4a0614f5d3a1d768346cc33a4a3b86522dda986` tree. The required comparison
endpoint `.coord/odp-0362yd-baseline-clean` remains unchanged.

The stale sparse checkout's lock, route overlay and probe are archived,
restored from the archive and reverified before removal. The route is
identical to its tracked historical source. The probe differs only in the
equivalent bounded `ALLOC_LOG` iteration spelling; its exact bytes and
historical diagnostic intent remain preserved. The three archived payloads
are unlinked only after regular-file, ancestor, size, mode and hash checks.

The fresh CWD/FD holder guard reports zero visible holders and zero fatal
findings. Permission-denied system/container and named service processes,
plus disappearance races, prevent complete process visibility. Those limits
are recorded; no complete visibility claim is made. Root verifies absent
paths/registrations and unchanged root/integration heads, refs, canonical
baseline, archives, approved locks and protected constructor probe.

Evidence is under external
`.coord/worktree-cleanup-2026-10-25/trunk-consolidation/`:

| Receipt | SHA-256 |
|---|---|
| `quarantine/wave101-root-removal/preflight.json` | `eb23980a2d745a5959269314403216058f48e1276585b27fcc2baec26458138c` |
| `quarantine/wave101-root-removal/execution.json` | `4a133775228f31b4664ab182623d1028927ab6dc2abe64f74daae76419ac7e31` |
| `quarantine/wave101-root-removal/postflight.json` | `4f0eb93b2d24aa02e302b6f671c18b4d932dde8b3f66d6a26712b4a7104f6c7f` |
| `cleanup101-root-allocation.json` | `572842122aceca084d1e4ab24f43feec2165234a90c367dd0997d14f32123201` |

Source allocation was 7,602,479,104 bytes; quoted quarantine directories
retain 229,376 bytes. The conservative ledger therefore increases by
7,602,249,728 to **785,182,834,688 bytes**, under all prior exclusions and
the separately disclosed top-level evidence overhead. This is not a
filesystem free-space delta. Totals are **268 worktrees removed and 430
refs deleted**. Required comparison endpoints remain until measurement.

A fresh census finds no overlooked unified-derived branch: only trunk,
private integration and the active runner remain. The runner is already an
integration ancestor and can be deleted when qualification no longer uses
it. Main/office and 68 independent office siblings are outside this cleanup.

The 126 performance adapter v3 passes independent review and eight focused
synthetic tests; its candidate measurements remain unrun. The notes packet
is rebound to C with the same eight harness blobs and passes sixteen
build-free tests. Census v7 is returned for failing to seal its binding
before capture; v8 must reject changes during execution before GO. Neither
prepared scripts nor synthetic receipts count as final measurements.

Runtime, glossary, full53/wsdoc, raw API census, fidelity/performance and
final integration into `feat/unified-ops-api` remain open. Nothing is pushed.
