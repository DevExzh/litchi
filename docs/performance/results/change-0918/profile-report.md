# Change 0918: scoped DOC repeated-edit profile

This packet profiles the current `f5485b3ca8582051247d31946b1c5f45c8a561ad`
DOC repeated-edit probe against the production implementation at
`bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`. HEAD `f5485b3c` adds the 0917
probe and evidence only; it does not change the production DOC implementation.
This is an attribution study for the already measured 0917 lifecycle; it
contains no production change and makes no optimization claim by itself.

The retained release binary is
`/home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe`
(55,146,832 bytes, SHA-256
`2ae48b4f7548c1468c8e7691d56d89b231c4a63f52ed48383f07fe3609c59a13`). The
fixtures are `FloatingPictures.doc` (335,360 bytes, SHA-256
`41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`) and
`NoHeadFoot.doc` (26,112 bytes, SHA-256
`45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48`).
The host is an AMD EPYC 9R45 with 32 logical CPUs, Linux 7.0.0-1013-aws,
Rust 1.95.0, and Valgrind 3.26.0. Each final run was pinned to CPU 14.

## Measurement boundary

Callgrind instrumentation was enabled for the process but collection began only
on entry to the non-inlined Rust symbol
`_ZN24doc_retention_probe_091711measure_one17h2b5462a44ecd6077E` using:

```text
--instr-atstart=yes --collect-atstart=no --toggle-collect='*measure_one*'
--collect-jumps=yes --simulate-cache=yes --dump-instr=yes
```

The probe used `--warmup 0 --samples 1`. The collected call edge from
`measure_one` to
`_ZN24doc_retention_probe_09178run_case17ha97621af2661a28eE` has `calls=1` in
each raw profile. Therefore the collected subtree is one
`Snapshot::open_bounded` through repeated `Edit` replacements/releases,
`commit`, output `Vec` copy, and local lifecycle destruction. Source loading,
the three untimed reference routes, correctness/oracle checks, hashing, and
JSON serialization are outside collection. `measure_one` and its clock/result
wrapper contribute only their small direct overhead.

Callgrind instruction and data-reference counts are simulated events, not
native cycles, wall time, or allocation bytes. The existing 0917 allocator
regions remain the source for lifecycle allocation totals. The raw profiles and
full `callgrind_annotate` outputs are retained outside the checkout under
`/home/zhuhe/litchi-performance-0918/callgrind-*`; their sizes and SHA-256
values are in `profile-receipt.json`.

## Scoped results

The first three event columns are instruction references (`Ir`), data reads
(`Dr`), and data writes (`Dw`). Percentages in the attribution table are
inclusive within that one collected lifecycle and must not be summed across
rows.

| case / edits / route | Ir | Dr | Dw |
| --- | ---: | ---: | ---: |
| `docfloat` / 2 / `default` | 27,097,518 | 14,357,540 | 15,802,260 |
| `docfloat` / 32 / `default` | 224,621,383 | 136,587,541 | 146,527,655 |
| `docnohf` / 32 / `default` | 24,163,805 | 12,188,466 | 13,127,269 |

| inclusive function or operation | docfloat e02 Ir | docfloat e32 Ir | docnohf e32 Ir |
| --- | ---: | ---: | ---: |
| `Edit::replace_paragraph` | 8,959,932 (33.07%) | 202,541,505 (90.17%) | 22,247,964 (92.07%) |
| `RevisionEditor::replace_plain_text` | 7,483,690 (27.62%) | 176,011,245 (78.36%) | 20,695,891 (85.65%) |
| `RevisionEditor::commit` | 5,544,328 (20.46%) | 128,114,569 (57.04%) | 16,034,819 (66.36%) |
| `Editor::commit_candidate_with_rendered` | 5,149,246 (19.00%) | 111,334,454 (49.57%) | 13,868,573 (57.39%) |
| `Editor::put_streams_shared_with_rendered` | 5,235,889 (19.32%) | 112,716,817 (50.18%) | 14,254,685 (58.99%) |
| `Package::render_with_layout` | 2,484,594 (9.17%) | 51,007,229 (22.71%) | 8,097,200 (33.51%) |
| `OleWriter::write_to` | 1,960,311 (7.23%) | 42,606,458 (18.97%) | 5,736,899 (23.74%) |
| libc `__memcpy_avx_unaligned_erms` | 12,273,510 (45.29%) | 119,678,734 (53.28%) | 9,822,557 (40.65%) |
| `Arc::from_box_in` | 836,121 (3.09%) | 1,169,203 (0.52%) | 92,897 (0.38%) |

The 32-stage FloatingPictures lifecycle therefore spends most of its collected
instruction and data-reference budget in the repeated paragraph replacement
path, with the final package publication/rendering accounting for a large
inclusive subtree. The libc copy routine is the largest direct copy signal in
all three runs. The no-header/footer control has the same repeated-edit shape
but a much smaller absolute lifecycle. These observations identify where a
candidate should be measured; they do not establish that a particular source
change is safe or beneficial.

The current `RevisionEditor::commit` source constructs the changed Word and
Table streams as `Arc::<[u8]>::from(self.word.clone())` and
`Arc::<[u8]>::from(self.table.clone())`. The profile shows the surrounding
publication and copy work, but cannot attribute allocated bytes to individual
source expressions. Any candidate that changes these conversions requires a
fresh build, exact-output qualification, and a prospective comparison against
this current-source baseline.

## Independent repeat

The same three commands were captured again later under the same binary,
fixture hashes, CPU affinity, and collection boundary. The repeat is retained
under `/home/zhuhe/litchi-performance-0918/replay-final` and is recorded as
separate runs in `profile-receipt.json`; it is not treated as a byte-identical
rerun.

| repeat | Ir | Dr | Dw | raw callgrind SHA-256 |
| --- | ---: | ---: | ---: | --- |
| `docfloat` / 2 | 27,107,421 | 14,367,664 | 15,812,356 | `996180ce347e4ee80524ac749b9ece7ea8056663dc2b40ee79e13ddb16245d00` |
| `docfloat` / 32 | 230,402,404 | 142,320,923 | 152,256,410 | `72ab2c4b91120039f353a345e5e75c87dde25277e6be3ce6d633acceecce9f03` |
| `docnohf` / 32 | 24,145,247 | 12,165,069 | 13,103,788 | `72dc177d0228321e7544bba3c97265c9ac1fda8408338455a28a07ef55ccfec4` |

Each repeat preserved the exact output hash, output size, and output capacity of
its corresponding pinned run, and each raw profile retained the
`measure_one` to `run_case` edge with `calls=1`. The 32-stage FloatingPictures
repeat is 2.57% higher in simulated instruction references than the pinned
224,621,383-Ir capture; the other two Ir deltas are +0.04% and -0.08%. The
different raw hashes and cache-reference counts show that this is launch-level
runtime drift (guest scheduling, layout, or cache state may contribute), so the
two captures have no identity claim and are not averaged. The pinned profiles
remain the attribution values reported above. `profile-reduce.py` emits every
retained run's nine event counts from its hash-bound raw file.

## Attempts and limitations

The first attempt used `--instr-atstart=no` with the toggle option and collected
zero events. Its probe still exited zero; Valgrind also printed a known `brk`
segment-overflow diagnostic. It is retained as a configuration failure, not a
probe failure. A separate collect-at-start run produced 1,206,470,389 whole
process instruction references and is retained as an unscoped diagnostic. The
final command uses instrumentation at start with collection disabled, then
the wildcard `measure_one` toggle; it produced nonzero scoped events and the
call-edge proof above.

Reproduce the three final profiles with
`docs/performance/results/change-0918/profile-reproduce.sh`. It requires a new
external output directory and verifies the binary and fixture hashes before
recording fresh start/end/status files. Validate retained profiles without
launching Valgrind with `profile-replay.py`; it checks each retained raw and
stdout/stderr hash against that run's receipt, the exact nine-event order and
summary, the one-call scope edge, output guards, the exact 20 correctness keys,
and the one-sample inactive-frame gates. `profile-reduce.py` shares those
validation gates before emitting its per-run event counts. Raw profiles are
intentionally not committed because each exceeds the repository evidence-size
threshold.
