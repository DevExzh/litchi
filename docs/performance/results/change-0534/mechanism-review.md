# 0534 physical reconciliation mechanism review

The 0534 runtime candidate is rejected. It replaces the final CFB physical
role/FAT loop with a paired common-prefix walk, but all eight primary XLS
native p50 rows are slower in the matched comparison. The regressions range
from **1.0596%** to **10.6386%**, so the candidate does not meet the plan's
3% improvement requirement in any primary row. The production runtime must
retain the restored baseline physical loop; this review does not accept a
physical-loop optimization.

The scope is exactly the private
`OleFile::validate_physical_sector_layout` hunk in
[`candidate.patch`](candidate.patch). `collect_exact`, `claim_sector`, role
publication, stream validation, FAT loading, call order, public APIs and
error contracts are outside the runtime diff. The focused contract tests may
remain with the restored runtime, but their correctness result cannot turn a
failed native admission gate into a performance acceptance.

## Native decision

The matched [comparison](comparison.json) has equal corpus and output
identities. It contains the frozen two-repeat ABBA run over the nine XLS and
three CFB workflows. The eight primary rows are:

| Primary case | Repeat | Baseline p50 (ns) | Candidate p50 (ns) | Candidate change |
| --- | ---: | ---: | ---: | ---: |
| `xls_source_backed_open` | 1 | 101,940 | 112,785 | **+10.6386%** |
| `xls_source_backed_open` | 2 | 108,531 | 109,681 | **+1.0596%** |
| `xls_source_backed_open_one_cell` | 1 | 107,005 | 115,941 | **+8.3510%** |
| `xls_source_backed_open_one_cell` | 2 | 112,321 | 115,130 | **+2.5009%** |
| `xls_owned_source_open` | 1 | 96,970 | 103,891 | **+7.1373%** |
| `xls_owned_source_open` | 2 | 100,240 | 105,180 | **+4.9282%** |
| `xls_owned_source_open_one_cell` | 1 | 99,220 | 105,830 | **+6.6620%** |
| `xls_owned_source_open_one_cell` | 2 | 102,955 | 109,195 | **+6.0609%** |

The comparison artifact's `pass` status means that its identity and custody
checks completed; its primary workflow p50 admission gate is false. The
allocation guard's calls, allocated bytes and incremental-region peak remain
equal. This passes the allocation guard and does not override the latency rejection.

## Actual generated code shape

The [assembly analysis](assembly-analysis.json) is bound to the measured
normal binaries: baseline
`44c80a425a719441e0182d4a7932fcbe6309a84a816a75d75c764e90a0a13e0f` and
candidate
`4e5e4e1eb6aa5128a0a9c8b6c7dc321b5901cf77e9a9028b0451c0786ba0f71b`.
The comparison is bound to plan SHA-256
`4e1a13bdfa7f7c0268065d52def107f085b07cc0587d39685d9587de181d40dc`,
baseline/candidate source manifests
`ce5a705508299624c2611d525940ce1502b3b9f20196814cf449e9286f8f11d6` and
`232a89a94851140ddbe47b9f48387248be900ebbdb81f2c75b826740fa744252`, and
the frozen run script; stage receipts bind the binaries and host artifacts.
Each stage has twelve monomorphized physical-layout variants. Every baseline
variant is 275 bytes with a `0x50` stack reservation; every candidate variant
is 220 bytes with a `0x30` reservation. The selected
`validate_stream_allocations` variants are unchanged in count, size and
stack reservation (28,373 bytes across twelve variants), which is consistent
with the runtime diff being limited to physical reconciliation. The complete
normal executable nevertheless changes from 60,140,440 to 60,142,888 bytes;
selected symbol size is not a whole-executable or latency measure.

The bound [baseline physical disassembly](baseline/assembly-validate_physical_sector_layout-0.stdout)
loads the role-map length and the FAT pointer/length, then keeps the
`fat.get(sector)` bounds comparison inside the loop. Each iteration stores the
index for a possible error, loads the same-index FAT marker, checks
`FREESECT`, loads the role byte, and branches to the existing formatted error
path when an unclaimed role has another marker. A missing FAT entry therefore
branches at the first missing index while the scan is in progress.

The bound [candidate physical disassembly](candidate/assembly-validate_physical_sector_layout-0.stdout)
loads both lengths first, selects their common-prefix length with a compare
and conditional move, and runs the loop over that prefix. The loop still
loads the FAT marker and role byte and retains the marker and unclaimed-role
branches. Once the prefix ends, one compare of the original lengths selects
the existing missing-entry error when the FAT is short. This is the generated
shape of `zip(...).enumerate()`, including the explicit post-loop error path;
it is not evidence that any individual instruction has a particular latency.
The candidate also leaves the extra FAT tail untouched, as required by the
padding contract.

The two disassemblies show a branch moved out of the per-sector loop, along
with a new minimum-length setup and a post-loop length branch. Static code
size and branch count do not predict the native result. In particular, the
candidate's smaller physical symbol cannot be treated as a measured speedup.

## Constructor attribution

The [profile comparison](profile-comparison.json) uses eight profiles per
stage, 40 timed constructor dumps and six separate CFB setup dumps. It selects
only dumps with a positive incoming edge from the benchmark caller to the
requested XLS or CFB constructor; setup dumps are kept separate. The parent
constructor total is reported independently from child rows. Inclusive child
rows overlap their parent and must not be subtracted or added to manufacture
an instruction saving. The physical row is a disjoint exclusive owner.

The relevant matched parent totals and exclusive owners are:

| Profile | Parent constructor inclusive Ir, baseline → candidate | Physical reconciliation self Ir, baseline → candidate | Chain collector self Ir | Stream validation self Ir |
| --- | ---: | ---: | ---: | ---: |
| XLS-owned, repeat 1 | 11,317,204 → 10,989,365 (**−2.8968%**) | 1,991,710 → 1,659,795 (**−16.6648%**) | 5,601,140 → 5,601,140 | 1,814,485 → 1,814,485 |
| XLS-owned, repeat 2 | 11,319,136 → 10,986,774 (**−2.9363%**) | 1,991,710 → 1,659,795 (**−16.6648%**) | 5,601,140 → 5,601,140 | 1,814,485 → 1,814,485 |
| CFB few-large, either repeat | 10,264,094 → 9,933,809 (**−3.2179%**) | 1,981,930 → 1,651,645 (**−16.6648%**) | 5,571,945 → 5,571,945 | 1,640,190 → 1,640,190 |
| CFB many-small, either repeat | 13,979,277 → 13,974,282 (**−0.0357%**) | 36,910 → 30,795 (**−16.5673%**) | 764,485 → 764,485 | 533,435 → 533,435 |
| CFB tiny, either repeat | 244,607 → 244,572 (**−0.0143%**) | 430 → 395 (**−8.1395%**) | 5,200 → 5,200 | 4,285 → 4,285 |

Each workload records fewer physical-reconciliation instructions while its
chain-collector and stream-validation exclusive totals remain unchanged.
These are local Callgrind results, not native latency gains.
The existing 0533 claim-helper layout is already represented by no positive
out-of-line `claim_sector` target in these selected dumps. That absence is
explicitly allowed by the analyzer and does not mean ownership work became
free; no claim code changed in 0534.

## Hardware context

Both stages have two measured, fully covered hardware groups. The
[hardware reports](baseline/hardware-analysis.json) and
[candidate hardware reports](candidate/hardware-analysis.json) cover the
whole `xls_owned_source_open_one_cell` child, including fixture setup,
copies, queries, correctness oracles, drop and report generation. They do
not measure operation-local physical reconciliation.

For context, candidate-versus-baseline grouped counters change as follows:

| Hardware repeat | Cycles | Instructions | Branches | Branch misses | IPC | Branch-miss ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | **+5.1379%** | −1.0508% | −3.2331% | +0.0668% | −5.8863% | +3.4101% |
| 2 | **+1.6248%** | −1.0985% | −3.2591% | +3.5423% | −2.6797% | +7.0305% |

The whole-child counters therefore cannot identify the cause of the native
regression. They do reinforce the boundary between generated-code evidence
and the constructor clock: fewer counted instructions in this broad child
did not produce lower cycles or lower p50 latency. No IPC, branch-prediction,
or operation-local throughput claim follows from these captures.

The evidence remains scoped to the frozen warm in-memory CFB/XLS corpus and
its exact source, binary, receipts, plan, scripts and host artifacts. It does
not establish cold I/O, other Office producers, provider breadth, or scaling.
OLE2 and OOXML remain the active optimization priority; ODF is deferred until
that goal completes, and iWork is excluded.
