# Common OLE2 publication handoffs are rejected

Date: 2026-08-11

Production base: `e0434ae15`

Disposition: the deterministic benchmark is retained; both production
prototypes were measured and fully reverted. OLE2, OOXML, RTF and ODF
production code are unchanged by this record. iWork/IWA was explicitly
excluded.

## New end-to-end attribution case

The opt-in `ole_common_one_edit_save` case exercises the public common OLE2
editor over a dedicated deterministic CFB artifact. The `few-large` /
incompressible form contains four unchanged 4 MiB regular-FAT streams and one
36-byte MiniFAT edit target: five streams and 16,777,252 logical payload bytes
in total. The 16,913,408-byte source SHA-256 is
`7ffbd37c3e472a21b382bcbb02e430a62164e58d2270bbee0deaa584ff47a94d`.

The timed operation clones prepared input and replacement bytes, then calls
`Editor::open`, `put_stream` and `finish`. Corpus construction, the expected
output, exact output comparison and final public `OleFile` reopen are outside
the timed interval. Verification checks the replacement, every byte of all
four unchanged streams and the exact stream count. The deterministic changed
output SHA-256 is
`b9323eeace80e2c9c88801879265bfdfac83690bb2550880f5ef6bf87b48d131`.
Tight common-editor limits admit exactly the five-stream workload rather than
using broad defaults.

The case is opt-in, so the 36-case / 198-record default matrix is unchanged.
It raises the selectable harness count to 112 and the harness test count to
25.

## Prototype A: shared CFB writer payload

The first prototype added a private `Owned(Vec<u8>)` / `Shared(Arc<[u8]>)`
writer payload, preserved the existing borrowed and owned allocation
contracts, and passed common package streams to the writer by shared
ownership. Candidate validation, CFB render/reopen/capture, allocation reuse,
object rediscovery and the final render remained unchanged.

Focused tests proved retained `Vec` and `Arc` allocation identity, repeated
writes, replacement/deletion release and byte-identical borrowed/owned/shared
serialization at 4095, 4096 and 4097 bytes for both 512- and 4096-byte sectors.
All 120 CFB tests and all 145 common-layer unit/integration tests passed before
the prototype was reverted.

Matched before-after-after-before runs used 10 warmups and 100 samples per leg
on CPU 2. Pooling the two same-state legs gives:

| Heavy common edit/save | Before | Shared prototype | Delta |
|---|---:|---:|---:|
| p50 | 16.663 ms | 21.999 ms | **+32.02%** |
| mean | 16.494 ms | 22.031 ms | **+33.58%** |
| p95 | 17.855 ms | 23.277 ms | **+30.36%** |
| p99 | 18.581 ms | 24.683 ms | **+32.84%** |

The writer's existing borrowed/owned staging-only guard remained flat, but the
end-to-end result is an unambiguous regression. A plausible inference is that
the staging copy also establishes favorable payload locality for the following
serialization; the experiment does not prove that microarchitectural cause.
No allocation or memory claim was pursued after the latency gate failed.

Before executable SHA-256:
`dfcc2953cdadbf8b817e9a2f576b3ca1aea6599649a342923b63e7681a6aa33c`.
Shared-payload prototype SHA-256:
`943a1620980c64c945ba3d5d49b713574c958ffbd57b9dd84c718b6bc760ae17`.

## Prototype B: retain the validated render

The second prototype retained the exact `Vec<u8>` produced by
`commit_candidate`. Those bytes had already passed package limits, complete
render, protected-container rejection, CFB reopen, complete package capture,
unchanged-allocation reuse and object rediscovery. A changed `finish` moved
that exact validated allocation instead of rendering the recaptured package a
second time. Snapshot-derived editors without a retained artifact kept the
existing render fallback, and exact no-ops kept returning the original source.

A private test proved that ordinary and cloned changed editors returned the
exact validated artifact and that it reopened with the replacement. The
complete common test suite and the deterministic heavy harness case passed.
The same 10-warmup / 100-sample-per-leg protocol showed the intended path was
material:

| Heavy common edit/save | Before | Cached prototype | Delta |
|---|---:|---:|---:|
| p50 | 18.296 ms | 12.064 ms | **-34.06%** |
| mean | 18.377 ms | 12.280 ms | **-33.17%** |
| p95 | 19.296 ms | 13.445 ms | **-30.32%** |
| p99 | 19.786 ms | 13.828 ms | **-30.11%** |

The mandatory native DOC/XLS guard used 30 warmups and 1,000 samples per leg
for open, exact no-op edit/save and one semantic edit/save on the large writer
corpora. Pooling 2,000 samples per state exposed an unacceptable DOC
regression:

| Guard | Before p50 | Cached p50 | p50 | Mean |
|---|---:|---:|---:|---:|
| DOC open | 0.741 ms | 0.902 ms | **+21.64%** | **+21.21%** |
| DOC exact no-op edit/save | 0.227 ms | 0.224 ms | -1.04% | -1.07% |
| DOC one-edit/save | 1.327 ms | 1.448 ms | **+9.08%** | **+8.57%** |
| XLS open | 1.468 ms | 1.341 ms | -8.64% | -9.64% |
| XLS exact no-op edit/save | 2.753 us | 2.734 us | -0.69% | +1.30% |
| XLS one-edit/save | 1.597 ms | 1.626 ms | +1.80% | +1.67% |

DOC open cannot consume the terminal cache. The stable two-leg movement may be
caused by editor-layout or binary-layout/code-generation effects, but that
inference does not make a measured 21% regression acceptable. It also repeats
the broader lesson of the rejected XLS-only terminal handoff in change 0028:
removing a provably duplicated render is insufficient unless every ordinary
owner guard remains within threshold.

Cached-render prototype SHA-256:
`d8664fd9066cee55418778d20bc1cd48a5885a84d94deaa814988d7194997d59`.

## Final state and next work

Both speculative production implementations and their private tests were
removed with `apply_patch`. No low-level shared-payload API, editor field,
cache, dependency, public archive type, runtime, lock, unsafe code, validation
shortcut or owner handoff remains. The unrelated existing
`docs/FORMAT_IMPLEMENTATION_REVIEW.md` worktree edit remains unstaged.

The retained benchmark makes future OLE2 proposals answerable at the complete
public transaction boundary. The next candidate must be materially different:
attribute and reduce recapture or final owner/public-reader work without
reviving either direct shared serialization or an editor-wide terminal cache,
and keep DOC open/no-op/edit plus XLS open/no-op/edit as mandatory guards.

Raw primary and guard ABBA reports are under `docs/performance/results/`; their
digests are in `ole-common-publication-sha256.txt`.
