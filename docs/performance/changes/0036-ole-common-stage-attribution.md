# Change 0036: common OLE2 publication stage attribution

## Scope and disposition

The standalone harness now splits the retained opaque-heavy common OLE2 case
into public editor open, candidate `put_stream` publication, changed `finish`
rendering, and the existing end-to-end control. The three new cases are opt-in;
the default 36-case / 198-record matrix is unchanged.

An inline recapture-allocation reuse prototype was measured and fully reverted.
It retained the complete candidate render, CFB reopen, protection check, every
stream read and exact comparison, fresh directory metadata, object rediscovery,
and final render. It merely compared each recaptured `Vec<u8>` with its staged
stream before converting it to a new `Arc<[u8]>`, so exact matches could reuse
the staged allocation without the otherwise discarded Arc allocation/copy.
The end-to-end improvement was below the practical gate, so no production or
test-only prototype code remains.

## Deterministic attribution case

All four cases use the existing fixed CFB artifact with four unchanged
incompressible 4 MiB regular-FAT streams and one 36-byte MiniFAT target. The
source contains 16,777,252 logical stream bytes and has SHA-256
`7ffbd37c3e472a21b382bcbb02e430a62164e58d2270bbee0deaa584ff47a94d`.
The changed output remains
`b9323eeace80e2c9c88801879265bfdfac83690bb2550880f5ef6bf87b48d131`.

Preparation is outside each timed boundary. Open clones the source before the
clock. Candidate publication edits a fresh editor derived from one pre-opened
snapshot. Changed finish uses a fresh editor derived from one already validated
changed snapshot. Exact output comparison, all-stream readback, and the
independent public `OleFile` reopen remain outside timing.

On the retained current-production state, CPU-2 ABBA with 20 warmups and 200
samples per leg gives 400 pooled samples per stage:

| Current stage | p50 | mean | p95 | p99 |
|---|---:|---:|---:|---:|
| Open | 1.382 ms | 1.391 ms | 1.497 ms | 1.585 ms |
| `put_stream` candidate publication | 7.979 ms | 8.040 ms | 9.236 ms | 9.775 ms |
| Changed `finish` render | 5.473 ms | 5.555 ms | 6.315 ms | 6.477 ms |
| End-to-end open/edit/finish | 26.086 ms | 26.143 ms | 27.159 ms | 27.905 ms |

The isolated p50 stages sum to 14.834 ms, only 56.86% of the chained p50.
They are therefore attribution boundaries, not additive components: freshly
opened payload locality, allocator state, and cache pressure materially change
the chained path. No claim derives an unmeasured stage by subtracting medians.

## Rejected inline recapture allocation reuse

The prototype and unchanged binaries used the same committed harness. Their
SHA-256 digests are `d22e7467...3aaa` and `fe22672f...c406`, respectively.
Matched ABBA used the same CPU, warmups, sample count, corpus, and case order.

| Case | Before p50 | Prototype p50 | p50 delta | mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|---:|---:|
| Open guard | 1.382 ms | 1.384 ms | +0.13% | -0.31% | -2.19% | -3.59% |
| Candidate publication | 7.979 ms | 7.461 ms | -6.49% | -5.95% | -8.49% | -10.46% |
| Changed finish guard | 5.473 ms | 5.299 ms | -3.17% | -4.21% | -10.92% | -11.41% |
| End-to-end control | 26.086 ms | 25.404 ms | **-2.61%** | **-2.30%** | +0.54% | +1.24% |

The intended isolated stage improved, but the complete public transaction did
not reach the 5% practical gate and its p95/p99 did not improve. The movement
of the logically unchanged finish stage also reinforces the binary-layout and
cache-coupling warning from change 0033. Pursuing DOC/XLS/PPT owner guards,
allocation/RSS claims, or a production handoff after this primary rejection
would not make the end-to-end result useful, so the prototype was removed.

## Profile attribution

The baseline publication-only cycle profile attributes resolved copy work to
`OleWriter::create_stream`, sector emission, `OleFile::open_stream` during full
`Package::capture`, and `capture_container`; exact post-capture comparisons
also reach libc `memcmp`. The chained profile is more copy-heavy still. Full
sample counts, commands, symbol-map limitation, and non-additivity warning are
retained in
[`ole-common-stage-profile.txt`](../results/ole-common-stage-profile.txt).

Raw ABBA reports are [`before A`](../results/abba-ole-recapture-before-a.json),
[`after A`](../results/abba-ole-recapture-after-a.json),
[`after B`](../results/abba-ole-recapture-after-b.json), and
[`before B`](../results/abba-ole-recapture-before-b.json). Binary and evidence
digests are indexed by
[`ole-common-stage-sha256.txt`](../results/ole-common-stage-sha256.txt).

## Verification and remaining work

- The stage-equivalence harness regression proves every case uses the same
  corpus and deterministic changed artifact and independently reopens all five
  streams.
- All 27 standalone harness tests and warning-denied all-target Clippy pass.
- The workflow smoke and scheduled release jobs now run all four stage/control
  cases with the exact corpus identity.
- Formatting, YAML parsing, and diff-hygiene gates pass.

Do not revive the rejected shared-payload, terminal-render cache, or inline
recapture-allocation reuse designs. A future OLE2 change must remove materially
different work and keep the end-to-end control primary. Higher-ranked current
program candidates are source-backed OPC same-topology publication and
selector-first ODT paragraph reads; RTF lexical scanning first needs the
formatted-corpus evidence gate.
