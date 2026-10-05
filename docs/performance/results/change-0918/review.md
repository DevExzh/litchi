# 0918 current DOC repeated-edit profile review

This is an independent read-only scope and correctness review of the current
0918 Callgrind packet. I did not build the probe, run a capture, change
production code, or modify the raw profiles. The review was performed on
2026-10-05.

## Identity and workload

The profiled executable is the retained 0917 native probe
`/home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe`,
SHA-256
`2ae48b4f7548c1468c8e7691d56d89b231c4a63f52ed48383f07fe3609c59a13`.
The 0917 source receipt binds it to revision
`bb19ca74bd52044fbd2c1c7cc9875e975bb334c0` and the frozen probe source.
The fixture identities in every retained JSON are correct: `FloatingPictures.doc`
is 335,360 bytes with SHA-256
`41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`, and
`NoHeadFoot.doc` is 26,112 bytes with SHA-256
`45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48`.

Each accepted profile runs the public probe with `--warmup 0 --samples 1`,
`--retention default`, and one of the following cases:

| case | edits | Callgrind Ir | Dr | Dw | raw profile (bytes, SHA-256) |
| --- | ---: | ---: | ---: | ---: | --- |
| `docfloat` | 2 | 27,097,518 | 14,357,540 | 15,802,260 | `callgrind-docfloat-e02-default-pinned/callgrind.out` (1,777,082; `543db852c16e40bd135f2c4b5e3679122a726feafc861364fb57f75ee96152b9`) |
| `docfloat` | 32 | 224,621,383 | 136,587,541 | 146,527,655 | `callgrind-docfloat-e32-default-pinned/callgrind.out` (1,865,959; `3a21e1da6b5626113ddabe18c4e5b8a4da5f65aab4f864fd5e79fa3fcbe28f32`) |
| `docnohf` | 32 | 24,163,805 | 12,188,466 | 13,127,269 | `callgrind-docnohf-e32-default-pinned/callgrind.out` (1,324,906; `ac8ec4553b3b6122a7db727fc8ea71690af296f550918d990b36408b0807db7e`) |

The profile event ratios are diagnostic: the 32-edit FloatingPictures
profile is 8.29× the two-edit Ir count, 9.51× the Dr count, and 9.27× the Dw
count. They do not predict native elapsed time.

## Boundary and correctness checks

The retained Callgrind profiles have nonzero summaries, process exit status
zero, and exactly one `doc_retention_probe_0917::measure_one` invocation and
one descendant `run_case` invocation. Independent `callgrind_annotate
--inclusive` runs assign the complete profile total to `measure_one`; the
`run_case` row is the expected nested lifecycle. The profiles contain no
oracle, `sample_report`, JSON serialization, or post-sample public-reader
frames. This establishes the intended boundary:

```text
measure_one
  -> timed_case
     -> run_case
        -> Snapshot::open_bounded
        -> one Edit with the requested repeated paragraph replacements
        -> commit
        -> output Vec copy
        -> local snapshot/edit/commit destruction
```

The JSON output for all three accepted profiles parses successfully. All 20
correctness fields are true, including exact cross-route output equality,
public-reader reopening, patch/refusal controls, retention transitions,
untouched stream and directory preservation, semantic oracle checks, and the
edit-count check. The output hashes are `a298ac14820cfa74f952367a2c9212f13654ac5aca306f46756745b7bfbe7e39`
for FloatingPictures with two edits,
`642063d159bdbcc65148087a1a491d7ff9a36814fcf4ee7a8a48c7c281021ad8` for
FloatingPictures with 32 edits, and
`0c91fc9286b8592e25f2f2cb0de23f5ae32244c43fec8ee7883cf437c710347f` for
NoHeadFoot with 32 edits.

The earlier `docfloat` two-edit profile with a 608-byte zero-summary output
is invalid and must remain a failed attempt. The direct whole-process profile
is also exploratory: it includes preparation and preservation checks and is
not interchangeable with the accepted measure-one boundary. Both attempts
remain outside the accepted result set.

Callgrind printed a `brk segment overflow` warning for the accepted runs, but
each process exited zero and emitted a nonzero profile with conserved
`summary`/`totals` values. The warning is retained as an environment
limitation; it must not be suppressed or used to claim allocator or native
memory behavior.

## Interpretation limits

Inclusive Callgrind rows overlap. For the two-edit FloatingPictures profile,
the largest diagnostic rows include `__memcpy_avx_unaligned_erms` at 45.29%
of Ir and `Edit::replace_paragraph` at 33.07%. At 32 edits,
`Edit::replace_paragraph` is 90.17% of Ir. These shares are ranking evidence
inside this one public lifecycle; they are not additive. Callgrind also prices
`rep movsb`/`memcpy` and software hashing differently from native hardware, so
the profile cannot establish a native latency, allocation, RSS, or hardware
cache-miss claim.

The accepted set profiles only the default 8 MiB retention route. It does not
compare zero versus default versus explicit release, does not cover cold
cache, range sources, concurrent execution, distinct-object bulk editing, or
native Office round trips, and does not authorize a production optimization.
The JSON `elapsed_ns` values emitted while running under Valgrind are not
native timing evidence and must not be mixed with the 0917 native sample
statistics.

## Acceptance gates

The scope and correctness of the three retained profiles are approved for
diagnostic use. Before the packet is treated as reproducible evidence, its
custody record must bind each raw profile and stdout/stderr file by size and
SHA-256 and record the exact Valgrind/Callgrind invocation, CPU affinity,
toolchain, Valgrind version, binary hash, fixture hash, start/end timestamps,
exit status, and the collection option that matched `measure_one`. The
failed zero-summary and whole-process attempts should remain listed as failed
attempts. No production or adoption claim follows from this profile packet.

## Packet-wrapper findings

The report's timing column has now been removed, which resolves the risk of
presenting Valgrind-run elapsed values as native timing. The receipt field is
now named `probe_elapsed_ns`; if it is retained, its description should state
that it is an ordinary probe value measured under Valgrind and has no native
timing meaning.

The reproducer refuses an existing output root, and the offline
`profile-replay.py` verifier now exists. I ran it and `profile-reduce.py`
without launching Valgrind; both exited zero. The verifier bound the binary,
fixtures, raw profiles, stdout files, exact input/output arguments, nonzero
scope edge, exit status, and all reported correctness controls for all three
pinned runs and all three repeats. The receipt still does not hash stderr, and
the verifier does not independently assert summary/totals conservation or the
absence of post-sample symbols; those remain useful hardening checks. The
retained raw files themselves have nonzero summaries equal to their totals and
contain no oracle, serializer, or post-sample reader frames in the independent
scope inspection.

## Prospective comparison runner

The prospective comparison packet is now authored, but no comparison manifest
or performance child has been frozen or run. I independently parsed both
scripts and the protocol without creating Python bytecode or launching a
probe. The native schedule is 192 children over 96 paired cells; the allocator
schedule is 96 children over 24 cells and two reversed repetitions. Labels are
unique, version order alternates in native rounds, and the allocator order is
reversed in repetition two. The bootstrap implementation uses the frozen
10,000 draws and seed 918, and the candidate patch identity check compares the
supplied packet patch with the candidate source receipt.

The candidate source and final-build receipts are internally consistent: the
977-byte patch has SHA-256
`7439f82ff725106b1782c7ce6e3d28b1978ec7d8831b24f3745f3813b466bc18`, the
before/after source manifests have the same file rows apart from their
generated timestamp, and the final native and allocator binaries match their
recorded hashes. The production tree has no differences between the baseline
revision and the current revision outside `docs/performance`; the candidate
therefore remains an isolated three-conversion source change. I independently
validated all 96 pre-capture qualification receipts: every expected version,
lane, fixture, edit count, and retention route appears exactly once, exits
zero, and passes the unchanged 0917 validator with its retained summary.
Parsing the retained reports and comparing the runner's metric-excluding
semantic signatures gives an exact match for all 48 baseline/candidate pairs;
the stored summary hashes are expected to differ where timing or allocation
metrics differ.

The quality failures are recorded as baseline-matched or pre-existing: the
workspace format check and public/private rustdoc checks remain failures, while
the changed-file format check, litchi-doc format, clippy, tests, doctests, and
targeted data-stream transaction pass. The final build receipt binds the
authoritative build; its quality record retains the historical initial-build
pointer because the binaries are byte-identical.

The synthetic-control and replay-relocation requirements below resolve these
pre-freeze custody gates. The current runner resolves repository-relative
packet descriptors from its checkout hierarchy and leaves raw artifacts and
binaries external, so an arbitrary copied packet is not self-contained. A
valid replay must use the documented reconstructed hierarchy or the original
hierarchy with its external custody paths; it must not silently resolve back to
the original packet when presented as a copy. These are custody gates, not
production or performance findings.

## Independent final scope conclusion

The three pinned Callgrind profiles and their three independent repeats pass
the requested diagnostic scope and correctness review. Each has a nonzero,
summary-conserved event record, process exit status zero, one
`measure_one`-to-`run_case` call edge, matching output identity, and all 20
retained correctness controls true. The `brk segment overflow` diagnostic is a
retained environment limitation and does not invalidate these profiles because
the processes exited successfully with conserved nonzero reports. The packet
supports scoped Callgrind attribution of the current DOC repeated-edit
lifecycle; it supports no native timing, allocation, RSS, or hardware-cache
claim and does not authorize production adoption.

## Pre-capture clearance

I independently reran the retained in-memory synthetic controls with
`qualification-synthetic-controls.py` (writing only a temporary report). The
full 192-child native and 96-child allocator schedules passed all positive
latency, allocation, peak-live, and adoption gates. The 1.20 latency, 1.02
allocation, and 1.06 peak-live controls each rejected adoption; missing-pair
and semantic-corruption controls failed closed with the expected
`EvidenceError`s. The retained witness is
`qualification-synthetic-controls.json` (2,570 bytes); its runner and
protocol descriptors match the retained 63,000-byte runner and 7,319-byte
protocol.

The retained `comparison-relocation-smoke.json` witness (4,503 bytes,
SHA-256
`21330a2da74a14d7197948e354e46208677113edc383400d8a6b0e79da604866`)
records a copied reconstructed hierarchy under
`/home/zhuhe/litchi-performance-0918/comparison-relocation-smoke`. Its
independent path check resolved the copied runner's repository root to that
hierarchy, found all eight probe files and both fixture hashes, and passed
`require_reconstructed_hierarchy` without using the original checkout. The
retained `verify-smoke.py` (1,326 bytes, SHA-256
`4b2b99aa44e3ed1b8f9b37111c591f321ad12ceb85d3009598e7da6ea92d28d0`) was
rerun with its recorded command and exited zero with the same `verified`
result. The README now documents this layout and separately states that the
supplementary qualification replay retains its original absolute raw, fixture,
and binary custody.

These checks clear the comparison packet for identity freeze and performance
capture. They do not approve any performance result or adoption decision. A
completed capture packet must still run the full copied-packet
`comparison-replay.py` end to end, with its real manifest, receipts, reduced
result, and external raw custody, before comparative evidence is accepted.

## Final independent comparison review

I independently recomputed the completed reduction from all 288 retained child
receipts and 864 compressed stdout, stderr, and process artifacts. Every child
exited zero; each packet gzip reproduced its external raw file by size and
SHA-256; the four executable descriptors, eight probe files, two fixtures, the
source heads, and the candidate patch all matched the frozen manifest. The
in-checkout `comparison-replay.py` also exited zero with 192 native and 96
allocator children and adoption false.

The independent 10,000-draw, seed-918 bootstrap exactly reproduced all eight
cells. NoHeadFoot passed at every edit count. FloatingPictures at two and
eight edits failed the lower-bound gate with 0.8712919852 and 0.9339403211.
The 32-edit default allocation target saved 5.7073% for NoHeadFoot and
4.6685% for FloatingPictures, so the savings gate passed while adoption
remained rejected. Native adverse latency flags were 17 p50, 19 p95, and 19
p99; no native RSS pair crossed +5%. Allocated-byte and peak-live flags were
zero. One allocator RSS adverse pair was retained: FloatingPictures, eight
edits, release route, repetition one, 13,768 to 14,604 KiB (+6.072%).

For relocation, I copied the complete packet file-for-file into a fresh
reconstructed hierarchy. A literal invocation of the frozen
`comparison-replay.py` failed closed on the retained reports' original
absolute fixture input path (`fixture path mismatch`); the pre-capture
path-only smoke did not exercise this validator boundary. The review-only
adapter then verified the copied fixture hashes and supplied the manifest's
repository-root/relative-fixture path as validation metadata without opening
the original path. Its replay verified all 192 native and 96 allocator
children. Flipping one copied packet gzip byte was rejected as `content
changed`, and toggling the copied reduction's adoption gate was rejected as a
reduction mismatch.

The small reviewer summary is
[`review-final.json`](review-final.json). The external recomputation and
relocation proofs are retained at
`/home/zhuhe/litchi-performance-0918/reviewer-final/recompute-proof.json`
(3,869 bytes, SHA-256
`4336b22ef5ad5768609d5a07ae9259a72335315d1ee974028b94b16f43513816`) and
`relocation-proof.json` (2,923 bytes, SHA-256
`45a8294ac51d95897838c0b859f6c4ef24ada35af1b10f249c0a44958e7b9966`). The
relocation result is verified through that review adapter; the unchanged
frozen replay's fixture-path limitation remains recorded and does not change
the scoped rejection or authorize production adoption.

## Committed relocation-adapter review

I then audited and executed the committed `comparison-analysis.py` bytes
(13,553 bytes, SHA-256
`1c72d1d84ba7a789fced7ec24d4d7e275ef9a6a8a1b9fbd92ef24f7f9bd60d87`) in a
fresh reconstructed hierarchy. Two default, stdout-only runs were identical;
each verified 288 canonical reports and reported adoption false. The copied
runner, replay, protocol, manifest, reduced result, 0917 validator, and both
fixture files matched their recorded hashes. The canonical report aggregate
was `6247bd652bf353250de1d960309dfaa31d4caf0095e8ce6bb4f73e9a693f00d0`.

The committed adapter rejected a tampered manifest fixture path/alias and a
modified copied fixture. Independent copies with one retained raw gzip byte
changed and with the reduced result modified also exited nonzero: the former
reported `content changed`, and the latter rejected the frozen reduced-result
byte count. These checks cover the alias, fixture, raw-artifact, and reduction
custody boundaries without changing the frozen runner or reports. The exact
commands, outputs, return codes, and hashes are retained in the external
`committed-analysis-proof.json` (7,017 bytes, SHA-256
`98435feb868534df914516e46d7ca7e07beaa9c87c2f98f78033def94d008fe2`). The
temporary reconstructed and corruption copies were removed after these
checks; the original packet and raw captures remain outside the repository.
The cleanup receipt is retained at
`/home/zhuhe/litchi-performance-0918/reviewer-final/cleanup-receipt.json`
(1,565 bytes, SHA-256
`9f435e7a94e0c6e2d9d87c64e371e4414bdb3cb9130f4a6f7e9e04ad6f88490c`).
