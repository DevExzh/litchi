# 0440: Borrowed ODP attribute namespace cache

This private parser change retains the shared attribute iterator and cached
namespace resolution, but borrows each bound namespace URI from the reader.
It removes the per-attribute URI vector while making the existing reader-scope
invariant part of the Rust lifetime contract. The one-shot lookup path, error
ordering, decoding, drawing-attribute order and public APIs remain unchanged.

The [decision](decision.json) keeps the change for its measured allocation
benefit. Medium/large owned ODP append cases make 20.101%/20.163% fewer
allocation calls and request 6.322%/6.562% fewer cumulative bytes in both
repeats. Peak above entry and retained live bytes are unchanged. There is no
normal latency or RSS improvement claim.

The [frozen main protocol](protocol.json) retains 24 reports and 720 samples in
A1/B1/B2/A2 order, with three shapes, normal/allocator binaries, 30 samples,
three warmups, CPU 2 and one worker. Main large R2 p95/p99 increased about 8%.
The separately frozen [confirmation plan](confirmation-plan.json) retains four
additional large normal reports and 120 samples; the adverse tails did not
recur in either confirmation pair. Original tail observations remain in the
main summary. See [measurements](measurements.md) for the individual results.

Four selected whole-process profiles retain stat counters, raw perf data and
symbolized text. They include setup, warmups and oracle work. The original
baseline record attempt and first C1 confirmation overlapped during profile
conversion; [overlap-exclusion.json](overlap-exclusion.json) excludes both and
binds their replacements. Both original attempts remain for review. The main
matrix was serialized and is unaffected.

The independent report oracle is byte-identical to 0439. Source and output
archive hashes, complete slide semantics, opaque-member preservation, exact
no-op sharing, reversible patches and stale-source refusal remain required.
Reports do not embed raw archives: the Rust fixture gates inspect actual
archives, while Python independently regenerates semantic and opaque hashes.

Portable verification requires Python 3; it does not require Cargo, perf,
original binaries, original temporary directories, or the checkout used for
capture. Run from this directory:

```sh
python3 -B verify.py --sealed --cleanup
python3 -B derive.py --check
python3 -B derive-supplement.py --check
```

The verifier checks frozen oracle and driver hashes, all selected report
oracles, source/build/binary identities, exact profiler and capture commands,
ordered phase indexes, measurement chronology, validation receipts, cleanup
proof and the complete [SHA256SUMS](SHA256SUMS) inventory. Mutation probes
accept an unmodified copied control before rejecting corrupted copies with
refreshed inventories; report mutations also refresh their artifact hashes.
Completed probe receipts are retained under checks. Adding any receipt requires
resealing afterward.

[compression.json](compression.json) binds original and deterministically
gzipped log/profile bytes. Four explicitly owned temporary directories are
removed only after precleanup proof; shared build caches and the unchanged
user-owned GOAL.md are preserved. See [validation history](validation-notes.md),
[source review](parser-review.md), and [ADR obligations](adr-compliance.md).

The selector registry remains 436 and the default matrix remains 36 cases.
This closes one allocation-ownership experiment. One-shot cache costs,
repeated staging/validation, bounded existing-document append, package-Part
addition, repackaging, native breadth, cold/range I/O and scaling remain part
of the active non-iWork goal.
