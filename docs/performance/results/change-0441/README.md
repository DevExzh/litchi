# 0441: Shared immutable source-slide projection

ODP transaction staging now shares the snapshot's immutable slide projection
for pristine-page comparisons. Its editable Vec remains detached. This removes
one deep temporary copy without changing namespace parsing, source coverage,
origin mapping, preservation comparisons, publication validation or public APIs.

**The original frozen practical gate failed.** Normal medium/large p50 is
0.664–1.279% slower, and allocation-call/requested-byte reductions are below 2%.
The [decision](decision.json) keeps the change under an explicitly
[post-hoc peak-memory review](acceptance-review.md): all observations in both
repeats show 9.834%/9.857% lower medium/large peak above operation entry.
The root omitted peak from its original gate despite naming peak in the
hypothesis. The frozen protocol and failed result are preserved unchanged.

The [main protocol](protocol.json) retains 24 reports and 720 samples in
A1/B1/B2/A2 order, with three shapes, normal/allocator binaries, 30 samples,
three warmups, CPU 2 and one worker. No matched adverse comparison exceeds
5%; baseline tiny normal p99 has a +5.807% repeat flag. All timing, allocation,
RSS and uncertainty values remain in the [measurements](measurements.md).
Four selected whole-process profiles include setup, warmups and Rust oracle
work. Baseline/candidate conversions retain 13/11 addr2line warnings each.

The independent report oracle is byte-identical to 0439 and 0440. Rust fixture
gates inspect actual archives and verify complete slide semantics, one exact
tail append, opaque preservation, no-op sharing, reversible patches and stale
source refusal. Python independently regenerates semantic/opaque expectations
and checks report contracts; reports do not embed full archives for Python to
reopen. No source-read, physical-copy, passthrough, native or scaling claim is
made. Source and committed snapshot remain live at the allocator endpoint;
retained live bytes are unchanged and append remains materialized.

Portable verification needs Python 3, without original binaries, temporary
directories, perf, Cargo or the capture checkout:

```sh
python3 -B verify.py --sealed --cleanup
python3 -B derive.py --check
python3 -B profile-summary.py --check
python3 -B measurements.py --check
```

The verifier binds frozen protocol/oracle/driver hashes, exact source/build
identities, commands, ordered capture indexes, all selected report oracles,
profile artifacts, measurement chronology, validation receipts, cleanup proof
and the complete SHA256SUMS inventory. Copied controls pass before mutated
copies are rejected after inventory refresh. Mutated report probes also
refresh report artifact hashes. Compression records bind raw and deterministic
gzip bytes. Completed probe receipts require resealing before final read-only
verification.

Capture, profile, derivation and sealing machinery is adapted from 0440.
Retained before/after source copies describe the candidate, with fresh build
commands and binary hashes in each role's descriptor. The optional historical
prior-attribution.py reads the sealed 0440 profile; it is context and is not
required for portable verification of this bundle.

The registry stays at 436 selectors and the default matrix at 36 cases.
Repeated staging XML scans, one-shot attribute caching, bounded existing
append, Part addition, repackaging, native breadth, cold/range I/O and scaling
remain open. This batch does not complete the non-iWork goal.
