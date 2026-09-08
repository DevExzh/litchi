# 0472: compact plain-cell snapshot tags

The candidate validates cell addresses and attributes as before, then omits
owned tags for exact unprefixed `c` cells with zero attributes or only one `r`.
Changed-cell writing already regenerates `r`, while untouched cells copy exact
source spans. Rich and prefixed cell tags remain owned. See the
[source review](source-review.md) and [frozen protocol](protocol.json).

Control `a82367b52` and candidate `3eac9a493` both use fresh release builds at
`/tmp/litchi-goal-0468/profile-tree`. The control contains the production code
restored after 0471's rejected experiment. Control binary SHA-256 is
`f64c7dd28b2e440a1d8bb69bf667719e557166e62c31234e3d59848469c8cdc9`;
candidate SHA-256 is
`b0299258993647b9c3087ceafd72ba2cacca30a37f1f057976dbf700ab76b89d`.
Both source inventories contain 6,993 files, with exactly five changed XLSX
implementation/test files and the same two compile-time fixtures. No harness
or corpus changes are part of the candidate.

The build uses Rust 1.98.1, release debug level 1, forced frame pointers/unwind
tables, four Cargo jobs and no incremental compilation. Measurements use CPU 2
and one worker. Before and after each capture, the driver authenticates the
clean role checkout, all source/fixture hashes and the immutable binary.
Builds, measurements and compiler gates are serialized under
`/tmp/litchi-goal-0472/cpu.lock`.

`build.py control` and `build.py candidate` retain fresh role build receipts;
`capture.py` retains exact commands, environments, timestamps and artifacts.
Reproduction requires role checkouts from the recorded revisions at the same
absolute build path, both new authenticated binaries, and the bound fixtures.
Run normal A1/B1/B2/A2, then A-full/B-full, A-heap/B-heap, exports and correctness
gates into a fresh bundle. Do not overwrite historical measurements.

The seven-row normal ABBA has 100 samples/five warmups: six ordinary XLSX
commit/save rows and payload-heavy PPT creation. The complete 201-row short
guard uses 15/3. Dense one-percent Heaptrack uses 5/1 and includes whole-process
generation, expected output, warmups, verification and teardown. Its rounded
heap display is not an exact byte count or operation-local peak, and its
instrumented timing/RSS are excluded from normal comparisons. No registered
latency claim is authorized by this diagnostic protocol; the existing
500-sample requirement remains unchanged.

`analyze.py` replays the canonical ABBA and full-guard comparators and retains
individual policy flags, explicit five-percent latency flags, normal RSS pair
deltas and whole-process heap totals. Only full-guard comparison copies treat
identical empty optional source vectors as absent; raw reports are unchanged.
`verify.py --live` checks live role binaries and the currently checked-out
shared tree. Flagless verification requires their absence and verifies the
sealed bundle. Portable replay needs this complete bundle and only
`tools/perf_abba_summary.py` and `tools/perf_compare.py` in the same relative
repository layout; it does not need Rust, Git or the temporary binaries.

Six new differential tests cover eligibility, rich/prefixed fallback, old/new
serialization, attribute/error order, scanner integration and the Option<Tag>
size guard. The initial raw-string fixture and duplicate-error-wording failures
remain in `validation/` alongside the passing 16-test focused run.

No new native Office run, fuzz campaign, arbitrary attribute-rich population
performance, physical-cold/range source, bounded-streaming or worker-scaling
claim follows. Those wider requirements and the non-iWork goal remain open.

The candidate is retained: whole-process allocation calls decrease 14.330%,
while rounded peak heap remains `104.38M`. All six correctness gates and
1,263 XLSX tests pass, as do ten evidence tests and live verification.
The short full guard retains 86 latency flags. An initial strict registry
attempt failed because the temporary filesystem quota was exhausted; its
receipt is preserved and the retry follows removal of authenticated builds.
See `decision.json`, `summary.json`, `validation/` and `cleanup.json`.

The strict registry retry passes all ten existing claims after cleanup.
Portable replay passes from a fresh copy containing this bundle and the two
canonical comparator tools. Temporary build and portable-copy paths are removed;
shared Cargo caches and the two user-owned files are preserved.
