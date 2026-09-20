# Change 0701 code review — first-byte MCE marker search

Review status: source, measurement, and retention review pass; scoped
retention recommended.
The frozen candidate codec has SHA-256
`084e65b63c7d2597c262d15401a69442a6b0efc5d927290345899ac1354dbd6b`, and
the candidate focused-test source has SHA-256
`14879f98747f30b78b4eca43a3657a9f74245e172a0f5b5cb71154b54f51df91`.
This review records the measured result without making a broad 0701
performance claim.

## Scope reviewed

The baseline is the restored 0701 production source at revision
`7e4c245673`, whose MCE codec witness is
`f39bc8088825b9a655f4ecbf4c30de9f7e272d574351693be02d69bcfe7f2dc6`. The
retained focused-test source witness is
`764df9cb00f993434e74cf50e2f6fe3b9ed9d3452fe0756c4543c52495a045eb`.

The frozen candidate is one private `#[inline(never)]` helper,
`contains_mce_namespace(&[u8]) -> bool`, and one call-site replacement in
`process_markup_compatibility`. It uses the existing direct `memchr` crate to
find the first byte of `NAMESPACE` only within positions where the complete
URI can fit, then checks `starts_with(NAMESPACE)` and advances one byte after
a miss. The existing `find_bytes` helper and all other MCE paths remain out
of scope. The source diff contains exactly the codec and focused-test paths;
there is no Cargo, lockfile, dependency, or public-API change.

The candidate is source-acceptable only if the final diff has no Cargo or
lockfile change, import or dependency addition, public item, unsafe code,
global state, cache, parser/frame rewrite, limit movement, or test behavior
that changes the production source beyond the documented retained tests.

## Static source findings

The frozen implementation has no semantic blocker under the stated shape.
The valid-start slice is the key safety condition:

```rust
while remaining.len() >= needle.len() {
    let valid_starts = remaining.len() - needle.len() + 1;
    let Some(offset) = memchr::memchr(needle[0], &remaining[..valid_starts]) else {
        return false;
    };
    if remaining[offset..].starts_with(needle) {
        return true;
    }
    remaining = &remaining[offset + 1..];
}
false
```

At loop entry, subtraction is bounded and `valid_starts` includes the final
complete window. `memchr` can return only an offset below that bound, so the
`starts_with` slice has at least the URI length. The update consumes at least
one byte and cannot loop forever. A first-byte hit that lies in the trailing
incomplete region is intentionally outside the search range. The URI is a
non-empty fixed constant, so its first-byte access is backed by a source
invariant; an equivalent checked first-byte extraction is acceptable if used
by the implementation.

The boolean result is equivalent to the existing `.windows(...).any(...)`
predicate for arbitrary bytes. It does not inspect XML structure or decode
UTF-8. Exact bytes in element text, comments, processing-invalid input, and
repeated occurrences retain the old dispatch decision. A false result leaves
the existing output-limit check, borrowed input, and default report in place;
a true result enters exactly the old reader, output, namespace, error, and
report path. The input-limit check remains before the helper.

The `#[inline(never)]` attribute is a layout/measurement control, not a
correctness mechanism. It may reduce caller frame growth from a search helper,
but it adds call overhead and can alter code placement. Assembly, helper
frame, processor frame, native text, and binary size must be inspected after
the candidate is built. No source review can infer those values. The current
source review confirms that the attribute is on the helper and that the
caller contains no inlined search implementation in source; generated-code
claims remain pending the packet's assembly receipts.

## Frozen assembly review

The bounded `nm`/`objdump` receipts now confirm the intended separation while
also showing why the frame result must be described carefully:

| symbol | baseline | candidate | change |
| --- | ---: | ---: | ---: |
| `process_markup_compatibility` | 6,305 B; `sub $0x248,%rsp` | 7,183 B; `sub $0x238,%rsp` | +878 B; caller reservation −16 B |
| `contains_mce_namespace` | absent | 170 B (`0xaa`) | new outlined helper |

The candidate caller performs the existing input-limit branch, moves the XML
pointer and length into the helper argument registers, calls
`contains_mce_namespace` at `0x170580` from `0x1706d1`, tests `%al`, and takes
the existing borrowed fast path when the result is false. The parser setup and
all later processing remain after that branch. The helper begins with a
`0x3b` (59-byte) length check, so short input returns false before computing a
valid-start range.

Within the helper, the disassembly shows the intended two-stage search. It
computes the inclusive end of the valid-start region (`remaining + len - 58`)
and dispatches an indirect `memchr` implementation with the first byte
`0x68` (`'h'`), the current pointer, and that end bound. When a candidate
pointer is returned, it computes the relative offset and dispatches the
existing `bcmp` symbol for exactly `0x3b` bytes against the URI literal. A
failed comparison advances the current pointer by the candidate offset plus
one and repeats; a matching comparison returns true. The no-hit and short
branches return false. Thus the generated code does not contain the baseline's
every-window loop in the caller, and it does not call the 0700
`FinderBuilder`/`memmem` searcher.

The baseline assembly directly loops in `process_markup_compatibility`,
loading `bcmp`, comparing 59 bytes, incrementing the window pointer, and
continuing while at least one complete window remains. The candidate retains
one 59-byte `bcmp` per first-byte candidate, but moves the first-byte scan to
the helper's indirect `memchr` dispatch. This matches the source algorithm;
it does not by itself establish a latency benefit.

The caller's reserved stack area decreases by 16 bytes, but the helper saves
six callee-saved registers plus one stack slot (`48 + 8 = 56` bytes) before
its return address and nested search/`bcmp` calls. Therefore the caller-frame
delta is not a peak call-chain or process-stack bound. The packet must report
the 16-byte caller reservation change and the helper's extra frame separately;
it must make no whole-program peak-stack saving claim without a bounded
call-chain measurement.

## Initial measurement assessment (provisional)

The first complete semantic and focused control evidence changes the 0700
tradeoff, but it does not settle retention. The 1,920-case shared oracle
(192 inputs × 5 profiles × 2 binaries) reports zero mismatches, and the
independent allocation comparison is identical. The focused suite has 104
passing tests on both source states.

The marker controls now show the intended setup behavior. The tiny marked
case improves by 8.51% and 7.45% at the two primary p50 pairings (about 40
and 35 ns), and the root-comment case improves by 4.08%. The long
near-prefix marker-free case improves by 87.72% and 85.67%, while the late
comment hit improves by 97.24% and 97.46%. Tiny marker-free p50 remains 30 ns
on both sides; its means differ by roughly 1.5 ns and produce small review
triggers, so they remain recorded rather than dismissed.

The initial 13-workflow native matrix has recurring real-work costs below the
goal's approximate 5% trigger: `one-real` is +1.10%/+1.14%, `noop-real` is
+1.64%/+2.79%, and `two-real` is +1.72%/+1.62% in the primary balanced
pairings. The other ten workflows improve between roughly 10.3% and 15.0%.
The native p99/phase trigger list is retained; in particular, `noop-real`
has a +6.12% p99 total trigger in one pairing. The refusal matrix improves at
the p50 in all ten cases, including the MCE-positive cases, and the
16-to-64 MiB over-limit case improves by about 99.27%.

The shared XML controls are mostly within a few percent. Declaration controls
improve by roughly 0.8–2.6% at p50, while mixed opaque controls cost about
2.1–3.6%. The only shared p99 review triggers above 5% are the DOCX opaque
case (+8.20%) and DOCX many-opaque case (+25.26%); their individual receipts
remain visible. These controls do not replace the native Office workflows.

This is a materially stronger result than 0700: the earlier
approximately 45% tiny-marked penalty is gone, the broad mechanism and
refusal improvements are substantial, and the source change reduces the
caller reservation while avoiding `FinderBuilder`. Against that, the common
real edit/no-op/two-target workflows are consistently slower by about 1–3%,
the processor symbol grows 878 bytes, and the helper adds a nested frame.
The candidate therefore required an explicit final retention review. The
seven integration and six evidence gates plus the longer post-gate follow-up
are recorded below; they establish the measured benefit and its limits rather
than authorizing an unconditional production performance claim.

## Focused-test coverage review

The three new focused tests add 104 total `mce::` tests from the restored
101-test baseline. Their coverage is sufficient for this lexical predicate:

- every insertion offset from 0 through 128, with zero through three trailing
  bytes, covers offset zero, the final complete window, short trailing
  regions, and arbitrary `0xff` surroundings;
- every byte substitution at every URI position covers first-, middle-, and
  final-byte near misses while retaining one exact case per position;
- repeated `h` candidates, late mismatches, all trailing URI prefixes, and
  one or two subsequent exact markers exercise advance-after-miss and
  repeated-occurrence behavior;
- the retained 0700 tests continue to cover short inputs, arbitrary bytes,
  input-before-output limit precedence, text/comment placement, truncated
  XML, and malformed UTF-8 attributes.

The helper in the test file compares the production dispatch outcome with
the baseline `windows` predicate. Marker-free cases require borrowed bytes
and a default report; marker-present cases require the owned parser route or
an error. The shared oracle remains responsible for exact transformed output,
reports, and typed/debug error identity. The tests use the current URI
literal, which matches `NAMESPACE` byte-for-byte in this source; that match
must remain bound by the candidate source and oracle receipts.

I found no missing semantic case that blocks the candidate. The repeated
candidate test is an advance/control test rather than a claim that the URI
has a self-overlapping exact occurrence; the all-offset and oracle cases
cover the actual boolean boundary. No test changes are requested from this
review.

## Resource and ownership review

The helper has constant explicit storage and performs no fallible allocation.
Its scan remains bounded by the already admitted `xml` slice and therefore
does not change source-byte, output-byte, XML-depth, namespace, or allocation
limits. It retains no pointer, byte owner, cache entry, lock, execution
context, or source identity after return. The existing `memchr` dependency is
already declared directly by `litchi-ooxml-common`; adding no dependency edge
preserves ADR 0002 and ADR 0024 topology.

No snapshot, edit, commit, patch, archive, relationship, physical package,
or facade contract is reached by this helper. The codec continues to borrow
the caller's marker-free bytes and to produce owned transformed bytes only
through the existing parser path. No-op sharing, typed refusal ordering,
determinism, and preservation remain downstream invariants that the
differential evidence must verify.

## Cost review

For URI length `m`, the helper's logical work is `O(n + k*m)`, where `k` is
the number of valid positions containing the URI's first byte. Its worst case
is `O(n*m)` when repeated near-prefixes force a full comparison at many
positions. That is the central risk: `memchr` avoids impossible starts, but
it does not provide the skip-ahead behavior of a general substring search.
The repeated-prefix marker control is therefore mandatory and must remain
separate from ordinary Office workflows.

The expected mechanism is narrower setup on short inputs and marker-bearing
inputs, plus efficient skipping when the first URI byte is rare. Both are
hypotheses. The candidate can regress tiny calls through the outlined helper
call itself, and it can regress hostile first-byte-dense inputs through the
bounded repeated `starts_with` comparisons. The three real workflows also
need fresh balanced measurements because 0700's direct `memmem` regressions
cannot be attributed to this different implementation.

The review will treat these as first-class outcomes:

- tiny marker-free and tiny marked controls;
- exact early, middle, and late marker positions;
- long marker-free and late-marker inputs;
- repeated-prefix and overlap-heavy near misses;
- marker-positive MCE work, including refusal cases;
- real one-target, no-op, and two-target workflows;
- processor code/stack and allocation effects.

## Required final evidence (completion checklist)

The final packet binds a fresh baseline and candidate source map, lockfiles,
build inputs, corpus hashes, environment, and binary hashes. It runs the
focused MCE tests against both source states and passes the exact shared
oracle (192 cases × 5 profiles × 2 binaries) for output bytes/length, `Cow`
ownership, complete `Report`, and typed/debug errors.

The 0700 control family and full gates are reused only as harnesses; every
0701 result has a fresh receipt. The completed measurement set is:

- marker controls with retained distributions and all individual triggers;
- 13 native workflows with balanced process legs and warmups;
- 10 refusal cases with the same balanced discipline;
- independent allocation and profile diagnostics;
- assembly, native size, and processor stack review;
- all seven integration and six evidence gate rows;
- the post-gate six-case native and ten-case refusal follow-up with its
  independent audit.

No one leg, microbenchmark, geometric mean, or favorable synthetic case can
authorize retention. The host's shared/warm state and any unavailable tools
must remain explicit in the packet, and no cold-cache, native Office, remote
I/O, or parallel-scaling claim may be added to this local experiment.

## Initial conditional disposition criteria

The initial source review passes the candidate to measurement with these
conditions:

1. Exact source scope and the loop invariants above remain true after
   formatting.
2. Focused, oracle, integration, evidence, and follow-up checks pass without
   weakening limits or semantic comparisons.
3. Representative real edit/no-op/two-target workflows show a statistically
   useful result, and every approximately greater-than-5% regression called
   out by [`docs/GOAL.md`](../../../GOAL.md) receives an explicit disposition.
4. Repeated-prefix, tiny marked, and MCE-positive costs are acceptable in
   context; no material cost is hidden by aggregate statistics.
5. Any code-size or stack change is measured and justified by end-to-end
   benefit rather than by the helper's local scan alone.

If these conditions had failed, the candidate would have been rejected and
the restored codec retained. The packet preserves the candidate source
witness and the focused tests that document the lexical dispatch boundary.
The final measurement review below evaluates these conditions and authorizes
no broad production performance claim.

## ADR compliance

The static review found no conflict with the accepted constraints consulted
before editing. The helper remains in the existing `litchi-ooxml-common` MCE
owner (ADRs 0001, 0002, 0010, 0011, and 0024), changes no public or physical
package boundary, and adds no dependency. It does not alter snapshot/edit/
patch publication or exact no-op behavior (ADR 0003). It performs a bounded,
allocation-free byte scan under the existing source admission and execution
policy, with no cache, worker, ambient I/O, or global pool (ADRs 0005 and
0031). It preserves the existing validation and preservation route, including
malformed-input refusal and unknown-byte handling (ADR 0006). The fresh
baseline/candidate, differential, adversarial, and full-gate protocol follows
ADR 0008. No ADR amendment or proposed record is required.

## Final retention assessment

The frozen candidate satisfies the review conditions for this scoped
experiment. Both source states pass the 104-test focused suite, the shared
1,920-invocation oracle reports zero mismatches, all seven integration gates
and six evidence gates pass, and the independent follow-up audit passes 24
native rows, 40 refusal rows, and 138 comparisons. Allocation metrics remain
identical across the retained comparison set. The helper's exact lexical
predicate, limit ordering, borrowed marker-free result, parser/error route,
and report behavior therefore have the required differential evidence.

The longer native follow-up gives the useful mechanism signal without
supporting a real-deck gain claim: one-control, generated, and notes-bearing
cases improve by 11.28% to 14.77% at the reported p50 pairings. The one-edit
real case changes by +0.07%/+0.68%, and the two-edit real case by
-0.61%/+0.34%. The initial real-work costs remain part of the decision:
one-real +1.10%/+1.14%, no-op real +1.64%/+2.79%, and two-real
+1.72%/+1.62%. The first longer no-op comparison of -7.14% is confounded by
-6.87% baseline drift; its second comparison is -0.27%, so no no-op speedup
is claimed.

All ten longer refusal medians improve, including the MCE-positive cases at
-1.97% to -5.78%; early-name refusal improves about 48% and the padded
over-limit refusal about 99.28%. These are bounded refusal-path results, not
a claim about every parser or Office workload. The initial shared DOCX
opaque p99 flags (+8.20% and +25.26%) remain recorded. The follow-up retains
all 17 trigger rows, including the three primary clone-p99 flags: generated
+50.39% (+6.45 microseconds), notes-bearing +43.31% (+2.75 microseconds),
and one-real +7.50% (+1.68 microseconds). Their small absolute durations do
not erase the review triggers.

The generated-code and resource costs remain explicit: native text grows by
3,096 bytes, the processor symbol by 878 bytes, and the helper is a 170-byte
outlined function with a 56-byte saved-register/slot frame on its full path.
The caller reservation decreases by 16 bytes, but that is not a peak
call-chain or process-stack saving. Diagnostic RSS changes from 5,652 to
5,764 KiB, while the allocation comparison is identical. These costs are
accepted only because the measured marker-free/refusal improvements are
substantial on the retained corpus and the common real-work p50 cost stays
within the review context; they remain limitations of the result.

Retain the private `memchr` plus exact-prefix helper and its single predicate
replacement, with the three focused regression tests, in the shared MCE
codec. The retention is scoped to this predicate, corpus, host, and measured
workflow mix. It does not claim a universal MCE speedup, real Office deck or
CRUD speedup, memory reduction, full-save throughput gain, cold-cache result,
remote-I/O result, or parallel-scaling result.
