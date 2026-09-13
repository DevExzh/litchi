# 0546 adaptive exact-bound scanner design review

This review covers the isolated 0546 follow-up to the rejected 0545
pair-reduction scanner. It evaluates the `counted` helper against the exact
0545 `memchr2` comparator. The review is a diagnostic gate only; it does not
admit a production change or make an end-to-end XLSX claim. OLE2 and OOXML
remain ahead of deferred ODF work.

## Inputs and scope

The experiment is frozen by `freeze.json`. The measured sources are
`baseline.rs` (the exact 0545 comparator) and `counted.rs` (the adaptive
candidate); their frozen hashes are respectively
`47f0cc99f5f5a9b4ea2773474bd868ba688677824c28370674619dc59c08fa5e` and
`64c95876734169f0e7612bbd26dee6e78162d2e52950b3db6bbf52a79b8bcdc5`.
Both functions are
`inline(never)` and are selected through a run-time function pointer in one
release executable. Input loading, the direct-reader oracle, warm-up, result
checks, sample storage, and JSON output are outside the timed call.

The candidate keeps the exact 0544 formula:

```text
B(content) = 1                         // terminal Event::Eof
           + initial_nonmarker_text    // zero or one event
           + count('<' or '&')
           + count('>' or ';' followed by a nonmarker)
```

It changes only how the first marker term is counted. It finds at most 16
markers with the existing `memchr2` skip path. If the source ends before that
probe is exhausted, no bulk scan is performed. Otherwise it counts `<` and `&`
independently in disjoint 64 KiB suffix chunks. `memchr::memchr_iter(...).count()`
uses the single-byte iterator's specialized `count` implementation in the
pinned 2.8.3 dependency; the caller uses only its safe API and adds no
dependency or unsafe code. The `>`/`;` text term remains the original
`memchr2` scan, including its successor check and per-hit cap check.

The ten fixtures cover the five numeric/cap shapes, a sparse one-megabyte
source, an early-rejection source, 16 clustered markers followed by a
one-megabyte text tail, reference-heavy markup, and text delimiter controls.
The numeric/cap screening rule requires both p50 and mean to improve by at
least 20% in both repeats. Each other control must remain within 100,000 ns in
both p50 and mean. Every individual change above 5% is retained for review;
there are no selective reruns. The rule is an isolated promotion screen and
does not replace the original workflow, refusal, allocation, resource, or
native Office gates.

## Correctness proof

Let `M` be the positions whose byte is `<` or `&`. The probe visits the first
`min(16, |M|)` elements of `M` in increasing order and charges one for each.
If `|M| < 16`, the failed `memchr2` search establishes that no unvisited
marker exists and the marker term is complete. If `|M| >= 16`, let `q` be the
position of the sixteenth marker. The suffix `content[q + 1..]` contains every
remaining element of `M` and no probed element. For each suffix chunk, the
bytes `<` and `&` are distinct, so

```text
count('<', chunk) + count('&', chunk)
```

is exactly the number of remaining marker positions in that chunk. The chunks
partition the suffix, hence their checked sum plus the 16 probe charges is
exactly `|M|`. An empty suffix is handled by the empty chunk iterator.

The initial text charge and the `>`/`;` successor term are byte-for-byte the
0545 comparator. Therefore the candidate computes the same mathematical bound
as the comparator. `add` uses checked addition and declines when the cap is
exceeded. The per-chunk subtotal is at most 65,536 because the two counted
byte classes are disjoint, so its local addition cannot overflow; the merged
bound remains checked. Under the production 8 MiB source fence this is also
well below `usize::MAX`, while the checked merge preserves the helper's
standalone overflow behavior.

The direct oracle is configured with the reader settings used by the proof:
no dangling ampersands, no unmatched ends, no comment checking, checked end
names, no empty-element expansion, trimmed closing markup names, and disabled
text trimming. It resolves each borrowed event and asserts the observed event
prefix is no larger than the independent full bound. Tests cover malformed
tails, BOM, declarations, comments, CDATA, doctypes, references, arbitrary
bytes, exact cap/cap-plus-one streams, delimiter/successor pairs at chunk
boundaries, and final-byte cases. This proves the conservative admission
implication for the pinned `quick-xml` call; it does not make the lexical
counter an XML validator.

## Resource and early-exit behavior

The probe retains the baseline's sparse skip behavior while fewer than 16
markers are present: each search jumps directly to the next marker and the
candidate does not scan the suffix. Once the probe is exhausted, the bulk
path may read one 64 KiB chunk after the mathematical cap crossing before it
checks the accumulated subtotal. It still returns before constructing a
provisional reader and cannot admit an over-cap bound. This delayed check is a
performance risk for an adversarial source whose first 16 markers are dense
but whose remaining source is sparse or whose cap is crossed at the beginning
of a chunk. The clustered control explicitly measures the first case, and the
early-reject control measures the cap path.

The candidate borrows the source, has no source-sized allocation, and retains
only scalar locals. These facts do not establish integrated allocation, RSS,
copy, I/O, or workflow behavior. The exact event-bound proof remains tied to
the existing eligibility fences, reader version/configuration, and
authoritative fallback; a future reader setting or dependency change requires
a new review.

## Isolated evidence and promotion decision

Read-only analysis of the 100-sample matched captures gives these p50/mean
changes for repeat 1 and repeat 2:

| Fixture | Repeat 1 | Repeat 2 |
| --- | ---: | ---: |
| numeric 96 | −46.7% / −46.7% | −46.4% / −46.4% |
| numeric 128 | −47.0% / −46.3% | −47.6% / −47.0% |
| cap 160 | −46.7% / −46.5% | −46.8% / −46.7% |
| cap 164 | −50.8% / −50.0% | −50.8% / −50.1% |
| cap 256 | −93.8% / −93.8% | −93.8% / −93.8% |
| sparse | −0.4% / −0.9% | +0.7% / +0.1% |
| early reject | −93.5% / −93.4% | −93.4% / −93.3% |
| clustered | +174.0% / +171.5% | +174.6% / +173.1% |
| references | −46.1% / −46.1% | −46.1% / −46.1% |
| text markers | +3.6% / +3.3% | +2.5% / +2.3% |

The candidate passes the stated isolated screening rule, including the
absolute 100,000 ns control budget, but the clustered source is an explicit
174% relative regression. The absolute control budget prevents that small
fixture's low baseline from vetoing promotion by itself; it must remain an
individually reviewed risk in the integrated campaign. The assembly contains
separate one-byte count calls in the bulk path and retains the `memchr2`
successor loop, so the measured mechanism matches the source hypothesis. No
SIMD claim beyond the pinned dependency's documented/runtime-selected count
implementation follows from this inspection.

This evidence is sufficient to justify a fresh integrated OOXML/XLSX
campaign, because every numeric/cap shape clears the predeclared scanner
screen and sparse/early-reject behavior is bounded and measured. It is not
sufficient to retain the helper. The integrated candidate must be rebuilt from
the accepted 0544 production baseline, preserve the authoritative fallback and
error ordering, and pass every original workflow, allocation, refusal,
resource, no-op, preservation, and native Office gate. Any failure restores
the baseline and retains this diagnostic as attribution evidence.
