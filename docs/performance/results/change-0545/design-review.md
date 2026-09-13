# 0545 exact-544 scanner design review

This review covers the isolated follow-up to the rejected 0544 two-scan event
bound. It evaluates the exact-544 formula and its 4,096-byte chunked scalar
implementation against the frozen 0544 `memchr` helper. It is a diagnostic
precondition only. It does not admit a runtime change or claim an XLSX
workflow improvement.

## Inputs and scope

The review read `docs/GOAL.md`, `docs/adr/README.md`, ADRs 0001, 0003, 0005,
0006, 0008, and 0024, and the frozen 0544 source and design records. The
relevant accepted constraints are bounded source/state, validation and
preservation precedence, safe Rust, explicit measured evidence, and no
unsupported end-to-end claim. OLE2 and OOXML remain ahead of deferred ODF
work.

The frozen diagnostic inputs are in this directory:

| Input | SHA-256 |
| --- | --- |
| `baseline.rs` | `47f0cc99f5f5a9b4ea2773474bd868ba688677824c28370674619dc59c08fa5e` |
| `chunked.rs` | `5307ddf85619fc803ac3c6d00d65d2f55f4eec3db44a7d97d6dfee4e77298bd2` |
| `main.rs` | `bfc732129784ad94332836fcae4a505f89a9aa77a23d130e16206fb5b30cb5e5` |
| `Cargo.lock` | `58c2b810de06b6f92fbf00ed86635767babea1e8973b8fd9b399b94299996e30` |
| `freeze.json` | `770fea53337d068b804a2512dffb98524f4ee57c42f6fac5d72931f54d8357b9` |

The baseline is the 0544 helper with the same checked counter and two
`memchr2_iter` passes. The candidate uses borrowed `split_last` and paired
`chunks(4096)` slices. The binary runs ten warm-up iterations and 100 samples
per fixture, pinned to CPU 2, in the frozen order baseline-1, chunked-1,
chunked-2, baseline-2. There are 28 isolated captures: seven fixtures and
four process roles. The quality receipt reports two tests, formatting, build,
and warning-denied Clippy passing. No build or test was run as part of this
independent review; those receipts and the raw captures are the evidence.

## Static verdict

The exact-544 formula is sound for the pinned reader call, and the chunked
implementation is extensionally equivalent to the frozen baseline for the
fixed cap. The implementation has no heap allocation and its checked
per-chunk accumulation is safe under the fixed 4,096-byte chunk invariant.

The performance hypothesis does not hold across the relevant input shapes.
The chunked loop is substantially faster on the dense synthetic grids, but it
takes 42.98–44.18 times the baseline time on the one-megabyte sparse source and is 22.6–22.8%
slower on the early-reject source. The captured release assembly contains a
scalar byte loop and conditional jumps at the `chunked` symbol; it does not
show the proposed vector reduction. These are isolated lexical predicate
timings, not workflow timings. The current implementation must remain a
rejected diagnostic and must not be integrated into the production 0544 path.

## Soundness of the bound

Let `M` be the positions whose byte is `<` or `&`, and let `T` be the
positions before the final byte whose byte is `>` or `;` and whose successor
is neither `<` nor `&`. The proposed bound is:

```text
B(content) = 1                         // one terminal Event::Eof
           + initial_nonmarker_text    // zero or one
           + |M|
           + |T|
```

For `quick-xml 0.41.0` using `NsReader::from_reader(&[u8])` and
`read_event()`:

* Every non-text, non-EOF event starts at a `<` or `&`: `Start`, `End`,
  `Empty`, `Comment`, `CData`, `DocType`, `Decl`, and `PI` start at `<`, while
  `GeneralRef` starts at `&`.
* In the initial state, a non-marker prefix can produce at most one `Text`
  event. The reader strips a UTF-8 BOM before entering that state; charging a
  non-marker BOM is therefore conservative.
* After markup, the reader enters text state. It emits at most one contiguous
  text event before the next `<` or `&`. The byte after a markup terminator
  `>` is charged exactly when it is not a marker.
* After a completed general reference, the reader again enters text state.
  The byte after its `;` is charged under the same condition. A semicolon in
  an attribute, comment, CDATA section, ordinary text, or malformed input only
  adds a false-positive charge.
* Namespace resolution observes or borrows the event; it does not create a
  second event. Reader errors stop the observed event prefix and cannot add an
  event that the formula failed to charge. EOF contributes one final event.

Thus every emitted event prefix is at most `B(content)`, which gives the
required admission implication:

```text
actual emitted events > 131,072  =>  B(content) > 131,072
```

The proof is tied to the current reader configuration. In particular,
`allow_dangling_amp` must stay `false`: with it enabled, text returned for an
unterminated reference can begin after `&` without a preceding `;`. Likewise,
`expand_empty_elements` must stay `false`: enabling it synthesizes an `End`
event for a self-closing tag without another source `<`. The current
`trim_text_start` and `trim_text_end` settings must remain `false`; the
production path explicitly calls `trim_text(false)`. The other pinned defaults
are `allow_unmatched_ends = false`, `check_comments = false`,
`check_end_names = true`, and `trim_markup_names_in_closing_tags = true`.
Changing the reader call, dependency version, feature set, or any of these
settings requires a new proof review.

The source eligibility and MCE/x14ac fences remain separate correctness
requirements. The lexical predicate does not validate XML, bypass
authoritative validation, or replace the runtime event/depth/resource caps.

## Chunk and arithmetic review

For an empty slice, `chunked` returns the baseline result `true`. For a
non-empty slice it initializes the bound with EOF, the initial non-marker
charge, and the last-byte marker charge. It then processes every position
except the last. If `C = 4096`, the `k`th pair is taken from

```text
prefix.chunks(C)[k]       // content[k*C ..]
content[1..].chunks(C)[k] // content[1 + k*C ..]
```

so the zipped elements are exactly `(content[i], content[i + 1])` for every
`0 <= i < content.len() - 1`. The pair at a chunk boundary, including
`i = C - 1`, is in the preceding chunk; no pair is lost or duplicated. The
last marker is counted once outside the pair reduction, and no text event can
start after EOF.

Each pair contributes at most two. Therefore the unchecked inner `subtotal +=`
operations cannot overflow with `C = 4096` (`subtotal <= 8192` before the
checked merge). The merge uses `bound.checked_add(subtotal)` and declines on
overflow or when the cap is exceeded. This is safe under the frozen constants,
but the invariant is implicit: changing the chunk size or contribution width
must either restore checked subtotal additions or add an equivalent proof.
The small initial `1 + initial + last` expression is also bounded by three.

Checking only after a chunk means an over-cap source can be scanned through
the rest of the current chunk after its mathematical bound first exceeds the
cap. The extra work is bounded by fewer than 4,096 pair positions. Counting a
final marker up front can instead make the chunked predicate decline before
the physical position that makes the baseline predicate decline. Both return
the same Boolean result; these different stopping points are material to the
early-reject timing and do not change safety.

Both functions borrow the input and keep only scalar locals. The benchmark's
argument, input, sample vector, and JSON output allocations are outside the
timed function. No allocator, source-copy, or I/O claim follows from the
scanner itself.

## Edge-oracle and test review

The diagnostic tests independently calculate `B`, compare both predicates,
exercise BOM, declarations, comments, CDATA,
doctype, attributes, references, malformed tails, arbitrary bytes, and sizes
around 4,096 and 8,192. The repeated-comment cases also cross the 131,072
predicate boundary. The main capture checks both implementations on the
96/128/160/164/256 grid fixtures, a sparse source, and an early-reject source.

Before any runtime admission, strengthen or retain the following checks:

1. Configure every reader field used by the proof in the direct oracle,
   rather than setting only `check_end_names`; this prevents a dependency
   default or harness change from silently invalidating the implication.
2. Include an exact-cap valid stream that exercises both a markup/reference
   charge and a following text charge. Repeated adjacent comments exercise
   marker counting but do not exercise the two sides of the pair formula.
3. Add explicit pair fixtures with a `>` and a `;` at offsets `4095`, `4096`,
   `8191`, and `8192`, with marker and non-marker successors, plus final-byte
   `<`, `&`, `>`, and `;` cases. The current padded atoms cover many lengths,
   but do not make each of these delimiter/successor combinations obvious.
4. Keep a direct-reader error prefix in the oracle and assert the observed
   prefix is bounded. For representative valid cases, assert the expected
   event kinds/count as well; `while let Ok(...)` is sufficient for the
   one-way safety check but can conceal an unintended early parser error.
5. Keep a checked reference counter or a parameterized small-cap test for
   arithmetic boundaries. The fixed 8 MiB production source fence prevents
   `usize` overflow in practice, while the standalone helper itself has no
   source-size fence.
6. Treat `sparse` and `early-reject` as mandatory performance controls. A
   replacement that wins only dense grids is not a safe general admission
   helper; the two controls expose the cost of scanning every byte and of
   deferring the cap check to chunk completion.

These are diagnostic requirements, not reasons to alter the frozen source in
this batch.

## Matched isolated measurements

The following p50 values are in nanoseconds. Each row uses the same frozen
fixture and has 100 observations in each of the two ABBA repeats; percentages
are `(chunked / baseline - 1) * 100`.

| Fixture | Baseline r1 / r2 | Chunked r1 / r2 | p50 change r1 / r2 |
| --- | ---: | ---: | ---: |
| grid-96 | 297,961 / 304,691 | 168,001 / 168,011 | −43.616% / −44.859% |
| grid-128 | 538,192.5 / 533,622 | 305,916.5 / 305,931 | −43.159% / −42.669% |
| grid-160 | 811,663.5 / 836,433 | 487,212 / 486,897 | −39.974% / −41.789% |
| grid-164 | 818,378 / 793,388 | 499,447 / 498,007 | −38.971% / −37.230% |
| grid-256 | 553,737 / 549,442 | 492,792 / 493,152 | −11.006% / −10.245% |
| sparse | 15,950 / 16,400 | 704,642 / 704,917.5 | +4,317.818% / +4,198.277% |
| early-reject | 582,852.5 / 583,592.5 | 715,333 / 715,517 | +22.730% / +22.606% |

The dense fixtures are generated or extracted worksheet XML, with the
160/164/256 files retaining the 0544 cap-boundary provenance. `sparse` is a
one-megabyte text body with few delimiters. `early-reject` begins with
131,073 repeated comments and appends a one-megabyte text body. These are
lexical controls and do not represent complete Office workflows.

The matched results establish a real tradeoff: the chunked loop wins on
dense, marker-rich inputs but loses badly when the baseline's `memchr` search
can skip long non-marker spans or stop as soon as the cap is crossed. The
assembly receipt further shows that the source-level Boolean conversions did
not produce the intended vectorized reduction on this compiler and CPU. No
unmeasured claim should be made that a different compiler, CPU, or Boolean
spelling will fix that result.

## Disposition and next gate

The exact-544 formula is suitable for a bounded isolated diagnostic and the
proof is adequate under the pinned reader defaults. The current chunked
implementation is rejected for production and for a full workflow campaign.
The rejected 0544 `memchr` predicate is only the diagnostic comparator;
production remains restored to the pre-0544 implementation.

A future candidate would need a new design review and fresh source freeze. It
must protect sparse and early-reject inputs, explicitly pin the reader
configuration, preserve checked arithmetic and authoritative fallback, and
show release assembly or other direct evidence for any vectorization claim.
It must pass the existing semantic, resource, no-op, native, allocation, and
quality gates before any runtime admission. The next scalar/branchless idea is
unmeasured and remains a proposal only.
