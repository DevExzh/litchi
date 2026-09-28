# 0803 source review

## Scope

This review covers the frozen source-only 0803 diagnostic in
`candidate/after/`, its `candidate/before/` control, and
`candidate.patch`. The control is the exact rejected 0802 candidate source
(`change-0802/candidate/after`), as recorded by the 0803 manifest; it is not
the current production source at `a629dceb73`. No production file is changed
or adopted by this packet. I did not run a build, test, native capture,
profiler, or workflow measurement.

The archive-relative patch applies cleanly with `patch --dry-run` to
`change-0803/candidate/before` and has five helper-file hunks plus the required
after-leg amendment to the shared empty-tail test. The five before helper
sources are byte-identical to the 0802 candidate/after helper sources. All
other shared test cases remain unchanged; the one amended test records the
intentional private-state transition and is not an implementation change.

## Single-factor change

Each of the five helper copies moves one conditional. `CheckedAttributes::new`
still constructs the unchecked quick-xml iterator, but it now always stores
`Phase::First`. At the start of `next_first`, the same raw-tail length check
sets `Phase::Done` and returns `None` when the tag has no bytes after its name.
For every other tail, `next_first` continues to call the same unchecked
iterator and applies the same state transition as the 0802 control.

The linear array, ordered map, `key_at` and `name_at` lexing, duplicate
preflight, error remapping, offsets, borrowed names, `end_of`, and all later
phase transitions are unchanged. No backend is allocated or seeded by this
move, no attribute is replayed, and no parser check is re-enabled. The
constructor still creates the `Attributes` value, so this diagnostic isolates
the placement of the exact-empty branch rather than removing construction
work.

## Lexer and error-boundary parity

The check remains an exact raw-tail test: `tag.as_ref().len() ==
tag.name().as_ref().len()`. A whitespace tail, malformed tail, missing
separator, or any other non-empty tail still reaches quick-xml's unchecked
first-item lexer. The first item has no earlier name, and the existing
`Second`, `ShortTwo`, linear, and ordered phases retain their raw-key
preflights. In particular, duplicate names are still reported before a
malformed value where the control reports them, while missing-equals and
other lexical errors retain quick-xml's precedence and positions.

The move therefore preserves the iterator's item/error sequence for the
empty, whitespace, malformed, duplicate, and long-tag cases represented by
the shared tests. It does not alter the bounded linear prefix or the ordered
map handoff after the 33rd successful name.

## Clone and fused behavior

The exact-empty iterator has a deliberate private-state difference before its
first request. In the 0802 control, construction leaves it in `Done`; cloning
before the first request copies `Done`. In 0803, construction and a clone
before the first request are both in `First`; the first `next` on either copy
performs the raw-tail check, stores `Done`, and returns `None`. Subsequent
calls on both copies continue to return `None`.

Thus clone results, first returned item, and fused behavior remain equivalent
for iterator consumers, while derived `Debug` output before the first request
can differ (`Done` versus `First`). `CheckedAttributes` derives `Debug`, so
that private phase representation is observable when a caller formats the
iterator; it is an intentional state-layout difference in this diagnostic,
not a claim of byte-for-byte debug-output parity. Clones made after the first
request, after an error, or after another terminal transition follow the same
state as the control.

The initial 0803 after quality run used the unchanged empty-tail test and
failed its `Phase::Done` assertion before the first request, while the other
helper tests passed. The final after archive amends only that test: it asserts
`First` before the request, `Done` after the first `None`, and checks clone and
repeated-`None` exhaustion. This is a test update required to observe the
intentional state-placement change; it does not broaden the implementation
factor or alter the lexer oracle. The amended after quality receipt passes,
while the failed attempt remains retained under `quality-failed-0/`.

## Source disposition

I found no lexer, first-error, ownership, bounds, clone-result, or fusion
correctness blocker in the code-only change. The archive is suitable for the
controlled before/after diagnostic with the explicit after-test amendment
that captures the private-state transition. This source review provides no
production-workflow speedup, resource, cross-format, or adoption claim; the
0802 rejection and no-adoption boundary remain in force.
