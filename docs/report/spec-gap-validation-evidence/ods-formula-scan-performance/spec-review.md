# ODS formula scan performance review

Status: bounded review complete; no blocking finding remains in the immutable
legacy sheet lookahead cache or the adjacent reference-parser annotations.
The review covers identifier and sheet lookahead work only. It does not certify
the full formula parser as linear in time or memory.

## Reviewed snapshot

The baseline is commit `5fae21a34e163a3300fc0ac4f9c302f5731e3cbb`. The frozen
candidate source digests are SHA-256 values:

| Path | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula.rs` | `8769d998a9c15299508892bb00ec7a2a52f6368bdb21ce4c0f252ddfdf37e8e6` |
| `crates/litchi-ods/src/codec/formula/functions.rs` | `1d736271e9f9cf3743295f8895f184e32842640eaba48084dfafc9ca7333e2a1` |
| `crates/litchi-ods/src/codec/formula/reference.rs` | `3707b3176e2714aa794bf28628eb27ec72923e001faaab22a987883b596bd27e` |
| `crates/litchi-ods/tests/ods_formula_scan_regression.rs` | `07f03b83546dd4b705637271d83bfa2449c6e841499d8ae2b7939455ca4a1e15` |

`reference.rs` is byte-for-byte unchanged from the baseline. Its existing
`#[inline]` annotations on component copying and bumping and its `#[cold]`
limit-error annotation therefore have no new semantic surface in this batch.
The paired A/B/A/B measurement in
[`performance/report.md`](performance/report.md) found no material change on
the long bracketed-sheet lanes (within 0.4% median); the over-limit IRI shift
was within the measured process spread. No annotation-related correctness or
release blocker was found. Whether to retain those hints remains a separate
performance-policy choice.

## Cache correctness and work bound

`FormulaParser::legacy_sheet_scan_end` (currently around lines 773–797) caches
the end of the same byte run that the preexisting legacy scanner consumed:
`[A-Za-z0-9_ ]+`. For an immutable input run `[a,b)`, every start `s` with
`a <= s < b` has the same first non-run byte `b`; returning the cached `b`
therefore preserves the old endpoint and all token boundaries. A start at `b`
is deliberately a cache miss and performs the old zero-byte scan. A start
before the cached interval rescans the prefix and replaces the interval, which
is required for the compact-to-legacy rollback case.

The parser borrows immutable formula bytes for its lifetime, initializes the
cache before parsing, and changes to the trimmed formula body before the first
lookahead. No mutable-input or stale-cache path is exposed. The cache stores
one fixed-size tuple and does not grow with formula length or token count. The
existing input-byte admission and pre-token token-count admission remain in
`parse_with_limits`.

The only backward movement relevant to this cache is the compact candidate's
syntax miss followed by the legacy parser retrying the candidate's earlier
start. Thus a run can first be scanned from the later handoff point and then,
at most, once more from that earlier candidate start; subsequent starts move
forward and hit the interval. The compact scanner itself is bounded by
`MAX_STANDARD_FUNCTION_NAME_BYTES = 19`, and `parse_compact_cell_parts` receives
no larger slice, so repeated compact misses add a constant amount of work per
token. The scoped identifier/sheet lookahead is consequently linear in the
input bytes, including contiguous `A1A1...`, space-separated partial cells,
range endpoints, and the one backward handoff.

This argument is deliberately scoped. Other parser paths, including literal
handling that reserves based on remaining input, are outside this cache review;
the change must not be described as a whole-parser linear-memory guarantee.

## Compatibility cases checked

The cache predicate and call sites preserve the established lexical behavior:

* Literal spaces remain eligible for the legacy unquoted sheet handoff, while
  tabs, newlines, dots, and operators terminate that scan exactly as before.
  `A1 Sheet.B2`, `A1  Sheet.B2`, and `Sheet1 .A1` retain their sheet-qualified
  token shape; tab-separated forms retain independent/current-sheet tokens.
* Contiguous partial cells such as `A1A1A1A1` still produce the established
  sequence of partial cell references. A dot after the run still hands the
  complete prefix to the legacy sheet parser.
* Absolute cells bypass this cache and retain both absolute flags. Range
  endpoints invoke the same legacy parser with a later position, so ordinary
  and absolute ranges retain their prior endpoint semantics.
* Malformed suffixes (`A1_A2`, `A1.2`, `A1..B2`, missing columns after a dot,
  and trailing malformed range text) retain refusal behavior because the cache
  changes only the endpoint of an identical legacy character run.
* Formula prefixes are stripped only for scanning and the original formula
  text remains retained.

The private work-counter unit test covers 1,024 space-separated cells,
contiguous cells, ranges, absolute cells, and the compact-to-legacy rollback;
the focused integration test covers the semantic cases above, long fallback
sheet names, whitespace boundaries, and malformed suffixes.

## Validation receipt

The focused formula receipt records 49 unit tests passing, including
`test_legacy_sheet_scan_cache_keeps_scan_work_linear`. The focused
`ods_formula_scan_regression` receipt records 6 tests passing, with formatting
and diff checks successful. The frozen package gate receipt at
[`gates/results.json`](gates/results.json) records status zero for test,
clippy, documentation, doctest, and formatting gates: 822 tests across 49
targets, with all nine reviewed source hashes stable. The crate-boundary check
also records status zero.

Disposition: no additional blocker was found in this bounded performance
review.
