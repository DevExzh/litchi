# 0361g — Required limits on canonical raw text-wire calls

- Date: 2026-10-21.
- Contract: [0351 §2.5](0351-api-contract.md#25-configured-limits-items-18-19).
- Status: accepted naming/required-argument slice; remaining raw audits stay required.
- Base: `2c4c47512b`.

`litchi_iwa_text_wire::from_bytes(source, Limits)` replaces the default-only
call and its `from_bytes_with_limits` twin. The rewrite, validation, decode
and boundary functions use their canonical names with required
`RewriteLimits`. Explicit behavior variants retain `with_behavior` and the
required limits argument; their `and_limits` suffix is removed.

All 30 affected source/consumer files migrate together, including iWork,
Keynote, Numbers, Pages and the independently built Pages text-codec fuzz
target. Production callers keep their existing explicit profiles, including
package-derived policies and the two footnote paths' defaults.
The eight former default-only test calls pass the former default explicitly.
No aliases or deprecated bridges remain. The new `iwa-text-wire` glossary
profile bans both removed suffixes.

The seven implementation bodies match the former explicit-limit bodies after
normalizing the renamed identifiers, whitespace and optional trailing call
commas changed by rustfmt. The retained
[body hashes](results/0361g/body-parity.json) substantiate that this slice
changes neither the wire scan, Buffa projection, allocations, rewrite
algorithm, admission sequence nor output bytes. It removes one default-only
wrapper and introduces no runtime adapter. Existing boundary, malformed-wire,
unknown-field, UTF-16 and exact-replay tests remain the behavioral checks.

The [validation receipt](results/0361g/validation.json) records all ten passing
checks: five-crate all-target/all-feature compilation and Clippy; 48 wire unit
tests and its compile-fail doctest; 2,705 consumer library tests (five existing
ignored tests); 13 Keynote rich-storage integration tests; the separate Pages
text-codec fuzz build; fresh rustdoc and glossary enforcement; inventory; and
formatting. In total, 2,767 tests pass. The inventory emits no rows, but its
heuristics do not establish complete raw-API closure. The implementation-body
check can be reproduced with
`python3 docs/unified-ops/results/0361g/verify_bodies.py`. No fresh native timing
is claimed for this naming-only slice.

The remaining iWork
raw metadata constructors, owner Snappy-policy propagation, detection/legacy
wrappers and other shared/raw API audits remain required. The explicit
`RewriteLimits::default()` calls in `litchi-iwa/src/pages/editor/footnotes.rs`
and `litchi-pages/src/package/body_footnote.rs` remain part of the owner-policy
audit; requiring a raw argument alone does not prove retained-policy use.
The same crate's seven character-style entry points (`char_style_runs`,
`styles_in_span`, `restyle_span`, `character_style_variation`,
`character_style_properties`, `character_style_parent`, and
`stylesheet_with_style`) still need required caller limits and bounded
publication. They currently use a private default wire view and do not
appear as `with_limits` inventory rows. Their migration is required work;
the naming gate is not evidence that the whole raw surface is complete.
This slice does not complete the pending wave integrations or the final
workspace gate.
