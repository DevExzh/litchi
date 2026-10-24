# 0362za — Explicit limits for font preparation and weighted cache

The final direct-library API audit found two non-substrate public
`*_with_limits` entries. Both now expose one canonical raw constructor/helper
with its limits required, as required by API contract 0351 §2.5:

| Before | After |
| --- | --- |
| `litchi_fonts::embedding::prepare_with(resolver, glyphs, mode)` and `prepare_with_limits(..., limits)` | `prepare_with(resolver, glyphs, mode, limits)` |
| `WeightedCache::new(max_weight)` and `new_with_limits(max_weight, max_flights)` | `new(max_weight, max_flights)` |

The default `prepare` system-font workflow and IWA archive cache pass their
previous defaults explicitly. Existing bounded callers retain their exact
profiles. The old names are removed without aliases or deprecations, and the
font obfuscation error is marked non-exhaustive so downstream matches retain a
typed future fallback. All font, cache, and archive tests were migrated in the
same change.

Validation:

- `cargo test --offline --locked -p litchi-fonts --all-features` — 25 passed.
- `cargo test --offline --locked -p litchi-iwa-cache` — 10 passed.
- `cargo test --offline --locked -p litchi-iwa-archive --lib` — 181 passed.
