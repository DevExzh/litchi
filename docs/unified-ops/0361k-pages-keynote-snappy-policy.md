# 0361k — Pages and Keynote retained Snappy policy

- Date: 2026-10-22.
- Scope: Pages and Keynote owner publication paths in this slice.
- Status: source migration, focused tests and `-D warnings` Clippy pass; no
  native timing is claimed here.

The changed production writers now derive `SnappyLimits` from the retained
archive/source policy before decoding, computing a compressed bound, or
compressing a rewritten component. This covers package metadata, settings,
identity, slide order/build order, notes, backgrounds, media, chart and table
components, placeholder visibility, transitions, audio creation, and the
Pages section/table/layout writers in the slice. The same retained profile is
used for the bound and the actual encoder call, so a caller-selected stream
ceiling cannot be replaced by a default during publication. Existing typed
limit and allocation errors continue through the owner mappings.

Two Keynote chart planners intentionally retain an uncapped default profile
for the encoder's temporary per-frame scratch bound:
`slide_chart_caption::snappy_output_plan` and
`slide_chart_value_axis::snappy_output_plan`. That scratch allocation occurs
before the encoder knows the actual frame length and is independent of the
stream output ceiling. Their stream bound and the final
`SnappyStream::compress` call use the retained `SnappyLimits`; the tests cover
a reduced configured chunk and a one-byte-tight output cap, including
identical frame output and round-trip behavior.

Fixture, verification, and performance-test constructors still use explicit
default profiles to create independent test inputs. The Pages footnote owner
is carried by the separate retained-policy slice (`3b31195c1d` plus its
aggregate-budget repairs and validation), and Numbers is being migrated by
W4b; neither is claimed by this record.

Validation on the dirty policy source passed:

```text
source .coord/w8-snappy-policy/env.sh
CARGO_BUILD_JOBS=1 cargo test --offline --locked -p litchi-pages --lib --all-features
CARGO_BUILD_JOBS=1 cargo test --offline --locked -p litchi-keynote --lib --all-features
CARGO_BUILD_JOBS=1 cargo clippy --offline --locked \
  -p litchi-pages -p litchi-keynote --all-features --lib -- -D warnings
```

The Pages suite ran 209 tests; Keynote ran 330 tests with one existing
ignored test. Clippy emitted no warnings. The source changes are ready for
coordinator integration after the separate Numbers and footnote commits.
