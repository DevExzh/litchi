# Numbers tile-storage codec fuzzing

## Numbers plain-number cell-format codec

`numbers_table_cell_number_format_codec` drives the neutral strict Number
format seam for one borrowed `FormatStructArchive` payload. Successful cases
compare scalar and measured lazy reads, preserve the caller's source slice,
round-trip fixed and automatic decimal places, and replay prepared
source-preserving rewrites with exact finite limits. Canonical appends are
checked independently, while unknown scalar/group spans are required to stay
byte-for-byte intact. Fixed recipes keep missing, duplicate, wrong-wire,
incompatible, invalid-domain, truncated, and unmatched-group errors hot.

The target accepts at most 64 KiB and uses 128 KiB output, 16,384 fields,
512 KiB of aggregate work, and nesting depth 64. Corpus entries are small
hand-authored `hex:` wire recipes; they are not copied native package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_cell_number_format_codec
```

Run a bounded smoke with mutable corpus, artifacts, and build output outside
the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-number-format-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_cell_number_format_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_cell_number_format_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers Percentage cell-format codec

`numbers_table_cell_percentage_format_codec` exercises the strict Percentage
variant of the same borrowed `FormatStructArchive` seam. It requires native
format type 258, so a valid Number (type 256) payload is rejected by this
wrapper even though both formats share fields for decimal places, negative
style, and thousands separators. Successful cases compare scalar and measured
reads, preserve unknown scalar/group spans, replay source-preserving rewrites
under exact finite limits, and round-trip canonical automatic and fixed
precision writes. Fixed recipes keep missing, duplicate, wrong-wire,
incompatible, noncanonical, invalid-domain, cross-family, and malformed-group
errors hot.

The target accepts raw inputs up to 64 KiB and uses 128 KiB output, 16,384
fields, 512 KiB of aggregate work, and nesting depth 64. Corpus entries are
small hand-authored `hex:` wire recipes; they are not copied native package
bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_cell_percentage_format_codec
```

Run a bounded smoke with mutable corpus, artifacts, and build output outside
the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-percentage-format-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_cell_percentage_format_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_cell_percentage_format_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers Currency cell-format codec

`numbers_table_cell_currency_format_codec` exercises the strict Currency
variant of the borrowed `FormatStructArchive` seam. It covers native format
type 257 with standard and accounting styles, automatic and fixed decimal
places, validated three-letter currency codes, source-preserving rewrites,
canonical appends, and exact scalar/report resource parity. Fixed recipes
keep invalid code, style, decimal, and Boolean domains hot, together with
missing-each-selected-field, duplicate, non-canonical, wrong-wire, truncated,
malformed-group, and Number/Percentage cross-family refusals.

Unknown scalar, length-delimited, and balanced-group spans are interleaved
with selected fields, repeated, and checked by parsed wire order and exact
multiplicity after every rewrite. The target accepts raw inputs up to 64 KiB
and uses 128 KiB output, 16,384 fields, 512 KiB of aggregate work, and nesting
depth 64. Corpus entries are small hand-authored `hex:` wire recipes; they
are not copied native package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_cell_currency_format_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-currency-format-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_cell_currency_format_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_cell_currency_format_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers Scientific cell-format codec

`numbers_table_cell_scientific_format_codec` exercises the strict Scientific
variant of the borrowed `FormatStructArchive` seam. It requires native format
type 259, every fixed precision from 0 through 30, the native minus-sign
negative style, and a hidden thousands separator. Scientific is intentionally
not interchangeable with Number (256), Currency (257), or Percentage (258),
and automatic decimal places are rejected.

Successful cases compare scalar and measured reads, retain a borrowed source,
replay exact source-preserving rewrites, and keep unknown scalar, fixed-width,
length-delimited, and balanced nested-group spans byte-for-byte in their wire
order and multiplicity. Fixed recipes keep every selected field's missing,
duplicate, wrong-wire, noncanonical, and invalid-domain path hot, together
with wrong-family and malformed-group refusals. Canonical writes independently
round-trip all valid precision values and every rewrite is retried against
one-below exact output, field, work, depth, allocation, and retained-byte
limits.

The target accepts raw inputs up to 64 KiB and uses 128 KiB output, 16,384
fields, 512 KiB of aggregate work, and nesting depth 64. Corpus entries are
small hand-authored `hex:` wire recipes; they are not copied native Numbers
package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_cell_scientific_format_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-scientific-format-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_cell_scientific_format_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_cell_scientific_format_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers Fraction cell-format codec

`numbers_table_cell_fraction_format_codec` exercises the strict native
Fraction variant of the borrowed `FormatStructArchive` seam. It covers native
format type 262 and all nine denominator strategies: up to one, two, or three
digits, halves, quarters, eighths, sixteenths, tenths, and hundredths. The
legacy `requires_fraction_replacement` field 20 is accepted only as a
canonical false marker; its exact source bytes are preserved by rewrites and
the canonical writer never synthesizes it. A true value, duplicate,
wrong-wire, or noncanonical field-20 value is rejected.

Successful cases compare scalar and measured reads, retain a borrowed source,
replay exact source-preserving rewrites, and keep unknown scalar, fixed-width,
length-delimited, and balanced nested-group spans byte-for-byte in their wire
order and multiplicity. Fixed recipes keep every selected field's missing,
duplicate, wrong-wire, noncanonical, and invalid-domain path hot, together
with Number/Percentage/Currency wrong-family and malformed-group refusals.
Canonical writes and rewrites are retried against one-below exact output,
field, work, depth, allocation, and retained-byte limits.

The target accepts raw inputs up to 64 KiB and uses 128 KiB output, 16,384
fields, 512 KiB of aggregate work, and nesting depth 64. Corpus entries are
small hand-authored `hex:` wire recipes; they are not copied native Numbers
package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_cell_fraction_format_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-fraction-format-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_cell_fraction_format_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_cell_fraction_format_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers formula-archive codec

`numbers_formula_archive` sends one bounded, caller-owned
`TSCE.FormulaArchive` source through strict inspection and the visitor
decoder, and through the legacy-compatible render-event adapter with a
bounded, source-borrowing event sink. Successful inputs must agree on report
bounds, node and precedent counts, semantic node values, source-order
callbacks, and generated-Prost structural counts. Render strings are checked
to point into the caller's unchanged source bytes; malformed and
max-minus-one render probes must not publish partial events. The target also
constructs independent duplicate, missing-required, wrong-wire, non-canonical,
invalid-UTF-8, aggregate-budget, and deep-recursion cases so those contracts
do not depend on libFuzzer discovering a valid formula first.

The target accepts raw inputs up to 64 KiB and uses finite limits of 8,192
fields, 256 KiB of work, 2,048 nodes, 64 KiB of text, wire recursion depth 32,
and independent render recursion depth 64. Render options explicitly permit
canonical opaque unknown fields and unknown function identifiers. The
checked-in `corpus/numbers_formula_archive/empty_ast_array.hex` is a minimal
empty AST envelope; additional hand-authored recipes cover scalar, string/date/
duration, array/list, unknown-function, local/cross-table/coordinate,
category/range, thunk, opaque unknown scalar, unknown-group handling, and
malformed-nested render paths.
Generated cases cover the malformed and resource boundaries above. Corpus
recipes use the `hex:` form and are hand-authored, not copied from a native
Numbers package.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_formula_archive
```

Run a bounded smoke with mutable corpus, artifacts, and build output outside
the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-formula-archive-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_formula_archive/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_formula_archive "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=1 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Pages section codec

`pages_section_codec` compares strict pagination and section-settings
projections for one unchanged source. Successful settings snapshots must agree
between scalar and report paths, borrow non-empty section names from the source,
and match generated `TP.SectionArchive` values when the complete message is
decodable. Pagination intentionally keeps unrelated fields opaque, so malformed
oracle-only fields are observed without weakening the selected-field checks.
The target also probes finite byte, field, work, name, and recursion limits.

The target accepts raw inputs up to 64 KiB and uses 8,192 fields, 256 KiB of
work, 64 KiB of section-name text, and recursion depth 64. The checked-in
recipes under `corpus/pages_section_codec/` cover canonical scalar settings,
pagination scalars, optional and empty envelopes, duplicate and missing
required fields, invalid UTF-8/NUL names, unknown scalar/group spans, wrong
wire types, truncation, and deep groups. They are hand-authored `hex:` recipes,
not copied from native Pages packages.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_section_codec
```

## Keynote chart-title codec

`keynote_chart_title` sends one bounded, caller-owned generated
`TSCH.ChartNonStyleArchive` extension through strict chart-title reads and
wire-local rewrites. A successful input covers borrowed field-21/23 reads,
title-only and visible-title helpers, every proto2 presence/value combination,
exact no-op/set/clear behavior, unknown wire-span preservation, and a semantic
inverse that restores the selected values and unknown spans when the requested
pair is representable. Exact source bytes are required when selected wire spans
remain positionally recoverable; removing selected spans can lose their original
interleaving, so those inverse cases are checked semantically while retaining
every unknown span. A clear request (`title_visible = Some(false)` with no title)
against a hidden or absent title is an intentional exact no-op; if an earlier
rewrite removed a hidden field-21 span while no title text was present, its
inverse therefore preserves the candidate rather than recreating that span.
Failed title- and
output-capped rewrites must not publish a partial candidate or modify their
source. Malformed mutations are required to remain rejected without modifying
their source. Error formatting is checked against a private sentinel, so
malformed content cannot be reflected in `Display` or `Debug`.

The target accepts raw inputs up to 64 KiB and uses finite limits of 8,192
fields, 256 KiB of aggregate work, 128 KiB of rewrite output, 64 KiB of title
text, and nesting depth 64. The checked-in recipes under
`corpus/keynote_chart_title/` cover empty and missing-visible states,
visible-empty, hidden-present and hidden-empty title states, unknown-only
payloads, Unicode, unknown scalar/group spans, duplicate selected fields,
non-canonical wire, invalid UTF-8, and truncation. They are hand-authored
`hex:` protobuf wire recipes rather than copied native package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check keynote_chart_title
```

Run a bounded AddressSanitizer/libFuzzer smoke with all mutable corpus,
artifact, and build locations outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-keynote-chart-title-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/keynote_chart_title/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  keynote_chart_title "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. Corpus additions,
artifacts, and build output stay in the temporary root; set
`KEEP_FUZZ_CORPUS=1` to retain it for review.

## Keynote chart value-axis settings codec

`keynote_chart_axis_value_settings_codec` drives the strict, source-preserving
projection for the selected value-axis scalar settings in
`TSCH.Generated.ChartAxisNonStyleArchive`. It covers manual or automatic
bounds (including partial endpoints), major/minor step counts, and linear,
logarithmic, or forward-compatible scale values. The harness compares the
scalar and reported decode paths, keeps the caller-owned source immutable,
replays every prepared rewrite with its exact resource requirements, and
requires the one-shot helper to produce the same candidate and report.

Successful rewrites are decoded again and checked for semantic readback,
exact no-op bytes, and preservation of every unrelated root and nested span,
including unknown fields, balanced groups, field order, and signed-zero
payloads. Malformed recipes exercise duplicate or wrong-wire selected fields,
non-canonical and truncated varints/lengths, missing nested bound values,
non-finite or inverted bounds, negative and zero step counts, and malformed
signed-int32 encodings. Every candidate is attempted only with finite input,
field, work, nesting, allocation, retained-source, scratch, and output
ceilings; max-minus-one replays must fail atomically without changing the
source.

The target accepts at most 64 KiB and uses 8,192 wire fields, 512 KiB of
aggregate work, 128 KiB of output, and nesting depth 64. Its checked-in
`corpus/keynote_chart_axis_value_settings/` recipes are small hand-authored
`hex:` protobuf payloads covering automatic/full/partial settings, unknown
interleaving and groups, nested unknown spans, and malformed boundaries.
They are not copied from native Keynote packages.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check keynote_chart_axis_value_settings_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-keynote-chart-axis-value-settings-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/keynote_chart_axis_value_settings/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  keynote_chart_axis_value_settings_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Keynote chart-axis-title codec

`keynote_chart_axis_title_codec` sends one bounded generated
`TSCH.Generated.ChartAxisNonStyleArchive` extension through the strict
category/value title projection. Successful cases cover category fields 13/15
and value fields 14/16 across absent, visible-empty, hidden-stale, set, clear,
and no-op states. Rewrites are checked for exact source atomicity, semantic
inverse/readback, and byte-for-byte preservation of unknown, opposite-axis,
and secondary marker spans. Fixed mutations cover wrong wire types, duplicate
singular fields, non-canonical lengths, invalid UTF-8, truncation,
unterminated groups, redacted malformed errors, and max-minus-one field,
work, output, title, input, and nesting limits.

The target accepts raw inputs up to 64 KiB and uses 8,192 fields, 256 KiB of
aggregate work, 128 KiB of rewrite output, 64 KiB of title text, and nesting
depth 64. Prepared rewrites also enforce four logical allocations, 256 KiB of
retained source/candidate bytes, and 128 KiB of candidate scratch. The harness
replays the exact prepared requirements and each max-minus-one execution
ceiling before publishing a candidate, so failed output is never observable.
The checked-in `corpus/keynote_chart_axis_title/` recipes are tiny
hand-authored `hex:` wire payloads covering category/value presence, visible
empty and hidden-stale values, Unicode, unknown interleaving, opposite and
secondary fields, fixed-width unknown values, balanced groups, and malformed
boundaries. They contain no native package bytes or crash artifacts.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check keynote_chart_axis_title_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-keynote-chart-axis-title-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/keynote_chart_axis_title/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  keynote_chart_axis_title_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

The `.hex` recipes are decoded by the harness into bounded caller-owned
inputs; they are not copied into a mutable in-checkout libFuzzer corpus.

## Keynote movie geometry and transform codec

`keynote_movie_geometry_codec` drives the strict, source-preserving geometry
projection for the drawable envelope inside `TSD.MovieArchive`, together with
the optional native flags/angle transform projection. Successful inputs
exercise scalar decode accounting, prepared rewrite/execute replay, optional
transform preserve/set/clear rewrites, candidate readback, and one-shot
equivalence without exposing generated protobuf values. Transform writes
toggle only the semantic reflection bit in generated commands while the
source-preserving codec retains all other native flag bits and unknown spans.
The target also probes malformed and truncated envelopes, duplicate and
non-finite transform fields, unknown balanced groups, unknown overlong
scalars, unterminated groups, and a deep group chain that must stop at the
typed nesting ceiling.

The target accepts at most 64 KiB and uses 8,192 fields, 512 KiB of work, 128
KiB of output, and recursion depth 64. Prepared execution is replayed against
exact and max-minus-one output, field, work, depth, allocation, retained-byte,
and scratch ceilings; every refusal is checked after preserving the caller's
source bytes. Recipes under `corpus/keynote_movie_geometry_codec/` are
hand-authored `hex:` wire inputs, not copied native packages.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check keynote_movie_geometry_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-keynote-movie-geometry-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/keynote_movie_geometry_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  keynote_movie_geometry_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers persisted table-sort-order codec

`numbers_table_sort_order_codec` fuzzes the strict field-44 projection of a
complete `TST.TableModelArchive`. Successful payloads are compared through
scalar and reported decode paths, then passed through exact prepared
set/clear/no-op rewrites and candidate readback. Preparation is followed by
the exact `RewriteExecutionRequirements` replay and independent max-minus-one
output, field, work, depth, rule, allocation, retained, and scratch probes.
The fixed recipes under `corpus/numbers_table_sort_order_codec/` cover absent
and canonical orders, selected-row rules, duplicate and wrong-wire known
fields, non-canonical varints, duplicate columns, truncation, unknown
overlong scalars, balanced unknown groups, and unterminated groups. They are
hand-authored `hex:` complete-model payloads, not native package members.

The target accepts raw inputs up to 64 KiB and uses finite ceilings of 128 KiB
output, 16,384 fields, 512 KiB work, 1,024 rules, 4,096 columns, and nesting
depth 64. Type-check it with:

```sh
cargo +nightly fuzz check numbers_table_sort_order_codec
```

Run a bounded sanitizer smoke with mutable state outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-table-sort-codec-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/numbers_table_sort_order_codec/*.hex "$fuzz_corpus/"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_sort_order_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers table-header settings codec

`numbers_table_header_settings_codec` drives the strict generated-free
projection for the required table dimensions and seven optional header,
footer, freeze, and print-repetition fields. Successful payloads are decoded,
prepared, executed under the exact `RewriteExecutionRequirements`, and read
back through the strict decoder. The one-shot rewrite must match prepared
execution byte-for-byte, and the caller-owned source must remain unchanged.

Every prepared candidate is replayed with one less input-byte, output, field,
work, nesting, allocation, retained-byte, and scratch ceiling; each
max-minus-one operation must fail before publishing a candidate. Fixed `hex:`
recipes under `corpus/numbers_table_header_settings_codec/` cover all optional
fields, duplicate and missing required fields, wrong wire types, non-canonical known
keys/values, truncation, unknown overlong scalars, balanced unknown groups,
unterminated groups, and deep nesting. These are hand-authored protobuf
payloads, not native Numbers package members.

The target accepts raw inputs up to 64 KiB and uses finite ceilings of 128 KiB
output, 8,192 fields, 512 KiB work, nesting depth 64, 16 allocations, 128 KiB
retained bytes, and 512 KiB scratch. Type-check it with:

```sh
cargo +nightly fuzz check numbers_table_header_settings_codec
```

Run a bounded sanitizer smoke with mutable state outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-table-header-codec-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/numbers_table_header_settings_codec/*.hex "$fuzz_corpus/"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_header_settings_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Format-neutral table-dimension codec

`table_dimension` drives the hidden `table_dimension_codec` facade over one
`TST.HeaderStorageBucket` payload. Successful inputs compare scalar and
reported reads, check source-ordered header callbacks borrow from the original
source, and exercise exact no-op, size-set, and size-remove rewrites. The
preflight/execute path is also exercised: the candidate must satisfy its exact
output requirement, while an output ceiling one byte below that requirement
must fail without changing the source. Malformed required fields, duplicate or
wrong-wire selected fields, non-canonical varints, truncation, unknown scalar
values, and balanced/unbalanced unknown groups are all kept in the fixed cases.

The target accepts raw inputs up to 64 KiB and uses finite limits of 8,192
fields, 256 KiB of work, 1,024 references, 64 KiB of text, and recursion depth
64. The recipes under `corpus/table_dimension/` are hand-authored `hex:` wire
encodings for empty, canonical, unknown-preserving, duplicate, wrong-wire,
and truncated buckets; they are not copied from a native Numbers package.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check table_dimension
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-table-dimension-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/table_dimension/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  table_dimension "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. Corpus additions,
artifacts, and build output stay in the temporary root; set
`KEEP_FUZZ_CORPUS=1` if the temporary campaign should be retained.

## Format-neutral table-model discovery codec

`table_model_discovery_codec` fuzzes the borrowed, strict discovery projection
used to admit a canonical `TST.TableModelArchive`. It returns only the
caller-owned table identifier, display name, and dimensions; it does not
construct a package, ZIP member, generated model, or rewrite candidate.
Reported scalar and facts entry points are replayed against the same bytes and
must agree exactly, including source borrowing and traversal accounting.

Fixed protobuf recipes and generated cases cover all required fields,
canonical unknown scalars/fixed values/length-delimited values, balanced and
deep unknown groups, malformed and mismatched groups, truncated framing,
duplicate and wrong-wire known fields, non-canonical keys/values, and invalid
wire types. Successful inputs are retried with one less input-byte, field,
work, and nesting ceiling; every refusal must be typed and leave the source
unchanged. Its codec-owned report categories for allocations, retained bytes,
and scratch bytes are zero because the snapshot retains only source borrows;
these values are not allocator telemetry, and the projection exposes no
mutable resource knobs for those axes.

The target accepts raw inputs up to 64 KiB and uses finite limits of 8,192
fields, 256 KiB of aggregate strict-plus-parity work, 64 KiB of borrowed text,
and nesting depth 64.
Recipes under `corpus/table_model_discovery_codec/` are hand-authored `hex:`
protobuf payloads, not native Keynote packages.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check table_model_discovery_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-table-model-discovery-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/table_model_discovery_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  table_model_discovery_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. Corpus additions,
artifacts, and build output stay in the temporary root.

## Keynote movie-playback codec

`movie_playback_codec` drives the hidden `movie_playback_codec` projection for
one bounded `TSD.MovieArchive` playback payload. Valid sources are decoded by
the scalar and reported paths, then passed through a prepared rewrite with
exact output/field/work/depth/allocation/retained/scratch limits and strict
candidate readback. The harness keeps the source caller-owned and exercises
unknown overlong scalars, balanced unknown groups, duplicate/wrong-wire
known fields, truncation, and max-minus-one execution ceilings. The recipes
under `corpus/movie_playback_codec/` are hand-authored protobuf wire cases,
not native Keynote package members.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check movie_playback_codec
```

## Format-neutral TableInfo ownership and lock codec

`table_info_codec` fuzzes the strict generated-free `TST.TableInfoArchive`
projection used by the Pages and Keynote table owners. It replays the complete
lock-aware snapshot and the model-reference convenience decoder against the
same caller-owned bytes, covering absent, explicit `locked=false`, and
`locked=true` envelopes. Required `super`, table-model, and nested identifier
routes are checked alongside zero identifiers, duplicate singular fields,
wrong-wire fields, non-canonical framing, truncated payloads, and malformed or
mismatched unknown groups. A successful decode must leave the source bytes
unchanged and both public decoder entry points must agree.

For every valid source, the target also exercises the prepared
`prepare_table_info_lock_rewrite` contract for absent, explicit false, and
explicit true lock representations. It checks prepare-report and execution-
requirement parity, exact execution, strict candidate reread, no-op byte
identity, inverse restoration, one-shot parity, and stale-fingerprint
rejection. Each prepared input and execution ceiling is probed at
max-minus-one for input bytes, output bytes, fields, work bytes, nesting,
logical allocations, retained bytes, and scratch bytes; rejected probes must
leave the caller-owned source unchanged. The rewrite path is raw-preserving:
unknown canonical scalars, fixed-width values, length-delimited values, and
balanced groups remain in the source framing while only the selected lock
field may change.

The target accepts raw inputs up to 64 KiB and uses finite limits of 8,192
fields, 256 KiB of strict-plus-projection work, and recursion depth 64. The
checked-in recipes under `corpus/table_info_codec/` are hand-authored `hex:`
protobuf payloads, not native package archives. The prepared rewrite keeps all
candidate bytes in the codec-owned output and applies its exact requirements
before execution; this target reports codec-owned logical limits rather than
allocator/RSS telemetry.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check table_info_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-table-info-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/table_info_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  table_info_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Format-neutral table-appearance codec

`table_appearance` drives the generated-free `table_appearance_codec` facade
over bounded TableModel, TableStyle, TableStylePreset, TableStyleNetwork, and
Stylesheet payloads. Successful model and stylesheet rewrites execute with
exact prepared requirements and decode again, exercising source preservation,
candidate verification, unknown-field retention, and typed
output/field/work/allocation ceilings. The preset and network discovery
projections are replayed through both plain and reported decoders and probe
their input/field/work/nesting ceilings one below the reported requirement.
Malformed, duplicate, wrong-wire, non-canonical, truncated, and unbalanced-
group inputs are observed without mutating their caller-owned source. The
recipes under `corpus/table_appearance/` include canonical known properties,
unknown overlong scalars/groups, preset/network reference failures,
style-edge rewrites, and strict failure shapes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check table_appearance
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-table-appearance-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/table_appearance/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  table_appearance "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. Corpus additions,
artifacts, and build output stay in the temporary root; set
`KEEP_FUZZ_CORPUS=1` if the temporary campaign should be retained.

## Numbers Pop-Up Menu codec

`numbers_table_cell_pop_up_menu_codec` drives the hidden
`numbers_table_cell_pop_up_menu_codec` projection for one bounded
`TST.PopUpMenuModel` payload. Successful inputs are decoded through the scalar
and reported paths, then passed through the prepared rewrite/execute route
with exact requirements and candidate readback. The harness checks that the
caller-owned source is unchanged, that prepared output/field/work/nesting and
allocation ceilings are honored, and that an output or work limit one below
the exact requirement fails before a candidate can be published.

The fixed recipes under
`corpus/numbers_table_cell_pop_up_menu_codec/` cover the required NIL sentinel,
string choices, deprecated `item`, missing/duplicate sentinel, wrong value
type, typed string options, duplicate known fields, unknown overlong scalars,
and balanced unknown groups. They are hand-authored protobuf wire recipes,
not copied from a native Numbers package. Unknown bytes are retained only
where the strict codec policy permits them; malformed known fields remain
fail-closed.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_cell_pop_up_menu_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-numbers-popup-codec-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_cell_pop_up_menu_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_cell_pop_up_menu_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. Corpus additions,
artifacts, and build output stay in the temporary root; set
`KEEP_FUZZ_CORPUS=1` if the temporary campaign should be retained.

## Numbers unified cell-control codec

`numbers_table_cell_control_codec` drives the neutral hidden
`numbers_table_cell_control_codec` facade for interactive `CellSpecArchive`
and `FormatStructArchive` payloads. It covers native interaction types 4--8,
including strict fixed64 slider/stepper/star ranges and canonical checkbox
formats, alongside popup references and the rejection path for unsupported
or deprecated fields. Scalar and report decoders must agree and leave the
caller-owned bytes unchanged. Prepared CellSpec and format writes are
executed with exact requirements, decoded again, and replayed with each
output/field/work/depth/reference/allocation/retained/scratch ceiling one
below the requirement to prove fail-closed execution.

The target also feeds duplicate/wrong-wire/non-canonical known fields,
non-finite or reversed ranges, truncated references, unknown overlong
scalars, balanced unknown groups, and deeply nested groups. The recipes under
`corpus/numbers_table_cell_control_codec/` are hand-authored protobuf wire
inputs and never native package copies.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_cell_control_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-numbers-control-codec-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_cell_control_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_cell_control_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. Corpus additions,
artifacts, and build output stay in the temporary root.

`numbers_tile_storage` sends one bounded, caller-owned byte source through both
tile entry points:

* `decode_tile_with_report` (scalar/report path); and
* `decode_tile_with_visitor` (streaming row path).

For a successful decode the target checks scalar equality, exact
`DecodeReport` equality, and source-order/count summaries from the streaming
callbacks. Every callback checks that each non-empty row payload points into
the unchanged source. The visitor retains only compact summaries for the
first `MAX_FIELDS / 32` rows (256 with the fixed limits below); pointer checks
continue for all later callbacks without retaining their payloads. When
decoding fails, it consumes no partial result and only checks that both strict
paths fail; a streaming visitor may have seen rows before the later error, as
specified by the visitor API. After strict scalar success, the target decodes
the same source with the generated Prost `tst::Tile` oracle and compares every
tile scalar plus every streamed row scalar, optional-field presence, payload,
and source-order ordinal. A Prost rejection after strict acceptance is a
fuzz failure.

The target accepts raw inputs up to 64 KiB. Checked-in corpus entries are
human-readable `hex:` recipes; the harness decodes those recipes before the
two calls. This keeps the corpus reviewable while ensuring the codec sees
the exact same binary source in each path. Invalid or oversized recipes are
skipped. The finite per-source policy is 8,192 fields, 256 KiB of work, 1,024
references, 64 KiB of text, and recursion depth 64.

## Corpus provenance

The seeds are hand-authored protobuf wire encodings from
`src/buffa-projections/TSTTableCellStorageArchive.proto` and the matching
`src/protos/TSTArchives.proto`; they are not copied from a private Numbers
document. `modern.hex` includes current/BNC row buffers and all optional tile
scalars. `pre_bnc.hex` contains only the required pre-BNC row buffers.
`empty.hex` is the required-field, zero-value tile. `wide.hex` has sixteen
wide-offset rows and 8,191-by-8,191 tile bounds. `unknown_fields.hex` appends
unknown scalar and length-delimited fields. The three `malformed_*.hex`
recipes are near-valid duplicate-required, overlong-varint, and truncated-row
inputs. The recipes are deterministic regression inputs; they do not claim to
represent a complete native Numbers corpus.

From this directory, list and type-check the target:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_tile_storage
```

Run a bounded nightly sanitizer smoke (the target itself is intentionally
short; increase `-runs` only for a longer local campaign). The seed corpus is
copied to a temporary directory so libFuzzer's corpus additions, artifacts,
and build output never land in this repository. Set `KEEP_FUZZ_CORPUS=1` to
retain that temporary directory for review; otherwise the exact temporary
directory is removed on exit.

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-tile-storage-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/numbers_tile_storage/*.hex "$fuzz_corpus/"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_tile_storage "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=1 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. The corpus and target
do not write generated artifacts into the repository. If a run is interrupted,
the temporary directory remains recoverable under the system temporary
directory; set `KEEP_FUZZ_CORPUS=1` before the command to print and retain its
path for review, then inspect or move it before removing it.

## Direct TableDataList and segment codec

`numbers_table_data_list` attempts the same unchanged source independently as
both a `TableDataList` root and a `TableDataListSegment`. For each attempted
shape it compares the strict scalar/report path with the strict visitor path.
After strict success only, it decodes the source with generated Prost and
compares scalar values, repeated-field source order, optional-field presence,
and nested message values. Every borrowed string, opaque message payload,
segment reference envelope, and segment range payload is checked to point into
the unchanged source. A bounded prefix of callback summaries is retained;
pointer and ordinal checks still run for every callback. The root and segment
attempts have independent success/failure outcomes, and malformed input for
one shape must not prevent the other attempt.

The target accepts at most 64 KiB and configures finite limits of 8,192 fields,
256 KiB of work, 1,024 references, 64 KiB of UTF-8 text, and recursion depth
64. The named recipes in `corpus/numbers_table_data_list/` cover:

* `minimal_root.hex`, `minimal_segment.hex`, and `segment_entries.hex`;
* `every_entry_field.hex`, including all optional entry fields;
* `segments_references.hex`, with repeated entries and references;
* `utf8_unknown_groups.hex`, with non-ASCII text and an unknown group;
* `duplicate_fields.hex`; and
* `malformed_varint.hex`, `malformed_truncated.hex`,
  `malformed_range.hex`, and `malformed_reference.hex`.

The additional root-focused recipes are
`root_duplicate_next_list_id.hex`, `root_duplicate_optional_bool.hex`,
`root_entry_missing_refcount.hex`, `root_invalid_optional_bool.hex`,
`root_missing_next_list_id.hex`, `root_overflow_next_list_id.hex`,
`root_truncated_segment_reference.hex`, and `root_unclosed_group.hex`.
The segment-focused recipes are `segment_duplicate_range.hex`,
`segment_duplicate_type.hex`, `segment_entry_missing_refcount.hex`,
`segment_missing_list_type.hex`, `segment_missing_range.hex`,
`segment_range_duplicate_location.hex`, `segment_range_overflow_location.hex`,
`segment_unclosed_group.hex`, `segment_wrong_list_type_wire.hex`, and
`segment_wrong_range_wire.hex`.

List and type-check this target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_data_list
```

Run a bounded AddressSanitizer/libFuzzer smoke with all mutable corpus,
artifact, and build locations outside the checkout. `cargo +nightly fuzz run`
uses the sanitizer fuzzing profile; increase `-runs` only for a longer local
campaign.

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-table-data-list-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/numbers_table_data_list/*.hex "$fuzz_corpus/"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_data_list "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=1 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

The checked-in files are `hex:` recipes rather than generated corpus output,
so they remain reviewable and the target feeds the exact decoded bytes to both
strict entry points. Invalid or oversized recipes are skipped.


## Pages body and section-boundary codec

`pages_body_footnote_codec` sends one bounded, caller-owned source through both
the document-body and section-boundary projections. Successful snapshots are
compared with generated Prost values, every borrowed reference is checked to
remain inside the unchanged source, and malformed or resource-limited inputs
are observed independently for both shapes. The target also checks source
atomicity across mutations and invalid limit configurations.

The target accepts at most 64 KiB and uses 8,192 fields, 256 KiB of aggregate
work, and recursion depth 64. Corpus entries under
`corpus/pages_body_footnote_codec/` are hand-authored `hex:` recipes covering
canonical document and boundary envelopes, optional fields, duplicate and
missing required fields, invalid references, non-canonical wire, truncation,
and unknown groups.

List and type-check this target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_body_footnote_codec
```

## Pages body-footnote graph codec

`pages_footnote_graph_codec` drives the hidden generated-free graph seam used
by the Pages body-footnote owner. It creates bounded canonical reference,
storage, marker, and body-entry payloads from each input, checks exact output
and resource reports, exercises one-under output/field/work budgets, then
decodes and rewrites the body table while retaining unknown scalar and group
spans. It also checks source immutability and candidate semantic readback.

The target accepts at most 64 KiB and uses 256 KiB output/work, 8,192 fields,
1,024 entries, and recursion depth 64. Corpus entries under
`corpus/pages_footnote_graph_codec/` are hand-authored `hex:` command seeds;
they are not native package bytes and are isolated from the existing body and
footnote codec corpora.

```sh
cargo +nightly fuzz check pages_footnote_graph_codec
```

## Pages footnote reference and marker codecs

`pages_footnote_codec` drives both strict, caller-owned Pages footnote
projections from one bounded source. Successful marker reads must retain the
exact raw source and successful reference reads must preserve the same
source-owned no-op candidate. Optional textual fields are checked for source
borrowing, and a generated Prost decode is used as a semantic oracle when it
accepts the complete wire payload. Every read and no-op candidate is checked
for source atomicity.

Arbitrary byte mutations are observed independently for the marker and
reference shapes. Fixed recipes cover duplicate singular fields, missing and
zero required identities, invalid UTF-8, wrong wire types, non-canonical
varints, truncation, unknown scalar/group spans, and Unicode text. The target
also probes finite input, output, field, work, and nesting ceilings. The
codec has no production encoding path, so output is modeled as the exact raw
no-op candidate and is capped explicitly in the harness.

The target accepts at most 64 KiB and uses 8,192 fields, 256 KiB of aggregate
work, 128 KiB of no-op output, and recursion depth 64. Corpus entries under
`corpus/pages_footnote_codec/` are hand-authored `hex:` recipes; they are not
copied from native Pages packages.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_footnote_codec
```

## Pages movie-caption inheritance and rewrite codec

`pages_movie_caption_codec` drives the selected `TSA.CaptionInfoArchive`
inheritance chain through strict borrowed reads and raw-preserving reference
remaps. Successful inputs cover the required drawable/shape/shape-info chain,
optional placement and storage references, unknown fields, proto2 boolean and
enum presence, exact no-op writes, wide identifier rewrites, source
atomicity, and candidate readback. Malformed and finite-limit recipes are
checked independently before any candidate is published.

The target accepts at most 64 KiB and uses 8,192 fields, 256 KiB of aggregate
work, 128 KiB of candidate output, and recursion depth 64. Corpus entries under
`corpus/pages_movie_caption_codec/` are hand-authored `hex:` recipes; they are
not copied from native Pages packages.

List and type-check this target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_movie_caption_codec
```

Run a bounded AddressSanitizer/libFuzzer smoke with all mutable corpus,
artifact, and build locations outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-pages-movie-caption-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/pages_movie_caption_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  pages_movie_caption_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Pages header/footer section-template codec

`pages_header_footer_codec` exercises the strict source-preserving
`TP.SectionTemplateArchive` projection used by the Pages header/footer owner.
It decodes repeated header and footer `TSP.Reference` records without owning
native identifiers and rewrites a complete set of
references while preserving unknown fields, balanced groups, source order,
and unknown overlong scalar values. The replacement path verifies output
counts, exact report accounting, and caller-source immutability.

The target accepts raw inputs up to 64 KiB and uses finite ceilings of 8,192
fields, 256 KiB of work, 128 KiB of output, nesting depth 64, and 8,192
references. The checked-in recipes under `corpus/pages_header_footer_codec/`
cover canonical and optional reference fields, unknown groups and overlong
unknown scalars, missing/duplicate/non-canonical identifiers, wrong wire
types, truncation, and unterminated groups. Recipes are hand-authored
`hex:` protobuf payloads, not copied native package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_header_footer_codec
```

## Numbers comment-storage codec

`comment_storage_codec` sends one bounded, caller-owned
`TSD.CommentStorageArchive` source through both strict entry points:

* `decode_comment_storage_archive_with_report` (scalar/report path); and
* `decode_comment_storage_archive_with_visitor` (source-ordered reply path).

For strict successes the target requires equal snapshots and exact
`DecodeReport` values from both paths. It then uses generated Prost
`tsd::CommentStorageArchive` only as an oracle for that strict success, and
compares text presence/content, IEEE-754 creation-date bits, author and UUID
presence/values, deprecated reference fields, and every reply in source order.
Every reply payload and the borrowed text are checked against the unchanged
source pointer range. Callback state keeps only a bounded prefix of compact
reply summaries; callback counts, order, and borrow checks still cover every
reply. Rejections (including malformed wire, unknown groups, and finite-limit
failures) must remain rejections on both strict paths; a visitor may have
observed a valid prefix before a later error.

The target accepts raw inputs up to 64 KiB and uses finite limits of 8,192
fields, 256 KiB of work, 1,024 references, 64 KiB of UTF-8 text, and recursion
depth 64. The 20 checked-in recipes under
`corpus/comment_storage_codec/` cover empty and text-only roots, Unicode,
negative-zero and NaN date bits, deprecated/default reference presence,
ordered replies, zero/wide UUIDs, mixed field order, unknown scalars and a
well-formed group, plus duplicate, missing, invalid-UTF-8, truncated,
malformed-group, wrong-wire, and noncanonical-varint inputs. They are
hand-authored protobuf wire encodings, not copied from a private Numbers
document.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check comment_storage_codec
```

Run a bounded AddressSanitizer/libFuzzer smoke with all mutable corpus,
artifact, and build locations outside the checkout. Increase `-runs` only for
a longer local campaign.

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-comment-storage-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/comment_storage_codec/*.hex "$fuzz_corpus/"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  comment_storage_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation. Corpus additions,
artifacts, and build output stay in the temporary root; set
`KEEP_FUZZ_CORPUS=1` to retain it for review.

## Numbers comment-storage direct-reply rewrite codec

`comment_storage_reply_codec` drives the prepared, source-preserving
`TSD.CommentStorageArchive` direct-reply rewrite seam. Each bounded source is
decoded, then exercised through append, replace, and remove preparation and
execution. Successful candidates are read back through the strict decoder;
the target also restores the original reply sequence through the inverse-like
operation and checks that the caller-owned source remains unchanged. A wrong
expected reply identifier must fail before candidate publication.

The harness checks `CommentStorageReplyRewrite` preparation reports and exact
`RewriteExecutionRequirements` replay, including output bytes, fields, work,
nesting, references, scratch, retained bytes, and allocation ceilings. For
each successful preparation it retries every finite axis at its exact bound
and at one below that bound, requiring the latter to fail without an output
allocation. Malformed framing, wrong wire types, duplicate or missing reply
references, unterminated and deeply nested groups, and unknown balanced-group
bytes are exercised without allowing a panic or unbounded recursion.

The target accepts raw inputs up to 64 KiB and keeps all decode/rewrite limits
finite (64-level nesting and bounded fields, work, references, text, scratch,
retained bytes, and allocations). Recipes under
`corpus/comment_storage_reply_codec/` are hand-authored `hex:` payloads; they
are not copied from native Numbers packages.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check comment_storage_reply_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build
output outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-comment-storage-reply-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cp corpus/comment_storage_reply_codec/*.hex "$fuzz_corpus/"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  comment_storage_reply_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## Numbers table-title codec

`numbers_table_title_codec` drives the strict generated-free projection for
the optional `TST.TableModelArchive` table-title settings.  It compares the
scalar and reported decode paths over caller-owned bytes, checks every
proto2 presence state for the visibility/border flags and every IEEE-754
fixed64 height bit pattern (including negative zero, NaN, and infinities),
and forces both style references through the private Buffa lazy view.  A
successful snapshot must retain non-zero reference identifiers and valid
legacy reference scalar presence without copying or changing the source.

Canonical unknown varint, fixed64, bytes, fixed32, and balanced nested-group
spans are exercised alongside duplicate selected fields, wrong wire types,
non-canonical keys/values, missing or zero reference identifiers, truncated
payloads, invalid legacy values, malformed groups, and invalid field tags.
The target replays one rich source at its exact bytes/fields/work/nesting/
reference report limits and requires each max-minus-one limit to fail with the
matching typed `DecodeLimit`.  It also probes a group chain beyond the finite
nesting ceiling.  Failed and successful paths assert that the caller-owned
source remains byte-for-byte unchanged and that scalar/report results agree.

Inputs are bounded at 64 KiB with 8,192 fields, 256 KiB of work, depth 64,
and two selected references.  The 25 checked-in recipes under
`corpus/numbers_table_title_codec/` are hand-authored `hex:` protobuf wire
payloads; they contain no native Numbers package bytes or crash artifacts.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check numbers_table_title_codec
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build
output outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-numbers-table-title-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/numbers_table_title_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  numbers_table_title_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

`cargo +nightly fuzz run` is the sanitizer invocation.  Corpus additions,
artifacts, and build output stay in the temporary root; set
`KEEP_FUZZ_CORPUS=1` to retain it for review.

## Keynote physical table-sort codec

`keynote_table_physical_sort_codec` fuzzes the archive-free wire seam used by
Keynote's selector-first physical table owner. It drives strict source-
borrowing plans for `Tile.row_infos` and sparse `HeaderStorageBucket.headers`,
then checks the prepared and one-shot rewrites against each other. The target
also validates the row/column UID-map projection and its checked source-indexed
permutation type.

Every successful rewrite is reparsed and executed as an exact no-op. The
caller-owned source is compared after every decode, preflight, and execution;
unknown scalar/length-delimited spans and opaque row/header payloads must
remain present after permutation. Fixed recipes keep missing required fields,
duplicate required fields and row indexes, wrong wire types, non-canonical
varints, truncated messages, malformed UID pairs, and dimension mismatches
reachable. Output, work, and input ceilings are each retried one byte below
the exact preflight requirement, so refusal happens before candidate
publication.

Inputs are capped at 64 KiB, output at 128 KiB, fields at 16,384, records at
4,096, elements at 16,384, work/scratch at 512 KiB, and nesting at 64. Corpus
entries are small hand-authored `hex:` wire recipes; they are not copied
native Keynote package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check keynote_table_physical_sort_codec
```

Run a bounded smoke with mutable corpus, artifacts, and build output outside
the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-keynote-physical-sort-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/keynote_table_physical_sort_codec/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  keynote_table_physical_sort_codec "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```
