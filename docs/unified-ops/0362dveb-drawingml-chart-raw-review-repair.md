# 0362dveb — DrawingML chart raw review return

- Date: 2026-10-24.
- Returned source: `b028116636` (`docs(drawingml): bind chart review receipt`).
- Repair source: `6637852e88` (`fix(drawingml): close chart raw review gaps`).
- Status: source repair complete; coordinator-capped Cargo validation remains
  pending until the slot is explicitly granted.

## Returned findings repaired

The classic chart writer no longer materializes escaped copies of chart
document fields before its bounded sink sees them. Pivot-source names and
header/footer text are admitted against `Limits::max_text_bytes` before the
writer starts, and relationship identifiers are admitted against the typed
`AttributeValueBytes` policy, including package-supplied overrides. XML
escaping for these fields, the chart language attribute, and the shared text
element helper is streamed directly to the sink, preserving the previous
wire spelling without an unbounded temporary.

The reader now preflights the decoded, normalized value of retained language,
picture-format, and number-format attributes before calling
`decoded_and_normalized_value(...).into_owned()`. Entity spellings therefore
use their decoded byte count for `ChartTextBytes` admission, including the
axis, data-label, point-label, and trendline number-format paths.

ChartEx attribute keys now retain namespace identity. Unqualified schema
attributes remain addressable by their local names; relationships are exposed
as `r:*` only when their namespace URI is the strict or transitional
relationships URI; bound foreign and undeclared-prefix attributes are kept as
opaque expanded keys and cannot override `id`, `val`, `idx`, or another
unqualified semantic field. Existing valid relationship and foreign-part
fidelity is unchanged.

ChartEx tree admission now requires exactly one complete root, rejects a
second top-level element, non-whitespace text/CDATA/entity references outside
that root, DTDs, and processing instructions, while accepting one XML
declaration in the XML prolog. This deliberately withdraws the earlier
unverified end-name/default claim; no quick-xml end-name behavior was changed.
Waterfall subtotal indexes are counted and admitted against `max_points` as
each value is pushed, before the vector can grow beyond policy.

## Regression coverage

Focused tests cover decoded escaped language and number-format values,
pre-sink writer refusal for pivot/header/footer/relationship fields,
namespace-qualified and undeclared-prefix ChartEx attributes, exact ChartEx
root/prolog grammar, and a one-series ChartEx part with more subtotal indexes
than the selected point limit. Existing round-trip and relationship tests
continue to exercise valid wire output.

Only `rustfmt --edition 2024`, `git diff --check`, and the repository's staged
file-size/binary guards were run for this source-only return. No Cargo command,
native timing, binary, or large evidence artifact is part of this record.
