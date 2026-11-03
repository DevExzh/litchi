# 0422 — W4-I: value-axis settings, chart image fills, captions and titles

- Status: slice 1 in progress; the wave continues (§4). Branch `uo2/iwa-w4i`
  (cut from trunk `feat/unified-ops-api` at `c599ef9cb8`).
- Date: 2026-11-02.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (§0 wave table, §3.2 API
  shape, §4 parity tiers, §5 performance); the predecessor's successor brief
  [0419](0419-iwa-w4f-chart-remainder-captions.md) §3; the chart patterns
  and gate list of [0414](0414-iwa-w4c-chart-style-families.md) (§8,
  §11–§14); owner decision D1 and the [macOS queue](0363-mac-queue.md).
- Scope, in order: value-axis bounds, steps and scale (`chart.legend_axis`,
  Pages 6, Numbers 6); chart image fills (background, series, symbol; 3 per
  format); the shared caption lift, then the Numbers shape, image and movie
  captions and titles (15) and the chart captions (Pages 3, Numbers 3).
  Chart add/duplicate/remove (`chart_lifecycle`) and `table_create` belong to
  W4-G; table sort and pivot to W4-H.
- `performance_claim`: every migrated route measured is cheaper than its
  host route under callgrind (per-slice tables below); none misses the 5%
  bar.

## 1. Slice 1: value-axis bounds, steps and scale

12 rows flip (Pages and Numbers: the three host setters and the three host
reads each).

| Crate | What |
|---|---|
| `litchi-iwa-style` | New `value_axis`: Keynote's rules, lifted. `validate` (finite endpoints, minimum ≤ maximum, no non-positive endpoint on a logarithmic axis; a known scale given by its native value is canonicalised), `equal_exact`, `from_snapshot` (a stored aggregate the native format could not have written is invalid), `codec_write` (presence-preserving: a component equal to the stored one keeps its bytes, an absent scale field stays absent while linear) over `litchi_iwa_protos::keynote_chart_axis_value_settings_codec`. Payload layer `settings`/`rewrite` on a `TSCH.ChartAxisNonStyleArchive` payload: the rewrite applies the components in the host's order (bounds, steps, scale; minimum before maximum, major before minor), so fields absent from the source are appended exactly where the host's `set_chart_value_axis_{bounds,steps,scale}` append them. Workspace layer `read`/`write`: the chart's primary value-axis non-style object, refused when another chart axis references it (the host's `AxisNonStyleSlot::ensure_exclusive`), rewritten in place under a residual `series::Limits` profile with the in-place `MutationScope` (extension field 10,000 of the axis object). New engine `Refusal::InvalidSettings`; `chart_style::{owners, style_object, in_place_scope, rewrite_in_place, read_payload}` are crate-visible for it |
| `litchi-keynote` | `slide_chart_value_axis` keeps its bounded transaction (budget, selection, locality) and now calls the shared rules; its private `common_settings`, `validate_settings`, `settings_equal_exact`, `codec_settings` and `codec_write` (124 lines) are deleted. Behaviour unchanged (`tests/slide_chart_value_axis.rs`, `tests/chart_axis_api.rs`) |
| `litchi-pages` | `chart::value_axis` (`Package::body_chart_value_axis_settings`, `edit_body_chart_value_axis_settings` → `Edit::{set, set_bounds, set_steps, set_scale}`, `apply_body_chart_value_axis_settings`) on the chart-property core, Keynote's aggregate shape. Chart-series `Refusal::InvalidSettings` (`RefusalKind::InvalidValue`) |
| `litchi-numbers` | `sheet::chart::value_axis`, the same family on the chart-series workspace path (`&mut` staging, as the other Numbers chart families). Chart-series `Refusal::InvalidSettings` |

Parity (`value_axis_cases.rs`): `{pages_body,numbers_sheet}_chart_value_axis_set_parity`
run 13 sequences per format through the facade with the inverse replayed:
bounds, steps and scale into an axis that does not store them, a single
maximum, in-place bounds and steps, a scale over stored bounds, back to
automatic bounds and steps, logarithmic back to linear (the stored scale
stays explicit), both step fields removed and appended again, and two
three-component edits (the host writes bounds, steps, scale in order). All
Tier A (byte-identical, whole ZIP). `…_read_parity` compares the host's three
reads with the aggregate read on every state (Tier A). Regressions:

- `numbers_sheet_chart_value_axis_shared_axis_refuses`: a second chart whose
  value axis references the first chart's non-style object; the host and the
  leaf refuse, the source is unchanged.
- `chart_value_axis_logarithmic_non_positive_bound_refuses` (stricter
  admission, §4.2 of 0363): the host stores a logarithmic scale over a
  negative minimum; the leaf refuses the staging (`Refusal::InvalidSettings`)
  as Keynote always has, and reads such a source as invalid. Listed in the
  macOS queue.

Eight macOS-queue rows added (the six setters and two reads).

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| pages_value_axis_set | 1,193,208,991 | 48,131,065 | 0.0403 |
| numbers_value_axis_set | 151,334,720 | 77,925,554 | 0.5149 |

Probe `chart_property_cost.rs` (`W4C_ROUTE`, release, CPU 3, ABBA, inclusive
Ir of `w4c_host_route`/`w4c_leaf_route`; the route sets fixed bounds 0–50).
Receipt `results/0422/callgrind-value-axis.json`; raw captures in
`.coord/iwa-w4i/cg1/`. Keynote has no host value-axis route to compare; its
transaction path is unchanged apart from the shared rule calls.
