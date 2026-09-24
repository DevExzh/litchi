# Condition-cache outside-shape review

Status: VP15-1 remains blocked on the reviewed value-VM snapshot. This is a
source-only review of condition-cache shape keys and the joint matrix-branch
shape planner. It excludes the concurrently changing reference-geometry
helper and the separate VP15-2 reservation work.

## Reviewed snapshot and reproduction

The frozen condition-cache snapshot reviewed for this finding is

```text
crates/litchi-ods/src/codec/formula/evaluation/value.rs
c76e3ddb2546b151a3189a578c5c08208e500c5264de4ba9fa92dff9ea0c464e
```

The independent reproduction is recorded in
`tests/ods_formula_array_reference_evaluation.rs` as
`finite_inner_condition_does_not_cache_out_of_shape_broadcast_positions`
(lines 708–750 in the reproduction checkout). The captured baseline log is retained as
`planner-outside-shape-baseline-test.log` in
[`planner-remaining-boundaries-baseline.tar.gz`](../performance/diagnostics/planner-remaining-boundaries-baseline.tar.gz).
It reports:

```text
ResourceLimit { resource: Objects, observed: 33, limit: 32,
                scope: "ods-formula-evaluation" }
```

The test uses a small expression and a selected widening sibling:

```text
=IF({TRUE();FALSE()};IF(IF({TRUE()|TRUE()};1;0);1;2);[.B1:.B33])
```

The outer `1×2` condition selects the nested branch in the first column and
the `33×1` reference in the second column. Consequently both selected
branches jointly produce a `33×2` result. The inner condition has only a
finite `2×1` source; its rows beyond that source correctly become `#N/A`.
The profile uses `max_array_cells = 66` and `max_stack_entries = 32`, so the
failure is from the condition cache rather than a large inline-array AST or
output admission. The selected reference itself is read once per row.

The earlier example that put a larger inline array in an unselected branch is
withdrawn: the accepted selected-branch-only shape profile correctly cannot
use that branch to widen the result.

## Finding

`condition_cache_set_shape` records one shape per condition node, but when a
previously cached condition at a smaller shape is probed at a widened demand,
`condition_cache_get` falls back to the widened `(demand, index)` key when
projection from the old source shape fails (around lines 2685–2707). The scalar
probe then caches the resulting scalar, including `#N/A`, through
`condition_cache_put` (around lines 3655–3720). The first such miss promotes
the node's canonical shape and rekeys the old entries (around lines
2626–2665); every later widened coordinate is then a distinct cache entry.

For the expression above, the inner `IF({TRUE()|TRUE()};1;0)` is first
cached at `2×2` for the initially selected positions. The selected reference
sibling widens the common demand to `33×2`. Rows three through thirty-three
probe the inner condition outside its finite source, produce `#N/A`, and are
inserted under the widened demand. The global cache reaches 33 entries and
refuses at the configured 32-entry limit. These entries are aliases or a
repeatable out-of-source result rather than distinct source cells; the
condition cache must not make an otherwise admissible finite matrix fail.

A bounded fix needs to retain an intrinsic source-shape key (with a bounded
outside-shape/sentinel result), or otherwise ensure that out-of-source
condition probes do not create one widened-demand entry per output position.
The fix must preserve selected-cell laziness and must not cache a value under
a source coordinate that was never valid.

## Cleared accounting properties

The reviewed key ordering remains internally consistent: entries are ordered
by `(node, demand rows, demand columns, index)`, and all entries for a node
are rekeyed together, so the binary-search invariant is preserved. The
rekeying pass charges the full existing cache length before `retain_mut`, and
shape-table lookup/insertion and cache insertion charge their binary-search
and shift work. No separate cancellation or uncharged retain pass was found
in this narrow review; `charge_work` checks execution before the pass.

The remaining disposition is therefore one concrete VP15-1 correctness and
resource blocker: widened-demand cache entries for finite/out-of-source
condition positions. The reference-geometry implementation and unrelated
planner paths are outside this artifact's scope.
