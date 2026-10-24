# 0362dvb — DrawingML owner policy and producer-limit return

This return records the final substantive repairs after the 0362dv review.
The source commit is `f40464c502d7ae17e94cbd0bdf58b8e78bf9b5c9`; exact
commands, lock provenance, and external log hashes are in
[`results/0362dvb/drawingml-owner-policy.json`](results/0362dvb/drawingml-owner-policy.json).

## Repairs

- `ink::actions::Edit::commit` is the only editor publication route. Fresh
  authoring still uses `Draft::finish`; editor `finish`/`prepare` spellings
  remain absent. `Prepared::edit` and `edit_with_options` retain the admitting
  `EditOptions` policy through the Arc-shared profile and reject wider child
  policies.
- Unchanged-owner-policy no-op commits reuse cached admission facts in O(1).
  Narrower child policies perform the bounded validation they require.
  Changed `Prepared::apply` checks the recipient output bound before copying
  patch bytes, then reads back and validates the recipient policy.
- Action, chart, transform, theme-family, model3d, SVG, InkML, and primitive
  coordinate producers preserve typed `LimitKind`, observed value, and
  maximum at real budget boundaries. Invalid lexical values remain invalid;
  they are not mislabeled as budget exhaustion. Stale immutable-owner replay
  returns typed `PatchConflict`.
- The explicit transform projection alias was removed. No compatibility alias
  or deprecated forwarding method was added.

## Validation

The final focused command used the pinned workspace lock (`52ed0c…`), the
repository capped shim, one build job, offline locked dependencies, zero debug
profiles, and an external target under 20 GB. It passed 237 library tests and
123 focused integration tests. The receipt distinguishes three pre-commit
qualification failures (all fixed before the final run) from the final result,
and preserves the earlier standalone `unified-chart-add --offline --locked`
failure plus its offline local-package-only lock repair.

No build target, binary, or evidence file is committed.
