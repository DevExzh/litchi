# 0396 — isolated iWork performance-probe call counts

- Scope: repair only the selected-call contamination in the formula-refresh
  and W2-A Numbers probe families. Runtime, budgets, fixtures, and operation
  semantics are unchanged. The D7 cached-formula probe remains E-owned and is
  deliberately excluded from this WP.
- Base: frozen integrated candidate `1bc8f22499d2691e85ebed37686c8381f0c985da`.
- Owned source: `formula_refresh_cases.rs` and `numbers_w2a_cost.rs`.
- Source handoff commits: `dc585d05cd` (formula/W2-A wrapper isolation) and
  `5d18c05ddb` (separate Numbers cached-formula wrapper isolation after the
  temporary ownership transfer).

## Raw-audit findings

The retained audit
`/home/zhuhe/code/litchi-worktrees/.coord/resume-20261028/perf/final-1bc8f22499-iwork/capture/coordinator-probe-call-audit.json`
contains 137 legs. The paired host/leaf contract is two selected calls per
side; the one existing Pages hidden-axis leaf-only route is the sole intentional
one-call exception. This WP repairs:

- `w1f_formula_refresh_keynote` leaf setup discovery, which called the selected
  `leaf_refresh_keynote` symbol while searching for a drawable and recorded six
  calls instead of two;
- six W2-A reset routes, whose reset setup called selected
  `w2a_host_route_3` and recorded three host calls instead of two.

Each setup path now calls an uninstrumented setup function. The exact
selected wrappers remain the ABBA calls only; setup names deliberately do not
contain the selected wrapper symbol, so historical wildcard toggles cannot
collect setup.

## External audit validator

The external `validate_call_audit.py` revision validates the raw audit against
the supplied 69- or 87-route manifest. It rejects missing/extra/duplicate
legs, unknown route IDs, wrong side sets, and every call count other than two
for `host_leaf`/`unified_leaf` or one for `leaf_only`. There are no route-specific
exemptions. Historical runners, manifests, raw audits, and receipts remain
untouched; the validator writes only a new external receipt.

External validator SHA-256:
`56cf47937fade54867c6393bb5d6a561bcb943a01514e152e165d60f3988667f`.

New measurement-pipeline copies require `--call-audit` in measurement mode,
strip it before delegating to the historical parser, and validate the raw audit
against the normalized manifest after the parent run succeeds:

- 69-route runner:
  `iwork-owner-probe-final/qualify_iwork_owner.py`, SHA-256
  `4898d5d5308f9059897be7f68e40a0c0dacafc0f034834087103c132f6fc7c8d`;
- 87-route runner:
  `iwork-chart-font-probe-final/qualify_iwork_chart_font.py`, SHA-256
  `881b8cabd750f60ce53ab4c29ba916c1b937ec5e42ea484ab7ee516407c92699`.

Historical runners and manifests remain the immutable parents of these copies.

After ownership acknowledgement, the cached-formula unified probe received a
separate source commit `numbers_unified.rs` wrapper split: warm-up and output
checks call setup implementations whose names cannot match the measured
symbols, while the exact ABBA wrappers remain
`numbers_unified::unified_cached_formula_route` and
`numbers_unified::direct_cached_formula_route`. This is intentionally a
separate follow-up commit so the E-owned comment proof can merge independently.

No capture, timing, binary, native, or performance result is claimed by this
record. Final qualification must run on the repaired source after coordinator
review and integrated-source freeze.
