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
`(deleted-coord-artifact)`
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

The external `validate_call_audit.py` revision derives the audit from the fresh
run's `results.json`, raw callgrind files, annotation files, source identity,
and binary-integrity records. It rejects stale results, capture hash/size
mismatches, missing/extra/duplicate legs, unknown route IDs, wrong side sets,
helper/setup collection, mismatched exact symbols, and every call count other
than two for `host_leaf`/`unified_leaf` or one for `leaf_only`. There are no
route-specific exemptions. Historical `--audit` input remains diagnostic-only;
the new 69/87 runners derive validation from their own fresh output directory.

External validator SHA-256:
`c0637e4304f5fdbfae6fca6b0e9a9c17f7d10cb9a56c7df6846951fa4bdf6b8e`.

Synthetic derived-audit controls (`test_derived_call_audit.py`) pass positive,
malformed-results, stale-source, and capture-hash-mismatch cases. Its SHA-256
is `dfd330302f395eb37169b61a9770b2c1c0109ab38bed48ada34dbfb7c4ef7a94`.

The coordinator's final validator also requires raw `summary` and `totals`
to agree. It sums the selected raw function's self and callee costs across
DWARF file contexts and requires that complete function cost to equal the
program total. Some annotations show both a complete function aggregate and
a smaller binary/file-context row; the latter must not be mistaken for the
whole function. A matching full-function annotation is required as additional
evidence. This preserves the measured totals and rejects collection outside
the selected wrapper; it does not normalize or discount instructions.
Controls cover split file contexts, outside-wrapper collection, missing raw
totals, and stale summaries in addition to the original refusal cases.

New measurement-pipeline copies require `--call-audit` in measurement mode,
strip it before delegating to the historical parser, and validate the raw audit
against the normalized manifest after the parent run succeeds:

- 69-route runner (derives audit after the delegated run):
  `iwork-owner-probe-final/qualify_iwork_owner.py`, SHA-256
  `ccb3374509cda8fc99a98dd51e059fc07ef0c40855ff26f2a59fb0894a2359aa`;
- 87-route runner (derives audit after the delegated run):
  `iwork-chart-font-probe-final/qualify_iwork_chart_font.py`, SHA-256
  `2fbe3746850e0fec6fa1ae92baf67984233dced6d38ad616b04223e88b3c2e9f`.

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
