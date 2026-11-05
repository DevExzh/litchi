# 0359qa — DOC body-toggle profile attribution correction

- **Status:** correction on `uo2/doc-toggle`; parent merge pending.
- **Date:** 2026-10-22.
- **Input:** review of [0359q](0359q-doc-body-toggle.md)'s shared-host
  Callgrind and Massif observations.
- **Scope:** the W5 probe selector and the attribution of its existing external
  profiles. No format implementation or public API changed.

## Selector repair

Before this correction, `legacy_w5_source_probe` computed
`doc_batch_leaf_write(&bytes, 1_000)` at the expected-output statement before
`Probe::time_rounds` could reject the scenario. Consequently a run with
`LITCHI_PROBE_ONLY=doc/leaf_write` still paid for one 1,000-toggle batch seed.
The `doc/batch1k_leaf_write` and `doc/leaf_write` expected-output calls now sit
inside their respective `selected` guards. A selected route still computes its
expected bytes and asserts every timed call against them; an unselected route
does no expected-output work. Both before and after binaries must be rebuilt
from this same probe source before a new profile is accepted.

The source control is
`crates/litchi-ops/examples/legacy_w5_source_probe.rs` at the 0359q source
commit `2452934481`, with the selector change in this correction commit. The
relevant control is
`Probe::time_rounds`'s two warmups and one measured call when
`LITCHI_PROBE_BATCH_ROUNDS=1`; the old `doc/leaf_write` profile therefore has
four route calls: the expected call, two warmups and one measured call. The
fixture `doc(1_000)` is built before the route and is outside the route node.

## Route-local Callgrind attribution

The existing external captures were re-read with
`callgrind_annotate --auto=no --inclusive=yes --tree=calling --show-percs=no
--threshold=99`. The `doc/leaf_write` node is the ordinary four-call route;
the following are inclusive `Ir` counts for those four calls, excluding
fixture construction and startup:

| Route operation | Calls | Before Ir | After Ir | After / before |
|---|---:|---:|---:|---:|
| `routes::doc_leaf_write` | 4 | 266,645,421 | 263,503,270 | **0.988216** |
| package open (before `from_bytes_with_options`; after `clone_bounded_source` + `from_owned_bytes`) | 4 | 64,774,129 | 65,459,941 | 1.010588 |
| `Package::edit` | 4 | 93,006,172 | 89,183,308 | 0.958897 |
| `body_text::Edit::set_character_property_override` | 4 | 32,366,281 | 22,452,250 | 0.693693 |
| `litchi_doc::package::edit::Edit::commit` | 4 | 73,491,486 | 83,336,939 | 1.133967 |
| `Package::write_to` | 4 | 2,349,978 | 2,414,429 | 1.027426 |

The after open count is `680,029` for `clone_bounded_source` plus `64,779,912`
for `Package::from_owned_bytes`. The route node's unchanged output SHA is
`9c1cc5368f99a988f1280446fcf43ec9485a1a88ff6e7245b26038676718731a` on both
legs. These counts show that the ordinary full commit-and-write route is
0.988216 in Callgrind; the old 0.320 row cannot stand in for it.

The old single-leg Callgrind files remain outside the repository as:

```text
(deleted-coord-artifact)/before-single.callgrind
  485,720 B  sha256 f61deedfd92bf21abf3f3ece78440538558a169837a2670c945dcca48c4d453f
(deleted-coord-artifact)/after-single.callgrind
  469,665 B  sha256 3904a436bf05771dd03bb73b2dcea583bbbc3d10231d9db99043c33b94a67fb5
```

The corresponding compact summary remains
`(deleted-coord-artifact)/profile-summary.txt`
(1,064 B, sha256
`5c66fa441092b06e6d8355c5a542eddc0ee3db3dcffbee148051092e4db9efd1`).

## Massif qualification boundary

The old `doc/leaf_write` Massif captures do not provide an ordinary-route
peak. Their maximum totals were 104,778,760 B before and 14,012,920 B after,
but both peak trees are attributed to `routes::doc_batch_leaf_write`; neither
peak tree contains `routes::doc_leaf_write`. The values retain the historical
observation and cannot be separated into fixture and ordinary-route heap from
these captures. The external files are:

```text
(deleted-coord-artifact)/before-leaf_write.massif
  812,109 B  sha256 8faa46ae90e493b49ed3cf891f2ef69db7d8a2ee0f5b45fa4bfe4de897326c6c
(deleted-coord-artifact)/after-leaf_write.massif
  299,874 B  sha256 8bf626c956ea00ce201e4ff4882d733f9029045b6107d7b89a7ec69bd3e0e9cb
```

The shared-host native ordinary-route timing, 3,106.697 µs before versus
3,684.751 µs after (+18.6%), remains diagnostic only. A quiet matched native
run and an isolated ordinary-route heap measurement are still required when a
build slot is available; this correction does not waive either qualification.
