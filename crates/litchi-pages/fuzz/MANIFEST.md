# Pages fuzz manifest

`Cargo.toml` registers two owner-level targets:

| Registered target | Harness | Seed directory |
| --- | --- | --- |
| `pages_body_drawable_order` | `fuzz_targets/pages_body_drawable_order.rs` | `corpus/pages_body_drawable_order/` |
| `pages_body_table_hidden_axes` | `fuzz_targets/pages_body_table_hidden_axes.rs` | `corpus/pages_body_table_hidden_axes/` |

The harness treats each bounded input as a fixture descriptor. It constructs a
small exact Pages package, so valid body drawable orders and native body-
storage slots remain reachable even when libFuzzer mutates the descriptor.
Corpus entries are not copied native packages and generated artifacts must not
be checked in.

`pages_body_table_hidden_axes` constructs a bounded two-table native graph for
each descriptor and probes selector-driven set/clear/reset/no-op transactions,
absent-owner nonempty-set refusal and existing-owner edits, stale
filtered/pivot state, locking,
shared ownership, malformed graph and wire records, data-dependent limits,
exact patch apply/inverse/conflict behavior, selected-table graph locality, and
strict nested codec resource ceilings. Valid descriptors can select a
non-identity canonical 6267 UID permutation; malformed descriptors also cover
missing/wrong model field-46 and table-info field-6 metadata plus unsupported
filter dependency metadata. Count-overflow metadata is expected to stop at
the bounded `WireFields` limit before allocation; the other malformed
root/graph/wire/UID/metadata cases report `InvalidSource` when ingress admits
the fixture.
Its deterministic corpus covers valid semantic, malformed, and boundary-limit
modes; generated artifacts stay outside the checkout.

Inspect the registered targets without starting a campaign:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_body_drawable_order
cargo +nightly fuzz check pages_body_table_hidden_axes
```
