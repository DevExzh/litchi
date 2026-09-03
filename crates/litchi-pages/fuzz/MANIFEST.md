# Pages fuzz manifest

`Cargo.toml` registers one owner-level target:

| Registered target | Harness | Seed directory |
| --- | --- | --- |
| `pages_body_drawable_order` | `fuzz_targets/pages_body_drawable_order.rs` | `corpus/pages_body_drawable_order/` |

The harness treats each bounded input as a fixture descriptor. It constructs a
small exact Pages package, so valid body drawable orders and native body-
storage slots remain reachable even when libFuzzer mutates the descriptor.
Corpus entries are not copied native packages and generated artifacts must not
be checked in.

Inspect the registered target without starting a campaign:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_body_drawable_order
```
