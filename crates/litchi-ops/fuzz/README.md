
Query plans: `cargo run --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bin
smoke_queries` runs deterministic mutations of harvested selectors across three
fake format families. The `query_plan` libFuzzer target also compiles and runs
plans with bounded output. This validates the engine, not production adapters.
