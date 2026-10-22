//! Engine-level coverage of the pivot extension owners (record 0360).
//!
//! These suites build schema-shaped synthetic OPC graphs and drive both the
//! package routes (`Package::edit_pivot_table`, `edit_pivot_table_data`,
//! `edit_pivot_cache` and their `apply_*` replays) and the crate-private
//! raw-graph engine behind them, whose stale-source, limit and signature
//! refusals cannot be reached through an immutable package.

mod cached_unique_names;
mod server_formats;
mod table_data;
