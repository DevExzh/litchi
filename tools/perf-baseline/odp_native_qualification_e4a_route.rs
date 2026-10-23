//! Adapter selected by the historical e4a ODP probe.
//!
//! The e4a checkout predates `Commit`/`Patch` on the source-backed append
//! family.  Keep its measured lifecycle in one reviewed route source instead
//! of making the baseline compile the current candidate probe.  The include
//! remains intentionally external to this perf helper crate: it is consumed
//! by `crates/litchi-odp/examples/odp_0362ye_probe.rs` in the pinned e4a
//! checkout through `ODP_PROBE_ROUTE`.

include!("../../docs/unified-ops/results/0362ye/route_before.rs");
