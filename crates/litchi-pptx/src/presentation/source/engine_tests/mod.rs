//! Engine-level tests of the source-backed slide and SVG publication routes
//! the `source::Package` families run on (record 0356a: moved from the
//! crate's integration tests when the routes became crate-private).

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

mod cross_copy;
mod edit;
mod slide_removal;
mod svg_lifecycle;
mod svg_transaction;
