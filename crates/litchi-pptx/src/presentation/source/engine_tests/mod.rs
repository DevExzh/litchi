//! Engine-level tests of the source-backed slide and SVG publication routes
//! the `source::Package` families run on (moved from the crate's integration
//! tests when the routes became crate-private).

mod cross_copy;
mod edit;
mod slide_removal;
mod svg_lifecycle;
mod svg_transaction;
