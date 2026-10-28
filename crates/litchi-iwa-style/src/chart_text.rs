//! Shared chart title, caption and axis-title wire engines (record 0375b).
//!
//! The package owners retain responsibility for selector resolution, native
//! graph ownership, transaction budgets and exact-source publication.  This
//! module is deliberately limited to the format-neutral generated-extension
//! codecs and the canonical inline caption graph.  Re-exporting the private
//! codec surface here keeps the three owners on one bounded wire
//! implementation without allowing native identifiers or package objects to
//! cross a public package API.

/// Lossless chart-title extension projection and rewrite.
pub mod title {
    pub use litchi_iwa_protos::keynote_chart_title_codec::*;
}

/// Lossless chart-axis-title extension projection and rewrite.
pub mod axis_title {
    pub use litchi_iwa_protos::keynote_chart_axis_title_codec::*;
}

/// Lossless chart-caption reference projection and rewrite.
pub mod caption {
    pub use litchi_iwa_protos::keynote_chart_caption_codec::*;
}

/// Canonical inline title/caption graph payload encoder.
pub mod caption_graph {
    pub use litchi_iwa_protos::keynote_chart_caption_graph_codec::*;
}
