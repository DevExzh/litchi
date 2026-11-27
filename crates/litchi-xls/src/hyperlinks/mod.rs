//! Layered, inert BIFF8 worksheet hyperlink retention.
//!
//! The facade exposes contextual hyperlink values, while the typed model,
//! BIFF8 codec, worksheet record linkage, and regression fixtures each live
//! in their own layer.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse_hlink_record;
pub use model::{
    FileMoniker, Hyperlink, HyperlinkMoniker, HyperlinkRange, HyperlinkTargetKind, ItemMoniker,
    RECORD_TYPE, TOOLTIP_RECORD_TYPE, UrlMoniker,
};
pub(crate) use package::HyperlinkCollector;
