//! Layered ODT master-page header/footer semantics.
//!
//! The owner facade keeps master-page content at the short contextual path and
//! groups page-layout property vocabulary below `properties`. XML conversion
//! and package/flat-document adapters live in their respective layers.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::MasterRegion as Region;
pub use model::{
    Block, Child, ChildKind, Column, Field, FieldKind, Inline, Kind, Master, SenderKind,
};

pub use codec::parse_page_layout_header_footer_properties;
pub(crate) use codec::{read, replace_range, set_text, set_xml};

/// Master-page header/footer content vocabulary.
pub mod content {
    pub use super::model::{Block, Column, Field, FieldKind, Inline, SenderKind};
}

/// Master-page structural vocabulary.
pub mod master {
    pub use super::model::MasterRegion as Region;
    pub use super::model::{Child, ChildKind, Kind, Master};
}

/// Page-layout header/footer property vocabulary and XML operations.
pub mod properties {
    pub use super::codec::parse_page_layout_header_footer_properties;
    pub(crate) use super::codec::{parse_region_properties, replace_page_layout_region_properties};
    pub use super::model::{
        Border, BorderLineWidth, BorderStyle, Color, Edges, Length, Properties, Region, Shadow,
        StyleProperties,
    };
}

pub(super) const MAX_VALUE: usize = 4096;

pub(super) fn bad(message: impl Into<String>) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(message.into())
}
