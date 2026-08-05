//! Layered facade for the legacy PowerPoint writer core.
//!
//! Semantic state and caller-facing types live in [model], binary record
//! conversion lives in [codec], and OLE2 stream assembly lives in [package].

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{ShapeProperties, ShapeType, TextAlignment, WriteError, Writer};

// Internal names retained for the focused owner tests.
#[cfg(test)]
pub(super) use super::blip::Kind as PictureKind;
#[cfg(test)]
pub(super) use super::comments::SlideComment;
#[cfg(test)]
pub(super) use super::custom_shows::CustomShow;
#[cfg(test)]
pub(super) use super::hyperlink::{Hyperlink, HyperlinkCollection};
#[cfg(test)]
pub(super) use super::shape_style::{
    ArrowStyle, FillStyle, LineCapStyle, LineJoinStyle, LineStyle, LineStyleConfig, ShapeStyle,
};
#[cfg(test)]
pub(super) use super::smart_tags::SmartTagDefinition;
#[cfg(test)]
pub(super) use super::text_format::{FontEntity, Paragraph, TextAlign};
