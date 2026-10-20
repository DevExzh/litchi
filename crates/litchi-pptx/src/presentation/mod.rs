//! Borrowed semantic view of one `PresentationML` package graph.
//!
//! The typed facade delegates XML scanning to `codec` and package
//! relationship traversal to `package`. Embedded resources and media
//! authoring remain available through their existing child modules.

mod codec;
mod model;
mod order;
pub(crate) mod owned_parts;
mod package;
pub(crate) mod source;
mod source_cross_copy;
mod source_removal;
pub(crate) mod transition;

#[cfg(test)]
mod tests;

pub mod embedded;
pub mod media;

pub use model::Presentation;

pub use source::{
    MAX_SOURCE_BACKED_SLIDE_BATCH, SourceBackedPresentation, SourceImage, SourceImageDescriptor,
    SourceImageTarget, SourceSlide, SourceSvgDescriptor, SourceSvgImage,
};

/// Source-backed per-slide modern-comment snapshots.
pub use source::modern_comments;
/// Source-backed slide related-part snapshots.
pub use source::related;
/// Source-backed scene retention diagnostics.
pub use source::scenes as scene;

/// Explicit access to the retained presentation root.
pub use source::raw;
/// Detached presentation section membership.
pub use source::sections;

/// The `firstSlideNum` of a `p:presentation` root (default 1), the number
/// slide-number fields count from (record 0279).
pub(crate) fn first_slide_number(xml: &[u8]) -> crate::Result<u32> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
        {
            Event::Start(root) | Event::Empty(root) => {
                for attribute in root.checked_attributes() {
                    let attribute =
                        attribute.map_err(|error| crate::Error::Xml(error.to_string()))?;
                    if attribute.key.local_name().as_ref() == b"firstSlideNum" {
                        return std::str::from_utf8(&attribute.value)
                            .ok()
                            .and_then(|value| value.trim().parse::<u32>().ok())
                            .ok_or_else(|| {
                                crate::Error::Xml("invalid presentation firstSlideNum".to_owned())
                            });
                    }
                }
                return Ok(1);
            },
            Event::Eof => return Err(crate::Error::Xml("empty presentation part".to_owned())),
            _ => {},
        }
    }
}
