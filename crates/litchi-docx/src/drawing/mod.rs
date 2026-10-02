//! Typed inventory of `WordprocessingML` drawing anchors.
//!
//! The owner is intentionally small: `model` contains the owned semantic
//! inventory, `codec` contains the streaming `<w:drawing>`, `w:object`, and
//! `w:pict` scanners, and `validation` owns checked Word 2010 extension
//! values. `DrawingML` preset tokens remain the closed domain supplied by
//! [`litchi_drawingml::geom::Preset`]; unknown drawing children are inert.

mod codec;
mod model;
pub mod source;
mod validation;

#[cfg(test)]
mod source_tests;
#[cfg(test)]
mod tests;

pub use model::{Anchor, AnchorId, Kind, LegacyAnchor, LegacyAnchorKind, Object};
pub use source::{
    ByteRange, DrawingDialect, DrawingPlacement, ElementRange, NamespaceContext, PictureSource,
    RelationshipDialect, RelationshipReference, ScanLimits, SourceDrawing, SvgOwner, SvgOwnerState,
};

pub(crate) use codec::{parse, parse_legacy};

/// Collect the drawing objects of `paragraphs` in document order, recording a
/// [`crate::leniency::DrawingDefect`] (context: paragraph index) for each
/// object kept with an unrecognized or missing preset.
pub(crate) fn collect_document_objects<'a>(
    paragraphs: impl IntoIterator<Item = &'a crate::Paragraph>,
    tolerance: &mut crate::leniency::ToleranceReport,
) -> crate::error::Result<Vec<Object>> {
    let mut objects = Vec::new();
    for (index, paragraph) in paragraphs.into_iter().enumerate() {
        for object in paragraph.drawing_objects()? {
            match object.unknown_preset() {
                Some("") => {
                    tolerance.record(crate::leniency::DrawingDefect::MissingPreset, index as u64);
                },
                Some(_) => {
                    tolerance.record(crate::leniency::DrawingDefect::UnknownPreset, index as u64);
                },
                None => {},
            }
            objects.push(object);
        }
    }
    Ok(objects)
}
pub(crate) use validation::append_word2010_anchor_id;
