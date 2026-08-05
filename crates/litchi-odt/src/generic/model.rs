//! Typed snapshots for format-neutral ODF package and flat-document access.

use crate::core::OwnedPackage;

/// Standard packaged OpenDocument document family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Family {
    /// Text document or template.
    Text,
    /// Spreadsheet document or template.
    Spreadsheet,
    /// Presentation document or template.
    Presentation,
    /// Drawing document or template.
    Drawing,
    /// Standalone chart document or template.
    Chart,
    /// Mathematical formula document or template.
    Formula,
    /// Image document or template.
    Image,
    /// Text master document or template.
    Master,
    /// Legacy producer-specific web-oriented text template.
    Web,
    /// Database front-end document.
    Database,
}

/// Validated, format-neutral OpenDocument package.
///
/// This provides lossless package access for every standard OpenDocument
/// family, including document types that do not yet have a specialized object
/// model. Saving an unmodified package returns the original bytes exactly.
pub struct Package {
    pub(super) package: OwnedPackage,
    pub(super) family: Family,
    pub(super) template: bool,
    pub(super) mimetype: String,
}

/// Validated flat OpenDocument XML file.
///
/// Flat documents combine content, styles, settings, and metadata under one
/// `office:document` root and are conventionally stored as `.fodt`, `.fods`,
/// `.fodp`, `.fodg`, `.fodc`, or `.fodi`. The `.fodf` extension is also
/// accepted for compatibility with odfdo's non-standard `office:formula`
/// convention; conforming packaged `.odf` formulas use a direct MathML root.
pub struct FlatDocument {
    pub(super) xml: String,
    pub(super) family: Family,
    pub(super) mimetype: String,
}
