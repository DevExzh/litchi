use crate::{RtfError, RtfResult};

/// Largest supported document gutter width, in twips.
pub const MAX_DOCUMENT_GUTTER_TWIPS: u32 = 31_680;

#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Passive document print-layout settings from the RTF header.
///
/// These values are retained for round-tripping only. This crate does not
/// alter gutter geometry or arrange logical pages for printing.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentPrintLayoutSettings {
    /// Whether facing pages were requested (`facingp`).
    pub facing_pages: bool,
    /// Whether inside/outside margins are mirrored (`margmirror`).
    pub mirror_margins: bool,
    /// Document-wide gutter width in twips (`gutter`).
    pub document_gutter_twips: Option<u32>,
    /// Whether a parallel/top gutter was requested (`gutterprl`).
    pub parallel_gutter: bool,
    /// Whether two logical pages per physical page were requested (`twoonone`).
    pub two_logical_pages_per_physical_page: bool,
}

impl DocumentPrintLayoutSettings {
    /// Validate values before installing or serializing these settings.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self
            .document_gutter_twips
            .is_some_and(|value| value > MAX_DOCUMENT_GUTTER_TWIPS)
        {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document gutter must be in 0..={MAX_DOCUMENT_GUTTER_TWIPS} twips"
            )));
        }
        Ok(())
    }
    /// Atomically replace the document-wide gutter width.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn set_document_gutter_twips(&mut self, value: Option<u32>) -> RtfResult<()> {
        let mut candidate = *self;
        candidate.document_gutter_twips = value;
        candidate.validate()?;
        *self = candidate;
        Ok(())
    }

    /// Return whether all print-layout settings were omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.facing_pages
            && !self.mirror_margins
            && self.document_gutter_twips.is_none()
            && !self.parallel_gutter
            && !self.two_logical_pages_per_physical_page
    }
}
