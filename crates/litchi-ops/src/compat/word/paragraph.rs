//! Paragraph wrapper, moved from the facade's `document::paragraph`.

use super::Run;
use crate::Result;

/// A paragraph of a word-processing document.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "moved unchanged from the facade, where boxing the large variant would have broken the public API"
)]
#[non_exhaustive]
pub enum Paragraph {
    /// A legacy Word (`.doc`) paragraph.
    #[cfg(feature = "doc")]
    Doc(litchi_doc::Paragraph),
    /// A WordprocessingML (`.docx`) paragraph.
    #[cfg(feature = "docx")]
    Docx(litchi_docx::Paragraph),
    /// A Rich Text Format paragraph.
    #[cfg(feature = "rtf")]
    Rtf(litchi_rtf::ParagraphContent<'static>),
    /// An OpenDocument Text paragraph.
    #[cfg(feature = "odt")]
    Odt(litchi_odt::elements::text::Paragraph),
}

impl Paragraph {
    /// The text content of the paragraph.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the paragraph's text cannot be
    /// decoded.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(paragraph) => Ok(paragraph.text()?.to_owned()),
            #[cfg(feature = "docx")]
            Self::Docx(paragraph) => Ok(paragraph.text()?.to_owned()),
            #[cfg(feature = "rtf")]
            Self::Rtf(paragraph) => Ok(paragraph.text()),
            #[cfg(feature = "odt")]
            Self::Odt(paragraph) => paragraph.text().map_err(super::odt_error),
        }
    }

    /// The runs of the paragraph.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the runs cannot be decoded.
    pub fn runs(&self) -> Result<Vec<Run>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(paragraph) => Ok(paragraph.runs()?.into_iter().map(Run::Doc).collect()),
            #[cfg(feature = "docx")]
            Self::Docx(paragraph) => Ok(paragraph.runs()?.into_iter().map(Run::Docx).collect()),
            #[cfg(feature = "rtf")]
            Self::Rtf(paragraph) => Ok(paragraph
                .runs()
                .iter()
                .map(|run| Run::Rtf(run.clone()))
                .collect()),
            #[cfg(feature = "odt")]
            Self::Odt(paragraph) => Ok(paragraph
                .runs()
                .map_err(super::odt_error)?
                .into_iter()
                .map(Run::Odt)
                .collect()),
        }
    }
}
