//! Text-run wrapper, moved from the facade's `document::run`.

use litchi_core::VerticalPosition;

use crate::Result;

/// A text run of a paragraph.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "moved unchanged from the facade, where boxing the large variant would have broken the public API"
)]
#[non_exhaustive]
pub enum Run {
    /// A legacy Word (`.doc`) run.
    #[cfg(feature = "doc")]
    Doc(litchi_doc::Run),
    /// A WordprocessingML (`.docx`) run.
    #[cfg(feature = "docx")]
    Docx(litchi_docx::Run),
    /// A Rich Text Format run.
    #[cfg(feature = "rtf")]
    Rtf(litchi_rtf::Run<'static>),
    /// An OpenDocument Text span.
    #[cfg(feature = "odt")]
    Odt(litchi_odt::elements::text::Span),
}

impl Run {
    /// The text content of the run.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the run's text cannot be
    /// decoded.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(run) => Ok(run.text()?.to_owned()),
            #[cfg(feature = "docx")]
            Self::Docx(run) => Ok(run.text()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(run) => Ok(run.text().to_owned()),
            #[cfg(feature = "odt")]
            Self::Odt(run) => run.text().map_err(super::odt_error),
        }
    }

    /// Whether the run is bold, when the run states it.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the run properties cannot be read.
    pub fn bold(&self) -> Result<Option<bool>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(run) => Ok(run.bold()),
            #[cfg(feature = "docx")]
            Self::Docx(run) => Ok(run.bold()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(run) => Ok(run.bold()),
            #[cfg(feature = "odt")]
            Self::Odt(run) => Ok(run.bold()),
        }
    }

    /// Whether the run is italic, when the run states it.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the run properties cannot be read.
    pub fn italic(&self) -> Result<Option<bool>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(run) => Ok(run.italic()),
            #[cfg(feature = "docx")]
            Self::Docx(run) => Ok(run.italic()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(run) => Ok(run.italic()),
            #[cfg(feature = "odt")]
            Self::Odt(run) => Ok(run.italic()),
        }
    }

    /// Whether the run is underlined, when the run states it.
    ///
    /// DOCX reports a direct `w:u` (`Some(false)` for `w:val="none"`); DOC
    /// reports `Some(true)` for any underline style and `None` otherwise; RTF
    /// always states it, as it does bold and italic; ODT spans report `None`,
    /// since ODT formatting lives in styles.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the run properties cannot be read.
    pub fn underline(&self) -> Result<Option<bool>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(run) => Ok(run.underline()),
            #[cfg(feature = "docx")]
            Self::Docx(run) => Ok(run.underline()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(run) => Ok(Some(run.underline())),
            #[cfg(feature = "odt")]
            Self::Odt(_) => Ok(None),
        }
    }

    /// Whether the run is struck through, when the run states it.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the run properties cannot be read.
    pub fn strikethrough(&self) -> Result<Option<bool>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(run) => Ok(run.strikethrough()),
            #[cfg(feature = "docx")]
            Self::Docx(run) => Ok(run.strikethrough()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(run) => Ok(run.strikethrough()),
            #[cfg(feature = "odt")]
            Self::Odt(run) => Ok(run.strikethrough()),
        }
    }

    /// The run's vertical position (superscript or subscript), or `None`
    /// for the normal baseline.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the run properties cannot be read.
    pub fn vertical_position(&self) -> Result<Option<VerticalPosition>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(run) => Ok(match run.properties().vertical_position {
                VerticalPosition::Normal => None,
                position => Some(position),
            }),
            #[cfg(feature = "docx")]
            Self::Docx(run) => Ok(match run.vertical_position()? {
                Some(VerticalPosition::Superscript) => Some(VerticalPosition::Superscript),
                Some(VerticalPosition::Subscript) => Some(VerticalPosition::Subscript),
                Some(VerticalPosition::Normal) | None => None,
            }),
            #[cfg(feature = "rtf")]
            Self::Rtf(run) => Ok(run.vertical_position()),
            #[cfg(feature = "odt")]
            Self::Odt(run) => Ok(run.vertical_position()),
        }
    }
}
