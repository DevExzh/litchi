//! Text run implementation for Word documents.

use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A text run in a paragraph.
pub enum Run {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Run),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Run),
    #[cfg(feature = "iwa")]
    Pages(String),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Run<'static>),
}

impl Run {
    /// Get the text content of the run.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "ole")]
            Run::Doc(r) => r.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Run::Docx(r) => r.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "iwa")]
            Run::Pages(text) => Ok(text.clone()),
            #[cfg(feature = "rtf")]
            Run::Rtf(r) => Ok(r.text().to_string()),
        }
    }

    /// Check if the run is bold.
    pub fn bold(&self) -> Result<Option<bool>> {
        match self {
            #[cfg(feature = "ole")]
            Run::Doc(r) => Ok(r.bold()),
            #[cfg(feature = "ooxml")]
            Run::Docx(r) => r.bold().map_err(Error::from),
            #[cfg(feature = "iwa")]
            Run::Pages(_) => Ok(None), // Pages doesn't support run-level formatting in the current API
            #[cfg(feature = "rtf")]
            Run::Rtf(r) => Ok(Some(r.formatting.bold)),
        }
    }

    /// Check if the run is italic.
    pub fn italic(&self) -> Result<Option<bool>> {
        match self {
            #[cfg(feature = "ole")]
            Run::Doc(r) => Ok(r.italic()),
            #[cfg(feature = "ooxml")]
            Run::Docx(r) => r.italic().map_err(Error::from),
            #[cfg(feature = "iwa")]
            Run::Pages(_) => Ok(None), // Pages doesn't support run-level formatting in the current API
            #[cfg(feature = "rtf")]
            Run::Rtf(r) => Ok(Some(r.formatting.italic)),
        }
    }

    /// Check if the run is strikethrough.
    pub fn strikethrough(&self) -> Result<Option<bool>> {
        match self {
            #[cfg(feature = "ole")]
            Run::Doc(r) => Ok(r.strikethrough()),
            #[cfg(feature = "ooxml")]
            Run::Docx(r) => r.strikethrough().map_err(Error::from),
            #[cfg(feature = "iwa")]
            Run::Pages(_) => Ok(None), // Pages doesn't support run-level formatting in the current API
            #[cfg(feature = "rtf")]
            Run::Rtf(r) => Ok(Some(r.formatting.strike)),
        }
    }

    /// Get the vertical position of the run (superscript/subscript).
    ///
    /// Returns the vertical positioning if specified, None if normal.
    ///
    /// **Note**: This method requires the `ole` or `ooxml` feature to be enabled.
    #[cfg(any(feature = "ole", feature = "ooxml", feature = "iwa"))]
    pub fn vertical_position(&self) -> Result<Option<crate::common::VerticalPosition>> {
        use crate::common::VerticalPosition;

        match self {
            #[cfg(feature = "ole")]
            Run::Doc(r) => {
                let pos = match r.properties().vertical_position {
                    VerticalPosition::Normal => None,
                    pos => Some(pos),
                };
                Ok(pos)
            },
            #[cfg(feature = "ooxml")]
            Run::Docx(r) => {
                use crate::ooxml::docx::paragraph::VerticalPosition as OoxmlVerticalPosition;
                match r.vertical_position().map_err(Error::from)? {
                    Some(OoxmlVerticalPosition::Superscript) => {
                        Ok(Some(VerticalPosition::Superscript))
                    },
                    Some(OoxmlVerticalPosition::Subscript) => Ok(Some(VerticalPosition::Subscript)),
                    Some(OoxmlVerticalPosition::Normal) | None => Ok(None),
                }
            },
            #[cfg(feature = "iwa")]
            Run::Pages(_) => Ok(None), // Pages doesn't support run-level formatting in the current API
            #[cfg(feature = "rtf")]
            Run::Rtf(r) => {
                if r.formatting.superscript {
                    Ok(Some(VerticalPosition::Superscript))
                } else if r.formatting.subscript {
                    Ok(Some(VerticalPosition::Subscript))
                } else {
                    Ok(None)
                }
            },
        }
    }
}
