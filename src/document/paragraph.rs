//! Paragraph implementation for Word documents.

use crate::common::{Error, Result};
use super::Run;

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A paragraph in a Word document.
pub enum Paragraph {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Paragraph),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Paragraph),
}

impl Paragraph {
    /// Get the text content of the paragraph.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "ole")]
            Paragraph::Doc(p) => p.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Paragraph::Docx(p) => p.text().map(|s| s.to_string()).map_err(Error::from),
        }
    }

    /// Get the runs in this paragraph.
    pub fn runs(&self) -> Result<Vec<Run>> {
        match self {
            #[cfg(feature = "ole")]
            Paragraph::Doc(p) => {
                let runs = p.runs().map_err(Error::from)?;
                Ok(runs.into_iter().map(Run::Doc).collect())
            }
            #[cfg(feature = "ooxml")]
            Paragraph::Docx(p) => {
                let runs = p.runs().map_err(Error::from)?;
                Ok(runs.into_iter().map(Run::Docx).collect())
            }
        }
    }
}

