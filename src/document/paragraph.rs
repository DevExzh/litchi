//! Paragraph implementation for Word documents.

use super::Run;
use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A paragraph in a Word document.
#[derive(Debug, Clone)]
pub enum Paragraph {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Paragraph),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Paragraph),
    #[cfg(feature = "iwa")]
    Pages(String),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Paragraph),
}

impl Paragraph {
    /// Get the text content of the paragraph.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "ole")]
            Paragraph::Doc(p) => p.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Paragraph::Docx(p) => p.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "iwa")]
            Paragraph::Pages(text) => Ok(text.clone()),
            #[cfg(feature = "rtf")]
            Paragraph::Rtf(_p) => {
                // RTF paragraphs are just formatting info, text comes from runs
                Ok(String::new())
            },
        }
    }

    /// Get the runs in this paragraph.
    pub fn runs(&self) -> Result<Vec<Run>> {
        match self {
            #[cfg(feature = "ole")]
            Paragraph::Doc(p) => {
                let runs = p.runs().map_err(Error::from)?;
                Ok(runs.into_iter().map(Run::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            Paragraph::Docx(p) => {
                let runs = p.runs().map_err(Error::from)?;
                Ok(runs.into_iter().map(Run::Docx).collect())
            },
            #[cfg(feature = "iwa")]
            Paragraph::Pages(text) => {
                // Pages paragraphs are simple strings without run-level formatting
                // Return a single run with the entire text
                Ok(vec![Run::Pages(text.clone())])
            },
            #[cfg(feature = "rtf")]
            Paragraph::Rtf(_p) => {
                // RTF runs come from style blocks, not paragraphs
                Ok(Vec::new())
            },
        }
    }
}
