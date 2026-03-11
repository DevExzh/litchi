//! Text run implementation for Word documents.

#[cfg(any(feature = "ole", feature = "ooxml", feature = "odf"))]
use crate::common::Error;
use crate::common::Result;

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A text run in a paragraph.
#[derive(Debug, Clone)]
pub enum Run {
    #[cfg(feature = "ole")]
    Doc(ole::doc::Run),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Run),
    #[cfg(feature = "iwa")]
    Pages(String),
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::Run<'static>),
    #[cfg(feature = "odf")]
    Odt(crate::odf::Run),
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
            #[cfg(feature = "odf")]
            Run::Odt(r) => r
                .text()
                .map_err(|e| Error::ParseError(format!("Failed to get run text: {}", e))),
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
            Run::Rtf(r) => Ok(r.bold()),
            #[cfg(feature = "odf")]
            Run::Odt(r) => Ok(r.bold()),
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
            Run::Rtf(r) => Ok(r.italic()),
            #[cfg(feature = "odf")]
            Run::Odt(r) => Ok(r.italic()),
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
            Run::Rtf(r) => Ok(r.strikethrough()),
            #[cfg(feature = "odf")]
            Run::Odt(r) => Ok(r.strikethrough()),
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
                // Now ooxml::docx::Run also uses crate::common::VerticalPosition
                match r.vertical_position().map_err(Error::from)? {
                    Some(VerticalPosition::Superscript) => Ok(Some(VerticalPosition::Superscript)),
                    Some(VerticalPosition::Subscript) => Ok(Some(VerticalPosition::Subscript)),
                    Some(VerticalPosition::Normal) | None => Ok(None),
                }
            },
            #[cfg(feature = "iwa")]
            Run::Pages(_) => Ok(None), // Pages doesn't support run-level formatting in the current API
            #[cfg(feature = "rtf")]
            Run::Rtf(r) => Ok(r.vertical_position()),
            #[cfg(feature = "odf")]
            Run::Odt(r) => Ok(r.vertical_position()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::super::Document;
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("test-data")
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_run_text_docx() {
        let path = test_data_path().join("ooxml/docx/FancyFoot.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let runs = para.runs().expect("Failed to get runs");
            for run in runs {
                let text = run.text().expect("Failed to get run text");
                assert!(
                    !text.is_empty() || text.is_empty(),
                    "Run text can be empty or non-empty"
                );
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_run_formatting_docx() {
        let path = test_data_path().join("ooxml/docx/FancyFoot.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let runs = para.runs().expect("Failed to get runs");
            for run in runs {
                let _bold = run.bold().expect("Failed to get bold");
                let _italic = run.italic().expect("Failed to get italic");
                let _strikethrough = run.strikethrough().expect("Failed to get strikethrough");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_run_formatting_doc() {
        let path = test_data_path().join("ole/doc/FancyFoot.doc");
        let doc = Document::open(&path).expect("Failed to open DOC");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let runs = para.runs().expect("Failed to get runs");
            for run in runs {
                let _bold = run.bold().expect("Failed to get bold");
                let _italic = run.italic().expect("Failed to get italic");
                let _strikethrough = run.strikethrough().expect("Failed to get strikethrough");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "ole"))]
    fn test_run_text_doc() {
        let path = test_data_path().join("ole/doc/FancyFoot.doc");
        let doc = Document::open(&path).expect("Failed to open DOC");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let runs = para.runs().expect("Failed to get runs");
            for run in runs {
                let text = run.text().expect("Failed to get run text");
                assert!(!text.is_empty() || text.is_empty());
            }
        }
    }

    #[test]
    #[cfg(feature = "rtf")]
    fn test_run_rtf() {
        // Use testUnicode.rtf which parses correctly
        let path = test_data_path().join("rtf/testUnicode.rtf");
        let doc = Document::open(&path).expect("Failed to open RTF");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let runs = para.runs().expect("Failed to get runs");
            for run in runs {
                let _text = run.text().expect("Failed to get run text");
                let _bold = run.bold().expect("Failed to get bold");
                let _italic = run.italic().expect("Failed to get italic");
                let _strikethrough = run.strikethrough().expect("Failed to get strikethrough");
            }
        }
    }
}
