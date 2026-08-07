//! Paragraph implementation for Word documents.

use super::Run;
#[cfg(any(feature = "doc", feature = "odt"))]
use litchi_core::Error;
use litchi_core::Result;

#[cfg(feature = "doc")]
use litchi_doc as doc;

/// A paragraph in a Word document.
#[derive(Debug, Clone)]
#[allow(clippy::large_enum_variant)] // public facade enum; boxing would break the API
pub enum Paragraph {
    #[cfg(feature = "doc")]
    Doc(doc::Paragraph),
    #[cfg(feature = "docx")]
    Docx(crate::docx::Paragraph),
    #[cfg(feature = "pages")]
    Pages(String),
    #[cfg(feature = "rtf")]
    Rtf(litchi_rtf::ParagraphContent<'static>),
    #[cfg(feature = "odt")]
    Odt(litchi_odt::elements::text::Paragraph),
}

impl Paragraph {
    /// Get the text content of the paragraph.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "doc")]
            Paragraph::Doc(p) => p.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "docx")]
            Paragraph::Docx(p) => p
                .text()
                .map(|s| s.to_string())
                .map_err(crate::map_ooxml_error),
            #[cfg(feature = "pages")]
            Paragraph::Pages(text) => Ok(text.clone()),
            #[cfg(feature = "rtf")]
            Paragraph::Rtf(p) => Ok(p.text()),
            #[cfg(feature = "odt")]
            Paragraph::Odt(p) => p
                .text()
                .map_err(|e| Error::ParseError(format!("Failed to get paragraph text: {}", e))),
        }
    }

    /// Get the runs in this paragraph.
    pub fn runs(&self) -> Result<Vec<Run>> {
        match self {
            #[cfg(feature = "doc")]
            Paragraph::Doc(p) => {
                let runs = p.runs().map_err(Error::from)?;
                Ok(runs.into_iter().map(Run::Doc).collect())
            },
            #[cfg(feature = "docx")]
            Paragraph::Docx(p) => {
                let runs = p.runs().map_err(crate::map_ooxml_error)?;
                Ok(runs.into_iter().map(Run::Docx).collect())
            },
            #[cfg(feature = "pages")]
            Paragraph::Pages(text) => {
                // Pages paragraphs are simple strings without run-level formatting
                // Return a single run with the entire text
                Ok(vec![Run::Pages(text.clone())])
            },
            #[cfg(feature = "rtf")]
            Paragraph::Rtf(p) => Ok(p.runs().iter().map(|r| Run::Rtf(r.clone())).collect()),
            #[cfg(feature = "odt")]
            Paragraph::Odt(p) => {
                let runs = p
                    .runs()
                    .map_err(|e| Error::ParseError(format!("Failed to get runs: {}", e)))?;
                Ok(runs.into_iter().map(Run::Odt).collect())
            },
        }
    }
}

#[cfg(all(test, any(all(feature = "docx", feature = "doc"), feature = "rtf")))]
mod tests {
    use super::super::Document;
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
    }

    #[test]
    #[cfg(all(feature = "docx", feature = "doc"))]
    fn test_paragraph_text_docx() {
        let path = test_data_path().join("ooxml/docx/FancyFoot.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let text = para.text().expect("Failed to get paragraph text");
            // Text should be extractable
            let _ = text.len();
        }
    }

    #[test]
    #[cfg(all(feature = "docx", feature = "doc"))]
    fn test_paragraph_text_doc() {
        let path = test_data_path().join("ole/doc/FancyFoot.doc");
        let doc = Document::open(&path).expect("Failed to open DOC");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let text = para.text().expect("Failed to get paragraph text");
            let _ = text.len();
        }
    }

    #[test]
    #[cfg(all(feature = "docx", feature = "doc"))]
    fn test_paragraph_runs_docx() {
        let path = test_data_path().join("ooxml/docx/FancyFoot.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let runs = para.runs().expect("Failed to get runs");
            for run in runs {
                let _text = run.text().expect("Failed to get run text");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "docx", feature = "doc"))]
    fn test_paragraph_runs_doc() {
        let path = test_data_path().join("ole/doc/FancyFoot.doc");
        let doc = Document::open(&path).expect("Failed to open DOC");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        for para in paragraphs {
            let runs = para.runs().expect("Failed to get runs");
            for run in runs {
                let _text = run.text().expect("Failed to get run text");
            }
        }
    }

    #[test]
    #[cfg(feature = "rtf")]
    fn test_paragraph_rtf() {
        // Use testUnicode.rtf which parses correctly
        let path = test_data_path().join("rtf/testUnicode.rtf");
        let doc = Document::open(&path).expect("Failed to open RTF");
        let paragraphs = doc.paragraphs().expect("Failed to get paragraphs");

        assert!(!paragraphs.is_empty(), "Expected at least one paragraph");

        for para in paragraphs {
            let text = para.text().expect("Failed to get paragraph text");
            let _ = text.len();
        }
    }
}
