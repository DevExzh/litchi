//! Bounded document-XML codecs for the main `WordprocessingML` document.

use crate::error::Result;
use crate::section::{Section, Sections};

/// Extract sections from the document XML.
///
/// Sections are defined by `<w:sectPr>` elements, which can appear
/// in two places:
/// 1. Inside `<w:pPr>` (paragraph properties) - defines a section break
/// 2. At the end of `<w:body>` - defines the last section
pub(super) fn extract_sections(part: &dyn litchi_opc::Part) -> Result<Sections> {
    let mut sections_xml = crate::section::footnote_columns::active_sections(part)?;

    // If no sections were found, create a default section
    if sections_xml.is_empty() {
        sections_xml.push(Section::from_xml_bytes(b"<w:sectPr/>".to_vec())?);
    }

    Ok(Sections::new(sections_xml))
}
