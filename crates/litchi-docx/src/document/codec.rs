//! Bounded document-XML codecs for the main WordprocessingML document.

use crate::error::{Error, Result};
use crate::section::{Section, Sections};

/// Extract sections from the document XML.
///
/// Sections are defined by `<w:sectPr>` elements, which can appear
/// in two places:
/// 1. Inside `<w:pPr>` (paragraph properties) - defines a section break
/// 2. At the end of `<w:body>` - defines the last section
pub(super) fn extract_sections(xml_bytes: &[u8]) -> Result<Sections> {
    let mut sections_xml = Vec::new();
    crate::namespace::scan_word_element_ranges(xml_bytes, &[b"sectPr"], |_, start, length| {
        let start = usize::try_from(start)
            .map_err(|_| Error::InvalidFormat("section offset overflow".to_string()))?;
        let length = usize::try_from(length)
            .map_err(|_| Error::InvalidFormat("section length overflow".to_string()))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| Error::InvalidFormat("section range overflow".to_string()))?;
        let raw = xml_bytes.get(start..end).ok_or_else(|| {
            Error::InvalidFormat("section range is outside document XML".to_string())
        })?;
        sections_xml.push(Section::from_xml_bytes(raw.to_vec())?);
        Ok(())
    })?;

    // If no sections were found, create a default section
    if sections_xml.is_empty() {
        sections_xml.push(Section::from_xml_bytes(b"<w:sectPr/>".to_vec())?);
    }

    Ok(Sections::new(sections_xml))
}
