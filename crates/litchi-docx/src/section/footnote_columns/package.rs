//! Package-facing discovery and section integration for `footnoteColumns`.

use crate::error::{Error, Result};
use crate::namespace::scan_word_element_ranges;
use litchi_opc::part::Part;

use super::transaction::Snapshot;

/// Parse every `w:sectPr` found in a Word main-document part.
///
/// The source part is scanned as authored bytes so an ignorable Word 2012
/// child is not discarded before the focused owner can snapshot it. Namespace
/// context inherited from the document root is handled by the bounded prefix
/// fallback in the section codec.
pub fn parse_part(part: &dyn Part) -> Result<Vec<Snapshot>> {
    let xml = part.blob();
    let mut snapshots = Vec::new();
    scan_word_element_ranges(xml.as_ref(), &[b"sectPr"], |_, start, length| {
        let start = usize::try_from(start)
            .map_err(|_| Error::InvalidFormat("section XML offset overflow".into()))?;
        let length = usize::try_from(length)
            .map_err(|_| Error::InvalidFormat("section XML length overflow".into()))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| Error::InvalidFormat("section XML range overflow".into()))?;
        let bytes = xml
            .get(start..end)
            .ok_or_else(|| Error::InvalidFormat("section XML range is outside its part".into()))?;
        snapshots.push(Snapshot::from_xml(bytes.to_vec())?);
        Ok(())
    })?;
    Ok(snapshots)
}
