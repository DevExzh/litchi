//! OOXML package boundary for WordprocessingML numbering.
//!
//! The semantic model and bounded XML codec live in
//! [`crate::numbering`]. Numbering snapshots retain the authored part bytes so
//! extension edits can publish a source-preserving candidate.

use crate::error::Result;
use crate::numbering::{Collection, Snapshot};
use litchi_opc::part::Part;

pub(crate) fn parse_part(part: &dyn Part) -> Result<Collection> {
    Ok(Snapshot::from_xml(part.blob().to_vec())?
        .collection()
        .clone())
}

pub(crate) fn parse_snapshot_part(part: &dyn Part) -> Result<Snapshot> {
    Snapshot::from_xml(part.blob().to_vec())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbering::Format;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    #[test]
    fn package_boundary_returns_the_owner_collection() {
        let part = BlobPart::new(
            PackURI::new("/word/numbering.xml").expect("valid numbering URI"),
            "application/xml".to_owned(),
            br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#.to_vec(),
        );
        let value = parse_part(&part).expect("valid numbering part");

        assert_eq!(value.abstract_num_count(), 1);
        assert_eq!(value.num_count(), 1);
        assert_eq!(value.abstract_nums()[0].levels()[0].format, Format::Decimal);
    }

    #[test]
    fn package_boundary_retains_raw_extension_bytes_for_snapshot_edits() {
        let source = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w12="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w12"><w:abstractNum w:abstractNumId="1" w12:restartNumberingAfterBreak="0"><x:future xmlns:x="urn:future"/></w:abstractNum></w:numbering>"#;
        let part = BlobPart::new(
            PackURI::new("/word/numbering.xml").expect("valid numbering URI"),
            "application/xml".to_owned(),
            source.to_vec(),
        );

        let value = parse_part(&part).expect("valid numbering part");
        assert_eq!(
            value.abstract_nums()[0].restart_numbering_after_break(),
            Some(false)
        );
        let snapshot = parse_snapshot_part(&part).expect("source-preserving snapshot");
        assert_eq!(snapshot.xml_bytes(), source);
        let mut edit = snapshot.edit();
        edit.set_restart_numbering_after_break(1, Some(true))
            .expect("valid edit");
        let committed = edit.commit().expect("commit");
        assert!(
            std::str::from_utf8(committed.snapshot().xml_bytes())
                .expect("UTF-8 numbering XML")
                .contains(r#"w12:restartNumberingAfterBreak="true""#)
        );
    }
}
