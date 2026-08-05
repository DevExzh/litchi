//! OOXML package boundary for WordprocessingML numbering.
//!
//! The semantic model and bounded XML codec live in
//! [`crate::numbering`]. This boundary only preprocesses the related
//! part with markup compatibility and maps owner errors into the host error
//! type.

use crate::error::Result;
use crate::numbering::Collection;
use litchi_opc::part::Part;

pub(crate) fn parse_part(part: &dyn Part) -> Result<Collection> {
    let xml = litchi_ooxml_common::mce::process_part(part)?;
    crate::numbering::parse_numbering(xml.as_ref())
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
}
