//! Compatibility adapter for XLSB External Link package authoring.
//!
//! The owner crate emits the bounded BIFF12 stream. This small host-side
//! wrapper retains the historical OPC part and relationship API.

use crate::xlsb::error::{Error, Result};
use crate::xlsb::external_link::Link;
use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type;
use litchi_opc::part::{BlobPart, Part};

const EXTERNAL_LINK_CONTENT_TYPE: &str = "application/vnd.ms-excel.externalLink";
const MAX_EXTERNAL_LINK_PART_BYTES: usize = 32 * 1024 * 1024;

pub(crate) fn author_external_link_part(link: &Link, one_based_index: usize) -> Result<BlobPart> {
    link.validate()?;
    if one_based_index == 0 {
        return Err(Error::InvalidFormula(
            "external-link part index must be one-based".to_string(),
        ));
    }
    let mut part = BlobPart::new(
        PackURI::new(format!(
            "/xl/externalLinks/externalLink{one_based_index}.bin"
        ))?,
        EXTERNAL_LINK_CONTENT_TYPE.to_string(),
        Vec::new(),
    );
    let relationship_id = match link.kind() {
        crate::xlsb::external_link::Kind::Workbook => {
            Some(part.relate_to_ext(link.source(), relationship_type::EXTERNAL_LINK_PATH))
        },
        crate::xlsb::external_link::Kind::Dde => None,
        crate::xlsb::external_link::Kind::Ole => {
            Some(part.relate_to_ext(link.source(), relationship_type::OLE_OBJECT))
        },
    };
    let bytes =
        litchi_xlsb::external_link::write_external_link_stream(link, relationship_id.as_deref())?;
    if bytes.len() > MAX_EXTERNAL_LINK_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_EXTERNAL_LINK_PART_BYTES,
            found: bytes.len(),
        });
    }
    part.set_blob(bytes);
    Ok(part)
}
