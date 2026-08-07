//! Physical package ownership for authored XLSB external-link parts.

use crate::external_link::{Kind, Link};
use crate::package::error::{Error, Result};
use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type;
use litchi_opc::part::{BlobPart, Part};

const CONTENT_TYPE: &str = "application/vnd.ms-excel.externalLink";
const MAX_PART_BYTES: usize = 32 * 1024 * 1024;

/// Author one validated external-link part and its external relationship.
///
/// # Errors
///
/// Returns an error when the link is invalid, the part index is not one-based,
/// the package URI cannot be constructed, or the bounded BIFF12 payload cannot
/// be serialized.
pub(crate) fn author_part(link: &Link, one_based_index: usize) -> Result<BlobPart> {
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
        CONTENT_TYPE.to_string(),
        Vec::new(),
    );
    let relationship_id = match link.kind() {
        Kind::Workbook => {
            Some(part.relate_to_ext(link.source(), relationship_type::EXTERNAL_LINK_PATH))
        },
        Kind::Dde => None,
        Kind::Ole => Some(part.relate_to_ext(link.source(), relationship_type::OLE_OBJECT)),
    };
    let bytes = crate::external_link::write_external_link_stream(link, relationship_id.as_deref())?;
    if bytes.len() > MAX_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_PART_BYTES,
            found: bytes.len(),
        });
    }
    part.set_blob(bytes);
    Ok(part)
}
