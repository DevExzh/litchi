//! Worksheet package ownership for smart-tag metadata.

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};

use super::codec::{parse, replace_worksheet};
use super::model::Collection;
use crate::error::{Result, invalid};

pub(crate) fn load(package: &OpcPackage, worksheet: &PackURI) -> Result<Option<Collection>> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    parse(part.blob())
}

/// Replace or remove the worksheet's `smartTags` element.
///
/// A replacement that leaves the worksheet bytes as they are (for example
/// removing absent smart tags) publishes nothing, so a signed package keeps
/// its signature byte for byte. Any other replacement changes the worksheet,
/// so a signed package is refused with [`crate::Error::Signed`] before
/// anything is changed.
pub(crate) fn store(
    package: &mut OpcPackage,
    worksheet: &PackURI,
    value: Option<&Collection>,
) -> Result<()> {
    let updated = {
        let part = package.get_part(worksheet)?;
        require_worksheet(part)?;
        let updated = replace_worksheet(part.blob(), value)?;
        if updated.as_slice() == part.blob() {
            return Ok(());
        }
        updated
    };
    crate::signature_policy::ensure_unsigned(package)?;
    package.get_part_mut(worksheet)?.set_blob(updated);
    Ok(())
}

fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::SML_WORKSHEET {
        Ok(())
    } else {
        Err(invalid(format!(
            "part '{}' is not a SpreadsheetML worksheet",
            part.partname()
        )))
    }
}
