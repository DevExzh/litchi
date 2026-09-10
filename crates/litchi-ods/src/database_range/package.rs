//! Package-level `database-ranges` content replacement.

use super::codec::{self, Location};
use crate::model::database_range::Range;
use crate::package::Package;
use litchi_core::{Error, Result};

/// Replace the source-checked `database-ranges` owner and rebuild only the ODS package.
pub(crate) fn replace(
    package: &Package,
    source_xml: &str,
    location: &Location,
    ranges: Option<&[Range]>,
) -> Result<Vec<u8>> {
    if package.content_xml() != source_xml {
        return Err(Error::InvalidFormat(
            "ODS database-ranges source changed before commit".to_string(),
        ));
    }
    if !package.package().digital_signatures()?.is_empty() {
        return Err(Error::InvalidFormat(
            "changed ODS database-range publication requires explicit signature disposition"
                .to_string(),
        ));
    }
    let content_xml = codec::replace(source_xml, location, ranges)?;
    package
        .replace_content_xml(&content_xml)
        .map(Package::into_bytes)
}
