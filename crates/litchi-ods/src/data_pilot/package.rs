//! Archive-level `DataPilot` content replacement.

use super::codec::{self, Location};
use crate::model::data_pilot::Table;
use crate::package::Archive;
use litchi_core::{Error, Result};
use litchi_odf_common::XmlStreamLimits;

/// Replace the source-checked `DataPilot` owner and rebuild only the ODS package.
pub(crate) fn replace(
    package: &Archive,
    source_xml: &str,
    location: &Location,
    tables: Option<&[Table]>,
    limits: &XmlStreamLimits,
) -> Result<Vec<u8>> {
    if package.content_xml() != source_xml {
        return Err(Error::InvalidFormat(
            "ODS DataPilot source changed before commit".to_string(),
        ));
    }
    let content_xml = codec::replace(source_xml, location, tables, limits)?;
    package
        .replace_content_xml(&content_xml)
        .map(Archive::into_bytes)
}
