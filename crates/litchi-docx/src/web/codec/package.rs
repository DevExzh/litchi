//! Bounded web-settings part ownership and relationship validation.

use super::super::model::{Conformance, Settings};
use super::super::{CONTENT_TYPE, MAX_XML_BYTES, invalid};
use super::relationship::validate_frame_relationships;
use super::xml::process_web_xml;
use crate::{Error, Result};
use litchi_opc::Part;

impl Settings {
    fn read_part(part: &dyn Part) -> Result<(Self, Conformance)> {
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        if part.blob().len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        let xml = process_web_xml(part.blob())?;
        let (settings, conformance) = Self::parse_xml(xml.as_ref())?;
        validate_frame_relationships(part, &settings, conformance)?;
        Ok((settings, conformance))
    }
}

/// Read one bounded web-settings part and validate its frame relationships.
pub fn read(part: &dyn Part) -> Result<(Settings, Conformance)> {
    Settings::read_part(part)
}
