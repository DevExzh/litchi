//! Bounded web-settings part ownership and relationship validation.

use super::super::CONTENT_TYPE;
use super::super::model::{Conformance, Settings};
use super::relationship::validate_frame_relationships;
use super::xml::process_web_xml;
use crate::{Error, Result};
use litchi_opc::Part;

impl Settings {
    fn read_part(part: &dyn Part, limits: &super::super::Limits) -> Result<(Self, Conformance)> {
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        if part.blob().len() > limits.xml_bytes() {
            return Err(Error::WebLimit {
                kind: crate::web::LimitKind::XmlBytes,
                actual: part.blob().len(),
                maximum: limits.xml_bytes(),
            });
        }
        // Keep the selected source policy on this one package read.  The
        // source scan precedes MCE processing; the second scan below proves
        // the expanded projection before the semantic model retains it.
        super::super::raw::admit_xml(part.blob(), limits)?;
        let xml = process_web_xml(part.blob(), limits)?;
        super::super::raw::admit_xml(xml.as_ref(), limits)?;
        let (settings, conformance) = Self::parse_xml(xml.as_ref())?;
        validate_frame_relationships(part, &settings, conformance)?;
        Ok((settings, conformance))
    }
}

/// Read one bounded web-settings part and validate its frame relationships.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
pub(crate) fn read(
    part: &dyn Part,
    limits: &super::super::Limits,
) -> Result<(Settings, Conformance)> {
    Settings::read_part(part, limits)
}
