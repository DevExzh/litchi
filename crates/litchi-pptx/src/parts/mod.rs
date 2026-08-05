//! Low-level PresentationML part views.
//!
//! These wrappers own only vocabulary validation and relationship traversal.
//! They deliberately borrow the OPC graph so an opened package can retain
//! every unmodeled part and byte range until an explicit managed write.

mod presentation;
mod slide;

pub use presentation::{PresentationPart, SlideReference};
pub use slide::{SlideLayoutPart, SlideMasterPart, SlidePart};

use std::borrow::Cow;

use litchi_ooxml_common::mce::process_part;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, Part};
use quick_xml::events::BytesStart;
use quick_xml::reader::NsReader;

use crate::{Error, Result};

pub(crate) const MAX_PART_XML_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_SLIDES: usize = 100_000;

pub(crate) fn processed_xml(part: &dyn Part) -> Result<Cow<'_, [u8]>> {
    if part.blob().len() > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "PresentationML part XML",
            limit: MAX_PART_XML_BYTES,
        });
    }
    Ok(process_part(part)?)
}

pub(crate) fn relationship_attribute(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<Option<String>> {
    crate::namespace::relationship_attribute_value(
        element,
        b"id",
        reader.decoder(),
        reader.resolver(),
    )
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn parse_u32(value: &str, field: &str) -> Result<u32> {
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {field} value '{value}'")))
}

pub(crate) fn parse_i64(value: &str, field: &str) -> Result<i64> {
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {field} value '{value}'")))
}

pub(crate) fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(invalid(format!("invalid {field} value '{value}'"))),
    }
}

pub(crate) fn validate_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        return Ok(());
    }
    Err(Error::ContentType {
        expected: expected.to_string(),
        actual: part.content_type().to_string(),
    })
}

pub(crate) fn is_relationship_type(actual: &str, transitional: &str, local: &str) -> bool {
    actual == transitional
        || actual == format!("http://purl.oclc.org/ooxml/officeDocument/relationships/{local}")
}

pub(crate) fn find_related_part<'a>(
    package: &'a OpcPackage,
    source: &dyn Part,
    relationship_id: &str,
    relationship_type: &str,
    local_type: &str,
    content_type: &str,
) -> Result<&'a dyn Part> {
    let relationship = source
        .rels()
        .get(relationship_id)
        .ok_or_else(|| Error::Relationship(format!("missing relationship '{relationship_id}'")))?;
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "relationship '{relationship_id}' must be internal"
        )));
    }
    if !is_relationship_type(relationship.reltype(), relationship_type, local_type) {
        return Err(Error::Relationship(format!(
            "relationship '{relationship_id}' has unexpected type '{}'",
            relationship.reltype()
        )));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    validate_content_type(part, content_type)?;
    Ok(part)
}

pub(crate) fn related_part_by_type<'a>(
    package: &'a OpcPackage,
    source: &dyn Part,
    relationship_type: &str,
    local_type: &str,
    content_type: &str,
) -> Result<Option<&'a dyn Part>> {
    let mut matching = source.rels().iter().filter(|relationship| {
        is_relationship_type(relationship.reltype(), relationship_type, local_type)
    });
    let Some(relationship) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() {
        return Err(Error::Relationship(format!(
            "part '{}' has multiple '{local_type}' relationships",
            source.partname()
        )));
    }
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "'{local_type}' relationship must be internal"
        )));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    validate_content_type(part, content_type)?;
    Ok(Some(part))
}

pub(crate) fn expected_main_content_type(content_type: &str) -> bool {
    matches!(
        content_type,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    )
}
