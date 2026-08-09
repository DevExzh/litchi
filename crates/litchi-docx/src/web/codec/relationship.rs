#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Relationship-aware validation and XML relationship attribute decoding.

use super::super::model::{Child, Conformance, Frameset, Settings};
use crate::{Error, Result};
use litchi_opc::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};

pub(super) fn validate_frame_relationships(
    part: &dyn Part,
    settings: &Settings,
    conformance: Conformance,
) -> Result<()> {
    const FRAME_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame";
    const STRICT_FRAME_RELATIONSHIP: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/frame";

    fn validate(part: &dyn Part, frameset: &Frameset, expected_type: &str) -> Result<()> {
        for child in &frameset.children {
            match child {
                Child::Frameset(nested) => {
                    validate(part, nested, expected_type)?;
                },
                Child::Frame(frame) => {
                    let Some(id) = &frame.source_file_relationship_id else {
                        continue;
                    };
                    let relationship = part.rels().get(id).ok_or_else(|| {
                        Error::Invalid(format!("frame source relationship '{id}' does not exist"))
                    })?;
                    if relationship.reltype() != expected_type {
                        return Err(Error::Invalid(format!(
                            "frame source relationship '{id}' has an invalid type"
                        )));
                    }
                },
            }
        }
        Ok(())
    }

    if let Some(frameset) = &settings.frameset {
        let expected = match conformance {
            Conformance::Transitional => FRAME_RELATIONSHIP,
            Conformance::Strict => STRICT_FRAME_RELATIONSHIP,
        };
        validate(part, frameset, expected)?;
    }
    Ok(())
}

pub(super) fn required_relationship_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    const RELATIONSHIPS: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(namespace))
                if namespace == RELATIONSHIPS || namespace == STRICT_RELATIONSHIPS
        );
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(Error::Invalid(
                "duplicate frame source relationship ID".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    value.ok_or_else(|| Error::Invalid("frame source relationship ID is required".into()))
}
