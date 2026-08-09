//! OPC relationship and part lifecycle for presentation properties.

use super::model::Properties;
use super::{CONTENT_TYPE, REL, STRICT_REL};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<Properties>> {
    let Some(uri) = find_properties_part(package)? else {
        return Ok(None);
    };
    let part = package.get_part(&uri)?;
    let mut value = Properties::parse(part.blob())?;
    hydrate_html_target(part, &mut value)?;
    Ok(Some(value))
}

/// Read only the typed document-level math defaults.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_math_from_package(
    package: &OpcPackage,
) -> Result<Option<crate::presentation_properties::math::Properties>> {
    Ok(load_from_package(package)?.and_then(|value| value.math().cloned()))
}

/// Replace the typed document-level math defaults as one package transaction.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn put_math_to_package(
    package: &mut OpcPackage,
    value: crate::presentation_properties::math::Properties,
) -> Result<Option<crate::presentation_properties::math::Properties>> {
    value.validate()?;
    let previous = load_math_from_package(package)?;
    if previous.as_ref() == Some(&value) {
        return Ok(previous);
    }

    let mut staged = package.clone();
    let properties = find_properties_part(&staged)?;
    let mut snapshot = if properties.is_some() {
        load_from_package(&staged)?.ok_or_else(|| {
            invalid("presentation-properties relationship disappeared during staging")
        })?
    } else {
        Properties::default()
    };
    snapshot.replace_math(value.clone());
    let strict = properties
        .as_ref()
        .map(|uri| is_strict_root(staged.get_part(uri)?.blob()))
        .transpose()?
        .unwrap_or(false);
    let blob = snapshot.to_xml(strict)?;
    let parsed = Properties::parse(&blob)?;
    if parsed.math() != Some(&value) {
        return Err(invalid("staged presentation math did not round-trip"));
    }

    if let Some(uri) = properties {
        staged.get_part_mut(&uri)?.set_blob(blob);
    } else {
        let uri = staged.next_partname("/ppt/presProps%d.xml")?;
        let presentation = staged.main_document_part()?.partname().clone();
        staged.try_add_part(Box::new(BlobPart::new(
            uri.clone(),
            ct::PML_PRES_PROPS.to_owned(),
            blob,
        )))?;
        let target = uri.relative_ref(presentation.base_uri());
        staged
            .get_part_mut(&presentation)?
            .relate_to(&target, rt::PRES_PROPS);
    }
    staged.unsign();
    if load_math_from_package(&staged)?.as_ref() != Some(&value) {
        return Err(invalid("published presentation math did not validate"));
    }
    *package = staged;
    Ok(previous)
}

/// Remove the typed document-level math defaults as one package transaction.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn remove_math_from_package(
    package: &mut OpcPackage,
) -> Result<Option<crate::presentation_properties::math::Properties>> {
    let previous = load_math_from_package(package)?;
    let Some(previous_value) = previous.clone() else {
        return Ok(None);
    };

    let mut staged = package.clone();
    let uri = find_properties_part(&staged)?.ok_or_else(|| {
        invalid("presentation-properties relationship disappeared during staging")
    })?;
    let mut snapshot = load_from_package(&staged)?.ok_or_else(|| {
        invalid("presentation-properties relationship disappeared during staging")
    })?;
    let removed = snapshot.remove_math();
    if removed.as_ref() != Some(&previous_value) {
        return Err(invalid("presentation math snapshot changed during staging"));
    }
    let strict = is_strict_root(staged.get_part(&uri)?.blob())?;
    let blob = snapshot.to_xml(strict)?;
    if Properties::parse(&blob)?.math().is_some() {
        return Err(invalid(
            "staged presentation math removal did not round-trip",
        ));
    }
    staged.get_part_mut(&uri)?.set_blob(blob);
    staged.unsign();
    if load_math_from_package(&staged)?.is_some() {
        return Err(invalid(
            "published presentation math removal did not validate",
        ));
    }
    *package = staged;
    Ok(Some(previous_value))
}

fn find_properties_part(package: &OpcPackage) -> Result<Option<PackURI>> {
    let presentation = package.main_document_part()?;
    let mut found = presentation
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL));
    let Some(relationship) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "presentation has multiple presentation-properties relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid(
            "presentation-properties relationship cannot be external",
        ));
    }
    let uri = relationship.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "presentation-properties part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    Ok(Some(uri))
}

fn hydrate_html_target(part: &dyn litchi_opc::Part, value: &mut Properties) -> Result<()> {
    if let Some(html) = value.html_publish.as_mut() {
        let target = part
            .rels()
            .get(&html.target.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "missing HTML publish relationship '{}'",
                    html.target.relationship_id
                ))
            })?;
        html.target.target = Some(target.target_ref().to_string());
        html.target.relationship_type = Some(target.reltype().to_string());
        html.target.external = Some(target.is_external());
    }
    Ok(())
}

fn is_strict_root(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::Xml(format!("invalid presentation-properties root: {error}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"presentationPr" {
                    return Err(invalid(
                        "presentation-properties root is not presentationPr",
                    ));
                }
                return match namespace {
                    ResolveResult::Bound(Namespace(value))
                        if value == super::P_STRICT.as_bytes() =>
                    {
                        Ok(true)
                    },
                    ResolveResult::Bound(Namespace(value)) if value == super::P_NS.as_bytes() => {
                        Ok(false)
                    },
                    _ => Err(invalid(
                        "unsupported presentation-properties root namespace",
                    )),
                };
            },
            Event::Decl(_) | Event::Comment(_) | Event::Text(_) => {},
            Event::CData(_) | Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid("unsupported presentation-properties prologue"));
            },
            Event::End(_) => return Err(invalid("presentation-properties root is missing")),
            Event::Eof => return Err(invalid("presentation-properties root is missing")),
        }
    }
}
