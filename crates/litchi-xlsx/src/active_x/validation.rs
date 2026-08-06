//! Structural and resource validation for the inert ActiveX graph.
//!
//! Validation is intentionally independent from XML and OPC mutation. It
//! checks semantic invariants before a codec allocates output or a package
//! transaction changes state; binary payloads are only size-accounted.

use super::model::*;
use super::{
    MAX_BINARY, MAX_CONTROL_NAME_CHARS, MAX_CONTROLS, MAX_DEPTH, MAX_PROPERTIES, MAX_SHAPE_ID,
    MAX_STRING, MAX_TOTAL_BINARY, Result, invalid, limit, relerr,
};
use std::collections::HashSet;

pub(super) fn validate_controls(value: &Controls) -> Result<()> {
    if value.controls.is_empty() || value.controls.len() > MAX_CONTROLS {
        return Err(invalid("controls requires 1..65535 controls"));
    }
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for control in &value.controls {
        if !(1..=MAX_SHAPE_ID).contains(&control.shape_id) || !ids.insert(control.shape_id) {
            return Err(invalid(
                "control shapeId must be unique and within Office's supported range",
            ));
        }
        nonempty(&control.relationship_id, "control relationship ID")?;
        bounded(&control.relationship_id, "control relationship ID")?;
        if let Some(name) = control.name.as_ref() {
            bounded(name, "control name")?;
            if name.chars().count() > MAX_CONTROL_NAME_CHARS {
                return Err(invalid("control name exceeds Office's 32-character limit"));
            }
            if !names.insert(name) {
                return Err(invalid("duplicate control name"));
            }
        }
        if let Some(properties) = control.properties.as_ref() {
            validate_control_properties(properties)?;
        }
    }
    Ok(())
}

fn validate_control_properties(value: &ControlProperties) -> Result<()> {
    if let Some(name) = value.macro_name.as_ref() {
        bounded(name, "control macro name")?;
    }
    if let Some(text) = value.alternate_text.as_ref() {
        bounded(text, "control alternate text")?;
    }
    if let Some(id) = value.preview_relationship_id.as_ref() {
        nonempty(id, "preview relationship ID")?;
        bounded(id, "preview relationship ID")?;
    }
    Ok(())
}

pub(super) fn validate_descriptor(value: &Descriptor) -> Result<()> {
    nonempty(&value.class_id, "ActiveX class ID")?;
    bounded(&value.class_id, "ActiveX class ID")?;
    if let Some(license) = value.license.as_ref() {
        bounded(license, "ActiveX license")?;
    }
    match value.persistence {
        Persistence::PropertyBag => {
            if value.properties.is_empty() || value.relationship_id.is_some() {
                return Err(invalid(
                    "property-bag ActiveX requires properties and forbids r:id",
                ));
            }
        },
        _ => {
            if !value.properties.is_empty()
                || value.relationship_id.as_deref().is_none_or(str::is_empty)
            {
                return Err(invalid(
                    "binary ActiveX persistence requires r:id and forbids properties",
                ));
            }
        },
    }
    let mut count = 0usize;
    validate_properties(&value.properties, 0, &mut count)
}

pub(super) fn validate_font(font: &Font) -> Result<()> {
    match font.persistence {
        Some(Persistence::PropertyBag) => {
            if font.properties.is_empty() || font.relationship_id.is_some() {
                return Err(invalid(
                    "property-bag font requires properties and forbids r:id",
                ));
            }
        },
        Some(_) => {
            if !font.properties.is_empty()
                || font.relationship_id.as_deref().is_none_or(str::is_empty)
            {
                return Err(invalid(
                    "binary font persistence requires r:id and forbids properties",
                ));
            }
        },
        None => {
            if font.relationship_id.is_some() {
                return Err(invalid("font r:id requires a binary persistence mode"));
            }
        },
    }
    Ok(())
}

fn validate_properties(values: &[Property], depth: usize, count: &mut usize) -> Result<()> {
    if depth >= MAX_DEPTH {
        return Err(limit("ActiveX property nesting"));
    }
    let mut names = HashSet::new();
    for property in values {
        *count = count
            .checked_add(1)
            .ok_or_else(|| limit("ActiveX properties"))?;
        if *count > MAX_PROPERTIES {
            return Err(limit("ActiveX properties"));
        }
        nonempty(&property.name, "ActiveX property name")?;
        bounded(&property.name, "ActiveX property name")?;
        if !names.insert(&property.name) {
            return Err(invalid("duplicate ActiveX property name"));
        }
        if let Some(value) = property.value.as_ref() {
            bounded(value, "ActiveX property value")?;
        }
        if property.value.is_some() && property.object.is_some() {
            return Err(invalid("ActiveX property value cannot coexist with object"));
        }
        if let Some(PropertyObject::Font(font)) = property.object.as_ref() {
            validate_font(font)?;
            validate_properties(&font.properties, depth + 1, count)?;
        }
    }
    Ok(())
}

pub(super) fn validate_control_set(value: &ControlSet) -> Result<()> {
    if value.controls.is_empty() || value.controls.len() > MAX_CONTROLS {
        return Err(invalid("ActiveX control set requires 1..65535 controls"));
    }
    validate_controls(&Controls {
        controls: value
            .controls
            .iter()
            .map(|item| item.control.clone())
            .collect(),
    })?;
    let mut total = 0usize;
    for item in &value.controls {
        validate_descriptor(&item.descriptor)?;
        for binary in &item.binaries {
            if binary.bytes.len() > MAX_BINARY {
                return Err(limit("ActiveX binary bytes"));
            }
            total = total
                .checked_add(binary.bytes.len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
        }
        if let Some(preview) = item.preview.as_ref() {
            if preview.bytes.len() > MAX_BINARY {
                return Err(limit("ActiveX preview image bytes"));
            }
            total = total
                .checked_add(preview.bytes.len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
        }
    }
    if total > MAX_TOTAL_BINARY {
        return Err(limit("aggregate ActiveX resource bytes"));
    }
    Ok(())
}

pub(super) fn validate_part_location(
    uri: &litchi_opc::PackURI,
    prefix: &str,
    kind: &str,
) -> Result<()> {
    let Some(filename) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{kind} must be stored below {prefix}")));
    };
    if filename.is_empty() || filename.contains('/') {
        return Err(invalid(format!(
            "{kind} must be a direct child of {prefix}"
        )));
    }
    Ok(())
}

pub(super) fn validate_rel_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    if !matches!(bytes.next(), Some(b'A'..=b'Z' | b'a'..=b'z' | b'_'))
        || !bytes.all(
            |byte| matches!(byte, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'.' | b'-'),
        )
    {
        return Err(relerr("relationship ID must be an XML NCName"));
    }
    bounded(value, "relationship ID")
}

pub(super) fn bounded(value: &str, what: &str) -> Result<()> {
    if value.len() > MAX_STRING {
        Err(limit(what))
    } else {
        Ok(())
    }
}

pub(super) fn nonempty(value: &str, what: &str) -> Result<()> {
    if value.is_empty() {
        Err(invalid(format!("{what} must not be empty")))
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn control_set() -> ControlSet {
        ControlSet {
            controls: vec![LoadedControl {
                control: Control {
                    shape_id: 1,
                    relationship_id: "rId1".into(),
                    name: Some("Button".into()),
                    properties: None,
                },
                descriptor_uri: litchi_opc::PackURI::new("/xl/activeX/activeX1.xml").unwrap(),
                descriptor: Descriptor {
                    class_id: "{inert}".into(),
                    license: None,
                    persistence: Persistence::StreamInit,
                    relationship_id: Some("rIdBinary".into()),
                    properties: Vec::new(),
                },
                binaries: vec![Binary {
                    relationship_id: "rIdBinary".into(),
                    part_uri: litchi_opc::PackURI::new("/xl/activeX/activeX1.bin").unwrap(),
                    bytes: vec![0, 1, 2],
                }],
                preview: None,
            }],
        }
    }

    #[test]
    fn validates_complete_inert_graph_without_interpreting_payload() {
        assert!(validate_control_set(&control_set()).is_ok());
    }

    #[test]
    fn rejects_duplicate_control_identity() {
        let mut value = control_set();
        value.controls.push(value.controls[0].clone());
        assert!(validate_control_set(&value).is_err());
    }

    #[test]
    fn rejects_invalid_relationship_name_before_package_mutation() {
        assert!(validate_rel_id("r:id").is_err());
        assert!(validate_rel_id("rId_1").is_ok());
    }
}
