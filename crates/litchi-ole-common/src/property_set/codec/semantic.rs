//! Semantic validation shared by parsing, serialization, and editors.

use super::super::model::*;
use litchi_cfb::OleError;

pub(super) const MAX_PROPERTY_COUNT: usize = 16_384;

fn minimum_property_set_version(value: &Value) -> u16 {
    match value {
        Value::I1(_) | Value::Int(_) | Value::UInt(_) => 1,
        Value::Vector(vector) => vector
            .values()
            .iter()
            .map(minimum_property_set_version)
            .max()
            .unwrap_or(Stream::VERSION_0),
        Value::Array(array) => array
            .values()
            .iter()
            .map(minimum_property_set_version)
            .max()
            .unwrap_or(Stream::VERSION_0),
        _ => Stream::VERSION_0,
    }
}

pub(super) fn validate_section(section: &Section, version: u16) -> Result<(), OleError> {
    if !matches!(version, Stream::VERSION_0 | Stream::VERSION_1) {
        return Err(invalid(format!(
            "Unsupported Property Set version {version}"
        )));
    }
    if section.properties.len() > MAX_PROPERTY_COUNT {
        return Err(invalid("Property count exceeds safety limit"));
    }
    let effective_codepage = section.codepage.map_or(DEFAULT_CODEPAGE, |page| page.id());
    for identifier in section.properties.keys() {
        if !valid_property_identifier(*identifier) {
            return Err(invalid(format!(
                "Property identifier {identifier} is outside the Property Set range"
            )));
        }
    }
    for (identifier, value) in &section.properties {
        let required_version = minimum_property_set_version(value);
        if required_version > version {
            return Err(invalid(format!(
                "Property {identifier} requires Property Set version {required_version}"
            )));
        }
        if let Value::VersionedStream(value) = value {
            value.validate_for_property(*identifier)?;
        }
        match value {
            Value::HeadingPairs(value) => value.validate()?,
            Value::DocParts(value) => value.validate_for_codepage(effective_codepage)?,
            _ => {},
        }
    }
    if let (Some(Value::HeadingPairs(headings)), Some(Value::DocParts(parts))) = (
        section.properties.get(&PID_HEADING_PAIRS),
        section.properties.get(&PID_DOC_PARTS),
    ) {
        let expected = headings.document_part_count();
        let actual = parts.len() as u64;
        if expected != actual {
            return Err(invalid(format!(
                "Heading pair part count {expected} does not match document-part count {actual}"
            )));
        }
    }
    if let Some(value) = section.properties.get(&PID_BEHAVIOR) {
        if version == Stream::VERSION_0 {
            return Err(invalid("Behavior property requires Property Set version 1"));
        }
        if !matches!(value, Value::UI4(0 | 1)) {
            return Err(invalid(
                "Behavior property must be VT_UI4 with value 0 or 1",
            ));
        }
    }
    let mut names =
        try_hash_set_with_capacity(section.dictionary.len(), "property dictionary names")?;
    for (identifier, name) in &section.dictionary {
        if !valid_named_property_identifier(*identifier) {
            return Err(invalid(format!(
                "Dictionary property identifier {identifier} is outside the normal range"
            )));
        }
        validate_property_name(name)?;
        if !names.insert(AsciiInsensitive(name)) {
            return Err(invalid("Duplicate dictionary property name"));
        }
        if !section.properties.contains_key(identifier) {
            return Err(invalid("Dictionary references a missing property"));
        }
    }
    match (section.codepage, section.properties.get(&PID_CODEPAGE)) {
        (Some(codepage), Some(Value::I2(value))) if *value == codepage.id() as i16 => {},
        (None, None) => {},
        _ => return Err(invalid("PID 1 does not match section codepage")),
    }
    Ok(())
}
