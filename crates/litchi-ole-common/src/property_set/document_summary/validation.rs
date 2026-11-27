//! Strict PIDDSI validation layered over the generic MS-OLEPS validator.

use super::super::codec::validate_section as validate_generic;
use super::super::model::{Section, Stream, Value, invalid};
use super::model::{
    BYTE_COUNT, CATEGORY, CHARACTER_COUNT_WITH_SPACES, COMPANY, CONTENT_STATUS, CONTENT_TYPE,
    DOCUMENT_PARTS, DOCUMENT_VERSION, HEADING_PAIRS, HIDDEN_COUNT, HYPERLINKS_CHANGED, LANGUAGE,
    LINE_COUNT, LINKS_DIRTY, MANAGER, MAX_TEXT_BYTES, MULTIMEDIA_CLIP_COUNT, NOTE_COUNT,
    PARAGRAPH_COUNT, PRESENTATION_FORMAT, SCALE, SHARED_DOCUMENT, SLIDE_COUNT, VERSION, Version,
};
use litchi_cfb::OleError;

pub(super) fn validate_section(section: &Section) -> Result<(), OleError> {
    if section.format_identifier != super::super::model::DOCUMENT_SUMMARY_INFORMATION_FMTID {
        return Err(invalid(
            "Document Summary Information format identifier does not match",
        ));
    }
    if section.named_properties().next().is_some() {
        return Err(invalid(
            "Document Summary Information must not contain a Dictionary property",
        ));
    }
    if section.page().is_none() {
        return Err(invalid(
            "Document Summary Information must contain a CodePage property",
        ));
    }
    validate_generic(section, Stream::VERSION_0)?;

    for identifier in section.property_ids() {
        let Some(value) = section.property(identifier) else {
            continue;
        };
        match identifier {
            CATEGORY | PRESENTATION_FORMAT | MANAGER | COMPANY | CONTENT_TYPE | CONTENT_STATUS
            | LANGUAGE | DOCUMENT_VERSION => validate_string(value, identifier)?,
            BYTE_COUNT
            | LINE_COUNT
            | PARAGRAPH_COUNT
            | SLIDE_COUNT
            | NOTE_COUNT
            | HIDDEN_COUNT
            | MULTIMEDIA_CLIP_COUNT
            | CHARACTER_COUNT_WITH_SPACES => validate_i4(value, identifier)?,
            SCALE | SHARED_DOCUMENT => {
                validate_bool(value, identifier)?;
                if matches!(value, Value::Bool(true)) {
                    return Err(invalid(format!(
                        "PIDDSI property {identifier} must be FALSE"
                    )));
                }
            },
            LINKS_DIRTY | HYPERLINKS_CHANGED => validate_bool(value, identifier)?,
            HEADING_PAIRS => {
                if !matches!(value, Value::HeadingPairs(_)) {
                    return Err(invalid(format!(
                        "PIDDSI property {identifier} must be HeadingPairs"
                    )));
                }
            },
            DOCUMENT_PARTS => {
                if !matches!(value, Value::DocParts(_)) {
                    return Err(invalid(format!(
                        "PIDDSI property {identifier} must be DocParts"
                    )));
                }
            },
            VERSION => {
                let Value::I4(raw) = value else {
                    return Err(invalid("PIDDSI Version must be VT_I4"));
                };
                Version::from_raw(*raw)?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_string(property: &Value, identifier: u32) -> Result<(), OleError> {
    let (Value::Lpstr(text) | Value::Lpwstr(text)) = property else {
        return Err(invalid(format!(
            "PIDDSI property {identifier} must be a VtString"
        )));
    };
    if text.len() > MAX_TEXT_BYTES || text.contains('\0') {
        return Err(invalid(format!(
            "PIDDSI property {identifier} exceeds the typed string limit"
        )));
    }
    Ok(())
}

fn validate_i4(value: &Value, identifier: u32) -> Result<(), OleError> {
    if matches!(value, Value::I4(_)) {
        Ok(())
    } else {
        Err(invalid(format!(
            "PIDDSI property {identifier} must be VT_I4"
        )))
    }
}

fn validate_bool(value: &Value, identifier: u32) -> Result<(), OleError> {
    if matches!(value, Value::Bool(_)) {
        Ok(())
    } else {
        Err(invalid(format!(
            "PIDDSI property {identifier} must be VT_BOOL"
        )))
    }
}
