//! Strict PIDSI validation layered over the generic MS-OLEPS validator.

use super::super::codec::validate_section as validate_generic;
use super::super::model::{CodePage, Section, Stream, Value, invalid};
use super::model::*;
use litchi_cfb::OleError;

pub(super) fn validate_section(section: &Section) -> Result<(), OleError> {
    if section.format_identifier != super::super::model::SUMMARY_INFORMATION_FMTID {
        return Err(invalid(
            "SummaryInformation format identifier does not match",
        ));
    }
    if section.named_properties().next().is_some() {
        return Err(invalid(
            "SummaryInformation must not contain a Dictionary property",
        ));
    }
    let page = section
        .page()
        .ok_or_else(|| invalid("SummaryInformation must contain a CodePage property"))?;
    validate_generic(section, Stream::VERSION_0)?;
    validate_codepage(section, page)?;

    for identifier in section.property_ids() {
        let Some(value) = section.property(identifier) else {
            continue;
        };
        match identifier {
            CODEPAGE => {},
            TITLE | SUBJECT | AUTHOR | KEYWORDS | COMMENTS | TEMPLATE | LAST_AUTHOR
            | REVISION_NUMBER | APP_NAME => validate_string(value, identifier, page)?,
            EDIT_TIME | LAST_PRINTED | CREATE_DTM | LAST_SAVE_DTM => {
                if !matches!(value, Value::Filetime(_)) {
                    return Err(invalid(format!(
                        "SummaryInformation property {identifier} must be VT_FILETIME"
                    )));
                }
            },
            PAGE_COUNT | WORD_COUNT | CHARACTER_COUNT => validate_count(value, identifier)?,
            THUMBNAIL => validate_thumbnail(value)?,
            DOC_SECURITY => validate_security(value)?,
            _ => {},
        }
    }
    Ok(())
}

pub(super) fn validate_codepage(section: &Section, page: CodePage) -> Result<(), OleError> {
    for identifier in section.property_ids() {
        match section.property(identifier) {
            Some(Value::Lpstr(value)) => {
                validate_text_for_page(value, page, "Property Set string")?
            },
            Some(Value::Lpwstr(value)) => {
                if page != CodePage::Utf16Le {
                    return Err(invalid(
                        "LPWSTR SummaryInformation strings require CP_WINUNICODE",
                    ));
                }
                validate_text_for_page(value, page, "Property Set string")?;
            },
            _ => {},
        }
    }
    for (name, _) in section.named_properties() {
        validate_text_for_page(name, page, "Property Set dictionary name")?;
    }
    Ok(())
}

pub(super) fn validate_text_for_page(
    value: &str,
    page: CodePage,
    field: &str,
) -> Result<(), OleError> {
    if value.len() > MAX_TEXT_BYTES || value.contains('\0') {
        return Err(invalid(format!("{field} exceeds the typed string limit")));
    }
    let encoded_len = match page {
        CodePage::Utf16Le => value
            .encode_utf16()
            .count()
            .checked_add(1)
            .and_then(|units| units.checked_mul(2))
            .ok_or_else(|| invalid(format!("{field} length overflows")))?,
        CodePage::Mbcs(page) => page
            .encode(value)
            .map_err(|error| {
                invalid(format!(
                    "{field} is not representable in its code page: {error}"
                ))
            })?
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid(format!("{field} length overflows")))?,
    };
    if encoded_len > MAX_TEXT_BYTES {
        return Err(invalid(format!("{field} exceeds the encoded string limit")));
    }
    Ok(())
}

fn validate_string(value: &Value, identifier: u32, page: CodePage) -> Result<(), OleError> {
    let value = match value {
        Value::Lpstr(value) => value,
        Value::Lpwstr(value) if page == CodePage::Utf16Le => value,
        Value::Lpwstr(_) => {
            return Err(invalid(format!(
                "SummaryInformation property {identifier} LPWSTR requires CP_WINUNICODE"
            )));
        },
        _ => {
            return Err(invalid(format!(
                "SummaryInformation property {identifier} must be a VtString"
            )));
        },
    };
    validate_text_for_page(value, page, "SummaryInformation string")?;
    if identifier == REVISION_NUMBER && (value.is_empty() || value.parse::<u64>().is_err()) {
        return Err(invalid(format!(
            "SummaryInformation revision number must be a nonnegative whole number"
        )));
    }
    Ok(())
}

fn validate_count(value: &Value, identifier: u32) -> Result<(), OleError> {
    match value {
        Value::I4(value) if *value >= 0 => Ok(()),
        Value::I4(_) => Err(invalid(format!(
            "SummaryInformation count property {identifier} must be nonnegative"
        ))),
        _ => Err(invalid(format!(
            "SummaryInformation property {identifier} must be VT_I4"
        ))),
    }
}

fn validate_thumbnail(value: &Value) -> Result<(), OleError> {
    let thumbnail = ThumbnailRef::from_value(value)?;
    if thumbnail.data().len() > MAX_THUMBNAIL_BYTES {
        return Err(invalid(
            "SummaryInformation thumbnail exceeds the safety limit",
        ));
    }
    Ok(())
}

fn validate_security(value: &Value) -> Result<(), OleError> {
    if matches!(value, Value::I4(_)) {
        Ok(())
    } else {
        Err(invalid("SummaryInformation DocSecurity must be VT_I4"))
    }
}
