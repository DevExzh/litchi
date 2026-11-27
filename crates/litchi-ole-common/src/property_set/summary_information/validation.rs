//! Strict PIDSI validation layered over the generic MS-OLEPS validator.

use super::super::codec::validate_section as validate_generic;
use super::super::model::{CodePage, Section, Stream, Value, invalid};
use super::model::{
    APP_NAME, AUTHOR, CHARACTER_COUNT, COMMENTS, CREATE_DTM, DOC_SECURITY, EDIT_TIME, KEYWORDS,
    LAST_AUTHOR, LAST_PRINTED, LAST_SAVE_DTM, MAX_TEXT_BYTES, MAX_THUMBNAIL_BYTES, PAGE_COUNT,
    REVISION_NUMBER, SUBJECT, TEMPLATE, THUMBNAIL, TITLE, ThumbnailRef, WORD_COUNT,
};
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
        if let Some(Value::Lpstr(text)) = section.property(identifier) {
            validate_text_for_page(text, page, "Property Set string")?;
        } else if let Some(Value::Lpwstr(text)) = section.property(identifier) {
            if page != CodePage::Utf16Le {
                return Err(invalid(
                    "LPWSTR SummaryInformation strings require CP_WINUNICODE",
                ));
            }
            validate_text_for_page(text, page, "Property Set string")?;
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
        CodePage::Mbcs(mbcs) => mbcs
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

fn validate_string(property: &Value, identifier: u32, page: CodePage) -> Result<(), OleError> {
    let text = if let Value::Lpstr(text) = property {
        text
    } else if let Value::Lpwstr(text) = property {
        if page != CodePage::Utf16Le {
            return Err(invalid(format!(
                "SummaryInformation property {identifier} LPWSTR requires CP_WINUNICODE"
            )));
        }
        text
    } else {
        return Err(invalid(format!(
            "SummaryInformation property {identifier} must be a VtString"
        )));
    };
    validate_text_for_page(text, page, "SummaryInformation string")?;
    if identifier == REVISION_NUMBER && (text.is_empty() || text.parse::<u64>().is_err()) {
        return Err(invalid(
            "SummaryInformation revision number must be a nonnegative whole number".to_string(),
        ));
    }
    Ok(())
}

fn validate_count(property: &Value, identifier: u32) -> Result<(), OleError> {
    if let Value::I4(count) = property {
        if *count >= 0 {
            Ok(())
        } else {
            Err(invalid(format!(
                "SummaryInformation count property {identifier} must be nonnegative"
            )))
        }
    } else {
        Err(invalid(format!(
            "SummaryInformation property {identifier} must be VT_I4"
        )))
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
