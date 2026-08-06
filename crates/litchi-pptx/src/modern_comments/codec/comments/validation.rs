use super::super::super::model::{List, NamespaceDeclaration};
use super::super::super::{MAX_BYTES, MAX_COMMENTS, MAX_REPLIES, MAX_STRING_BYTES};
use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::custom_xml::valid_guid;
use std::collections::HashSet;

pub(super) fn validate_model(value: &List) -> Result<()> {
    validate_prefix(&value.root_prefix)?;
    validate_namespaces(&value.namespace_declarations, Some(&value.root_prefix))?;
    if value.comments.len() > MAX_COMMENTS {
        return Err(limit("modern comments"));
    }
    let mut replies = 0usize;
    for comment in &value.comments {
        validate_guid(&comment.id)?;
        validate_guid(&comment.author_id)?;
        validate_date_time(&comment.created)?;
        if let Some(value) = &comment.start_date {
            validate_date_time(value)?;
        }
        if let Some(value) = &comment.due_date {
            validate_date_time(value)?;
        }
        if let Some(ids) = &comment.assigned_to {
            for id in ids {
                validate_guid(id)?;
            }
        }
        validate_optional_string(comment.title.as_deref())?;
        validate_namespaces(&comment.namespace_declarations, None)?;
        validate_namespaces(&comment.reply_list_namespace_declarations, None)?;
        if !comment.reply_list_present && !comment.replies.is_empty() {
            return Err(invalid("modern comment replies require replyLst presence"));
        }
        replies = replies
            .checked_add(comment.replies.len())
            .ok_or_else(|| limit("modern comment replies"))?;
        if replies > MAX_REPLIES {
            return Err(limit("modern comment replies"));
        }
        for anchor in &comment.anchors {
            if anchor.xml.len() > MAX_BYTES {
                return Err(limit("modern comment anchor bytes"));
            }
        }
        validate_fragment(comment.text_body_xml.as_deref())?;
        validate_fragment(comment.extension_xml.as_deref())?;
        for reply in &comment.replies {
            validate_guid(&reply.id)?;
            validate_guid(&reply.author_id)?;
            validate_date_time(&reply.created)?;
            validate_namespaces(&reply.namespace_declarations, None)?;
            validate_fragment(reply.text_body_xml.as_deref())?;
            validate_fragment(reply.extension_xml.as_deref())?;
        }
    }
    Ok(())
}

pub(super) fn validate_namespaces(
    value: &[NamespaceDeclaration],
    excluded: Option<&str>,
) -> Result<()> {
    let mut seen = HashSet::new();
    for declaration in value {
        validate_prefix(&declaration.prefix)?;
        bounded(&declaration.uri)?;
        if declaration.prefix == "xml" || declaration.prefix == "xmlns" {
            return Err(invalid(
                "reserved XML namespace prefix cannot be redeclared",
            ));
        }
        if excluded == Some(declaration.prefix.as_str()) {
            return Err(invalid("modern Comment namespace prefix is declared twice"));
        }
        if !seen.insert(&declaration.prefix) {
            return Err(invalid("duplicate namespace declaration"));
        }
    }
    Ok(())
}

fn validate_prefix(value: &str) -> Result<()> {
    if value.is_empty()
        || (value.bytes().enumerate().all(|(index, byte)| {
            byte.is_ascii_alphanumeric()
                || byte == b'_'
                || byte == b'-'
                || (byte == b'.' && index != 0)
        }) && !value.as_bytes()[0].is_ascii_digit()
            && !value.starts_with('-'))
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML namespace prefix '{value}'")))
    }
}

pub(super) fn validate_guid(value: &str) -> Result<()> {
    if valid_guid(value) {
        Ok(())
    } else {
        Err(invalid(format!("invalid modern Comment GUID '{value}'")))
    }
}

pub(super) fn validate_date_time(value: &str) -> Result<()> {
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML dateTime '{value}'")))
    }
}

fn validate_optional_string(value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        bounded(value)?;
    }
    Ok(())
}

fn validate_fragment(value: Option<&[u8]>) -> Result<()> {
    if value.is_some_and(|value| value.len() > MAX_BYTES) {
        Err(limit("modern Comment opaque fragment bytes"))
    } else {
        Ok(())
    }
}

pub(super) fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("modern Comment string bytes"))
    }
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(label: &str) -> Error {
    invalid(format!("{label} exceeds implementation limit"))
}
