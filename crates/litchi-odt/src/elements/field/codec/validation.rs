//! Typed ODF field and wire-boundary validation.

use super::super::model::*;
use super::super::{
    MAX_DATABASE_AGGREGATE, MAX_DATABASE_VALUE, MAX_FIELD_DEPTH, TEXT_DATABASE_NAMESPACE,
};
use litchi_core::{Error, Result};
use std::collections::HashMap;

pub(in crate::elements::field) fn checked_field_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("field nesting depth overflow".to_string()))?;
    if depth > MAX_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "field nesting exceeds {MAX_FIELD_DEPTH} levels"
        )));
    }
    Ok(depth)
}

pub(super) type DatabaseAttributes = HashMap<(String, String), String>;

pub(super) fn validate_note_body_attributes(
    source: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    for attribute in source.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid text:note-body attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        return Err(Error::InvalidFormat(
            "text:note-body does not permit attributes".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_meta_field_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    let valid = parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    });
    if valid {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "text:meta-field occurs outside an ODF inline-text host".to_string(),
        ))
    }
}

pub(super) fn validate_drop_down_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    if parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "text:drop-down occurs outside an ODF inline-text host".to_string(),
        ))
    }
}

pub(super) fn validate_database_field(field: DatabaseField) -> Result<DatabaseField> {
    match field.kind {
        DatabaseFieldKind::Display if field.column_name.is_none() => {
            return Err(Error::InvalidFormat(
                "text:database-display requires text:column-name".to_string(),
            ));
        },
        DatabaseFieldKind::Next | DatabaseFieldKind::RowSelect
            if !field.display_text.is_empty() =>
        {
            return Err(Error::InvalidFormat(
                "database selection fields cannot contain character data".to_string(),
            ));
        },
        _ => {},
    }
    if field.number_letter_sync.is_some()
        && !matches!(field.number_format.as_deref(), Some("a" | "A"))
    {
        return Err(Error::InvalidFormat(
            "style:num-letter-sync requires style:num-format a or A".to_string(),
        ));
    }
    Ok(field)
}

pub(super) fn validate_constructed_database_field(field: &DatabaseField) -> Result<()> {
    let mut aggregate = 0usize;
    for value in [
        field.source.database_name.as_deref(),
        Some(field.source.table_name.as_str()),
        field.column_name.as_deref(),
        field.condition.as_deref(),
        field.data_style_name.as_deref(),
        field.number_format.as_deref(),
        Some(field.display_text.as_str()),
        field
            .source
            .connection_resource
            .as_ref()
            .map(|resource| resource.href.as_str()),
    ]
    .into_iter()
    .flatten()
    {
        if !value.chars().all(is_xml_1_0_char) {
            return Err(Error::InvalidFormat(
                "database field contains forbidden XML characters".to_string(),
            ));
        }
        append_database_size(&mut aggregate, value.len())?;
    }
    if field
        .source
        .connection_resource
        .as_ref()
        .is_some_and(|resource| !resource.simple_link)
    {
        return Err(Error::InvalidFormat(
            "ODF form:connection-resource only supports xlink:href".to_string(),
        ));
    }
    let forbidden = match field.kind {
        DatabaseFieldKind::Display => {
            field.condition.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        DatabaseFieldKind::Next => {
            field.column_name.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        DatabaseFieldKind::RowSelect => {
            field.column_name.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
        DatabaseFieldKind::RowNumber => {
            field.column_name.is_some()
                || field.condition.is_some()
                || field.row_number.is_some()
                || field.data_style_name.is_some()
        },
        DatabaseFieldKind::Name => {
            field.column_name.is_some()
                || field.condition.is_some()
                || field.row_number.is_some()
                || field.value.is_some()
                || field.data_style_name.is_some()
                || field.number_format.is_some()
                || field.number_letter_sync.is_some()
        },
    };
    if forbidden {
        return Err(Error::InvalidFormat(
            "database field contains attributes from another field kind".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn reject_drop_down_attributes(
    attributes: &DatabaseAttributes,
    allowed: &[(&str, &str)],
) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) {
            return Err(Error::InvalidFormat(format!(
                "unexpected drop-down field attribute {namespace}:{local}"
            )));
        }
    }
    Ok(())
}

pub(super) fn reject_database_attributes(
    attributes: &DatabaseAttributes,
    allowed: &[(&str, &str)],
) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) {
            return Err(Error::InvalidFormat(format!(
                "unexpected database field attribute {namespace}:{local}"
            )));
        }
    }
    Ok(())
}

pub(super) fn append_database_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    if amount > MAX_DATABASE_VALUE {
        return Err(Error::InvalidFormat(
            "database field value exceeds 64 KiB".to_string(),
        ));
    }
    *aggregate = aggregate.checked_add(amount).ok_or_else(|| {
        Error::InvalidFormat("database field aggregate size overflow".to_string())
    })?;
    if *aggregate > MAX_DATABASE_AGGREGATE {
        return Err(Error::InvalidFormat(
            "database field metadata exceeds 16 MiB".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_database_parent(parent: Option<&(Option<String>, String)>) -> Result<()> {
    if parent.is_some_and(|(namespace, local)| {
        namespace.as_deref() == Some(TEXT_DATABASE_NAMESPACE)
            && matches!(
                local.as_str(),
                "a" | "h" | "meta" | "meta-field" | "p" | "ruby-base" | "span"
            )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "database field occurs outside an ODF inline-text host".to_string(),
        ))
    }
}
