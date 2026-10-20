//! The `/metadata` set path of the legacy OLE formats (DOC, XLS, PPT).
//!
//! `set("/metadata", props)` stages document properties through the leaf
//! `Package::edit_metadata` transaction of each format. The keys are the
//! OfficeCLI root-metadata spellings that the DOCX/XLSX/PPTX roots read
//! (`title`, `author`, `lastModifiedBy`, `revisionNumber`, ...), plus the
//! `DocumentSummaryInformation` fields and `custom.<name>` for user-defined
//! properties. `null` removes a property. `keywords` takes a list, or text
//! split at `;` and `,`; PIDSI `KEYWORDS` holds the list joined with `, `,
//! and a list equal to the stored one keeps the stored text (record 0354).
//! The registry below is a static sorted table; lookup allocates nothing.
//!
//! Creation (`create_with(.., CreateOptions::metadata(..))`) uses the leaf
//! builders' `set_metadata` instead, which keeps the root class identifier
//! the transaction route also preserves.

use litchi_core::FileFormat;
use litchi_ole_common::property_set::summary_properties::MetadataError;
use litchi_ole_common::property_set::{SummaryProperties, Value};

use crate::Props;
use crate::error::{Code, Error, Result};
use crate::props::PropValue;
use crate::target::TargetArg;

/// The path segment of the metadata node.
pub(crate) const SEGMENT: &str = "metadata";

/// The prefix of user-defined (custom) property keys.
const CUSTOM_PREFIX: &str = "custom.";

/// One registered metadata key: its unified spelling and the name
/// [`SummaryProperties::set`] resolves. Every key takes text (RFC 3339 for
/// the timestamps).
#[derive(Debug, Clone, Copy)]
pub(crate) struct Key {
    /// The key as callers write it.
    pub(crate) name: &'static str,
    /// The property name the shared OLE model resolves.
    target: &'static str,
}

const fn key(name: &'static str, target: &'static str) -> Key {
    Key { name, target }
}

/// Every key, sorted by ASCII-lowercase name for binary search.
pub(crate) static KEYS: [Key; 22] = [
    key("application", "applicationname"),
    key("author", "author"),
    key("category", "category"),
    key("company", "company"),
    key("contentStatus", "contentstatus"),
    key("contentType", "contenttype"),
    key("created", "created"),
    key("description", "comments"),
    key("extended.company", "company"),
    key("extended.manager", "manager"),
    key("extended.template", "template"),
    key("keywords", "keywords"),
    key("language", "language"),
    key("lastModifiedBy", "lastauthor"),
    key("lastPrinted", "lastprinted"),
    key("manager", "manager"),
    key("modified", "modified"),
    key("revisionNumber", "revisionnumber"),
    key("subject", "subject"),
    key("template", "template"),
    key("title", "title"),
    key("version", "documentversion"),
];

const _: () = {
    let mut index = 1;
    while index < KEYS.len() {
        assert!(
            ascii_lower_cmp(KEYS[index - 1].name, KEYS[index].name).is_lt(),
            "metadata keys must be sorted"
        );
        index += 1;
    }
};

const fn ascii_lower_cmp(left: &str, right: &str) -> std::cmp::Ordering {
    let (left, right) = (left.as_bytes(), right.as_bytes());
    let mut index = 0;
    while index < left.len() && index < right.len() {
        let (a, b) = (
            left[index].to_ascii_lowercase(),
            right[index].to_ascii_lowercase(),
        );
        if a < b {
            return std::cmp::Ordering::Less;
        }
        if a > b {
            return std::cmp::Ordering::Greater;
        }
        index += 1;
    }
    if left.len() < right.len() {
        std::cmp::Ordering::Less
    } else if left.len() > right.len() {
        std::cmp::Ordering::Greater
    } else {
        std::cmp::Ordering::Equal
    }
}

fn lookup(name: &str) -> Option<&'static Key> {
    KEYS.binary_search_by(|key| ascii_lower_cmp(key.name, name))
        .ok()
        .and_then(|index| KEYS.get(index))
}

/// Whether `target` is exactly the `/metadata` node: one absolute step named
/// `metadata`, without predicates.
pub(crate) fn is_target(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        let path = path.path();
        path.is_absolute()
            && matches!(path.steps(), [step]
                if step.predicates().is_empty() && step.name().eq_ignore_ascii_case(SEGMENT))
    })
}

/// Apply every property of `props` to the staged OLE properties, in order.
///
/// # Errors
///
/// `unsupported_property` for an unknown key, `invalid_value` for a value of
/// the wrong kind or one the property refuses; the staged properties may hold
/// earlier keys of the same call, which the caller discards with the edit.
pub(crate) fn apply(
    properties: &mut SummaryProperties,
    props: &Props<'_>,
    format: FileFormat,
) -> Result<usize> {
    if props.is_empty() {
        return Err(Error::new(
            Code::InvalidArgument,
            "a metadata set needs at least one property",
        )
        .with_format(format));
    }
    for (key, value) in props.iter() {
        let name = key.as_str();
        if let Some(custom) = strip_custom(name) {
            apply_custom(properties, name, custom, value, format)?;
            continue;
        }
        let Some(entry) = lookup(name) else {
            return Err(Error::new(
                Code::UnsupportedProperty,
                "the metadata key is not supported by OLE property sets",
            )
            .with_property(name.to_owned())
            .with_format(format)
            .with_suggestion("use a root-metadata key such as `title`, or `custom.<name>`"));
        };
        if entry.name == "keywords" {
            apply_keywords(properties, name, value, format)?;
            continue;
        }
        let result = if value.is_null() {
            properties.remove(entry.target).map(|_| ())
        } else {
            let text = scalar_text(value).ok_or_else(|| wrong_kind(name, format))?;
            properties.set(entry.target, &text)
        };
        result.map_err(|_| invalid(name, format))?;
    }
    Ok(props.len())
}

/// Stage the keyword list: removed when empty, kept when it equals the
/// stored text's split, else the list joined with `, `.
fn apply_keywords(
    properties: &mut SummaryProperties,
    name: &str,
    value: &PropValue<'_>,
    format: FileFormat,
) -> Result<()> {
    use litchi_core::metadata::keywords;
    let list = crate::props::keyword_list(value).ok_or_else(|| wrong_kind(name, format))?;
    let current = properties
        .summary_information()
        .keywords()
        .map(keywords::split)
        .unwrap_or_default();
    if list == current {
        return Ok(());
    }
    let result = if list.is_empty() {
        properties.remove("keywords").map(|_| ())
    } else {
        let text = keywords::join(&list).map_err(|_| {
            Error::new(
                Code::InvalidValue,
                "a keyword holds ';' or ',' or surrounding whitespace, which the OLE keyword string cannot keep",
            )
            .with_property(name.to_owned())
            .with_format(format)
        })?;
        properties.set("keywords", &text)
    };
    result.map_err(|_| invalid(name, format))
}

fn strip_custom(name: &str) -> Option<&str> {
    let prefix = name.get(..CUSTOM_PREFIX.len())?;
    prefix
        .eq_ignore_ascii_case(CUSTOM_PREFIX)
        .then(|| &name[CUSTOM_PREFIX.len()..])
        .filter(|custom| !custom.is_empty())
}

fn apply_custom(
    properties: &mut SummaryProperties,
    key: &str,
    name: &str,
    value: &PropValue<'_>,
    format: FileFormat,
) -> Result<()> {
    if name.starts_with('_') {
        return Err(invalid(key, format));
    }
    let typed = match value {
        PropValue::Null => {
            return properties
                .edit_user_defined_properties(|section| {
                    section.remove_named(name);
                    Ok(())
                })
                .map(|_| ())
                .map_err(|_| invalid(key, format));
        },
        PropValue::Bool(value) => Value::Bool(*value),
        PropValue::Int(value) => match i32::try_from(*value) {
            Ok(value) => Value::I4(value),
            Err(_) => Value::R8(*value as f64),
        },
        PropValue::Number(value) if value.is_finite() => Value::R8(*value),
        PropValue::Text(text) => Value::Lpstr(text.to_string()),
        _ => return Err(wrong_kind(key, format)),
    };
    properties
        .set_user_defined(name, typed)
        .map_err(|_| invalid(key, format))
}

/// Text for a text key: text as is, integers in decimal (a revision number
/// may be written as `7`).
fn scalar_text(value: &PropValue<'_>) -> Option<String> {
    match value {
        PropValue::Text(text) => Some(text.to_string()),
        PropValue::Int(_) | PropValue::Unsigned(_) => Some(value.to_string()),
        _ => None,
    }
}

fn wrong_kind(key: &str, format: FileFormat) -> Error {
    Error::new(Code::InvalidValue, "the metadata property requires text")
        .with_property(key.to_owned())
        .with_format(format)
}

fn invalid(key: &str, format: FileFormat) -> Error {
    Error::new(
        Code::InvalidValue,
        "the metadata value is not representable in OLE property sets",
    )
    .with_property(key.to_owned())
    .with_format(format)
}

/// Map a unified-metadata conversion refusal (create-time).
pub(crate) fn metadata_error(error: &MetadataError, format: FileFormat) -> Error {
    match error {
        MetadataError::Unsupported { field } => Error::new(
            Code::UnsupportedCapability,
            "the metadata field cannot be stored in OLE property sets",
        )
        .with_property(*field)
        .with_format(format)
        .with_suggestion("leave the field unset for legacy binary formats"),
        _ => Error::new(
            Code::InvalidValue,
            "the metadata value is not representable in OLE property sets",
        )
        .with_property(error.field())
        .with_format(format),
    }
}

/// The patch a legacy write step retains: the format's primary patch, or a
/// metadata transaction's exact-source patch. Both expose the endpoint bytes
/// the generic history compares. (XLS keeps its own, with worksheet-feature
/// patches; record 0327.)
#[cfg(any(feature = "doc", feature = "ppt"))]
pub(crate) enum StepPatch<P, M> {
    /// A primary content edit.
    Primary(P),
    /// A metadata edit.
    Metadata(M),
}

/// Exact package endpoints of a retained leaf patch.
#[cfg(any(feature = "doc", feature = "ppt"))]
pub(crate) trait Endpoints {
    fn before_image(&self) -> &[u8];
    fn after_image(&self) -> &[u8];
}

#[cfg(any(feature = "doc", feature = "ppt"))]
impl<P: Endpoints, M: Endpoints> StepPatch<P, M> {
    pub(crate) fn before(&self) -> &[u8] {
        match self {
            Self::Primary(patch) => patch.before_image(),
            Self::Metadata(patch) => patch.before_image(),
        }
    }

    pub(crate) fn after(&self) -> &[u8] {
        match self {
            Self::Primary(patch) => patch.after_image(),
            Self::Metadata(patch) => patch.after_image(),
        }
    }
}

#[cfg(any(feature = "doc", feature = "ppt"))]
macro_rules! endpoints {
    ($type:ty, $before:ident, $after:ident) => {
        impl Endpoints for $type {
            fn before_image(&self) -> &[u8] {
                self.$before()
            }

            fn after_image(&self) -> &[u8] {
                self.$after()
            }
        }
    };
}

#[cfg(feature = "doc")]
endpoints!(litchi_doc::Patch, before_bytes, after_bytes);
#[cfg(feature = "doc")]
endpoints!(
    litchi_doc::package::metadata::Patch,
    before_bytes,
    after_bytes
);
#[cfg(feature = "ppt")]
endpoints!(litchi_ppt::Patch, before, after);
#[cfg(feature = "ppt")]
endpoints!(
    litchi_ppt::package::metadata::Patch,
    before_bytes,
    after_bytes
);

/// Retained bytes of a metadata patch: both endpoint images.
pub(crate) fn snapshot_bytes(before: &[u8], after: &[u8]) -> usize {
    before.len().saturating_add(after.len())
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "test assertions panic by design")]

    use super::*;

    #[test]
    fn keys_resolve_case_insensitively_and_custom_keys_are_separate() {
        assert_eq!(lookup("Title").unwrap().target, "title");
        assert_eq!(lookup("LASTMODIFIEDBY").unwrap().target, "lastauthor");
        assert!(lookup("custom.Department").is_none());
        assert_eq!(strip_custom("Custom.Department"), Some("Department"));
        assert_eq!(strip_custom("custom."), None);
        assert!(lookup("identifier").is_none());
    }

    #[test]
    fn apply_sets_removes_and_refuses() {
        let mut properties = SummaryProperties::new();
        let props = Props::new()
            .with("title", "Report")
            .with("revisionNumber", 7)
            .with("custom.Reviewed", true)
            .with("created", "2026-10-05T09:30:00Z");
        assert_eq!(apply(&mut properties, &props, FileFormat::Doc).unwrap(), 4);
        let summary = properties.summary_information();
        assert_eq!(summary.title(), Some("Report"));
        assert_eq!(summary.revision_number(), Some("7"));

        let clear = Props::new()
            .with("title", PropValue::Null)
            .with("custom.Reviewed", PropValue::Null);
        apply(&mut properties, &clear, FileFormat::Doc).unwrap();
        assert!(properties.summary_information().title().is_none());
        assert!(properties.user_defined_properties().is_none());

        for (props, code) in [
            (
                Props::new().with("identifier", "x"),
                Code::UnsupportedProperty,
            ),
            (
                Props::new().with("created", "yesterday"),
                Code::InvalidValue,
            ),
            (Props::new().with("title", true), Code::InvalidValue),
            (
                Props::new().with("custom._PID_LINKBASE", "x"),
                Code::InvalidValue,
            ),
            (Props::new(), Code::InvalidArgument),
        ] {
            let error = apply(&mut properties, &props, FileFormat::Doc).unwrap_err();
            assert_eq!(error.code(), code);
        }
    }
}
