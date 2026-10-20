//! The `/metadata` write path of the ODF adapters over the shared ODF
//! metadata family (ADR 0034 §7, record 0288).
//!
//! `set("/metadata", props)` stages each key on the leaf's
//! `Package::edit_metadata()` (`litchi_odf_common::metadata::Edit`) and
//! commits once: the leaf splices only the touched `office:meta` children
//! (or adds a `meta.xml` to a package without one) and never stamps a date.
//! `title`, `subject`, `description`, `author` (`meta:initial-creator`),
//! `lastModifiedBy` (`dc:creator`) and `language` take text; `keywords`
//! takes a list (one `meta:keyword` per keyword) or text split at `;` and
//! `,` (record 0354); `null` removes the element. The other root-metadata keys of the OLE and
//! OOXML formats are a typed `unsupported_capability`; an unknown key is
//! `unsupported_property`.

use litchi_core::FileFormat;
#[cfg(odf_objects)]
use litchi_core::Metadata;
use litchi_odf_common::metadata::{Edit, Field, Owner, Patch};

use crate::error::{Code, Error, Result};
use crate::props::{PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Whether `target` is exactly the `/metadata` path.
pub(crate) fn is_target(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        let path = path.path();
        path.is_absolute()
            && matches!(path.steps(), [step]
                if step.predicates().is_empty() && step.name().eq_ignore_ascii_case("metadata"))
    })
}

/// The `/metadata` keys and the fields they write (`creator` is the ODP
/// adapter's alias of `author`).
pub(crate) const KEYS: [(&str, Field); 8] = [
    ("title", Field::Title),
    ("subject", Field::Subject),
    ("description", Field::Description),
    ("author", Field::Author),
    ("creator", Field::Author),
    ("lastModifiedBy", Field::LastModifiedBy),
    ("language", Field::Language),
    ("keywords", Field::Keywords),
];

/// Root-metadata keys other formats write that `meta.xml` holds only as
/// producer-maintained statistics or that ODF has no element for.
const UNOWNED: [&str; 16] = [
    "application",
    "category",
    "company",
    "contentStatus",
    "contentType",
    "created",
    "extended.company",
    "extended.manager",
    "extended.template",
    "lastPrinted",
    "manager",
    "modified",
    "revision",
    "revisionNumber",
    "template",
    "version",
];

/// The field a `/metadata` key writes.
///
/// # Errors
///
/// `unsupported_capability` for a root-metadata key `meta.xml` does not
/// carry as an editable element, `unsupported_property` otherwise.
pub(crate) fn field(key: &str, format: FileFormat) -> Result<Field> {
    if let Some((_, field)) = KEYS.iter().find(|(name, _)| name.eq_ignore_ascii_case(key)) {
        return Ok(*field);
    }
    let unowned = UNOWNED.iter().any(|name| name.eq_ignore_ascii_case(key))
        || key
            .get(..7)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("custom."));
    Err(if unowned {
        Error::new(
            Code::UnsupportedCapability,
            "the ODF metadata family writes title, subject, description, author, lastModifiedBy, language and keywords",
        )
        .with_suggestion("follow-up: ODF user-defined and statistics metadata (record 0288)")
    } else {
        Error::new(
            Code::UnsupportedProperty,
            "the metadata key is not an ODF document property",
        )
        .with_suggestion(
            "use title, subject, description, author, lastModifiedBy, language or keywords",
        )
    }
    .with_property(key.to_owned())
    .with_format(format))
}

/// Stage every `/metadata` assignment of `props` on `edit`.
///
/// # Errors
///
/// As [`field`]; `invalid_value` for a value that is not text or null, or
/// text XML cannot carry.
pub(crate) fn stage<P: Owner>(
    edit: &mut Edit<P>,
    props: &Props<'_>,
    format: FileFormat,
    limits: &SemanticLimits,
) -> Result<()> {
    if props.is_empty() {
        return Err(Error::new(
            Code::InvalidArgument,
            "a metadata set needs at least one property",
        )
        .with_format(format));
    }
    if props.len() > limits.max_operations() {
        return Err(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: props.len(),
            maximum: limits.max_operations(),
        }
        .into());
    }
    for (key, value) in props.iter() {
        let name = key.as_str();
        let field = field(name, format)?;
        if field == Field::Keywords {
            let keywords = crate::props::keyword_list(value).ok_or_else(|| {
                Error::new(
                    Code::InvalidValue,
                    "the keywords take a list, or text separated by ';' or ','",
                )
                .with_format(format)
                .with_property(name.to_owned())
            })?;
            edit.set_keywords(&keywords)
                .map_err(|error| staging(error, format).with_property(name.to_owned()))?;
            continue;
        }
        let value = match value {
            PropValue::Null => None,
            other => Some(other.as_text().ok_or_else(|| {
                Error::new(Code::InvalidValue, "a metadata value must be text or null")
                    .with_format(format)
                    .with_property(name.to_owned())
            })?),
        };
        edit.set(field, value)
            .map_err(|error| staging(error, format).with_property(name.to_owned()))?;
    }
    Ok(())
}

/// Stage create-time `metadata` on `edit`.
#[cfg(odf_objects)]
///
/// # Errors
///
/// `unsupported_capability` when `metadata` sets a field the family does
/// not write; `invalid_value` for text XML cannot carry.
pub(crate) fn stage_created<P: Owner>(
    edit: &mut Edit<P>,
    metadata: &Metadata,
    format: FileFormat,
) -> Result<()> {
    let unsupported = metadata.identifier.is_some()
        || metadata.template.is_some()
        || metadata.revision.is_some()
        || metadata.created.is_some()
        || metadata.modified.is_some()
        || metadata.category.is_some()
        || metadata.company.is_some()
        || metadata.manager.is_some()
        || metadata.content_status.is_some()
        || metadata.content_type.is_some()
        || metadata.version.is_some()
        || metadata.last_printed.is_some()
        || metadata.hyperlink_base.is_some()
        || metadata.application.is_some();
    if unsupported {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "OpenDocument creation writes only title, subject, description, author, lastModifiedBy, language and keywords",
        )
        .with_format(format));
    }
    for (field, value) in [
        (Field::Title, &metadata.title),
        (Field::Subject, &metadata.subject),
        (Field::Description, &metadata.description),
        (Field::Author, &metadata.author),
        (Field::LastModifiedBy, &metadata.last_modified_by),
        (Field::Language, &metadata.language),
    ] {
        if let Some(value) = value {
            edit.set(field, Some(value))
                .map_err(|error| staging(error, format))?;
        }
    }
    if !metadata.keywords.is_empty() {
        edit.set_keywords(&metadata.keywords)
            .map_err(|error| staging(error, format))?;
    }
    Ok(())
}

/// Commit a staged metadata edit: the committed package, its patch and
/// whether the bytes changed.
///
/// # Errors
///
/// A refused candidate is the edit's fault (`invalid_value`, or
/// `unsupported_capability` for a signed or encrypted package); typed
/// limits keep their code.
pub(crate) fn commit<P: Owner>(edit: Edit<P>, format: FileFormat) -> Result<(P, Patch<P>, bool)> {
    let commit = edit.commit().map_err(|error| match error {
        litchi_odf_common::Error::Core(error) => staging_core(error, format),
        other => Error::from(other).with_format(format),
    })?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    Ok((package, patch, changed))
}

/// A staging refusal: the caller's value, not the file, is at fault.
fn staging(error: litchi_odf_common::Error, format: FileFormat) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => staging_core(error, format),
        other => Error::from(other).with_format(format),
    }
}

fn staging_core(error: litchi_core::Error, format: FileFormat) -> Error {
    let code = if matches!(
        error,
        litchi_core::Error::Unsupported(_) | litchi_core::Error::Refused(_)
    ) {
        Code::UnsupportedCapability
    } else {
        Code::InvalidValue
    };
    Error::from(error)
        .with_format(format)
        .caused_by_argument(code)
}
