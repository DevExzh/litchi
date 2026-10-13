//! Document metadata on the document root (`/`), shared by the OOXML
//! adapters (record 0253).
//!
//! OfficeCLI reads and sets document properties on the root path: the core
//! keys `title`, `author` (alias `creator`), `subject`, `keywords`,
//! `description`, `category`, `lastModifiedBy` and `revisionNumber`, and the
//! application keys `extended.template`, `extended.manager` and
//! `extended.company` (`ExtendedPropertiesHandler.TrySetExtendedProperty`).
//! This module admits those keys into a [`Metadata`] change; each format
//! adapter commits it through its leaf `metadata` transaction. Setting `null`
//! clears a field. Unlike OfficeCLI, a set never stamps `modified`
//! (ADR 0006, determinism).

use litchi_core::{FileFormat, Metadata};
use litchi_ooxml_common::properties;

use crate::error::{Code, Error, Result};
use crate::node::spec::PropSpec;
use crate::props::Props;
use crate::registry::{Access, Alias, Constraint, Entry, Operation, Table};
use crate::target::TargetArg;
use crate::value::{Scalar, ValueType};

/// One settable root metadata key.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Key {
    Title,
    Author,
    Subject,
    Keywords,
    Description,
    Category,
    LastModifiedBy,
    RevisionNumber,
    Template,
    Manager,
    Company,
}

const fn text(name: &'static str, key: Key) -> Entry<Key> {
    Entry {
        spec: PropSpec::new(name, ValueType::Text),
        access: Access::Set(key),
        constraint: Constraint::None,
        clear: true,
    }
}

/// The settable root keys, with OfficeCLI's names.
pub(crate) const WRITE_TABLE: Table<Key, 11, 1> = Table::new(
    [
        text("title", Key::Title),
        text("author", Key::Author),
        text("subject", Key::Subject),
        text("keywords", Key::Keywords),
        text("description", Key::Description),
        text("category", Key::Category),
        text("lastModifiedBy", Key::LastModifiedBy),
        text("revisionNumber", Key::RevisionNumber),
        text("extended.template", Key::Template),
        text("extended.manager", Key::Manager),
        text("extended.company", Key::Company),
    ],
    [Alias {
        name: "creator",
        entry: 1,
    }],
);

/// Whether a root-key name is one of [`WRITE_TABLE`]'s, so a root spec can
/// mark it settable.
pub(crate) const fn settable(name: &str) -> bool {
    let entries = [
        "title",
        "author",
        "subject",
        "keywords",
        "description",
        "category",
        "lastModifiedBy",
        "revisionNumber",
        "extended.template",
        "extended.manager",
        "extended.company",
    ];
    let mut index = 0;
    while index < entries.len() {
        if eq(entries[index], name) {
            return true;
        }
        index += 1;
    }
    false
}

const fn eq(left: &str, right: &str) -> bool {
    let (left, right) = (left.as_bytes(), right.as_bytes());
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if left[index] != right[index] {
            return false;
        }
        index += 1;
    }
    true
}

/// Mark the [`WRITE_TABLE`] keys of a root property table settable.
pub(crate) const fn root_capabilities<const N: usize>(props: &[PropSpec]) -> [PropSpec; N] {
    let mut result = [PropSpec::new("", ValueType::Text); N];
    let mut index = 0;
    while index < N {
        result[index] = props[index];
        if settable(props[index].name) {
            result[index].access.set = true;
        }
        index += 1;
    }
    result
}

/// A timestamp in .NET's round-trip form (`"o"`), as OfficeCLI prints
/// core-properties times: seven fraction digits, and `Z` for an instant.
pub(crate) fn round_trip(timestamp: litchi_core::metadata::Timestamp) -> String {
    let naive = timestamp.naive();
    format!(
        "{}.{:07}{}",
        naive.format("%Y-%m-%dT%H:%M:%S"),
        naive.and_utc().timestamp_subsec_nanos() / 100,
        if timestamp.utc().is_some() { "Z" } else { "" }
    )
}

/// Admitted root metadata assignments; `None` clears a field.
#[derive(Debug, Default)]
pub(crate) struct Assignments(Vec<(Key, Option<String>)>);

impl Assignments {
    /// Apply the assignments to `metadata`.
    pub(crate) fn apply(&self, metadata: &mut Metadata) {
        for (key, value) in &self.0 {
            let slot = match key {
                Key::Title => &mut metadata.title,
                Key::Author => &mut metadata.author,
                Key::Subject => &mut metadata.subject,
                Key::Keywords => &mut metadata.keywords,
                Key::Description => &mut metadata.description,
                Key::Category => &mut metadata.category,
                Key::LastModifiedBy => &mut metadata.last_modified_by,
                Key::RevisionNumber => &mut metadata.revision,
                Key::Template => &mut metadata.template,
                Key::Manager => &mut metadata.manager,
                Key::Company => &mut metadata.company,
            };
            slot.clone_from(value);
        }
    }
}

/// Whether a mutation targets the document root, where metadata lives.
pub(crate) fn is_root(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| path.path().is_root())
}

/// Admit root property assignments.
///
/// # Errors
///
/// Returns [`Code::UnsupportedProperty`] for a key that is not a settable
/// metadata key (with a suggestion), [`Code::InvalidArgument`] for a key set
/// twice, and the coercion errors of the registry.
pub(crate) fn admit(props: &Props<'_>, format: FileFormat) -> Result<Assignments> {
    let mut assignments = Vec::new();
    let mut seen = 0_u16;
    for (key, value) in props.iter() {
        let (index, entry) = WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_property(key.as_str().to_owned())
                .with_format(format)
        })?;
        let bit = 1_u16 << index;
        if seen & bit != 0 {
            return Err(Error::new(
                Code::InvalidArgument,
                "a metadata property was assigned more than once",
            )
            .with_property(entry.spec.name)
            .with_format(format));
        }
        seen |= bit;
        let assignment = entry.prepare(Operation::Set, value, |key, value| {
            Ok((
                key,
                match value {
                    Some(Scalar::Text(text)) => Some(text.to_owned()),
                    None => None,
                    Some(_) => {
                        return Err(Error::new(
                            Code::InternalError,
                            "metadata text coercion returned another type",
                        ));
                    },
                },
            ))
        })?;
        assignments.push(assignment);
    }
    Ok(Assignments(assignments))
}

/// Map a leaf metadata error onto the unified taxonomy.
pub(crate) fn native_error(error: properties::Error, format: FileFormat) -> Error {
    let error = match error {
        properties::Error::Unsupported { field } => Error::new(
            Code::UnsupportedProperty,
            "the format cannot store this metadata field",
        )
        .with_property(field.name()),
        properties::Error::Invalid(message) => Error::new(Code::InvalidValue, message),
        properties::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::new(
            Code::LimitExceeded,
            format!("metadata {kind:?} limit exceeded: {observed} > {maximum}"),
        ),
        properties::Error::Signed => crate::format::signature::refusal(format),
        properties::Error::Conflict(message) => Error::new(Code::ValidationError, message),
        other => Error::new(Code::CorruptFile, other.to_string()),
    };
    error.with_format(format)
}

/// Fold the application properties of an eagerly opened package into
/// open-time metadata. A malformed application part leaves those fields
/// unset here; a metadata edit reads the part strictly and reports it.
pub(crate) fn with_app_props(mut metadata: Metadata, package: &litchi_opc::OpcPackage) -> Metadata {
    if let Ok(Some(app)) = properties::app::read(package) {
        app.project_into(&mut metadata);
    }
    metadata
}

/// [`with_app_props`] over a source-backed package.
pub(crate) fn with_source_app_props(
    mut metadata: Metadata,
    package: &litchi_opc::SourceBackedPackage,
) -> Metadata {
    if let Ok(Some(app)) = properties::app::read_source_backed(package) {
        app.project_into(&mut metadata);
    }
    metadata
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn admission_uses_officecli_names_and_clears_with_null() {
        let props = Props::new()
            .with("creator", "Ada")
            .with("extended.company", "Litchi")
            .with("title", crate::props::PropValue::Null);
        let assignments = admit(&props, FileFormat::Xlsx).unwrap();
        let mut metadata = Metadata {
            title: Some("Old".to_owned()),
            ..Metadata::default()
        };
        assignments.apply(&mut metadata);
        assert_eq!(metadata.author.as_deref(), Some("Ada"));
        assert_eq!(metadata.company.as_deref(), Some("Litchi"));
        assert_eq!(metadata.title, None);

        let error = admit(&Props::new().with("titel", "x"), FileFormat::Xlsx).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        let error = admit(
            &Props::new().with("author", "a").with("creator", "b"),
            FileFormat::Pptx,
        )
        .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        assert!(settable("extended.manager"));
        assert!(!settable("created"));
    }
}
