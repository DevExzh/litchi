//! The `/metadata` set path of the OOXML formats (XLSX, PPTX, XLSB), record
//! 0253, and DOCX (records 0338 and 0342).
//!
//! `set("/metadata", props)` stages document properties through each
//! format's leaf `metadata` transaction. OfficeCLI's spelling, a set on the
//! root `/` (`set / title=...`), routes to the same writer for DOCX, XLSX
//! and PPTX (record 0338); [`is_root`] and [`names_metadata`] decide it. The keys are the 22 OfficeCLI
//! root-metadata spellings the legacy OLE formats accept on the same path
//! (record 0254; matched ASCII case-insensitively), plus `custom.<name>` for
//! custom properties. `null` removes a property. `keywords` takes a list,
//! or text split at `;` and `,` (record 0354). A key no OOXML part stores
//! (`contentType`) is refused by the leaf with a typed
//! `UnsupportedProperty` refusal whose reason is fixed and content-free.
//! Unlike OfficeCLI, a set never stamps `modified` (ADR 0006,
//! determinism). The path is set-only; reads go through `File::metadata()`
//! and the root node's properties.

#![cfg_attr(
    not(any(feature = "pptx", feature = "xlsx", feature = "xlsb")),
    allow(
        dead_code,
        reason = "a DOCX-only build uses the core-properties admission only"
    )
)]

use litchi_core::FileFormat;
use litchi_core::metadata::{Metadata, Timestamp};
use litchi_ooxml_common::custom;
use litchi_ooxml_common::properties;
#[cfg(test)]
use std::mem::size_of;

use crate::error::{Code, Error, Result};
use crate::props::{PropValue, Props};
use crate::target::TargetArg;

/// The path segment of the metadata node.
const SEGMENT: &str = "metadata";

/// The prefix of custom property keys.
const CUSTOM_PREFIX: &str = "custom.";

/// One [`Metadata`] field a key writes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Field {
    Application,
    Author,
    Category,
    Company,
    ContentStatus,
    ContentType,
    Created,
    Description,
    Keywords,
    Language,
    LastModifiedBy,
    LastPrinted,
    Manager,
    Modified,
    Revision,
    Subject,
    Template,
    Title,
    Version,
}

/// Every key, sorted by ASCII-lowercase name for binary search; the same
/// spellings as record 0254's OLE table.
static KEYS: [(&str, Field); 22] = [
    ("application", Field::Application),
    ("author", Field::Author),
    ("category", Field::Category),
    ("company", Field::Company),
    ("contentStatus", Field::ContentStatus),
    ("contentType", Field::ContentType),
    ("created", Field::Created),
    ("description", Field::Description),
    ("extended.company", Field::Company),
    ("extended.manager", Field::Manager),
    ("extended.template", Field::Template),
    ("keywords", Field::Keywords),
    ("language", Field::Language),
    ("lastModifiedBy", Field::LastModifiedBy),
    ("lastPrinted", Field::LastPrinted),
    ("manager", Field::Manager),
    ("modified", Field::Modified),
    ("revisionNumber", Field::Revision),
    ("subject", Field::Subject),
    ("template", Field::Template),
    ("title", Field::Title),
    ("version", Field::Version),
];

const _: () = {
    let mut index = 1;
    while index < KEYS.len() {
        assert!(
            ascii_lower_cmp(KEYS[index - 1].0, KEYS[index].0).is_lt(),
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

impl Field {
    /// Whether the core-properties part stores the field.
    #[cfg_attr(not(feature = "docx"), allow(dead_code, reason = "DOCX only"))]
    const fn is_core(self) -> bool {
        !matches!(
            self,
            Self::Application | Self::Company | Self::ContentType | Self::Manager | Self::Template
        )
    }

    /// Whether the application-properties part stores the field.
    #[cfg_attr(not(feature = "docx"), allow(dead_code, reason = "DOCX only"))]
    const fn is_app(self) -> bool {
        matches!(
            self,
            Self::Application | Self::Company | Self::Manager | Self::Template
        )
    }
}

fn lookup(name: &str) -> Option<Field> {
    KEYS.binary_search_by(|(key, _)| ascii_lower_cmp(key, name))
        .ok()
        .map(|index| KEYS[index].1)
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

/// Whether `target` is the document root `/`: OfficeCLI's spelling of a
/// metadata set (record 0338).
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
pub(crate) fn is_root(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        let path = path.path();
        path.is_absolute() && path.steps().is_empty()
    })
}

/// Whether any of `props` is a metadata key or `custom.<name>`.
#[cfg(feature = "pptx")]
pub(crate) fn names_metadata(props: &Props<'_>) -> bool {
    props.iter().any(|(key, _)| {
        let name = key.as_str();
        lookup(name).is_some() || strip_custom(name).is_some()
    })
}

/// Whether `name` is one of the metadata keys, ASCII case-insensitively
/// (usable in `const` tables).
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
const fn is_key(name: &str) -> bool {
    let mut index = 0;
    while index < KEYS.len() {
        if ascii_lower_cmp(KEYS[index].0, name).is_eq() {
            return true;
        }
        index += 1;
    }
    false
}

/// `props` with every metadata key settable: the root element's capability
/// table, whose set route is the metadata writer (record 0338). The keys
/// stay in their read order; only `access.set` changes.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
pub(crate) const fn root_capability<const N: usize>(
    props: &[crate::node::spec::PropSpec],
) -> [crate::node::spec::PropSpec; N] {
    assert!(
        props.len() == N,
        "the root capability table keeps every key"
    );
    let mut out = [crate::node::spec::PropSpec::new("", crate::value::ValueType::Text); N];
    let mut index = 0;
    while index < N {
        let mut prop = props[index];
        if is_key(prop.name) {
            prop.access.set = true;
        }
        out[index] = prop;
        index += 1;
    }
    out
}

/// One admitted value.
#[derive(Debug, Clone)]
enum Assigned {
    Text(Option<String>),
    Time(Option<Timestamp>),
    /// The keyword list; empty removes the keywords.
    Keywords(Vec<String>),
}

/// Admitted `/metadata` assignments; `None` removes.
#[derive(Debug, Default)]
pub(crate) struct Assignments {
    fields: Vec<(Field, Assigned)>,
    custom: Vec<(String, Option<custom::Value>)>,
}

impl Assignments {
    /// Conservative logical bytes retained by the admitted metadata request.
    /// This is used before the leaf metadata editor materializes its before /
    /// after property scope; it includes names, authored text and the
    /// structural descriptors retained by the patch.
    #[cfg(test)]
    pub(crate) fn retained_bytes(&self) -> usize {
        let fields =
            self.fields
                .iter()
                .fold(size_of::<(Field, Assigned)>(), |bytes, (_, value)| {
                    bytes.saturating_add(match value {
                        Assigned::Text(text) => text.as_ref().map_or(0, String::len),
                        Assigned::Time(_) => size_of::<Timestamp>(),
                        Assigned::Keywords(values) => values
                            .iter()
                            .fold(0usize, |total, value| total.saturating_add(value.len())),
                    })
                });
        self.custom.iter().fold(
            fields.saturating_add(size_of::<(String, Option<custom::Value>)>()),
            |bytes, (name, value)| {
                bytes
                    .saturating_add(name.len())
                    .saturating_add(size_of::<custom::Value>())
                    .saturating_add(match value {
                        Some(custom::Value::Text(text)) => text.len(),
                        _ => 0,
                    })
            },
        )
    }

    /// Apply the field assignments to `metadata`.
    pub(crate) fn apply(&self, metadata: &mut Metadata) {
        for (field, value) in &self.fields {
            match value {
                Assigned::Text(text) => {
                    let slot = match field {
                        Field::Application => &mut metadata.application,
                        Field::Author => &mut metadata.author,
                        Field::Category => &mut metadata.category,
                        Field::Company => &mut metadata.company,
                        Field::ContentStatus => &mut metadata.content_status,
                        Field::ContentType => &mut metadata.content_type,
                        Field::Description => &mut metadata.description,
                        Field::Language => &mut metadata.language,
                        Field::LastModifiedBy => &mut metadata.last_modified_by,
                        Field::Manager => &mut metadata.manager,
                        Field::Revision => &mut metadata.revision,
                        Field::Subject => &mut metadata.subject,
                        Field::Template => &mut metadata.template,
                        Field::Title => &mut metadata.title,
                        Field::Version => &mut metadata.version,
                        Field::Created | Field::Modified | Field::LastPrinted | Field::Keywords => {
                            continue;
                        },
                    };
                    slot.clone_from(text);
                },
                Assigned::Keywords(keywords) => metadata.keywords.clone_from(keywords),
                Assigned::Time(time) => {
                    let slot = match field {
                        Field::Created => &mut metadata.created,
                        Field::Modified => &mut metadata.modified,
                        Field::LastPrinted => &mut metadata.last_printed,
                        _ => continue,
                    };
                    *slot = *time;
                },
            }
        }
    }

    /// Apply the core-property assignments to `metadata`; the application
    /// fields are left to [`Self::apply_app`] (DOCX, record 0342).
    #[cfg(feature = "docx")]
    pub(crate) fn apply_core(&self, metadata: &mut Metadata) {
        let core = Self {
            fields: self
                .fields
                .iter()
                .filter(|(field, _)| field.is_core())
                .cloned()
                .collect(),
            custom: Vec::new(),
        };
        core.apply(metadata);
    }

    /// Whether any application property (`company`, `manager`, `template`,
    /// `application`, or their `extended.*` spellings) is assigned.
    #[cfg(feature = "docx")]
    pub(crate) fn has_app(&self) -> bool {
        self.fields.iter().any(|(field, _)| field.is_app())
    }

    /// The canonical facade field name of an assigned field no OOXML part
    /// stores.  The input spelling is `contentType`, but errors name the
    /// shared [`Metadata`] field (`content_type`).
    pub(crate) fn unsupported_property(&self) -> Option<&'static str> {
        self.fields
            .iter()
            .any(|(field, _)| *field == Field::ContentType)
            .then_some("content_type")
    }

    /// `base` with the application-property assignments applied: every
    /// element the assignments do not name keeps its value (DOCX, record
    /// 0342).
    #[cfg(feature = "docx")]
    pub(crate) fn apply_app(&self, base: Option<properties::app::Props>) -> properties::app::Props {
        let mut props = base.unwrap_or_default();
        for (field, value) in &self.fields {
            let Assigned::Text(text) = value else {
                continue;
            };
            let slot = match field {
                Field::Application => &mut props.application,
                Field::Company => &mut props.company,
                Field::Manager => &mut props.manager,
                Field::Template => &mut props.template,
                _ => continue,
            };
            slot.clone_from(text);
        }
        props
    }

    /// Whether any custom property is assigned.
    pub(crate) fn has_custom(&self) -> bool {
        !self.custom.is_empty()
    }

    /// Apply the custom assignments to a copy of the staged custom
    /// properties.
    ///
    /// # Errors
    ///
    /// `invalid_value` for a name or value the custom part refuses.
    pub(crate) fn apply_custom(
        &self,
        props: &custom::Props,
        format: FileFormat,
    ) -> Result<custom::Props> {
        let mut props = props.clone();
        for (name, value) in &self.custom {
            match value {
                Some(value) => {
                    props
                        .insert(name.clone(), value.clone())
                        .map_err(|_| invalid(&format!("{CUSTOM_PREFIX}{name}"), format))?;
                },
                None => {
                    props.remove(name);
                },
            }
        }
        Ok(props)
    }
}

/// Return the canonical shared field name that an OOXML metadata transaction
/// cannot store, when the authored value contains one.  This is used for
/// complete metadata creation, where the native refusal deliberately carries
/// no property detail so its reason remains fixed and content-free.
pub(crate) fn unsupported_field(metadata: &Metadata) -> Option<&'static str> {
    litchi_core::metadata::Field::ALL
        .into_iter()
        .find_map(|field| {
            (metadata.is_set(field)
                && !properties::CORE_FIELDS.contains(&field)
                && !properties::app::METADATA_FIELDS.contains(&field))
            .then_some(field.name())
        })
}

/// Restore caller-facing property context after mapping a typed native
/// refusal.  The leaf refusal intentionally redacts the field from its fixed
/// reason; only the facade boundary has the request context needed to name
/// the shared field.  Existing property context, when any, always wins.
pub(crate) fn with_unsupported_property(mut error: Error, property: Option<&'static str>) -> Error {
    if error.code() == Code::UnsupportedProperty
        && error.property().is_none()
        && let Some(property) = property
    {
        error = error.with_property(property);
    }
    error
}

/// Admit `/metadata` property assignments.
///
/// # Errors
///
/// `invalid_argument` for an empty or repeated assignment,
/// `unsupported_property` for an unknown key (with a suggestion), and
/// `invalid_value` for a value of the wrong kind.
pub(crate) fn admit(props: &Props<'_>, format: FileFormat) -> Result<Assignments> {
    if props.is_empty() {
        return Err(Error::new(
            Code::InvalidArgument,
            "a metadata set needs at least one property",
        )
        .with_format(format));
    }
    let mut assignments = Assignments::default();
    for (key, value) in props.iter() {
        let name = key.as_str();
        if let Some(custom) = strip_custom(name) {
            if assignments
                .custom
                .iter()
                .any(|(seen, _)| seen.eq_ignore_ascii_case(custom))
            {
                return Err(repeated(name, format));
            }
            assignments
                .custom
                .push((custom.to_owned(), custom_value(name, value, format)?));
            continue;
        }
        let Some(field) = lookup(name) else {
            return Err(Error::new(
                Code::UnsupportedProperty,
                "the metadata key is not a document property",
            )
            .with_property(name.to_owned())
            .with_format(format)
            .with_suggestion("use a root-metadata key such as `title`, or `custom.<name>`"));
        };
        if assignments.fields.iter().any(|(seen, _)| *seen == field) {
            return Err(repeated(name, format));
        }
        let assigned = match field {
            Field::Created | Field::Modified | Field::LastPrinted => Assigned::Time(match value {
                PropValue::Null => None,
                PropValue::Text(text) => {
                    Some(Timestamp::parse(text).map_err(|_| invalid(name, format))?)
                },
                _ => return Err(wrong_kind(name, format)),
            }),
            Field::Keywords => Assigned::Keywords(
                crate::props::keyword_list(value).ok_or_else(|| wrong_kind(name, format))?,
            ),
            _ => Assigned::Text(match value {
                PropValue::Null => None,
                PropValue::Text(text) => Some(text.to_string()),
                PropValue::Int(_) | PropValue::Unsigned(_) => Some(value.to_string()),
                _ => return Err(wrong_kind(name, format)),
            }),
        };
        assignments.fields.push((field, assigned));
    }
    Ok(assignments)
}

fn strip_custom(name: &str) -> Option<&str> {
    let prefix = name.get(..CUSTOM_PREFIX.len())?;
    prefix
        .eq_ignore_ascii_case(CUSTOM_PREFIX)
        .then(|| &name[CUSTOM_PREFIX.len()..])
        .filter(|custom| !custom.is_empty())
}

fn custom_value(
    key: &str,
    value: &PropValue<'_>,
    format: FileFormat,
) -> Result<Option<custom::Value>> {
    Ok(Some(match value {
        PropValue::Null => return Ok(None),
        PropValue::Bool(value) => custom::Value::Bool(*value),
        PropValue::Int(value) => match i32::try_from(*value) {
            Ok(value) => custom::Value::I32(value),
            Err(_) => custom::Value::I64(*value),
        },
        PropValue::Number(value) if value.is_finite() => custom::Value::F64(*value),
        PropValue::Text(text) => custom::Value::Text(text.to_string()),
        _ => return Err(wrong_kind(key, format)),
    }))
}

fn repeated(key: &str, format: FileFormat) -> Error {
    Error::new(
        Code::InvalidArgument,
        "a metadata property was assigned more than once",
    )
    .with_property(key.to_owned())
    .with_format(format)
}

fn wrong_kind(key: &str, format: FileFormat) -> Error {
    Error::new(
        Code::InvalidValue,
        "the metadata property requires text (RFC 3339 for timestamps)",
    )
    .with_property(key.to_owned())
    .with_format(format)
}

fn invalid(key: &str, format: FileFormat) -> Error {
    Error::new(
        Code::InvalidValue,
        "the metadata value is not representable in OOXML document properties",
    )
    .with_property(key.to_owned())
    .with_format(format)
}

/// Map a leaf metadata error onto the unified taxonomy.
pub(crate) fn native_error(error: properties::Error, format: FileFormat) -> Error {
    let error = match error {
        properties::Error::Refused(refusal) => {
            Error::new(crate::error::refusal_code(refusal.kind()), refusal.reason())
        },
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
        properties::Error::PatchConflict => Error::new(
            Code::InvalidInput,
            "the metadata patch base does not match the package",
        ),
        properties::Error::SourceChanged { .. } => Error::new(
            Code::IoError,
            "the metadata source changed after it was opened",
        ),
        properties::Error::Internal(_) => Error::new(
            Code::InternalError,
            "the metadata transaction invariant failed",
        ),
        other => Error::new(Code::CorruptFile, other.to_string()),
    };
    error.with_format(format)
}

/// A timestamp in .NET's round-trip form (`"o"`), as OfficeCLI prints
/// core-properties times: seven fraction digits, and `Z` for an instant.
#[cfg_attr(
    not(feature = "pptx"),
    allow(dead_code, reason = "only PPTX node views print core-property times")
)]
pub(crate) fn round_trip(timestamp: Timestamp) -> String {
    let naive = timestamp.naive();
    format!(
        "{}.{:07}{}",
        naive.format("%Y-%m-%dT%H:%M:%S"),
        naive.and_utc().timestamp_subsec_nanos() / 100,
        if timestamp.utc().is_some() { "Z" } else { "" }
    )
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
            .with("Author", "Ada")
            .with("extended.company", "Litchi")
            .with("created", "2024-01-02T03:04:05Z")
            .with("custom.Reviewed", true)
            .with("title", PropValue::Null);
        let assignments = admit(&props, FileFormat::Xlsx).unwrap();
        let mut metadata = Metadata {
            title: Some("Old".to_owned()),
            ..Metadata::default()
        };
        assignments.apply(&mut metadata);
        assert_eq!(metadata.author.as_deref(), Some("Ada"));
        assert_eq!(metadata.company.as_deref(), Some("Litchi"));
        assert_eq!(metadata.title, None);
        assert_eq!(
            metadata.created,
            Some(Timestamp::parse("2024-01-02T03:04:05Z").unwrap())
        );
        let custom = assignments
            .apply_custom(&custom::Props::new(), FileFormat::Xlsx)
            .unwrap();
        assert_eq!(custom.get("Reviewed"), Some(&custom::Value::Bool(true)));

        let error = admit(&Props::new().with("titel", "x"), FileFormat::Xlsx).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
        let error = admit(
            &Props::new()
                .with("company", "a")
                .with("extended.company", "b"),
            FileFormat::Pptx,
        )
        .unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        let error =
            admit(&Props::new().with("created", "yesterday"), FileFormat::Xlsb).unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn stale_ooxml_metadata_replay_is_invalid_input() {
        for format in [
            FileFormat::Docx,
            FileFormat::Pptx,
            FileFormat::Xlsx,
            FileFormat::Xlsb,
        ] {
            let error = native_error(properties::Error::PatchConflict, format);
            assert_eq!(error.code(), Code::InvalidInput, "{format:?}");
        }
    }

    #[test]
    fn native_metadata_refusal_and_source_change_keep_their_codes() {
        let refusal = properties::Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedProperty,
            "set_metadata",
            "the OOXML metadata route does not support this property",
        ));
        let error = native_error(refusal, FileFormat::Xlsx);
        assert_eq!(error.code(), Code::UnsupportedProperty);
        assert_eq!(error.property(), None);

        let error = with_unsupported_property(error, Some("content_type"));
        assert_eq!(error.property(), Some("content_type"));

        let error = native_error(
            properties::Error::SourceChanged {
                expected: litchi_core::SourceVersion::new(8, 1),
                observed: litchi_core::SourceVersion::new(8, 2),
            },
            FileFormat::Xlsb,
        );
        assert_eq!(error.code(), Code::IoError);

        let error = native_error(
            properties::Error::Package(litchi_ooxml_common::Error::Invalid(
                "extended property Company contains XML 1.0-forbidden character U+0001".to_owned(),
            )),
            FileFormat::Xlsx,
        );
        assert_eq!(error.code(), Code::CorruptFile);

        let error = native_error(
            properties::Error::Invalid(
                "extended property Company contains XML 1.0-forbidden character U+0000".to_owned(),
            ),
            FileFormat::Xlsx,
        );
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn unsupported_metadata_fields_use_shared_names_at_the_facade_boundary() {
        let assignments = admit(
            &Props::new().with("contentType", "application/example"),
            FileFormat::Xlsx,
        )
        .expect("contentType admission");
        assert_eq!(assignments.unsupported_property(), Some("content_type"));

        for (metadata, expected) in [
            (
                Metadata {
                    content_type: Some("application/example".to_owned()),
                    ..Metadata::default()
                },
                "content_type",
            ),
            (
                Metadata {
                    last_backup: Some(Timestamp::parse("2024-01-01T00:00:00").unwrap()),
                    ..Metadata::default()
                },
                "last_backup",
            ),
            (
                Metadata {
                    codepage: Some(1252),
                    ..Metadata::default()
                },
                "codepage",
            ),
        ] {
            assert_eq!(unsupported_field(&metadata), Some(expected));
        }

        assert_eq!(unsupported_field(&Metadata::default()), None);
        assert_eq!(
            unsupported_field(&Metadata {
                company: Some("Litchi".to_owned()),
                ..Metadata::default()
            }),
            None
        );
    }

    #[test]
    fn retained_metadata_charge_includes_names_values_and_custom_shape() {
        let short = admit(&Props::new().with("title", "x"), FileFormat::Xlsx).unwrap();
        let long = admit(
            &Props::new()
                .with("title", "a much longer authored title")
                .with("custom.ReviewedBy", "an authored reviewer"),
            FileFormat::Xlsx,
        )
        .unwrap();

        assert!(long.retained_bytes() > short.retained_bytes());
        assert!(long.retained_bytes() >= "ReviewedBy".len() + "an authored reviewer".len());
    }
}
