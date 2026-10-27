//! The shared, exact and reversible metadata transaction of an OOXML
//! package.
//!
//! An OOXML package holds three document-property families, each in its own
//! package-level part: core properties (`docProps/core.xml`), application
//! properties (`docProps/app.xml`, [`super::app`]) and custom properties
//! (`docProps/custom.xml`, [`crate::custom`]). Every host format edits them
//! the same way:
//!
//! 1. [`Snapshot::read`] (owning [`OpcPackage`]) or
//!    [`Snapshot::read_source_backed`] ([`SourceBackedPackage`]) captures the
//!    strictly validated graph and the exact part bytes of every family;
//! 2. [`Edit::new`] starts an [`Edit`] over an owning package, which stages
//!    a whole [`Metadata`], application properties, or custom properties;
//! 3. [`Edit::commit`] computes a [`Patch`]: the exact before and after bytes
//!    of each part that changes, the part created when a family was absent,
//!    or the part removed when the custom family is emptied, and publishes
//!    it as a new [`OpcPackage`] in a [`Commit`] (the source is unchanged);
//! 4. the host wraps the new package as its new owner, and replays a patch
//!    with [`Patch::apply`] on a copy of its package.
//!
//! An owner that publishes over a retained source ([`SourceBackedPackage`])
//! stages the same change with a detached [`Snapshot::plan`] and publishes
//! the [`Plan::finish`] patch itself.
//!
//! Each [`Metadata`] field lands in the family that stores it
//! ([`super::CORE_FIELDS`], [`super::app::METADATA_FIELDS`]); changing a
//! field no OOXML family stores (`content_type`, `last_backup`, `codepage`)
//! is refused with [`Error::Refused`], never dropped. A field left
//! unchanged keeps its exact lexical form, and an exact no-op commits an
//! empty patch that publishes the package unchanged, a signed one included.
//! A changed patch refuses a signed package ([`Error::Signed`], ADR 0006).
//! Nothing stamps `modified`, `lastModifiedBy` or a revision unless the
//! caller sets it (ADR 0006, determinism).

use std::collections::BTreeSet;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::edit::{Refusal, RefusalKind};
use litchi_core::metadata::{Field, Metadata};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::rel::TargetMode;
use litchi_opc::{
    OpcError, OpcPackage, PackURI, PreparedTopology, SourceBackedPackage, SourceTopologyPlan,
};

use super::package_graph::{
    GraphPart as _, PackageMetadata, ensure_only_inbound, free_part_name, names_conflict,
    other_inbound,
};
use super::{CORE_FIELDS, Core, Dialect, app, graph, read};
use crate::custom::{self, Host};
use crate::leniency::ToleranceReport;
use crate::signature::{refuse_signed_change, settle_unsigned_change};

fn merge_core_metadata(target: &mut Metadata, core: &Metadata) {
    target.title.clone_from(&core.title);
    target.subject.clone_from(&core.subject);
    target.author.clone_from(&core.author);
    target.keywords.clone_from(&core.keywords);
    target.description.clone_from(&core.description);
    target.identifier.clone_from(&core.identifier);
    target.language.clone_from(&core.language);
    target.last_modified_by.clone_from(&core.last_modified_by);
    target.revision.clone_from(&core.revision);
    target.created = core.created;
    target.modified = core.modified;
    target.category.clone_from(&core.category);
    target.content_status.clone_from(&core.content_status);
    target.version.clone_from(&core.version);
    target.last_printed = core.last_printed;
}

fn clear_core_metadata(metadata: &mut Metadata) {
    metadata.title = None;
    metadata.subject = None;
    metadata.author = None;
    metadata.keywords.clear();
    metadata.description = None;
    metadata.identifier = None;
    metadata.language = None;
    metadata.last_modified_by = None;
    metadata.revision = None;
    metadata.created = None;
    metadata.modified = None;
    metadata.category = None;
    metadata.content_status = None;
    metadata.version = None;
    metadata.last_printed = None;
}

/// The resource a metadata limit bounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The UTF-8 bytes of one text value.
    TextBytes,
    /// The bytes of one property part.
    XmlBytes,
    /// The XML events of one property part.
    XmlEvents,
    /// The number of custom properties.
    CustomProperties,
    /// Package relationship identifiers or part names available for a new
    /// family part.
    Names,
    /// Retained bytes of the complete property patch, including its before
    /// and after part payloads. This is checked before the patch is built.
    RetainedBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::TextBytes => "metadata text bytes",
            Self::XmlBytes => "metadata XML bytes",
            Self::XmlEvents => "metadata XML events",
            Self::CustomProperties => "metadata custom properties",
            Self::Names => "metadata property names",
            Self::RetainedBytes => "metadata retained bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Options selected before a document-property edit captures its replacement
/// families. The retained ceiling covers the exact semantic patch, including
/// before/after payloads and change descriptors.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    max_retained_bytes: Option<usize>,
    owner: Option<OwnerAdmission>,
}

impl EditOptions {
    /// Creates an unbounded metadata edit.
    pub const fn new() -> Self {
        Self {
            max_retained_bytes: None,
            owner: None,
        }
    }

    /// Sets the retained-patch ceiling for this edit.
    pub const fn max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = Some(maximum);
        self
    }

    /// Returns the selected retained-patch ceiling.
    #[must_use]
    pub const fn max_retained_bytes_value(self) -> Option<usize> {
        self.max_retained_bytes
    }

    /// Supplies the immutable host-owner census used by the exact package
    /// endpoint admission. This is a format bridge; ordinary callers keep
    /// the default patch-only policy.
    #[doc(hidden)]
    pub const fn with_owner_admission(mut self, owner: OwnerAdmission) -> Self {
        self.owner = Some(owner);
        self
    }
}

/// Why a metadata read, edit or patch failed.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed metadata request or source topology cannot be carried by
    /// this lossless route.
    #[error(transparent)]
    Refused(Refusal),
    /// A value violates its family's lexical or structural rules, or a part
    /// holds content its writer cannot carry.
    #[error("invalid OOXML metadata: {0}")]
    Invalid(String),
    /// A bounded resource exceeds its maximum.
    #[error("OOXML metadata {kind:?} limit exceeded: {observed} > {maximum}")]
    Limit {
        /// The bounded resource.
        kind: LimitKind,
        /// The observed amount, or `usize::MAX` on overflow.
        observed: usize,
        /// The largest accepted amount.
        maximum: usize,
    },
    /// A changed edit of a signed package needs the signature removed first.
    #[error("a changed metadata edit of a signed package requires removing the signature first")]
    Signed,
    /// A patch or staged composition does not match the recipient base.
    #[error("OOXML metadata patch base does not match the recipient")]
    PatchConflict,
    /// The retained positional source changed while a source-backed property
    /// operation was reading or authenticating it.
    #[error("OOXML metadata source changed")]
    SourceChanged {
        /// The source version captured at open.
        expected: litchi_core::SourceVersion,
        /// The source version observed during the operation.
        observed: litchi_core::SourceVersion,
    },
    /// The transaction found an invariant that its own validation should have
    /// guaranteed.
    #[error("OOXML metadata transaction invariant failed: {0}")]
    Internal(&'static str),
    /// Another package, graph or XML failure. A nested source change is
    /// promoted to [`Error::SourceChanged`] before it reaches this wrapper.
    #[error(transparent)]
    Package(crate::Error),
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        match error {
            crate::Error::Refused(refusal) => Self::Refused(refusal),
            crate::Error::PatchConflict => Self::PatchConflict,
            crate::Error::Signed => Self::Signed,
            crate::Error::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            crate::Error::Internal(reason) => Self::Internal(reason),
            crate::Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy) => Self::Signed,
            crate::Error::Opc(error) => match source_changed(&error) {
                Some((expected, observed)) => Self::SourceChanged { expected, observed },
                None => Self::Package(crate::Error::Opc(error)),
            },
            crate::Error::Invalid(message) => Self::Invalid(message),
            crate::Error::Limit {
                kind: crate::LimitKind::Properties(kind),
                maximum,
                observed,
            } => Self::Limit {
                kind,
                observed,
                maximum,
            },
            crate::Error::Limit {
                kind: resource,
                maximum: max,
                observed: actual,
            } => {
                let name = resource.as_str();
                let kind = if name.contains("text") {
                    Some(LimitKind::TextBytes)
                } else if name.contains("events") {
                    Some(LimitKind::XmlEvents)
                } else if name.contains("XML bytes") {
                    Some(LimitKind::XmlBytes)
                } else if name.contains("custom") {
                    Some(LimitKind::CustomProperties)
                } else if name.contains("names") || name.contains("identifiers") {
                    Some(LimitKind::Names)
                } else {
                    None
                };
                match kind {
                    Some(kind) => Self::Limit {
                        kind,
                        observed: actual,
                        maximum: max,
                    },
                    None => Self::Package(crate::Error::Limit {
                        kind: resource,
                        maximum: max,
                        observed: actual,
                    }),
                }
            },
            other => Self::Package(other),
        }
    }
}

impl From<OpcError> for Error {
    fn from(error: OpcError) -> Self {
        Self::from(crate::Error::Opc(error))
    }
}

/// Preserve malformed source XML and metadata graphs as package corruption;
/// authored codec failures use [`From<crate::Error>`] and remain invalid
/// values. A source change still gets its typed freshness error.
fn source_error(error: crate::Error) -> Error {
    match error {
        error @ crate::Error::Invalid(_) => Error::Package(error),
        other => Error::from(other),
    }
}

/// Preserve a malformed source graph discovered by a local transaction
/// helper as package corruption.  The local `Invalid` variant is also used
/// for authored value validation, so only callers that have just inspected
/// retained source state use this conversion.
fn source_local_error(error: Error) -> Error {
    match error {
        Error::Invalid(message) => Error::Package(crate::Error::Invalid(message)),
        other => other,
    }
}

fn app_plan_error(error: crate::Error) -> Error {
    match error {
        crate::Error::Limit { .. } => Error::from(error),
        crate::Error::Invalid(message) => Error::Package(crate::Error::Invalid(message)),
        other => Error::from(other),
    }
}

fn generated_codec_error(error: crate::Error, reason: &'static str) -> Error {
    match error {
        crate::Error::Limit { .. } => Error::from(error),
        crate::Error::Invalid(_)
        | crate::Error::Xml(_)
        | crate::Error::Decode(_)
        | crate::Error::Allocation { .. } => Error::Internal(reason),
        other => Error::from(other),
    }
}

fn authored_codec_error(error: crate::Error, reason: &'static str) -> Error {
    match error {
        crate::Error::Limit { .. } => Error::from(error),
        crate::Error::Invalid(message) => Error::Invalid(message),
        crate::Error::Xml(_) | crate::Error::Decode(_) | crate::Error::Allocation { .. } => {
            Error::Internal(reason)
        },
        other => Error::from(other),
    }
}

fn app_authored_error(error: crate::Error) -> Error {
    authored_codec_error(
        error,
        "application-properties prepared output invariant failed",
    )
}

fn app_prepared_error(error: crate::Error) -> Error {
    generated_codec_error(
        error,
        "application-properties prepared output invariant failed",
    )
}

fn core_generated_error(error: crate::Error) -> Error {
    generated_codec_error(error, "core-properties prepared output invariant failed")
}

fn custom_generated_error(error: crate::Error) -> Error {
    generated_codec_error(error, "custom-properties prepared output invariant failed")
}

fn patch_output_error(error: crate::Error) -> Error {
    generated_codec_error(error, "metadata patch output failed validation")
}

fn source_changed(
    error: &OpcError,
) -> Option<(litchi_core::SourceVersion, litchi_core::SourceVersion)> {
    match error {
        OpcError::SourceChanged { expected, actual } => Some((*expected, *actual)),
        OpcError::IncompleteOutput { source, .. } => source_changed(source),
        _ => None,
    }
}

fn unsupported_property() -> Error {
    Error::Refused(Refusal::new(
        RefusalKind::UnsupportedProperty,
        "set_metadata",
        "the OOXML metadata route does not support this property",
    ))
}

fn unsupported_custom_content() -> Error {
    Error::Refused(Refusal::new(
        RefusalKind::UnsupportedContent,
        "set_custom_props",
        "rewriting unsupported custom-property content is refused",
    ))
}

fn unsupported_core_content() -> Error {
    Error::Refused(Refusal::new(
        RefusalKind::UnsupportedContent,
        "set_metadata",
        "rewriting unsupported core-properties content is refused",
    ))
}

fn shared_family_removal() -> Error {
    Error::Refused(Refusal::new(
        RefusalKind::UnsupportedContent,
        "remove_metadata",
        "removing a shared OOXML metadata part is refused",
    ))
}

type Result<T> = std::result::Result<T, Error>;

/// One package-level document-property family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum Family {
    /// Core properties: `docProps/core.xml`.
    Core,
    /// Application (extended) properties: `docProps/app.xml`.
    App,
    /// Custom properties: `docProps/custom.xml`.
    Custom,
}

impl Family {
    /// The content type of the family's part.
    #[must_use]
    pub const fn content_type(self) -> &'static str {
        match self {
            Self::Core => ct::OPC_CORE_PROPERTIES,
            Self::App => ct::OFC_EXTENDED_PROPERTIES,
            Self::Custom => ct::OFC_CUSTOM_PROPERTIES,
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Core => "core-properties",
            Self::App => "application-properties",
            Self::Custom => "custom-properties",
        }
    }

    const fn stem(self) -> &'static str {
        match self {
            Self::Core => "/docProps/core",
            Self::App => "/docProps/app",
            Self::Custom => "/docProps/custom",
        }
    }
}

const fn family_index(family: Family) -> usize {
    match family {
        Family::Core => 0,
        Family::App => 1,
        Family::Custom => 2,
    }
}

/// The exact state of a present family part.
#[derive(Debug, Clone, PartialEq, Eq)]
struct Present {
    part: PackURI,
    relationship_id: String,
    relationship_type: String,
    relationship_target: String,
    bytes: Arc<Vec<u8>>,
}

/// One family as a snapshot captured it.
#[derive(Debug, Clone)]
struct FamilyState {
    present: Option<Present>,
    /// The part name a creation of this family uses.
    new_part: PackURI,
    /// Why removing the present part would leave a relationship dangling.
    shared_by: Option<String>,
}

/// The captured document properties of one package.
///
/// A snapshot is immutable: it holds the strictly validated graph and the
/// exact part bytes of the three families, the decoded core and custom
/// values, and decodes the application part when it is asked for (so a
/// malformed `app.xml` fails only the operations that read it).
#[derive(Debug, Clone)]
pub struct Snapshot {
    host: Host,
    core: FamilyState,
    core_value: Option<Core>,
    core_dialect: Option<Dialect>,
    core_report: ToleranceReport,
    app: FamilyState,
    custom: FamilyState,
    custom_value: custom::Props,
    /// The package-level relationship IDs in use.
    relationship_ids: BTreeSet<String>,
}

impl Snapshot {
    /// Capture the document properties of an owning package.
    ///
    /// `host` decides which reserved custom properties belong to the
    /// document ([`custom::Props::read_for`]).
    ///
    /// # Errors
    ///
    /// Returns the strict graph errors of each family (an orphaned part, a
    /// second or part-owned relationship, an external or dangling target, a
    /// target of another content type), and the read errors of the core and
    /// custom parts.
    pub fn read(package: &OpcPackage, host: Host) -> Result<Self> {
        let core_graph = graph::inspect(package).map_err(source_error)?;
        let app_graph = app::inspect(package).map_err(source_error)?;
        let custom_graph = custom::inspect_graph(package).map_err(source_error)?;
        let bytes = |part: &PackURI| -> Result<Arc<Vec<u8>>> {
            Ok(package
                .get_part(part)
                .map_err(|error| source_error(crate::Error::Opc(error)))?
                .blob_arc())
        };
        Self::capture(
            package,
            host,
            Graphs {
                core: (
                    core_graph.part,
                    core_graph.relationship_id,
                    core_graph.dialect,
                ),
                app: (app_graph.part, app_graph.relationship_id),
                custom: (custom_graph.part_name, custom_graph.relationship_id),
            },
            bytes,
        )
    }

    /// Capture only the core-properties family of an owning package; the
    /// application and custom families are not read (record 0377a).
    ///
    /// Only [`publish_core_metadata`] uses such a snapshot, and only for a
    /// changed core patch, which reads and writes nothing of the other two
    /// families: their states here are placeholders, never published.
    fn read_core_only(package: &OpcPackage, host: Host) -> Result<Self> {
        let core_graph = graph::inspect(package).map_err(source_error)?;
        let bytes = |part: &PackURI| -> Result<Arc<Vec<u8>>> {
            Ok(package
                .get_part(part)
                .map_err(|error| source_error(crate::Error::Opc(error)))?
                .blob_arc())
        };
        let core = family_state(
            package,
            Family::Core,
            core_graph.part,
            core_graph.relationship_id,
            &bytes,
        )
        .map_err(source_local_error)?;
        let core_dialect = core_graph.dialect;
        let mut core_report = ToleranceReport::default();
        let core_value = core
            .present
            .as_ref()
            .map(|present| {
                read::decode_selected(
                    &present.bytes,
                    present.part.as_str(),
                    core_dialect,
                    &mut core_report,
                )
            })
            .transpose()
            .map_err(source_error)?;
        // The placeholders are never staged or published; they reuse the
        // core state's part name instead of parsing two more.
        let unread = |new_part: PackURI| FamilyState {
            present: None,
            new_part,
            shared_by: None,
        };
        let (app, custom) = (unread(core.new_part.clone()), unread(core.new_part.clone()));
        let relationship_ids = package
            .package_relationships()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        Ok(Self {
            host,
            core,
            core_value,
            core_dialect,
            core_report,
            app,
            custom,
            custom_value: custom::Props::new(),
            relationship_ids,
        })
    }

    /// Capture the document properties of a source-backed package.
    ///
    /// Only the three relationship-selected parts are read, with the
    /// package's source-freshness and execution checks around them.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::read`], and a source-change or cancellation
    /// error.
    pub fn read_source_backed(package: &SourceBackedPackage, host: Host) -> Result<Self> {
        package.check_execution()?;
        let core_graph = graph::inspect_source_strict(package).map_err(source_error)?;
        let app_graph = app::inspect_source(package).map_err(source_error)?;
        let custom_graph = custom::inspect_source_graph(package).map_err(source_error)?;
        let bytes = |part: &PackURI| -> Result<Arc<Vec<u8>>> {
            let data = package
                .part(part)
                .map_err(|error| source_error(crate::Error::Opc(error)))?
                .data()
                .map_err(|error| source_error(crate::Error::Opc(error)))?;
            package.check_execution()?;
            Ok(Arc::new(data.as_bytes().to_vec()))
        };
        let snapshot = Self::capture(
            package,
            host,
            Graphs {
                core: (
                    core_graph.part,
                    core_graph.relationship_id,
                    core_graph.dialect,
                ),
                app: (app_graph.part, app_graph.relationship_id),
                custom: (custom_graph.part_name, custom_graph.relationship_id),
            },
            bytes,
        )?;
        package.source_version()?;
        Ok(snapshot)
    }

    fn capture<P: PackageMetadata>(
        package: &P,
        host: Host,
        graphs: Graphs,
        bytes: impl Fn(&PackURI) -> Result<Arc<Vec<u8>>>,
    ) -> Result<Self> {
        let Graphs {
            core: (core_part, core_id, core_dialect),
            app: (app_part, app_id),
            custom: (custom_part, custom_id),
        } = graphs;
        let core = family_state(package, Family::Core, core_part, core_id, &bytes)
            .map_err(source_local_error)?;
        let app = family_state(package, Family::App, app_part, app_id, &bytes)
            .map_err(source_local_error)?;
        let custom = family_state(package, Family::Custom, custom_part, custom_id, &bytes)
            .map_err(source_local_error)?;
        let mut core_report = ToleranceReport::default();
        let core_value = core
            .present
            .as_ref()
            .map(|present| {
                read::decode_selected(
                    &present.bytes,
                    present.part.as_str(),
                    core_dialect,
                    &mut core_report,
                )
            })
            .transpose()
            .map_err(source_error)?;
        let custom_value = match &custom.present {
            Some(present) => custom::Props::decode_for(&present.bytes, &present.part, host)
                .map_err(source_error)?,
            None => custom::Props::new(),
        };
        let relationship_ids = package
            .package_relationships()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        Ok(Self {
            host,
            core,
            core_value,
            core_dialect,
            core_report,
            app,
            custom,
            custom_value,
            relationship_ids,
        })
    }

    /// The host the custom properties were read for.
    #[must_use]
    pub const fn host(&self) -> Host {
        self.host
    }

    /// The semantic metadata: the core fields, then the application fields.
    ///
    /// # Errors
    ///
    /// Returns the strict decoding errors of the application part.
    pub fn metadata(&self) -> Result<Metadata> {
        let mut metadata = self
            .core_value
            .as_ref()
            .map(Core::metadata)
            .unwrap_or_default();
        if let Some(app) = self.app_props()? {
            app.project_into(&mut metadata);
        }
        Ok(metadata)
    }

    /// Whether the package has a core-properties part.
    #[must_use]
    pub fn has_core(&self) -> bool {
        self.core.present.is_some()
    }

    /// The repairs reading the core-properties part applied (see
    /// [`crate::leniency`]). A core edit of a part that needed repairs is
    /// refused, because the writer never rewrites content it cannot carry.
    #[must_use]
    pub fn core_tolerance_report(&self) -> &ToleranceReport {
        &self.core_report
    }

    /// The application properties, decoded now; `None` when the package has
    /// no application-properties part.
    ///
    /// # Errors
    ///
    /// Returns the strict decoding errors of [`super::app::read`].
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        self.app
            .present
            .as_ref()
            .map(|present| app::decode(&present.bytes).map_err(source_error))
            .transpose()
    }

    /// The custom properties; empty when the package has no
    /// custom-properties part.
    #[must_use]
    pub fn custom_props(&self) -> &custom::Props {
        &self.custom_value
    }

    /// The part of `family`, when present.
    #[must_use]
    pub fn part_name(&self, family: Family) -> Option<&PackURI> {
        self.family(family)
            .present
            .as_ref()
            .map(|present| &present.part)
    }

    /// The package-level relationship selecting `family`, when present.
    #[doc(hidden)]
    #[must_use]
    pub fn relationship_id_value(&self, family: Family) -> Option<&str> {
        self.family(family)
            .present
            .as_ref()
            .map(|present| present.relationship_id.as_str())
    }

    /// The part name selected if `family` is created by a later edit.
    #[must_use]
    pub fn new_part_name(&self, family: Family) -> &PackURI {
        &self.family(family).new_part
    }

    /// The exact bytes of the part of `family`, when present.
    #[must_use]
    pub fn part_bytes(&self, family: Family) -> Option<&[u8]> {
        self.family(family)
            .present
            .as_ref()
            .map(|present| present.bytes.as_slice())
    }

    /// Start a detached [`Plan`] of a change of this snapshot.
    #[must_use]
    pub fn plan(&self) -> Plan {
        self.plan_with_options(EditOptions::default())
    }

    /// Start a detached plan with explicit edit options.
    ///
    /// The limit is retained by the plan, so [`Plan::finish`] performs its
    /// exact pre-capture admission before it materializes any replacement
    /// property payload.
    #[must_use]
    pub fn plan_with_options(&self, options: EditOptions) -> Plan {
        Plan {
            base: self.clone(),
            metadata: None,
            clear_core: false,
            force_core: false,
            app: None,
            custom: None,
            remove_custom: false,
            options,
        }
    }

    /// Read the snapshot that would result from replaying `patch` over a
    /// source-backed package.
    ///
    /// This keeps source-backed replay deferred: only the three property
    /// parts named by the patch are loaded, while every other package member
    /// remains in the retained source. The patch is checked against the
    /// source before any replacement is admitted.
    pub fn read_source_backed_after(
        package: &SourceBackedPackage,
        host: Host,
        patch: &Patch,
    ) -> Result<Self> {
        patch.check_source_backed(package)?;
        let mut snapshot = Self::read_source_backed(package, host)?;
        for change in &patch.changes {
            snapshot.apply_change(change)?;
        }
        package.source_version()?;
        Ok(snapshot)
    }

    fn apply_change(&mut self, change: &Change) -> Result<()> {
        let state = self.family_mut(change.family);
        let present = match (&change.before, &change.after) {
            (Some(_), Some(after)) => Some(Present {
                part: change.part.clone(),
                relationship_id: change.relationship_id.clone(),
                relationship_type: change.relationship_type.clone(),
                relationship_target: change.relationship_target.clone(),
                bytes: Arc::clone(after),
            }),
            (Some(_), None) => None,
            (None, Some(after)) => Some(Present {
                part: change.part.clone(),
                relationship_id: change.relationship_id.clone(),
                relationship_type: change.relationship_type.clone(),
                relationship_target: change.relationship_target.clone(),
                bytes: Arc::clone(after),
            }),
            (None, None) => return Ok(()),
        };
        if change.before.is_none() {
            state.new_part = change.part.clone();
        }
        state.present = present;
        state.shared_by = None;
        match change.before.is_some() {
            true if change.after.is_none() => {
                self.relationship_ids.remove(&change.relationship_id);
            },
            false if change.after.is_some() => {
                self.relationship_ids.insert(change.relationship_id.clone());
            },
            _ => {},
        }

        match change.family {
            Family::Core => {
                if let Some(after) = change.after.as_deref() {
                    let dialect =
                        if change.relationship_type == super::STRICT_CORE_PROPERTIES_RELATIONSHIP {
                            Dialect::Strict
                        } else {
                            Dialect::Transitional
                        };
                    let mut report = ToleranceReport::default();
                    self.core_value = Some(
                        read::decode_selected(
                            after,
                            change.part.as_str(),
                            Some(dialect),
                            &mut report,
                        )
                        .map_err(patch_output_error)?,
                    );
                    self.core_dialect = Some(dialect);
                    self.core_report = report;
                } else {
                    self.core_value = None;
                    self.core_dialect = None;
                    self.core_report = ToleranceReport::default();
                }
            },
            Family::Custom => {
                self.custom_value = match change.after.as_deref() {
                    Some(after) => custom::Props::decode_for(after, &change.part, self.host)
                        .map_err(patch_output_error)?,
                    None => custom::Props::new(),
                };
            },
            Family::App => {},
        }
        Ok(())
    }

    fn family(&self, family: Family) -> &FamilyState {
        match family {
            Family::Core => &self.core,
            Family::App => &self.app,
            Family::Custom => &self.custom,
        }
    }

    fn family_mut(&mut self, family: Family) -> &mut FamilyState {
        match family {
            Family::Core => &mut self.core,
            Family::App => &mut self.app,
            Family::Custom => &mut self.custom,
        }
    }
}

/// The three graphs a capture reads: part, relationship ID (and the core
/// dialect).
struct Graphs {
    core: (Option<PackURI>, Option<String>, Option<Dialect>),
    app: (Option<PackURI>, Option<String>),
    custom: (Option<PackURI>, Option<String>),
}

fn family_state<P: PackageMetadata>(
    package: &P,
    family: Family,
    part: Option<PackURI>,
    relationship_id: Option<String>,
    bytes: &impl Fn(&PackURI) -> Result<Arc<Vec<u8>>>,
) -> Result<FamilyState> {
    let new_part = free_part_name(
        package,
        family.stem(),
        crate::LimitKind::DocumentPropertyPartNames,
    );
    let present = match (part, relationship_id) {
        (Some(part), Some(relationship_id)) => {
            let relationship = package
                .package_relationships()
                .get(&relationship_id)
                .ok_or_else(|| {
                    Error::Invalid(format!(
                        "{} relationship '{relationship_id}' disappeared",
                        family.label()
                    ))
                })?;
            Some(Present {
                relationship_type: relationship.reltype().to_owned(),
                relationship_target: relationship.target_ref().to_owned(),
                bytes: bytes(&part)?,
                part,
                relationship_id,
            })
        },
        (None, None) => None,
        _ => {
            return Err(Error::Invalid(format!(
                "{} graph has inconsistent part and relationship state",
                family.label()
            )));
        },
    };
    let shared_by = match &present {
        Some(present) => other_inbound(
            package,
            &present.part,
            &present.relationship_id,
            family.label(),
        )?,
        None => None,
    };
    let new_part = match (&present, new_part) {
        // A present family is never created; a name collision is moot.
        (Some(present), Err(_)) => present.part.clone(),
        (_, new_part) => new_part?,
    };
    Ok(FamilyState {
        present,
        new_part,
        shared_by,
    })
}

/// Capture the complete document-properties graph for an exact no-op patch.
///
/// Changed patches carry before bytes only for the families they replace.
/// An empty patch has no changed family to authenticate, so it records every
/// property family instead. The relationship ID retained for an absent family
/// is diagnostic metadata; replay treats it as a wildcard and authenticates
/// the actual absent-family state on the recipient.
fn snapshot_scope(snapshot: &Snapshot) -> Result<Box<[Change]>> {
    let mut relationship_ids = snapshot.relationship_ids.clone();
    [Family::Core, Family::App, Family::Custom]
        .into_iter()
        .map(|family| {
            let state = snapshot.family(family);
            let (relationship_id, part, relationship_type, relationship_target, before, after) =
                match &state.present {
                    Some(present) => (
                        retained_string(&present.relationship_id),
                        exact_part(&present.part)?,
                        retained_string(&present.relationship_type),
                        retained_string(&present.relationship_target),
                        Some(Arc::clone(&present.bytes)),
                        Some(Arc::clone(&present.bytes)),
                    ),
                    None => (
                        // Keep distinct diagnostic placeholders. Replay
                        // authenticates absence without requiring these IDs
                        // to remain free on the recipient.
                        retained_string(&next_relationship_id(&relationship_ids)?),
                        exact_part(&state.new_part)?,
                        retained_string(match family {
                            Family::Core => rt::CORE_PROPERTIES,
                            Family::App => rt::EXTENDED_PROPERTIES,
                            Family::Custom => rt::CUSTOM_PROPERTIES,
                        }),
                        retained_string(state.new_part.as_str().trim_start_matches('/')),
                        None,
                        None,
                    ),
                };
            if state.present.is_none() {
                relationship_ids.insert(relationship_id.clone());
            }
            Ok(Change {
                family,
                part,
                relationship_id,
                relationship_type,
                relationship_target,
                before,
                after,
                free_relationship_id: state.present.is_none(),
            })
        })
        .collect()
}

/// A detached plan of a change of a [`Snapshot`]'s document properties.
///
/// A plan stages like an [`Edit`] but holds no package: [`Plan::finish`]
/// returns the exact [`Patch`] (with the read view after it and the
/// [`Diagnostics`]) for an owner that publishes over a retained source
/// rather than an owning [`OpcPackage`] (the XLSB source-backed workbook,
/// through [`Patch::check_source_backed`] and
/// [`Patch::prepare_source_backed`]). An owning package uses [`Edit`],
/// whose commit publishes the new package.
///
/// Each family is either untouched (its part is never read or written by
/// the commit) or staged with a complete target value. Staging a value
/// equal to the current one changes nothing.
#[derive(Debug, Clone)]
pub struct Plan {
    base: Snapshot,
    metadata: Option<Metadata>,
    clear_core: bool,
    /// Whether an explicit metadata set must create a core part even when
    /// the requested value is the empty semantic projection.
    force_core: bool,
    app: Option<Option<app::Props>>,
    custom: Option<custom::Props>,
    /// Whether an explicit custom-family removal was requested. This is
    /// distinct from setting an equal empty collection, which is a semantic
    /// no-op and preserves a present raw part.
    remove_custom: bool,
    options: EditOptions,
}

/// Retained endpoint accounting supplied by a host's immutable package.
///
/// The common planner combines this owner census with its prepared family
/// lengths before it allocates replacement payloads. It carries no package,
/// snapshot or publication handle and is therefore safe to retain in
/// [`EditOptions`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[doc(hidden)]
pub struct OwnerAdmission {
    source_bytes: usize,
    source_archive_bytes: usize,
    families: [FamilyAdmission; 3],
}

impl OwnerAdmission {
    /// Creates a host endpoint census in `Family::Core`, `Family::App`, and
    /// `Family::Custom` order.
    #[doc(hidden)]
    pub const fn new(
        source_bytes: usize,
        source_archive_bytes: usize,
        families: [FamilyAdmission; 3],
    ) -> Self {
        Self {
            source_bytes,
            source_archive_bytes,
            families,
        }
    }
}

/// Fixed retained bytes for one property family in a host owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[doc(hidden)]
pub struct FamilyAdmission {
    present_part_bytes: Option<usize>,
    present_relationship_bytes: Option<usize>,
    created_part_bytes: usize,
    created_relationship_base_bytes: usize,
}

impl FamilyAdmission {
    /// Creates one family census. Part values exclude their payload; the
    /// planner adds the exact before/after payload capacities itself.
    #[doc(hidden)]
    pub const fn new(
        present_part_bytes: Option<usize>,
        present_relationship_bytes: Option<usize>,
        created_part_bytes: usize,
        created_relationship_base_bytes: usize,
    ) -> Self {
        Self {
            present_part_bytes,
            present_relationship_bytes,
            created_part_bytes,
            created_relationship_base_bytes,
        }
    }
}

impl OwnerAdmission {
    fn family(&self, family: Family) -> &FamilyAdmission {
        &self.families[family_index(family)]
    }

    fn check(&self, prepared: &PreparedPlan, maximum: Option<usize>) -> Result<()> {
        let Some(maximum) = maximum else {
            return Ok(());
        };
        let mut target = if prepared.planned_changes.iter().all(Option::is_none) {
            self.source_bytes
        } else {
            self.source_bytes
                .checked_sub(self.source_archive_bytes)
                .ok_or(Error::Internal(
                    "metadata owner admission endpoint underflow",
                ))?
        };
        for change in prepared.planned_changes.iter().flatten() {
            let family = self.family(change.family);
            target = match (change.before_len, change.after_len) {
                (Some(before), Some(after)) => target
                    .checked_sub(before)
                    .and_then(|value| value.checked_add(after))
                    .ok_or(Error::Internal(
                        "metadata owner admission replacement accounting overflow",
                    ))?,
                (Some(before), None) => {
                    let part_bytes = family.present_part_bytes.ok_or(Error::Internal(
                        "metadata owner admission removal part missing",
                    ))?;
                    let relationship_bytes = family.present_relationship_bytes.ok_or(
                        Error::Internal("metadata owner admission removal relationship missing"),
                    )?;
                    let removed = part_bytes
                        .checked_add(before)
                        .and_then(|value| value.checked_add(relationship_bytes))
                        .ok_or(Error::Internal(
                            "metadata owner admission removal accounting overflow",
                        ))?;
                    target.checked_sub(removed).ok_or(Error::Internal(
                        "metadata owner admission endpoint underflow",
                    ))?
                },
                (None, Some(after)) => target
                    .checked_add(family.created_part_bytes)
                    .and_then(|value| value.checked_add(after))
                    .and_then(|value| value.checked_add(family.created_relationship_base_bytes))
                    .and_then(|value| value.checked_add(change.relationship_id.len()))
                    .ok_or(Error::Internal(
                        "metadata owner admission creation accounting overflow",
                    ))?,
                (None, None) => {
                    return Err(Error::Internal(
                        "metadata owner admission carried an empty topology change",
                    ));
                },
            };
        }
        let observed = self
            .source_bytes
            .checked_add(target)
            .and_then(|value| value.checked_add(prepared.planned_retained_bytes))
            .ok_or(Error::Internal(
                "metadata owner admission retained-byte accounting overflow",
            ))?;
        if observed > maximum {
            return Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed,
                maximum,
            });
        }
        Ok(())
    }
}

impl Plan {
    /// The snapshot this plan started from.
    #[must_use]
    pub const fn base(&self) -> &Snapshot {
        &self.base
    }

    /// The staged metadata: the staged value, or the snapshot's.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of [`Snapshot::metadata`] when no value is
    /// staged.
    pub fn metadata(&self) -> Result<Metadata> {
        let mut metadata = match &self.metadata {
            Some(staged) => staged.clone(),
            None => self.base.metadata()?,
        };
        if self.clear_core {
            clear_core_metadata(&mut metadata);
        }
        Ok(metadata)
    }

    /// Stage the complete metadata. Fields equal to the snapshot's change
    /// nothing; a changed field lands in the family that stores it.
    ///
    /// Setting a value equal to the snapshot's is an exact no-op even when
    /// the source has no core part: these hosts read an absent part as
    /// `Metadata::default()`, so creating one would change bytes without
    /// changing the value (0351 §2.3 *No-op*). A host whose API tells an
    /// absent core part from an empty one records explicit creation through
    /// [`Self::set_core_metadata`].
    pub fn set_metadata(&mut self, metadata: Metadata) -> &mut Self {
        self.metadata = Some(metadata);
        self.clear_core = false;
        self.force_core = false;
        self
    }

    /// Stage only the core-properties fields, retaining the staged or source
    /// application projection unchanged.
    ///
    /// This is the compatibility boundary for hosts whose core-properties
    /// family is edited separately from application and custom properties.
    /// Non-core fields are rejected instead of being silently discarded.
    pub fn set_core_metadata(&mut self, metadata: Metadata) -> Result<&mut Self> {
        super::ensure_core_only_transaction(&metadata)?;
        let mut target = self.metadata()?;
        merge_core_metadata(&mut target, &metadata);
        self.metadata = Some(target);
        self.clear_core = false;
        self.force_core = self.base.core.present.is_none();
        Ok(self)
    }

    /// Stage removal of the core-properties part.
    pub fn clear_metadata(&mut self) -> &mut Self {
        self.metadata = None;
        self.clear_core = true;
        self.force_core = false;
        self
    }

    /// Stage a change of the metadata, starting from the staged value.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of [`Snapshot::metadata`].
    pub fn update_metadata(&mut self, update: impl FnOnce(&mut Metadata)) -> Result<&mut Self> {
        let mut metadata = self.metadata()?;
        let before = metadata.clone();
        update(&mut metadata);
        if self.clear_core
            && before
                .changed_fields(&metadata)
                .any(|field| CORE_FIELDS.contains(&field))
        {
            self.clear_core = false;
        }
        self.metadata = Some(metadata);
        Ok(self)
    }

    /// The staged application properties: the staged value, or the
    /// snapshot's.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of [`Snapshot::app_props`] when the family
    /// is not staged.
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        match &self.app {
            Some(staged) => Ok(staged.clone()),
            None => self.base.app_props(),
        }
    }

    /// Stage the complete application properties; `None` removes the part.
    ///
    /// Metadata staged with [`Self::set_metadata`] is merged on top of this
    /// value at commit.
    pub fn set_app_props(&mut self, props: Option<app::Props>) -> &mut Self {
        self.app = Some(props);
        self
    }

    /// The staged custom properties: the staged value, or the snapshot's.
    #[must_use]
    pub fn custom_props(&self) -> &custom::Props {
        self.custom.as_ref().unwrap_or(&self.base.custom_value)
    }

    /// Stage the complete custom properties; an empty collection removes the
    /// part when it changes the semantic collection. Setting an equal empty
    /// collection is an exact no-op, preserving a present raw part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> &mut Self {
        self.custom = Some(props);
        self.remove_custom = false;
        self
    }

    /// Stage explicit removal of the custom-properties part and relationship.
    ///
    /// This is distinct from setting an equal empty collection: it removes a
    /// present raw part even when its tolerant semantic projection is empty.
    /// An absent graph remains an exact no-op.
    pub fn remove_custom_props(&mut self) -> &mut Self {
        self.custom = Some(custom::Props::new());
        self.remove_custom = true;
        self
    }

    /// Prepare the exact family targets and retained-byte charge.
    ///
    /// This is deliberately a single pass over the decoded source values.
    /// The resulting prepared targets and lengths are consumed by
    /// [`PreparedPlan::finish`], so admission does not decode, splice-plan or
    /// otherwise rescan the same property families a second time.
    fn prepare(self) -> Result<PreparedPlan> {
        let Self {
            base,
            metadata,
            clear_core,
            force_core,
            app: staged_app,
            custom,
            remove_custom,
            options: _,
        } = self;
        let (core_metadata, app_target, changed_fields) = if clear_core {
            match metadata {
                Some(target) => {
                    let before = base.metadata()?;
                    let changed_fields: Vec<_> = before
                        .changed_fields(&target)
                        .filter(|field| !CORE_FIELDS.contains(field))
                        .collect();
                    if changed_fields
                        .iter()
                        .any(|field| !app::METADATA_FIELDS.contains(field))
                    {
                        return Err(unsupported_property());
                    }
                    let app_changed = changed_fields
                        .iter()
                        .any(|field| app::METADATA_FIELDS.contains(field));
                    let app_base = match &staged_app {
                        Some(staged) => staged.clone(),
                        None => base.app_props()?,
                    };
                    let app_target = if app_changed {
                        Some(Some(app::Props::merged(app_base.as_ref(), &target)?))
                    } else {
                        staged_app
                    };
                    (Some(None), app_target, changed_fields)
                },
                None => (Some(None), staged_app, Vec::new()),
            }
        } else {
            match metadata {
                Some(target) => {
                    let before = base.metadata()?;
                    let changed_fields: Vec<_> = before.changed_fields(&target).collect();
                    if changed_fields.iter().any(|field| {
                        !CORE_FIELDS.contains(field) && !app::METADATA_FIELDS.contains(field)
                    }) {
                        return Err(unsupported_property());
                    }
                    let core_changed = changed_fields
                        .iter()
                        .any(|field| CORE_FIELDS.contains(field));
                    let core_requested = core_changed || force_core;
                    let app_changed = changed_fields
                        .iter()
                        .any(|field| app::METADATA_FIELDS.contains(field));
                    let app_base = match &staged_app {
                        Some(staged) => staged.clone(),
                        None => base.app_props()?,
                    };
                    let app_target = if app_changed {
                        Some(Some(app::Props::merged(app_base.as_ref(), &target)?))
                    } else {
                        staged_app
                    };
                    (
                        core_requested.then_some(Some(target)),
                        app_target,
                        changed_fields,
                    )
                },
                None => (None, staged_app, Vec::new()),
            }
        };

        let core = match core_metadata {
            Some(Some(target)) => {
                if base.core.present.is_some() && !base.core_report.is_clean() {
                    return Err(unsupported_core_content());
                }
                let (current, dialect) = match &base.core.present {
                    Some(present) => {
                        let xml = std::str::from_utf8(&present.bytes).map_err(|error| {
                            Error::Package(crate::Error::Invalid(format!(
                                "invalid UTF-8 in core properties: {error}"
                            )))
                        })?;
                        let (current, dialect) = read::decode(xml).map_err(source_error)?;
                        if Some(dialect) != base.core_dialect {
                            return Err(Error::Package(crate::Error::Invalid(
                                "core-properties XML namespace does not match its relationship dialect"
                                    .to_owned(),
                            )));
                        }
                        (current, dialect)
                    },
                    None => (Core::default(), Dialect::Transitional),
                };
                super::write::validate_target(&target).map_err(Error::from)?;
                let merged = Core::merged(base.core.present.as_ref().map(|_| &current), &target)?;
                (force_core || merged != current)
                    .then(|| {
                        super::write::encoded_len(&merged, dialect)
                            .map_err(core_generated_error)
                            .map(|after_len| PreparedCore {
                                target: Some(merged),
                                dialect,
                                after_len: Some(after_len),
                            })
                    })
                    .transpose()?
            },
            Some(None) if base.core.present.is_some() => Some(PreparedCore {
                target: None,
                dialect: base.core_dialect.unwrap_or(Dialect::Transitional),
                after_len: None,
            }),
            _ => None,
        };

        let app = if let Some(target) = app_target {
            let current = base.app_props()?;
            if target == current {
                None
            } else {
                let (after_len, plan) = match (&base.app.present, &target) {
                    (Some(present), Some(props)) => {
                        app::validate_target(props).map_err(Error::from)?;
                        let plan =
                            app::plan_splice(&present.bytes, props).map_err(app_plan_error)?;
                        (plan.as_ref().map(|plan| plan.output_len()), plan)
                    },
                    (None, Some(props)) => {
                        app::validate_target(props).map_err(Error::from)?;
                        (
                            Some(app::encode_new_len(props).map_err(app_authored_error)?),
                            None,
                        )
                    },
                    (_, None) => (None, None),
                };
                Some(PreparedApp {
                    target,
                    after_len,
                    plan,
                })
            }
        } else {
            None
        };

        let custom = custom
            .filter(|target| {
                (remove_custom && base.custom.present.is_some())
                    || (!remove_custom && !target.same_values(&base.custom_value))
            })
            .map(|target| {
                if !remove_custom {
                    target.validate_for(base.host)?;
                }
                if !remove_custom && !base.custom_value.tolerance_report().is_clean() {
                    return Err(unsupported_custom_content());
                }
                let after_len = (!remove_custom && !target.is_empty())
                    .then(|| custom::encoded_len(&target).map_err(custom_generated_error))
                    .transpose()?;
                Ok(PreparedCustom { target, after_len })
            })
            .transpose()?;

        let mut relationship_ids = base.relationship_ids.clone();
        let mut retained_changes = 0usize;
        let mut changed = false;
        let mut planned_changes = std::array::from_fn(|_| None);
        if let Some(core) = &core {
            let (bytes, change) =
                planned_change_bytes(&base, Family::Core, core.after_len, &mut relationship_ids)?;
            retained_changes = retained_changes.checked_add(bytes).ok_or(Error::Internal(
                "metadata prepared retained-byte accounting overflow",
            ))?;
            planned_changes[family_index(Family::Core)] = Some(change);
            changed = true;
        }
        if let Some(app) = &app {
            let (bytes, change) =
                planned_change_bytes(&base, Family::App, app.after_len, &mut relationship_ids)?;
            retained_changes = retained_changes.checked_add(bytes).ok_or(Error::Internal(
                "metadata prepared retained-byte accounting overflow",
            ))?;
            planned_changes[family_index(Family::App)] = Some(change);
            changed = true;
        }
        if let Some(custom) = &custom {
            let (bytes, change) = planned_change_bytes(
                &base,
                Family::Custom,
                custom.after_len,
                &mut relationship_ids,
            )?;
            retained_changes = retained_changes.checked_add(bytes).ok_or(Error::Internal(
                "metadata prepared retained-byte accounting overflow",
            ))?;
            planned_changes[family_index(Family::Custom)] = Some(change);
            changed = true;
        }
        let planned_retained_bytes = if changed {
            size_of::<Patch>()
                .checked_add(retained_changes)
                .ok_or(Error::Internal(
                    "metadata prepared retained-byte accounting overflow",
                ))?
        } else {
            let scope = planned_scope_bytes(&base)?;
            let scope = scope.checked_mul(2).ok_or(Error::Internal(
                "metadata prepared scope accounting overflow",
            ))?;
            size_of::<Patch>()
                .checked_add(scope)
                .ok_or(Error::Internal(
                    "metadata prepared retained-byte accounting overflow",
                ))?
        };
        Ok(PreparedPlan {
            base,
            core,
            app,
            custom,
            changed_fields: changed_fields.into_boxed_slice(),
            planned_changes,
            planned_retained_bytes,
        })
    }

    /// Compute the exact patch of the staged values: the read view after
    /// it, the patch and what changed. Nothing is published; the owner
    /// checks and publishes the patch over its source.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a changed field no OOXML family stores
    /// or content that this lossless route cannot carry; [`Error::Invalid`]
    /// for a core or custom part that holds content its writer cannot carry
    /// or a value outside its lexical space; [`Error::Internal`] for a
    /// violated prepared-plan or admission invariant; [`Error::Limit`] for
    /// an exceeded bound; and [`Error::SourceChanged`] for source-backed
    /// freshness failure.
    pub fn finish(self) -> Result<(Snapshot, Patch, Diagnostics)> {
        let maximum = self.options.max_retained_bytes_value();
        let owner = self.options.owner;
        let prepared = self.prepare()?;
        prepared.finish_with_limit(maximum, owner)
    }
}

fn retained_string(value: &str) -> String {
    let mut retained = String::with_capacity(value.len());
    retained.push_str(value);
    retained.shrink_to_fit();
    retained
}

fn exact_part(part: &PackURI) -> Result<PackURI> {
    PackURI::new(retained_string(part.as_str()))
        .map_err(|_| Error::Internal("retained metadata part URI became invalid"))
}

#[derive(Debug, Clone)]
struct PreparedChange {
    family: Family,
    part: PackURI,
    relationship_id: String,
    relationship_type: String,
    relationship_target: String,
    before_len: Option<usize>,
    after_len: Option<usize>,
}

#[derive(Debug)]
struct PreparedPlan {
    base: Snapshot,
    core: Option<PreparedCore>,
    app: Option<PreparedApp>,
    custom: Option<PreparedCustom>,
    changed_fields: Box<[Field]>,
    planned_changes: [Option<PreparedChange>; 3],
    planned_retained_bytes: usize,
}

#[derive(Debug)]
struct PreparedCore {
    target: Option<Core>,
    dialect: Dialect,
    after_len: Option<usize>,
}

#[derive(Debug)]
struct PreparedApp {
    target: Option<app::Props>,
    after_len: Option<usize>,
    plan: Option<app::SplicePlan>,
}

#[derive(Debug)]
struct PreparedCustom {
    target: custom::Props,
    after_len: Option<usize>,
}

impl PreparedPlan {
    fn check_planned_limit(
        &self,
        maximum: Option<usize>,
        owner: Option<OwnerAdmission>,
    ) -> Result<()> {
        if let Some(maximum) = maximum
            && self.planned_retained_bytes > maximum
        {
            return Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed: self.planned_retained_bytes,
                maximum,
            });
        }
        if let Some(owner) = owner {
            owner.check(self, maximum)?;
        }
        Ok(())
    }

    fn finish_with_limit(
        self,
        maximum: Option<usize>,
        owner: Option<OwnerAdmission>,
    ) -> Result<(Snapshot, Patch, Diagnostics)> {
        self.check_planned_limit(maximum, owner)?;
        let result = self.finish_prepared();
        if let Ok((_, patch, _)) = &result
            && let Some(maximum) = maximum
            && patch.retained_bytes() > maximum
        {
            return Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed: patch.retained_bytes(),
                maximum,
            });
        }
        result
    }

    fn finish_prepared(self) -> Result<(Snapshot, Patch, Diagnostics)> {
        let Self {
            base,
            core,
            app,
            custom,
            mut changed_fields,
            mut planned_changes,
            ..
        } = self;
        let mut after = base.clone();
        let mut changes = Vec::new();
        if let Some(core) = core {
            let planned = take_planned_change(&mut planned_changes, Family::Core)?;
            let (after_bytes, merged) = match (core.target, core.after_len) {
                (Some(merged), Some(after_len)) => (
                    Some(
                        super::write::encode_preflighted(&merged, core.dialect, after_len)
                            .map_err(core_generated_error)?
                            .into_bytes(),
                    ),
                    Some(merged),
                ),
                (None, None) => (None, None),
                _ => {
                    return Err(Error::Internal(
                        "core-properties prepared target and length disagree",
                    ));
                },
            };
            if let Some(change) = stage_prepared_change(&mut after, planned, after_bytes)? {
                changes.push(change);
                if let Some(merged) = merged {
                    after.core_dialect = after.core_dialect.or(Some(core.dialect));
                    after.core_value = Some(merged);
                    after.core_report = ToleranceReport::default();
                } else {
                    after.core_dialect = None;
                    after.core_value = None;
                    after.core_report = ToleranceReport::default();
                }
            }
        }

        if let Some(app) = app {
            let planned = take_planned_change(&mut planned_changes, Family::App)?;
            let after_bytes = match (&base.app.present, &app.target) {
                (Some(present), Some(props)) => Some(
                    app::splice_planned(
                        &present.bytes,
                        props,
                        app.plan.ok_or(Error::Internal(
                            "missing prepared application-properties layout",
                        ))?,
                    )
                    .map_err(app_prepared_error)?,
                ),
                (None, Some(props)) => Some(
                    app::encode_new_preflighted(
                        props,
                        app.after_len.ok_or(Error::Internal(
                            "missing prepared application-properties length",
                        ))?,
                    )
                    .map_err(app_authored_error)?,
                ),
                (_, None) => None,
            };
            if let Some(change) = stage_prepared_change(&mut after, planned, after_bytes)? {
                changes.push(change);
            }
        }

        if let Some(custom) = custom {
            let planned = take_planned_change(&mut planned_changes, Family::Custom)?;
            let after_bytes = if custom.target.is_empty() {
                None
            } else {
                Some(
                    custom::encode_preflighted(
                        &custom.target,
                        custom
                            .after_len
                            .ok_or(Error::Internal("missing prepared custom-properties length"))?,
                    )
                    .map_err(custom_generated_error)?,
                )
            };
            if let Some(change) = stage_prepared_change(&mut after, planned, after_bytes)? {
                changes.push(change);
                after.custom_value = custom.target;
            }
        }

        let families = changes.iter().map(|change| change.family).collect();
        if changes.is_empty() {
            changed_fields = Box::new([]);
        }
        // An empty patch still has to authenticate the package it came from.
        // Keep the complete document-properties scope separate from the
        // changed-family list so `is_noop()` and diagnostics remain truthful
        // while replay can reject an equal-looking patch on another owner.
        let scope = if changes.is_empty() {
            Some(snapshot_scope(&base)?)
        } else {
            None
        };
        Ok((
            after,
            Patch {
                changes: changes.into_boxed_slice(),
                after_scope: scope.clone(),
                scope,
            },
            Diagnostics {
                families,
                fields: changed_fields,
            },
        ))
    }
}

fn take_planned_change(
    planned_changes: &mut [Option<PreparedChange>; 3],
    family: Family,
) -> Result<PreparedChange> {
    planned_changes[family_index(family)]
        .take()
        .ok_or(Error::Internal("missing prepared property change"))
}

/// Publish one already-planned family while reusing its topology strings and
/// admitted lengths. The prepared descriptor is consumed into the final patch;
/// no family graph or relationship-ID scan is repeated here.
fn stage_prepared_change(
    after: &mut Snapshot,
    prepared: PreparedChange,
    bytes: Option<Vec<u8>>,
) -> Result<Option<Change>> {
    let family = prepared.family;
    let state = after.family(family).clone();
    let is_removal = prepared.after_len.is_none();
    if let Some(bytes) = bytes.as_ref()
        && Some(bytes.len()) != prepared.after_len
    {
        return Err(Error::Internal("prepared property length changed"));
    }
    match (state.present, prepared.before_len, is_removal, bytes) {
        (Some(present), Some(_), true, None) => {
            if state.shared_by.is_some() {
                return Err(shared_family_removal());
            }
            after.relationship_ids.remove(&prepared.relationship_id);
            let slot = after.family_mut(family);
            slot.new_part = prepared.part.clone();
            slot.present = None;
            Ok(Some(Change {
                family,
                part: prepared.part,
                relationship_id: prepared.relationship_id,
                relationship_type: prepared.relationship_type,
                relationship_target: prepared.relationship_target,
                before: Some(present.bytes),
                after: None,
                free_relationship_id: false,
            }))
        },
        (Some(present), Some(_), false, Some(bytes)) => {
            if bytes.as_slice() == present.bytes.as_slice() {
                return Ok(None);
            }
            let bytes = Arc::new(bytes);
            let slot = after.family_mut(family);
            let Some(slot) = slot.present.as_mut() else {
                return Err(Error::Internal("prepared property part disappeared"));
            };
            slot.bytes = Arc::clone(&bytes);
            let change = Change {
                family,
                part: prepared.part,
                relationship_id: prepared.relationship_id,
                relationship_type: prepared.relationship_type,
                relationship_target: prepared.relationship_target,
                before: Some(present.bytes),
                after: Some(bytes),
                free_relationship_id: false,
            };
            Ok(Some(change))
        },
        (None, None, false, Some(bytes)) => {
            let bytes = Arc::new(bytes);
            let present = Present {
                part: prepared.part.clone(),
                relationship_id: prepared.relationship_id.clone(),
                relationship_type: prepared.relationship_type.clone(),
                relationship_target: prepared.relationship_target.clone(),
                bytes: Arc::clone(&bytes),
            };
            after
                .relationship_ids
                .insert(prepared.relationship_id.clone());
            let slot = after.family_mut(family);
            slot.present = Some(present);
            slot.shared_by = None;
            Ok(Some(Change {
                family,
                part: prepared.part,
                relationship_id: prepared.relationship_id,
                relationship_type: prepared.relationship_type,
                relationship_target: prepared.relationship_target,
                before: None,
                after: Some(bytes),
                free_relationship_id: false,
            }))
        },
        (_, _, _, _) => Err(Error::Internal("prepared property topology changed")),
    }
}

fn planned_change_bytes(
    snapshot: &Snapshot,
    family: Family,
    after_len: Option<usize>,
    relationship_ids: &mut BTreeSet<String>,
) -> Result<(usize, PreparedChange)> {
    let state = snapshot.family(family);
    // Keep every relationship ID from the base snapshot reserved for the
    // entire plan. A family removed earlier in this plan still occupies its
    // ID in the patch's authenticated before graph; letting a later family
    // creation reuse it would make the one-patch check fail (and would make
    // its inverse ambiguous). IDs become free only between independent
    // plans, after the first patch has been published.
    let (part, relationship_id, relationship_type, relationship_target, before_len) =
        match &state.present {
            Some(present) => (
                exact_part(&present.part)?,
                retained_string(&present.relationship_id),
                retained_string(&present.relationship_type),
                retained_string(&present.relationship_target),
                Some(present.bytes.capacity()),
            ),
            None => {
                let relationship_id = retained_string(&next_relationship_id(relationship_ids)?);
                relationship_ids.insert(relationship_id.clone());
                let relationship_target =
                    retained_string(state.new_part.as_str().trim_start_matches('/'));
                let relationship_type = retained_string(match family {
                    Family::Core => rt::CORE_PROPERTIES,
                    Family::App => rt::EXTENDED_PROPERTIES,
                    Family::Custom => rt::CUSTOM_PROPERTIES,
                });
                (
                    exact_part(&state.new_part)?,
                    relationship_id,
                    relationship_type,
                    relationship_target,
                    None,
                )
            },
        };
    let change = PreparedChange {
        family,
        part,
        relationship_id,
        relationship_type,
        relationship_target,
        before_len,
        after_len,
    };
    let bytes = size_of::<Change>()
        .checked_add(change.part.as_str().len())
        .and_then(|value| value.checked_add(change.relationship_id.capacity()))
        .and_then(|value| value.checked_add(change.relationship_type.capacity()))
        .and_then(|value| value.checked_add(change.relationship_target.capacity()))
        .and_then(|value| value.checked_add(change.before_len.unwrap_or(0)))
        .and_then(|value| value.checked_add(change.after_len.unwrap_or(0)))
        .ok_or(Error::Internal(
            "metadata planned-change accounting overflow",
        ))?;
    Ok((bytes, change))
}

fn planned_scope_bytes(snapshot: &Snapshot) -> Result<usize> {
    let mut relationship_ids = snapshot.relationship_ids.clone();
    let mut total = 0usize;
    for family in [Family::Core, Family::App, Family::Custom] {
        let state = snapshot.family(family);
        let bytes = match &state.present {
            Some(present) => size_of::<Change>()
                .checked_add(present.part.as_str().len())
                .and_then(|value| {
                    value.checked_add(retained_string(&present.relationship_id).capacity())
                })
                .and_then(|value| {
                    value.checked_add(retained_string(&present.relationship_type).capacity())
                })
                .and_then(|value| {
                    value.checked_add(retained_string(&present.relationship_target).capacity())
                })
                .and_then(|value| {
                    present
                        .bytes
                        .capacity()
                        .checked_mul(2)
                        .and_then(|bytes| value.checked_add(bytes))
                })
                .ok_or(Error::Internal(
                    "metadata planned-scope accounting overflow",
                ))?,
            None => {
                let relationship_id = retained_string(&next_relationship_id(&relationship_ids)?);
                relationship_ids.insert(relationship_id.clone());
                let relationship_type = retained_string(match family {
                    Family::Core => rt::CORE_PROPERTIES,
                    Family::App => rt::EXTENDED_PROPERTIES,
                    Family::Custom => rt::CUSTOM_PROPERTIES,
                });
                let relationship_target =
                    retained_string(state.new_part.as_str().trim_start_matches('/'));
                size_of::<Change>()
                    .checked_add(state.new_part.as_str().len())
                    .and_then(|value| value.checked_add(relationship_id.capacity()))
                    .and_then(|value| value.checked_add(relationship_type.capacity()))
                    .and_then(|value| value.checked_add(relationship_target.capacity()))
                    .ok_or(Error::Internal(
                        "metadata planned-scope accounting overflow",
                    ))?
            },
        };
        total = total.checked_add(bytes).ok_or(Error::Internal(
            "metadata planned-scope accounting overflow",
        ))?;
    }
    Ok(total)
}

/// A staged edit of an owning package's document properties.
///
/// The edit shares the source package (an [`OpcPackage`] clone shares every
/// part payload) and the [`Snapshot`] read from it, and never changes them:
/// [`Edit::commit`] publishes a new package, and dropping the edit
/// discards it. Each family is either untouched (its part is never read or
/// written by the commit) or staged with a complete target value. Staging
/// a value equal to the current one commits nothing.
#[derive(Debug, Clone)]
pub struct Edit {
    package: OpcPackage,
    plan: Plan,
    /// Whether the plan's snapshot was read from `package` by this edit, so
    /// commit need not re-authenticate the patch against it.
    own_snapshot: bool,
}

impl Edit {
    /// Capture the document properties of `package` and start an edit.
    ///
    /// `host` decides which reserved custom properties belong to the
    /// document ([`custom::Props::read_for`]).
    ///
    /// # Errors
    ///
    /// The errors of [`Snapshot::read`].
    pub fn new(package: &OpcPackage, host: Host) -> Result<Self> {
        Self::new_with_options(package, host, EditOptions::default())
    }

    /// Capture `package` with edit options selected before replacement payload
    /// materialization.
    pub fn new_with_options(
        package: &OpcPackage,
        host: Host,
        options: EditOptions,
    ) -> Result<Self> {
        let snapshot = Snapshot::read(package, host)?;
        Ok(Self {
            package: package.clone(),
            plan: snapshot.plan_with_options(options),
            own_snapshot: true,
        })
    }

    /// Start an edit from a snapshot already captured by the host bridge.
    ///
    /// This keeps owner-specific admission policy in [`EditOptions`] without
    /// capturing the same property families twice. The snapshot remains an
    /// immutable read view; only [`Self::commit`] publishes a package.
    #[doc(hidden)]
    pub fn new_with_snapshot(
        package: &OpcPackage,
        snapshot: Snapshot,
        options: EditOptions,
    ) -> Self {
        Self {
            package: package.clone(),
            plan: snapshot.plan_with_options(options),
            own_snapshot: false,
        }
    }

    /// [`Self::new_with_snapshot`] for a snapshot the caller has just read
    /// from `package` itself with [`Snapshot::read`], so commit can skip
    /// re-authenticating the patch against `package` (record 0377a).
    #[doc(hidden)]
    pub fn new_with_own_snapshot(
        package: &OpcPackage,
        snapshot: Snapshot,
        options: EditOptions,
    ) -> Self {
        Self {
            package: package.clone(),
            plan: snapshot.plan_with_options(options),
            own_snapshot: true,
        }
    }

    /// The snapshot this edit started from.
    #[must_use]
    pub const fn base(&self) -> &Snapshot {
        self.plan.base()
    }

    /// The staged metadata: the staged value, or the snapshot's.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of [`Snapshot::metadata`] when no value is
    /// staged.
    pub fn metadata(&self) -> Result<Metadata> {
        self.plan.metadata()
    }

    /// Stage the complete metadata. Fields equal to the snapshot's change
    /// nothing; a changed field lands in the family that stores it.
    pub fn set_metadata(&mut self, metadata: Metadata) -> &mut Self {
        self.plan.set_metadata(metadata);
        self
    }

    /// Stage only the core-properties fields, retaining the staged or source
    /// application projection unchanged.
    pub fn set_core_metadata(&mut self, metadata: Metadata) -> Result<&mut Self> {
        self.plan.set_core_metadata(metadata)?;
        Ok(self)
    }

    /// Stage removal of the core-properties part.
    pub fn clear_metadata(&mut self) -> &mut Self {
        self.plan.clear_metadata();
        self
    }

    /// Stage a change of the metadata, starting from the staged value.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of [`Snapshot::metadata`].
    pub fn update_metadata(&mut self, update: impl FnOnce(&mut Metadata)) -> Result<&mut Self> {
        self.plan.update_metadata(update)?;
        Ok(self)
    }

    /// The staged application properties: the staged value, or the
    /// snapshot's.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of [`Snapshot::app_props`] when the family
    /// is not staged.
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        self.plan.app_props()
    }

    /// Stage the complete application properties; `None` removes the part.
    ///
    /// Metadata staged with [`Self::set_metadata`] is merged on top of this
    /// value at commit.
    pub fn set_app_props(&mut self, props: Option<app::Props>) -> &mut Self {
        self.plan.set_app_props(props);
        self
    }

    /// The staged custom properties: the staged value, or the snapshot's.
    #[must_use]
    pub fn custom_props(&self) -> &custom::Props {
        self.plan.custom_props()
    }

    /// Stage the complete custom properties; an empty collection removes the
    /// part when it changes the semantic collection. Setting an equal empty
    /// collection is an exact no-op, preserving a present raw part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> &mut Self {
        self.plan.set_custom_props(props);
        self
    }

    /// Stage explicit removal of the custom-properties part and relationship.
    ///
    /// This is distinct from setting an equal empty collection: it removes a
    /// present raw part even when its tolerant semantic projection is empty.
    /// An absent graph remains an exact no-op.
    pub fn remove_custom_props(&mut self) -> &mut Self {
        self.plan.remove_custom_props();
        self
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        let options = self.plan.options;
        self.plan = self.plan.base.plan_with_options(options);
    }

    /// Compute the exact patch of the staged values and publish it as a new
    /// package.
    ///
    /// An exact no-op publishes the source package itself, so a signed
    /// package keeps every byte; a changed edit of a signed package is
    /// refused with [`Error::Signed`] (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a changed field no OOXML family stores
    /// or content that this lossless route cannot carry; [`Error::Invalid`]
    /// for a core or custom part that holds content its writer cannot carry
    /// or a value outside its lexical space; [`Error::Internal`] for a
    /// violated prepared-plan or admission invariant; [`Error::Limit`] for
    /// an exceeded bound; [`Error::SourceChanged`] for source-backed
    /// freshness failure; and [`Error::Signed`] for a changed edit of a
    /// signed package.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            mut package,
            plan,
            own_snapshot,
        } = self;
        let (_, patch, diagnostics) = plan.finish()?;
        if !own_snapshot {
            let package = patch.apply(&package)?.into_package();
            return Ok(Commit::new(package, patch, diagnostics));
        }
        // The patch was planned from this edit's own snapshot of `package`,
        // which nothing has changed since, so `Patch::apply`'s source check
        // and second package clone are redundant here (record 0377a).
        if !patch.is_noop() {
            refuse_signed_change(&package)?;
            patch.apply_changes(&mut package)?;
        }
        Ok(Commit::new(package, patch, diagnostics))
    }
}

/// Publish a changed core-properties value of an owning package, reading
/// and diffing only the core family (record 0377a).
///
/// `target` replaces the core fields (`None` removes the part), as
/// [`Edit::set_core_metadata`] and [`Edit::clear_metadata`] stage them. The
/// application and custom families are neither read nor written. When the
/// core patch is an exact no-op the complete edit runs instead, so the
/// no-op patch keeps its authenticated scope of every family.
///
/// Returns the commit and the core metadata the new package holds (`None`
/// when the part is absent), so a host can adopt it without decoding the
/// part it has just written.
///
/// # Errors
///
/// The errors of [`Edit::new`], [`Edit::set_core_metadata`] and
/// [`Edit::commit`] for the core family.
#[doc(hidden)]
pub fn publish_core_metadata(
    package: &OpcPackage,
    host: Host,
    target: Option<Metadata>,
) -> Result<(Commit, Option<Metadata>)> {
    let snapshot = Snapshot::read_core_only(package, host)?;
    let mut plan = snapshot.plan();
    match &target {
        Some(metadata) => {
            plan.set_core_metadata(metadata.clone())?;
        },
        None => {
            plan.clear_metadata();
        },
    }
    let (after, patch, diagnostics) = plan.finish()?;
    if patch.is_noop() {
        let mut edit = Edit::new(package, host)?;
        match target {
            Some(metadata) => {
                edit.set_core_metadata(metadata)?;
            },
            None => {
                edit.clear_metadata();
            },
        }
        let commit = edit.commit()?;
        let core = Snapshot::read(commit.package(), host)?
            .core_value
            .as_ref()
            .map(Core::metadata);
        return Ok((commit, core));
    }
    refuse_signed_change(package)?;
    let mut candidate = package.clone();
    patch.apply_changes(&mut candidate)?;
    let core = after.core_value.as_ref().map(Core::metadata);
    Ok((Commit::new(candidate, patch, diagnostics), core))
}

fn next_relationship_id(used: &BTreeSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("rId{index}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        kind: LimitKind::Names,
        observed: usize::MAX,
        maximum: u32::MAX as usize,
    })
}

/// What a [`Commit`] changed.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Diagnostics {
    families: Box<[Family]>,
    fields: Box<[Field]>,
}

impl Diagnostics {
    /// Whether the commit changed any part.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.families.is_empty()
    }

    /// The families whose part changed, in [`Family`] order.
    #[must_use]
    pub fn families(&self) -> &[Family] {
        &self.families
    }

    /// The [`Metadata`] fields whose value changed, in [`Field::ALL`]
    /// order. Empty when only application or custom properties were staged
    /// directly.
    #[must_use]
    pub fn fields(&self) -> &[Field] {
        &self.fields
    }
}

/// A published document-properties change: the new package, the exact
/// [`Patch`] and what changed. Read the properties after it with
/// [`Snapshot::read`]`(commit.package(), host)`.
pub type Commit = litchi_core::edit::Commit<OpcPackage, Patch, Diagnostics>;

/// The exact change of one family's part.
///
/// `before` and `after` are the complete part payloads; `None` means the
/// part (and its package-level relationship) is absent on that side. The
/// relationship is the same on both sides.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    family: Family,
    part: PackURI,
    relationship_id: String,
    relationship_type: String,
    relationship_target: String,
    before: Option<Arc<Vec<u8>>>,
    after: Option<Arc<Vec<u8>>>,
    /// An absent family has no meaningful relationship ID. Snapshot scopes
    /// retain a free placeholder for diagnostics, but replay accepts any
    /// free ID because a later creation may choose a different gap.
    free_relationship_id: bool,
}

impl Change {
    /// The family whose part changes.
    #[must_use]
    pub const fn family(&self) -> Family {
        self.family
    }

    /// The part that changes.
    #[must_use]
    pub const fn part_name(&self) -> &PackURI {
        &self.part
    }

    /// The part's payload before the change; `None` when it is created.
    #[must_use]
    pub fn before(&self) -> Option<&[u8]> {
        self.before.as_deref().map(Vec::as_slice)
    }

    /// The part's payload after the change; `None` when it is removed.
    #[must_use]
    pub fn after(&self) -> Option<&[u8]> {
        self.after.as_deref().map(Vec::as_slice)
    }

    fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            ..self.clone()
        }
    }
}

/// An exact, reversible patch of a package's document properties: at most
/// one [`Change`] per family.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Patch {
    changes: Box<[Change]>,
    /// The recorded before state for an exact no-op. Changed patches already
    /// carry their before state in `changes`; an empty patch needs the full
    /// property-family scope or it could be replayed onto any package.
    scope: Option<Box<[Change]>>,
    /// The recorded after state for a patch whose before state carries a
    /// complete property-family scope. This is present for exact no-ops and
    /// for compositions that retain a complete scope while changing one or
    /// more families, so their inverse authenticates the published state too.
    after_scope: Option<Box<[Change]>>,
}

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.changes.is_empty()
    }

    /// The changes, one per changed family, in [`Family`] order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Conservative retained-data charge for this metadata patch.
    ///
    /// The package endpoints are charged by the owning format. This method
    /// covers the patch's own change descriptors, relationship strings,
    /// property-part payloads, and the complete before/after scope retained
    /// by an exact no-op. Scope entries are intentionally charged even when
    /// they share an `Arc` with a changed entry: history admission accounts
    /// for every retained patch owner independently and never relies on an
    /// allocator-layout assumption.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        fn change_bytes(change: &Change) -> usize {
            size_of::<Change>()
                .saturating_add(change.part.as_str().len())
                .saturating_add(change.relationship_id.capacity())
                .saturating_add(change.relationship_type.capacity())
                .saturating_add(change.relationship_target.capacity())
                .saturating_add(change.before.as_ref().map_or(0, |bytes| bytes.capacity()))
                .saturating_add(change.after.as_ref().map_or(0, |bytes| bytes.capacity()))
        }

        self.changes
            .iter()
            .chain(self.scope.iter().flat_map(|scope| scope.iter()))
            .chain(self.after_scope.iter().flat_map(|scope| scope.iter()))
            .fold(size_of::<Self>(), |total, change| {
                total.saturating_add(change_bytes(change))
            })
    }

    /// The patch that restores the before state of every change.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            changes: self.changes.iter().map(Change::inverse).collect(),
            scope: self.after_scope.clone().or_else(|| self.scope.clone()),
            after_scope: self.scope.clone().or_else(|| self.after_scope.clone()),
        }
    }

    fn check_changes(&self) -> &[Change] {
        self.scope.as_deref().unwrap_or(&self.changes)
    }

    /// What replaying this patch changes: its families, in [`Family`]
    /// order. A patch records parts, not [`Metadata`] fields, so the fields
    /// are empty.
    #[must_use]
    pub fn diagnostics(&self) -> Diagnostics {
        Diagnostics {
            families: self.changes.iter().map(|change| change.family).collect(),
            fields: Box::default(),
        }
    }

    /// Check that an owning package is in this patch's before state: each
    /// recorded family's graph selects exactly the recorded part through the
    /// recorded relationship with the recorded bytes, or (for a creation)
    /// holds no such family and the recorded relationship ID and part name
    /// are free; and a removed part has no other inbound relationship.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the first family whose state
    /// differs, and the graph errors of the family inspections.
    pub fn check(&self, package: &OpcPackage) -> Result<()> {
        for change in self.check_changes() {
            let graph = owning_graph(package, change.family)?;
            let bytes = match &graph.0 {
                Some(part) => Some(
                    package
                        .get_part(part)
                        .map_err(|error| source_error(crate::Error::Opc(error)))?
                        .blob_arc(),
                ),
                None => None,
            };
            check_change(package, change, graph, bytes.as_deref().map(Vec::as_slice))?;
        }
        // A composed patch keeps a supplemental scope for families it does
        // not change.  The scope authenticates the untouched state, but it
        // may intentionally wildcard a future relationship ID for an absent
        // family.  Check the actual changes as well so a creation still
        // authenticates its selected ID and a removal still checks that its
        // part has no additional inbound relationships.  Ordinary patches
        // have no scope and therefore take only the single loop above.
        if self.scope.is_some() {
            for change in &self.changes {
                let graph = owning_graph(package, change.family)?;
                let bytes = match &graph.0 {
                    Some(part) => Some(
                        package
                            .get_part(part)
                            .map_err(|error| source_error(crate::Error::Opc(error)))?
                            .blob_arc(),
                    ),
                    None => None,
                };
                check_change(package, change, graph, bytes.as_deref().map(Vec::as_slice))?;
            }
        }
        Ok(())
    }

    /// Check and apply the patch to an owning package, publishing a new
    /// package together with the patch and diagnostics.
    ///
    /// An empty patch changes nothing, so a signed package keeps its
    /// signature. A changed patch refuses a signed package before anything is
    /// mutated (ADR 0006). The source package is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a signed package, and the errors of
    /// [`Self::check`].
    pub fn apply(&self, package: &OpcPackage) -> Result<Commit> {
        self.check(package)?;
        if self.is_noop() {
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                self.diagnostics(),
            ));
        }
        refuse_signed_change(package)?;
        let mut candidate = package.clone();
        self.apply_changes(&mut candidate)?;
        Ok(Commit::new(candidate, self.clone(), self.diagnostics()))
    }

    #[cfg(test)]
    pub(super) fn apply_in_place(&self, package: &mut OpcPackage) -> Result<bool> {
        let commit = self.apply(package)?;
        let changed = commit.diagnostics().changed();
        *package = commit.into_package();
        Ok(changed)
    }

    fn apply_changes(&self, candidate: &mut OpcPackage) -> Result<()> {
        for change in &self.changes {
            match (&change.before, &change.after) {
                (Some(_), Some(after)) => {
                    candidate
                        .get_part_mut(&change.part)
                        .map_err(crate::Error::from)?
                        .set_blob_shared(Arc::clone(after));
                },
                (Some(_), None) => {
                    if !candidate.remove_part(&change.part) {
                        return Err(Error::Internal(
                            "metadata part disappeared after validation",
                        ));
                    }
                    if candidate
                        .rels_mut()
                        .remove(&change.relationship_id)
                        .is_none()
                    {
                        return Err(Error::Internal(
                            "metadata relationship disappeared after validation",
                        ));
                    }
                },
                (None, Some(after)) => {
                    candidate.try_add_part(Box::new(BlobPart::new_shared(
                        change.part.clone(),
                        change.family.content_type().to_owned(),
                        Arc::clone(after),
                    )))?;
                    candidate.rels_mut().try_add_relationship(
                        change.relationship_type.clone(),
                        change.relationship_target.clone(),
                        change.relationship_id.clone(),
                        TargetMode::Internal,
                    )?;
                },
                (None, None) => {},
            }
        }
        settle_unsigned_change(candidate)?;
        Ok(())
    }
}

impl Patch {
    /// The patch that applies this patch and then `next`.
    ///
    /// `next` must start from this patch's after state: for every family
    /// both patches change, `next`'s before bytes equal this patch's after
    /// bytes. A family whose composed before and after bytes are equal is
    /// dropped, so composing a patch with its inverse is a no-op.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when `next` does not continue this patch,
    /// or when one family is removed and then created again under another
    /// part or relationship (one [`Change`] keeps one relationship).
    pub fn compose(&self, next: &Self) -> Result<Self> {
        // A complete scope belongs to the package state at one end of the
        // patch. Preserve it when composing a changed patch: the changed
        // patch supplies the exact before/after image for the families it
        // touches, while the scope continues to authenticate every family it
        // leaves untouched. This keeps same-base history composition useful
        // without weakening the no-op patch's cross-family guard.
        let self_after = self.after_scope();
        if let (Some(first), Some(second)) = (self_after, next.scope.as_deref()) {
            check_scope_overlap(first, second)?;
        }
        if let Some(first) = self_after {
            check_scope_transition(first, &next.changes, false, "second")?;
        }
        if let Some(second) = next.scope.as_deref() {
            check_scope_transition(second, &self.changes, true, "first")?;
        }

        let mut changes = Vec::new();
        for family in [Family::Core, Family::App, Family::Custom] {
            let first = self.changes.iter().find(|change| change.family == family);
            let second = next.changes.iter().find(|change| change.family == family);
            let composed = match (first, second) {
                (None, None) => continue,
                (Some(change), None) | (None, Some(change)) => change.clone(),
                (Some(first), Some(second)) => {
                    if first.after != second.before {
                        return Err(Error::PatchConflict);
                    }
                    if first.part != second.part
                        || first.relationship_id != second.relationship_id
                        || first.relationship_type != second.relationship_type
                        || first.relationship_target != second.relationship_target
                    {
                        return Err(Error::PatchConflict);
                    }
                    Change {
                        before: first.before.clone(),
                        after: second.after.clone(),
                        ..first.clone()
                    }
                },
            };
            if composed.before != composed.after {
                changes.push(composed);
            }
        }
        let is_noop = changes.is_empty();
        let mut scope = compose_before_scope(self, next);
        if is_noop && scope.is_none() && !self.changes.is_empty() {
            let affected = scope_from_changes(&self.changes, false);
            scope = Some(affected.clone());
        }
        let after_scope = scope
            .as_deref()
            .map(|scope| scope_with_changes(scope, &changes, true));
        Ok(Self {
            changes: changes.into_boxed_slice(),
            scope,
            after_scope,
        })
    }

    /// Check that a source-backed package is in this patch's before state,
    /// with the rules of [`Self::check`]. Only the recorded families' parts
    /// are read, inside the package's source-freshness checks.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::check`], and a source-change or cancellation
    /// error.
    pub fn check_source_backed(&self, package: &SourceBackedPackage) -> Result<()> {
        package.check_execution()?;
        for change in self.check_changes() {
            let graph = source_graph(package, change.family)?;
            let bytes = match &graph.0 {
                Some(part) => Some(
                    package
                        .part(part)
                        .map_err(|error| source_error(crate::Error::Opc(error)))?
                        .data()
                        .map_err(|error| source_error(crate::Error::Opc(error)))?,
                ),
                None => None,
            };
            check_change(
                package,
                change,
                graph,
                bytes.as_ref().map(|data| data.as_bytes()),
            )?;
        }
        if self.scope.is_some() {
            for change in &self.changes {
                let graph = source_graph(package, change.family)?;
                let bytes = match &graph.0 {
                    Some(part) => Some(
                        package
                            .part(part)
                            .map_err(|error| source_error(crate::Error::Opc(error)))?
                            .data()
                            .map_err(|error| source_error(crate::Error::Opc(error)))?,
                    ),
                    None => None,
                };
                check_change(
                    package,
                    change,
                    graph,
                    bytes.as_ref().map(|data| data.as_bytes()),
                )?;
            }
        }
        package.check_execution()?;
        Ok(())
    }

    /// Plan this patch over a source-backed package without materializing
    /// it: each change becomes a part replacement, or a part and
    /// package-relationship addition or removal, of one
    /// [`SourceTopologyPlan`]. Every other member keeps its source bytes and
    /// compression when the returned topology is published.
    ///
    /// The package is checked first ([`Self::check_source_backed`]). An empty
    /// patch prepares the exact-source no-op topology, so a signed package
    /// keeps its signature; a changed patch refuses a package that carries
    /// signature infrastructure (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed patch of a signed package, the
    /// errors of [`Self::check_source_backed`], and the topology planner's
    /// source, limit and allocation errors.
    pub fn prepare_source_backed(&self, package: &SourceBackedPackage) -> Result<PreparedTopology> {
        let mut plan = SourceTopologyPlan::new();
        self.check_source_backed(package)?;
        if !self.is_noop() {
            if package.has_signature_infrastructure() {
                return Err(Error::Signed);
            }
            let root = PackURI::new("/")
                .map_err(|_| Error::Internal("the OPC package root URI became invalid"))?;
            for change in &self.changes {
                match (&change.before, &change.after) {
                    (Some(_), Some(after)) => {
                        plan.try_replace_part_shared(change.part.clone(), Arc::clone(after))?;
                    },
                    (Some(_), None) => {
                        plan.try_remove_part(change.part.clone())?;
                        plan.try_remove_relationship(root.clone(), change.relationship_id.clone())?;
                    },
                    (None, Some(after)) => {
                        plan.try_add_part_shared(
                            change.part.clone(),
                            change.family.content_type(),
                            Arc::clone(after),
                        )?;
                        plan.try_add_internal_relationship(
                            root.clone(),
                            change.relationship_id.clone(),
                            change.relationship_type.clone(),
                            change.part.clone(),
                        )?;
                    },
                    (None, None) => {},
                }
            }
        }
        Ok(package.clone().prepare_topology(plan)?)
    }

    fn after_scope(&self) -> Option<&[Change]> {
        self.after_scope
            .as_deref()
            .or_else(|| self.is_noop().then_some(self.scope.as_deref()).flatten())
    }
}

/// Build the before-state guard for a composed patch. A complete no-op scope
/// contributes every untouched family; a partial cancellation scope
/// contributes only its retained families, and disjoint changed families are
/// appended from their own before images.
fn compose_before_scope(first: &Patch, second: &Patch) -> Option<Box<[Change]>> {
    let mut entries = Vec::new();
    for family in [Family::Core, Family::App, Family::Custom] {
        let first_change = first.changes.iter().find(|change| change.family == family);
        let second_change = second.changes.iter().find(|change| change.family == family);
        let entry = first_change
            .map(|change| state_change(change, false))
            .or_else(|| {
                first
                    .scope
                    .as_deref()
                    .and_then(|scope| scope.iter().find(|change| change.family == family))
                    .cloned()
            })
            .or_else(|| {
                second
                    .scope
                    .as_deref()
                    .and_then(|scope| scope.iter().find(|change| change.family == family))
                    .cloned()
            })
            .or_else(|| second_change.map(|change| state_change(change, false)));
        if let Some(entry) = entry {
            entries.push(entry);
        }
    }
    (!entries.is_empty()).then(|| entries.into_boxed_slice())
}

/// Check the changed families at one end of a complete property scope.
///
/// `after` selects the first patch's after image; otherwise the next patch's
/// before image is selected. The scope itself still contains a state image in
/// both `before` and `after`, so it can be used as an exact owner guard.
fn check_scope_transition(
    scope: &[Change],
    changes: &[Change],
    after: bool,
    _side: &str,
) -> Result<()> {
    let complete = scope_is_complete(scope);
    for change in changes {
        let Some(expected) = scope.iter().find(|scoped| scoped.family == change.family) else {
            if complete {
                return Err(Error::PatchConflict);
            }
            // A partial cancellation scope authenticates only the families it
            // retained. A disjoint changed family still carries its own
            // before image and may be composed independently.
            continue;
        };
        if !same_change_identity(expected, change)
            || expected.before.as_deref()
                != (if after {
                    change.after.as_deref()
                } else {
                    change.before.as_deref()
                })
        {
            return Err(Error::PatchConflict);
        }
    }
    Ok(())
}

fn check_scope_overlap(first: &[Change], second: &[Change]) -> Result<()> {
    for first in first {
        let Some(second) = second.iter().find(|change| change.family == first.family) else {
            continue;
        };
        if !same_change_identity(first, second)
            || first.before != second.before
            || first.after != second.after
        {
            return Err(Error::PatchConflict);
        }
    }
    Ok(())
}

fn scope_is_complete(scope: &[Change]) -> bool {
    scope.len() == 3
        && [Family::Core, Family::App, Family::Custom]
            .into_iter()
            .all(|family| scope.iter().any(|change| change.family == family))
}

fn same_change_identity(first: &Change, second: &Change) -> bool {
    first.family == second.family
        && first.part == second.part
        && (first.free_relationship_id
            || second.free_relationship_id
            || first.relationship_id == second.relationship_id)
        && first.relationship_type == second.relationship_type
        && first.relationship_target == second.relationship_target
}

/// Turn one changed family into a state entry whose before and after images
/// are the selected endpoint. The graph identity is retained, including the
/// free relationship ID and part name for a creation/removal.
fn state_change(change: &Change, after: bool) -> Change {
    let bytes = if after {
        change.after.clone()
    } else {
        change.before.clone()
    };
    Change {
        before: bytes.clone(),
        after: bytes,
        ..change.clone()
    }
}

/// Replace changed-family entries in a complete scope with an endpoint of a
/// patch. Unchanged families retain their exact bytes and graph identity.
fn scope_with_changes(scope: &[Change], changes: &[Change], after: bool) -> Box<[Change]> {
    scope
        .iter()
        .map(|entry| {
            changes
                .iter()
                .find(|change| change.family == entry.family)
                .map(|change| state_change(change, after))
                .unwrap_or_else(|| entry.clone())
        })
        .collect()
}

fn scope_from_changes(changes: &[Change], after: bool) -> Box<[Change]> {
    changes
        .iter()
        .map(|change| state_change(change, after))
        .collect()
}

/// The selected part and relationship ID of one family of a source-backed
/// package, strictly inspected.
fn source_graph(
    package: &SourceBackedPackage,
    family: Family,
) -> Result<(Option<PackURI>, Option<String>)> {
    Ok(match family {
        Family::Core => {
            let graph = graph::inspect_source_strict(package).map_err(source_error)?;
            (graph.part, graph.relationship_id)
        },
        Family::App => {
            let graph = app::inspect_source(package).map_err(source_error)?;
            (graph.part, graph.relationship_id)
        },
        Family::Custom => {
            let graph = custom::inspect_source_graph(package).map_err(source_error)?;
            (graph.part_name, graph.relationship_id)
        },
    })
}

/// The selected part and relationship ID of one family of an owning
/// package, strictly inspected.
fn owning_graph(package: &OpcPackage, family: Family) -> Result<(Option<PackURI>, Option<String>)> {
    Ok(match family {
        Family::Core => {
            let graph = graph::inspect(package).map_err(source_error)?;
            (graph.part, graph.relationship_id)
        },
        Family::App => {
            let graph = app::inspect(package).map_err(source_error)?;
            (graph.part, graph.relationship_id)
        },
        Family::Custom => {
            let graph = custom::inspect_graph(package).map_err(source_error)?;
            (graph.part_name, graph.relationship_id)
        },
    })
}

fn check_change<P: PackageMetadata>(
    package: &P,
    change: &Change,
    (part, relationship_id): (Option<PackURI>, Option<String>),
    bytes: Option<&[u8]>,
) -> Result<()> {
    let stale = || Error::PatchConflict;
    let relationships = package.package_relationships();
    match &change.before {
        Some(before) => {
            let (Some(part), Some(relationship_id)) = (part, relationship_id) else {
                return Err(stale());
            };
            if part != change.part
                || relationship_id != change.relationship_id
                || bytes != Some(before.as_slice())
            {
                return Err(stale());
            }
            let relationship = relationships.get(&relationship_id).ok_or_else(stale)?;
            if relationship.reltype() != change.relationship_type
                || relationship.target_ref() != change.relationship_target
                || relationship.is_external()
            {
                return Err(stale());
            }
            if change.after.is_none() {
                ensure_only_inbound(
                    package,
                    &change.part,
                    &change.relationship_id,
                    change.family.label(),
                )
                .map_err(source_error)?;
            }
        },
        None => {
            if part.is_some()
                || relationship_id.is_some()
                || (!change.free_relationship_id
                    && relationships.get(&change.relationship_id).is_some())
                || package
                    .metadata_parts()
                    .any(|candidate| names_conflict(candidate.part_name(), &change.part))
            {
                return Err(stale());
            }
        },
    }
    Ok(())
}

#[cfg(test)]
mod classification_tests {
    use super::{
        Error, app_authored_error, app_plan_error, app_prepared_error, core_generated_error,
        custom_generated_error, patch_output_error,
    };

    #[test]
    fn source_and_prepared_codec_errors_keep_their_contract_classification() {
        assert!(matches!(
            app_plan_error(crate::Error::Invalid(
                "unexpected source element".to_owned()
            )),
            Error::Package(crate::Error::Invalid(_))
        ));
        assert!(matches!(
            app_plan_error(crate::Error::Invalid(
                "extended property Company contains XML 1.0-forbidden character U+0000"
                    .to_owned()
            )),
            Error::Package(crate::Error::Invalid(message)) if message.contains("U+0000")
        ));

        for error in [
            app_prepared_error(crate::Error::Invalid(
                "extended-properties edits overlap".to_owned(),
            )),
            core_generated_error(crate::Error::Invalid(
                "core-properties encoded-size invariant failed".to_owned(),
            )),
            custom_generated_error(crate::Error::Invalid(
                "custom-properties encoded-size invariant failed".to_owned(),
            )),
            patch_output_error(crate::Error::Invalid(
                "generated metadata patch did not decode".to_owned(),
            )),
        ] {
            assert!(matches!(error, Error::Internal(_)));
        }

        let authored = app_authored_error(crate::Error::Invalid(
            "extended property Company contains XML 1.0-forbidden character U+0000".to_owned(),
        ));
        assert!(matches!(authored, Error::Invalid(message) if message.contains("U+0000")));

        let bounded = app_authored_error(crate::Error::Limit {
            kind: crate::LimitKind::ExtendedPropertiesXmlBytes,
            maximum: 8 * 1024 * 1024,
            observed: 8 * 1024 * 1024 + 1,
        });
        assert!(matches!(bounded, Error::Limit { .. }));
    }
}
