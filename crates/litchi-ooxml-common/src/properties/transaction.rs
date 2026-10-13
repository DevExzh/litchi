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
//! 2. [`Snapshot::edit`] starts an [`Edit`], which stages a whole
//!    [`Metadata`], application properties, or custom properties;
//! 3. [`Edit::commit`] computes a [`Patch`]: the exact before and after bytes
//!    of each part that changes, the part created when a family was absent,
//!    or the part removed when the custom family is emptied;
//! 4. the host applies the patch ([`Patch::apply`]) to a copy of its package
//!    and publishes the result as a new owner.
//!
//! Each [`Metadata`] field lands in the family that stores it
//! ([`super::CORE_FIELDS`], [`super::app::METADATA_FIELDS`]); changing a
//! field no OOXML family stores (`content_type`, `last_backup`, `codepage`)
//! is refused with [`Error::Unsupported`], never dropped. A field left
//! unchanged keeps its exact lexical form, and an exact no-op commits an
//! empty patch that publishes the package unchanged, a signed one included.
//! A changed patch refuses a signed package ([`Error::Signed`], ADR 0006).
//! Nothing stamps `modified`, `lastModifiedBy` or a revision unless the
//! caller sets it (ADR 0006, determinism).

use std::collections::BTreeSet;
use std::sync::Arc;

use litchi_core::metadata::{Field, Metadata};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::rel::TargetMode;
use litchi_opc::{OpcError, OpcPackage, PackURI, SourceBackedPackage};

use super::package_graph::{
    GraphPart as _, PackageMetadata, ensure_only_inbound, free_part_name, names_conflict,
    other_inbound,
};
use super::{CORE_FIELDS, Core, Dialect, app, graph, read};
use crate::custom::{self, Host};
use crate::leniency::ToleranceReport;
use crate::signature::{refuse_signed_change, settle_unsigned_change};

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
}

/// Why a metadata read, edit or patch failed.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The field cannot be stored by any OOXML property family.
    #[error("OOXML packages cannot store the metadata field `{field}`")]
    Unsupported {
        /// The refused field.
        field: Field,
    },
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
    /// The patch does not match the package's current property parts.
    #[error("the metadata patch does not match the package: {0}")]
    Conflict(String),
    /// Another package, graph or XML failure.
    #[error(transparent)]
    Package(crate::Error),
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        match error {
            crate::Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy) => Self::Signed,
            crate::Error::Invalid(message) => Self::Invalid(message),
            crate::Error::Limit {
                resource,
                max,
                actual,
            } => {
                let kind = if resource.contains("text") {
                    Some(LimitKind::TextBytes)
                } else if resource.contains("events") {
                    Some(LimitKind::XmlEvents)
                } else if resource.contains("XML bytes") {
                    Some(LimitKind::XmlBytes)
                } else if resource.contains("custom") {
                    Some(LimitKind::CustomProperties)
                } else if resource.contains("names") || resource.contains("identifiers") {
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
                        resource,
                        max,
                        actual,
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
        let core_graph = graph::inspect(package)?;
        let app_graph = app::inspect(package)?;
        let custom_graph = custom::inspect_graph(package)?;
        let bytes = |part: &PackURI| -> Result<Arc<Vec<u8>>> {
            Ok(package
                .get_part(part)
                .map_err(crate::Error::from)?
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
        let core_graph = graph::inspect_source_strict(package)?;
        let app_graph = app::inspect_source(package)?;
        let custom_graph = custom::inspect_source_graph(package)?;
        let bytes = |part: &PackURI| -> Result<Arc<Vec<u8>>> {
            let data = package.part(part)?.data()?;
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
        let core = family_state(package, Family::Core, core_part, core_id, &bytes)?;
        let app = family_state(package, Family::App, app_part, app_id, &bytes)?;
        let custom = family_state(package, Family::Custom, custom_part, custom_id, &bytes)?;
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
            .transpose()?;
        let custom_value = match &custom.present {
            Some(present) => custom::Props::decode_for(&present.bytes, &present.part, host)?,
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
        Ok(self
            .app
            .present
            .as_ref()
            .map(|present| app::decode(&present.bytes))
            .transpose()?)
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

    /// The exact bytes of the part of `family`, when present.
    #[must_use]
    pub fn part_bytes(&self, family: Family) -> Option<&[u8]> {
        self.family(family)
            .present
            .as_ref()
            .map(|present| present.bytes.as_slice())
    }

    /// Start an edit of this snapshot.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            base: self.clone(),
            metadata: None,
            app: None,
            custom: None,
        }
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
    let new_part = free_part_name(package, family.stem(), "document-property part names");
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

/// A staged edit of a [`Snapshot`]'s document properties.
///
/// Each family is either untouched (its part is never read or written by
/// the commit) or staged with a complete target value. Staging a value
/// equal to the current one commits nothing.
#[derive(Debug, Clone)]
pub struct Edit {
    base: Snapshot,
    metadata: Option<Metadata>,
    app: Option<Option<app::Props>>,
    custom: Option<custom::Props>,
}

impl Edit {
    /// The snapshot this edit started from.
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
        match &self.metadata {
            Some(staged) => Ok(staged.clone()),
            None => self.base.metadata(),
        }
    }

    /// Stage the complete metadata. Fields equal to the snapshot's change
    /// nothing; a changed field lands in the family that stores it.
    pub fn set_metadata(&mut self, metadata: Metadata) -> &mut Self {
        self.metadata = Some(metadata);
        self
    }

    /// Stage a change of the metadata, starting from the staged value.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of [`Snapshot::metadata`].
    pub fn update_metadata(&mut self, update: impl FnOnce(&mut Metadata)) -> Result<&mut Self> {
        let mut metadata = self.metadata()?;
        update(&mut metadata);
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
    /// part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> &mut Self {
        self.custom = Some(props);
        self
    }

    /// Compute the exact patch of the staged values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Unsupported`] for a changed field no OOXML family
    /// stores; [`Error::Invalid`] for a core or custom part that holds
    /// content its writer cannot carry, a value outside its lexical space,
    /// or removing a part another relationship targets; and
    /// [`Error::Limit`] for an exceeded bound.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            base,
            metadata,
            app: staged_app,
            custom,
        } = self;
        let mut after = base.clone();
        let mut changes = Vec::new();
        let mut changed_fields = Vec::new();

        let (core_target, app_target) = match metadata {
            Some(target) => {
                let before = base.metadata()?;
                changed_fields = before.changed_fields(&target).collect();
                if let Some(field) = changed_fields.iter().copied().find(|field| {
                    !CORE_FIELDS.contains(field) && !app::METADATA_FIELDS.contains(field)
                }) {
                    return Err(Error::Unsupported { field });
                }
                let core_changed = changed_fields
                    .iter()
                    .any(|field| CORE_FIELDS.contains(field));
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
                (core_changed.then_some(target), app_target)
            },
            None => (None, staged_app),
        };

        if let Some(target) = core_target {
            let after_bytes = match &base.core.present {
                Some(present) => {
                    let xml = std::str::from_utf8(&present.bytes).map_err(|error| {
                        Error::Invalid(format!("invalid UTF-8 in core properties: {error}"))
                    })?;
                    let (current, dialect) = read::decode(xml)?;
                    if Some(dialect) != base.core_dialect {
                        return Err(Error::Invalid(
                            "core-properties XML namespace does not match its relationship dialect"
                                .to_owned(),
                        ));
                    }
                    let merged = Core::merged(Some(&current), &target)?;
                    if merged == current {
                        None
                    } else {
                        Some((super::write::encode(&merged, dialect)?.into_bytes(), merged))
                    }
                },
                None => {
                    let merged = Core::merged(None, &target)?;
                    Some((
                        super::write::encode(&merged, Dialect::Transitional)?.into_bytes(),
                        merged,
                    ))
                },
            };
            if let Some((bytes, merged)) = after_bytes
                && let Some(change) = stage_change(
                    &mut after,
                    Family::Core,
                    false,
                    Some(bytes),
                    Dialect::Transitional.relationship_type(),
                )?
            {
                changes.push(change);
                after.core_dialect = after.core_dialect.or(Some(Dialect::Transitional));
                after.core_value = Some(merged);
                after.core_report = ToleranceReport::default();
            }
        }

        if let Some(target) = app_target {
            let current = base.app_props()?;
            if target != current {
                let after_bytes = match (&base.app.present, &target) {
                    (Some(present), Some(props)) => app::splice(&present.bytes, props)?,
                    (None, Some(props)) => Some(app::encode_new(props)?),
                    (_, None) => None,
                };
                if let Some(change) = stage_change(
                    &mut after,
                    Family::App,
                    target.is_none(),
                    after_bytes,
                    rt::EXTENDED_PROPERTIES,
                )? {
                    changes.push(change);
                }
            }
        }

        if let Some(target) = custom
            && !target.same_values(&base.custom_value)
        {
            target.validate_for(base.host)?;
            if !base.custom_value.tolerance_report().is_clean() {
                return Err(Error::Invalid(
                    "the custom properties were repaired when they were read; rewriting them would drop the properties that were skipped"
                        .to_owned(),
                ));
            }
            let after_bytes = if target.is_empty() {
                None
            } else {
                Some(target.encode()?)
            };
            if let Some(change) = stage_change(
                &mut after,
                Family::Custom,
                target.is_empty(),
                after_bytes,
                rt::CUSTOM_PROPERTIES,
            )? {
                changes.push(change);
                after.custom_value = target;
            }
        }

        let families = changes.iter().map(|change| change.family).collect();
        if changes.is_empty() {
            changed_fields.clear();
        }
        Ok(Commit {
            snapshot: after,
            patch: Patch {
                changes: changes.into_boxed_slice(),
            },
            diagnostics: Diagnostics {
                families,
                fields: changed_fields.into_boxed_slice(),
            },
        })
    }
}

/// Turn one family's target bytes into a change, and update `after`.
///
/// `remove` stages the removal of a present part; otherwise `bytes` is the
/// new payload, `None` when it equals the current one.
fn stage_change(
    after: &mut Snapshot,
    family: Family,
    remove: bool,
    bytes: Option<Vec<u8>>,
    new_relationship_type: &str,
) -> Result<Option<Change>> {
    let state = after.family(family).clone();
    let change = match (state.present, remove, bytes) {
        (Some(present), true, _) => {
            if let Some(reason) = state.shared_by {
                return Err(Error::Invalid(reason));
            }
            after.relationship_ids.remove(&present.relationship_id);
            let slot = after.family_mut(family);
            slot.new_part = present.part.clone();
            slot.present = None;
            Change {
                family,
                part: present.part,
                relationship_id: present.relationship_id,
                relationship_type: present.relationship_type,
                relationship_target: present.relationship_target,
                before: Some(present.bytes),
                after: None,
            }
        },
        (Some(present), false, Some(bytes)) => {
            if bytes.as_slice() == present.bytes.as_slice() {
                return Ok(None);
            }
            let bytes = Arc::new(bytes);
            let slot = after.family_mut(family);
            if let Some(slot) = slot.present.as_mut() {
                slot.bytes = Arc::clone(&bytes);
            }
            Change {
                family,
                part: present.part,
                relationship_id: present.relationship_id,
                relationship_type: present.relationship_type,
                relationship_target: present.relationship_target,
                before: Some(present.bytes),
                after: Some(bytes),
            }
        },
        (None, false, Some(bytes)) => {
            let relationship_id = next_relationship_id(&after.relationship_ids)?;
            after.relationship_ids.insert(relationship_id.clone());
            let part = state.new_part;
            let bytes = Arc::new(bytes);
            let present = Present {
                relationship_type: new_relationship_type.to_owned(),
                relationship_target: part.as_str().trim_start_matches('/').to_owned(),
                bytes: Arc::clone(&bytes),
                part: part.clone(),
                relationship_id: relationship_id.clone(),
            };
            let change = Change {
                family,
                part,
                relationship_id,
                relationship_type: present.relationship_type.clone(),
                relationship_target: present.relationship_target.clone(),
                before: None,
                after: Some(bytes),
            };
            let slot = after.family_mut(family);
            slot.present = Some(present);
            slot.shared_by = None;
            change
        },
        (_, _, None) | (None, true, _) => return Ok(None),
    };
    Ok(Some(change))
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

/// The result of [`Edit::commit`]: the document properties after the edit,
/// the patch that produces them, and what changed.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The document properties after the patch.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The exact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the commit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit into its snapshot, patch and diagnostics.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, Diagnostics) {
        (self.snapshot, self.patch, self.diagnostics)
    }
}

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

    /// The patch that restores the before state of every change.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            changes: self.changes.iter().map(Change::inverse).collect(),
        }
    }

    /// Check that an owning package is in this patch's before state: each
    /// family's graph selects exactly the recorded part through the
    /// recorded relationship with the recorded bytes, or (for a creation)
    /// holds no such family and the recorded relationship ID and part name
    /// are free; and a removed part has no other inbound relationship.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] naming the first family whose state
    /// differs, and the graph errors of the family inspections.
    pub fn check(&self, package: &OpcPackage) -> Result<()> {
        for change in &self.changes {
            let graph = owning_graph(package, change.family)?;
            let bytes = match &graph.0 {
                Some(part) => Some(
                    package
                        .get_part(part)
                        .map_err(crate::Error::from)?
                        .blob_arc(),
                ),
                None => None,
            };
            check_change(package, change, graph, bytes.as_deref().map(Vec::as_slice))?;
        }
        Ok(())
    }

    /// Apply the patch to an owning package after [`Self::check`]. Returns
    /// whether the package changed.
    ///
    /// An empty patch changes nothing, so a signed package keeps its
    /// signature. A changed patch refuses a signed package before anything is
    /// mutated (ADR 0006). The package changes atomically: a failure leaves
    /// it as it was.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a signed package, and the errors of
    /// [`Self::check`].
    pub fn apply(&self, package: &mut OpcPackage) -> Result<bool> {
        if self.is_noop() {
            return Ok(false);
        }
        self.check(package)?;
        refuse_signed_change(package)?;
        let mut candidate = package.clone();
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
                        return Err(Error::Conflict(change.part.to_string()));
                    }
                    if candidate
                        .rels_mut()
                        .remove(&change.relationship_id)
                        .is_none()
                    {
                        return Err(Error::Conflict(format!(
                            "{} relationship '{}' disappeared during removal",
                            change.family.label(),
                            change.relationship_id
                        )));
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
        settle_unsigned_change(&mut candidate)?;
        *package = candidate;
        Ok(true)
    }
}

/// The selected part and relationship ID of one family of an owning
/// package, strictly inspected.
fn owning_graph(package: &OpcPackage, family: Family) -> Result<(Option<PackURI>, Option<String>)> {
    Ok(match family {
        Family::Core => {
            let graph = graph::inspect(package)?;
            (graph.part, graph.relationship_id)
        },
        Family::App => {
            let graph = app::inspect(package)?;
            (graph.part, graph.relationship_id)
        },
        Family::Custom => {
            let graph = custom::inspect_graph(package)?;
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
    let stale = || {
        Error::Conflict(format!(
            "the package's {} part is not the patch's before state",
            change.family.label()
        ))
    };
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
                )?;
            }
        },
        None => {
            if part.is_some()
                || relationship_id.is_some()
                || relationships.get(&change.relationship_id).is_some()
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
