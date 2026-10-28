//! Slide-table duplication and removal (record 0372; the Keynote leaf of the
//! host's `duplicate_slide_table` and `remove_slide_table`) over the shared
//! table structure set (`litchi_iwa_table::structure`, record 0379).
//!
//! Duplication is the shared table-graph copy (record 0370h): fresh
//! identifiers for the drawable, model and owned storage, the host's table
//! UUID, its CalculationEngine owner family and the clones' registry
//! records. This owner names the copy as the host does (`"<name> copy"`,
//! then a numeric suffix, unique on the slide), offsets it by ten points,
//! appends it to the slide's owned drawables and z-order (in front), and
//! records the cross-component references to it.
//!
//! Removal takes out the drawable's private object closure and the
//! CalculationEngine owners of its contexts, detaches it from the slide's
//! owned drawables, z-order, template tags and instructional text, and
//! releases a removed identifier suffix, as the host does. A surviving
//! object that still references the table refuses the removal.

use std::collections::BTreeSet;
use std::fmt;
use std::num::NonZeroU64;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::WireLimits;
use litchi_iwa_table::structure::budget::WireBudget;
use litchi_iwa_table::structure::collect::{Census, verify_removed};
use litchi_iwa_table::structure::duplicate::{self as engine, Drawable};
use litchi_iwa_table::structure::metadata::{self, Transition};
use litchi_iwa_table::structure::plan::{
    DeletionObjectIndex, GraphPlan, ObjectRemoval, RemovalPlan,
};
use litchi_iwa_table::structure::{
    Fault, FaultLimit, Published, Working, copy_bytes, formula, references, reserve, unique_message,
};
use thiserror::Error as ThisError;

use super::structure::wire::{self, Fields, LENGTH_DELIMITED};
use super::{Package, slide_table_core as core};
use crate::SlideSelector;
use crate::slide::table::TableSelector;

/// `KN.SlideArchive`.
const SLIDE_MESSAGE_TYPE: u32 = 5;
/// `KN.SlideArchive.owned_drawables`.
const OWNED_DRAWABLES: u32 = 7;
/// `KN.SlideArchive.sage_tag_to_info_map`.
const TEMPLATE_TAGS: u32 = 28;
/// `KN.SlideArchive.SageTagMapEntry.info`.
const TEMPLATE_TAG_INFO: u32 = 2;
/// `KN.SlideArchive.drawables_z_order`.
const DRAWABLES_Z_ORDER: u32 = 42;
/// `KN.SlideArchive.instructional_text_map`.
const INSTRUCTIONAL_TEXT: u32 = 45;
/// `KN.InstructionalTextMap.instructional_text_for_infos`.
const INSTRUCTIONAL_TEXT_ENTRIES: u32 = 1;
/// `KN.InstructionalTextMap.InstructionalTextMapEntry.info`.
const INSTRUCTIONAL_TEXT_INFO: u32 = 1;
/// `TST.TableInfoArchive`.
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
/// `TST.TableInfoArchive.tableModel`.
const TABLE_INFO_MODEL: u32 = 2;
/// `TST.TableModelArchive`.
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
/// `TST.TableModelArchive.table_name`.
const TABLE_MODEL_NAME: u32 = 8;
/// How far the copy is offset from its source (the host's
/// `TABLE_DUPLICATE_OFFSET`).
const OFFSET: f32 = 10.0;

/// A finite resource governed by the transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
    /// Work spent resolving the slide and table selectors.
    Selection,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "object identifiers",
            Self::Selection => "selection work",
        })
    }
}

/// A typed, content-redacted slide-table lifecycle failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted table graph is malformed or ambiguous.
    #[error("the selected Keynote slide-table graph is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error(
        "Keynote slide-table lifecycle {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Keynote slide-table lifecycle items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested tables.
    #[error("the Keynote slide-table lifecycle candidate failed verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Keynote slide-table lifecycle patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by
/// [`Refusal::kind`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No slide or slide table matches the selectors.
    #[error("no Keynote slide or slide table matches the selectors")]
    TableNotFound,
    /// A selector is ambiguous.
    #[error("the Keynote slide or table selector is ambiguous")]
    AmbiguousSelector,
    /// The source keeps no exact artifact, or the table has a native shape
    /// this route does not own.
    #[error("this Keynote source does not support exact slide-table lifecycle changes")]
    UnsupportedSource,
    /// The table's formula owner state cannot be copied (volatile,
    /// spanning, whole-owner, error or spill state, as the host refuses
    /// it), or a surviving object depends on the removed table.
    #[error("the Keynote slide table has a dependency this route cannot carry")]
    UnsupportedDependency,
    /// The slide has no free copy name for the table.
    #[error("the Keynote slide has no free name for the table copy")]
    NoCopyName,
    /// This edit already holds its single operation.
    #[error("a Keynote slide-table lifecycle operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested before staging an operation.
    #[error("no Keynote slide-table lifecycle operation is staged")]
    NoStagedOperation,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound => RefusalKind::NotFound,
            Self::AmbiguousSelector => RefusalKind::InvalidValue,
            Self::UnsupportedSource | Self::UnsupportedDependency | Self::NoCopyName => {
                RefusalKind::UnsupportedContent
            },
            Self::OperationAlreadyStaged | Self::NoStagedOperation => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
            Fault::Invalid => Self::InvalidSource,
            Fault::Dependency => Self::Refused(Refusal::UnsupportedDependency),
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

fn select_error(error: core::Error) -> Error {
    match error {
        core::Error::SlideNameNotFound
        | core::Error::SlidePositionNotFound(_)
        | core::Error::TablePositionNotFound(_)
        | core::Error::EmptySlideName => Error::Refused(Refusal::TableNotFound),
        core::Error::AmbiguousSelector => Error::Refused(Refusal::AmbiguousSelector),
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        core::Error::UnsupportedDependency | core::Error::UnsupportedTopology => {
            Error::Refused(Refusal::UnsupportedDependency)
        },
        core::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::Selection,
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// Preserve bounded-work failures while redacting the candidate's native
/// read detail.  Candidate verification is still a focused-route operation:
/// a limit or allocation refusal must not turn into an untyped verification
/// failure (record 0351 §2.4).
fn reopen_error(error: super::ReadError) -> Error {
    match error {
        super::ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Identifiers,
                _ => LimitKind::Selection,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        super::ReadError::PayloadLimit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        super::ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        super::ReadError::Archive(error) => match error {
            litchi_iwa_archive::Error::Limit {
                observed, maximum, ..
            } => Error::Limit {
                kind: LimitKind::PackageBytes,
                observed,
                maximum,
            },
            litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
            _ => Error::Verification,
        },
        _ => Error::Verification,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Action {
    Duplicate,
    Remove,
}

/// The one staged operation and the table it resolved to.
#[derive(Debug, Clone)]
struct Staged {
    action: Action,
    target: core::Target,
}

/// One immutable edit that stages exactly one duplication or removal of a
/// slide table.
///
/// Start it with [`Package::edit_slide_table_lifecycle`]; the source package
/// is never changed, and dropping the edit discards the staged operation.
pub struct Edit<'a> {
    source: &'a Package,
    staged: Option<Staged>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.staged.as_ref().map(|staged| staged.action))
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn stage(
        mut self,
        slide: SlideSelector<'_>,
        table: TableSelector,
        action: Action,
    ) -> Result<Self, Error> {
        if self.staged.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        let mut budget = core::Budget::new(self.source).map_err(select_error)?;
        let target =
            core::select_table(self.source, slide, table, &mut budget).map_err(select_error)?;
        self.staged = Some(Staged { action, target });
        Ok(self)
    }

    /// Stage a populated copy of one slide table on its slide, named
    /// `"<name> copy"` (or `"<name> copy 2"`, ... when that name is taken on
    /// the slide), offset by ten points and placed in front (host
    /// `duplicate_slide_table`).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] when the selectors resolve no table;
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn duplicate<'s>(
        self,
        slide: impl Into<SlideSelector<'s>>,
        table: impl Into<TableSelector>,
    ) -> Result<Self, Error> {
        self.stage(slide.into(), table.into(), Action::Duplicate)
    }

    /// Stage the removal of one slide table with its private storage and
    /// formula owners (host `remove_slide_table`).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] when the selectors resolve no table;
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn remove<'s>(
        self,
        slide: impl Into<SlideSelector<'s>>,
        table: impl Into<TableSelector>,
    ) -> Result<Self, Error> {
        self.stage(slide.into(), table.into(), Action::Remove)
    }

    /// Publish and reopen the staged operation as one exact-source
    /// transaction.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit; the shared
    /// engine's typed refusals; [`Error::Verification`] when the reopened
    /// package does not show the slide's expected tables. Failure never
    /// publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let Staged { action, target } = self
            .staged
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let slide = target.slide.identifier;
        let mut expected = slide_table_names(self.source, slide)?;
        let position = target.table_position.get();
        let published = match action {
            Action::Duplicate => {
                let current = expected.get(position).ok_or(Error::InvalidSource)?;
                let existing = expected.iter().map(String::as_str).collect::<Vec<_>>();
                let name = engine::copy_name(current, &existing)
                    .ok_or(Error::Refused(Refusal::NoCopyName))?;
                let (published, copy) = duplicate(self.source, &target, &name)?;
                expected
                    .try_reserve(1)
                    .map_err(|_| Error::Allocation { amount: 1 })?;
                expected.push(name);
                verify_copy(&published.package, &target, copy, expected.len() - 1)?;
                published
            },
            Action::Remove => {
                if position >= expected.len() {
                    return Err(Error::InvalidSource);
                }
                expected.remove(position);
                remove(self.source, &target)?
            },
        };
        if slide_table_names(&published.package, slide)? != expected {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

/// A reversible, exact-source slide-table lifecycle patch.
///
/// The patch retains the exact source and target artifacts; it authorizes
/// application only to its source, and its inverse restores the source byte
/// for byte.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the exact artifact is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics for one slide-table lifecycle publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten or deleted native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of stale root previews deleted.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
}

/// The validated immutable result of one slide-table lifecycle edit.
#[must_use = "a table lifecycle commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    fn new(source: &Package, published: Published<Package>) -> Result<Self, Error> {
        Ok(Self {
            patch: Patch {
                artifacts: ExactArtifacts::new(owner(source)?, owner(&published.package)?),
            },
            diagnostics: Diagnostics {
                changed: true,
                touched_components: published.touched_components,
                deleted_previews: published.deleted_previews,
            },
            package: published.package,
        })
    }

    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

fn owner(package: &Package) -> Result<std::sync::Arc<[u8]>, Error> {
    let catalog = core::physical_catalog(package).map_err(|error| match error {
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        _ => Error::InvalidSource,
    })?;
    Ok(catalog.shared_source())
}

fn wire_limits(package: &Package) -> Result<WireLimits, Error> {
    package.wire_limits().map_err(|_| Error::InvalidSource)
}

/// The payload of object `identifier`'s message of `message_type`, when it
/// has exactly one; several are ambiguous.
fn message(package: &Package, identifier: u64, message_type: u32) -> Result<Option<&[u8]>, Error> {
    let object = package.object(identifier).ok_or(Error::InvalidSource)?;
    let mut found = object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type);
    let first = found.next();
    if found.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(first.map(|message| message.data.as_slice()))
}

/// The referenced identifier of the single field `number` of `payload`.
fn reference(payload: &[u8], number: u32, wire: WireLimits) -> Result<Option<u64>, Error> {
    let invalid = |_| Error::InvalidSource;
    let fields = Fields::parse(payload, wire).map_err(invalid)?;
    let Some(field) = fields.unique(number).map_err(invalid)? else {
        return Ok(None);
    };
    Ok(Some(
        wire::reference_identifier(field.message().map_err(invalid)?, wire).map_err(invalid)?,
    ))
}

/// The names of the tables in `slide`'s z-order, back to front (the host's
/// `slide_tables` order and the table selector's positions).
fn slide_table_names(package: &Package, slide: u64) -> Result<Vec<String>, Error> {
    let invalid = |_| Error::InvalidSource;
    let wire = wire_limits(package)?;
    let payload = message(package, slide, SLIDE_MESSAGE_TYPE)?.ok_or(Error::InvalidSource)?;
    let fields = Fields::parse(payload, wire).map_err(invalid)?;
    let mut names = Vec::new();
    for field in fields.all(DRAWABLES_Z_ORDER) {
        let drawable =
            wire::reference_identifier(field.message().map_err(invalid)?, wire).map_err(invalid)?;
        let Some(info) = message(package, drawable, TABLE_INFO_MESSAGE_TYPE)? else {
            continue;
        };
        let model = reference(info, TABLE_INFO_MODEL, wire)?.ok_or(Error::InvalidSource)?;
        let model =
            message(package, model, TABLE_MODEL_MESSAGE_TYPE)?.ok_or(Error::InvalidSource)?;
        let model = Fields::parse(model, wire).map_err(invalid)?;
        let name = match model.unique(TABLE_MODEL_NAME).map_err(invalid)? {
            Some(name) => std::str::from_utf8(name.message().map_err(invalid)?)
                .map_err(|_| Error::InvalidSource)?
                .to_owned(),
            None => String::new(),
        };
        names
            .try_reserve(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        names.push(name);
    }
    Ok(names)
}

/// Copy the selected table as `name` and attach the copy in front of the
/// slide's drawables; returns the publication and the copy's drawable.
fn duplicate(
    source: &Package,
    target: &core::Target,
    name: &str,
) -> Result<(Published<Package>, u64), Error> {
    let slide = target.slide.identifier;
    let mut working = Working::new(source)?;
    let mut transition = Transition::default();
    let copy = engine::clone_graph(
        &mut working,
        &mut transition,
        target.table_info.identifier,
        Drawable {
            message_types: &[TABLE_INFO_MESSAGE_TYPE],
            info_path: &[],
            offset: OFFSET,
            parent: None,
        },
        name,
    )?;

    let limits = working.archive_limits();
    let index = unique_message(working.object(slide)?, &[SLIDE_MESSAGE_TYPE])?;
    for path in [[OWNED_DRAWABLES], [DRAWABLES_Z_ORDER]] {
        let object = working.object_mut(slide)?;
        let mut after = references::identifiers(&object.messages[index].data, &path)?;
        after
            .try_reserve(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        after.push(copy.info);
        let data = references::rewrite(&object.messages[index].data, &path, &after)?;
        references::replace(object, index, data, &[], &[copy.info], &path, limits)?;
    }

    // The slide's and the engine's references to the copy across components.
    if let Some(registry) = metadata::registry(&working)? {
        let info_component = registry
            .component_for_entry(&copy.info_component)?
            .identifier;
        let slide_member = working.locate(slide)?.ok_or(Fault::Invalid)?;
        let slide_component = registry.component_for_entry(&slide_member)?.identifier;
        transition.reference(&registry, slide_component, info_component, copy.info)?;
        if let Some(component) = &copy.engine_component {
            let engine_component = registry.component_for_entry(component)?.identifier;
            transition.reference(&registry, engine_component, info_component, copy.info)?;
        }
    }
    Ok((working.publish(&transition)?, copy.info))
}

/// Prove the copy is the selectable table at `position` with its source's
/// dimensions.
fn verify_copy(
    candidate: &Package,
    source: &core::Target,
    copy: u64,
    position: usize,
) -> Result<(), Error> {
    let mut budget = core::Budget::new(candidate).map_err(select_error)?;
    let selected = core::select_table(
        candidate,
        SlideSelector::index(source.slide_position.get()),
        TableSelector::index(position),
        &mut budget,
    )
    .map_err(select_error)?;
    if selected.table_info.identifier != copy
        || (selected.rows, selected.columns) != (source.rows, source.columns)
    {
        return Err(Error::Verification);
    }
    Ok(())
}

/// Remove the selected table's private graph and formula owners and detach
/// it from its slide.
fn remove(source: &Package, target: &core::Target) -> Result<Published<Package>, Error> {
    let slide = target.slide.identifier;
    let info = target.table_info.identifier;
    let mut working = Working::new(source)?;
    let census = Census::new(&working)?;
    let closure = census.private_closure(&working, &[info])?;

    // The CalculationEngine owners of the removed contexts.
    let index = DeletionObjectIndex::new(source)?;
    let mut graph = GraphPlan {
        index,
        removals: RemovalPlan::default(),
    };
    reserve(&mut graph.removals.object_removals, closure.len())?;
    reserve(&mut graph.removals.formula_contexts, closure.len())?;
    for identifier in &closure {
        let nonzero = NonZeroU64::new(*identifier).ok_or(Error::InvalidSource)?;
        let location = graph.index.get(nonzero).ok_or(Error::InvalidSource)?;
        graph.removals.object_removals.push(ObjectRemoval {
            component_index: location.component_index,
            object_index: location.object_index,
            identifier: nonzero,
        });
        graph.removals.formula_contexts.push(nonzero);
    }
    let mut budget = WireBudget::new(source.source_bytes().len());
    let formula = formula::prepare(source, &graph, &mut budget)?;
    let mut removed: BTreeSet<u64> = closure;
    for removal in &formula.removals.object_removals {
        removed.insert(removal.identifier.get());
    }
    let mut detached = Vec::new();
    reserve(&mut detached, 1)?;
    detached.push((slide, info));
    for edit in &formula.removals.message_edits {
        reserve(&mut detached, edit.remove_object_references.len())?;
        for referenced in &edit.remove_object_references {
            detached.push((edit.object_identifier.get(), referenced.get()));
        }
    }
    census.require_detached(&removed, &detached)?;

    detach(&mut working, wire_limits(source)?, slide, info)?;
    // The CalculationEngine edits.
    let limits = working.archive_limits();
    for edit in formula.removals.message_edits {
        if !edit.remove_data_references.is_empty()
            || removed.contains(&edit.object_identifier.get())
        {
            return Err(Error::InvalidSource);
        }
        let object = working.object_mut(edit.object_identifier.get())?;
        if object
            .messages
            .get(edit.message_index)
            .map(|message| message.type_)
            != Some(edit.message_type)
        {
            return Err(Error::InvalidSource);
        }
        let mut pruned = Vec::new();
        reserve(&mut pruned, edit.remove_object_references.len())?;
        for identifier in &edit.remove_object_references {
            pruned.push(identifier.get());
        }
        references::replace(
            object,
            edit.message_index,
            edit.data,
            &pruned,
            &[],
            &[],
            limits,
        )?;
    }
    let mut removed_list = Vec::new();
    reserve(&mut removed_list, removed.len())?;
    removed_list.extend(removed.iter().copied());
    let mut transition = Transition::default();
    if let Some(registry) = metadata::registry(&working)? {
        transition.remove_objects(&registry, &removed_list)?;
    }
    // A removed suffix of the identifier space lowers the watermark, as the
    // host editor releases it (record 0370g).
    transition.release_suffix(&removed_list)?;
    for identifier in &removed_list {
        working.remove(*identifier)?;
    }
    let published = working.publish(&transition)?;
    verify_removed(&published.package, &removed)?;
    Ok(published)
}

/// Whether message `entry` references `info` in field `number`.
fn names_info(entry: &[u8], number: u32, info: u64, wire: WireLimits) -> Result<bool, Error> {
    Ok(reference(entry, number, wire)? == Some(info))
}

/// The entries of repeated message field `number` that do not reference
/// `info` in their field `entry_number` (`None`: the entry is a reference),
/// and how many did.
fn kept_entries(
    fields: &Fields,
    number: u32,
    entry_number: Option<u32>,
    info: u64,
    wire: WireLimits,
) -> Result<(Vec<Vec<u8>>, usize), Error> {
    let invalid = |_| Error::InvalidSource;
    let mut kept = Vec::new();
    let mut dropped = 0;
    for field in fields.all(number) {
        let entry = field.message().map_err(invalid)?;
        let names = match entry_number {
            Some(entry_number) => names_info(entry, entry_number, info, wire)?,
            None => wire::reference_identifier(entry, wire).map_err(invalid)? == info,
        };
        if names {
            dropped += 1;
        } else {
            kept.try_reserve(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
            kept.push(copy_bytes(entry).map_err(Error::from)?);
        }
    }
    Ok((kept, dropped))
}

/// Detach `info` from the slide's owned drawables and z-order (each must
/// list it once) and drop the template-tag and instructional-text entries
/// that name it (host `patch_slide_drawable_references`).
fn detach(
    working: &mut Working<'_, Package>,
    wire: WireLimits,
    slide: u64,
    info: u64,
) -> Result<(), Error> {
    let invalid = |_| Error::InvalidSource;
    let limits = working.archive_limits();
    let object = working.object(slide)?;
    let index = unique_message(object, &[SLIDE_MESSAGE_TYPE])?;
    let mut fields = Fields::parse(&object.messages[index].data, wire).map_err(invalid)?;
    for number in [OWNED_DRAWABLES, DRAWABLES_Z_ORDER] {
        let (kept, dropped) = kept_entries(&fields, number, None, info, wire)?;
        if dropped != 1 {
            return Err(Error::InvalidSource);
        }
        fields.replace_all(number, LENGTH_DELIMITED, &kept);
    }
    let (kept, dropped) =
        kept_entries(&fields, TEMPLATE_TAGS, Some(TEMPLATE_TAG_INFO), info, wire)?;
    if dropped != 0 {
        fields.replace_all(TEMPLATE_TAGS, LENGTH_DELIMITED, &kept);
    }
    if let Some(map) = fields.unique(INSTRUCTIONAL_TEXT).map_err(invalid)? {
        let mut map = Fields::parse(map.message().map_err(invalid)?, wire).map_err(invalid)?;
        let (kept, dropped) = kept_entries(
            &map,
            INSTRUCTIONAL_TEXT_ENTRIES,
            Some(INSTRUCTIONAL_TEXT_INFO),
            info,
            wire,
        )?;
        if dropped != 0 {
            map.replace_all(INSTRUCTIONAL_TEXT_ENTRIES, LENGTH_DELIMITED, &kept);
            fields.set_message(INSTRUCTIONAL_TEXT, &map.encode());
        }
    }
    let data = fields.encode();
    let object = working.object_mut(slide)?;
    references::replace(object, index, data, &[info], &[], &[], limits)?;
    Ok(())
}

impl Package {
    /// Start an immutable single-operation slide-table duplication or
    /// removal.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::table::TableSelector;
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let commit = package
    ///     .edit_slide_table_lifecycle()
    ///     .duplicate(SlideSelector::index(0), TableSelector::index(0))?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_slide_table_lifecycle(&commit.patch().inverse())?;
    /// assert!(restored.diagnostics().changed());
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_slide_table_lifecycle(&self) -> Edit<'_> {
        Edit {
            source: self,
            staged: None,
        }
    }

    /// Replay an exact-source slide-table lifecycle patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact, and [`Error::Verification`] when its target does
    /// not reopen.
    pub fn apply_slide_table_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_source(&owner(self)?) {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.snapshot()
        } else {
            self.structure_reopen_shared(patch.artifacts.target())
                .map_err(reopen_error)?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                touched_components: 0,
                deleted_previews: 0,
            },
        })
    }
}
