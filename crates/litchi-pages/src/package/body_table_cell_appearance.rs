//! Cell fill and text layout of rooted body tables (record 0371; the
//! migration host's `table_cell_fill`, `table_cell_layout` and their set and
//! reset routes).
//!
//! The format-neutral cell-style engine
//! (`litchi_iwa_table::cell_appearance`) plans each change; this owner
//! resolves the selector, allocates the new variation's identifier, applies
//! the plan to the parsed components and the package metadata, publishes and
//! reopens. Each staged operation is applied to the package the previous one
//! produced, as the host applies them one after another.

use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_archive::{ComponentCatalog, SourceCatalog};
use litchi_iwa_core::{Archive, RawMessage};
use litchi_iwa_table::cell_appearance::{self as engine, Kind, Plan, Property, Registration};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{self as graph, GraphError, GraphLimit};
use super::body_storage_objects as objects;
use crate::selector::BodyTableSelector;
use crate::table::cell_appearance::{Appearance, Fill, Layout};

const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

/// A finite resource charged by one cell-appearance edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected.
    InputBytes,
    /// Bytes in the rewritten payloads or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selector resolved no body table.
    #[error("the Pages body has no such table")]
    TableNotFound,
    /// The cell lies outside the table.
    #[error("the cell lies outside the table")]
    OutsideTable,
    /// The cell's row has no stored tile.
    #[error("the cell's row has no stored tile")]
    MissingTile,
    /// The table's style list is segmented or holds other entries.
    #[error("the table's style list is not supported")]
    UnsupportedStyleList,
    /// Image fills are not supported by this family.
    #[error("image cell fills are not supported")]
    ImageFill,
    /// The cell record is legacy or malformed.
    #[error("the cell record is not supported")]
    UnsupportedCell,
    /// Nothing was staged.
    #[error("no cell appearance change is staged")]
    NoStagedOperation,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact cell appearance edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound => RefusalKind::NotFound,
            Self::OutsideTable => RefusalKind::InvalidValue,
            Self::MissingTile
            | Self::UnsupportedStyleList
            | Self::ImageFill
            | Self::UnsupportedCell
            | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
}

/// Failure from a cell-appearance read or edit.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The table's style graph is malformed.
    #[error("the selected Pages table's cell styles are invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error("Pages cell appearance {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for a Pages cell appearance edit")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the new appearance.
    #[error("the edited Pages cell failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages cell appearance patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                },
                observed,
                maximum,
            },
        }
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        match error {
            engine::Error::Refused(refusal) => Self::Refused(match refusal {
                engine::Refusal::OutsideTable => Refusal::OutsideTable,
                engine::Refusal::MissingTile => Refusal::MissingTile,
                engine::Refusal::UnsupportedStyleList => Refusal::UnsupportedStyleList,
                engine::Refusal::ImageFill => Refusal::ImageFill,
                _ => Refusal::UnsupportedCell,
            }),
            _ => Self::InvalidSource,
        }
    }
}

/// The package view the shared engine reads.
struct Source<'a> {
    package: &'a Package,
    index: HashMap<u64, (usize, usize)>,
}

impl<'a> Source<'a> {
    fn new(package: &'a Package) -> Result<Self, Error> {
        let mut index = HashMap::new();
        index
            .try_reserve(package.state.object_count)
            .map_err(|_error| Error::Allocation {
                amount: package.state.object_count,
            })?;
        for (component_index, component) in package.state.source.components().iter().enumerate() {
            for (object_index, object) in component.archive().objects.iter().enumerate() {
                if let Some(identifier) = object.archive_info.identifier
                    && index
                        .insert(identifier, (component_index, object_index))
                        .is_some()
                {
                    return Err(Error::InvalidSource);
                }
            }
        }
        Ok(Self { package, index })
    }
}

impl litchi_iwa_table::TableSource for Source<'_> {
    fn catalog(&self) -> &ComponentCatalog {
        self.package.state.source.components()
    }

    fn physical(&self) -> Option<&SourceCatalog> {
        Some(&self.package.state.source)
    }

    fn resolve_object(
        &self,
        identifier: u64,
    ) -> Result<Option<litchi_iwa_table::ResolvedObject<'_>>, litchi_iwa_table::ResolveError> {
        let Some(&(component_index, object_index)) = self.index.get(&identifier) else {
            return Ok(None);
        };
        let object = self
            .package
            .state
            .source
            .components()
            .get_index(component_index)
            .and_then(|component| component.archive().objects.get(object_index))
            .ok_or(litchi_iwa_table::ResolveError)?;
        Ok(Some(litchi_iwa_table::ResolvedObject {
            messages: &object.messages,
            component_index,
            object_index,
        }))
    }

    fn source_bytes(&self) -> &[u8] {
        self.package.state.source.source_bytes()
    }

    fn archive_limits(&self) -> litchi_iwa_archive::Limits {
        self.package.state.source.limits()
    }

    fn max_objects(&self) -> usize {
        self.package.state.object_count
    }

    fn max_references(&self) -> usize {
        usize::MAX
    }
}

/// One staged change.
#[derive(Debug, Clone, PartialEq)]
enum Operation {
    SetFill(u32, u32, Fill),
    ResetFill(u32, u32),
    SetLayout(u32, u32, Layout),
    ResetLayout(u32, u32),
}

/// Cell appearance changes staged against one body table.
pub struct Edit<'a> {
    source: &'a Package,
    table: BodyTableSelector<'static>,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage the cell's own fill (an image fill is refused).
    ///
    /// # Errors
    ///
    /// [`Refusal::ImageFill`] for an image fill.
    pub fn set_fill(&mut self, row: u32, column: u32, fill: Fill) -> Result<&mut Self, Error> {
        if matches!(fill, Fill::Image(_)) {
            return Err(Refusal::ImageFill.into());
        }
        self.operations.push(Operation::SetFill(row, column, fill));
        Ok(self)
    }

    /// Stage removing the cell's own fill, so it shows its style's fill.
    pub fn reset_fill(&mut self, row: u32, column: u32) -> &mut Self {
        self.operations.push(Operation::ResetFill(row, column));
        self
    }

    /// Stage the cell's text wrapping, vertical alignment and insets.
    pub fn set_layout(&mut self, row: u32, column: u32, layout: Layout) -> &mut Self {
        self.operations
            .push(Operation::SetLayout(row, column, layout));
        self
    }

    /// Stage removing the cell's own text-layout overrides.
    pub fn reset_layout(&mut self, row: u32, column: u32) -> &mut Self {
        self.operations.push(Operation::ResetLayout(row, column));
        self
    }

    /// Apply every staged change in order, publish and reopen.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] without a change; the engine's typed
    /// refusals; [`Error::Verification`] when the reopened cell does not show
    /// the change.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut package = self.source.clone();
        let mut changes = 0;
        for operation in &self.operations {
            let model = model_identifier(&package, self.table)?;
            let (row, column) = match operation {
                Operation::SetFill(row, column, _)
                | Operation::ResetFill(row, column)
                | Operation::SetLayout(row, column, _)
                | Operation::ResetLayout(row, column) => (*row, *column),
            };
            let target = engine::Target {
                model_identifier: model,
                row,
                column,
            };
            match operation {
                Operation::SetFill(_, _, fill) => {
                    if let Some(next) = set(&package, target, &[Property::Fill(fill.clone())])? {
                        package = next;
                        changes += 1;
                    }
                },
                Operation::SetLayout(_, _, layout) => {
                    let properties = [
                        Property::TextWrap(layout.text_wrap()),
                        Property::VerticalAlignment(layout.vertical_alignment()),
                        Property::Padding(layout.insets()),
                    ];
                    if let Some(next) = set(&package, target, &properties)? {
                        package = next;
                        changes += 1;
                    }
                },
                Operation::ResetFill(..) => {
                    if let Some(next) = reset(&package, target, Kind::Fill)? {
                        package = next;
                        changes += 1;
                    }
                },
                Operation::ResetLayout(..) => {
                    let mut any = false;
                    for kind in [Kind::TextWrap, Kind::VerticalAlignment, Kind::Padding] {
                        let target = engine::Target {
                            model_identifier: model_identifier(&package, self.table)?,
                            ..target
                        };
                        if let Some(next) = reset(&package, target, kind)? {
                            package = next;
                            any = true;
                        }
                    }
                    changes += usize::from(any);
                },
            }
        }
        let target = package.state.source.shared_source();
        let changed = !Arc::ptr_eq(&target, &self.source.state.source.shared_source());
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(self.source.state.source.shared_source(), target),
            },
            diagnostics: Diagnostics {
                changed,
                applied: changes,
            },
        })
    }
}

/// Exact-source reversible cell-appearance patch.
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

/// Content-free diagnostics of one cell-appearance edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    applied: usize,
}

impl Diagnostics {
    /// Whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// How many staged changes changed the package.
    #[must_use]
    pub const fn applied(self) -> usize {
        self.applied
    }
}

/// The validated immutable result of one cell-appearance edit.
#[must_use = "a cell appearance commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
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

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

fn model_identifier(package: &Package, table: BodyTableSelector<'_>) -> Result<u64, Error> {
    let mut budget = super::table_lock::WireBudget::new(package.state.source.limits())
        .map_err(|_error| Error::InvalidSource)?;
    package
        .resolve_body_table_with_budget(table, &mut budget)
        .map(|target| target.model_identifier.get())
        .map_err(|_error| Error::Refused(Refusal::TableNotFound))
}

fn limits(package: &Package) -> Result<litchi_numbers_wire::Limits, Error> {
    package
        .numbers_wire_limits()
        .map_err(|_error| Error::InvalidSource)
}

fn fresh_identifier(package: &Package) -> Result<u64, Error> {
    let metadata = objects::metadata(package)?;
    Ok(objects::next_identifier(package, metadata.as_ref())?)
}

fn set(
    package: &Package,
    target: engine::Target,
    properties: &[Property],
) -> Result<Option<Package>, Error> {
    let limits = limits(package)?;
    let fresh = fresh_identifier(package)?;
    let plan = {
        let source = Source::new(package)?;
        engine::plan_set(&source, target, properties, fresh, &limits)?
    };
    let Some(plan) = plan else {
        return Ok(None);
    };
    let candidate = publish(package, &plan)?;
    let source = Source::new(&candidate)?;
    for property in properties {
        let shown = engine::value(&source, target, property.kind(), &limits)?;
        if &shown != property {
            return Err(Error::Verification);
        }
    }
    Ok(Some(candidate))
}

fn reset(package: &Package, target: engine::Target, kind: Kind) -> Result<Option<Package>, Error> {
    let limits = limits(package)?;
    let fresh = fresh_identifier(package)?;
    let plan = {
        let source = Source::new(package)?;
        engine::plan_reset(&source, target, kind, fresh, &limits)?
    };
    let Some(plan) = plan else {
        return Ok(None);
    };
    Ok(Some(publish(package, &plan)?))
}

fn editable<'a>(
    source: &Package,
    archives: &'a mut Vec<(Arc<str>, Archive)>,
    component: &str,
) -> Result<&'a mut Archive, Error> {
    let position = match archives
        .iter()
        .position(|(name, _)| name.as_ref() == component)
    {
        Some(position) => position,
        None => {
            archives.push((
                Arc::from(component),
                graph::editable_archive(source, component)?,
            ));
            archives.len() - 1
        },
    };
    Ok(&mut archives[position].1)
}

/// The component (member) name holding `identifier` in the source.
fn component_name(package: &Package, identifier: u64) -> Result<Arc<str>, Error> {
    package
        .state
        .source
        .components()
        .iter()
        .find(|component| component.archive().object(identifier).is_some())
        .map(|component| Arc::from(component.name()))
        .ok_or(Error::InvalidSource)
}

/// Apply one plan to the parsed components and the package metadata.
fn publish(source: &Package, plan: &Plan) -> Result<Package, Error> {
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_error| Error::InvalidSource)?;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    for rewrite in &plan.rewrites {
        let component = component_name(source, rewrite.identifier)?;
        let archive = editable(source, &mut archives, &component)?;
        let object = archive
            .object_mut(rewrite.identifier)
            .ok_or(Error::InvalidSource)?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == rewrite.message_type)
            .ok_or(Error::InvalidSource)?;
        object
            .replace_message_preserving_header(
                index,
                RawMessage {
                    type_: rewrite.message_type,
                    data: rewrite.data.clone(),
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
        let info = &mut object.archive_info.message_infos[index];
        for reference in &rewrite.add_references {
            if !info.object_references.contains(reference) {
                info.object_references.push(*reference);
            }
        }
        for reference in &rewrite.remove_references {
            info.object_references
                .retain(|current| current != reference);
            for field in &mut info.field_infos {
                field
                    .object_references
                    .retain(|current| current != reference);
            }
        }
    }
    let mut style_component = None;
    if let Some(insertion) = &plan.insertion {
        let component = component_name(source, insertion.stylesheet)?;
        let mut object = objects::new_object(
            source,
            insertion.identifier,
            CELL_STYLE_MESSAGE_TYPE,
            insertion.data.clone(),
        )?;
        let info = &mut object.archive_info.message_infos[0];
        info.versions = engine::CELL_STYLE_VERSIONS.to_vec().into();
        info.object_references
            .extend(insertion.object_references.iter().copied());
        editable(source, &mut archives, &component)?
            .insert_object(object, archive_limits)
            .map_err(graph::map_core_error)?;
        style_component = Some(component);
    }
    let mut removed_component = None;
    if let Some(removed) = plan.removal {
        let component = component_name(source, removed)?;
        editable(source, &mut archives, &component)?
            .remove_object(removed)
            .ok_or(Error::InvalidSource)?;
        removed_component = Some(component);
    }

    if !plan.registrations.is_empty()
        && let Some(metadata) = objects::metadata(source)?
    {
        let mut payload = metadata.payload.clone();
        // A UUID registration raises the watermark with it (the registry
        // batch admits new objects up to the new watermark only).
        let watermark = plan
            .registrations
            .iter()
            .find_map(|registration| match registration {
                Registration::Watermark(identifier) => Some(*identifier),
                _ => None,
            })
            .unwrap_or(metadata.last)
            .max(metadata.last);
        let registers_uuid = plan
            .registrations
            .iter()
            .any(|registration| matches!(registration, Registration::Uuid { .. }));
        for registration in &plan.registrations {
            match *registration {
                Registration::Uuid { object } => {
                    let component = style_component.as_deref().ok_or(Error::InvalidSource)?;
                    if let Some(component) = metadata.component_for_entry(component)? {
                        let staged = objects::Metadata::with_payload(&metadata, payload);
                        payload = staged.registered(source, component, &[object], watermark)?;
                    }
                },
                Registration::ExternalReference { from, object } => {
                    let from = component_name(source, from)?;
                    let target = style_component.as_deref().ok_or(Error::InvalidSource)?;
                    if let (Some(from), Some(target)) = (
                        metadata.component_for_entry(&from)?,
                        metadata.component_for_entry(target)?,
                    ) && from != target
                    {
                        payload = objects::add_external_reference(&payload, from, target, object)?;
                    }
                },
                Registration::Watermark(_) => {
                    if !registers_uuid {
                        let staged = objects::Metadata::with_payload(&metadata, payload);
                        payload = staged.raised(source, watermark)?;
                    }
                },
                Registration::Removed { object } => {
                    let component = removed_component.as_deref().ok_or(Error::InvalidSource)?;
                    if let Some(component) = metadata.component_for_entry(component)? {
                        if metadata.uuid_owner(object)? == Some(component) {
                            payload = objects::remove_object_uuid(&payload, component, object)?;
                        }
                        payload = objects::remove_external_references(&payload, component, object)?;
                    }
                    let remaining = objects::max_identifier(source, &archives);
                    payload = objects::released(&payload, metadata.last, &[object], remaining)?;
                },
            }
        }
        if payload != metadata.payload {
            editable(source, &mut archives, &metadata.component)?
                .object_mut(metadata.object)
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header(
                    metadata.message_index,
                    RawMessage {
                        type_: METADATA_MESSAGE_TYPE,
                        data: payload,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
        }
    }
    Ok(graph::publish_changes(source, archives, &[], &[], &[])?)
}

fn appearance(package: &Package, model: u64, row: u32, column: u32) -> Result<Appearance, Error> {
    let limits = limits(package)?;
    let source = Source::new(package)?;
    let target = engine::Target {
        model_identifier: model,
        row,
        column,
    };
    let value = |kind| engine::value(&source, target, kind, &limits);
    let Property::Fill(fill) = value(Kind::Fill)? else {
        return Err(Error::InvalidSource);
    };
    let Property::TextWrap(text_wrap) = value(Kind::TextWrap)? else {
        return Err(Error::InvalidSource);
    };
    let Property::VerticalAlignment(vertical_alignment) = value(Kind::VerticalAlignment)? else {
        return Err(Error::InvalidSource);
    };
    let Property::Padding(insets) = value(Kind::Padding)? else {
        return Err(Error::InvalidSource);
    };
    Ok(Appearance {
        fill,
        layout: Layout::new(text_wrap, vertical_alignment, insets),
    })
}

impl Package {
    /// Read one body-table cell's effective fill and text layout.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::OutsideTable`] and the typed
    /// refusals of an unsupported style graph.
    pub fn body_table_cell_appearance<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        row: u32,
        column: u32,
    ) -> Result<Appearance, Error> {
        let model = model_identifier(self, selector.into())?;
        appearance(self, model, row, column)
    }

    /// Begin staging cell fill and text-layout changes on one body table.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] when the selector resolves no table.
    pub fn edit_body_table_cell_appearance<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        model_identifier(self, selector)?;
        let catalog = self.body_tables().map_err(|_error| Error::InvalidSource)?;
        let position = match selector {
            BodyTableSelector::Position(position) => position.get(),
            BodyTableSelector::Name(name) => (0..catalog.len())
                .find(|index| {
                    catalog
                        .get(*index)
                        .is_some_and(|entry| entry.name() == name)
                })
                .ok_or(Refusal::TableNotFound)?,
        };
        Ok(Edit {
            source: self,
            table: BodyTableSelector::index(position),
            operations: Vec::new(),
        })
    }

    /// Replay an exact-source cell-appearance patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact.
    pub fn apply_body_table_cell_appearance(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            graph::reopen(self, patch.artifacts.target())?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                applied: usize::from(!patch.is_noop()),
            },
        })
    }
}
