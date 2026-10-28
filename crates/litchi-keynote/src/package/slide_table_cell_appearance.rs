//! Cell fill, text layout and borders of slide-table cells (record 0372;
//! the Keynote leaf of the host's `*_slide_table_cell_{fill,layout}`,
//! `slide_table_cell_borders`, `set_slide_table_cell_border` and
//! `clear_slide_table_cell_border`) over the shared cell-appearance engine
//! (`litchi_iwa_table::cell_appearance`).
//!
//! The engine plans each change; this owner resolves the slide-table
//! selector, supplies the next object identifier, applies the plan to the
//! structural workspace (payload rewrites, the new or removed style object
//! and the registry), publishes and reopens. Each staged operation is
//! applied to the package the previous one produced, as the host applies
//! them one after another.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_table::cell_appearance::{self as engine, Kind, Plan, Property, Registration};
use thiserror::Error as ThisError;

use super::structure::metadata::ExternalReference;
use super::structure::workspace::Workspace;
use super::{Package, slide_table_core as core};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::cell_appearance::{Appearance, BorderSide, Fill, Layout, Stroke};

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selectors resolved no slide table.
    #[error("the Keynote slide has no such table")]
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
            | Self::UnsupportedCell => RefusalKind::UnsupportedContent,
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
    /// The table's style graph is malformed or the source is not exact.
    #[error("the selected Keynote table's cell styles are invalid")]
    InvalidSource,
    /// Reopening the candidate did not reproduce the new appearance.
    #[error("the edited Keynote table cell failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Keynote cell appearance patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
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

/// One staged change.
#[derive(Debug, Clone, PartialEq)]
enum Operation {
    SetFill(u32, u32, Fill),
    ResetFill(u32, u32),
    SetLayout(u32, u32, Layout),
    ResetLayout(u32, u32),
    Border(u32, u32, BorderSide, Option<Stroke>),
}

/// Cell appearance changes staged against one slide table.
pub struct Edit<'a> {
    source: &'a Package,
    slide: usize,
    table: usize,
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

    /// Stage an explicit stroke on one edge of the cell.
    pub fn set_border(
        &mut self,
        row: u32,
        column: u32,
        side: BorderSide,
        stroke: Stroke,
    ) -> &mut Self {
        self.operations
            .push(Operation::Border(row, column, side, Some(stroke)));
        self
    }

    /// Stage an explicitly cleared edge (no stroke, overriding the table
    /// style), as the host's `clear_slide_table_cell_border` writes it.
    pub fn clear_border(&mut self, row: u32, column: u32, side: BorderSide) -> &mut Self {
        self.operations
            .push(Operation::Border(row, column, side, None));
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
        let mut package = self.source.snapshot();
        let mut applied = 0;
        // Every intermediate candidate advances the source's save token, so
        // the transaction advances it once.
        let origin = Workspace::new(self.source)
            .map_err(|_| Error::InvalidSource)?
            .registry
            .save_token();
        for operation in &self.operations {
            let (row, column) = match operation {
                Operation::SetFill(row, column, _)
                | Operation::ResetFill(row, column)
                | Operation::SetLayout(row, column, _)
                | Operation::ResetLayout(row, column)
                | Operation::Border(row, column, ..) => (*row, *column),
            };
            let target = target(&package, self.slide, self.table, row, column)?;
            match operation {
                Operation::SetFill(_, _, fill) => {
                    if let Some(next) =
                        set(&package, target, &[Property::Fill(fill.clone())], origin)?
                    {
                        package = next;
                        applied += 1;
                    }
                },
                Operation::SetLayout(_, _, layout) => {
                    let properties = [
                        Property::TextWrap(layout.text_wrap()),
                        Property::VerticalAlignment(layout.vertical_alignment()),
                        Property::Padding(layout.insets()),
                    ];
                    if let Some(next) = set(&package, target, &properties, origin)? {
                        package = next;
                        applied += 1;
                    }
                },
                Operation::Border(_, _, side, stroke) => {
                    package = border(&package, target, *side, *stroke, origin)?;
                    applied += 1;
                },
                Operation::ResetFill(..) => {
                    if let Some(next) = reset(&package, target, Kind::Fill, origin)? {
                        package = next;
                        applied += 1;
                    }
                },
                Operation::ResetLayout(..) => {
                    let mut any = false;
                    for kind in [Kind::TextWrap, Kind::VerticalAlignment, Kind::Padding] {
                        let target = self::target(&package, self.slide, self.table, row, column)?;
                        if let Some(next) = reset(&package, target, kind, origin)? {
                            package = next;
                            any = true;
                        }
                    }
                    applied += usize::from(any);
                },
            }
        }
        let source = owner(self.source)?;
        let target = owner(&package)?;
        let changed = !Arc::ptr_eq(&source, &target);
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(source, target),
            },
            diagnostics: Diagnostics { changed, applied },
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

fn owner(package: &Package) -> Result<Arc<[u8]>, Error> {
    Ok(core::physical_catalog(package)
        .map_err(|_| Error::InvalidSource)?
        .shared_source())
}

/// The selected table's model and position: (slide, table, model).
fn select(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
) -> Result<(usize, usize, u64), Error> {
    let mut budget = core::Budget::new(package).map_err(|_| Error::InvalidSource)?;
    let selected = core::select_table(package, slide, table, &mut budget)
        .map_err(|_| Error::Refused(Refusal::TableNotFound))?;
    Ok((
        selected.slide_position.get(),
        selected.table_position.get(),
        selected.model.identifier,
    ))
}

fn target(
    package: &Package,
    slide: usize,
    table: usize,
    row: u32,
    column: u32,
) -> Result<engine::Target, Error> {
    let (_, _, model) = select(
        package,
        SlideSelector::index(slide),
        TableSelector::index(table),
    )?;
    Ok(engine::Target {
        model_identifier: model,
        row,
        column,
    })
}

fn limits(package: &Package) -> Result<litchi_numbers_wire::Limits, Error> {
    package
        .numbers_wire_limits()
        .map_err(|_| Error::InvalidSource)
}

fn fresh(package: &Package) -> Result<u64, Error> {
    Ok(Workspace::new(package)
        .map_err(|_| Error::InvalidSource)?
        .next_identifier())
}

fn set(
    package: &Package,
    target: engine::Target,
    properties: &[Property],
    origin: Option<u64>,
) -> Result<Option<Package>, Error> {
    let limits = limits(package)?;
    let Some(plan) = engine::plan_set(package, target, properties, fresh(package)?, &limits)?
    else {
        return Ok(None);
    };
    let candidate = publish(package, &plan, origin)?;
    for property in properties {
        if &engine::value(&candidate, target, property.kind(), &limits)? != property {
            return Err(Error::Verification);
        }
    }
    Ok(Some(candidate))
}

fn reset(
    package: &Package,
    target: engine::Target,
    kind: Kind,
    origin: Option<u64>,
) -> Result<Option<Package>, Error> {
    let limits = limits(package)?;
    let Some(plan) = engine::plan_reset(package, target, kind, fresh(package)?, &limits)? else {
        return Ok(None);
    };
    Ok(Some(publish(package, &plan, origin)?))
}

fn border(
    package: &Package,
    target: engine::Target,
    side: BorderSide,
    stroke: Option<Stroke>,
    origin: Option<u64>,
) -> Result<Package, Error> {
    let plan = engine::plan_border(package, target, side, stroke, fresh(package)?)?;
    let candidate = publish(package, &plan, origin)?;
    if engine::borders(&candidate, target)?.get(side) != stroke {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

fn component_id(workspace: &Workspace<'_>, object: u64) -> Result<u64, Error> {
    let member = workspace.component_of(object).ok_or(Error::InvalidSource)?;
    Ok(workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier)
}

/// Apply one plan on the structural workspace, publish and reopen. The
/// save token advances from `origin`'s, once per transaction.
fn publish(source: &Package, plan: &Plan, origin: Option<u64>) -> Result<Package, Error> {
    let invalid = |_| Error::InvalidSource;
    let mut workspace = Workspace::new(source).map_err(invalid)?;
    if let Some(origin) = origin {
        workspace.registry.advance_from(Some(origin));
    }
    let archive_limits = workspace.archive_limits();
    for rewrite in &plan.rewrites {
        let object = workspace.object_mut(rewrite.identifier).map_err(invalid)?;
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
            .map_err(|_| Error::InvalidSource)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Error::InvalidSource)?;
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
    for field in &plan.field_references {
        let object = workspace.object_mut(field.object).map_err(invalid)?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == field.message_type)
            .ok_or(Error::InvalidSource)?;
        for record in &mut object
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Error::InvalidSource)?
            .field_infos
        {
            if record.object_references.contains(&field.sibling)
                && !record.object_references.contains(&field.reference)
            {
                record.object_references.push(field.reference);
            }
        }
    }
    let mut inserted = None;
    if let Some(insertion) = &plan.insertion {
        if workspace.allocate().map_err(invalid)? != insertion.identifier {
            return Err(Error::InvalidSource);
        }
        let member = workspace
            .component_of(insertion.beside)
            .ok_or(Error::InvalidSource)?;
        let mut object = ArchiveObject::new(
            insertion.identifier,
            vec![RawMessage {
                type_: insertion.message_type,
                data: insertion.data.clone(),
            }],
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or(Error::InvalidSource)?;
        if let Some(versions) = insertion.versions {
            info.versions = versions.to_vec().into();
        }
        info.object_references
            .extend(insertion.object_references.iter().copied());
        workspace.insert(&member, object).map_err(invalid)?;
        inserted = Some(insertion.identifier);
    }
    let removed_component = match plan.removal {
        Some(removed) => {
            let component = component_id(&workspace, removed)?;
            workspace.remove_object(removed).map_err(invalid)?;
            Some(component)
        },
        None => None,
    };
    for registration in &plan.registrations {
        match *registration {
            Registration::Uuid { object } => {
                if Some(object) != inserted {
                    return Err(Error::InvalidSource);
                }
                let component = component_id(&workspace, object)?;
                let uuid = workspace.clone_uuid(None, object).map_err(invalid)?;
                workspace
                    .registry
                    .component_mut(component)
                    .ok_or(Error::InvalidSource)?
                    .add_object_uuid(object, uuid);
            },
            Registration::ExternalReference { from, object } => {
                let from = component_id(&workspace, from)?;
                let to = component_id(&workspace, object)?;
                if from != to {
                    workspace
                        .registry
                        .component_mut(from)
                        .ok_or(Error::InvalidSource)?
                        .add_external_reference(ExternalReference {
                            component: to,
                            object: Some(object),
                            is_weak: None,
                        });
                }
            },
            // The identifier allocation raised the watermark.
            Registration::Watermark(_) => {},
            Registration::Removed { object } => {
                let component = removed_component.ok_or(Error::InvalidSource)?;
                workspace
                    .registry
                    .component_mut(component)
                    .ok_or(Error::InvalidSource)?
                    .remove_object_uuids(&[object]);
                let referrers = workspace
                    .registry
                    .components()
                    .filter(|entry| {
                        entry.external_references().iter().any(|reference| {
                            reference.component == component && reference.object == Some(object)
                        })
                    })
                    .map(|entry| entry.identifier)
                    .collect::<Vec<_>>();
                for referrer in referrers {
                    let entry = workspace
                        .registry
                        .component_mut(referrer)
                        .ok_or(Error::InvalidSource)?;
                    let kept = entry
                        .external_references()
                        .iter()
                        .filter(|reference| {
                            !(reference.component == component && reference.object == Some(object))
                        })
                        .copied()
                        .collect();
                    entry.set_external_references(kept);
                }
                workspace.release_removed_identifier_suffix();
            },
        }
    }
    let bytes = workspace.publish().map_err(invalid)?;
    source
        .structure_reopen(bytes)
        .map_err(|_| Error::Verification)
}

fn appearance(package: &Package, target: engine::Target) -> Result<Appearance, Error> {
    let limits = limits(package)?;
    let value = |kind| engine::value(package, target, kind, &limits);
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
        borders: engine::borders(package, target)?,
    })
}

impl Package {
    /// Read one slide-table cell's effective fill, text layout and explicit
    /// edge strokes (host `slide_table_cell_fill`, `slide_table_cell_layout`,
    /// `slide_table_cell_borders`).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::OutsideTable`] and the typed
    /// refusals of an unsupported style graph.
    pub fn slide_table_cell_appearance<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        row: u32,
        column: u32,
    ) -> Result<Appearance, Error> {
        let (_, _, model) = select(self, slide.into(), table.into())?;
        appearance(
            self,
            engine::Target {
                model_identifier: model,
                row,
                column,
            },
        )
    }

    /// Begin staging cell fill, text-layout and border changes on one slide
    /// table (host `set_slide_table_cell_{fill,layout,border}`,
    /// `reset_slide_table_cell_{fill,layout}`, `clear_slide_table_cell_border`).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] when the selectors resolve no table.
    pub fn edit_slide_table_cell_appearance<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let (slide, table, _) = select(self, slide.into(), table.into())?;
        Ok(Edit {
            source: self,
            slide,
            table,
            operations: Vec::new(),
        })
    }

    /// Replay an exact-source cell-appearance patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact.
    pub fn apply_slide_table_cell_appearance(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_source(&owner(self)?) {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.snapshot()
        } else {
            let candidate =
                Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                    .map_err(|_| Error::Verification)?;
            candidate.validate().map_err(|_| Error::Verification)?;
            candidate
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
