//! Implementation of the public [`crate::sheet::drawable`] structural
//! transaction: the back-to-front list of the drawables a sheet owns, the
//! removal of a drawable and its private graph, and z-order moves
//! (record 0378).
//!
//! An [`Edit`] stages any number of operations against one sheet of an
//! immutable package. Selectors resolve against that base snapshot and the
//! edit remembers native identity privately, so a later operation is never
//! shifted by an earlier one. [`Edit::commit`] applies the operations in
//! order in one structural working set, publishes one candidate, reopens it
//! and reads the sheet back.
//!
//! This is the behaviour of the host editor's `sheet_drawable_order`,
//! `set_sheet_drawable_order`, `move_sheet_drawable` and
//! `remove_sheet_{text_box,shape}` routes.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::BTreeSet;
use std::fmt;

use litchi_core::Position;
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, SheetDrawables};
use super::structure::{
    self, Working,
    collect::{Census, verify_removed},
    metadata::{self, Transition},
    references,
};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{Drawable, DrawableSelector, Geometry, Kind};
use litchi_iwa_common::shape::path::Preset;

use super::sheet_drawable_author::{self as author, Context};

structure::structural_family!("sheet-drawable");

/// A typed, content-redacted sheet-drawable failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted sheet-drawable source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} sheet-drawable items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("sheet-drawable candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("sheet-drawable patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The z-order destination is outside the sheet's drawable list.
    #[error("z-order position {position:?} is outside {count} drawables")]
    ZPositionOutOfRange {
        /// The requested back-to-front position.
        position: Position,
        /// The number of drawables on the sheet.
        count: usize,
    },
    /// A new drawable's frame needs a finite position and a positive size
    /// (and a rounded-rectangle radius at most half the shorter side).
    #[error("invalid frame for a new drawable")]
    InvalidGeometry,
    /// The drawable was already removed by an earlier operation of this edit.
    #[error("the drawable was removed earlier in this edit")]
    AlreadyRemoved,
    /// The operation does not apply to the selected kind (removing a table,
    /// image, movie, audio clip or chart through this family).
    #[error("{operation} does not apply to a Numbers sheet {kind}")]
    UnsupportedOperation {
        /// The refused verb.
        operation: &'static str,
        /// The selected kind.
        kind: &'static str,
    },
    /// Commit was requested before staging an operation.
    #[error("no sheet-drawable operation is staged")]
    NoStagedOperation,
    /// A surviving object depends on the removed graph.
    #[error("a surviving object depends on the removed drawable")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support sheet-drawable edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } | Self::AlreadyRemoved => {
                RefusalKind::NotFound
            },
            Self::ZPositionOutOfRange { .. } | Self::InvalidGeometry => RefusalKind::InvalidValue,
            Self::UnsupportedOperation { .. } | Self::NoStagedOperation => {
                RefusalKind::UnsupportedOperation
            },
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One staged operation, with the native identity resolved from the base.
#[derive(Debug, Clone, PartialEq)]
enum Operation {
    Remove { drawable: u64 },
    MoveTo { drawable: u64, position: usize },
    Add(Box<NewShape>),
}

/// One staged text box or preset shape.
#[derive(Debug, Clone, PartialEq)]
struct NewShape {
    preset: Preset,
    geometry: Geometry,
    text: String,
    text_box: bool,
}

/// Drawable operations staged against one sheet of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    base: SheetDrawables,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.base.sheet.position)
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn resolve(&self, selector: DrawableSelector) -> Result<(u64, Kind), Error> {
        let entry = self
            .base
            .select(selector)
            .ok_or(Refusal::DrawableNotFound { selector })?;
        if self.operations.iter().any(|operation| {
            matches!(operation, Operation::Remove { drawable } if *drawable == entry.identifier)
        }) {
            return Err(Refusal::AlreadyRemoved.into());
        }
        Ok((entry.identifier, entry.kind))
    }

    /// The number of drawables the sheet holds after the staged operations.
    fn staged_count(&self) -> usize {
        let removed = self
            .operations
            .iter()
            .filter(|operation| matches!(operation, Operation::Remove { .. }))
            .count();
        let added = self
            .operations
            .iter()
            .filter(|operation| matches!(operation, Operation::Add(_)))
            .count();
        self.base.entries.len() - removed + added
    }

    fn stage_shape(
        &mut self,
        preset: Preset,
        geometry: Geometry,
        text: &str,
        text_box: bool,
    ) -> Result<&mut Self, Error> {
        let valid = geometry.is_valid()
            && geometry.position.is_some()
            && geometry
                .size
                .is_some_and(|size| size.width > 0.0 && size.height > 0.0);
        if !valid {
            return Err(Refusal::InvalidGeometry.into());
        }
        let mut owned = String::new();
        owned
            .try_reserve_exact(text.len())
            .map_err(|_error| Error::Allocation { amount: text.len() })?;
        owned.push_str(text);
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::Add(Box::new(NewShape {
            preset,
            geometry,
            text: owned,
            text_box,
        })));
        Ok(self)
    }

    /// Stage a new text box holding `text` at `geometry` (position and a
    /// positive size required), styled by the theme's text-box preset. It
    /// joins the front of the sheet's list.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidGeometry`], [`Error::Allocation`].
    pub fn add_text_box(&mut self, geometry: Geometry, text: &str) -> Result<&mut Self, Error> {
        self.stage_shape(Preset::Rectangle, geometry, text, true)
    }

    /// Stage a new preset shape holding `text` at `geometry`, styled by the
    /// theme's text-box preset (as Numbers' own shape insertion does). It
    /// joins the front of the sheet's list.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidGeometry`] (also a rounded-rectangle radius above
    /// half the shorter side, at commit), [`Error::Allocation`].
    pub fn add_shape(
        &mut self,
        preset: Preset,
        geometry: Geometry,
        text: &str,
    ) -> Result<&mut Self, Error> {
        self.stage_shape(preset, geometry, text, false)
    }

    /// Stage the removal of a text box or ordinary shape and the objects
    /// only it references (its text storage, title and caption stand-ins).
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::AlreadyRemoved`], and
    /// [`Refusal::UnsupportedOperation`] for the other kinds.
    pub fn remove(&mut self, drawable: impl Into<DrawableSelector>) -> Result<&mut Self, Error> {
        let (identifier, kind) = self.resolve(drawable.into())?;
        if !matches!(kind, Kind::TextBox | Kind::Shape) {
            return Err(Refusal::UnsupportedOperation {
                operation: "remove",
                kind: kind.name(),
            }
            .into());
        }
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::Remove {
            drawable: identifier,
        });
        Ok(self)
    }

    /// Stage a move of the drawable to the zero-based back-to-front
    /// `position` of the sheet's list (interpreted after the drawable is
    /// taken out, as the native Arrange commands do). Selectors of later
    /// operations still resolve against the base snapshot.
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::AlreadyRemoved`],
    /// [`Refusal::ZPositionOutOfRange`].
    pub fn move_to(
        &mut self,
        drawable: impl Into<DrawableSelector>,
        position: Position,
    ) -> Result<&mut Self, Error> {
        let (identifier, _) = self.resolve(drawable.into())?;
        let count = self.staged_count();
        if position.get() >= count {
            return Err(Refusal::ZPositionOutOfRange { position, count }.into());
        }
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::MoveTo {
            drawable: identifier,
            position: position.get(),
        });
        Ok(self)
    }

    /// Validate and atomically publish the staged operations.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit; source, resource and
    /// verification errors. Failure never publishes a partial package. An
    /// edit whose operations leave the sheet unchanged is an exact no-op.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let structural = self
            .operations
            .iter()
            .any(|operation| !matches!(operation, Operation::MoveTo { .. }));
        if structural {
            let (published, order) = publish_structural(self.source, &self.base, &self.operations)?;
            let reopened = graph::sheet_drawables(&published.package, self.base.sheet.position)?;
            if reopened.identifiers() != order {
                return Err(Error::Verification);
            }
            return Commit::new(self.source, published);
        }
        let before = self.base.identifiers();
        let mut order = before.clone();
        for operation in &self.operations {
            if let Operation::MoveTo { drawable, position } = *operation {
                move_in(&mut order, drawable, position)?;
            }
        }
        if order == before {
            return Commit::unchanged(self.source);
        }
        let published = publish_order(self.source, &self.base, &order)?;
        let reopened = graph::sheet_drawables(&published.package, self.base.sheet.position)?;
        if reopened.identifiers() != order {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

fn move_in(order: &mut Vec<u64>, drawable: u64, position: usize) -> Result<(), Error> {
    let current = order
        .iter()
        .position(|identifier| *identifier == drawable)
        .ok_or(Error::InvalidSource)?;
    let moved = order.remove(current);
    if position > order.len() {
        return Err(Refusal::ZPositionOutOfRange {
            position: Position::new(position),
            count: order.len() + 1,
        }
        .into());
    }
    order.insert(position, moved);
    Ok(())
}

/// A pure z-order change: the sheet's list is rewritten in place.
fn publish_order(
    source: &Package,
    base: &SheetDrawables,
    order: &[u64],
) -> Result<structure::Published, Error> {
    let sheet = &base.sheet;
    let mut rewrite = |object: &mut litchi_iwa_core::ArchiveObject,
                       limits: litchi_iwa_core::Limits| {
        let data = references::rewrite(
            &object.messages[sheet.message_index].data,
            sheet.path,
            order,
        )?;
        references::replace(
            object,
            sheet.message_index,
            data,
            &[],
            &[],
            sheet.path,
            limits,
        )
    };
    Ok(graph::publish_in_place(
        source,
        &mut [(sheet.identifier, &mut rewrite)],
    )?)
}

/// Removals, additions and moves in one structural working set; returns
/// the candidate and the sheet's final list.
#[allow(
    clippy::too_many_lines,
    reason = "one staged publication: removals, creations, the sheet list, the registry"
)]
fn publish_structural(
    source: &Package,
    base: &SheetDrawables,
    operations: &[Operation],
) -> Result<(structure::Published, Vec<u64>), Error> {
    let sheet = &base.sheet;
    let mut working = Working::new(source)?;
    let registry = metadata::registry(&working)?;
    let sheet_member = working
        .locate(sheet.identifier)?
        .ok_or(Error::InvalidSource)?;
    let mut order = base.identifiers();
    let mut removed = Vec::new();
    let mut created: Vec<u64> = Vec::new();
    let mut context = None;
    let mut transition = Transition::default();
    for operation in operations {
        match operation {
            Operation::Remove { drawable } => {
                order.retain(|identifier| identifier != drawable);
                removed.push(*drawable);
            },
            Operation::MoveTo { drawable, position } => move_in(&mut order, *drawable, *position)?,
            Operation::Add(shape) => {
                if context.is_none() {
                    context = Some(Context::read(source)?);
                }
                let context = context.as_ref().ok_or(Error::InvalidSource)?;
                let identifiers = working.allocate(4)?;
                let identifiers: [u64; 4] = identifiers
                    .try_into()
                    .map_err(|_error| Error::InvalidSource)?;
                let authored = author::author(
                    context,
                    identifiers,
                    sheet.identifier,
                    shape.geometry,
                    shape.preset,
                    &shape.text,
                    shape.text_box,
                    working.archive_limits(),
                )?
                .ok_or(Refusal::InvalidGeometry)?;
                for object in authored.objects {
                    working.insert(&sheet_member, object)?;
                }
                if let Some(registry) = &registry {
                    let component = registry.component_for_entry(&sheet_member)?.identifier;
                    for identifier in authored.identifiers {
                        let seed = metadata::identity_seed(&working, identifier)?;
                        transition.register_uuid(registry, component, identifier, seed)?;
                    }
                    let style_member = working
                        .locate(context.shape_style)?
                        .ok_or(Error::InvalidSource)?;
                    let style_component = registry.component_for_entry(&style_member)?.identifier;
                    transition.reference(
                        registry,
                        component,
                        style_component,
                        context.shape_style,
                    )?;
                }
                order.push(authored.identifiers[0]);
                created.push(authored.identifiers[0]);
            },
        }
    }
    let mut closure = BTreeSet::new();
    if !removed.is_empty() {
        let census = Census::new(&working)?;
        closure = census.private_closure(&working, &removed)?;
        let detached: Vec<(u64, u64)> = removed
            .iter()
            .map(|drawable| (sheet.identifier, *drawable))
            .collect();
        census.require_detached(&closure, &detached)?;
    }
    let limits = working.archive_limits();
    let object = working.object_mut(sheet.identifier)?;
    let data = references::rewrite(
        &object.messages[sheet.message_index].data,
        sheet.path,
        &order,
    )?;
    references::replace(
        object,
        sheet.message_index,
        data,
        &removed,
        &created,
        sheet.path,
        limits,
    )?;
    if !closure.is_empty() {
        let list: Vec<u64> = closure.iter().copied().collect();
        if let Some(registry) = &registry {
            transition.remove_objects(registry, &list)?;
        }
        transition.release_suffix(&list)?;
        for identifier in &list {
            working.remove(*identifier)?;
        }
        if let Some(registry) = &registry {
            transition.release_unreferenced_data(&working, registry)?;
        }
    }
    let published = working.publish(&transition)?;
    verify_removed(&published.package, &closure)?;
    Ok((published, order))
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The drawables one sheet owns, from the back-most to the front-most
    /// layer (`TN.SheetArchive.drawable_infos`), with their kind and frame.
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] and [`Error::InvalidSource`] for a
    /// malformed drawable list.
    pub fn sheet_drawable_list<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
    ) -> Result<Box<[Drawable]>, Error> {
        let position = resolve_sheet(self, sheet.into())?;
        let drawables = graph::sheet_drawables(self, position)?;
        Ok(graph::list(self, &drawables)?.into_boxed_slice())
    }

    /// Start a structural drawable edit of one sheet: removals and z-order
    /// moves, published together by [`Edit::commit`].
    ///
    /// ```no_run
    /// use litchi_core::Position;
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::drawable::DrawableSelector;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_drawables(0)?;
    /// edit.move_to(DrawableSelector::text_box(0), Position::new(0))?;
    /// let commit = edit.commit()?;
    /// let restored = commit.package().apply_sheet_drawables(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] and [`Error::InvalidSource`].
    pub fn edit_sheet_drawables<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
    ) -> Result<Edit<'_>, Error> {
        let position = resolve_sheet(self, sheet.into())?;
        Ok(Edit {
            source: self,
            base: graph::sheet_drawables(self, position)?,
            operations: Vec::new(),
        })
    }

    /// Apply a sheet-drawable patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_drawables(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
