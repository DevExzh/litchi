//! Exact-source shape and text-frame styles of Numbers sheet drawables.
//!
//! This adapter resolves a semantic sheet/drawable selector and delegates the
//! copy-on-write style graph to the shared `litchi_iwa_style::shape` engine.
//! `style_graph` performs the bounded exact-source reassembly and strict
//! candidate reopen.  The public surface intentionally keeps native object
//! identifiers, archives and protobuf payloads private.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;
use std::mem::{size_of, size_of_val};

use litchi_iwa_core::Limits;
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_style::shape::{
    self as engine, ShapeProperty, Slot, Styles,
    property::{Fills, ShapeEffects, Strokes, TextColumns, TextLayout},
};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, SHAPE_INFO_MESSAGE_TYPE};
use super::{structure, style_graph as numbers_style_graph};
use crate::selector::SheetSelector;
use crate::sheet::drawable::style::{Columns, Fill, Property, PropertyKind};
use crate::sheet::drawable::{DrawableSelector, Kind};

structure::structural_family!("sheet drawable style", Wire(litchi_iwa_common::LimitKind));

/// A well-formed request this family refuses, classified by
/// [`Refusal::kind`].
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
    /// The selected kind does not own the requested style property.
    #[error("{property:?} does not apply to a Numbers sheet {kind}")]
    UnsupportedProperty {
        /// The selected drawable kind.
        kind: &'static str,
        /// The requested semantic property.
        property: PropertyKind,
    },
    /// Image fills require the format family's media/reference transaction.
    #[error("image fills are not written by the sheet-drawable style family")]
    ImageFill,
    /// The selected style graph contains overrides this route cannot rewrite.
    #[error("the selected drawable's style cannot be rewritten safely")]
    UnsupportedStyle,
    /// A surviving object depends on the style graph being rewritten.
    #[error("a surviving object depends on the selected drawable's style")]
    UnsupportedDependency,
    /// Commit was requested without a staged change.
    #[error("no sheet-drawable style change is staged")]
    NoStagedOperation,
    /// The source has no exact physical artifact suitable for publication.
    #[error("source representation does not support sheet-drawable style edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedProperty { .. } => RefusalKind::UnsupportedTarget,
            Self::ImageFill
            | Self::UnsupportedStyle
            | Self::UnsupportedDependency
            | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
}

/// A typed, content-redacted sheet-drawable style failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or style payload is malformed.
    #[error("the selected Numbers sheet-drawable style source is invalid")]
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
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} sheet-drawable style units or bytes")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("sheet-drawable style candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("sheet-drawable style patch conflicts with this package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<litchi_iwa_style::graph::Error> for Error {
    fn from(error: litchi_iwa_style::graph::Error) -> Self {
        match &error {
            litchi_iwa_style::graph::Error::Invalid(_) => Refusal::UnsupportedStyle.into(),
            _ => match numbers_style_graph::fault(&error) {
                numbers_style_graph::StyleGraphFault::Shared(fault) => fault.into(),
                numbers_style_graph::StyleGraphFault::Limit {
                    kind,
                    observed,
                    maximum,
                } => Error::Limit {
                    kind: match kind {
                        litchi_iwa_style::graph::GraphLimitKind::ScratchBytes => {
                            LimitKind::ScratchBytes
                        },
                        litchi_iwa_style::graph::GraphLimitKind::AllocationEvents => {
                            LimitKind::AllocationEvents
                        },
                    },
                    observed,
                    maximum,
                },
            },
        }
    }
}

/// One staged style change.
#[derive(Debug, Clone, PartialEq)]
enum Change {
    Set(Property),
    Reset(PropertyKind),
}

/// A deliberately small bound on caller-controlled staging.  The value is
/// charged before the change enters the vector, so `Vec::reserve` is only a
/// fallible allocation after the semantic budget has already been enforced.
const MAX_STAGED_OPERATIONS: usize = 64;
const MAX_STAGED_BYTES: usize = 32 * 1024;
const STAGED_ENTRY_OVERHEAD: usize = 128;

fn property_heap_bytes(property: &Property) -> usize {
    match property {
        Property::Fill(Fill::Gradient(gradient)) => gradient
            .stops()
            .len()
            .saturating_mul(size_of::<litchi_iwa_common::shape::fill::Stop>()),
        Property::Columns(Columns::Variable(columns)) => columns
            .following()
            .len()
            .saturating_mul(size_of::<litchi_iwa_text::columns::Following>()),
        _ => 0,
    }
}

fn change_cost(change: &Change) -> usize {
    STAGED_ENTRY_OVERHEAD
        .saturating_add(size_of_val(change))
        .saturating_add(match change {
            Change::Set(property) => property_heap_bytes(property),
            Change::Reset(_) => 0,
        })
}

fn property_index(kind: PropertyKind) -> usize {
    match kind {
        PropertyKind::Fill => 0,
        PropertyKind::Stroke => 1,
        PropertyKind::Effects => 2,
        PropertyKind::TextLayout => 3,
        PropertyKind::Columns => 4,
    }
}

fn reset_slots(kind: PropertyKind) -> &'static [Slot] {
    match kind {
        PropertyKind::Fill => &[Slot::Fill],
        PropertyKind::Stroke => &[Slot::Stroke],
        PropertyKind::Effects => &[Slot::Opacity, Slot::Reflection],
        PropertyKind::TextLayout => &[
            Slot::ShrinkToFit,
            Slot::VerticalAlignment,
            Slot::PaddingNull,
            Slot::Padding,
        ],
        PropertyKind::Columns => &[Slot::ColumnsNull, Slot::Columns],
    }
}

#[derive(Clone)]
struct ResetExpectation {
    effective: Property,
    changed: bool,
    was_exclusive: bool,
}

enum FinalState {
    Set(Property),
    Reset(ResetExpectation),
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)
        .map_err(Error::from)?
        .ok_or(Refusal::SheetNotFound.into())
}

fn resolve_drawable(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> Result<(u64, Kind), Error> {
    let drawables = graph::sheet_drawables(package, sheet).map_err(Error::from)?;
    let entry = drawables
        .select(selector)
        .ok_or(Refusal::DrawableNotFound { selector })?;
    if !matches!(entry.kind, Kind::Shape | Kind::TextBox) {
        return Err(Refusal::UnsupportedProperty {
            kind: entry.kind.name(),
            property: PropertyKind::TextLayout,
        }
        .into());
    }
    let Some((_, message_type)) = entry.message else {
        return Err(Error::InvalidSource);
    };
    if message_type != SHAPE_INFO_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    Ok((entry.identifier, entry.kind))
}

fn ensure_property(kind: Kind, property: PropertyKind) -> Result<(), Error> {
    let valid = match kind {
        Kind::Shape => matches!(
            property,
            PropertyKind::Fill
                | PropertyKind::Stroke
                | PropertyKind::Effects
                | PropertyKind::TextLayout
        ),
        Kind::TextBox => matches!(property, PropertyKind::Columns | PropertyKind::TextLayout),
        _ => false,
    };
    if valid {
        Ok(())
    } else {
        Err(Refusal::UnsupportedProperty {
            kind: kind.name(),
            property,
        }
        .into())
    }
}

fn workspace(
    package: &Package,
) -> Result<(numbers_style_graph::LeafComponents<'_>, Limits), Error> {
    let catalog = structure::exact_source(package).map_err(Error::from)?;
    let limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(|error| Error::from(structure::map_archive(&error)))?;
    let components = numbers_style_graph::LeafComponents::new(package).map_err(Error::from)?;
    Ok((components, limits))
}

fn read(workspace: &Workspace<'_>, drawable: u64, kind: PropertyKind) -> Result<Property, Error> {
    Ok(match kind {
        PropertyKind::Fill => {
            let fill = engine::value(workspace, drawable, &Fills)?;
            if matches!(fill, Fill::Image(_)) {
                return Err(Refusal::ImageFill.into());
            }
            Property::Fill(fill)
        },
        PropertyKind::Stroke => Property::Stroke(engine::value(workspace, drawable, &Strokes)?),
        PropertyKind::Effects => {
            Property::Effects(engine::value(workspace, drawable, &ShapeEffects)?)
        },
        PropertyKind::TextLayout => {
            Property::TextLayout(engine::value(workspace, drawable, &TextLayout)?)
        },
        PropertyKind::Columns => {
            Property::Columns(engine::value(workspace, drawable, &TextColumns)?)
        },
    })
}

fn apply_change(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    change: &Change,
) -> Result<bool, Error> {
    Ok(match change {
        Change::Set(Property::Fill(fill)) => engine::set(workspace, drawable, &Fills, fill)?,
        Change::Set(Property::Stroke(stroke)) => {
            engine::set(workspace, drawable, &Strokes, stroke)?
        },
        Change::Set(Property::Effects(effects)) => {
            engine::set(workspace, drawable, &ShapeEffects, effects)?
        },
        Change::Set(Property::TextLayout(layout)) => {
            engine::set(workspace, drawable, &TextLayout, layout)?
        },
        Change::Set(Property::Columns(columns)) => {
            engine::set(workspace, drawable, &TextColumns, columns)?
        },
        Change::Reset(PropertyKind::Fill) => engine::reset(workspace, drawable, &Fills)?,
        Change::Reset(PropertyKind::Stroke) => engine::reset(workspace, drawable, &Strokes)?,
        Change::Reset(PropertyKind::Effects) => engine::reset(workspace, drawable, &ShapeEffects)?,
        Change::Reset(PropertyKind::TextLayout) => engine::reset(workspace, drawable, &TextLayout)?,
        Change::Reset(PropertyKind::Columns) => engine::reset(workspace, drawable, &TextColumns)?,
    })
}

fn inherited_property<P: ShapeProperty>(
    workspace: &Workspace<'_>,
    style: u64,
    property: &P,
) -> Result<P::Value, Error> {
    let mut values = Vec::new();
    for group in property.groups() {
        values.extend(engine::effective_group(workspace, style, group)?);
    }
    let borrowed = values.iter().map(Option::as_deref).collect::<Vec<_>>();
    Ok(property.read(&borrowed)?)
}

fn inherited_reset_value(
    workspace: &Workspace<'_>,
    drawable: u64,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let target = engine::target(workspace, drawable)?;
    let payload = workspace
        .shape_style(target.style())
        .ok_or(Error::Verification)?;
    let Some(parent) = engine::parent_of(payload)? else {
        // A base style has no inherited parent.  The engine will reject a
        // reset of a direct base-style property; for a no-op reset, its
        // current effective value is the only valid expectation.
        return read(workspace, drawable, kind);
    };
    Ok(match kind {
        PropertyKind::Fill => Property::Fill(inherited_property(workspace, parent, &Fills)?),
        PropertyKind::Stroke => Property::Stroke(inherited_property(workspace, parent, &Strokes)?),
        PropertyKind::Effects => {
            Property::Effects(inherited_property(workspace, parent, &ShapeEffects)?)
        },
        PropertyKind::TextLayout => {
            Property::TextLayout(inherited_property(workspace, parent, &TextLayout)?)
        },
        PropertyKind::Columns => {
            Property::Columns(inherited_property(workspace, parent, &TextColumns)?)
        },
    })
}

fn reset_baselines(
    workspace: &Workspace<'_>,
    drawable: u64,
    changes: &[Change],
) -> Result<[Option<Property>; 5], Error> {
    let mut baselines: [Option<Property>; 5] = std::array::from_fn(|_| None);
    for change in changes {
        let Change::Reset(kind) = change else {
            continue;
        };
        let slot = &mut baselines[property_index(*kind)];
        if slot.is_none() {
            *slot = Some(inherited_reset_value(workspace, drawable, *kind)?);
        }
    }
    Ok(baselines)
}

fn verify_reset_graph(
    workspace: &Workspace<'_>,
    drawable: u64,
    kind: PropertyKind,
    expectation: &ResetExpectation,
    before: engine::Target,
) -> Result<(), Error> {
    if read(workspace, drawable, kind)? != expectation.effective {
        return Err(Error::Verification);
    }
    if !expectation.changed {
        return Ok(());
    }
    let after = engine::target(workspace, drawable)?;
    let payload = workspace
        .shape_style(after.style())
        .ok_or(Error::Verification)?;
    let direct = engine::direct_slots(payload)?;
    if before.exclusive() {
        // An exclusive variation can be rewritten or collapsed; either way
        // the selected property must no longer be a direct override.
        if direct.is_some_and(|slots| slots.iter().any(|slot| reset_slots(kind).contains(slot))) {
            return Err(Error::Verification);
        }
    } else {
        // A shared style cannot be mutated in place.  The engine derives a
        // child that preserves the inherited effective value, so prove that
        // the child is linked into the former shared graph and carries the
        // effective fallback rather than silently changing the sibling.
        let Some(parent) = engine::parent_of(payload)? else {
            return Err(Error::Verification);
        };
        if parent != before.style() {
            return Err(Error::Verification);
        }
        let Some(slots) = direct else {
            return Err(Error::Verification);
        };
        if !slots.iter().any(|slot| reset_slots(kind).contains(slot)) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn verify_reset_candidate(
    workspace: &Workspace<'_>,
    drawable: u64,
    kind: PropertyKind,
    expectation: &ResetExpectation,
) -> Result<(), Error> {
    if read(workspace, drawable, kind)? != expectation.effective {
        return Err(Error::Verification);
    }
    if !expectation.changed {
        return Ok(());
    }
    let target = engine::target(workspace, drawable)?;
    let payload = workspace
        .shape_style(target.style())
        .ok_or(Error::Verification)?;
    let direct = engine::direct_slots(payload)?;
    if expectation.was_exclusive {
        if direct.is_some_and(|slots| slots.iter().any(|slot| reset_slots(kind).contains(slot))) {
            return Err(Error::Verification);
        }
    } else {
        let Some(_parent) = engine::parent_of(payload)? else {
            return Err(Error::Verification);
        };
        let Some(slots) = direct else {
            return Err(Error::Verification);
        };
        if !slots.iter().any(|slot| reset_slots(kind).contains(slot)) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

/// Style changes staged against one selected Numbers sheet drawable.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    selector: DrawableSelector,
    drawable: u64,
    kind: Kind,
    changes: Vec<Change>,
    staged_bytes: usize,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("selector", &self.selector)
            .field("kind", &self.kind)
            .field("changes", &self.changes.len())
            .field("staged_bytes", &self.staged_bytes)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(&mut self, change: Change) -> Result<&mut Self, Error> {
        let observed_operations = self.changes.len().saturating_add(1);
        if observed_operations > MAX_STAGED_OPERATIONS {
            return Err(Error::Limit {
                kind: LimitKind::Wire(litchi_iwa_common::LimitKind::RewriteWork),
                observed: observed_operations as u64,
                maximum: MAX_STAGED_OPERATIONS as u64,
            });
        }
        let cost = change_cost(&change);
        let observed_bytes = self.staged_bytes.saturating_add(cost);
        if observed_bytes > MAX_STAGED_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::Wire(litchi_iwa_common::LimitKind::OutputBytes),
                observed: observed_bytes as u64,
                maximum: MAX_STAGED_BYTES as u64,
            });
        }
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(change);
        self.staged_bytes = observed_bytes;
        Ok(self)
    }

    /// Stage a semantic style value.
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        ensure_property(self.kind, property.kind())?;
        if matches!(property, Property::Fill(Fill::Image(_))) {
            return Err(Refusal::ImageFill.into());
        }
        self.stage(Change::Set(property))
    }

    /// Stage removal of the selected drawable's own override.
    pub fn reset(&mut self, kind: PropertyKind) -> Result<&mut Self, Error> {
        ensure_property(self.kind, kind)?;
        self.stage(Change::Reset(kind))
    }

    /// Validate, publish and strictly reopen all staged changes.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let (components, limits) = workspace(self.source)?;
        let source_view = Workspace::new(&components, limits)?;
        let reset_baselines = reset_baselines(&source_view, self.drawable, &self.changes)?;
        let mut working = Workspace::new(&components, limits)?;
        let mut final_states: [Option<FinalState>; 5] = std::array::from_fn(|_| None);
        for change in &self.changes {
            match change {
                Change::Set(property) => {
                    apply_change(&mut working, self.drawable, change)?;
                    final_states[property_index(property.kind())] =
                        Some(FinalState::Set(property.clone()));
                },
                Change::Reset(kind) => {
                    let before = engine::target(&working, self.drawable)?;
                    let changed = apply_change(&mut working, self.drawable, change)?;
                    let expectation = ResetExpectation {
                        effective: reset_baselines[property_index(*kind)]
                            .clone()
                            .ok_or(Error::Verification)?,
                        changed,
                        was_exclusive: before.exclusive(),
                    };
                    verify_reset_graph(&working, self.drawable, *kind, &expectation, before)?;
                    final_states[property_index(*kind)] = Some(FinalState::Reset(expectation));
                },
            }
        }
        let changes = working.finish();
        if changes.is_empty() {
            return Commit::unchanged(self.source);
        }
        let published = numbers_style_graph::publish(self.source, changes).map_err(Error::from)?;
        let (components, limits) = workspace(&published.package)?;
        let reopened = Workspace::new(&components, limits)?;
        for state in final_states.into_iter().flatten() {
            match state {
                FinalState::Set(expected) => {
                    let actual = read(&reopened, self.drawable, expected.kind())?;
                    if actual != expected {
                        return Err(Error::Verification);
                    }
                },
                FinalState::Reset(expectation) => {
                    let kind = expectation.effective.kind();
                    verify_reset_candidate(&reopened, self.drawable, kind, &expectation)?;
                },
            }
        }
        Commit::new(self.source, published)
    }
}

impl Package {
    /// Read one effective shape or text-frame style value.
    pub fn sheet_drawable_style<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let (drawable, drawable_kind) = resolve_drawable(self, sheet, drawable.into())?;
        ensure_property(drawable_kind, kind)?;
        let (components, limits) = workspace(self)?;
        let workspace = Workspace::new(&components, limits)?;
        read(&workspace, drawable, kind)
    }

    /// Begin a selector-first shape or text-frame style edit.
    pub fn edit_sheet_drawable_style<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let selector = drawable.into();
        let (drawable, kind) = resolve_drawable(self, sheet, selector)?;
        Ok(Edit {
            source: self,
            sheet,
            selector,
            drawable,
            kind,
            changes: Vec::new(),
            staged_bytes: 0,
        })
    }

    /// Apply a reversible exact-source style patch.
    pub fn apply_sheet_drawable_style(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
