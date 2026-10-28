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

use litchi_iwa_core::Limits;
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_style::shape::{
    self as engine,
    property::{Fills, ShapeEffects, Strokes, TextColumns, TextLayout},
};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, SHAPE_INFO_MESSAGE_TYPE};
use super::{style_graph as numbers_style_graph, structure};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{DrawableSelector, Kind};
use crate::sheet::drawable::style::{Property, PropertyKind};

structure::structural_family!("sheet drawable style");

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
            | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
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
            _ => numbers_style_graph::fault(&error).into(),
        }
    }
}

/// One staged style change.
#[derive(Debug, Clone, PartialEq)]
enum Change {
    Set(Property),
    Reset(PropertyKind),
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

fn read(
    workspace: &Workspace<'_>,
    drawable: u64,
    kind: PropertyKind,
) -> Result<Property, Error> {
    Ok(match kind {
        PropertyKind::Fill => {
            let fill = engine::value(workspace, drawable, &Fills)?;
            if matches!(fill, litchi_iwa_common::shape::fill::Fill::Image(_)) {
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
        Change::Set(Property::Stroke(stroke)) => engine::set(workspace, drawable, &Strokes, stroke)?,
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
        Change::Reset(PropertyKind::TextLayout) => {
            engine::reset(workspace, drawable, &TextLayout)?
        },
        Change::Reset(PropertyKind::Columns) => engine::reset(workspace, drawable, &TextColumns)?,
    })
}

/// Style changes staged against one selected Numbers sheet drawable.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    selector: DrawableSelector,
    drawable: u64,
    kind: Kind,
    changes: Vec<Change>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("selector", &self.selector)
            .field("kind", &self.kind)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage a semantic style value.
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        ensure_property(self.kind, property.kind())?;
        if matches!(property, Property::Fill(litchi_iwa_common::shape::fill::Fill::Image(_))) {
            return Err(Refusal::ImageFill.into());
        }
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Change::Set(property));
        Ok(self)
    }

    /// Stage removal of the selected drawable's own override.
    pub fn reset(&mut self, kind: PropertyKind) -> Result<&mut Self, Error> {
        ensure_property(self.kind, kind)?;
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Change::Reset(kind));
        Ok(self)
    }

    /// Validate, publish and strictly reopen all staged changes.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let (components, limits) = workspace(self.source)?;
        let mut working = Workspace::new(&components, limits)?;
        for change in &self.changes {
            apply_change(&mut working, self.drawable, change)?;
        }
        let changes = working.finish();
        if changes.is_empty() {
            return Commit::unchanged(self.source);
        }
        let published = numbers_style_graph::publish(self.source, changes).map_err(Error::from)?;
        let (components, limits) = workspace(&published.package)?;
        let reopened = Workspace::new(&components, limits)?;
        for (position, change) in self.changes.iter().enumerate() {
            let property = match change {
                Change::Set(property) => Some(property),
                Change::Reset(_) => None,
            };
            let kind = match change {
                Change::Set(property) => property.kind(),
                Change::Reset(kind) => *kind,
            };
            if self.changes[position + 1..].iter().any(|later| match later {
                Change::Set(property) => property.kind() == kind,
                Change::Reset(later) => *later == kind,
            }) {
                continue;
            }
            let actual = read(&reopened, self.drawable, kind)?;
            if property.is_some_and(|expected| actual != *expected) {
                return Err(Error::Verification);
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
        })
    }

    /// Apply a reversible exact-source style patch.
    pub fn apply_sheet_drawable_style(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
