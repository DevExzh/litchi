//! Implementation of the public [`crate::slide::table::cell_text`] transactions:
//! one character or paragraph property of whole slide-table cells (record
//! 0412).
//!
//! A cell's text style is a key in its table's style list whose entry names
//! a `TSWP.ParagraphStyleArchive`; a cell without a key uses its section's
//! base style. Each staged change runs the shared cell rich-text engine
//! (`litchi_iwa_table::cell_rich_text::text`), which plans through the
//! shared paragraph-style engine (`litchi_iwa_text_wire::paragraph_style`)
//! and rewrites the cell's own variation in place, derives a new variation
//! beside its stylesheet, or collapses it back onto its parent. This is the
//! behaviour of the host editor's `set_table_cell_*`/`reset_table_cell_*`
//! families on Pages slide tables, collapsed into one property verb.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_table::cell_rich_text::{self as rich, text as engine};
use litchi_iwa_table::structure::Working;
use litchi_iwa_text_wire::paragraph_style::{Change, Property, PropertyKind};
use thiserror::Error as ThisError;

use super::Package;
use super::slide_table_cell_rich_text::{self as shared, rich_text_family};
use crate::SlideSelector;
use crate::slide::table::TableSelector;

rich_text_family!("slide-table cell text");

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No slide or slide table matches the selectors.
    #[error("no slide table matches the selectors")]
    TableNotFound,
    /// The cell is outside the table.
    #[error("cell outside the table")]
    CellNotFound,
    /// Commit was requested before staging a change.
    #[error("no cell text change is staged")]
    NoStagedOperation,
    /// The cell's style lives outside what this transaction rewrites (no
    /// stylesheet, a segmented list, a style in another component).
    #[error("a surviving object depends on the cell's style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell text properties")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }

    const fn from_engine(refusal: rich::Refusal) -> Self {
        match refusal {
            rich::Refusal::CellNotFound => Self::CellNotFound,
            rich::Refusal::UnsupportedSource => Self::UnsupportedSource,
            _ => Self::UnsupportedDependency,
        }
    }
}

/// One staged change of one cell.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind),
}

/// One immutable edit that changes text properties of cells of one slide
/// table.
pub struct Edit<'a> {
    source: &'a Package,
    table: shared::Table,
    changes: Vec<(At, Staged)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.table.slide)
            .field("table", &self.table.position)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, at: At, change: Staged) -> Result<Self, Error> {
        if !self.table.contains(at) {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.changes.push((at, change));
        Ok(self)
    }

    /// Stage `property` as the effective value of the cell at `at`. Changes
    /// apply in staging order.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set(self, at: At, property: Property) -> Result<Self, Error> {
        self.stage(at, Staged::Set(property))
    }

    /// Stage removing the cell variation's own override of `kind`, so the
    /// value is inherited again.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn reset(self, at: At, kind: PropertyKind) -> Result<Self, Error> {
        self.stage(at, Staged::Reset(kind))
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves a property as it is is skipped; an edit whose
    /// changes are all skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; [`Refusal::UnsupportedDependency`]
    /// for a style graph this transaction does not rewrite; source,
    /// resource and verification errors. Failure never publishes a partial
    /// package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (at, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let change = match change {
                Staged::Set(property) => Change::Set(property.clone()),
                Staged::Reset(kind) => Change::Reset(*kind),
            };
            if let Some(published) =
                engine::apply(Working::new(base)?, self.table.model, *at, &change)?
            {
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Ok(Commit::unchanged(self.source));
        };
        // Every staged set that is the cell's last change reads back.
        let working = Working::new(&package)?;
        for (index, (at, change)) in self.changes.iter().enumerate() {
            if let Staged::Set(property) = change
                && !self.changes[index + 1..]
                    .iter()
                    .any(|(other, _)| other == at)
                && &engine::read(&working, self.table.model, *at, property.kind())? != property
            {
                return Err(Error::Verification);
            }
        }
        Ok(Commit::new(self.source, package, touched))
    }
}

impl Package {
    /// The effective value of `kind` for one slide-table cell's text (its text
    /// style and that style's parents; a cell without its own style uses its
    /// section's base style).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::CellNotFound`], source errors.
    pub fn slide_table_cell_text<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let table = shared::resolve(self, slide.into(), table.into())?;
        if !table.contains(at) {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        Ok(engine::read(&Working::new(self)?, table.model, at, kind)?)
    }

    /// Start an immutable cell text-property edit of one slide table.
    ///
    /// ```no_run
    /// use litchi_keynote::Package;
    /// use litchi_keynote::slide::table::cell_text::{At, PropertyKind};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.key")?;
    /// let current = package.slide_table_cell_text(0, 0, At::new(1, 1), PropertyKind::Alignment)?;
    /// let commit = package
    ///     .edit_slide_table_cell_text(0, 0)?
    ///     .reset(At::new(1, 1), PropertyKind::Alignment)?
    ///     .commit()?;
    /// # let _ = (current, commit);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] and source errors.
    pub fn edit_slide_table_cell_text<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            table: shared::resolve(self, slide.into(), table.into())?,
            changes: Vec::new(),
        })
    }

    /// Apply a cell text patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; [`Error::Verification`] when its target does not reopen.
    pub fn apply_slide_table_cell_text(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
