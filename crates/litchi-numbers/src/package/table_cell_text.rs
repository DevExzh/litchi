//! Implementation of the public [`crate::cell::text`] transactions: one
//! character or paragraph property of whole cells (record 0370).
//!
//! A cell's text style is a key in its table's style list whose entry names
//! a `TSWP.ParagraphStyleArchive`; a cell without a key uses its section's
//! base style. Each staged change runs the shared paragraph-style engine
//! ([`litchi_iwa_text_wire::paragraph_style`], record 0374) over the cell's
//! style, with the cell's own variation counted exclusive when its list
//! entry is its only referrer, nothing else references the style and no
//! style names it as parent. The engine's plan is published here:
//!
//! - **Rewrite**: the cell's own variation takes the new payload in place;
//! - **Derive**: a new child variation of the cell's style is stored beside
//!   its stylesheet and registered there, a list entry names it and the cell
//!   points at the entry (an empty cell takes a style-only cell);
//! - **Collapse**: the cell's variation no longer overrides anything, so the
//!   cell points back at the parent (no key when the parent is the section's
//!   base style), the variation's entry is released and the variation is
//!   removed with its stylesheet registration.
//!
//! This is the behaviour of the host editor's `cell_paragraph_style`
//! `set_table_cell_*`/`reset_table_cell_*` families, collapsed into one
//! property verb.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_table::cell_rich_text::{self as rich, text as engine};
use litchi_iwa_text_wire::paragraph_style::{Change, Property, PropertyKind};
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{self, Published, Working};
use super::table_topology as topology;
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;

structure::structural_family!("cell-text", Wire(litchi_iwa_common::LimitKind));

/// A typed, content-redacted cell text failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted cell text source is invalid")]
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
    #[error("could not allocate {amount} cell text items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("cell text candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("cell text patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
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
    /// The value cannot be represented by the native style.
    #[error("the property value cannot be stored")]
    InvalidValue,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::InvalidValue => RefusalKind::InvalidValue,
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

impl From<topology::Error> for Error {
    fn from(error: topology::Error) -> Self {
        use topology::{Error as E, Refusal as R};
        match error {
            E::Refused(R::SheetNotFound) => Self::Refused(Refusal::SheetNotFound),
            E::Refused(R::TableNotFound) => Self::Refused(Refusal::TableNotFound),
            E::Refused(R::UnsupportedSource) => Self::Refused(Refusal::UnsupportedSource),
            E::Allocation { amount } => Self::Allocation { amount },
            _ => Self::InvalidSource,
        }
    }
}

/// The shared engine's error (record 0412).
impl From<rich::Error> for Error {
    fn from(error: rich::Error) -> Self {
        match error {
            rich::Error::Refused(rich::Refusal::CellNotFound) => {
                Self::Refused(Refusal::CellNotFound)
            },
            rich::Error::Refused(rich::Refusal::UnsupportedSource) => {
                Self::Refused(Refusal::UnsupportedSource)
            },
            rich::Error::Refused(_) => Self::Refused(Refusal::UnsupportedDependency),
            rich::Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    rich::LimitKind::PackageBytes => LimitKind::PackageBytes,
                    rich::LimitKind::Identifiers => LimitKind::Identifiers,
                    rich::LimitKind::Wire(kind) => LimitKind::Wire(kind),
                    _ => LimitKind::PayloadBytes,
                },
                observed,
                maximum,
            },
            rich::Error::Allocation { amount } => Self::Allocation { amount },
            rich::Error::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

/// One staged change of one cell.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind),
}

/// One immutable edit that changes text properties of cells of one table.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    rows: u32,
    columns: u32,
    changes: Vec<(At, Staged)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, position: At, change: Staged) -> Result<Self, Error> {
        if position.row() >= self.rows || position.column() >= self.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.changes.push((position, change));
        Ok(self)
    }

    /// Stage `property` as the effective value of the cell at `position`.
    /// Changes apply in staging order.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set(self, position: At, property: Property) -> Result<Self, Error> {
        self.stage(position, Staged::Set(property))
    }

    /// Stage removing the cell variation's own override of `kind`, so the
    /// value is inherited again.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn reset(self, position: At, kind: PropertyKind) -> Result<Self, Error> {
        self.stage(position, Staged::Reset(kind))
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
        for (position, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let change = match change {
                Staged::Set(property) => Change::Set(property.clone()),
                Staged::Reset(kind) => Change::Reset(*kind),
            };
            if let Some(published) = apply_one(base, self.sheet, self.table, *position, &change)? {
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        // Every staged set reads back.
        for (position, change) in &self.changes {
            if let Staged::Set(property) = change
                && let Some(later) = self
                    .changes
                    .iter()
                    .rposition(|(other, _)| other == position)
                && self.changes[later].1 == *change
                && &text(&package, self.sheet, self.table, *position, property.kind())? != property
            {
                return Err(Error::Verification);
            }
        }
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

/// The table model of `table` of sheet `sheet` in the working set.
fn locate_model(working: &Working<'_>, sheet: usize, table: usize) -> Result<u64, Error> {
    Ok(topology::locate(working, sheet, table)?.model)
}

/// The effective `kind` of the cell at `position`.
fn text(
    package: &Package,
    sheet: usize,
    table_index: usize,
    position: At,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let working = Working::new(package)?;
    let model = locate_model(&working, sheet, table_index)?;
    Ok(engine::read(&working, model, position, kind)?)
}

/// One staged change, through the shared engine.
fn apply_one(
    source: &Package,
    sheet: usize,
    table_index: usize,
    position: At,
    change: &Change,
) -> Result<Option<Published>, Error> {
    let working = Working::new(source)?;
    let model = locate_model(&working, sheet, table_index)?;
    Ok(engine::apply(working, model, position, change)?)
}

impl Package {
    /// The effective value of `kind` for one cell's text (its text style and
    /// that style's parents; a cell without its own style uses its section's
    /// base style).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], source errors.
    pub fn table_cell_text<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let (sheet, table) = self.resolve_cell_text_target(sheet.into(), table.into())?;
        text(self, sheet, table, position, kind)
    }

    /// Start an immutable cell text-property edit of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::cell::text::{Property, PropertyKind};
    /// use litchi_numbers::{At, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let current = package.table_cell_text(0, 0, At::new(1, 1), PropertyKind::Alignment)?;
    /// let commit = package
    ///     .edit_table_cell_text(0, 0)?
    ///     .reset(At::new(1, 1), PropertyKind::Alignment)?
    ///     .commit()?;
    /// # let _ = (current, commit);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_text<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet, table) = self.resolve_cell_text_target(sheet.into(), table.into())?;
        let leaf = self.workbook_ref().sheets()[sheet]
            .tables()
            .nth(table)
            .ok_or(Error::Refused(Refusal::TableNotFound))?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            rows: leaf.row_count(),
            columns: leaf.column_count(),
            changes: Vec::new(),
        })
    }

    /// Apply a cell text patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_text(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }

    fn resolve_cell_text_target(
        &self,
        sheet: SheetSelector<'_>,
        table: TableSelector<'_>,
    ) -> Result<(usize, usize), Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = topology::table_position(self, sheet, table).map_err(Error::from)?;
        Ok((sheet, table))
    }
}
