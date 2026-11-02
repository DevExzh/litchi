//! Implementation of the public [`crate::cell::list`] transactions:
//! paragraph lists of the text in one cell (record 0370b).
//!
//! A cell's text carries lists only as rich text: a `TST.RichTextPayload`
//! in its table's rich-text list names a `TSWP.StorageArchive` whose
//! list-style table gives each paragraph its list. A plain cell (empty, or a
//! string key) reads as unlisted paragraphs; the first change that needs a
//! list promotes it the way the host editor's `cell_paragraph_list` does:
//!
//! - the table gains a rich-text list when it has none;
//! - a cell storage holding the cell's text is stored beside the table
//!   model, its paragraphs styled by the cell's text style and the
//!   stylesheet's `None` list preset, with a rich-text payload naming it;
//! - a rich-text entry names the payload, the cell becomes a rich-text cell
//!   pointing at it, and the cell's string entry is released.
//!
//! The list change itself runs W1-D's engine
//! ([`litchi_iwa_text_wire::list`], record 0374) and is published by
//! [`super::storage_list`]. Host `set_x`/`reset_x` pairs collapse into one
//! property verb ([`Edit::set`]/[`Edit::reset`]); `reset_list` and
//! `reset_level` are [`Edit::set_list`] with [`ParagraphList::None`] and
//! [`Edit::set_level`] with [`ParagraphListLevel::ZERO`].

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_table::cell_rich_text::{self as rich, list as engine};
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumbering,
    ParagraphListPlacement,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{Property, PropertyKind};
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{self, Published, Working};
use super::table_topology as topology;
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;

structure::structural_family!("cell-list", Wire(litchi_iwa_common::LimitKind));

/// A typed, content-redacted cell list failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted cell list source is invalid")]
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
    #[error("could not allocate {amount} cell list items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("cell list candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("cell list patch conflicts with this package")]
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
    #[error("no cell list change is staged")]
    NoStagedOperation,
    /// The position is not the start of a paragraph of the cell's text.
    #[error("position is not a paragraph start")]
    NotAParagraphStart,
    /// The paragraph's list kind has no such property (a bullet of a
    /// numbered paragraph, any list property of an unlisted paragraph).
    #[error("the paragraph's list has no such property")]
    NotApplicable,
    /// Placements must start at zero and strictly increase.
    #[error("list placements must start at zero and strictly increase")]
    InvalidPlacements,
    /// The cell's paragraphs do not share one list.
    #[error("the cell's paragraphs do not share one list")]
    NotUniform,
    /// The cell holds a value other than text (a number, a formula, …).
    #[error("the cell does not hold text")]
    NotText,
    /// The cell's rich text or styles live outside what this transaction
    /// rewrites (a shared rich-text payload, a segmented list, no `None`
    /// list preset).
    #[error("a surviving object depends on the cell's rich-text graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell lists")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::NotAParagraphStart | Self::InvalidPlacements => RefusalKind::InvalidValue,
            Self::NotApplicable | Self::NotUniform | Self::NotText => {
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
        use rich::Refusal as R;
        match error {
            rich::Error::Refused(refusal) => Self::Refused(match refusal {
                R::CellNotFound => Refusal::CellNotFound,
                R::UnsupportedSource => Refusal::UnsupportedSource,
                R::NotAParagraphStart => Refusal::NotAParagraphStart,
                R::NotApplicable => Refusal::NotApplicable,
                R::InvalidPlacements => Refusal::InvalidPlacements,
                R::NotUniform => Refusal::NotUniform,
                R::NotText => Refusal::NotText,
                _ => Refusal::UnsupportedDependency,
            }),
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

/// One staged change of one cell: the shared engine's operation.
type Staged = engine::Operation;

/// One immutable edit that changes paragraph lists of cells of one table.
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

    /// Stage giving every paragraph of the cell at `position` the list
    /// preset `list` ([`ParagraphList::None`] removes the lists).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_list(self, position: At, list: ParagraphList) -> Result<Self, Error> {
        self.stage(position, Staged::List(list))
    }

    /// Stage giving the cell's paragraphs the presets of `placements`: each
    /// placement's list runs from its paragraph to the next placement.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_lists(
        self,
        position: At,
        placements: impl IntoIterator<Item = ParagraphListPlacement>,
    ) -> Result<Self, Error> {
        let mut staged = Vec::new();
        for placement in placements {
            staged
                .try_reserve(1)
                .map_err(|_error| Error::Allocation { amount: 1 })?;
            staged.push(placement);
        }
        self.stage(position, Staged::Lists(staged))
    }

    /// Stage the list level of the paragraph starting at `paragraph`
    /// ([`ParagraphListLevel::ZERO`] resets it).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_level(
        self,
        position: At,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Level(paragraph, level))
    }

    /// Stage whether the paragraph starting at `paragraph` continues or
    /// restarts its list's numbering.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_numbering(
        self,
        position: At,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Numbering(paragraph, numbering))
    }

    /// Stage `property` as the effective list property of the paragraph
    /// starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set(
        self,
        position: At,
        paragraph: TextPosition,
        property: Property,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Set(paragraph, property))
    }

    /// Stage restoring the list's own value of `kind` for the paragraph
    /// starting at `paragraph` (its variation's override goes).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn reset(
        self,
        position: At,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<Self, Error> {
        self.stage(position, Staged::Reset(paragraph, kind))
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves the cell as it is is skipped; an edit whose
    /// changes are all skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; the list refusals; source, resource
    /// and verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        // The table model is located once, on the first working set; the
        // engine never renumbers it.
        let mut located: Option<u64> = None;
        for (position, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let working = Working::new(base)?;
            let model = match located {
                Some(model) => model,
                None => *located.insert(locate_model(&working, self.sheet, self.table)?),
            };
            if let Some(published) = engine::apply(working, model, *position, change)? {
                engine::verify(&published.package, model, *position, change)?;
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
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

/// The reader of one cell's lists.
fn reader(
    package: &Package,
    sheet: usize,
    table: usize,
    position: At,
) -> Result<engine::Reader<'_, Package>, Error> {
    let working = Working::new(package)?;
    let model = locate_model(&working, sheet, table)?;
    Ok(engine::Reader::new(working, model, position)?)
}

impl Package {
    /// The list preset every paragraph of one cell's text shares
    /// ([`ParagraphList::None`] for plain text).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], [`Refusal::NotUniform`]
    /// when the paragraphs differ, [`Refusal::NotText`], source errors.
    pub fn table_cell_list<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<ParagraphList, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Ok(reader(self, sheet, table, position)?.list()?)
    }

    /// The list preset of each run of paragraphs of one cell's text.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_list`], without [`Refusal::NotUniform`].
    pub fn table_cell_lists<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Vec<ParagraphListPlacement>, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Ok(reader(self, sheet, table, position)?.lists()?)
    }

    /// The list level of each run of paragraphs of one cell's text.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_lists`].
    pub fn table_cell_list_levels<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Ok(reader(self, sheet, table, position)?.levels()?)
    }

    /// The list level of the paragraph starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_lists`], [`Refusal::NotAParagraphStart`].
    pub fn table_cell_list_level<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLevel, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Ok(reader(self, sheet, table, position)?.level(paragraph)?)
    }

    /// Whether the paragraph starting at `paragraph` continues or restarts
    /// its list's numbering.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_list_level`].
    pub fn table_cell_list_numbering<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumbering, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Ok(reader(self, sheet, table, position)?.numbering(paragraph)?)
    }

    /// The effective list property `kind` of the paragraph starting at
    /// `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_list_level`], [`Refusal::NotApplicable`] when
    /// the paragraph's list has no such property.
    pub fn table_cell_list_property<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
        Ok(reader(self, sheet, table, position)?.property(paragraph, kind)?)
    }

    /// Start an immutable cell-list edit of one table.
    ///
    /// ```no_run
    /// use litchi_numbers::cell::list::ParagraphList;
    /// use litchi_numbers::{At, Package};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package
    ///     .edit_table_cell_lists(0, 0)?
    ///     .set_list(At::new(1, 1), ParagraphList::Bullet)?
    ///     .commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_lists<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet, table) = self.resolve_cell_list_target(sheet.into(), table.into())?;
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

    /// Apply a cell-list patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_lists(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }

    fn resolve_cell_list_target(
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
