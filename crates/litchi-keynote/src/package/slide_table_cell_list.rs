//! Implementation of the public [`crate::slide::table::cell_list`] transactions:
//! paragraph lists of the text in one slide-table cell (record 0412).
//!
//! A cell's text carries lists only as rich text: a `TST.RichTextPayload`
//! in its table's rich-text list names a `TSWP.StorageArchive` whose
//! list-style table gives each paragraph its list. A plain cell (empty, or a
//! string key) reads as unlisted paragraphs; the first change that needs a
//! list promotes it to rich text, as the host editor's `cell_paragraph_list`
//! does. Every change runs the shared cell rich-text engine
//! (`litchi_iwa_table::cell_rich_text::list`) over W1-D's list engine
//! (`litchi_iwa_text_wire::list`). Host `set_x`/`reset_x` pairs collapse into
//! one property verb ([`Edit::set`]/[`Edit::reset`]); `reset_list` and
//! `reset_level` are [`Edit::set_list`] with [`ParagraphList::None`] and
//! [`Edit::set_level`] with [`ParagraphListLevel::ZERO`].

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_table::cell_rich_text::{self as rich, list as engine};
use litchi_iwa_table::structure::Working;
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumbering,
    ParagraphListPlacement,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{Property, PropertyKind};
use thiserror::Error as ThisError;

use super::Package;
use super::slide_table_cell_rich_text::{self as shared, rich_text_family};
use crate::SlideSelector;
use crate::slide::table::TableSelector;

rich_text_family!("slide-table cell list");

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
            Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
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

    const fn from_engine(refusal: rich::Refusal) -> Self {
        match refusal {
            rich::Refusal::CellNotFound => Self::CellNotFound,
            rich::Refusal::UnsupportedSource => Self::UnsupportedSource,
            rich::Refusal::NotAParagraphStart => Self::NotAParagraphStart,
            rich::Refusal::NotApplicable => Self::NotApplicable,
            rich::Refusal::InvalidPlacements => Self::InvalidPlacements,
            rich::Refusal::NotUniform => Self::NotUniform,
            rich::Refusal::NotText => Self::NotText,
            _ => Self::UnsupportedDependency,
        }
    }
}

/// One immutable edit that changes paragraph lists of cells of one slide
/// table.
pub struct Edit<'a> {
    source: &'a Package,
    table: shared::Table,
    changes: Vec<(At, engine::Operation)>,
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
    fn stage(mut self, at: At, change: engine::Operation) -> Result<Self, Error> {
        if !self.table.contains(at) {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.changes.push((at, change));
        Ok(self)
    }

    /// Stage giving every paragraph of the cell at `at` the list preset
    /// `list` ([`ParagraphList::None`] removes the lists).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_list(self, at: At, list: ParagraphList) -> Result<Self, Error> {
        self.stage(at, engine::Operation::List(list))
    }

    /// Stage giving the cell's paragraphs the presets of `placements`: each
    /// placement's list runs from its paragraph to the next placement.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_lists(
        self,
        at: At,
        placements: impl IntoIterator<Item = ParagraphListPlacement>,
    ) -> Result<Self, Error> {
        let mut staged = Vec::new();
        for placement in placements {
            staged
                .try_reserve(1)
                .map_err(|_error| Error::Allocation { amount: 1 })?;
            staged.push(placement);
        }
        self.stage(at, engine::Operation::Lists(staged))
    }

    /// Stage the list level of the paragraph starting at `paragraph`
    /// ([`ParagraphListLevel::ZERO`] resets it).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_level(
        self,
        at: At,
        paragraph: TextPosition,
        level: ParagraphListLevel,
    ) -> Result<Self, Error> {
        self.stage(at, engine::Operation::Level(paragraph, level))
    }

    /// Stage whether the paragraph starting at `paragraph` continues or
    /// restarts its list's numbering.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set_numbering(
        self,
        at: At,
        paragraph: TextPosition,
        numbering: ParagraphListNumbering,
    ) -> Result<Self, Error> {
        self.stage(at, engine::Operation::Numbering(paragraph, numbering))
    }

    /// Stage `property` as the effective list property of the paragraph
    /// starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn set(self, at: At, paragraph: TextPosition, property: Property) -> Result<Self, Error> {
        self.stage(at, engine::Operation::Set(paragraph, property))
    }

    /// Stage restoring the list's own value of `kind` for the paragraph
    /// starting at `paragraph` (its variation's override goes).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`] outside the table, [`Error::Allocation`].
    pub fn reset(self, at: At, paragraph: TextPosition, kind: PropertyKind) -> Result<Self, Error> {
        self.stage(at, engine::Operation::Reset(paragraph, kind))
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
        for (at, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) =
                engine::apply(Working::new(base)?, self.table.model, *at, change)?
            {
                engine::verify(&published.package, self.table.model, *at, change)?;
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Ok(Commit::unchanged(self.source));
        };
        Ok(Commit::new(self.source, package, touched))
    }
}

impl Package {
    fn slide_table_cell_list_reader(
        &self,
        slide: SlideSelector<'_>,
        table: TableSelector,
        at: At,
    ) -> Result<engine::Reader<'_, Self>, Error> {
        let table = shared::resolve(self, slide, table)?;
        if !table.contains(at) {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        Ok(engine::Reader::new(Working::new(self)?, table.model, at)?)
    }

    /// The list preset every paragraph of one slide-table cell's text shares
    /// ([`ParagraphList::None`] for plain text).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::CellNotFound`],
    /// [`Refusal::NotUniform`] when the paragraphs differ,
    /// [`Refusal::NotText`], source errors.
    pub fn slide_table_cell_list<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
    ) -> Result<ParagraphList, Error> {
        Ok(self
            .slide_table_cell_list_reader(slide.into(), table.into(), at)?
            .list()?)
    }

    /// The list preset of each run of paragraphs of one slide-table cell's
    /// text.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_table_cell_list`], without [`Refusal::NotUniform`].
    pub fn slide_table_cell_lists<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
    ) -> Result<Vec<ParagraphListPlacement>, Error> {
        Ok(self
            .slide_table_cell_list_reader(slide.into(), table.into(), at)?
            .lists()?)
    }

    /// The list level of each run of paragraphs of one slide-table cell's
    /// text.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_table_cell_lists`].
    pub fn slide_table_cell_list_levels<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
    ) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
        Ok(self
            .slide_table_cell_list_reader(slide.into(), table.into(), at)?
            .levels()?)
    }

    /// The list level of the paragraph starting at `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_table_cell_lists`], [`Refusal::NotAParagraphStart`].
    pub fn slide_table_cell_list_level<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
        paragraph: TextPosition,
    ) -> Result<ParagraphListLevel, Error> {
        Ok(self
            .slide_table_cell_list_reader(slide.into(), table.into(), at)?
            .level(paragraph)?)
    }

    /// Whether the paragraph starting at `paragraph` continues or restarts
    /// its list's numbering.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_table_cell_list_level`].
    pub fn slide_table_cell_list_numbering<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
        paragraph: TextPosition,
    ) -> Result<ParagraphListNumbering, Error> {
        Ok(self
            .slide_table_cell_list_reader(slide.into(), table.into(), at)?
            .numbering(paragraph)?)
    }

    /// The effective list property `kind` of the paragraph starting at
    /// `paragraph`.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_table_cell_list_level`], [`Refusal::NotApplicable`]
    /// when the paragraph's list has no such property.
    pub fn slide_table_cell_list_property<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
        paragraph: TextPosition,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        Ok(self
            .slide_table_cell_list_reader(slide.into(), table.into(), at)?
            .property(paragraph, kind)?)
    }

    /// Start an immutable cell-list edit of one slide table.
    ///
    /// ```no_run
    /// use litchi_keynote::Package;
    /// use litchi_keynote::slide::table::cell_list::{At, ParagraphList};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.key")?;
    /// let commit = package
    ///     .edit_slide_table_cell_lists(0, 0)?
    ///     .set_list(At::new(0, 0), ParagraphList::Bullet)?
    ///     .commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`] and source errors.
    pub fn edit_slide_table_cell_lists<'slide>(
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

    /// Apply a cell-list patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; [`Error::Verification`] when its target does not reopen.
    pub fn apply_slide_table_cell_lists(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
