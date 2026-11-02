//! Cell fill, text layout and borders of Numbers table cells (record 0412b;
//! the Numbers leaf of the host's `set_table_cell_{fill,layout,border}`,
//! `reset_table_cell_{fill,layout}`, `clear_table_cell_border` and the
//! `table_cell_{fill,layout,borders}` reads) over the shared cell-appearance
//! engine (`litchi_iwa_table::cell_appearance`), which also serves the Pages
//! and Keynote families.
//!
//! The engine plans each change; this owner resolves the sheet and table
//! selectors, publishes the plan on the shared structural working set
//! (`cell_appearance::working::publish`) and reopens. Each staged operation
//! is applied to the package the previous one produced, as the host applies
//! them one after another; the save token advances once per transaction.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_table::cell_appearance::{self as engine, Kind, Plan, Property};
use thiserror::Error as ThisError;

use super::Package;
use super::structure::{self, Published, Working};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::cell_appearance::{Appearance, BorderSide, Fill, Layout, Stroke};

structure::structural_family!("cell-appearance");

/// A typed, content-redacted cell-appearance failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The table's style graph is malformed or ambiguous.
    #[error("the selected table's cell styles are invalid")]
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
    #[error("could not allocate {amount} cell-appearance items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the new appearance.
    #[error("cell-appearance candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("cell-appearance patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by
/// [`Refusal::kind`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
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
    /// A surviving object depends on what the transaction would rewrite.
    #[error("a surviving object depends on the cell's style graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support cell appearance")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::TableNotFound => RefusalKind::NotFound,
            Self::OutsideTable => RefusalKind::InvalidValue,
            Self::MissingTile
            | Self::UnsupportedStyleList
            | Self::ImageFill
            | Self::UnsupportedCell
            | Self::UnsupportedDependency
            | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
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

/// Cell appearance changes staged against one table.
pub struct Edit<'a> {
    source: &'a Package,
    model: u64,
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
    fn stage(&mut self, operation: Operation) -> Result<&mut Self, Error> {
        self.operations
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.operations.push(operation);
        Ok(self)
    }

    /// Stage the cell's own fill (an image fill is refused).
    ///
    /// # Errors
    ///
    /// [`Refusal::ImageFill`] for an image fill; [`Error::Allocation`].
    pub fn set_fill(&mut self, row: u32, column: u32, fill: Fill) -> Result<&mut Self, Error> {
        if matches!(fill, Fill::Image(_)) {
            return Err(Refusal::ImageFill.into());
        }
        self.stage(Operation::SetFill(row, column, fill))
    }

    /// Stage removing the cell's own fill, so it shows its style's fill.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn reset_fill(&mut self, row: u32, column: u32) -> Result<&mut Self, Error> {
        self.stage(Operation::ResetFill(row, column))
    }

    /// Stage the cell's text wrapping, vertical alignment and insets.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set_layout(
        &mut self,
        row: u32,
        column: u32,
        layout: Layout,
    ) -> Result<&mut Self, Error> {
        self.stage(Operation::SetLayout(row, column, layout))
    }

    /// Stage removing the cell's own text-layout overrides.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn reset_layout(&mut self, row: u32, column: u32) -> Result<&mut Self, Error> {
        self.stage(Operation::ResetLayout(row, column))
    }

    /// Stage an explicit stroke on one edge of the cell.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set_border(
        &mut self,
        row: u32,
        column: u32,
        side: BorderSide,
        stroke: Stroke,
    ) -> Result<&mut Self, Error> {
        self.stage(Operation::Border(row, column, side, Some(stroke)))
    }

    /// Stage an explicitly cleared edge (no stroke, overriding the table
    /// style), as the host's `clear_table_cell_border` writes it.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn clear_border(
        &mut self,
        row: u32,
        column: u32,
        side: BorderSide,
    ) -> Result<&mut Self, Error> {
        self.stage(Operation::Border(row, column, side, None))
    }

    /// Apply every staged change in order, publish and reopen. A change that
    /// leaves the cell as it is is skipped; an edit whose changes are all
    /// skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] without a change; the engine's typed
    /// refusals; resource errors; [`Error::Verification`] when the reopened
    /// cell does not show the change.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        let mut token = None;
        for operation in &self.operations {
            let (row, column) = match operation {
                Operation::SetFill(row, column, _)
                | Operation::ResetFill(row, column)
                | Operation::SetLayout(row, column, _)
                | Operation::ResetLayout(row, column)
                | Operation::Border(row, column, ..) => (*row, *column),
            };
            let target = engine::Target {
                model_identifier: self.model,
                row,
                column,
            };
            let mut plans = Vec::new();
            match operation {
                Operation::SetFill(_, _, fill) => {
                    plans.push(Step::Set(vec![Property::Fill(fill.clone())]));
                },
                Operation::SetLayout(_, _, layout) => plans.push(Step::Set(vec![
                    Property::TextWrap(layout.text_wrap()),
                    Property::VerticalAlignment(layout.vertical_alignment()),
                    Property::Padding(layout.insets()),
                ])),
                Operation::Border(_, _, side, stroke) => plans.push(Step::Border(*side, *stroke)),
                Operation::ResetFill(..) => plans.push(Step::Reset(Kind::Fill)),
                Operation::ResetLayout(..) => {
                    for kind in [Kind::TextWrap, Kind::VerticalAlignment, Kind::Padding] {
                        plans.push(Step::Reset(kind));
                    }
                },
            }
            for step in plans {
                let base = current.as_ref().unwrap_or(self.source);
                if let Some(published) = step.apply(base, target, token)? {
                    if token.is_none() {
                        token = Working::new(&published.package)?.save_token()?;
                    }
                    touched = touched.saturating_add(published.touched_components);
                    current = Some(published.package);
                }
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

/// One engine step of a staged operation.
enum Step {
    Set(Vec<Property>),
    Reset(Kind),
    Border(BorderSide, Option<Stroke>),
}

impl Step {
    fn apply(
        &self,
        package: &Package,
        target: engine::Target,
        token: Option<u64>,
    ) -> Result<Option<Published>, Error> {
        let limits = limits(package)?;
        let fresh = Working::new(package)?.next_identifier();
        let plan: Option<Plan> = match self {
            Self::Set(properties) => engine::plan_set(package, target, properties, fresh, &limits)?,
            Self::Reset(kind) => engine::plan_reset(package, target, *kind, fresh, &limits)?,
            Self::Border(side, stroke) => {
                Some(engine::plan_border(package, target, *side, *stroke, fresh)?)
            },
        };
        let Some(plan) = plan else {
            return Ok(None);
        };
        let published = engine::working::publish(package, &plan, token)?;
        let candidate = &published.package;
        match self {
            Self::Set(properties) => {
                let limits = self::limits(candidate)?;
                for property in properties {
                    if &engine::value(candidate, target, property.kind(), &limits)? != property {
                        return Err(Error::Verification);
                    }
                }
            },
            Self::Border(side, stroke) => {
                if engine::borders(candidate, target)?.get(*side) != *stroke {
                    return Err(Error::Verification);
                }
            },
            Self::Reset(_) => {},
        }
        Ok(Some(published))
    }
}

/// The raw-cell profile of the package's IWA streams.
fn limits(package: &Package) -> Result<litchi_numbers_wire::Limits, Error> {
    let maximum = package.state.options.archive().max_iwa_stream_bytes();
    let wire = litchi_iwa_common::WireLimits::default()
        .with_input_bytes(maximum)
        .and_then(|limits| limits.with_output_bytes(maximum))
        .map_err(|_| Error::InvalidSource)?;
    Ok(litchi_numbers_wire::Limits::from_wire_limits(wire))
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
    /// The selected table's model.
    fn cell_appearance_model(
        &self,
        sheet: SheetSelector<'_>,
        table: TableSelector<'_>,
    ) -> Result<u64, Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table =
            super::table_topology::table_position(self, sheet, table).map_err(
                |error| match error {
                    super::table_topology::Error::Refused(
                        super::table_topology::Refusal::TableNotFound,
                    ) => Error::Refused(Refusal::TableNotFound),
                    _ => Error::InvalidSource,
                },
            )?;
        let working = Working::new(self)?;
        super::table_topology::locate(&working, sheet, table)
            .map(|target| target.model)
            .map_err(|_| Error::InvalidSource)
    }

    /// Read one cell's effective fill, text layout and explicit edge strokes
    /// (host `table_cell_fill`, `table_cell_layout`, `table_cell_borders`).
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::TableNotFound`],
    /// [`Refusal::OutsideTable`] and the typed refusals of an unsupported
    /// style graph.
    pub fn table_cell_appearance<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        row: u32,
        column: u32,
    ) -> Result<Appearance, Error> {
        let model = self.cell_appearance_model(sheet.into(), table.into())?;
        appearance(
            self,
            engine::Target {
                model_identifier: model,
                row,
                column,
            },
        )
    }

    /// Begin staging cell fill, text-layout and border changes on one table
    /// (host `set_table_cell_{fill,layout,border}`,
    /// `reset_table_cell_{fill,layout}`, `clear_table_cell_border`).
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::table::cell_appearance::BorderSide;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_table_cell_appearance(0, 0)?;
    /// edit.clear_border(1, 1, BorderSide::Top)?;
    /// let commit = edit.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::TableNotFound`].
    pub fn edit_table_cell_appearance<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            model: self.cell_appearance_model(sheet.into(), table.into())?,
            operations: Vec::new(),
        })
    }

    /// Replay an exact-source cell-appearance patch.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when the patch was made against another
    /// artifact.
    pub fn apply_table_cell_appearance(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
