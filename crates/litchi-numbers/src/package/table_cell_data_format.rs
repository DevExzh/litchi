//! Any display format or control of a stored Numbers table cell, as the
//! neutral [`DataFormat`] (record 0370k; the host editor's
//! `set_table_cell_data_format` and `table_cell_data_format`).
//!
//! The read is the shared `litchi_iwa_table::cell_format::data_format`
//! dispatch. A write composes the Numbers per-family transactions: a change
//! inside one family (or between Automatic and a format) is that family's
//! transaction, and a change between two families is two (the first to
//! Automatic) on a private intermediate candidate. Checkbox, Star Rating,
//! Slider, Stepper and Pop-Up Menu go through the cell-control family, which
//! also rewrites the package metadata registry. The patch keeps only the
//! source and final artifacts, so its inverse restores the source byte for
//! byte.
//!
//! Numeral systems are a display family of the shared engine since this
//! record; the Numbers owner writes them through the engine directly, as it
//! has no dedicated numeral-system transaction.

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_table::cell_format::{self as engine, data_format};
use thiserror::Error as ThisError;

use super::{Package, table_cell_pop_up_menu as popup};
use crate::cell::data_format::{CellControl, DataFormat};
use crate::{At, SheetSelector, TableSelector};

/// A finite resource governed by a cell data-format transaction (the
/// family transactions' resource, coarsened).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Work charged to the family transaction the change runs.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("transaction work")
    }
}

/// A typed, content-redacted cell data-format failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted table graph is malformed or ambiguous.
    #[error("the Numbers cell data-format source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error(
        "Numbers cell data-format {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Numbers cell data-format units")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the staged format.
    #[error("Numbers cell data-format candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source, or
    /// the cell changed since the edit was staged.
    #[error("Numbers cell data-format patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No sheet or table matches the selectors.
    #[error("no Numbers sheet or table matches the selectors")]
    TableNotFound,
    /// The cell is outside the table or holds no stored cell record.
    #[error("the Numbers cell is outside the table or not stored")]
    CellNotFound,
    /// The table is locked.
    #[error("the Numbers table is locked")]
    TableLocked,
    /// The change is not one this route writes.
    #[error("this Numbers cell format is not written by this route")]
    UnsupportedFormat,
    /// The value is not valid for its family.
    #[error("the Numbers cell format value is invalid")]
    InvalidValue,
    /// A native dependency of the cell is outside the engines.
    #[error("the Numbers cell has an unsupported native dependency")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support Numbers cell formats")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::TableLocked => RefusalKind::Protected,
            Self::UnsupportedFormat => RefusalKind::UnsupportedOperation,
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

/// Map a family refusal by its shared kind.
fn refusal(kind: litchi_core::edit::RefusalKind, cell: bool) -> Refusal {
    use litchi_core::edit::RefusalKind;
    match kind {
        RefusalKind::NotFound if cell => Refusal::CellNotFound,
        RefusalKind::NotFound => Refusal::TableNotFound,
        RefusalKind::Protected => Refusal::TableLocked,
        RefusalKind::InvalidValue => Refusal::InvalidValue,
        RefusalKind::UnsupportedOperation => Refusal::UnsupportedFormat,
        _ => Refusal::UnsupportedDependency,
    }
}

macro_rules! from_family {
    ($($module:ident),* $(,)?) => {$(
        impl From<super::$module::Error> for Error {
            fn from(error: super::$module::Error) -> Self {
                use super::$module::{Error as E, Refusal as R};
                match error {
                    E::Refused(R::UnsupportedSource) => Self::Refused(Refusal::UnsupportedSource),
                    E::Refused(R::CellNotFound) => Self::Refused(Refusal::CellNotFound),
                    E::Refused(other) => Self::Refused(refusal(other.kind(), false)),
                    E::InvalidSource { .. } => Self::InvalidSource,
                    E::Limit { observed, maximum, .. } => Self::Limit {
                        kind: LimitKind::TransactionWork,
                        observed,
                        maximum,
                    },
                    E::Allocation { amount, .. } => Self::Allocation { amount },
                    E::Verification => Self::Verification,
                    E::PatchConflict => Self::PatchConflict,
                }
            }
        }
    )*};
}

from_family!(
    table_cell_number_format,
    table_cell_currency_format,
    table_cell_percentage_format,
    table_cell_scientific_format,
    table_cell_fraction_format,
    table_cell_date_time_format,
    table_cell_duration_format,
    table_cell_text_format,
    table_cell_custom_format,
    table_cell_pop_up_menu,
);

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        match error {
            engine::Error::Refused(refusal) => Self::Refused(match refusal {
                engine::Refusal::CellNotFound => Refusal::CellNotFound,
                engine::Refusal::TableLocked { .. } => Refusal::TableLocked,
                engine::Refusal::UnsupportedDependency { .. } => Refusal::UnsupportedDependency,
                engine::Refusal::UnsupportedSource => Refusal::UnsupportedSource,
            }),
            engine::Error::InvalidSource { .. } => Self::InvalidSource,
            engine::Error::Limit {
                observed, maximum, ..
            } => Self::Limit {
                kind: LimitKind::TransactionWork,
                observed,
                maximum,
            },
            engine::Error::Allocation { amount, .. } => Self::Allocation { amount },
            engine::Error::Verification => Self::Verification,
            engine::Error::PatchConflict => Self::PatchConflict,
        }
    }
}

/// A reversible exact-source cell data-format patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: usize,
    table: usize,
    at: At,
    before: DataFormat,
    after: DataFormat,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("at", &self.at)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The zero-based sheet position the patch edits.
    #[must_use]
    pub const fn sheet(&self) -> usize {
        self.sheet
    }

    /// The zero-based table position on that sheet.
    #[must_use]
    pub const fn table(&self) -> usize {
        self.table
    }

    /// The edited cell.
    #[must_use]
    pub const fn at(&self) -> At {
        self.at
    }

    /// The source format.
    #[must_use]
    pub const fn before(&self) -> &DataFormat {
        &self.before
    }

    /// The target format.
    #[must_use]
    pub const fn after(&self) -> &DataFormat {
        &self.after
    }

    /// Whether the patch is an exact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            table: self.table,
            at: self.at,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Content-free diagnostics for one cell data-format publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    rewrites: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Family transactions published: one, or two for a change between
    /// families.
    #[must_use]
    pub const fn rewrites(self) -> usize {
        self.rewrites
    }
}

/// One validated cell data-format publication.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume this publication and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// One staged format change of one stored table cell.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    at: At,
    before: DataFormat,
    after: DataFormat,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("at", &self.at)
            .finish_non_exhaustive()
    }
}

/// The family transaction a format belongs to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Owner {
    Display(data_format::Family),
    Control,
    NumeralSystem,
}

fn owner(format: &DataFormat) -> Option<Owner> {
    use data_format::Family;
    Some(match data_format::family(format)? {
        Family::ScalarControl | Family::PopUpMenu => Owner::Control,
        Family::NumeralSystem => Owner::NumeralSystem,
        family => Owner::Display(family),
    })
}

impl Edit<'_> {
    /// The format observed when the edit was created.
    #[must_use]
    pub const fn before(&self) -> &DataFormat {
        &self.before
    }

    /// The staged format.
    #[must_use]
    pub const fn after(&self) -> &DataFormat {
        &self.after
    }

    /// Stage `format` for the cell ([`DataFormat::Automatic`] clears it).
    #[must_use]
    pub fn set(mut self, format: impl Into<DataFormat>) -> Self {
        self.after = format.into();
        self
    }

    /// Stage iWork's automatic format for the cell.
    #[must_use]
    pub fn reset(mut self) -> Self {
        self.after = DataFormat::Automatic;
        self
    }

    /// Validate and atomically publish the staged format.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableLocked`]; the family transactions' source,
    /// dependency and resource failures; [`Error::Verification`]. A failure
    /// never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let source_owner = artifact(self.source)?;
        let patch = |artifacts: OwnedExactArtifacts| Patch {
            artifacts,
            sheet: self.sheet,
            table: self.table,
            at: self.at,
            before: self.before.clone(),
            after: self.after.clone(),
        };
        if self.before == self.after {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: patch(OwnedExactArtifacts::new(source_owner.clone(), source_owner)),
                diagnostics: Diagnostics::default(),
            });
        }
        let (before, after) = (owner(&self.before), owner(&self.after));
        // A change that involves a numeral system runs every step through
        // the shared engine, appending list keys as the host editor's
        // compatibility writer does (record 0370k).
        let numeral = before == Some(Owner::NumeralSystem) || after == Some(Owner::NumeralSystem);
        let rewrite =
            |source: &Package, family: Owner, before: &DataFormat, after: Option<&DataFormat>| {
                if numeral && family != Owner::Control {
                    engine_rewrite(
                        source,
                        SheetSelector::index(self.sheet),
                        TableSelector::index(self.table),
                        self.at,
                        before,
                        after,
                    )
                } else {
                    rewrite(source, self.sheet, self.table, self.at, family, after)
                }
            };
        let (package, rewrites) = match (before, after) {
            (Some(left), Some(right)) if left != right => {
                let cleared = rewrite(self.source, left, &self.before, None)?;
                (
                    rewrite(&cleared, right, &DataFormat::Automatic, Some(&self.after))?,
                    2,
                )
            },
            (_, Some(family)) => (
                rewrite(self.source, family, &self.before, Some(&self.after))?,
                1,
            ),
            (Some(family), None) => (rewrite(self.source, family, &self.before, None)?, 1),
            (None, None) => return Err(Error::Verification),
        };
        if read(&package, self.sheet, self.table, self.at)? != self.after {
            return Err(Error::Verification);
        }
        let target_owner = artifact(&package)?;
        Ok(Commit {
            package,
            patch: patch(OwnedExactArtifacts::new(source_owner, target_owner)),
            diagnostics: Diagnostics {
                changed: true,
                rewrites,
            },
        })
    }
}

fn artifact(package: &Package) -> Result<litchi_iwa_archive::package::SharedBytes, Error> {
    Ok(super::table_headers::rewrite::physical_source(package)
        .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?
        .__source_owner())
}

/// One family transaction: set `after` (a value of `family`) or, for
/// `None`, return the cell to Automatic.
fn rewrite(
    source: &Package,
    sheet: usize,
    table: usize,
    at: At,
    family: Owner,
    after: Option<&DataFormat>,
) -> Result<Package, Error> {
    use data_format::Family;
    let (sheet, table) = (SheetSelector::index(sheet), TableSelector::index(table));
    macro_rules! display {
        ($edit:ident, $variant:ident) => {{
            let edit = source.$edit(sheet, table, at)?;
            let edit = match after {
                Some(DataFormat::$variant(value)) => edit.set(value.clone()),
                Some(_) => return Err(Error::Verification),
                None => edit.clear(),
            };
            Ok(edit.commit()?.into_package())
        }};
    }
    match family {
        Owner::Display(Family::Number) => display!(edit_table_cell_number_format, Number),
        Owner::Display(Family::Currency) => display!(edit_table_cell_currency_format, Currency),
        Owner::Display(Family::Percentage) => {
            display!(edit_table_cell_percentage_format, Percentage)
        },
        Owner::Display(Family::Scientific) => {
            display!(edit_table_cell_scientific_format, Scientific)
        },
        Owner::Display(Family::Fraction) => display!(edit_table_cell_fraction_format, Fraction),
        Owner::Display(Family::DateTime) => display!(edit_table_cell_date_time_format, DateTime),
        Owner::Display(Family::Duration) => display!(edit_table_cell_duration_format, Duration),
        Owner::Display(Family::Text) => display!(edit_table_cell_text_format, Text),
        Owner::Display(Family::Custom) => display!(edit_table_cell_custom_format, Custom),
        Owner::Control => {
            let edit = source.edit_table_cell_control_format(sheet, table, at)?;
            let edit = match after {
                Some(format) => edit.set(
                    CellControl::try_from(format.clone()).map_err(|_error| Error::Verification)?,
                ),
                None => edit.clear(),
            };
            Ok(edit.commit()?.into_package())
        },
        // A numeral system always goes through `engine_rewrite`.
        Owner::Display(_) | Owner::NumeralSystem => Err(Error::Refused(Refusal::UnsupportedFormat)),
    }
}

/// One numeral-system rewrite straight through the shared engine (the
/// Numbers owner has no dedicated numeral-system transaction), published by
/// the engine's exact-source reassembly and reopened.
fn engine_rewrite(
    source: &Package,
    sheet: SheetSelector<'_>,
    table: TableSelector<'_>,
    at: At,
    before: &DataFormat,
    after: Option<&DataFormat>,
) -> Result<Package, Error> {
    let target = popup::resolve_cell(source, sheet, table, at)?;
    let path = target.path();
    let after = after.cloned().unwrap_or(DataFormat::Automatic);
    let catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
    if !catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    // The host editor writes a numeral system through its compatibility
    // writer, which appends list keys at the watermark; the Numbers
    // per-family owners reuse free keys (record 0379).
    let mut budget = engine::TransactionBudget::for_cell_control(source)
        .with_list_keys(engine::ListKeys::Append);
    budget.charge_package_source(catalog, path)?;
    let bound =
        engine::control::preflight_copy_on_write_scalar_control(source, target, path, &mut budget)?;
    let output = data_format::rewrite(source, target, before, &after, path, &mut budget)?;
    let bytes = engine::publish::reassemble(
        source,
        &output.members,
        data_format::maximum_members(data_format::Family::NumeralSystem),
        bound,
        path,
        &mut budget,
    )?;
    let candidate = source
        .reopen_candidate(std::sync::Arc::from(bytes))
        .map_err(|_error| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_error| Error::Verification)?;
    budget.charge_candidate_reopen(candidate_catalog, path)?;
    let candidate_target = popup::resolve_cell(&candidate, sheet, table, at)?;
    if data_format::read_with_budget(&candidate, candidate_target, path, &mut budget)? != after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

fn read(package: &Package, sheet: usize, table: usize, at: At) -> Result<DataFormat, Error> {
    let target = popup::resolve_cell(
        package,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        at,
    )?;
    Ok(data_format::read(package, target, target.path())?)
}

impl Package {
    /// Read the display format or control of one stored table cell.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::CellNotFound`] (outside the
    /// table, or a cell with no stored record),
    /// [`Refusal::UnsupportedDependency`] for a format no engine family
    /// admits, and source failures.
    pub fn table_cell_data_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        at: At,
    ) -> Result<DataFormat, Error> {
        let target = popup::resolve_cell(self, sheet, table, at)?;
        Ok(data_format::read(self, target, target.path())?)
    }

    /// Start a data-format edit of one stored table cell.
    ///
    /// # Errors
    ///
    /// As [`Self::table_cell_data_format`].
    pub fn edit_table_cell_data_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        at: At,
    ) -> Result<Edit<'_>, Error> {
        let target = popup::resolve_cell(self, sheet, table, at)?;
        let before = data_format::read(self, target, target.path())?;
        Ok(Edit {
            source: self,
            sheet: target.scope,
            table: target.table,
            at,
            after: before.clone(),
            before,
        })
    }

    /// Replay a cell data-format patch on its exact source.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] on another artifact or a changed cell;
    /// [`Error::Verification`] when the target does not reopen to the
    /// patch's format.
    pub fn apply_table_cell_data_format(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_owner(&artifact(self)?) {
            return Err(Error::PatchConflict);
        }
        if read(self, patch.sheet, patch.table, patch.at)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::default(),
            });
        }
        let candidate = Package::from_source_owner_with_options(
            patch.artifacts.target_owner(),
            self.state.options,
        )
        .map_err(|_error| Error::Verification)?;
        if read(&candidate, patch.sheet, patch.table, patch.at)? != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                rewrites: 1,
            },
        })
    }
}
