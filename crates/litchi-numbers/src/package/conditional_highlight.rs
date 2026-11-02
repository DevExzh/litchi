//! Implementation of the public [`crate::cell::conditional_highlight`]
//! transactions (record 0370d): the ordered conditional-highlight rules of
//! one cell.
//!
//! The native rule graph surgery is the shared conditional-highlight engine
//! (`litchi_iwa_table::conditional_highlight`, lifted from this module in
//! record 0412f so Pages and Keynote run it too); this module keeps the
//! Numbers selectors, the staged edit and the exact-source patch.

use std::fmt;

use litchi_iwa_table::conditional_highlight::{self as engine, Staged};
use litchi_iwa_table::topology::Target;
use thiserror::Error as ThisError;

pub(in crate::package) use litchi_iwa_common::table::cell::conditional_highlight::Rule;

use super::Package;
use super::structure::{self, Published, Working};
use super::table_topology as topology;
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;

structure::structural_family!("conditional-highlight", Wire(litchi_iwa_common::LimitKind));

type Result<T, E = Error> = std::result::Result<T, E>;

/// A typed, content-redacted conditional-highlight failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed, or its rules use a representation
    /// this route does not read.
    #[error("rooted conditional-highlight source is invalid")]
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
    #[error("could not allocate {amount} conditional-highlight items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("conditional-highlight candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("conditional-highlight patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
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
    /// The cell is outside the table.
    #[error("cell outside the table")]
    CellNotFound,
    /// Commit was requested before staging a change.
    #[error("no conditional-highlight change is staged")]
    NoStagedOperation,
    /// A rule set must hold between one and fifteen rules.
    #[error("conditional highlighting takes one to fifteen rules")]
    InvalidRuleCount,
    /// An object outside the rule graph references the graph's styles.
    #[error("a surviving object depends on the conditional-highlight graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support conditional highlights")]
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
            Self::InvalidRuleCount => RefusalKind::InvalidValue,
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
            E::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    topology::LimitKind::PackageBytes => LimitKind::PackageBytes,
                    topology::LimitKind::Identifiers => LimitKind::Identifiers,
                    _ => LimitKind::PayloadBytes,
                },
                observed,
                maximum,
            },
            E::Allocation { amount } => Self::Allocation { amount },
            E::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        use engine::{Error as E, LimitKind as K, Refusal as R};
        match error {
            E::Refused(R::CellNotFound) => Self::Refused(Refusal::CellNotFound),
            E::Refused(R::InvalidRuleCount) => Self::Refused(Refusal::InvalidRuleCount),
            E::Refused(R::UnsupportedDependency) => Self::Refused(Refusal::UnsupportedDependency),
            E::Refused(R::UnsupportedSource) => Self::Refused(Refusal::UnsupportedSource),
            E::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    K::PackageBytes => LimitKind::PackageBytes,
                    K::Identifiers => LimitKind::Identifiers,
                    K::Wire(kind) => LimitKind::Wire(kind),
                    _ => LimitKind::PayloadBytes,
                },
                observed,
                maximum,
            },
            E::Allocation { amount } => Self::Allocation { amount },
            E::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

/// The located table of `(sheet, table)`.
fn target(source: &Package, sheet: usize, table: usize) -> Result<Target> {
    Ok(topology::locate(&Working::new(source)?, sheet, table)?)
}

/// The cell's rules (`None` without conditional highlighting).
fn rules_of(
    package: &Package,
    sheet: usize,
    table: usize,
    position: At,
) -> Result<Option<Vec<Rule>>> {
    Ok(engine::rules_of(
        package,
        target(package, sheet, table)?,
        position,
    )?)
}

fn apply_one(
    source: &Package,
    sheet: usize,
    table: usize,
    position: At,
    staged: &Staged,
) -> Result<Option<Published>> {
    Ok(engine::apply_one(
        source,
        target(source, sheet, table)?,
        position,
        staged,
    )?)
}

/// One immutable edit that changes the conditional highlighting of cells of
/// one table.
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
    fn stage(mut self, position: At, change: Staged) -> Result<Self> {
        if position.row() >= self.rows || position.column() >= self.columns {
            return Err(Error::Refused(Refusal::CellNotFound));
        }
        self.changes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.changes.push((position, change));
        Ok(self)
    }

    /// Stage `rules`, in order, as the cell's conditional highlighting
    /// (replacing any it has).
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`], [`Refusal::InvalidRuleCount`] for no rule
    /// or more than fifteen, [`Error::Allocation`]. A volatile date rule
    /// registers the cell with the calculation engine's volatile owner
    /// family (record 0376a).
    pub fn set(self, position: At, rules: impl IntoIterator<Item = Rule>) -> Result<Self> {
        let mut staged = Vec::new();
        for rule in rules {
            staged
                .try_reserve(1)
                .map_err(|_error| Error::Allocation { amount: 1 })?;
            staged.push(rule);
        }
        if staged.is_empty() || staged.len() > engine::MAX_RULES {
            return Err(Error::Refused(Refusal::InvalidRuleCount));
        }
        self.stage(position, Staged::Set(staged))
    }

    /// Stage removing the cell's conditional highlighting, keeping its value
    /// and base style.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`], [`Error::Allocation`].
    pub fn clear(self, position: At) -> Result<Self> {
        self.stage(position, Staged::Clear)
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves the cell as it is is skipped.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], the refusals, source, resource and
    /// verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit> {
        if self.changes.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (position, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.sheet, self.table, *position, change)? {
                let expected = match change {
                    Staged::Set(rules) => Some(rules.clone()),
                    Staged::Clear => None,
                };
                if rules_of(&published.package, self.sheet, self.table, *position)? != expected {
                    return Err(Error::Verification);
                }
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

impl Package {
    /// The ordered conditional-highlight rules of one cell (`None` when it
    /// has none).
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::CellNotFound`], [`Error::InvalidSource`]
    /// for a rule representation this route does not read.
    pub fn table_cell_conditional_highlight<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Option<Vec<Rule>>> {
        let (sheet, table) = self.resolve_conditional_target(sheet.into(), table.into())?;
        rules_of(self, sheet, table, position)
    }

    /// Start an immutable conditional-highlight edit of one table.
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_table_cell_conditional_highlights<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>> {
        let (sheet, table) = self.resolve_conditional_target(sheet.into(), table.into())?;
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

    /// Apply a conditional-highlight patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_cell_conditional_highlights(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }

    fn resolve_conditional_target(
        &self,
        sheet: SheetSelector<'_>,
        table: TableSelector<'_>,
    ) -> Result<(usize, usize)> {
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
