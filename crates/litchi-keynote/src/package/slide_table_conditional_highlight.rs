//! The ordered conditional-highlight rules of slide-table cells (record
//! 0412f).
//!
//! The rule graph surgery is the shared conditional-highlight engine
//! (`litchi_iwa_table::conditional_highlight`), the engine the Numbers
//! family runs: the conditional-style list, the style set with its pre-pivot
//! and current predicates and formula graphs, the per-rule paragraph and
//! cell styles, the cell's key and applied rule, and a volatile rule's
//! calculation-engine owner. This module resolves the Keynote slide and
//! table selectors and publishes an exact-source patch.

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_table::conditional_highlight::{self as engine, Staged};
use litchi_iwa_table::structure::{Fault, FaultLimit};
use litchi_iwa_table::topology::Target;
use thiserror::Error as ThisError;

use super::Package;
use super::slide_table_cell_rich_text::owner;
use super::slide_table_core as core;
use crate::SlideSelector;
use crate::slide::table::TableSelector;

use litchi_iwa_common::table::cell::conditional_highlight::Rule;

/// A finite resource governed by the transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
    /// A bounded wire operation of the shared engine.
    Wire(litchi_iwa_common::LimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "object identifiers",
            Self::Wire(kind) => return fmt::Display::fmt(kind, formatter),
        })
    }
}

/// A typed, content-redacted slide-table conditional-highlight failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted slide-table conditional-highlight source is invalid")]
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
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selector did not resolve one rooted table.
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
            Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
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

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                    _ => LimitKind::PayloadBytes,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
            Fault::Dependency => Self::Refused(Refusal::UnsupportedDependency),
            Fault::Verification => Self::Verification,
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

/// A reversible, exact-source slide-table conditional-highlight patch.
///
/// The patch retains the exact source and target artifacts; it authorizes
/// application only to its source, and its inverse restores the source byte
/// for byte.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact artifact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics for one slide-table conditional-highlight publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// One validated slide-table conditional-highlight publication.
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

/// One immutable edit that changes the conditional highlighting of cells of
/// one table, applied in order.
pub struct Edit<'a> {
    source: &'a Package,
    slide: usize,
    table: usize,
    rows: u32,
    columns: u32,
    changes: Vec<(At, Staged)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.slide)
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

    /// Stage `rules`, in order, as the cell's conditional highlighting
    /// (replacing any it has). A volatile date rule registers the cell with
    /// the calculation engine's volatile owner family.
    ///
    /// # Errors
    ///
    /// [`Refusal::CellNotFound`], [`Refusal::InvalidRuleCount`] for no rule
    /// or more than fifteen, [`Error::Allocation`].
    pub fn set(self, position: At, rules: impl IntoIterator<Item = Rule>) -> Result<Self, Error> {
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
    pub fn clear(self, position: At) -> Result<Self, Error> {
        self.stage(position, Staged::Clear)
    }

    /// Validate and atomically publish the staged changes, in order. A
    /// change that leaves the cell as it is is skipped.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], the refusals, source, resource and
    /// verification errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (position, change) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let target = locate(
                base,
                SlideSelector::index(self.slide),
                TableSelector::index(self.table),
            )?
            .2;
            if let Some(published) = engine::apply_one(base, target, *position, change)? {
                let expected = match change {
                    Staged::Set(rules) => Some(rules.clone()),
                    Staged::Clear => None,
                };
                let after = locate(
                    &published.package,
                    SlideSelector::index(self.slide),
                    TableSelector::index(self.table),
                )?
                .2;
                if engine::rules_of(&published.package, after, *position)? != expected {
                    return Err(Error::Verification);
                }
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            let owner = owner(self.source);
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
                },
                diagnostics: Diagnostics::default(),
            });
        };
        Ok(Commit {
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(owner(self.source), owner(&package)),
            },
            diagnostics: Diagnostics {
                changed: true,
                touched_components: touched,
            },
            package,
        })
    }
}

/// Resolve a slide table to its slide and table positions and native
/// drawable and model.
fn locate(
    source: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
) -> Result<(usize, usize, Target), Error> {
    let mut budget = core::Budget::new(source).map_err(locate_error)?;
    let selected = core::select_table(source, slide, table, &mut budget).map_err(locate_error)?;
    Ok((
        selected.slide_position.get(),
        selected.table_position.get(),
        Target {
            info: selected.table_info.identifier,
            model: selected.model.identifier,
        },
    ))
}

fn locate_error(error: core::Error) -> Error {
    match error {
        core::Error::SlideNameNotFound
        | core::Error::SlidePositionNotFound(_)
        | core::Error::TablePositionNotFound(_)
        | core::Error::EmptySlideName
        | core::Error::AmbiguousSelector => Error::Refused(Refusal::TableNotFound),
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        core::Error::Allocation(amount) => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

impl Package {
    /// The ordered conditional-highlight rules of one slide-table cell
    /// (`None` when it has none).
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::CellNotFound`],
    /// [`Error::InvalidSource`] for a rule representation the engine does
    /// not read.
    pub fn slide_table_cell_conditional_highlight<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        position: At,
    ) -> Result<Option<Vec<Rule>>, Error> {
        let (_, _, target) = locate(self, slide.into(), table.into())?;
        Ok(engine::rules_of(self, target, position)?)
    }

    /// Start an immutable conditional-highlight edit of one slide table.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`]; source errors.
    pub fn edit_slide_table_cell_conditional_highlights<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = core::Budget::new(self).map_err(locate_error)?;
        let selected = core::select_table(self, slide.into(), table.into(), &mut budget)
            .map_err(locate_error)?;
        Ok(Edit {
            source: self,
            slide: selected.slide_position.get(),
            table: selected.table_position.get(),
            rows: selected.rows,
            columns: selected.columns,
            changes: Vec::new(),
        })
    }

    /// Apply a slide-table conditional-highlight patch to its exact source
    /// package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_slide_table_cell_conditional_highlights(
        &self,
        patch: &Patch,
    ) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_owner(&owner(self)) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::default(),
            });
        }
        let target = patch.artifacts.target_owner();
        let candidate = Package::from_source_with_options(
            std::sync::Arc::from(target.as_ref()),
            self.state.options,
        )
        .map_err(|_error| Error::Verification)?;
        candidate.validate().map_err(|_error| Error::Verification)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components: 0,
            },
        })
    }
}
