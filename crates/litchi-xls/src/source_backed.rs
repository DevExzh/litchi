//! Source-backed XLS edit families (record 0359a F2, the 0355a §1 shape).
//!
//! ```text
//! SourceBackedWorkbook::edit_cells()            -> source_backed::cells::Edit
//! SourceBackedWorkbook::edit_comments()         -> source_backed::comments::Edit
//! SourceBackedWorkbook::edit_sheet_visibility() -> source_backed::sheet_visibility::Edit
//! x::Edit::{verbs..}(&mut self, ..)                 // staging only
//! x::Edit::commit(self) -> Result<x::Commit>        // validates, writes nothing
//! x::Commit = litchi_core::edit::Commit<SourceBackedWorkbook, x::Patch, x::Diagnostics>
//! SourceBackedWorkbook::apply_x(&self, &x::Patch) -> Result<x::Commit>
//! SourceBackedWorkbook::write_to(&self, impl Write) / save(path)   // the only output
//! ```
//!
//! A commit runs every check the former `commit_source_backed` publishers
//! ran (fixed-width value fields, same-length `BoundSheet8`/`TXO`
//! splices, protection, signature, macro and source-freshness refusals, the
//! complete reader's readback of the composed candidate) and returns a new
//! [`crate::SourceBackedWorkbook`] that reads through the plan's composed
//! source and writes the plan: `write_to`/`save` stream exactly the bytes the
//! former publishers streamed, and nothing is materialized. Chained edits
//! work on the committed workbook.
//!
//! A family [`cells::Patch`] (likewise `comments::Patch`,
//! `sheet_visibility::Patch`) is an exact pair: it replays onto a share of the
//! workbook it was committed on, its inverse onto a share of the commit's
//! workbook (restoring the source workbook), and refuses any other workbook
//! with [`crate::Error::PatchConflict`].

use crate::SourceBackedWorkbook;

/// Declare a source-backed family's `Patch`, `Commit`, `LimitKind` and the
/// shared publication and replay helpers over its `Diagnostics`.
macro_rules! source_family {
    ($noun:literal, $operation:literal) => {
        /// The family error: the workbook error of the source or committed
        /// workbook.
        pub use crate::Error;

        #[doc = concat!("A published source-backed ", $noun, " edit: the committed workbook, its exact-pair patch and diagnostics.")]
        pub type Commit = litchi_core::edit::Commit<crate::SourceBackedWorkbook, Patch, Diagnostics>;

        #[doc = concat!(
            "The source-backed ", $noun, " family has no caller-selected limit ",
            "dimension: the committed workbook is opened under the source ",
            "workbook's options."
        )]
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {}

        #[doc = concat!(
            "An exact, reversible source-backed ", $noun, " patch: the pair of ",
            "workbooks it was committed between."
        )]
        #[derive(Clone)]
        pub struct Patch {
            before: crate::SourceBackedWorkbook,
            after: crate::SourceBackedWorkbook,
            diagnostics: Diagnostics,
        }

        impl ::std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                formatter
                    .debug_struct("Patch")
                    .field("is_noop", &self.is_noop())
                    .finish_non_exhaustive()
            }
        }

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.before.same_workbook(&self.after)
            }

            /// The patch that undoes this one, applicable to its result.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self {
                    before: self.after.clone(),
                    after: self.before.clone(),
                    diagnostics: self.diagnostics,
                }
            }
        }

        fn published(
            before: &crate::SourceBackedWorkbook,
            plan: litchi_cfb::ValidatedOverlayPlan,
            diagnostics: Diagnostics,
        ) -> crate::Result<Commit> {
            let after = before.committed(plan).map_err($crate::source_backed::workbook_error)?;
            let patch = Patch {
                before: before.clone(),
                after: after.clone(),
                diagnostics,
            };
            Ok(Commit::new(after, patch, diagnostics))
        }

        fn replay(workbook: &crate::SourceBackedWorkbook, patch: &Patch) -> crate::Result<Commit> {
            if !workbook.same_workbook(&patch.before) {
                return Err(crate::Error::PatchConflict {
                    operation: $operation,
                    reason: "the patch was made from another workbook",
                });
            }
            Ok(Commit::new(
                patch.after.clone(),
                patch.clone(),
                patch.diagnostics,
            ))
        }
    };
}

pub mod cells;
pub mod comments;
pub mod sheet_visibility;

/// A source-backed workbook output failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum OutputError {
    /// The committed workbook's checked overlay publication failed; a sink
    /// failure keeps its typed output progress.
    Overlay(litchi_cfb::OverlayError),
    /// The opened workbook's source could not be streamed (for example it
    /// changed after it was opened).
    Package(crate::package::Error),
}

impl std::fmt::Display for OutputError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Overlay(error) => std::fmt::Display::fmt(error, formatter),
            Self::Package(error) => std::fmt::Display::fmt(error, formatter),
        }
    }
}

impl std::error::Error for OutputError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Overlay(error) => Some(error),
            Self::Package(error) => Some(error),
        }
    }
}

impl From<litchi_cfb::OverlayError> for OutputError {
    fn from(error: litchi_cfb::OverlayError) -> Self {
        Self::Overlay(error)
    }
}

impl From<crate::package::Error> for OutputError {
    fn from(error: crate::package::Error) -> Self {
        Self::Package(error)
    }
}

/// The workbook error behind a source-backed open of a committed workbook.
pub(crate) fn workbook_error(error: crate::SourceBackedError) -> crate::Error {
    match error {
        crate::SourceBackedError::Parse(error) => error,
        crate::SourceBackedError::Cfb(error) => crate::Error::Cfb(error),
        crate::SourceBackedError::Object(error) => crate::Error::Object(error),
        crate::SourceBackedError::Io(error) => crate::Error::Io(error),
        _ => crate::Error::Internal("the committed source-backed workbook could not be reopened"),
    }
}

impl SourceBackedWorkbook {
    /// The source-backed package view of this workbook, the staging base of
    /// a source-backed family edit.
    fn staging_package(&self) -> crate::Result<crate::Package> {
        crate::Package::from_source(self.clone()).map_err(crate::scoped::package_error)
    }
}
