//! Lossless structural edits of worksheet cell records.
//!
//! The owner inventories all ordinary scalar cell families from `[MS-XLSB]`
//! sections 2.4.319-330 plus formula cached-result records from sections
//! 2.4.684-687. Existing values retain their BIFF12 record family, while
//! explicit insert/remove operations author or remove complete records.
//! Length-changing strings and inert formula token/cache changes rebuild only
//! selected records; unknown records remain byte-exact. Row spans and worksheet
//! dimensions expand with structural insertions, and whole-workbook publication
//! validates style, shared-string, formula, and package dependency closure.
//! The worksheet and structural families carry deterministic semantic deltas,
//! bounded durable transfer (including ordinary drawing graphs), three-way
//! conflict reporting, and undo/redo. Package-owned cell-value patches carry
//! exact whole-package images and their exact inverse.
//!
//! Package-owned cell edits use the canonical commit triad:
//!
//! ```
//! # fn main() -> litchi_xlsb::package::PackageResult<()> {
//! use litchi_xlsb::Package;
//!
//! let source = Package::new()?;
//! let edit = source.edit_cell_values(0usize)?.resolve()?.expect("worksheet");
//! let commit = edit.commit()?;
//! assert_eq!(commit.package().to_bytes()?, commit.patch().after());
//! assert_eq!(commit.diagnostics().changed(), !commit.patch().is_noop());
//! let _package = commit.into_package();
//! # Ok(())
//! # }
//! ```

mod drawing_transfer;
mod format;
mod formula_transfer;
mod names;
mod package;
pub(crate) mod resources;
mod scalar_sst;
pub mod structure;
pub(crate) mod workbook;
pub mod worksheet;

pub use drawing_transfer::DrawingTransferRefusal;
pub use format::{BorderLine, Format};
pub use package::{Commit, Diagnostics, Edit, EditOptions, Patch, Scalar};
pub use worksheet::{
    CellError, CellFormula, Change, LimitKind, Limits, Number, Reference, Snapshot, StoredCell,
    StyleIndex, TransferLimits, Value,
};

pub use crate::package::error::{Error, Result};
