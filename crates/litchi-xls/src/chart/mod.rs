//! Layered BIFF8 chart-sheet and embedded-chart owner.
//!
//! The public surface is contextual to an XLS chart: the model lives in the
//! model module, record parsing and replay in the codec module, and workbook
//! hosting in the package module. Chart formulas and cached values remain
//! inert. Unsupported records are retained rather than normalized or evaluated.

pub mod frt;

mod codec;
mod inventory;
mod model;
mod package;
mod wire;

#[cfg(test)]
mod tests;

pub use litchi_ograph::chart::{format, group};
pub use litchi_ograph::record::{chart3d, frame, line, marker, pie, series};

pub use inventory::{ChartInventory, SemanticCompleteness};
pub use model::*;
pub use package::build_workbook;
pub use package::{Editor, Entry};
