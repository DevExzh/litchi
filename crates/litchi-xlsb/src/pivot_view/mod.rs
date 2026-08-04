//! Layered, lossless PivotTable-view framing for XLSB.
//!
//! The model owns the semantic view binding, while the codec owns bounded
//! BIFF12 framing and lossless stream retention. Workbook, worksheet,
//! relationship, and package orchestration remain in the OOXML host adapter.
//!
//! The PivotTable part is an extensible BIFF12 record collection described by
//! [MS-XLSB] sections 2.1.7.40, 2.4.278, and 2.4.631.

use crate::raw::Error as WireError;
use thiserror::Error as ThisError;

mod codec;
mod model;
#[cfg(test)]
mod tests;

/// Result type for the standalone PivotTable-view codec.
pub type Result<T> = std::result::Result<T, Error>;

/// Error returned by the bounded PivotTable-view codec.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A BIFF12 header or scalar failed raw validation.
    #[error(transparent)]
    Wire(#[from] WireError),
    /// A fixed-width field or enclosing stream boundary is malformed.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A PivotTable framing or identity invariant is invalid.
    #[error("invalid PivotTable view: {0}")]
    Invalid(String),
}

pub use model::Part;

// Historical names remain aliases at the owner facade. New code should use
// the contextual `pivot_view::Part` spelling.
pub type PivotTableViewPart = Part;
