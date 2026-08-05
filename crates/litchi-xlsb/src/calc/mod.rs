//! Strict workbook calculation properties (`BrtCalcProp`).
//!
//! Every [`Props`] value is directly writable: its fields are private and its
//! numeric inputs are represented by checked types. Short setters and
//! consuming builders keep ordinary authoring compact:
//!
//! ```
//! use litchi_xlsb::calc::{Delta, Mode, Opts, Props, Threads};
//!
//! # fn example() -> litchi_xlsb::calc::Result<()> {
//! let mut props = Props::new()
//!     .with_mode(Mode::Manual)
//!     .with_iters(25)
//!     .with_delta(Delta::new(0.000_01)?)
//!     .with_threads(Threads::new(4)?)
//!     .with_opts(Opts::A1 | Opts::ITERATE | Opts::MTR | Opts::USER_THREADS)?;
//!
//! props.set_mode(Mode::Auto).set_id(0x0001_EB1D);
//! assert_eq!(props.threads().get(), 4);
//! assert!(props.has(Opts::ITERATE));
//! assert!(Threads::new(0).is_err());
//! assert!(Delta::new(-0.0).is_err());
//! # Ok(())
//! # }
//! ```
//!
//! Invalid numeric state cannot be constructed by bypassing the checked
//! constructors:
//!
//! ```compile_fail
//! use litchi_xlsb::calc::Threads;
//!
//! let _ = Threads(0);
//! ```
//!
//! ```compile_fail
//! use litchi_xlsb::calc::Delta;
//!
//! let _ = Delta(-0.0);
//! ```

mod codec;
mod model;

#[cfg(test)]
mod tests;

use thiserror::Error;

/// Exact byte length of a conforming `BrtCalcProp` payload.
pub const LEN: usize = 26;

/// Result of reading or writing calculation properties.
pub type Result<T> = std::result::Result<T, Error>;

/// A typed `BrtCalcProp` failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The BIFF12 payload was truncated, had trailing bytes, or could not be
    /// written to its destination.
    #[error("invalid BrtCalcProp wire payload: {0}")]
    Wire(#[from] crate::raw::Error),

    /// `fAutoRecalc` was outside its closed enumeration.
    #[error("invalid BrtCalcProp calculation mode {value}")]
    Mode {
        /// Rejected wire value.
        value: u32,
    },

    /// A flag outside the nine defined option bits was set.
    #[error("BrtCalcProp contains reserved option bits {bits:#06x}")]
    ReservedOpts {
        /// Reserved bits that were set.
        bits: u16,
    },

    /// `xnumDelta` violated the strict `Xnum` domain.
    #[error("invalid BrtCalcProp xnumDelta bit pattern {bits:#018x}")]
    Delta {
        /// Rejected IEEE-754 bit pattern.
        bits: u64,
    },

    /// `cUserThreadCount` was outside the specification's inclusive range.
    #[error("BrtCalcProp thread count {value} is outside 1..=1024")]
    Threads {
        /// Rejected signed wire value.
        value: i32,
    },
}

pub use codec::{read, write};
pub use model::{Delta, Mode, Opts, Props, Threads};
