//! Lossless chart record parsing and encoding.
//!
//! The parent chart module keeps this private codec facade at
//! [`crate::chart::codec`]. Parser state is kept in [`model`], while BIFF
//! traversal and emission live in [`records`].

mod model;
mod records;

pub(crate) use records::{PLOT_AREA, SHT_PROPS, valid_props};
pub(super) use records::{encode, parse, patch};
