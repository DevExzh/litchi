//! ODS `content.xml` owner.
//!
//! The historical `crate::codec::content` facade remains the narrow entry
//! point. XML traversal lives in [`codec`], parser assembly state lives in
//! [`model`], and the existing regression coverage lives in [`tests`].

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) use codec::Parser;
pub(crate) use model::{CellBuilder, RowBuilder, SheetBuilder};
