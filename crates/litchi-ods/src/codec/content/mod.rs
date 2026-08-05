//! ODS `content.xml` owner.
//!
//! The historical `crate::codec::content` facade remains the narrow entry
//! point. Streaming XML traversal lives in [`codec`], semantic assembly state
//! lives in [`model`], content-level package joins live in [`package`], and
//! regression coverage lives in [`tests`].

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(crate) use codec::Parser;
pub(crate) use model::{CellBuilder, RowBuilder, SheetBuilder};
