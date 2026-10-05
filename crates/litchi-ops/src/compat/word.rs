//! Word-processing element wrappers: [`Paragraph`], [`Run`], [`Table`],
//! [`Row`], [`Cell`], [`CellMerge`] and [`DocumentElement`].
//!
//! Moved from the facade's `document` module (record 0003 §2.4). The
//! [`DocumentView`](crate::DocumentView) family view returns them.
//!
//! **Scheduled for deletion in M1**, with the rest of [`crate::compat`].

mod cell_merge;
mod element;
mod paragraph;
mod run;
mod table;

pub use cell_merge::CellMerge;
pub use element::DocumentElement;
pub use paragraph::Paragraph;
pub use run::Run;
pub use table::{Cell, Row, Table};

#[cfg(feature = "odt")]
use litchi_core::detection::FileFormat;

/// Type an ODT element error: the ODT leaf reports `litchi_core::Error`.
#[cfg(feature = "odt")]
fn odt_error(error: litchi_core::Error) -> crate::Error {
    crate::Error::from(error).with_format(FileFormat::Odt)
}
