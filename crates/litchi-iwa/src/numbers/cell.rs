//! Private Numbers cell migration adapter.
//!
//! The canonical dependency-free cell vocabulary is owned by
//! [`litchi_numbers::cell`]. The monolith keeps these contextual names only so
//! the archive reader, editor, and Pages table bridge can migrate without
//! duplicating the value implementation.

pub(crate) use litchi_numbers::cell::{Update as TableCellUpdate, Value as CellValue};
