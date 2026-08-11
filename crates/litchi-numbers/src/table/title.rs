//! Lossless table-title values and exact-source transactions.
//!
//! Start from [`crate::Package::table_title_settings`] to read one title, or
//! [`crate::Package::edit_table_title`] to stage a replacement. Both methods
//! require a [`crate::SheetSelector`] followed by a sheet-scoped
//! [`crate::TableSelector`]; native object identifiers and raw package bytes
//! are intentionally absent from this API.
//!
//! [`Settings`](crate::table::title::Settings) preserves each optional
//! Boolean's native presence: `None` is
//! absent on the wire, whereas `Some(false)` is an explicitly stored false.
//! Use [`Edit::set`](crate::table::title::Edit::set) to replace the complete
//! lossless value, then [`Edit::commit`](crate::table::title::Edit::commit) to
//! publish it. An unchanged value returns an exact no-op; a changed commit
//! produces a source-bound [`Patch`](crate::table::title::Patch), removes every
//! existing canonical root preview, and can be reversed with
//! [`Patch::inverse`](crate::table::title::Patch::inverse) plus
//! [`crate::Package::apply_table_title`].
//!
//! Changed publication refuses locked tables. A setting with effective title
//! visibility also requires valid native title height, paragraph-style, and
//! shape-style prerequisites. These checks preserve the source's native
//! structure without exposing it as an editable public detail.

pub use crate::package::table_title::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path};
pub use litchi_iwa_common::table::title::Settings;
