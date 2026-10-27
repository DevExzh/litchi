//! Resizing a table at its trailing edges, started by
//! [`crate::Package::edit_table_resize`] (record 0370e).

pub use crate::package::table_resize::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
};
