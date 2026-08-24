//! Lossless title visibility and outline settings for Pages body tables.
//!
//! The compact semantic value preserves native optional-field presence:
//! `None` is absent, while `Some(false)` is an explicit native false.

pub use crate::package::body_table_title::{
    BodyTableTitleCommit as Commit, BodyTableTitleDiagnostics as Diagnostics,
    BodyTableTitleEdit as Edit, BodyTableTitleError as Error, BodyTableTitleLimitKind as LimitKind,
    BodyTableTitlePatch as Patch,
};
pub use litchi_iwa_common::table::title::Settings;
