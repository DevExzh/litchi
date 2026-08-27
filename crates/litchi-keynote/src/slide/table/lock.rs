//! Archive-free interactive lock state for a Keynote slide table.
//!
//! The common table vocabulary owns the compact lock value. Keynote owns only
//! selector resolution and the exact-source package transaction; native
//! identifiers, archive objects, protobuf messages, and package bytes remain
//! private to the package adapter.

/// Interactive editing state of a native iWork table.
pub use litchi_iwa_common::table::lock::State;

/// Transaction types for persisted Keynote slide-table lock state.
pub mod transaction {
    pub use crate::package::slide_table_lock_state::{
        SlideTableLockStateCommit as Commit, SlideTableLockStateDiagnostics as Diagnostics,
        SlideTableLockStateEdit as Edit, SlideTableLockStateError as Error,
        SlideTableLockStateLimitKind as LimitKind, SlideTableLockStatePatch as Patch,
        SlideTableLockStatePath as Path,
    };
}
