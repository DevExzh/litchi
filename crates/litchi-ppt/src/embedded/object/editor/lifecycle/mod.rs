//! Editor lifecycle operations: opening a package and reading its live state.

pub(super) mod open;
pub(super) mod read;

pub(super) use self::open::{open, open_records};
pub(super) use self::read::{persist_ids, persisted_record};
