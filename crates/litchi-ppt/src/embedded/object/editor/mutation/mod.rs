//! Staged object and persisted-record mutations.

pub(super) mod objects;
pub(super) mod records;

pub(super) use objects::{add, remove, reorder, replace_storage};
pub(super) use records::replace_persisted_record;
