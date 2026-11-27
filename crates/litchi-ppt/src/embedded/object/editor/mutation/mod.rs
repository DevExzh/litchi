//! Staged object and persisted-record mutations.

pub(super) mod objects;
pub(super) mod records;

#[cfg(feature = "equation")]
pub(super) use objects::next_persist_id;
pub(super) use objects::{add, remove, reorder, replace_storage};
#[cfg(feature = "equation")]
pub(super) use records::remove_persisted_record;
pub(super) use records::{insert_persisted_record, replace_persisted_record};
