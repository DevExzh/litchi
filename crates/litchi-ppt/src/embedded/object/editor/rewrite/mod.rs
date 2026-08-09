//! Lossless `PowerPoint` record rewriting with bounded recursion.

pub(super) mod nested;
pub(super) mod record;

#[cfg(test)]
pub(super) use nested::MAX_NESTED_RECORD_DEPTH;
pub(super) use nested::{external_object_list, replace_nested_record};
pub(super) use record::{slice, type_of, u32_at};
