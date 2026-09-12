//! Final package topology admission helpers.
//!
//! The implementation remains next to the lifecycle projection because it
//! uses the projection's private source proofs and budget types. This module
//! is the narrow boundary used by the planner and is intentionally separate
//! from source scanning and XML mutation helpers.

pub(super) use super::final_part_budget as admit_parts;
