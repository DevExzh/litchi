//! Numbers formula function token registry.
//!
//! The canonical native mapping lives in [`litchi_numbers_wire::function_map`].
//! This module retains the package-private visibility used by the standalone
//! Numbers parser and formula author.

pub(super) use litchi_numbers_wire::function_map::{function_identifier, function_name};
