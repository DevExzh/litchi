//! Numbers formula function token registry.
//!
//! The canonical native mapping lives in [`litchi_numbers_wire::function_map`].
//! This module retains the host crate visibility used by the legacy Numbers
//! readers and writers.

pub(crate) use litchi_numbers_wire::function_map::{function_identifier, function_name};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_observed_function_tokens_both_ways() {
        assert_eq!(function_name(168), Some("SUM"));
        assert_eq!(function_name(62), Some("IF"));
        assert_eq!(function_identifier("sum"), Some(168));
        assert_eq!(function_identifier("XLOOKUP"), Some(314));
        assert_eq!(function_name(999), None);
    }
}
