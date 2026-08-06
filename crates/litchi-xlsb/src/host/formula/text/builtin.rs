//! Built-in function metadata used by the formula-text compiler.

use super::super::function_table::BUILTIN_FUNCTIONS;

#[derive(Debug, Clone, Copy)]
pub(super) struct BuiltinFunction {
    pub(super) index: u16,
    pub(super) name: &'static str,
    pub(super) min_args: u8,
    pub(super) max_args: u8,
}

impl BuiltinFunction {
    pub(super) fn accepts_arg_count(self, count: u8) -> bool {
        if count < self.min_args || count > self.max_args {
            return false;
        }
        match self.index {
            // GETPIVOTDATA permits the two mandatory arguments, one optional
            // field, or complete field/item pairs thereafter.
            358 => count <= 3 || count.is_multiple_of(2),
            // COUNTIFS is made solely of range/criteria pairs.
            481 => count.is_multiple_of(2),
            // SUMIFS and AVERAGEIFS have one leading aggregate range followed
            // by range/criteria pairs.
            482 | 484 => !count.is_multiple_of(2),
            _ => true,
        }
    }
}

pub(super) fn builtin_function_by_name(name: &str) -> Option<BuiltinFunction> {
    BUILTIN_FUNCTIONS
        .iter()
        .find_map(|&(index, function_name, min_args, max_args)| {
            function_name
                .eq_ignore_ascii_case(name)
                .then_some(BuiltinFunction {
                    index,
                    name: function_name,
                    min_args,
                    max_args,
                })
        })
}

#[cfg(test)]
pub(crate) fn has_builtin_function(index: u16) -> bool {
    let position = BUILTIN_FUNCTIONS
        .binary_search_by_key(&index, |entry| entry.0)
        .ok();
    position.is_some()
}
