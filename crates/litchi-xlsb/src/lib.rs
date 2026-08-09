//! Typed Excel Binary Workbook documents.
//!
//! [`raw`] owns the validated BIFF12 wire substrate, while [`calc`] owns the
//! strictly typed workbook calculation record. Additional semantic snapshots
//! and edits will be layered over them without exposing package identifiers in
//! ordinary APIs.

#![forbid(unsafe_code)]
#![cfg_attr(
    test,
    allow(
        clippy::cloned_ref_to_slice_refs,
        clippy::decimal_bitwise_operands,
        clippy::default_constructed_unit_structs,
        clippy::default_trait_access,
        clippy::field_reassign_with_default,
        clippy::identity_op,
        clippy::inconsistent_digit_grouping,
        clippy::items_after_test_module,
        clippy::manual_string_new,
        clippy::match_wildcard_for_single_variants,
        clippy::print_stdout,
        clippy::unwrap_used,
        reason = "unit tests use panic-on-failure extraction and literal wire-layout assertions to keep fixtures compact"
    )
)]
// This crate models an existing binary format and exposes a broad, stable API. The
// following compatibility lints are intentionally deferred until their suggested
// signature and documentation changes can be made as a dedicated API migration.
#![allow(
    clippy::doc_markdown,
    clippy::missing_errors_doc,
    clippy::missing_panics_doc,
    clippy::must_use_candidate,
    clippy::needless_pass_by_value,
    clippy::return_self_not_must_use,
    clippy::trivially_copy_pass_by_ref,
    clippy::unused_self,
    reason = "retrofitting these API lints would churn public signatures and documentation independently of BIFF12 correctness"
)]
// Parser and writer conversions below mirror BIFF12's fixed-width fields. Changing
// their failure or wrapping behavior as lint cleanup would be a wire-format change.
#![allow(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unnecessary_unwrap,
    clippy::wildcard_enum_match_arm,
    reason = "BIFF12 decoding deliberately preserves fixed-width casts, exact sentinels, and established error mapping"
)]
// These style-only rewrites touch much of the legacy codec and obscure review of
// semantic changes; keep them local to this crate rather than weakening workspace
// policy for newer crates.
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::assigning_clones,
    clippy::bool_to_int_with_if,
    clippy::collapsible_if,
    clippy::derivable_impls,
    clippy::double_must_use,
    clippy::drop_non_drop,
    clippy::elidable_lifetime_names,
    clippy::format_push_string,
    clippy::if_not_else,
    clippy::implicit_clone,
    clippy::inconsistent_struct_constructor,
    clippy::items_after_statements,
    clippy::manual_contains,
    clippy::manual_is_multiple_of,
    clippy::manual_let_else,
    clippy::map_unwrap_or,
    clippy::match_same_arms,
    clippy::missing_fields_in_debug,
    clippy::needless_lifetimes,
    clippy::needless_question_mark,
    clippy::needless_range_loop,
    clippy::nonminimal_bool,
    clippy::question_mark,
    clippy::redundant_closure,
    clippy::redundant_closure_for_method_calls,
    clippy::ref_option,
    clippy::semicolon_if_nothing_returned,
    clippy::single_match_else,
    clippy::unnecessary_wraps,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::useless_conversion,
    clippy::verbose_bit_mask,
    clippy::while_let_on_iterator,
    reason = "style-only legacy codec cleanup is isolated from behavior-preserving strict-lint adoption"
)]
// BIFF12 terminology and record layouts naturally trigger naming, layout, and
// boolean-density heuristics. Renaming these public concepts would reduce fidelity
// to the specification and can break downstream imports.
#![allow(
    clippy::fn_params_excessive_bools,
    clippy::module_inception,
    clippy::module_name_repetitions,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::struct_excessive_bools,
    clippy::wildcard_imports,
    reason = "names and dense flag structures intentionally follow BIFF12 records and existing public API vocabulary"
)]

pub mod calc;
pub mod cell_values;
pub mod cell_watches;
pub mod chart;
pub mod comments;
pub mod conditional_formatting;
pub mod data_validation;
pub mod date_utils;
pub mod external_link;
pub mod formula;
pub mod hyperlinks;
pub mod merged_cells;
pub mod named_ranges;
pub mod package;
mod pivot_chart;
pub mod pivot_view;
pub mod raw;
pub mod shapes;
pub mod shared_workbook;
pub mod sheet;
pub mod slicer;
pub mod sparkline;
pub mod styles;
pub mod timeline;
pub mod workbook;
pub mod xml_maps;

/// OPC resource limits used by XLSB package and workbook ingress.
pub use litchi_opc::ReadLimits;
pub mod writer;

pub use raw::Error;

pub use package::Package;
pub use package::scenarios;
pub use sheet::Worksheet;
pub use workbook::Workbook;

pub use pivot_view::Part;

pub use data_validation::{FormulaBinary, RecordKind, Settings, Validation};
pub use formula::ptg_types;
pub use formula::{
    ArrayValue, BinaryOperator, Compiler, Error as FormulaError, ExternalTableReference, Group,
    GroupKind, MAX_CELL_FORMULA_BYTES, MemoryKind, ParsedFormula, Parser, Range, Resolution,
    Result as FormulaResult, TableColumns, TableDataType, TableNamedColumns, TableReference,
    TableRowType, Token, UnaryOperator,
};
