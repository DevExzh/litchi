//! Typed Excel Binary Workbook documents.
//!
//! [`raw`] owns the validated BIFF12 wire substrate, while [`calc`] owns the
//! strictly typed workbook calculation record. Additional semantic snapshots
//! and edits will be layered over them without exposing package identifiers in
//! ordinary APIs.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]
#![deny(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "ordinary and newly authored workbook paths require checked wire conversions, typed error propagation, and explicit variants"
)]
#![cfg_attr(
    test,
    allow(
        clippy::cloned_ref_to_slice_refs,
        clippy::expect_used,
        clippy::inconsistent_digit_grouping,
        clippy::items_after_test_module,
        reason = "unit tests use panic-on-failure extraction and literal wire-layout assertions to keep fixtures compact"
    )
)]
// These style-only rewrites touch much of the legacy codec and obscure review of
// semantic changes; keep them local to this crate rather than weakening workspace
// policy for newer crates.
#![allow(
    clippy::derivable_impls,
    clippy::double_must_use,
    clippy::drop_non_drop,
    clippy::manual_contains,
    clippy::needless_lifetimes,
    clippy::needless_question_mark,
    clippy::needless_range_loop,
    clippy::nonminimal_bool,
    clippy::question_mark,
    clippy::redundant_closure,
    clippy::useless_conversion,
    clippy::while_let_on_iterator,
    reason = "style-only legacy codec cleanup is isolated from behavior-preserving strict-lint adoption"
)]
// BIFF12 terminology and record layouts naturally trigger naming, layout, and
// boolean-density heuristics. Renaming these public concepts would reduce fidelity
// to the specification and can break downstream imports.
#![allow(
    clippy::module_inception,
    reason = "names and dense flag structures intentionally follow BIFF12 records and existing public API vocabulary"
)]

pub mod auto_filter;
pub mod binary_index;
pub mod builder;
pub mod calc;
pub mod calculation_chain;
pub mod cell_values;
pub mod cell_watches;
pub mod chart;
pub mod comments;
pub mod conditional_formatting;
pub mod custom_data;
pub mod data_model;
pub mod data_validation;
pub mod date_utils;
pub mod drawing;
pub mod external_link;
pub mod formula;
pub mod hyperlinks;
pub mod leniency;
pub mod merged_cells;
pub mod metadata;
pub mod named_ranges;
pub mod opc;
mod open_options;
pub mod package;
pub mod page_breaks;
pub(crate) mod part_edit;
mod pivot_chart;
pub mod pivot_view;
pub mod raw;
pub mod ribbon;
mod semantic_limits;
pub mod shapes;
pub mod shared_workbook;
pub mod sheet;
pub(crate) mod signature_policy;
pub mod slicer;
pub mod sparkline;
pub mod styles;
pub mod task_panes;
#[cfg(all(test, feature = "vba-inspection"))]
mod test_support;
pub mod theme;
pub mod timeline;
#[cfg(feature = "vba-inspection")]
pub mod vba;
pub mod volatile_dependencies;
pub mod workbook;
pub(crate) mod worksheet_index;
pub mod xml_maps;

/// OPC resource limits and diagnostics used by XLSB package and workbook ingress.
pub use litchi_opc::{SourceCacheDiagnostics, SourceCacheLimits, SourceLimitKind, SourceLimits};
pub(crate) mod writer;

/// Canonical builder for authoring a new XLSB package.
pub use builder::Builder;

pub use open_options::{OpenOptions, OpenPolicy};
pub use raw::Error;
pub use semantic_limits::SemanticLimits;

pub use leniency::{
    CachedNumberDefect, Defect, LayoutDefect, Leniency, ReservedFieldDefect, ToleranceReport,
    ToleratedDefect, ValueEncodingDefect,
};

pub use package::scenarios;
pub use package::{Cell, FormulaOpacityReason, FormulaResolutionStatus, Package};
pub use sheet::Worksheet;
pub use workbook::{
    SourceBackedExternalLink, SourceBackedIndexedWorksheet, SourceBackedWorkbook,
    SourceBackedWorksheet, Workbook,
};

pub use pivot_view::Part;

pub use data_validation::{FormulaBinary, RecordKind, Settings, Validation};
pub use formula::ptg_types;
pub use formula::{
    ArrayValue, BinaryOperator, Compiler, Error as FormulaError, ExternalTableReference, Group,
    GroupKind, MAX_CELL_FORMULA_BYTES, MemoryKind, ParsedFormula, Parser, Range, Resolution,
    Result as FormulaResult, TableColumns, TableDataType, TableNamedColumns, TableReference,
    TableRowType, Token, UnaryOperator,
};
