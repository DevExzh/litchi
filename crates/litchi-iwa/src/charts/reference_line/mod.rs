//! Chart reference-line ownership boundary.
//!
//! The semantic values are archive-free and live in `litchi-iwa-common`.
//! This module retains only the IWA archive, wire, graph, and package
//! adapters.

pub(crate) mod archive;

pub(crate) use archive::{
    chart_reference_line_objects, chart_reference_lines, set_chart_reference_lines,
};

pub use litchi_iwa_common::chart::reference_line::{Kind, Line, Value};
