//! Typed `DrawingML` chart record families.
//!
//! The writer is organized around chart-space, presentation, plot-area,
//! series, axis, legend, and shared-line records. All families share the
//! bounded validation and XML wire primitives in the parent writer module.

mod axis;
mod common;
mod document;
mod legend;
mod plot_area;
mod presentation;
mod series;

pub(super) use document::write_chart_space;
