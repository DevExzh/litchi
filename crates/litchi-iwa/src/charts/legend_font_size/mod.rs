//! Lossless native chart-legend font-size storage and mutation.
//!
//! iWork stores legend typography indirectly: field 2 of the generated legend
//! style selects one entry in the chart's paragraph-style table. A direct font
//! size is a private paragraph-style variation appended to that table.

mod graph;
mod wire;

use crate::charts::ChartFontSize;

pub(crate) use graph::{chart_legend_font_size, set_chart_legend_font_size};

/// Exact direct font-size state for a native chart legend.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartLegendFontSize {
    /// No direct override; iWork resolves the legend's base paragraph style.
    #[default]
    Inherited,
    /// A direct positive point size.
    Size(ChartFontSize),
}
