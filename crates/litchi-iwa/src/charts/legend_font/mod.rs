//! Lossless native chart-legend typography storage and mutation.
//!
//! iWork stores legend typography indirectly: field 2 of the generated legend
//! style selects one entry in the chart's paragraph-style table. Direct font
//! identity, face traits, and size share one private paragraph-style variation.

mod graph;
mod wire;

use crate::charts::{ChartFont, ChartFontSize};

pub(crate) use graph::{
    chart_legend_font, chart_legend_font_size, set_chart_legend_font, set_chart_legend_font_size,
};

/// Exact direct font-identity and face-trait state for a native chart legend.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum ChartLegendFont {
    /// No direct override; iWork resolves the legend's base paragraph style.
    #[default]
    Inherited,
    /// A direct font identity with effective bold and italic traits.
    Font(ChartFont),
}

/// Exact direct font-size state for a native chart legend.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartLegendFontSize {
    /// No direct override; iWork resolves the legend's base paragraph style.
    #[default]
    Inherited,
    /// A direct positive point size.
    Size(ChartFontSize),
}
