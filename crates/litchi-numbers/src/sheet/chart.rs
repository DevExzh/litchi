//! Sheet-chart transaction families: arrangement, numeric data and the
//! metadata reader's errors.

/// Selector-first exact-source chart arrangement transactions:
/// [`crate::Package::edit_sheet_chart_arrangement`].
pub mod arrangement {
    pub use crate::ChartArrangement;
    pub use crate::package::families::chart_arrangement::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Selector-first exact-source chart numeric-data transactions:
/// [`crate::Package::edit_sheet_chart_data`].
pub mod data {
    pub use crate::package::families::chart_data::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Errors of the read-only chart metadata projection.
pub mod metadata {
    pub use crate::package::families::chart_metadata::{Error, LimitKind};
}

/// Exact-source visible title transactions of a selected sheet chart:
/// [`crate::Package::edit_sheet_chart_title`].
pub mod title {
    pub use crate::package::sheet_chart_title::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}

/// Exact-source category/value axis-title transactions of a selected sheet
/// chart: [`crate::Package::edit_sheet_chart_axis_title`].
pub mod axis_title {
    pub use crate::package::sheet_chart_axis_title::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::chart::axis::Axis;
}

/// Exact-source chart-font identity and point-size transactions:
/// [`crate::Package::edit_sheet_chart_font`].
pub mod font {
    pub use crate::package::sheet_chart_font::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Property, PropertyKind, Refusal,
    };
    pub use litchi_iwa_style::chart_font::ChartFont;
    pub use litchi_iwa_text::character::TextPointSize;
}

/// Exact-source axis-style properties of a chart's primary value or
/// category axis: [`crate::Package::edit_sheet_chart_axis`].
pub mod axis {
    pub use crate::package::sheet_chart_axis::{
        Commit, Diagnostics, Edit, Error, Frequency, Layout, LimitKind, Patch, Property,
        PropertyKind, Refusal,
    };
    pub use litchi_iwa_common::chart::axis::style::Visibility;
    pub use litchi_iwa_common::chart::axis::{Axis, TickMarkLocation};
    pub use litchi_iwa_common::chart::category_labels::Interval;
    pub use litchi_iwa_style::chart_style::axis::{
        GridlineStroke, LabelAffixes, LabelAngle, LabelPosition3d, NumberFormat, Stroke,
    };
}

/// Exact-source per-series chart properties (one value per series):
/// [`crate::Package::edit_sheet_chart_series`].
pub mod series {
    pub use crate::package::sheet_chart_series::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Property, PropertyKind, Refusal,
    };
    pub use litchi_iwa_common::chart::error_bar;
    pub use litchi_iwa_common::chart::number_format::{LabelAffixes, NumberFormat};
    pub use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
    pub use litchi_iwa_common::chart::series::{
        ConnectionLine, MovingAveragePeriod, PolynomialOrder, Stroke, StrokePattern, Symbol,
        SymbolFill, SymbolShape, SymbolSize, Trendline, TrendlineType,
    };
    pub use litchi_iwa_common::chart::series_labels::Visibility;
    pub use litchi_iwa_common::shape::fill::{Fill, Gradient};
}

/// Chart lifecycle of a sheet chart: the native kind and series direction
/// (record 0416). The frame is [`crate::sheet::drawable::properties`], the
/// grid [`data`].
pub mod lifecycle {
    pub use crate::package::sheet_chart_lifecycle::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::chart::Direction;
    pub use litchi_iwa_common::chart::kind::Kind;
}

/// Selector-first exact-source chart-level style transactions on the
/// shared chart style engine: [`crate::Package::edit_sheet_chart_style`].
pub mod style {
    pub use crate::package::sheet_chart_style::{
        Commit, Diagnostics, Drop, Edit, Error, Lighting3d, LimitKind, Patch, Property,
        PropertyKind, RadarGridShape, Refusal, Shadow,
    };
    pub use litchi_iwa_common::chart::{corner, donut, gaps, legend, pie, radar};
    pub use litchi_iwa_common::chart3d;
}

/// Exact-source legend fill, border and shadow transactions of one sheet
/// chart on the shared legend-style variation graph:
/// [`crate::Package::edit_sheet_chart_legend_style`].
pub mod legend_style {
    pub use crate::package::sheet_chart_legend_style::{
        Commit, Diagnostics, Edit, LegendFill, LegendFont, LegendFontSize, LegendShadow,
        LegendStroke, Patch, Property, PropertyKind,
    };
    pub use crate::package::sheet_chart_series::{Error, LimitKind, Refusal};
    pub use litchi_iwa_common::shape::fill::{Fill, Gradient};
    pub use litchi_iwa_common::shape::shadow::Drop;
    pub use litchi_iwa_common::shape::stroke::Stroke;
    pub use litchi_iwa_text::character::TextPointSize;
    pub use litchi_iwa_text::font::Font;
}
