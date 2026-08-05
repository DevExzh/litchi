//! Typed ODF chart model used by standalone and embedded chart owners.

use super::data::CachedTable;
use super::extensions::Extensions;
use litchi_odf_common::calculation::Settings;
use litchi_odf_common::chart::{Class, Dimension, Labels, Position};

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Text {
    pub text: String,
    pub cell_range: Option<String>,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub extensions: Extensions,
}

impl Text {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegendSpec {
    pub position: Position,
    pub style_name: Option<String>,
    pub title: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub expansion: Option<String>,
    pub expansion_aspect_ratio: Option<String>,
    pub extensions: Extensions,
}

impl Default for LegendSpec {
    fn default() -> Self {
        Self {
            position: Position::End,
            style_name: None,
            title: None,
            x: None,
            y: None,
            expansion: None,
            expansion_aspect_ratio: None,
            extensions: Extensions::default(),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleElement {
    pub style_name: Option<String>,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GridSpec {
    pub class: Class,
    pub style_name: Option<String>,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DataLabelSpec {
    pub text: Option<String>,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataPointSpec {
    pub repeated: u32,
    pub style_name: Option<String>,
    pub label: Option<DataLabelSpec>,
    pub extensions: Extensions,
}

impl Default for DataPointSpec {
    fn default() -> Self {
        Self {
            repeated: 1,
            style_name: None,
            label: None,
            extensions: Extensions::default(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DomainSpec {
    pub cell_range_address: String,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EquationSpec {
    pub display_equation: bool,
    pub display_r_square: bool,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RegressionSpec {
    pub style_name: Option<String>,
    pub equation: Option<EquationSpec>,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SeriesSpec {
    pub xml_id: Option<String>,
    pub class: Option<String>,
    pub values_cell_range_address: Option<String>,
    pub label_cell_address: Option<String>,
    pub attached_axis: Option<String>,
    pub style_name: Option<String>,
    pub domains: Vec<DomainSpec>,
    pub data_points: Vec<DataPointSpec>,
    pub data_label: Option<DataLabelSpec>,
    pub mean_value: Option<StyleElement>,
    pub error_indicator: Option<StyleElement>,
    pub regression_curves: Vec<RegressionSpec>,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AxisSpec {
    pub dimension: Dimension,
    pub name: Option<String>,
    pub style_name: Option<String>,
    pub title: Option<Text>,
    pub categories_cell_range_address: Option<String>,
    pub grids: Vec<GridSpec>,
    pub extensions: Extensions,
}

impl AxisSpec {
    pub fn new(dimension: Dimension) -> Self {
        Self {
            dimension,
            name: None,
            style_name: None,
            title: None,
            categories_cell_range_address: None,
            grids: Vec::new(),
            extensions: Extensions::default(),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PlotAreaSpec {
    pub cell_range_address: Option<String>,
    pub data_source_labels: Option<Labels>,
    pub style_name: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    pub axes: Vec<AxisSpec>,
    pub series: Vec<SeriesSpec>,
    pub wall: Option<StyleElement>,
    pub floor: Option<StyleElement>,
    pub stock_gain_marker: Option<StyleElement>,
    pub stock_loss_marker: Option<StyleElement>,
    pub stock_range_line: Option<StyleElement>,
    pub extensions: Extensions,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Definition {
    pub class: String,
    pub style_name: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    pub title: Option<Text>,
    pub subtitle: Option<Text>,
    pub footer: Option<Text>,
    pub legend: Option<LegendSpec>,
    pub plot_area: PlotAreaSpec,
    pub cached_table: Option<CachedTable>,
    /// Inert cached-formula recalculation metadata; never executed by this crate.
    pub calculation_settings: Option<Settings>,
    pub extensions: Extensions,
}

impl Definition {
    pub fn new(class: impl Into<String>) -> Self {
        Self {
            class: class.into(),
            style_name: None,
            width: None,
            height: None,
            title: None,
            subtitle: None,
            footer: None,
            legend: None,
            plot_area: PlotAreaSpec::default(),
            cached_table: None,
            calculation_settings: None,
            extensions: Extensions::default(),
        }
    }
}
