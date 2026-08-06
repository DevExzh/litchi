//! Layered semantic chart-reader implementation.
//!
//! Each included owner below keeps one coherent part of the DrawingML chart
//! grammar in a separate source unit while sharing the same borrowed XML
//! reader and typed validation context.

use super::super::super::model::{ChartXmlReader, IGNORED_NAMESPACE_ELEMENT};
use super::super::validation::{
    bounded_percentage_i32_attr, bounded_percentage_u32_attr, bounded_u32_attr, get_attr,
    invalid_attribute, missing_attribute, optional_bool_attr, optional_i32_attr, optional_u32_attr,
    parse_axis_cross_between, parse_axis_cross_mode, parse_axis_label_align,
    parse_axis_orientation, parse_axis_position, parse_bool_attr, parse_built_in_unit,
    parse_data_label_position, parse_grouping, parse_marker_style, parse_number_format,
    parse_tick_label_position, parse_tick_mark, parse_time_unit, required_enum_attr,
    required_f64_attr, required_named_f64_attr, required_nonnegative_f64_attr,
    required_positive_f64_attr, required_positive_u32_attr, required_string_attr,
    required_u32_attr,
};
use super::super::xml::consume_empty_chart_element;

use crate::chart::axis::{
    Axis, AxisCommon, AxisCrossBetween, AxisCrossMode, CategoryAxis, DateAxis, DisplayUnits,
    SeriesAxis, ValueAxis,
};
use crate::chart::bubble::{Scale as BubbleScale, Size as BubbleSize};
use crate::chart::data::{
    DataSourceRef, Layout, NumberFormat, NumericData, RichText, StringData, TitleText,
};
use crate::chart::legend::{Legend, LegendEntry};
use crate::chart::model::{
    ColorMapOverride, ColorMapping, ColorSchemeIndex, ExtensionList, ExternalData, HeaderFooter,
    PageMargins, PageOrientation, PageSetup, PictureFormat, PictureOptions, PivotFormat,
    PivotSource, PrintSettings, Protection, ShapeProperties, TextProperties, View3D, WallFloor,
};
use crate::chart::plot_area::{
    Area3DTypeGroup, AreaTypeGroup, BandFormat, Bar3DTypeGroup, BarShape, BarTypeGroup,
    BubbleTypeGroup, DataTable, DoughnutTypeGroup, Line3DTypeGroup, LineTypeGroup, Lines,
    OfPieTypeGroup, Pie3DTypeGroup, PieTypeGroup, PlotArea, RadarTypeGroup, ScatterTypeGroup,
    StockTypeGroup, Surface3DTypeGroup, SurfaceTypeGroup, TypeGroup, TypeGroupCommon, UpDownBars,
};
use crate::chart::series::{
    DataLabel, DataLabels, DataPoint, ErrorBar, ErrorBarDirection, ErrorBarType, ErrorBarValueType,
    Marker, Series, Trendline, TrendlineType,
};
use crate::chart::types::{
    AxisOrientation, AxisPosition, BarDirection, BarGrouping, LayoutMode, LayoutTarget,
    LegendPosition, OfPieSplitType, OfPieType, RadarStyle, ScatterStyle, TickLabelPosition,
    TickMark,
};
use crate::{Error, Result};
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use std::io::BufRead;

mod axis;
mod document;
mod legend;
mod plot_area;
mod series;

pub(crate) use axis::*;
pub(crate) use document::*;
pub(crate) use legend::*;
pub(crate) use plot_area::*;
pub(crate) use series::*;
