//! Inert LibreOffice `calcext` sparkline metadata.
//!
//! LibreOffice Calc persists sparklines in the experimental `calcext`
//! namespace (`urn:org:documentfoundation:names:experimental:calc:xmlns:
//! calcext:1.0`) as a `calcext:sparkline-groups` container attached to
//! `table:table`, written after the rows and any conditional formats. Each
//! `calcext:sparkline-group` carries rendering attributes and a
//! `calcext:sparklines` list of `calcext:sparkline` cells with their data
//! ranges.
//!
//! Litchi parses and stores this metadata as typed data only: sparklines are
//! never rendered, cell and range addresses are never resolved, and the
//! `loext:*` theme-color families of the `calcext:sparkline-*-complex-color`
//! children are never resolved against a theme.

use super::structure::validate_cell_range_addresses;
use litchi_core::{Error, Result, xml::escape_xml};

/// Namespace URI of the LibreOffice `loext` extension used by theme colors.
pub(crate) const LOEXT_NAMESPACE_URI: &str =
    "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
/// Namespace declaration written on each complex-color element.
const LOEXT_NAMESPACE_DECLARATION: &str =
    " xmlns:loext=\"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0\"";

/// Most sparkline groups one sheet may declare.
pub(crate) const MAX_SPARKLINE_GROUPS_PER_SHEET: usize = 16_384;
/// Most sparklines one group may contain.
pub(crate) const MAX_SPARKLINES_PER_GROUP: usize = 65_536;
/// Most `loext:transformation` children one complex color may carry.
pub(crate) const MAX_COLOR_TRANSFORMATIONS: usize = 64;
/// Largest accepted length of any single lexical attribute value.
pub(crate) const MAX_SPARKLINE_ATTRIBUTE_BYTES: usize = 64 * 1024;

/// The rendering kind of a sparkline group (`calcext:type`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SparklineType {
    Line,
    Column,
    Stacked,
}

impl SparklineType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "line" => Ok(Self::Line),
            "column" => Ok(Self::Column),
            "stacked" => Ok(Self::Stacked),
            _ => Err(Error::InvalidFormat(format!(
                "invalid calcext sparkline type '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Line => "line",
            Self::Column => "column",
            Self::Stacked => "stacked",
        }
    }
}

/// How a sparkline renders empty source cells (`calcext:display-empty-cells-as`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SparklineEmptyCells {
    Gap,
    Span,
    Zero,
}

impl SparklineEmptyCells {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "gap" => Ok(Self::Gap),
            "span" => Ok(Self::Span),
            "zero" => Ok(Self::Zero),
            _ => Err(Error::InvalidFormat(format!(
                "invalid calcext:display-empty-cells-as value '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Gap => "gap",
            Self::Span => "span",
            Self::Zero => "zero",
        }
    }
}

/// The axis scaling of a sparkline group (`calcext:min-axis-type` /
/// `calcext:max-axis-type`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SparklineAxisType {
    Individual,
    Group,
    Custom,
}

impl SparklineAxisType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "individual" => Ok(Self::Individual),
            "group" => Ok(Self::Group),
            "custom" => Ok(Self::Custom),
            _ => Err(Error::InvalidFormat(format!(
                "invalid calcext axis type '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Individual => "individual",
            Self::Group => "group",
            Self::Custom => "custom",
        }
    }
}

/// Boolean rendering switches of a sparkline group, all optional.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SparklineFlags {
    /// Whether the x axis is a date axis (`calcext:date-axis`).
    pub date_axis: Option<bool>,
    /// Whether every point is marked (`calcext:markers`).
    pub markers: Option<bool>,
    /// Whether the highest point is marked (`calcext:high`).
    pub high: Option<bool>,
    /// Whether the lowest point is marked (`calcext:low`).
    pub low: Option<bool>,
    /// Whether the first point is marked (`calcext:first`).
    pub first: Option<bool>,
    /// Whether the last point is marked (`calcext:last`).
    pub last: Option<bool>,
    /// Whether negative points are marked (`calcext:negative`).
    pub negative: Option<bool>,
    /// Whether the x axis is displayed (`calcext:display-x-axis`).
    pub display_x_axis: Option<bool>,
    /// Whether hidden rows and columns are plotted (`calcext:display-hidden`).
    pub display_hidden: Option<bool>,
    /// Whether the series is plotted right to left (`calcext:right-to-left`).
    pub right_to_left: Option<bool>,
}

/// Optional `#RRGGBB` color slots of a sparkline group.
///
/// Colors are never rendered by litchi. Theme-based colors are modeled
/// separately by [`SparklineComplexColors`].
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SparklineColors {
    /// Series color (`calcext:color-series`).
    pub series: Option<String>,
    /// Negative-points color (`calcext:color-negative`).
    pub negative: Option<String>,
    /// Axis color (`calcext:color-axis`).
    pub axis: Option<String>,
    /// Markers color (`calcext:color-markers`).
    pub markers: Option<String>,
    /// First-point color (`calcext:color-first`).
    pub first: Option<String>,
    /// Last-point color (`calcext:color-last`).
    pub last: Option<String>,
    /// Highest-point color (`calcext:color-high`).
    pub high: Option<String>,
    /// Lowest-point color (`calcext:color-low`).
    pub low: Option<String>,
}

/// The theme color family of a complex color (`loext:theme-type`).
///
/// The family names follow the OOXML theme slots LibreOffice mirrors; litchi
/// never resolves them against a theme.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ThemeColorType {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl ThemeColorType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "dark1" => Self::Dark1,
            "light1" => Self::Light1,
            "dark2" => Self::Dark2,
            "light2" => Self::Light2,
            "accent1" => Self::Accent1,
            "accent2" => Self::Accent2,
            "accent3" => Self::Accent3,
            "accent4" => Self::Accent4,
            "accent5" => Self::Accent5,
            "accent6" => Self::Accent6,
            "hyperlink" => Self::Hyperlink,
            "followed-hyperlink" => Self::FollowedHyperlink,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid loext:theme-type value '{value}'"
                )));
            },
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dark1",
            Self::Light1 => "light1",
            Self::Dark2 => "dark2",
            Self::Light2 => "light2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followed-hyperlink",
        }
    }
}

/// The kind of a color transformation (`loext:type`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColorTransformationType {
    Tint,
    Shade,
    LumMod,
    LumOff,
}

impl ColorTransformationType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "tint" => Ok(Self::Tint),
            "shade" => Ok(Self::Shade),
            "lummod" => Ok(Self::LumMod),
            "lumoff" => Ok(Self::LumOff),
            _ => Err(Error::InvalidFormat(format!(
                "invalid loext transformation type '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Tint => "tint",
            Self::Shade => "shade",
            Self::LumMod => "lummod",
            Self::LumOff => "lumoff",
        }
    }
}

/// One inert `loext:transformation` of a complex color.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SparklineColorTransformation {
    /// The transformation kind (`loext:type`).
    pub transformation_type: ColorTransformationType,
    /// The transformation amount (`loext:value`), an integer in the range
    /// LibreOffice accepts (`i16`).
    pub value: i16,
}

impl SparklineColorTransformation {
    /// Create an inert color transformation.
    pub fn new(transformation_type: ColorTransformationType, value: i16) -> Self {
        Self {
            transformation_type,
            value,
        }
    }
}

/// An inert theme-based `calcext:sparkline-*-complex-color` color.
///
/// The theme family is never resolved against a theme by litchi.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SparklineComplexColor {
    /// The theme color family (`loext:theme-type`).
    pub theme_type: ThemeColorType,
    /// Color transformations in document order.
    pub transformations: Vec<SparklineColorTransformation>,
}

impl SparklineComplexColor {
    /// Create an inert theme-based color without transformations.
    pub fn new(theme_type: ThemeColorType) -> Self {
        Self {
            theme_type,
            transformations: Vec::new(),
        }
    }

    /// Append one color transformation.
    pub fn with_transformation(mut self, transformation: SparklineColorTransformation) -> Self {
        self.transformations.push(transformation);
        self
    }
}

/// Optional theme-based color slots of a sparkline group, one per plain
/// `#RRGGBB` slot of [`SparklineColors`].
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SparklineComplexColors {
    /// Series color (`calcext:sparkline-series-complex-color`).
    pub series: Option<SparklineComplexColor>,
    /// Negative-points color (`calcext:sparkline-negative-complex-color`).
    pub negative: Option<SparklineComplexColor>,
    /// Axis color (`calcext:sparkline-axis-complex-color`).
    pub axis: Option<SparklineComplexColor>,
    /// Markers color (`calcext:sparkline-markers-complex-color`).
    pub markers: Option<SparklineComplexColor>,
    /// First-point color (`calcext:sparkline-first-complex-color`).
    pub first: Option<SparklineComplexColor>,
    /// Last-point color (`calcext:sparkline-last-complex-color`).
    pub last: Option<SparklineComplexColor>,
    /// Highest-point color (`calcext:sparkline-high-complex-color`).
    pub high: Option<SparklineComplexColor>,
    /// Lowest-point color (`calcext:sparkline-low-complex-color`).
    pub low: Option<SparklineComplexColor>,
}

/// The element names of the complex-color slots, in LibreOffice write order.
pub(crate) const COMPLEX_COLOR_SLOTS: [&str; 8] = [
    "sparkline-series-complex-color",
    "sparkline-negative-complex-color",
    "sparkline-axis-complex-color",
    "sparkline-markers-complex-color",
    "sparkline-first-complex-color",
    "sparkline-last-complex-color",
    "sparkline-high-complex-color",
    "sparkline-low-complex-color",
];

impl SparklineComplexColors {
    /// The slot for an element name, or `None` when it is not a complex color.
    pub(crate) fn slot_mut(
        &mut self,
        element_name: &str,
    ) -> Option<&mut Option<SparklineComplexColor>> {
        Some(match element_name {
            "sparkline-series-complex-color" => &mut self.series,
            "sparkline-negative-complex-color" => &mut self.negative,
            "sparkline-axis-complex-color" => &mut self.axis,
            "sparkline-markers-complex-color" => &mut self.markers,
            "sparkline-first-complex-color" => &mut self.first,
            "sparkline-last-complex-color" => &mut self.last,
            "sparkline-high-complex-color" => &mut self.high,
            "sparkline-low-complex-color" => &mut self.low,
            _ => return None,
        })
    }

    /// Assign a slot, rejecting duplicates. `element_name` must be one of
    /// [`COMPLEX_COLOR_SLOTS`].
    pub(crate) fn assign_slot(
        &mut self,
        element_name: &str,
        color: SparklineComplexColor,
    ) -> Result<()> {
        let slot = self.slot_mut(element_name).ok_or_else(|| {
            Error::InvalidFormat(format!("unknown complex color slot '{element_name}'"))
        })?;
        if slot.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate calcext:{element_name} element"
            )));
        }
        *slot = Some(color);
        Ok(())
    }
}

/// One inert `calcext:sparkline` cell assignment.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Sparkline {
    /// Lexical address of the cell hosting the sparkline
    /// (`calcext:cell-address`). Never resolved by litchi.
    pub cell_address: String,
    /// ODF cell-range addresses supplying the data (`calcext:data-range`,
    /// split on unquoted whitespace). Never resolved by litchi.
    pub data_ranges: Vec<String>,
}

/// One inert `calcext:sparkline-group` element of a sheet.
///
/// Numeric values (`calcext:line-width`, `calcext:manual-min`,
/// `calcext:manual-max`) are stored lexically and never rendered.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SparklineGroup {
    /// Optional producer-assigned group identifier (`calcext:id`).
    pub id: Option<String>,
    /// Optional rendering kind (`calcext:type`).
    pub sparkline_type: Option<SparklineType>,
    /// Optional lexical line width with unit, such as `1pt`
    /// (`calcext:line-width`).
    pub line_width: Option<String>,
    /// Optional empty-cell rendering (`calcext:display-empty-cells-as`).
    pub display_empty_cells_as: Option<SparklineEmptyCells>,
    /// Boolean rendering switches.
    pub flags: SparklineFlags,
    /// Optional minimum axis scaling (`calcext:min-axis-type`).
    pub min_axis_type: Option<SparklineAxisType>,
    /// Optional maximum axis scaling (`calcext:max-axis-type`).
    pub max_axis_type: Option<SparklineAxisType>,
    /// Optional lexical custom minimum (`calcext:manual-min`).
    pub manual_min: Option<String>,
    /// Optional lexical custom maximum (`calcext:manual-max`).
    pub manual_max: Option<String>,
    /// Optional `#RRGGBB` color slots.
    pub colors: SparklineColors,
    /// Optional theme-based color slots. Parsed and stored inertly; theme
    /// families are never resolved against a theme by litchi.
    pub complex_colors: SparklineComplexColors,
    /// Sparkline cell assignments in document order.
    pub sparklines: Vec<Sparkline>,
}

impl Sparkline {
    /// Create an inert sparkline cell assignment.
    pub fn new(cell_address: impl Into<String>, data_ranges: Vec<String>) -> Self {
        Self {
            cell_address: cell_address.into(),
            data_ranges,
        }
    }
}

impl SparklineGroup {
    /// Create an inert sparkline group with only cell assignments set.
    pub fn new(sparklines: Vec<Sparkline>) -> Self {
        Self {
            sparklines,
            ..Self::default()
        }
    }

    /// Set the optional group identifier.
    pub fn with_id(mut self, id: impl Into<String>) -> Self {
        self.id = Some(id.into());
        self
    }

    /// Set the optional rendering kind.
    pub fn with_type(mut self, sparkline_type: SparklineType) -> Self {
        self.sparkline_type = Some(sparkline_type);
        self
    }

    /// Set the optional lexical line width, such as `1pt`.
    pub fn with_line_width(mut self, line_width: impl Into<String>) -> Self {
        self.line_width = Some(line_width.into());
        self
    }

    /// Set the optional empty-cell rendering.
    pub fn with_display_empty_cells_as(mut self, display: SparklineEmptyCells) -> Self {
        self.display_empty_cells_as = Some(display);
        self
    }

    /// Set the boolean rendering switches.
    pub fn with_flags(mut self, flags: SparklineFlags) -> Self {
        self.flags = flags;
        self
    }

    /// Set the optional axis scaling types and custom bounds.
    pub fn with_axis(
        mut self,
        min_axis_type: Option<SparklineAxisType>,
        max_axis_type: Option<SparklineAxisType>,
        manual_min: Option<String>,
        manual_max: Option<String>,
    ) -> Self {
        self.min_axis_type = min_axis_type;
        self.max_axis_type = max_axis_type;
        self.manual_min = manual_min;
        self.manual_max = manual_max;
        self
    }

    /// Set the optional color slots.
    pub fn with_colors(mut self, colors: SparklineColors) -> Self {
        self.colors = colors;
        self
    }

    /// Set the optional theme-based color slots.
    pub fn with_complex_colors(mut self, complex_colors: SparklineComplexColors) -> Self {
        self.complex_colors = complex_colors;
        self
    }
}

pub(crate) fn validate_sparkline_group(group: &SparklineGroup) -> Result<()> {
    validate_sparkline_group_attributes(group)?;
    if group.sparklines.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:sparkline-group requires at least one calcext:sparkline".to_string(),
        ));
    }
    if group.sparklines.len() > MAX_SPARKLINES_PER_GROUP {
        return Err(Error::InvalidFormat(format!(
            "sparkline group exceeds the {MAX_SPARKLINES_PER_GROUP} sparkline safety limit"
        )));
    }
    for sparkline in &group.sparklines {
        validate_sparkline(sparkline)?;
    }
    Ok(())
}

/// Validate only the element attributes of a group, not its sparklines.
pub(crate) fn validate_sparkline_group_attributes(group: &SparklineGroup) -> Result<()> {
    if let Some(id) = &group.id {
        validate_attribute_length("calcext:id", id)?;
    }
    if let Some(line_width) = &group.line_width {
        validate_measure("calcext:line-width", line_width)?;
    }
    if let Some(min) = &group.manual_min {
        validate_lexical_number("calcext:manual-min", min)?;
    }
    if let Some(max) = &group.manual_max {
        validate_lexical_number("calcext:manual-max", max)?;
    }
    let colors = &group.colors;
    for (name, color) in [
        ("calcext:color-series", &colors.series),
        ("calcext:color-negative", &colors.negative),
        ("calcext:color-axis", &colors.axis),
        ("calcext:color-markers", &colors.markers),
        ("calcext:color-first", &colors.first),
        ("calcext:color-last", &colors.last),
        ("calcext:color-high", &colors.high),
        ("calcext:color-low", &colors.low),
    ] {
        if let Some(color) = color {
            validate_color(name, color)?;
        }
    }
    let complex_colors = &group.complex_colors;
    for complex_color in [
        &complex_colors.series,
        &complex_colors.negative,
        &complex_colors.axis,
        &complex_colors.markers,
        &complex_colors.first,
        &complex_colors.last,
        &complex_colors.high,
        &complex_colors.low,
    ]
    .into_iter()
    .flatten()
    {
        validate_complex_color(complex_color)?;
    }
    Ok(())
}

pub(crate) fn validate_complex_color(complex_color: &SparklineComplexColor) -> Result<()> {
    if complex_color.transformations.len() > MAX_COLOR_TRANSFORMATIONS {
        return Err(Error::InvalidFormat(format!(
            "complex color exceeds the {MAX_COLOR_TRANSFORMATIONS} transformation safety limit"
        )));
    }
    Ok(())
}

pub(crate) fn validate_sparkline(sparkline: &Sparkline) -> Result<()> {
    if sparkline.cell_address.is_empty() || sparkline.cell_address.trim() != sparkline.cell_address
    {
        return Err(Error::InvalidFormat(format!(
            "invalid calcext:cell-address '{}'",
            sparkline.cell_address
        )));
    }
    validate_attribute_length("calcext:cell-address", &sparkline.cell_address)?;
    if sparkline.data_ranges.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:sparkline requires at least one calcext:data-range".to_string(),
        ));
    }
    validate_cell_range_addresses(&sparkline.data_ranges)?;
    for range in &sparkline.data_ranges {
        validate_attribute_length("calcext:data-range", range)?;
    }
    Ok(())
}

pub(crate) fn validate_sparkline_groups(groups: &[SparklineGroup]) -> Result<()> {
    if groups.len() > MAX_SPARKLINE_GROUPS_PER_SHEET {
        return Err(Error::InvalidFormat(format!(
            "sheet exceeds the {MAX_SPARKLINE_GROUPS_PER_SHEET} sparkline group safety limit"
        )));
    }
    for group in groups {
        validate_sparkline_group(group)?;
    }
    Ok(())
}

fn validate_measure(name: &str, value: &str) -> Result<()> {
    validate_attribute_length(name, value)?;
    let unit_start = value
        .find(|character: char| character.is_ascii_alphabetic())
        .unwrap_or(value.len());
    let (number, unit) = value.split_at(unit_start);
    validate_lexical_number(name, number)?;
    if !unit.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return Err(Error::InvalidFormat(format!(
            "{name} has an invalid unit in '{value}'"
        )));
    }
    Ok(())
}

fn validate_lexical_number(name: &str, value: &str) -> Result<()> {
    let parsed: f64 = value.parse().map_err(|_| {
        Error::InvalidFormat(format!("{name} requires a numeric value, found '{value}'"))
    })?;
    if !parsed.is_finite() {
        return Err(Error::InvalidFormat(format!(
            "{name} requires a finite numeric value, found '{value}'"
        )));
    }
    Ok(())
}

fn validate_color(name: &str, color: &str) -> Result<()> {
    if color.len() != 7
        || !color.starts_with('#')
        || !color[1..].bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(Error::InvalidFormat(format!(
            "invalid {name} color '{color}'"
        )));
    }
    Ok(())
}

fn validate_attribute_length(name: &str, value: &str) -> Result<()> {
    if value.len() > MAX_SPARKLINE_ATTRIBUTE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds the {MAX_SPARKLINE_ATTRIBUTE_BYTES} byte safety limit"
        )));
    }
    Ok(())
}

/// Write a sheet's `calcext:sparkline-groups` container after its conditional
/// formats, matching LibreOffice's element order inside `table:table`.
pub(crate) fn write_sparkline_groups(out: &mut String, groups: &[SparklineGroup]) -> Result<()> {
    validate_sparkline_groups(groups)?;
    if groups.is_empty() {
        return Ok(());
    }
    out.push_str("<calcext:sparkline-groups>");
    for group in groups {
        out.push_str("<calcext:sparkline-group");
        if let Some(id) = &group.id {
            write_attribute(out, "calcext:id", id);
        }
        if let Some(sparkline_type) = group.sparkline_type {
            write_attribute(out, "calcext:type", sparkline_type.as_str());
        }
        if let Some(line_width) = &group.line_width {
            write_attribute(out, "calcext:line-width", line_width);
        }
        write_optional_bool_attribute(out, "calcext:date-axis", group.flags.date_axis);
        if let Some(display) = group.display_empty_cells_as {
            write_attribute(out, "calcext:display-empty-cells-as", display.as_str());
        }
        write_optional_bool_attribute(out, "calcext:markers", group.flags.markers);
        write_optional_bool_attribute(out, "calcext:high", group.flags.high);
        write_optional_bool_attribute(out, "calcext:low", group.flags.low);
        write_optional_bool_attribute(out, "calcext:first", group.flags.first);
        write_optional_bool_attribute(out, "calcext:last", group.flags.last);
        write_optional_bool_attribute(out, "calcext:negative", group.flags.negative);
        write_optional_bool_attribute(out, "calcext:display-x-axis", group.flags.display_x_axis);
        write_optional_bool_attribute(out, "calcext:display-hidden", group.flags.display_hidden);
        if let Some(axis_type) = group.min_axis_type {
            write_attribute(out, "calcext:min-axis-type", axis_type.as_str());
        }
        if let Some(axis_type) = group.max_axis_type {
            write_attribute(out, "calcext:max-axis-type", axis_type.as_str());
        }
        write_optional_bool_attribute(out, "calcext:right-to-left", group.flags.right_to_left);
        if let Some(max) = &group.manual_max {
            write_attribute(out, "calcext:manual-max", max);
        }
        if let Some(min) = &group.manual_min {
            write_attribute(out, "calcext:manual-min", min);
        }
        let colors = &group.colors;
        if let Some(color) = &colors.series {
            write_attribute(out, "calcext:color-series", color);
        }
        if let Some(color) = &colors.negative {
            write_attribute(out, "calcext:color-negative", color);
        }
        if let Some(color) = &colors.axis {
            write_attribute(out, "calcext:color-axis", color);
        }
        if let Some(color) = &colors.markers {
            write_attribute(out, "calcext:color-markers", color);
        }
        if let Some(color) = &colors.first {
            write_attribute(out, "calcext:color-first", color);
        }
        if let Some(color) = &colors.last {
            write_attribute(out, "calcext:color-last", color);
        }
        if let Some(color) = &colors.high {
            write_attribute(out, "calcext:color-high", color);
        }
        if let Some(color) = &colors.low {
            write_attribute(out, "calcext:color-low", color);
        }
        out.push('>');
        let complex_colors = &group.complex_colors;
        for (slot, complex_color) in [
            ("sparkline-series-complex-color", &complex_colors.series),
            ("sparkline-negative-complex-color", &complex_colors.negative),
            ("sparkline-axis-complex-color", &complex_colors.axis),
            ("sparkline-markers-complex-color", &complex_colors.markers),
            ("sparkline-first-complex-color", &complex_colors.first),
            ("sparkline-last-complex-color", &complex_colors.last),
            ("sparkline-high-complex-color", &complex_colors.high),
            ("sparkline-low-complex-color", &complex_colors.low),
        ] {
            if let Some(complex_color) = complex_color {
                write_complex_color(out, slot, complex_color);
            }
        }
        out.push_str("<calcext:sparklines>");
        for sparkline in &group.sparklines {
            out.push_str("<calcext:sparkline calcext:cell-address=\"");
            out.push_str(&escape_xml(&sparkline.cell_address));
            out.push_str("\" calcext:data-range=\"");
            out.push_str(&escape_xml(&sparkline.data_ranges.join(" ")));
            out.push_str("\"/>");
        }
        out.push_str("</calcext:sparklines></calcext:sparkline-group>");
    }
    out.push_str("</calcext:sparkline-groups>");
    Ok(())
}

fn write_complex_color(out: &mut String, slot: &str, complex_color: &SparklineComplexColor) {
    out.push_str("<calcext:");
    out.push_str(slot);
    out.push_str(LOEXT_NAMESPACE_DECLARATION);
    out.push_str(" loext:theme-type=\"");
    out.push_str(complex_color.theme_type.as_str());
    out.push_str("\" loext:color-type=\"theme\"");
    if complex_color.transformations.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for transformation in &complex_color.transformations {
        out.push_str("<loext:transformation loext:type=\"");
        out.push_str(transformation.transformation_type.as_str());
        out.push_str("\" loext:value=\"");
        out.push_str(&transformation.value.to_string());
        out.push_str("\"/>");
    }
    out.push_str("</calcext:");
    out.push_str(slot);
    out.push('>');
}

fn write_attribute(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    out.push_str(&escape_xml(value));
    out.push('"');
}

fn write_optional_bool_attribute(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str(if value { "=\"true\"" } else { "=\"false\"" });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_group() -> SparklineGroup {
        SparklineGroup::new(vec![
            Sparkline::new("Sheet1.A2", vec!["Sheet1.B1:Sheet1.M1".to_string()]),
            Sparkline::new("Sheet1.A3", vec!["Sheet1.B2:Sheet1.M2".to_string()]),
        ])
        .with_id("{1C5C5DE0-3C09-4CB3-A3EC-9E763301EC82}")
        .with_type(SparklineType::Column)
        .with_line_width("1pt")
        .with_display_empty_cells_as(SparklineEmptyCells::Gap)
        .with_flags(SparklineFlags {
            markers: Some(true),
            high: Some(true),
            display_x_axis: Some(true),
            ..SparklineFlags::default()
        })
        .with_axis(
            Some(SparklineAxisType::Custom),
            Some(SparklineAxisType::Individual),
            Some("-5".to_string()),
            None,
        )
        .with_colors(SparklineColors {
            series: Some("#0369a3".to_string()),
            low: Some("#c9211e".to_string()),
            ..SparklineColors::default()
        })
        .with_complex_colors(SparklineComplexColors {
            series: Some(
                SparklineComplexColor::new(ThemeColorType::Accent3).with_transformation(
                    SparklineColorTransformation::new(ColorTransformationType::LumMod, 6000),
                ),
            ),
            last: Some(SparklineComplexColor::new(ThemeColorType::Light1)),
            ..SparklineComplexColors::default()
        })
    }

    #[test]
    fn writes_groups_in_libreoffice_attribute_order() {
        let mut xml = String::new();
        write_sparkline_groups(&mut xml, &[sample_group()]).unwrap();
        assert!(xml.starts_with("<calcext:sparkline-groups>"));
        assert!(xml.contains(
            r##"<calcext:sparkline-group calcext:id="{1C5C5DE0-3C09-4CB3-A3EC-9E763301EC82}" calcext:type="column" calcext:line-width="1pt" calcext:display-empty-cells-as="gap" calcext:markers="true" calcext:high="true" calcext:display-x-axis="true" calcext:min-axis-type="custom" calcext:max-axis-type="individual" calcext:manual-min="-5" calcext:color-series="#0369a3" calcext:color-low="#c9211e">"##
        ));
        assert!(xml.contains(
            r#"<calcext:sparkline calcext:cell-address="Sheet1.A2" calcext:data-range="Sheet1.B1:Sheet1.M1"/>"#
        ));
        assert!(xml.contains(
            r#"<calcext:sparkline-series-complex-color xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" loext:theme-type="accent3" loext:color-type="theme"><loext:transformation loext:type="lummod" loext:value="6000"/></calcext:sparkline-series-complex-color>"#
        ));
        assert!(xml.contains(
            r#"<calcext:sparkline-last-complex-color xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" loext:theme-type="light1" loext:color-type="theme"/>"#
        ));
        assert!(xml.ends_with("</calcext:sparkline-groups>"));
    }

    #[test]
    fn writes_nothing_for_an_empty_collection() {
        let mut xml = String::new();
        write_sparkline_groups(&mut xml, &[]).unwrap();
        assert!(xml.is_empty());
    }

    #[test]
    fn rejects_invalid_groups_sparklines_and_attributes() {
        // Groups require at least one sparkline.
        assert!(validate_sparkline_group(&SparklineGroup::new(Vec::new())).is_err());
        // Sparklines require a cell address and at least one data range.
        assert!(validate_sparkline(&Sparkline::new("", vec![".A1".to_string()])).is_err());
        assert!(validate_sparkline(&Sparkline::new(" A1", vec![".A1".to_string()])).is_err());
        assert!(validate_sparkline(&Sparkline::new(".A1", Vec::new())).is_err());
        // Colors must be #RRGGBB and numbers must be numeric.
        let mut group = sample_group();
        group.colors.series = Some("blue".to_string());
        assert!(validate_sparkline_group(&group).is_err());
        let mut group = sample_group();
        group.manual_min = Some("low".to_string());
        assert!(validate_sparkline_group(&group).is_err());
        let mut group = sample_group();
        group.line_width = Some("wide".to_string());
        assert!(validate_sparkline_group(&group).is_err());
        let mut group = sample_group();
        group.line_width = Some("1px!".to_string());
        assert!(validate_sparkline_group(&group).is_err());
    }
}
