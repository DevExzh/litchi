//! Inert LibreOffice `calcext` conditional-format metadata.
//!
//! LibreOffice Calc persists conditional formatting in the experimental
//! `calcext` namespace (`urn:org:documentfoundation:names:experimental:calc:
//! xmlns:calcext:1.0`) as a `calcext:conditional-formats` container attached to
//! `table:table`. Each `calcext:conditional-format` names the cell ranges it
//! covers and carries an ordered list of rules: `calcext:condition` expression
//! rules, `calcext:color-scale` gradients, `calcext:data-bar` bars,
//! `calcext:icon-set` icon thresholds, and `calcext:date-is` date rules.
//!
//! Litchi parses and stores this metadata as typed data only: conditions and
//! formulas are never evaluated, style references are never resolved or
//! applied, and colors and thresholds are never rendered.

use super::structure::validate_cell_range_addresses;
use litchi_core::{Error, Result, xml::escape_xml};

/// Namespace URI of the LibreOffice `calcext` extension.
pub(crate) const CALCEXT_NAMESPACE_URI: &str =
    "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
/// Namespace declaration written when a document contains conditional formats.
pub(crate) const CALCEXT_NAMESPACE_DECLARATION: &str =
    " xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\"";

/// Most conditional formats one sheet may declare.
pub(crate) const MAX_CONDITIONAL_FORMATS_PER_SHEET: usize = 16_384;
/// Most inert rules one conditional format may carry.
pub(crate) const MAX_RULES_PER_FORMAT: usize = 1_024;
/// Most threshold entries one color scale or icon set may carry.
pub(crate) const MAX_ENTRIES_PER_RULE: usize = 256;
/// Number of limit entries LibreOffice writes for a data bar.
pub(crate) const DATA_BAR_ENTRY_COUNT: usize = 2;
/// Largest accepted length of any single lexical attribute value.
pub(crate) const MAX_CONDITIONAL_ATTRIBUTE_BYTES: usize = 64 * 1024;

/// One inert `calcext:conditional-format` element of a sheet.
///
/// Rules are retained in document order. Litchi does not evaluate them or
/// compute an effective style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Format {
    /// ODF cell-range addresses the format applies to
    /// (`calcext:target-range-address`, split on unquoted whitespace).
    pub target_range_addresses: Vec<String>,
    /// Inert rules in document order.
    pub rules: Vec<Rule>,
}

/// One inert rule of a `calcext:conditional-format` element.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Rule {
    /// A `calcext:condition` expression rule.
    Condition(Condition),
    /// A `calcext:color-scale` gradient rule.
    ColorScale(ColorScale),
    /// A `calcext:data-bar` rule.
    DataBar(DataBar),
    /// A `calcext:icon-set` rule.
    IconSet(IconSet),
    /// A `calcext:date-is` date rule.
    DateIs(DateIs),
}

impl From<Condition> for Rule {
    fn from(condition: Condition) -> Self {
        Self::Condition(condition)
    }
}

impl From<ColorScale> for Rule {
    fn from(color_scale: ColorScale) -> Self {
        Self::ColorScale(color_scale)
    }
}

impl From<DataBar> for Rule {
    fn from(data_bar: DataBar) -> Self {
        Self::DataBar(data_bar)
    }
}

impl From<IconSet> for Rule {
    fn from(icon_set: IconSet) -> Self {
        Self::IconSet(icon_set)
    }
}

impl From<DateIs> for Rule {
    fn from(date_is: DateIs) -> Self {
        Self::DateIs(date_is)
    }
}

/// One inert `calcext:condition` rule of a conditional format.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Condition {
    /// The decoded condition expression (`calcext:value`). It is never
    /// evaluated by litchi.
    pub condition: String,
    /// Name of the table-cell style referenced by the rule
    /// (`calcext:apply-style-name`). It is never resolved by litchi.
    pub apply_style_name: String,
    /// Optional lexical base cell address for relative formula references
    /// (`calcext:base-cell-address`).
    pub base_cell_address: Option<String>,
}

/// The threshold kind of a color-scale, data-bar, or icon-set entry.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EntryType {
    /// The lowest value of the range (`minimum`).
    Minimum,
    /// The highest value of the range (`maximum`).
    Maximum,
    /// A percent threshold (`percent`).
    Percent,
    /// A percentile threshold (`percentile`).
    Percentile,
    /// An inert formula threshold (`formula`). Never evaluated by litchi.
    Formula,
    /// A literal number threshold (`number`).
    Number,
    /// An automatic lower bound, used by data bars (`auto-minimum`).
    AutomaticMinimum,
    /// An automatic upper bound, used by data bars (`auto-maximum`).
    AutomaticMaximum,
}

impl EntryType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "minimum" => Ok(Self::Minimum),
            "maximum" => Ok(Self::Maximum),
            "percent" => Ok(Self::Percent),
            "percentile" => Ok(Self::Percentile),
            "formula" => Ok(Self::Formula),
            "number" => Ok(Self::Number),
            "auto-minimum" => Ok(Self::AutomaticMinimum),
            "auto-maximum" => Ok(Self::AutomaticMaximum),
            _ => Err(Error::InvalidFormat(format!(
                "invalid calcext entry type '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Minimum => "minimum",
            Self::Maximum => "maximum",
            Self::Percent => "percent",
            Self::Percentile => "percentile",
            Self::Formula => "formula",
            Self::Number => "number",
            Self::AutomaticMinimum => "auto-minimum",
            Self::AutomaticMaximum => "auto-maximum",
        }
    }

    /// Whether the entry value is an inert formula rather than a number.
    fn holds_formula(self) -> bool {
        self == Self::Formula
    }
}

/// One inert `calcext:color-scale-entry` threshold of a color scale.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ColorScaleEntry {
    /// The threshold kind (`calcext:type`).
    pub entry_type: EntryType,
    /// The lexical threshold (`calcext:value`): a number, or an inert formula
    /// when `entry_type` is [`EntryType::Formula`].
    pub value: String,
    /// The entry color in `#RRGGBB` form (`calcext:color`).
    pub color: String,
}

/// An inert `calcext:color-scale` gradient rule.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ColorScale {
    /// Threshold entries in document order.
    pub entries: Vec<ColorScaleEntry>,
}

/// One inert `calcext:formatting-entry` limit of a data bar.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataBarEntry {
    /// The limit kind (`calcext:type`).
    pub entry_type: EntryType,
    /// The lexical limit (`calcext:value`): a number, or an inert formula when
    /// `entry_type` is [`EntryType::Formula`].
    pub value: String,
}

/// The axis placement of a data bar (`calcext:axis-position`).
///
/// When the attribute is absent the axis is placed automatically.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DataBarAxisPosition {
    /// The axis is drawn in the middle of the cell (`middle`).
    Middle,
    /// No axis is drawn (`none`).
    None,
}

impl DataBarAxisPosition {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "middle" => Ok(Self::Middle),
            "none" => Ok(Self::None),
            _ => Err(Error::InvalidFormat(format!(
                "invalid calcext:axis-position value '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Middle => "middle",
            Self::None => "none",
        }
    }
}

/// An inert `calcext:data-bar` rule.
///
/// Numeric lengths are stored lexically and are never rendered by litchi.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataBar {
    /// Optional bar color for positive values in `#RRGGBB` form
    /// (`calcext:positive-color`).
    pub positive_color: Option<String>,
    /// Optional bar color for negative values in `#RRGGBB` form
    /// (`calcext:negative-color`).
    pub negative_color: Option<String>,
    /// Whether the bar is drawn as a gradient (`calcext:gradient`).
    pub gradient: Option<bool>,
    /// Optional axis placement (`calcext:axis-position`).
    pub axis_position: Option<DataBarAxisPosition>,
    /// Whether the cell value is shown next to the bar (`calcext:show-value`).
    pub show_value: Option<bool>,
    /// Optional axis color in `#RRGGBB` form (`calcext:axis-color`).
    pub axis_color: Option<String>,
    /// Optional minimum bar length as a lexical number (`calcext:min-length`).
    pub min_length: Option<String>,
    /// Optional maximum bar length as a lexical number (`calcext:max-length`).
    pub max_length: Option<String>,
    /// Exactly two limit entries: lower bound first, upper bound second.
    pub entries: Vec<DataBarEntry>,
}

/// One inert `calcext:formatting-entry` threshold of an icon set.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct IconSetEntry {
    /// The threshold kind (`calcext:type`).
    pub entry_type: EntryType,
    /// The lexical threshold (`calcext:value`): a number, or an inert formula
    /// when `entry_type` is [`EntryType::Formula`].
    pub value: String,
    /// Whether the threshold comparison includes equality
    /// (`calcext:greater-equal`, defaulting to true when absent).
    pub greater_equal: Option<bool>,
}

/// The icon family of an icon set (`calcext:icon-set-type`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum IconSetType {
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    ThreeSmilies,
    ThreeColorSmilies,
    ThreeStars,
    ThreeTriangles,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
    FiveBoxes,
}

impl IconSetType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "3Arrows" => Self::ThreeArrows,
            "3ArrowsGray" => Self::ThreeArrowsGray,
            "3Flags" => Self::ThreeFlags,
            "3TrafficLights1" => Self::ThreeTrafficLights1,
            "3TrafficLights2" => Self::ThreeTrafficLights2,
            "3Signs" => Self::ThreeSigns,
            "3Symbols" => Self::ThreeSymbols,
            "3Symbols2" => Self::ThreeSymbols2,
            "3Smilies" => Self::ThreeSmilies,
            "3ColorSmilies" => Self::ThreeColorSmilies,
            "3Stars" => Self::ThreeStars,
            "3Triangles" => Self::ThreeTriangles,
            "4Arrows" => Self::FourArrows,
            "4ArrowsGray" => Self::FourArrowsGray,
            "4RedToBlack" => Self::FourRedToBlack,
            "4Rating" => Self::FourRating,
            "4TrafficLights" => Self::FourTrafficLights,
            "5Arrows" => Self::FiveArrows,
            "5ArrowsGray" => Self::FiveArrowsGray,
            "5Rating" => Self::FiveRating,
            "5Quarters" => Self::FiveQuarters,
            "5Boxes" => Self::FiveBoxes,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid calcext:icon-set-type value '{value}'"
                )));
            },
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ThreeArrows => "3Arrows",
            Self::ThreeArrowsGray => "3ArrowsGray",
            Self::ThreeFlags => "3Flags",
            Self::ThreeTrafficLights1 => "3TrafficLights1",
            Self::ThreeTrafficLights2 => "3TrafficLights2",
            Self::ThreeSigns => "3Signs",
            Self::ThreeSymbols => "3Symbols",
            Self::ThreeSymbols2 => "3Symbols2",
            Self::ThreeSmilies => "3Smilies",
            Self::ThreeColorSmilies => "3ColorSmilies",
            Self::ThreeStars => "3Stars",
            Self::ThreeTriangles => "3Triangles",
            Self::FourArrows => "4Arrows",
            Self::FourArrowsGray => "4ArrowsGray",
            Self::FourRedToBlack => "4RedToBlack",
            Self::FourRating => "4Rating",
            Self::FourTrafficLights => "4TrafficLights",
            Self::FiveArrows => "5Arrows",
            Self::FiveArrowsGray => "5ArrowsGray",
            Self::FiveRating => "5Rating",
            Self::FiveQuarters => "5Quarters",
            Self::FiveBoxes => "5Boxes",
        }
    }
}

/// One inert `calcext:custom-iconset` icon replacement of an icon set.
///
/// LibreOffice writes these ahead of the threshold entries when an icon set
/// mixes icons from different families (`calcext:custom="true"`). Litchi
/// stores the assignment as typed data and never renders it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CustomIcon {
    /// The icon family the replacement icon comes from
    /// (`calcext:custom-iconset-name`).
    pub icon_set_type: IconSetType,
    /// The replaced icon position (`calcext:custom-iconset-index`).
    pub index: u32,
}

impl CustomIcon {
    /// Create an inert custom icon assignment.
    pub fn new(icon_set_type: IconSetType, index: u32) -> Self {
        Self {
            icon_set_type,
            index,
        }
    }
}

/// An inert `calcext:icon-set` rule.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct IconSet {
    /// The icon family (`calcext:icon-set-type`).
    pub icon_set_type: IconSetType,
    /// Whether the cell value is shown next to the icon (`calcext:show-value`).
    pub show_value: Option<bool>,
    /// Whether the icon set uses custom icons (`calcext:custom`).
    pub custom: Option<bool>,
    /// Inert custom icon assignments (`calcext:custom-iconset` children) in
    /// document order.
    pub custom_icons: Vec<CustomIcon>,
    /// Threshold entries in document order.
    pub entries: Vec<IconSetEntry>,
}

/// The date bucket of a `calcext:date-is` rule (`calcext:date`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DateType {
    Today,
    Yesterday,
    Tomorrow,
    Last7Days,
    ThisWeek,
    LastWeek,
    NextWeek,
    ThisMonth,
    LastMonth,
    NextMonth,
    ThisYear,
    LastYear,
    NextYear,
}

impl DateType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "today" => Self::Today,
            "yesterday" => Self::Yesterday,
            "tomorrow" => Self::Tomorrow,
            "last-7-days" => Self::Last7Days,
            "this-week" => Self::ThisWeek,
            "last-week" => Self::LastWeek,
            "next-week" => Self::NextWeek,
            "this-month" => Self::ThisMonth,
            "last-month" => Self::LastMonth,
            "next-month" => Self::NextMonth,
            "this-year" => Self::ThisYear,
            "last-year" => Self::LastYear,
            "next-year" => Self::NextYear,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "invalid calcext:date value '{value}'"
                )));
            },
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Today => "today",
            Self::Yesterday => "yesterday",
            Self::Tomorrow => "tomorrow",
            Self::Last7Days => "last-7-days",
            Self::ThisWeek => "this-week",
            Self::LastWeek => "last-week",
            Self::NextWeek => "next-week",
            Self::ThisMonth => "this-month",
            Self::LastMonth => "last-month",
            Self::NextMonth => "next-month",
            Self::ThisYear => "this-year",
            Self::LastYear => "last-year",
            Self::NextYear => "next-year",
        }
    }
}

/// An inert `calcext:date-is` date rule.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DateIs {
    /// The date bucket (`calcext:date`).
    pub date: DateType,
    /// Name of the table-cell style applied on a match (`calcext:style`). It
    /// is never resolved by litchi.
    pub style: String,
}

impl Format {
    /// Create a validated inert conditional format.
    pub fn new(target_range_addresses: Vec<String>, rules: Vec<Rule>) -> Result<Self> {
        let format = Self {
            target_range_addresses,
            rules,
        };
        validate_conditional_format(&format)?;
        Ok(format)
    }

    /// Iterate over the `calcext:condition` expression rules, in document order.
    pub fn conditions(&self) -> impl Iterator<Item = &Condition> {
        self.rules.iter().filter_map(|rule| match rule {
            Rule::Condition(condition) => Some(condition),
            _ => None,
        })
    }
}

impl Condition {
    /// Create an inert condition rule without a base cell address.
    pub fn new(condition: impl Into<String>, apply_style_name: impl Into<String>) -> Self {
        Self {
            condition: condition.into(),
            apply_style_name: apply_style_name.into(),
            base_cell_address: None,
        }
    }

    /// Set the optional lexical base cell address.
    pub fn with_base_cell_address(mut self, address: impl Into<String>) -> Self {
        self.base_cell_address = Some(address.into());
        self
    }
}

impl ColorScaleEntry {
    /// Create an inert color-scale threshold entry.
    pub fn new(entry_type: EntryType, value: impl Into<String>, color: impl Into<String>) -> Self {
        Self {
            entry_type,
            value: value.into(),
            color: color.into(),
        }
    }
}

impl ColorScale {
    /// Create an inert color-scale rule.
    pub fn new(entries: Vec<ColorScaleEntry>) -> Self {
        Self { entries }
    }
}

impl DataBarEntry {
    /// Create an inert data-bar limit entry.
    pub fn new(entry_type: EntryType, value: impl Into<String>) -> Self {
        Self {
            entry_type,
            value: value.into(),
        }
    }
}

impl DataBar {
    /// Create an inert data-bar rule with only limit entries set.
    pub fn new(entries: Vec<DataBarEntry>) -> Self {
        Self {
            positive_color: None,
            negative_color: None,
            gradient: None,
            axis_position: None,
            show_value: None,
            axis_color: None,
            min_length: None,
            max_length: None,
            entries,
        }
    }

    /// Set the optional bar colors for positive and negative values.
    pub fn with_colors(
        mut self,
        positive_color: impl Into<String>,
        negative_color: Option<String>,
    ) -> Self {
        self.positive_color = Some(positive_color.into());
        self.negative_color = negative_color;
        self
    }

    /// Set whether the bar is drawn as a gradient.
    pub fn with_gradient(mut self, gradient: bool) -> Self {
        self.gradient = Some(gradient);
        self
    }

    /// Set the optional axis placement and color.
    pub fn with_axis(mut self, position: DataBarAxisPosition, color: Option<String>) -> Self {
        self.axis_position = Some(position);
        self.axis_color = color;
        self
    }

    /// Set whether the cell value is shown next to the bar.
    pub fn with_show_value(mut self, show_value: bool) -> Self {
        self.show_value = Some(show_value);
        self
    }

    /// Set the optional minimum and maximum bar lengths as lexical numbers.
    pub fn with_lengths(mut self, min_length: Option<String>, max_length: Option<String>) -> Self {
        self.min_length = min_length;
        self.max_length = max_length;
        self
    }
}

impl IconSetEntry {
    /// Create an inert icon-set threshold entry.
    pub fn new(entry_type: EntryType, value: impl Into<String>) -> Self {
        Self {
            entry_type,
            value: value.into(),
            greater_equal: None,
        }
    }

    /// Set whether the threshold comparison includes equality.
    pub fn with_greater_equal(mut self, greater_equal: bool) -> Self {
        self.greater_equal = Some(greater_equal);
        self
    }
}

impl IconSet {
    /// Create an inert icon-set rule.
    pub fn new(icon_set_type: IconSetType, entries: Vec<IconSetEntry>) -> Self {
        Self {
            icon_set_type,
            show_value: None,
            custom: None,
            custom_icons: Vec::new(),
            entries,
        }
    }

    /// Set whether the cell value is shown next to the icon.
    pub fn with_show_value(mut self, show_value: bool) -> Self {
        self.show_value = Some(show_value);
        self
    }

    /// Set whether the icon set uses custom icons.
    pub fn with_custom(mut self, custom: bool) -> Self {
        self.custom = Some(custom);
        self
    }

    /// Set the inert custom icon assignments.
    pub fn with_custom_icons(mut self, custom_icons: Vec<CustomIcon>) -> Self {
        self.custom_icons = custom_icons;
        self
    }
}

impl DateIs {
    /// Create an inert date rule referencing a table-cell style.
    pub fn new(date: DateType, style: impl Into<String>) -> Self {
        Self {
            date,
            style: style.into(),
        }
    }
}

pub(crate) fn validate_conditional_format(format: &Format) -> Result<()> {
    if format.target_range_addresses.is_empty() {
        return Err(Error::InvalidFormat(
            "conditional formats require at least one target range".to_string(),
        ));
    }
    validate_cell_range_addresses(&format.target_range_addresses)?;
    for range in &format.target_range_addresses {
        validate_attribute_length("calcext:target-range-address", range)?;
    }
    if format.rules.is_empty() {
        return Err(Error::InvalidFormat(
            "conditional formats require at least one rule".to_string(),
        ));
    }
    if format.rules.len() > MAX_RULES_PER_FORMAT {
        return Err(Error::InvalidFormat(format!(
            "conditional format exceeds the {MAX_RULES_PER_FORMAT} rule safety limit"
        )));
    }
    for rule in &format.rules {
        validate_rule(rule)?;
    }
    Ok(())
}

pub(crate) fn validate_conditional_formats(formats: &[Format]) -> Result<()> {
    if formats.len() > MAX_CONDITIONAL_FORMATS_PER_SHEET {
        return Err(Error::InvalidFormat(format!(
            "sheet exceeds the {MAX_CONDITIONAL_FORMATS_PER_SHEET} conditional format safety limit"
        )));
    }
    for format in formats {
        validate_conditional_format(format)?;
    }
    Ok(())
}

pub(crate) fn validate_rule(rule: &Rule) -> Result<()> {
    match rule {
        Rule::Condition(condition) => validate_condition(condition),
        Rule::ColorScale(color_scale) => validate_color_scale(color_scale),
        Rule::DataBar(data_bar) => validate_data_bar(data_bar),
        Rule::IconSet(icon_set) => validate_icon_set(icon_set),
        Rule::DateIs(date_is) => validate_date_is(date_is),
    }
}

pub(crate) fn validate_condition(condition: &Condition) -> Result<()> {
    if condition.condition.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:condition requires a non-empty calcext:value".to_string(),
        ));
    }
    validate_attribute_length("calcext:value", &condition.condition)?;
    if condition.apply_style_name.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:condition requires a non-empty calcext:apply-style-name".to_string(),
        ));
    }
    validate_attribute_length("calcext:apply-style-name", &condition.apply_style_name)?;
    if let Some(address) = &condition.base_cell_address {
        if address.trim() != address || address.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "invalid calcext:base-cell-address '{address}'"
            )));
        }
        validate_attribute_length("calcext:base-cell-address", address)?;
    }
    Ok(())
}

pub(crate) fn validate_color_scale(color_scale: &ColorScale) -> Result<()> {
    if color_scale.entries.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:color-scale requires at least one calcext:color-scale-entry".to_string(),
        ));
    }
    validate_entry_count(color_scale.entries.len(), "calcext:color-scale")?;
    for entry in &color_scale.entries {
        validate_color_scale_entry(entry)?;
    }
    Ok(())
}

pub(crate) fn validate_color_scale_entry(entry: &ColorScaleEntry) -> Result<()> {
    validate_entry_value(entry.entry_type, &entry.value)?;
    validate_color("calcext:color", &entry.color)
}

pub(crate) fn validate_data_bar(data_bar: &DataBar) -> Result<()> {
    if data_bar.entries.len() != DATA_BAR_ENTRY_COUNT {
        return Err(Error::InvalidFormat(format!(
            "calcext:data-bar requires exactly {DATA_BAR_ENTRY_COUNT} calcext:formatting-entry elements"
        )));
    }
    for entry in &data_bar.entries {
        validate_data_bar_entry(entry)?;
    }
    validate_data_bar_attributes(data_bar)
}

/// Validate only the element attributes of a data bar, not its entries.
pub(crate) fn validate_data_bar_attributes(data_bar: &DataBar) -> Result<()> {
    if let Some(color) = &data_bar.positive_color {
        validate_color("calcext:positive-color", color)?;
    }
    if let Some(color) = &data_bar.negative_color {
        validate_color("calcext:negative-color", color)?;
    }
    if let Some(color) = &data_bar.axis_color {
        validate_color("calcext:axis-color", color)?;
    }
    if let Some(length) = &data_bar.min_length {
        validate_lexical_number("calcext:min-length", length)?;
    }
    if let Some(length) = &data_bar.max_length {
        validate_lexical_number("calcext:max-length", length)?;
    }
    Ok(())
}

pub(crate) fn validate_data_bar_entry(entry: &DataBarEntry) -> Result<()> {
    validate_entry_value(entry.entry_type, &entry.value)
}

pub(crate) fn validate_icon_set(icon_set: &IconSet) -> Result<()> {
    if icon_set.entries.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:icon-set requires at least one calcext:formatting-entry".to_string(),
        ));
    }
    validate_entry_count(icon_set.entries.len(), "calcext:icon-set")?;
    if icon_set.custom_icons.len() > MAX_ENTRIES_PER_RULE {
        return Err(Error::InvalidFormat(format!(
            "calcext:icon-set exceeds the {MAX_ENTRIES_PER_RULE} custom icon safety limit"
        )));
    }
    for entry in &icon_set.entries {
        validate_icon_set_entry(entry)?;
    }
    Ok(())
}

pub(crate) fn validate_icon_set_entry(entry: &IconSetEntry) -> Result<()> {
    validate_entry_value(entry.entry_type, &entry.value)
}

pub(crate) fn validate_date_is(date_is: &DateIs) -> Result<()> {
    if date_is.style.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:date-is requires a non-empty calcext:style".to_string(),
        ));
    }
    validate_attribute_length("calcext:style", &date_is.style)?;
    Ok(())
}

fn validate_entry_count(count: usize, element: &str) -> Result<()> {
    if count > MAX_ENTRIES_PER_RULE {
        return Err(Error::InvalidFormat(format!(
            "{element} exceeds the {MAX_ENTRIES_PER_RULE} entry safety limit"
        )));
    }
    Ok(())
}

fn validate_entry_value(entry_type: EntryType, value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext formatting entries require a non-empty calcext:value".to_string(),
        ));
    }
    validate_attribute_length("calcext:value", value)?;
    if !entry_type.holds_formula() {
        validate_lexical_number("calcext:value", value)?;
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
    if value.len() > MAX_CONDITIONAL_ATTRIBUTE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds the {MAX_CONDITIONAL_ATTRIBUTE_BYTES} byte safety limit"
        )));
    }
    Ok(())
}

/// Write a sheet's `calcext:conditional-formats` container after its rows.
pub(crate) fn write_conditional_formats(out: &mut String, formats: &[Format]) -> Result<()> {
    validate_conditional_formats(formats)?;
    if formats.is_empty() {
        return Ok(());
    }
    out.push_str("<calcext:conditional-formats>");
    for format in formats {
        out.push_str("<calcext:conditional-format calcext:target-range-address=\"");
        out.push_str(&escape_xml(&format.target_range_addresses.join(" ")));
        out.push_str("\">");
        for rule in &format.rules {
            write_rule(out, rule);
        }
        out.push_str("</calcext:conditional-format>");
    }
    out.push_str("</calcext:conditional-formats>");
    Ok(())
}

fn write_rule(out: &mut String, rule: &Rule) {
    match rule {
        Rule::Condition(condition) => write_condition(out, condition),
        Rule::ColorScale(color_scale) => write_color_scale(out, color_scale),
        Rule::DataBar(data_bar) => write_data_bar(out, data_bar),
        Rule::IconSet(icon_set) => write_icon_set(out, icon_set),
        Rule::DateIs(date_is) => write_date_is(out, date_is),
    }
}

fn write_condition(out: &mut String, condition: &Condition) {
    out.push_str("<calcext:condition calcext:apply-style-name=\"");
    out.push_str(&escape_xml(&condition.apply_style_name));
    out.push_str("\" calcext:value=\"");
    out.push_str(&escape_xml(&condition.condition));
    out.push('"');
    if let Some(address) = &condition.base_cell_address {
        out.push_str(" calcext:base-cell-address=\"");
        out.push_str(&escape_xml(address));
        out.push('"');
    }
    out.push_str("/>");
}

fn write_color_scale(out: &mut String, color_scale: &ColorScale) {
    out.push_str("<calcext:color-scale>");
    for entry in &color_scale.entries {
        out.push_str("<calcext:color-scale-entry calcext:value=\"");
        out.push_str(&escape_xml(&entry.value));
        out.push_str("\" calcext:type=\"");
        out.push_str(entry.entry_type.as_str());
        out.push_str("\" calcext:color=\"");
        out.push_str(&entry.color);
        out.push_str("\"/>");
    }
    out.push_str("</calcext:color-scale>");
}

fn write_data_bar(out: &mut String, data_bar: &DataBar) {
    out.push_str("<calcext:data-bar");
    write_optional_bool_attribute(out, "calcext:gradient", data_bar.gradient);
    write_optional_bool_attribute(out, "calcext:show-value", data_bar.show_value);
    if let Some(length) = &data_bar.min_length {
        write_attribute(out, "calcext:min-length", length);
    }
    if let Some(length) = &data_bar.max_length {
        write_attribute(out, "calcext:max-length", length);
    }
    if let Some(color) = &data_bar.negative_color {
        write_attribute(out, "calcext:negative-color", color);
    }
    if let Some(position) = data_bar.axis_position {
        write_attribute(out, "calcext:axis-position", position.as_str());
    }
    if let Some(color) = &data_bar.positive_color {
        write_attribute(out, "calcext:positive-color", color);
    }
    if let Some(color) = &data_bar.axis_color {
        write_attribute(out, "calcext:axis-color", color);
    }
    out.push('>');
    for entry in &data_bar.entries {
        out.push_str("<calcext:formatting-entry calcext:value=\"");
        out.push_str(&escape_xml(&entry.value));
        out.push_str("\" calcext:type=\"");
        out.push_str(entry.entry_type.as_str());
        out.push_str("\"/>");
    }
    out.push_str("</calcext:data-bar>");
}

fn write_icon_set(out: &mut String, icon_set: &IconSet) {
    out.push_str("<calcext:icon-set calcext:icon-set-type=\"");
    out.push_str(icon_set.icon_set_type.as_str());
    out.push('"');
    write_optional_bool_attribute(out, "calcext:custom", icon_set.custom);
    write_optional_bool_attribute(out, "calcext:show-value", icon_set.show_value);
    out.push('>');
    for custom_icon in &icon_set.custom_icons {
        out.push_str("<calcext:custom-iconset calcext:custom-iconset-name=\"");
        out.push_str(custom_icon.icon_set_type.as_str());
        out.push_str("\" calcext:custom-iconset-index=\"");
        out.push_str(&custom_icon.index.to_string());
        out.push_str("\"/>");
    }
    for entry in &icon_set.entries {
        out.push_str("<calcext:formatting-entry calcext:value=\"");
        out.push_str(&escape_xml(&entry.value));
        out.push('"');
        if let Some(greater_equal) = entry.greater_equal {
            out.push_str(" calcext:greater-equal=\"");
            out.push_str(bool_str(greater_equal));
            out.push('"');
        }
        out.push_str(" calcext:type=\"");
        out.push_str(entry.entry_type.as_str());
        out.push_str("\"/>");
    }
    out.push_str("</calcext:icon-set>");
}

fn write_date_is(out: &mut String, date_is: &DateIs) {
    out.push_str("<calcext:date-is calcext:style=\"");
    out.push_str(&escape_xml(&date_is.style));
    out.push_str("\" calcext:date=\"");
    out.push_str(date_is.date.as_str());
    out.push_str("\"/>");
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
        out.push_str("=\"");
        out.push_str(bool_str(value));
        out.push('"');
    }
}

fn bool_str(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_condition() -> Condition {
        Condition::new("cell-content()>5", "Good").with_base_cell_address("Sheet1.A1")
    }

    fn sample_color_scale() -> ColorScale {
        ColorScale::new(vec![
            ColorScaleEntry::new(EntryType::Minimum, "0", "#ff0000"),
            ColorScaleEntry::new(EntryType::Percentile, "50", "#ffff00"),
            ColorScaleEntry::new(EntryType::Maximum, "0", "#00ff00"),
        ])
    }

    fn sample_data_bar() -> DataBar {
        DataBar::new(vec![
            DataBarEntry::new(EntryType::AutomaticMinimum, "0"),
            DataBarEntry::new(EntryType::AutomaticMaximum, "0"),
        ])
        .with_colors("#638ec6", Some("#ff0000".to_string()))
        .with_gradient(false)
        .with_axis(DataBarAxisPosition::Middle, Some("#000000".to_string()))
        .with_show_value(false)
        .with_lengths(Some("10".to_string()), Some("90".to_string()))
    }

    fn sample_icon_set() -> IconSet {
        IconSet::new(
            IconSetType::ThreeTrafficLights1,
            vec![
                IconSetEntry::new(EntryType::Percent, "0"),
                IconSetEntry::new(EntryType::Percent, "33").with_greater_equal(false),
                IconSetEntry::new(EntryType::Percent, "67"),
            ],
        )
        .with_show_value(false)
        .with_custom(true)
        .with_custom_icons(vec![CustomIcon::new(IconSetType::ThreeStars, 0)])
    }

    fn sample_date_is() -> DateIs {
        DateIs::new(DateType::Last7Days, "Recent")
    }

    fn sample_format() -> Format {
        Format::new(
            vec!["Sheet1.A1:Sheet1.A5".to_string()],
            vec![
                sample_condition().into(),
                sample_color_scale().into(),
                sample_data_bar().into(),
                sample_icon_set().into(),
                sample_date_is().into(),
            ],
        )
        .unwrap()
    }

    #[test]
    fn writes_all_rule_types_with_escaped_values() {
        let mut xml = String::new();
        write_conditional_formats(&mut xml, &[sample_format()]).unwrap();
        assert!(xml.starts_with("<calcext:conditional-formats>"));
        assert!(xml.contains(
            r#"<calcext:condition calcext:apply-style-name="Good" calcext:value="cell-content()&gt;5" calcext:base-cell-address="Sheet1.A1"/>"#
        ));
        assert!(xml.contains(
            r##"<calcext:color-scale-entry calcext:value="50" calcext:type="percentile" calcext:color="#ffff00"/>"##
        ));
        assert!(xml.contains(
            r##"<calcext:data-bar calcext:gradient="false" calcext:show-value="false" calcext:min-length="10" calcext:max-length="90" calcext:negative-color="#ff0000" calcext:axis-position="middle" calcext:positive-color="#638ec6" calcext:axis-color="#000000">"##
        ));
        assert!(xml.contains(
            r#"<calcext:formatting-entry calcext:value="0" calcext:type="auto-minimum"/>"#
        ));
        assert!(xml.contains(
            r#"<calcext:icon-set calcext:icon-set-type="3TrafficLights1" calcext:custom="true" calcext:show-value="false">"#
        ));
        assert!(xml.contains(
            r#"<calcext:custom-iconset calcext:custom-iconset-name="3Stars" calcext:custom-iconset-index="0"/>"#
        ));
        assert!(xml.contains(
            r#"<calcext:formatting-entry calcext:value="33" calcext:greater-equal="false" calcext:type="percent"/>"#
        ));
        assert!(
            xml.contains(r#"<calcext:date-is calcext:style="Recent" calcext:date="last-7-days"/>"#)
        );
        assert!(xml.ends_with("</calcext:conditional-formats>"));
    }

    #[test]
    fn writes_nothing_for_an_empty_collection() {
        let mut xml = String::new();
        write_conditional_formats(&mut xml, &[]).unwrap();
        assert!(xml.is_empty());
    }

    #[test]
    fn rejects_missing_ranges_rules_and_blank_values() {
        assert!(Format::new(Vec::new(), vec![sample_condition().into()]).is_err());
        assert!(Format::new(vec![".A1".to_string()], Vec::new()).is_err());
        assert!(
            Format::new(
                vec![".A1".to_string()],
                vec![Condition::new("", "S").into()],
            )
            .is_err()
        );
        assert!(
            Format::new(
                vec![".A1".to_string()],
                vec![Condition::new("x", "").into()],
            )
            .is_err()
        );
        assert!(
            Format::new(
                vec![".A1".to_string()],
                vec![
                    Condition::new("x", "S")
                        .with_base_cell_address(" A1 ")
                        .into()
                ],
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_invalid_rule_bodies() {
        let format = |rule: Rule| Format::new(vec![".A1".to_string()], vec![rule]);
        // Color scale without entries, with a bad color, or a non-numeric value.
        assert!(format(ColorScale::new(Vec::new()).into()).is_err());
        assert!(
            format(
                ColorScale::new(vec![ColorScaleEntry::new(EntryType::Minimum, "0", "red",)]).into(),
            )
            .is_err()
        );
        assert!(
            format(
                ColorScale::new(vec![ColorScaleEntry::new(
                    EntryType::Percent,
                    "soon",
                    "#ff0000",
                )])
                .into(),
            )
            .is_err()
        );
        // Data bars require exactly two limit entries and numeric lengths.
        assert!(
            format(DataBar::new(vec![DataBarEntry::new(EntryType::AutomaticMinimum, "0",)]).into())
                .is_err()
        );
        assert!(
            format(
                DataBar::new(vec![
                    DataBarEntry::new(EntryType::AutomaticMinimum, "0"),
                    DataBarEntry::new(EntryType::AutomaticMaximum, "0"),
                ])
                .with_lengths(Some("wide".to_string()), None)
                .into(),
            )
            .is_err()
        );
        // Icon sets require entries; date rules require a style.
        assert!(format(IconSet::new(IconSetType::FiveBoxes, Vec::new()).into()).is_err());
        assert!(format(DateIs::new(DateType::Today, "").into()).is_err());
    }

    #[test]
    fn formula_entries_keep_inert_formula_values() {
        let scale = ColorScale::new(vec![
            ColorScaleEntry::new(EntryType::Minimum, "0", "#ff0000"),
            ColorScaleEntry::new(EntryType::Formula, "MAX([.A1:.A9])/2", "#00ff00"),
        ]);
        Format::new(vec![".A1".to_string()], vec![scale.into()]).unwrap();
    }
}
