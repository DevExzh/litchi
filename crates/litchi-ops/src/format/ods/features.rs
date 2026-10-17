//! ODS worksheet features as unified nodes (record 0327): hyperlinks,
//! pictures and shapes, conditional formats (`cf`), content validations,
//! the auto-filter and manual page breaks, with the XLSX property vocabulary
//! of record 0320. Cell annotations (`comment`) keep their own node.
//!
//! Every feature is a child of its worksheet, addressed by its position
//! among the worksheet's features of its kind (`/Sheet1/cf[1]`; the
//! auto-filter is `/Sheet1/autofilter`). A node holds one [`Feature`]
//! behind one `Arc`. Hyperlinks are the `text:a` anchors of the sheet's
//! cells in row-major order; every other kind is read from `content.xml`
//! through `litchi_ods::advanced::sheet_features`, once per sheet and
//! snapshot, without decoding a cell. The write verbs live in
//! `feature_write`.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::Position;
use litchi_ods::advanced::sheet_features::{
    self as sf, AutoFilter, DrawingKind, FormatStyle, Validation, from_odff,
};
use litchi_ods::model::conditional_format::{self as cf, EntryType};

use super::Ods;
use super::node::{Item, Node};
use super::spec::Kind;
use crate::error::{Code, Error, Result};
use crate::format::workbook::Label;
use crate::node::spec::{Access as Flags, GapSpec, PropSpec, Subject};
use crate::node::{PathStep, Preview, Value};
use crate::path::navigate::{Lookup, Select};
use crate::registry::{Access, Alias, Constraint, Entry, Table};
use crate::value::{Color, ColorContext, Length, LengthSpec, ValueType};

/// The feature kinds of a worksheet this module lists, in listing order.
pub(super) const KINDS: [Kind; 8] = [
    Kind::Hyperlink,
    Kind::Picture,
    Kind::Shape,
    Kind::Cf,
    Kind::DataValidation,
    Kind::Autofilter,
    Kind::RowBreak,
    Kind::ColBreak,
];

/// Whether `kind` is a worksheet feature this module lists.
pub(super) fn is_feature(kind: Kind) -> bool {
    KINDS.contains(&kind)
}

/// One worksheet feature: its kind, its position among the worksheet's
/// features of that kind, and the leaf value it was listed from.
#[derive(Clone)]
pub(crate) struct Feature {
    pub(super) kind: Kind,
    pub(super) index: u32,
    pub(super) data: Data,
}

/// The leaf value of a feature.
#[derive(Clone)]
pub(crate) enum Data {
    Hyperlink(Box<Link>),
    Drawing(Box<sf::Drawing>),
    Cf(Box<CfFeature>),
    Validation(Box<Validation>),
    Autofilter(AutoFilter),
    Break(u32),
}

/// A cell hyperlink.
#[derive(Clone)]
pub(crate) struct Link {
    pub(super) row: u32,
    pub(super) column: u32,
    /// Its position among the cell's links.
    pub(super) slot: usize,
    pub(super) href: String,
    pub(super) text: String,
    pub(super) title: Option<String>,
}

/// A conditional format with the cell style its first rule applies.
#[derive(Clone)]
pub(crate) struct CfFeature {
    pub(super) format: cf::Format,
    pub(super) rule: Option<Rule>,
    pub(super) style: FormatStyle,
}

/// The A1 label of a zero-based `(first_row, first_column, last_row,
/// last_column)` rectangle.
pub(super) fn rect_label((first_row, first_column, last_row, last_column): sf::Rect) -> String {
    Label::range((first_column, first_row), (last_column, last_row))
        .as_str()
        .to_owned()
}

/// The space-separated A1 ranges of ODF range addresses.
pub(super) fn address_labels(addresses: &[String]) -> String {
    addresses
        .iter()
        .filter_map(|address| sf::parse_address(address))
        .map(|(_, rect)| rect_label(rect))
        .collect::<Vec<_>>()
        .join(" ")
}

impl Link {
    pub(super) fn reference(&self) -> String {
        Label::cell(self.column, self.row).as_str().to_owned()
    }

    /// The external target; `None` for an in-document location.
    pub(super) fn url(&self) -> Option<&str> {
        (!self.href.starts_with('#')).then_some(self.href.as_str())
    }

    /// The in-document location in A1 syntax (`Sheet2!A1`).
    pub(super) fn location(&self) -> Option<String> {
        let target = self.href.strip_prefix('#')?;
        Some(match target.rfind('.') {
            Some(dot) if dot > 0 => format!(
                "{}!{}",
                &target[..dot].trim_start_matches('$'),
                target[dot + 1..].replace('$', "")
            ),
            _ => target.to_owned(),
        })
    }
}

impl Feature {
    /// The node's last path step.
    pub(super) fn step(&self) -> PathStep<'static> {
        match self.kind {
            Kind::Autofilter => PathStep::new("autofilter"),
            kind => PathStep::at(step_name(kind), Position::new(self.index as usize)),
        }
    }

    /// The text a query matches and the node view shows.
    pub(super) fn text(&self) -> Option<Cow<'_, str>> {
        Some(match &self.data {
            Data::Hyperlink(link) => Cow::Borrowed(if link.text.is_empty() {
                link.href.as_str()
            } else {
                link.text.as_str()
            }),
            Data::Drawing(drawing) => match drawing.kind {
                DrawingKind::Shape => Cow::Borrowed(drawing.text.as_deref()?),
                DrawingKind::Picture => Cow::Borrowed(drawing.description.as_deref()?),
            },
            Data::Cf(format) => Cow::Owned(address_labels(&format.format.target_range_addresses)),
            Data::Validation(validation) => Cow::Owned(ranges(&validation.ranges)),
            Data::Autofilter(filter) => Cow::Owned(rect_label(filter.range)),
            Data::Break(_) => return None,
        })
    }

    /// The node view's one-line preview.
    pub(super) fn preview(&self) -> Option<Preview<'_>> {
        Some(Preview::Text(match &self.data {
            Data::Hyperlink(link) => Cow::Owned(link.reference()),
            Data::Drawing(drawing) => Cow::Borrowed(drawing.name.as_deref()?),
            Data::Cf(format) => Cow::Owned(format!(
                "{} {}",
                format.rule.as_ref().map_or("", |rule| rule.kind.name()),
                address_labels(&format.format.target_range_addresses)
            )),
            Data::Validation(validation) => Cow::Owned(ranges(&validation.ranges)),
            Data::Autofilter(filter) => Cow::Owned(rect_label(filter.range)),
            Data::Break(index) => Cow::Owned(index.to_string()),
        }))
    }

    /// The property table of the feature's kind.
    pub(super) fn table(&self) -> &'static [PropSpec] {
        props(self.kind)
    }

    /// The value at `index` of the feature's property table.
    pub(super) fn value(&self, index: usize) -> Option<Value<'_>> {
        match &self.data {
            Data::Hyperlink(link) => hyperlink_value(link, *HYPERLINK_KEYS.get(index)?),
            Data::Drawing(drawing) => match self.kind {
                Kind::Picture => picture_value(drawing, *PICTURE_KEYS.get(index)?),
                _ => shape_value(drawing, *SHAPE_KEYS.get(index)?),
            },
            Data::Cf(format) => cf_value(format, *CF_KEYS.get(index)?),
            Data::Validation(validation) => {
                validation_value(validation, *VALIDATION_KEYS.get(index)?)
            },
            Data::Autofilter(filter) => match *AUTOFILTER_KEYS.get(index)? {
                AutofilterKey::Range => Some(owned(rect_label(filter.range))),
                AutofilterKey::Criteria => None,
            },
            Data::Break(item) => match *BREAK_KEYS.get(index)? {
                BreakKey::Index => Some(int(*item)),
                BreakKey::Manual => Some(Value::Bool(true)),
                BreakKey::Min | BreakKey::Max => None,
            },
        }
    }
}

/// The space-separated A1 labels of rectangles.
pub(super) fn ranges(rects: &[sf::Rect]) -> String {
    rects
        .iter()
        .map(|rect| rect_label(*rect))
        .collect::<Vec<_>>()
        .join(" ")
}

/// The path step name of a feature kind.
pub(super) const fn step_name(kind: Kind) -> &'static str {
    match kind {
        Kind::Hyperlink => "hyperlink",
        Kind::Picture => "picture",
        Kind::Shape => "shape",
        Kind::Cf => "cf",
        Kind::DataValidation => "dataValidation",
        Kind::Autofilter => "autofilter",
        Kind::RowBreak => "rowbreak",
        _ => "colbreak",
    }
}

/// The property table of a feature kind.
pub(super) fn props(kind: Kind) -> &'static [PropSpec] {
    match kind {
        Kind::Hyperlink => HYPERLINK_PROPS,
        Kind::Picture => PICTURE_PROPS,
        Kind::Shape => SHAPE_PROPS,
        Kind::Cf => CF_PROPS,
        Kind::DataValidation => VALIDATION_PROPS,
        Kind::Autofilter => AUTOFILTER_PROPS,
        Kind::RowBreak => ROW_BREAK_PROPS,
        _ => COL_BREAK_PROPS,
    }
}

fn text(value: &str) -> Value<'_> {
    Value::Text(Cow::Borrowed(value))
}

fn owned(value: String) -> Value<'static> {
    Value::Text(Cow::Owned(value))
}

fn int(value: impl Into<i64>) -> Value<'static> {
    Value::Int(value.into())
}

/// A `#RRGGBB` color value.
fn color(value: &str) -> Option<Value<'static>> {
    Color::parse(value.trim(), ColorContext::Spreadsheet)
        .ok()
        .map(Value::Color)
}

/// A length in centimetres as OfficeCLI's EMU spelling.
fn emu(cm: f64) -> Value<'static> {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "drawing lengths are bounded"
    )]
    let emu = (cm * 360_000.0).round() as i64;
    owned(format!("{emu}emu"))
}

// ---------------------------------------------------------------------------
// Listing

/// The features of `kind` on the worksheet at `sheet`, in position order.
pub(super) fn list(file: &Ods, sheet: u32, kind: Kind) -> Result<Vec<Feature>> {
    let data: Vec<Data> = match kind {
        Kind::Hyperlink => links(file, sheet)?
            .into_iter()
            .map(|link| Data::Hyperlink(Box::new(link)))
            .collect(),
        Kind::Picture | Kind::Shape => {
            let wanted = if kind == Kind::Picture {
                DrawingKind::Picture
            } else {
                DrawingKind::Shape
            };
            sheet_features(file, sheet)?
                .drawings
                .iter()
                .filter(|drawing| drawing.kind == wanted)
                .map(|drawing| Data::Drawing(Box::new(drawing.clone())))
                .collect()
        },
        Kind::Cf => {
            let features = sheet_features(file, sheet)?;
            if features.formats.is_empty() {
                Vec::new()
            } else {
                let styles = file.cache().format_styles(|| {
                    sf::format_styles(file.materialized()?.as_bytes()).map_err(super::core_error)
                })?;
                features
                    .formats
                    .iter()
                    .map(|format| {
                        let rule = Rule::of(format);
                        let style = style_name(format)
                            .and_then(|name| styles.get(name))
                            .cloned()
                            .unwrap_or_default();
                        Data::Cf(Box::new(CfFeature {
                            format: format.clone(),
                            rule,
                            style,
                        }))
                    })
                    .collect()
            }
        },
        Kind::DataValidation => sheet_features(file, sheet)?
            .validations
            .iter()
            .map(|validation| Data::Validation(Box::new(validation.clone())))
            .collect(),
        Kind::Autofilter => sheet_features(file, sheet)?
            .auto_filter
            .clone()
            .map(Data::Autofilter)
            .into_iter()
            .collect(),
        Kind::RowBreak => sheet_features(file, sheet)?
            .row_breaks
            .iter()
            .map(|index| Data::Break(*index))
            .collect(),
        Kind::ColBreak => sheet_features(file, sheet)?
            .column_breaks
            .iter()
            .map(|index| Data::Break(*index))
            .collect(),
        _ => Vec::new(),
    };
    Ok(data
        .into_iter()
        .enumerate()
        .map(|(index, data)| Feature {
            kind,
            index: u32::try_from(index).unwrap_or(u32::MAX),
            data,
        })
        .collect())
}

/// The number of features of `kind` on a worksheet.
pub(super) fn count(file: &Ods, sheet: u32, kind: Kind) -> Result<usize> {
    Ok(match kind {
        Kind::Hyperlink => links(file, sheet)?.len(),
        Kind::Picture | Kind::Shape => {
            let wanted = if kind == Kind::Picture {
                DrawingKind::Picture
            } else {
                DrawingKind::Shape
            };
            sheet_features(file, sheet)?
                .drawings
                .iter()
                .filter(|drawing| drawing.kind == wanted)
                .count()
        },
        Kind::Cf => sheet_features(file, sheet)?.formats.len(),
        Kind::DataValidation => sheet_features(file, sheet)?.validations.len(),
        Kind::Autofilter => usize::from(sheet_features(file, sheet)?.auto_filter.is_some()),
        Kind::RowBreak => sheet_features(file, sheet)?.row_breaks.len(),
        Kind::ColBreak => sheet_features(file, sheet)?.column_breaks.len(),
        _ => 0,
    })
}

/// The cached `content.xml` features of the worksheet at `sheet`.
fn sheet_features(file: &Ods, sheet: u32) -> Result<&sf::Features> {
    let sheets = file.worksheets()?;
    let name = &sheets
        .get(sheet as usize)
        .ok_or_else(super::node::missing_sheet)?
        .name;
    file.cache().features(sheets.len(), sheet as usize, || {
        let xml = file.content_xml()?;
        sf::read(xml, name).map_err(super::core_error)
    })
}

/// The hyperlinks of a worksheet's cells, row-major.
pub(super) fn links(file: &Ods, sheet: u32) -> Result<Vec<Link>> {
    let data = file
        .worksheets()?
        .get(sheet as usize)
        .ok_or_else(super::node::missing_sheet)?;
    let index = super::node::sheet_index(file, sheet)?;
    let mut out = Vec::new();
    for entry in index.cells() {
        let Some(cell) = entry.cell(data) else {
            continue;
        };
        for (slot, link) in cell.hyperlinks().iter().enumerate() {
            out.push(Link {
                row: entry.row,
                column: entry.column,
                slot,
                href: link.href.clone(),
                text: link.text.clone(),
                title: link.title.clone(),
            });
        }
    }
    Ok(out)
}

/// The cell style a format's first styled rule applies.
fn style_name(format: &cf::Format) -> Option<&str> {
    format.rules.iter().find_map(|rule| match rule {
        cf::Rule::Condition(condition) => Some(condition.apply_style_name.as_str()),
        cf::Rule::DateIs(date) => Some(date.style.as_str()),
        _ => None,
    })
}

/// Resolve one feature step under a worksheet.
pub(super) fn lookup<'f>(
    file: &'f Ods,
    sheet: u32,
    kind: Kind,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let index = match select {
        Select::Ordinal(position) => Some(position.get()),
        Select::None => Some(0),
        Select::Last => None,
        _ => return Ok(Lookup::NotApplicable),
    };
    let mut features = list(file, sheet, kind)?;
    let index = index.or_else(|| features.len().checked_sub(1));
    Ok(match index.filter(|index| *index < features.len()) {
        Some(index) => Lookup::Found(Node::new_feature(file, sheet, features.swap_remove(index))),
        None => Lookup::Missing,
    })
}

impl<'f> Node<'f> {
    /// A worksheet feature.
    pub(super) fn new_feature(file: &'f Ods, sheet: u32, feature: Feature) -> Self {
        Self::in_sheet_item(file, sheet, Item::Feature(Arc::new(feature)))
    }
}

// ---------------------------------------------------------------------------
// Conditional-format rules

/// What a conditional format's rule does, in OfficeCLI's vocabulary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CfKind {
    CellIs,
    Formula,
    ColorScale,
    DataBar,
    IconSet,
    TopN,
    AboveAverage,
    DuplicateValues,
    UniqueValues,
    /// A text rule: 0 contains, 1 not contains, 2 begins with, 3 ends with.
    Text(u8),
    ContainsBlanks,
    NotContainsBlanks,
    ContainsErrors,
    NotContainsErrors,
    TimePeriod,
}

impl CfKind {
    /// OfficeCLI's `type`.
    pub(super) const fn name(self) -> &'static str {
        match self {
            Self::CellIs => "cellIs",
            Self::Formula => "formula",
            Self::ColorScale => "colorScale",
            Self::DataBar => "dataBar",
            Self::IconSet => "iconSet",
            Self::TopN => "topN",
            Self::AboveAverage => "aboveAverage",
            Self::DuplicateValues => "duplicateValues",
            Self::UniqueValues => "uniqueValues",
            Self::Text(1) => "notContainsText",
            Self::Text(2) => "beginsWith",
            Self::Text(3) => "endsWith",
            Self::Text(_) => "containsText",
            Self::ContainsBlanks => "containsBlanks",
            Self::NotContainsBlanks => "notContainsBlanks",
            Self::ContainsErrors => "containsErrors",
            Self::NotContainsErrors => "notContainsErrors",
            Self::TimePeriod => "timePeriod",
        }
    }
}

/// The OfficeCLI names of the cellIs comparisons with their ODF spelling.
pub(super) const CF_OPERATORS: [(&str, &str); 8] = [
    ("between", "between"),
    ("notBetween", "not-between"),
    ("equal", "="),
    ("notEqual", "!="),
    ("greaterThan", ">"),
    ("lessThan", "<"),
    ("greaterThanOrEqual", ">="),
    ("lessThanOrEqual", "<="),
];

/// The OfficeCLI time periods with their `calcext:date` value.
pub(super) const PERIODS: [(&str, cf::DateType); 10] = [
    ("today", cf::DateType::Today),
    ("tomorrow", cf::DateType::Tomorrow),
    ("yesterday", cf::DateType::Yesterday),
    ("last7Days", cf::DateType::Last7Days),
    ("lastMonth", cf::DateType::LastMonth),
    ("nextMonth", cf::DateType::NextMonth),
    ("thisWeek", cf::DateType::ThisWeek),
    ("nextWeek", cf::DateType::NextWeek),
    ("lastWeek", cf::DateType::LastWeek),
    ("thisMonth", cf::DateType::ThisMonth),
];

/// The projection of a format's first rule.
#[derive(Debug, Clone)]
pub(crate) struct Rule {
    pub(super) kind: CfKind,
    /// A cellIs comparison, as an index of [`CF_OPERATORS`].
    pub(super) operator: Option<usize>,
    /// cellIs operands in A1 syntax.
    pub(super) values: Vec<String>,
    /// A formula rule's formula in A1 syntax.
    pub(super) formula: Option<String>,
    pub(super) text: Option<String>,
    pub(super) rank: Option<u32>,
    pub(super) percent: bool,
    pub(super) bottom: bool,
    pub(super) below: bool,
    pub(super) equal: bool,
    pub(super) period: Option<usize>,
}

/// The arguments of `name(…)` split on top-level commas or semicolons.
fn call_args<'a>(text: &'a str, name: &str) -> Option<Vec<&'a str>> {
    let inner = text
        .strip_prefix(name)?
        .strip_prefix('(')?
        .strip_suffix(')')?;
    let mut out = Vec::new();
    let (mut depth, mut quoted, mut start) = (0usize, false, 0usize);
    for (at, char) in inner.char_indices() {
        match char {
            '"' => quoted = !quoted,
            '(' | '[' if !quoted => depth += 1,
            ')' | ']' if !quoted => depth = depth.saturating_sub(1),
            ',' | ';' if !quoted && depth == 0 => {
                out.push(inner[start..at].trim());
                start = at + 1;
            },
            _ => {},
        }
    }
    out.push(inner[start..].trim());
    Some(out)
}

/// A string literal's text (`"a""b"` -> `a"b`); other text as is.
pub(super) fn unquote(text: &str) -> String {
    text.strip_prefix('"')
        .and_then(|rest| rest.strip_suffix('"'))
        .map_or_else(|| text.to_owned(), |inner| inner.replace("\"\"", "\""))
}

impl Rule {
    fn new(kind: CfKind) -> Self {
        Self {
            kind,
            operator: None,
            values: Vec::new(),
            formula: None,
            text: None,
            rank: None,
            percent: false,
            bottom: false,
            below: false,
            equal: false,
            period: None,
        }
    }

    /// The projection of a format's first rule.
    pub(super) fn of(format: &cf::Format) -> Option<Self> {
        Some(match format.rules.first()? {
            cf::Rule::ColorScale(_) => Self::new(CfKind::ColorScale),
            cf::Rule::DataBar(_) => Self::new(CfKind::DataBar),
            cf::Rule::IconSet(_) => Self::new(CfKind::IconSet),
            cf::Rule::DateIs(date) => Self {
                period: PERIODS.iter().position(|(_, held)| *held == date.date),
                ..Self::new(CfKind::TimePeriod)
            },
            cf::Rule::Condition(condition) => Self::condition(&condition.condition)?,
        })
    }

    /// The projection of a `calcext:condition` value.
    fn condition(value: &str) -> Option<Self> {
        let value = value.trim();
        let a1 = |text: &str| from_odff(text.trim());
        for (index, name) in [(0, "between"), (1, "not-between")] {
            if let Some(args) = call_args(value, name) {
                return Some(Self {
                    operator: Some(index),
                    values: args.into_iter().map(a1).collect(),
                    ..Self::new(CfKind::CellIs)
                });
            }
        }
        if let Some(args) = call_args(value, "formula-is") {
            let formula = a1(args.first()?);
            let blank = |suffix: &str| {
                formula
                    .strip_prefix("LEN(TRIM(")
                    .and_then(|rest| rest.strip_suffix(suffix))
                    .is_some_and(|cell| !cell.contains('('))
            };
            let kind = if blank("))=0") {
                CfKind::ContainsBlanks
            } else if blank("))>0") {
                CfKind::NotContainsBlanks
            } else {
                CfKind::Formula
            };
            return Some(Self {
                formula: Some(formula),
                ..Self::new(kind)
            });
        }
        for (name, kind) in [
            ("contains-text", CfKind::Text(0)),
            ("not-contains-text", CfKind::Text(1)),
            ("begins-with", CfKind::Text(2)),
            ("ends-with", CfKind::Text(3)),
        ] {
            if let Some(args) = call_args(value, name) {
                return Some(Self {
                    text: Some(unquote(args.first()?)),
                    ..Self::new(kind)
                });
            }
        }
        for (name, percent, bottom) in [
            ("top-elements", false, false),
            ("bottom-elements", false, true),
            ("top-percent", true, false),
            ("bottom-percent", true, true),
        ] {
            if let Some(args) = call_args(value, name) {
                return Some(Self {
                    rank: args.first()?.parse().ok(),
                    percent,
                    bottom,
                    ..Self::new(CfKind::TopN)
                });
            }
        }
        let average = |below: bool, equal: bool| Self {
            below,
            equal,
            ..Self::new(CfKind::AboveAverage)
        };
        Some(match value {
            "duplicate" => Self::new(CfKind::DuplicateValues),
            "unique" => Self::new(CfKind::UniqueValues),
            "above-average" => average(false, false),
            "below-average" => average(true, false),
            "above-equal-average" => average(false, true),
            "below-equal-average" => average(true, true),
            "is-error" => Self::new(CfKind::ContainsErrors),
            "is-no-error" => Self::new(CfKind::NotContainsErrors),
            _ => {
                // A comparison: `>=5`, `!=[.A1]`, `=1`.
                for (index, (_, spelling)) in CF_OPERATORS.iter().enumerate().skip(2).rev() {
                    if let Some(operand) = value.strip_prefix(spelling) {
                        if index == 2 && operand.starts_with('=') {
                            continue;
                        }
                        return Some(Self {
                            operator: Some(index),
                            values: vec![a1(operand)],
                            ..Self::new(CfKind::CellIs)
                        });
                    }
                }
                return None;
            },
        })
    }
}

// ---------------------------------------------------------------------------
// Property tables

const fn read(name: &'static str, ty: ValueType) -> PropSpec {
    PropSpec::new(name, ty)
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access = Flags {
        get: false,
        set: false,
        add: false,
    };
    spec
}

const fn entry<S: Copy>(spec: PropSpec, access: Access<S>) -> Entry<S> {
    Entry {
        spec,
        access,
        constraint: Constraint::None,
        clear: false,
    }
}

const fn clearable<S: Copy>(spec: PropSpec, access: Access<S>) -> Entry<S> {
    Entry {
        spec,
        access,
        constraint: Constraint::None,
        clear: true,
    }
}

const fn alias(name: &'static str, entry: usize) -> Alias {
    Alias { name, entry }
}

const TEXT: ValueType = ValueType::Text;
const BOOL: ValueType = ValueType::Bool;
const INT: ValueType = ValueType::Int;
const NUMBER: ValueType = ValueType::Number;
const SCALAR: ValueType = ValueType::Scalar;
const SHEET_COLOR: ValueType = ValueType::Color(ColorContext::Spreadsheet);
const DRAWING_COLOR: ValueType = ValueType::Color(ColorContext::Drawing);
const FONT_SIZE: ValueType = ValueType::Length(LengthSpec::FONT_SIZE);

// Pending setters refuse with `unsupported_property`, so `--unsupported
// warn` skips them as OfficeCLI skips what it cannot apply.
const fn ods_gap(
    element: &'static str,
    key: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Property { element, key },
        code: Code::UnsupportedProperty,
        follow_up,
        reason,
    }
}

static DIRECTION_GAP: GapSpec = ods_gap(
    "conditionalFormatting",
    "direction",
    "record 0327: wontfix (format)",
    "an ODF data bar has no direction; LibreOffice draws bars from the axis",
);
static PRIORITY_GAP: GapSpec = ods_gap(
    "conditionalFormatting",
    "stopIfTrue",
    "record 0327: wontfix (format)",
    "ODF conditional formats have no stop-if-true flag; LibreOffice applies the first matching rule",
);
static STD_DEV_GAP: GapSpec = ods_gap(
    "conditionalFormatting",
    "stdDev",
    "record 0327: wontfix (format)",
    "an ODF average condition has no standard-deviation band",
);
static REVERSE_GAP: GapSpec = ods_gap(
    "conditionalFormatting",
    "reverse",
    "record 0327: wontfix (format)",
    "an ODF icon set has no reversed order",
);
static CRITERIA_GAP: GapSpec = ods_gap(
    "autofilter",
    "criteria0",
    "record 0327 follow-up: ODS auto-filter column criteria",
    "the auto-filter database range is authored; its table:filter conditions are kept but not authored",
);
static BREAK_SPAN_GAP: GapSpec = ods_gap(
    "rowbreak",
    "min",
    "record 0327: wontfix (format)",
    "an ODF page break is a row or column style property and spans the whole sheet",
);
static BREAK_SPAN_MAX_GAP: GapSpec = ods_gap(
    "rowbreak",
    "max",
    "record 0327: wontfix (format)",
    "an ODF page break is a row or column style property and spans the whole sheet",
);
static COL_BREAK_SPAN_GAP: GapSpec = ods_gap(
    "colbreak",
    "min",
    "record 0327: wontfix (format)",
    "an ODF page break is a row or column style property and spans the whole sheet",
);
static COL_BREAK_SPAN_MAX_GAP: GapSpec = ods_gap(
    "colbreak",
    "max",
    "record 0327: wontfix (format)",
    "an ODF page break is a row or column style property and spans the whole sheet",
);

const fn drawing_gap(element: &'static str, key: &'static str) -> GapSpec {
    ods_gap(
        element,
        key,
        "record 0327 follow-up: ODS drawing transforms, effects and text styling",
        "the ODS drawing editor authors frames, names, descriptions, text, solid fills and strokes; draw:transform, effects, crops and shape text styles are not authored",
    )
}

static ANCHOR_GAP: GapSpec = ods_gap(
    "picture",
    "anchor",
    "record 0327 follow-up: ODS cell-anchored drawings",
    "an ODS drawing added through the unified layer is placed on the sheet in lengths; anchoring it to cells needs the sheet's row and column geometry",
);
static SHAPE_ANCHOR_GAP: GapSpec = ods_gap(
    "shape",
    "anchor",
    "record 0327 follow-up: ODS cell-anchored drawings",
    "an ODS drawing added through the unified layer is placed on the sheet in lengths; anchoring it to cells needs the sheet's row and column geometry",
);
static PICTURE_ROTATION_GAP: GapSpec = drawing_gap("picture", "rotation");
static PICTURE_FLIP_GAP: GapSpec = drawing_gap("picture", "flip");
static PICTURE_FLIP_H_GAP: GapSpec = drawing_gap("picture", "flipH");
static PICTURE_FLIP_V_GAP: GapSpec = drawing_gap("picture", "flipV");
static PICTURE_FLIP_BOTH_GAP: GapSpec = drawing_gap("picture", "flipBoth");
static PICTURE_CROP_GAP: GapSpec = drawing_gap("picture", "crop");
static PICTURE_OPACITY_GAP: GapSpec = drawing_gap("picture", "opacity");
static PICTURE_SHADOW_GAP: GapSpec = drawing_gap("picture", "shadow");
static PICTURE_GLOW_GAP: GapSpec = drawing_gap("picture", "glow");
static PICTURE_REFLECTION_GAP: GapSpec = drawing_gap("picture", "reflection");
static PICTURE_SOFT_EDGE_GAP: GapSpec = drawing_gap("picture", "softEdge");
static SHAPE_ROTATION_GAP: GapSpec = drawing_gap("shape", "rotation");
static SHAPE_FLIP_GAP: GapSpec = drawing_gap("shape", "flip");
static SHAPE_FLIP_H_GAP: GapSpec = drawing_gap("shape", "flipH");
static SHAPE_FLIP_V_GAP: GapSpec = drawing_gap("shape", "flipV");
static SHAPE_FLIP_BOTH_GAP: GapSpec = drawing_gap("shape", "flipBoth");
static SHAPE_FONT_GAP: GapSpec = drawing_gap("shape", "font");
static SHAPE_SIZE_GAP: GapSpec = drawing_gap("shape", "size");
static SHAPE_BOLD_GAP: GapSpec = drawing_gap("shape", "bold");
static SHAPE_ITALIC_GAP: GapSpec = drawing_gap("shape", "italic");
static SHAPE_UNDERLINE_GAP: GapSpec = drawing_gap("shape", "underline");
static SHAPE_COLOR_GAP: GapSpec = drawing_gap("shape", "color");
static SHAPE_ALIGN_GAP: GapSpec = drawing_gap("shape", "align");
static SHAPE_VALIGN_GAP: GapSpec = drawing_gap("shape", "valign");
static SHAPE_MARGIN_GAP: GapSpec = drawing_gap("shape", "margin");
static SHAPE_GRADIENT_GAP: GapSpec = drawing_gap("shape", "gradientFill");
static SHAPE_SHADOW_GAP: GapSpec = drawing_gap("shape", "shadow");
static SHAPE_GLOW_GAP: GapSpec = drawing_gap("shape", "glow");
static SHAPE_REFLECTION_GAP: GapSpec = drawing_gap("shape", "reflection");
static SHAPE_SOFT_EDGE_GAP: GapSpec = drawing_gap("shape", "softEdge");

macro_rules! keyed {
    ($key:ident, $setter:ty, $table:ident, $props:ident, $keys:ident, [$($entry:expr => $variant:ident),+ $(,)?], [$($alias:expr),* $(,)?]) => {
        pub(crate) const $table: Table<
            $setter,
            { <[&str]>::len(&[$(stringify!($variant)),+]) },
            { <[&str]>::len(&[$(stringify!($alias)),*]) },
        > =
            Table::new([$($entry),+], [$($alias),*]);
        pub(crate) const $props: &[PropSpec] = &$table.readable();
        const $keys: &[$key] = &[$($key::$variant),+];
    };
}

/// Hyperlink keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HyperlinkKey {
    Ref,
    Url,
    Location,
    Display,
    Tooltip,
}

keyed!(HyperlinkKey, HyperlinkKey, HYPERLINK_TABLE, HYPERLINK_PROPS, HYPERLINK_KEYS, [
    entry(read("ref", TEXT), Access::SetAndAdd(HyperlinkKey::Ref)) => Ref,
    clearable(read("url", TEXT), Access::SetAndAdd(HyperlinkKey::Url)) => Url,
    clearable(read("location", TEXT), Access::SetAndAdd(HyperlinkKey::Location)) => Location,
    clearable(read("display", TEXT), Access::SetAndAdd(HyperlinkKey::Display)) => Display,
    clearable(read("tooltip", TEXT), Access::SetAndAdd(HyperlinkKey::Tooltip)) => Tooltip,
], [alias("link", 1), alias("screentip", 4)]);

fn hyperlink_value(link: &Link, key: HyperlinkKey) -> Option<Value<'_>> {
    match key {
        HyperlinkKey::Ref => Some(owned(link.reference())),
        HyperlinkKey::Url => link.url().map(text),
        HyperlinkKey::Location => link.location().map(owned),
        HyperlinkKey::Display => (!link.text.is_empty()).then(|| text(&link.text)),
        HyperlinkKey::Tooltip => link.title.as_deref().map(text),
    }
}

/// Picture keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PictureKey {
    Name,
    Alt,
    Title,
    Decorative,
    Hyperlink,
    X,
    Y,
    Width,
    Height,
    AnchorMode,
    ContentType,
    FileSize,
    Src,
    Pending,
}

keyed!(PictureKey, PictureKey, PICTURE_TABLE, PICTURE_PROPS, PICTURE_KEYS, [
    entry(read("name", TEXT), Access::SetAndAdd(PictureKey::Name)) => Name,
    clearable(read("alt", TEXT), Access::SetAndAdd(PictureKey::Alt)) => Alt,
    clearable(read("title", TEXT), Access::Add(PictureKey::Title)) => Title,
    entry(read("decorative", BOOL), Access::ReadOnly) => Decorative,
    entry(read("hyperlink", TEXT), Access::ReadOnly) => Hyperlink,
    entry(read("x", SCALAR), Access::SetAndAdd(PictureKey::X)) => X,
    entry(read("y", SCALAR), Access::SetAndAdd(PictureKey::Y)) => Y,
    entry(read("width", SCALAR), Access::SetAndAdd(PictureKey::Width)) => Width,
    entry(read("height", SCALAR), Access::SetAndAdd(PictureKey::Height)) => Height,
    entry(read("anchorMode", TEXT), Access::ReadOnly) => AnchorMode,
    entry(write_only("rotation", NUMBER), Access::Gap(&PICTURE_ROTATION_GAP)) => Pending,
    entry(write_only("flip", TEXT), Access::Gap(&PICTURE_FLIP_GAP)) => Pending,
    entry(read("contentType", TEXT), Access::ReadOnly) => ContentType,
    entry(read("fileSize", INT), Access::ReadOnly) => FileSize,
    entry(write_only("src", TEXT), Access::Add(PictureKey::Src)) => Src,
    entry(write_only("anchor", TEXT), Access::Gap(&ANCHOR_GAP)) => Pending,
    entry(write_only("flipH", BOOL), Access::Gap(&PICTURE_FLIP_H_GAP)) => Pending,
    entry(write_only("flipV", BOOL), Access::Gap(&PICTURE_FLIP_V_GAP)) => Pending,
    entry(write_only("flipBoth", BOOL), Access::Gap(&PICTURE_FLIP_BOTH_GAP)) => Pending,
    entry(write_only("crop", TEXT), Access::Gap(&PICTURE_CROP_GAP)) => Pending,
    entry(write_only("opacity", TEXT), Access::Gap(&PICTURE_OPACITY_GAP)) => Pending,
    entry(write_only("shadow", TEXT), Access::Gap(&PICTURE_SHADOW_GAP)) => Pending,
    entry(write_only("glow", TEXT), Access::Gap(&PICTURE_GLOW_GAP)) => Pending,
    entry(write_only("reflection", TEXT), Access::Gap(&PICTURE_REFLECTION_GAP)) => Pending,
    entry(write_only("softEdge", TEXT), Access::Gap(&PICTURE_SOFT_EDGE_GAP)) => Pending,
], [
    alias("alttext", 1),
    alias("description", 1),
    alias("image", 14),
    alias("path", 14),
]);

/// Shape keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ShapeKey {
    Name,
    Text,
    Geometry,
    X,
    Y,
    Width,
    Height,
    Fill,
    Line,
    Hyperlink,
    Alt,
    Pending,
}

keyed!(ShapeKey, ShapeKey, SHAPE_TABLE, SHAPE_PROPS, SHAPE_KEYS, [
    entry(read("name", TEXT), Access::SetAndAdd(ShapeKey::Name)) => Name,
    entry(read("text", TEXT), Access::Add(ShapeKey::Text)) => Text,
    entry(read("geometry", TEXT), Access::Add(ShapeKey::Geometry)) => Geometry,
    entry(read("x", SCALAR), Access::SetAndAdd(ShapeKey::X)) => X,
    entry(read("y", SCALAR), Access::SetAndAdd(ShapeKey::Y)) => Y,
    entry(read("width", SCALAR), Access::SetAndAdd(ShapeKey::Width)) => Width,
    entry(read("height", SCALAR), Access::SetAndAdd(ShapeKey::Height)) => Height,
    entry(read("fill", TEXT), Access::Add(ShapeKey::Fill)) => Fill,
    entry(read("line", TEXT), Access::Add(ShapeKey::Line)) => Line,
    entry(write_only("font", TEXT), Access::Gap(&SHAPE_FONT_GAP)) => Pending,
    entry(write_only("size", FONT_SIZE), Access::Gap(&SHAPE_SIZE_GAP)) => Pending,
    entry(write_only("bold", BOOL), Access::Gap(&SHAPE_BOLD_GAP)) => Pending,
    entry(write_only("italic", BOOL), Access::Gap(&SHAPE_ITALIC_GAP)) => Pending,
    entry(write_only("underline", TEXT), Access::Gap(&SHAPE_UNDERLINE_GAP)) => Pending,
    entry(write_only("color", DRAWING_COLOR), Access::Gap(&SHAPE_COLOR_GAP)) => Pending,
    entry(write_only("align", TEXT), Access::Gap(&SHAPE_ALIGN_GAP)) => Pending,
    entry(write_only("valign", TEXT), Access::Gap(&SHAPE_VALIGN_GAP)) => Pending,
    entry(write_only("rotation", NUMBER), Access::Gap(&SHAPE_ROTATION_GAP)) => Pending,
    entry(write_only("flip", TEXT), Access::Gap(&SHAPE_FLIP_GAP)) => Pending,
    entry(read("hyperlink", TEXT), Access::ReadOnly) => Hyperlink,
    clearable(read("alt", TEXT), Access::SetAndAdd(ShapeKey::Alt)) => Alt,
    entry(write_only("anchor", TEXT), Access::Gap(&SHAPE_ANCHOR_GAP)) => Pending,
    entry(write_only("flipH", BOOL), Access::Gap(&SHAPE_FLIP_H_GAP)) => Pending,
    entry(write_only("flipV", BOOL), Access::Gap(&SHAPE_FLIP_V_GAP)) => Pending,
    entry(write_only("flipBoth", BOOL), Access::Gap(&SHAPE_FLIP_BOTH_GAP)) => Pending,
    entry(write_only("margin", TEXT), Access::Gap(&SHAPE_MARGIN_GAP)) => Pending,
    entry(write_only("gradientFill", TEXT), Access::Gap(&SHAPE_GRADIENT_GAP)) => Pending,
    entry(write_only("shadow", TEXT), Access::Gap(&SHAPE_SHADOW_GAP)) => Pending,
    entry(write_only("glow", TEXT), Access::Gap(&SHAPE_GLOW_GAP)) => Pending,
    entry(write_only("reflection", TEXT), Access::Gap(&SHAPE_REFLECTION_GAP)) => Pending,
    entry(write_only("softEdge", TEXT), Access::Gap(&SHAPE_SOFT_EDGE_GAP)) => Pending,
], [
    alias("border", 8),
    alias("font.bold", 11),
    alias("font.color", 14),
    alias("font.italic", 12),
    alias("font.name", 9),
    alias("font.size", 10),
    alias("font.underline", 13),
    alias("linecolor", 8),
    alias("preset", 2),
    alias("ref", 21),
    alias("rot", 17),
    alias("rotate", 17),
]);

/// OfficeCLI geometry of a drawing: lengths on the sheet, or a cell origin
/// with a length (one end cell) or cell (two cells) extent.
fn geometry(drawing: &sf::Drawing, axis: usize) -> Option<Value<'static>> {
    let frame = drawing.frame;
    match (drawing.cell, drawing.end_cell) {
        (None, _) => frame[axis].map(emu),
        (Some((row, column)), end) => Some(match (axis, end) {
            (0, _) => int(column),
            (1, _) => int(row),
            (2, Some((_, end))) => int(end.saturating_sub(column)),
            (3, Some((end, _))) => int(end.saturating_sub(row)),
            (_, None) => frame[axis].map(emu)?,
            _ => return None,
        }),
    }
}

fn anchor_mode(drawing: &sf::Drawing) -> Value<'static> {
    Value::Token(match (drawing.cell, drawing.end_cell) {
        (None, _) => "absolute",
        (Some(_), Some(_)) => "twoCell",
        (Some(_), None) => "oneCell",
    })
}

/// The OfficeCLI presets with their `draw:type`.
pub(super) const GEOMETRIES: [(&str, &str); 12] = [
    ("rect", "rectangle"),
    ("roundRect", "round-rectangle"),
    ("ellipse", "ellipse"),
    ("diamond", "diamond"),
    ("triangle", "isosceles-triangle"),
    ("rtTriangle", "right-triangle"),
    ("parallelogram", "parallelogram"),
    ("trapezoid", "trapezoid"),
    ("hexagon", "hexagon"),
    ("octagon", "octagon"),
    ("star5", "star5"),
    ("rightArrow", "right-arrow"),
];

/// The media type of a picture by its file name.
pub(super) fn media_type(path: &str) -> Option<&'static str> {
    let extension = path.rsplit('.').next()?.to_ascii_lowercase();
    Some(match extension.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "svg" => "image/svg+xml",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        _ => return None,
    })
}

fn picture_value(drawing: &sf::Drawing, key: PictureKey) -> Option<Value<'_>> {
    match key {
        PictureKey::Name => drawing.name.as_deref().map(text),
        PictureKey::Alt => drawing.description.as_deref().map(text),
        PictureKey::Title => drawing.title.as_deref().map(text),
        PictureKey::X => geometry(drawing, 0),
        PictureKey::Y => geometry(drawing, 1),
        PictureKey::Width => geometry(drawing, 2),
        PictureKey::Height => geometry(drawing, 3),
        PictureKey::AnchorMode => Some(anchor_mode(drawing)),
        PictureKey::ContentType => drawing
            .image
            .as_deref()
            .and_then(media_type)
            .map(Value::Token),
        PictureKey::Decorative
        | PictureKey::Hyperlink
        | PictureKey::FileSize
        | PictureKey::Src
        | PictureKey::Pending => None,
    }
}

/// A stroke width as points.
fn stroke_points(width: &str) -> Option<f64> {
    let length = Length::parse(width.trim()).ok()?;
    Some((length.points() * 100.0).round() / 100.0)
}

fn shape_value(drawing: &sf::Drawing, key: ShapeKey) -> Option<Value<'_>> {
    match key {
        ShapeKey::Name => drawing.name.as_deref().map(text),
        ShapeKey::Text => drawing.text.as_deref().map(text),
        ShapeKey::Geometry => {
            let held = drawing.geometry.as_deref()?;
            Some(
                GEOMETRIES
                    .iter()
                    .find(|(_, odf)| *odf == held)
                    .map_or_else(|| text(held), |(name, _)| Value::Token(name)),
            )
        },
        ShapeKey::X => geometry(drawing, 0),
        ShapeKey::Y => geometry(drawing, 1),
        ShapeKey::Width => geometry(drawing, 2),
        ShapeKey::Height => geometry(drawing, 3),
        ShapeKey::Fill => match drawing.fill.as_ref()? {
            None => Some(Value::Token("none")),
            Some(color) => Some(owned(color.to_ascii_uppercase())),
        },
        ShapeKey::Line => match drawing.line.as_ref()? {
            None => Some(Value::Token("none")),
            Some((color, width)) => Some(owned(match width.as_deref().and_then(stroke_points) {
                Some(points) => format!(
                    "{}:{}",
                    color.to_ascii_uppercase(),
                    crate::value::DotNetDouble(points)
                ),
                None => color.to_ascii_uppercase(),
            })),
        },
        ShapeKey::Alt => drawing.description.as_deref().map(text),
        ShapeKey::Hyperlink | ShapeKey::Pending => None,
    }
}

/// Conditional-format keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CfKey {
    Type,
    Ref,
    Priority,
    DxfId,
    Operator,
    Value,
    Value2,
    Formula,
    Text,
    Rank,
    Percent,
    Bottom,
    AboveAverage,
    EqualAverage,
    Period,
    Color,
    Min,
    Max,
    MinLength,
    MaxLength,
    ShowValue,
    MinColor,
    MidColor,
    MaxColor,
    Midpoint,
    Iconset,
    Fill,
    FontColor,
    FontBold,
    FontItalic,
    FontStrike,
    FontUnderline,
    FontSize,
    FontName,
    NegativeColor,
    AxisColor,
    AxisPosition,
    Pending,
}

keyed!(CfKey, CfKey, CF_TABLE, CF_PROPS, CF_KEYS, [
    entry(read("type", TEXT), Access::Add(CfKey::Type)) => Type,
    entry(read("ref", TEXT), Access::SetAndAdd(CfKey::Ref)) => Ref,
    entry(read("priority", INT), Access::ReadOnly) => Priority,
    entry(write_only("stopIfTrue", BOOL), Access::Gap(&PRIORITY_GAP)) => Pending,
    entry(read("dxfId", INT), Access::ReadOnly) => DxfId,
    entry(read("operator", TEXT), Access::SetAndAdd(CfKey::Operator)) => Operator,
    entry(read("value", TEXT), Access::SetAndAdd(CfKey::Value)) => Value,
    entry(read("value2", TEXT), Access::SetAndAdd(CfKey::Value2)) => Value2,
    entry(read("formula", TEXT), Access::SetAndAdd(CfKey::Formula)) => Formula,
    entry(read("text", TEXT), Access::SetAndAdd(CfKey::Text)) => Text,
    entry(read("rank", INT), Access::SetAndAdd(CfKey::Rank)) => Rank,
    entry(read("percent", BOOL), Access::SetAndAdd(CfKey::Percent)) => Percent,
    entry(read("bottom", BOOL), Access::SetAndAdd(CfKey::Bottom)) => Bottom,
    entry(read("aboveAverage", BOOL), Access::SetAndAdd(CfKey::AboveAverage)) => AboveAverage,
    entry(read("equalAverage", BOOL), Access::SetAndAdd(CfKey::EqualAverage)) => EqualAverage,
    entry(write_only("stdDev", INT), Access::Gap(&STD_DEV_GAP)) => Pending,
    entry(read("period", TEXT), Access::SetAndAdd(CfKey::Period)) => Period,
    entry(read("color", SHEET_COLOR), Access::SetAndAdd(CfKey::Color)) => Color,
    entry(read("min", TEXT), Access::SetAndAdd(CfKey::Min)) => Min,
    entry(read("max", TEXT), Access::SetAndAdd(CfKey::Max)) => Max,
    entry(read("minLength", INT), Access::SetAndAdd(CfKey::MinLength)) => MinLength,
    entry(read("maxLength", INT), Access::SetAndAdd(CfKey::MaxLength)) => MaxLength,
    entry(read("showValue", BOOL), Access::SetAndAdd(CfKey::ShowValue)) => ShowValue,
    entry(read("minColor", SHEET_COLOR), Access::SetAndAdd(CfKey::MinColor)) => MinColor,
    entry(read("midColor", SHEET_COLOR), Access::SetAndAdd(CfKey::MidColor)) => MidColor,
    entry(read("maxColor", SHEET_COLOR), Access::SetAndAdd(CfKey::MaxColor)) => MaxColor,
    entry(read("midpoint", TEXT), Access::SetAndAdd(CfKey::Midpoint)) => Midpoint,
    entry(read("iconset", TEXT), Access::SetAndAdd(CfKey::Iconset)) => Iconset,
    entry(write_only("reverse", BOOL), Access::Gap(&REVERSE_GAP)) => Pending,
    entry(read("fill", SHEET_COLOR), Access::SetAndAdd(CfKey::Fill)) => Fill,
    entry(read("font.color", SHEET_COLOR), Access::SetAndAdd(CfKey::FontColor)) => FontColor,
    entry(read("font.bold", BOOL), Access::SetAndAdd(CfKey::FontBold)) => FontBold,
    entry(read("font.italic", BOOL), Access::SetAndAdd(CfKey::FontItalic)) => FontItalic,
    entry(read("font.strike", BOOL), Access::SetAndAdd(CfKey::FontStrike)) => FontStrike,
    entry(read("font.underline", TEXT), Access::SetAndAdd(CfKey::FontUnderline)) => FontUnderline,
    entry(read("font.size", FONT_SIZE), Access::SetAndAdd(CfKey::FontSize)) => FontSize,
    entry(read("font.name", TEXT), Access::SetAndAdd(CfKey::FontName)) => FontName,
    entry(read("negativeColor", SHEET_COLOR), Access::SetAndAdd(CfKey::NegativeColor)) => NegativeColor,
    entry(read("axisColor", SHEET_COLOR), Access::SetAndAdd(CfKey::AxisColor)) => AxisColor,
    entry(read("axisPosition", TEXT), Access::SetAndAdd(CfKey::AxisPosition)) => AxisPosition,
    entry(write_only("direction", TEXT), Access::Gap(&DIRECTION_GAP)) => Pending,
], [
    alias("above", 13),
    alias("formula1", 8),
    alias("formula2", 7),
    alias("icons", 27),
    alias("maxvalue", 7),
    alias("range", 1),
    alias("rule", 0),
    alias("sqref", 1),
    alias("timeperiod", 16),
    alias("value1", 6),
]);

/// A threshold's value: a number, percent, percentile or formula.
fn threshold(kind: EntryType, value: &str) -> Option<Value<'static>> {
    match kind {
        EntryType::Number | EntryType::Percent | EntryType::Percentile => {
            Some(owned(value.to_owned()))
        },
        EntryType::Formula => Some(owned(from_odff(value))),
        _ => None,
    }
}

fn cf_value(feature: &CfFeature, key: CfKey) -> Option<Value<'_>> {
    let format = &feature.format;
    let style = &feature.style;
    let rule = feature.rule.as_ref();
    let flag = |value: bool| value.then_some(Value::Bool(true));
    let first = format.rules.first();
    let bar = match first {
        Some(cf::Rule::DataBar(bar)) => Some(bar),
        _ => None,
    };
    let scale = match first {
        Some(cf::Rule::ColorScale(scale)) => Some(scale),
        _ => None,
    };
    let icons = match first {
        Some(cf::Rule::IconSet(icons)) => Some(icons),
        _ => None,
    };
    let kind = rule.map(|rule| rule.kind);
    match key {
        CfKey::Type => kind.map(|kind| Value::Token(kind.name())),
        CfKey::Ref => Some(owned(address_labels(&format.target_range_addresses))),
        CfKey::Operator => match kind? {
            CfKind::CellIs => rule?
                .operator
                .and_then(|operator| CF_OPERATORS.get(operator))
                .map(|(name, _)| Value::Token(name)),
            CfKind::Text(operator) => Some(Value::Token(
                ["containsText", "notContains", "beginsWith", "endsWith"][usize::from(operator)],
            )),
            _ => None,
        },
        CfKey::Value => (kind? == CfKind::CellIs)
            .then(|| rule?.values.first().map(|value| text(value)))
            .flatten(),
        CfKey::Value2 => (kind? == CfKind::CellIs)
            .then(|| rule?.values.get(1).map(|value| text(value)))
            .flatten(),
        CfKey::Formula => (kind? == CfKind::Formula)
            .then(|| rule?.formula.as_deref().map(text))
            .flatten(),
        CfKey::Text => rule?.text.as_deref().map(text),
        CfKey::Rank => rule?.rank.map(int),
        CfKey::Percent => flag(rule?.percent),
        CfKey::Bottom => flag(rule?.bottom),
        CfKey::AboveAverage => {
            (kind? == CfKind::AboveAverage && rule?.below).then_some(Value::Bool(false))
        },
        CfKey::EqualAverage => flag(kind? == CfKind::AboveAverage && rule?.equal),
        CfKey::Period => rule?
            .period
            .and_then(|index| PERIODS.get(index))
            .map(|(name, _)| Value::Token(name)),
        CfKey::Color => bar?.positive_color.as_deref().and_then(color),
        CfKey::NegativeColor => bar?.negative_color.as_deref().and_then(color),
        CfKey::AxisColor => bar?.axis_color.as_deref().and_then(color),
        CfKey::AxisPosition => bar?
            .axis_position
            .map(|position| Value::Token(position.as_str())),
        CfKey::Min => bar?
            .entries
            .first()
            .and_then(|entry| threshold(entry.entry_type, &entry.value)),
        CfKey::Max => bar?
            .entries
            .get(1)
            .and_then(|entry| threshold(entry.entry_type, &entry.value)),
        CfKey::MinLength => bar?
            .min_length
            .as_deref()
            .and_then(|value| value.parse::<f64>().ok())
            .and_then(Value::number),
        CfKey::MaxLength => bar?
            .max_length
            .as_deref()
            .and_then(|value| value.parse::<f64>().ok())
            .and_then(Value::number),
        CfKey::ShowValue => match (bar, icons) {
            (Some(bar), _) => (bar.show_value == Some(false)).then_some(Value::Bool(false)),
            (None, Some(icons)) => (icons.show_value == Some(false)).then_some(Value::Bool(false)),
            (None, None) => None,
        },
        CfKey::MinColor => scale?.entries.first().and_then(|entry| color(&entry.color)),
        CfKey::MidColor => scale
            .filter(|scale| scale.entries.len() == 3)
            .and_then(|scale| color(&scale.entries[1].color)),
        CfKey::MaxColor => scale?.entries.last().and_then(|entry| color(&entry.color)),
        CfKey::Midpoint => scale
            .filter(|scale| scale.entries.len() == 3)
            .and_then(|scale| threshold(scale.entries[1].entry_type, &scale.entries[1].value)),
        CfKey::Iconset => icons.map(|icons| owned(icons.icon_set_type.as_str().to_owned())),
        CfKey::Fill => style.background.as_deref().and_then(color),
        CfKey::FontColor => style.color.as_deref().and_then(color),
        CfKey::FontBold => style.bold.map(Value::Bool),
        CfKey::FontItalic => style.italic.map(Value::Bool),
        CfKey::FontStrike => style.strike.map(Value::Bool),
        CfKey::FontUnderline => style
            .underline
            .filter(|underline| *underline)
            .map(|_| Value::Token("single")),
        CfKey::FontSize => style
            .size_pt
            .map(|points| Value::Length(Length::from_emu((points * 12_700.0).round() as i64))),
        CfKey::FontName => style.font.as_deref().map(text),
        CfKey::Priority | CfKey::DxfId | CfKey::Pending => None,
    }
}

/// Data-validation keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ValidationKey {
    Ref,
    Type,
    Operator,
    Formula1,
    Formula2,
    AllowBlank,
    ShowError,
    ShowInput,
    ErrorTitle,
    Error,
    PromptTitle,
    Prompt,
    ErrorStyle,
    InCellDropdown,
}

keyed!(ValidationKey, ValidationKey, VALIDATION_TABLE, VALIDATION_PROPS, VALIDATION_KEYS, [
    entry(read("ref", TEXT), Access::SetAndAdd(ValidationKey::Ref)) => Ref,
    entry(read("type", TEXT), Access::SetAndAdd(ValidationKey::Type)) => Type,
    entry(read("operator", TEXT), Access::SetAndAdd(ValidationKey::Operator)) => Operator,
    clearable(read("formula1", TEXT), Access::SetAndAdd(ValidationKey::Formula1)) => Formula1,
    clearable(read("formula2", TEXT), Access::SetAndAdd(ValidationKey::Formula2)) => Formula2,
    entry(read("allowBlank", BOOL), Access::SetAndAdd(ValidationKey::AllowBlank)) => AllowBlank,
    entry(read("showError", BOOL), Access::SetAndAdd(ValidationKey::ShowError)) => ShowError,
    entry(read("showInput", BOOL), Access::SetAndAdd(ValidationKey::ShowInput)) => ShowInput,
    clearable(read("errorTitle", TEXT), Access::SetAndAdd(ValidationKey::ErrorTitle)) => ErrorTitle,
    clearable(read("error", TEXT), Access::SetAndAdd(ValidationKey::Error)) => Error,
    clearable(read("promptTitle", TEXT), Access::SetAndAdd(ValidationKey::PromptTitle)) => PromptTitle,
    clearable(read("prompt", TEXT), Access::SetAndAdd(ValidationKey::Prompt)) => Prompt,
    entry(read("errorStyle", TEXT), Access::SetAndAdd(ValidationKey::ErrorStyle)) => ErrorStyle,
    entry(read("inCellDropdown", BOOL), Access::SetAndAdd(ValidationKey::InCellDropdown)) => InCellDropdown,
], [alias("range", 0), alias("sqref", 0)]);

/// The OfficeCLI validation operators with their ODF comparison.
pub(super) const VALIDATION_OPERATORS: [(&str, &str); 8] = [
    ("between", "between"),
    ("notBetween", "not-between"),
    ("equal", "="),
    ("notEqual", "!="),
    ("greaterThan", ">"),
    ("lessThan", "<"),
    ("greaterThanOrEqual", ">="),
    ("lessThanOrEqual", "<="),
];

/// The projection of a `table:condition`: OfficeCLI type, operator index
/// and formulas.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Condition {
    pub(super) kind: &'static str,
    pub(super) operator: Option<usize>,
    pub(super) formula1: Option<String>,
    pub(super) formula2: Option<String>,
}

/// A validation operand in A1 syntax.
fn operand(text: &str) -> String {
    from_odff(text.trim())
}

/// The projection of a `table:condition` value.
pub(super) fn parse_condition(condition: &str) -> Option<Condition> {
    let condition = condition.trim();
    let condition = condition
        .split_once(':')
        .filter(|(prefix, _)| matches!(*prefix, "of" | "oooc" | "msoxl"))
        .map_or(condition, |(_, rest)| rest)
        .trim();
    let plain = |kind| Condition {
        kind,
        operator: None,
        formula1: None,
        formula2: None,
    };
    if let Some(args) = call_args(condition, "cell-content-is-in-list") {
        let joined = args.join(";");
        let formula1 = if args.iter().all(|arg| arg.starts_with('"')) {
            format!(
                "\"{}\"",
                args.iter()
                    .map(|arg| unquote(arg))
                    .collect::<Vec<_>>()
                    .join(",")
            )
        } else {
            operand(&joined)
        };
        return Some(Condition {
            formula1: Some(formula1),
            ..plain("list")
        });
    }
    if let Some(args) = call_args(condition, "is-true-formula") {
        return Some(Condition {
            formula1: Some(operand(args.first()?)),
            ..plain("custom")
        });
    }
    // `cell-content-is-whole-number() and <comparison>`, or a text-length
    // comparison.
    let (kind, comparison) = if let Some((head, tail)) = condition.split_once(" and ") {
        let kind = match head.trim() {
            "cell-content-is-whole-number()" => "whole",
            "cell-content-is-decimal-number()" => "decimal",
            "cell-content-is-date()" => "date",
            "cell-content-is-time()" => "time",
            _ => return None,
        };
        (kind, tail.trim())
    } else {
        ("textLength", condition)
    };
    let subject = if kind == "textLength" {
        "cell-content-text-length"
    } else {
        "cell-content"
    };
    for (index, name) in [(0, "is-between"), (1, "is-not-between")] {
        if let Some(args) = call_args(comparison, &format!("{subject}-{name}")) {
            return Some(Condition {
                operator: Some(index),
                formula1: args.first().map(|arg| operand(arg)),
                formula2: args.get(1).map(|arg| operand(arg)),
                ..plain(kind)
            });
        }
    }
    let rest = comparison
        .strip_prefix(subject)?
        .strip_prefix("()")?
        .trim_start();
    for (index, (_, spelling)) in VALIDATION_OPERATORS.iter().enumerate().skip(2).rev() {
        if let Some(value) = rest.strip_prefix(spelling) {
            if index == 2 && value.starts_with('=') {
                continue;
            }
            return Some(Condition {
                operator: Some(index),
                formula1: Some(operand(value)),
                ..plain(kind)
            });
        }
    }
    None
}

fn validation_value(validation: &Validation, key: ValidationKey) -> Option<Value<'_>> {
    let condition = parse_condition(&validation.condition);
    match key {
        ValidationKey::Ref => Some(owned(ranges(&validation.ranges))),
        ValidationKey::Type => Some(Value::Token(
            condition
                .as_ref()
                .map_or("custom", |condition| condition.kind),
        )),
        ValidationKey::Operator => condition?
            .operator
            .and_then(|operator| VALIDATION_OPERATORS.get(operator))
            .map(|(name, _)| Value::Token(name)),
        ValidationKey::Formula1 => match condition {
            Some(condition) => condition.formula1.map(owned),
            None => Some(text(&validation.condition)),
        },
        ValidationKey::Formula2 => condition?.formula2.map(owned),
        ValidationKey::AllowBlank => Some(Value::Bool(validation.allow_empty)),
        ValidationKey::ShowError => Some(Value::Bool(
            validation.error.as_ref().is_some_and(|error| error.3),
        )),
        ValidationKey::ShowInput => Some(Value::Bool(
            validation.help.as_ref().is_some_and(|help| help.2),
        )),
        ValidationKey::ErrorTitle => validation.error.as_ref()?.1.as_deref().map(text),
        ValidationKey::Error => validation.error.as_ref()?.2.as_deref().map(text),
        ValidationKey::PromptTitle => validation.help.as_ref()?.0.as_deref().map(text),
        ValidationKey::Prompt => validation.help.as_ref()?.1.as_deref().map(text),
        ValidationKey::ErrorStyle => Some(Value::Token(
            match validation.error.as_ref().map(|error| error.0.as_str()) {
                Some("warning") => "warning",
                Some("information") => "information",
                _ => "stop",
            },
        )),
        ValidationKey::InCellDropdown => Some(Value::Bool(validation.display_list)),
    }
}

/// Auto-filter keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum AutofilterKey {
    Range,
    Criteria,
}

keyed!(AutofilterKey, AutofilterKey, AUTOFILTER_TABLE, AUTOFILTER_PROPS, AUTOFILTER_KEYS, [
    entry(read("range", TEXT), Access::SetAndAdd(AutofilterKey::Range)) => Range,
    entry(write_only("criteria0", TEXT), Access::Gap(&CRITERIA_GAP)) => Criteria,
], [alias("ref", 0)]);

/// Page-break keys; row breaks name their row, column breaks their column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BreakKey {
    Index,
    Manual,
    Min,
    Max,
}

keyed!(BreakKey, BreakKey, ROW_BREAK_TABLE, ROW_BREAK_PROPS, BREAK_KEYS, [
    entry(read("row", INT), Access::SetAndAdd(BreakKey::Index)) => Index,
    entry(read("manual", BOOL), Access::SetAndAdd(BreakKey::Manual)) => Manual,
    entry(write_only("min", INT), Access::Gap(&BREAK_SPAN_GAP)) => Min,
    entry(write_only("max", INT), Access::Gap(&BREAK_SPAN_MAX_GAP)) => Max,
], [alias("index", 0)]);

pub(crate) const COL_BREAK_TABLE: Table<BreakKey, 4, 2> = Table::new(
    [
        entry(read("col", SCALAR), Access::SetAndAdd(BreakKey::Index)),
        entry(read("manual", BOOL), Access::SetAndAdd(BreakKey::Manual)),
        entry(write_only("min", INT), Access::Gap(&COL_BREAK_SPAN_GAP)),
        entry(write_only("max", INT), Access::Gap(&COL_BREAK_SPAN_MAX_GAP)),
    ],
    [alias("column", 0), alias("index", 0)],
);
pub(crate) const COL_BREAK_PROPS: &[PropSpec] = &COL_BREAK_TABLE.readable();

/// The `Code` of a missing feature, for the write module.
pub(super) fn missing(message: &'static str) -> Error {
    Error::new(Code::PathNotFound, message).with_format(litchi_core::FileFormat::Ods)
}
