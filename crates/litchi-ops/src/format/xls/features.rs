//! XLS worksheet features as unified nodes (record 0327): legacy notes
//! (`comment`), hyperlinks, drawing pictures and shapes, conditional
//! formats (`cf`), data validations, the auto-filter and manual page
//! breaks, with the XLSX property vocabulary of record 0320 (as XLSB took it
//! in record 0322).
//!
//! Every feature is a child of its worksheet, addressed by its position
//! among the worksheet's features of its kind (`/Sheet1/comment[2]`,
//! `/Sheet1/cf[1]`; the auto-filter is `/Sheet1/autofilter`). A node holds
//! one [`Feature`] behind one `Arc`. Listing a kind reads the worksheet
//! substream's records through `litchi_xls::sheet_features` and never
//! decodes the worksheet's cells.
//!
//! Positions follow the record order, except notes, which are numbered in
//! row-major order of their cells, as XLSX numbers them. Conditional
//! formats are the classic `CondFmt` and the Office 2007 `CondFmt12`
//! records, in record order. The write verbs live in `feature_write`.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::Position;
use litchi_xls::sheet_features::{
    self as sf, AutoFilter, ConditionalFormat, DrawingKind, FeatureSet, Fill, Line, PageBreak,
    RuleType, Threshold, Validation,
};

use super::Xls;
use super::node::{Item, Node, Sheet};
use super::owner::SheetHandle;
use super::spec::Kind;
use crate::error::{Code, Result};
use crate::node::spec::{Access as Flags, GapSpec, PropSpec, Subject};
use crate::node::{PathStep, Preview, Value};
use crate::path::navigate::{Lookup, Select};
use crate::registry::{Access, Alias, Constraint, Entry, Table};
use crate::value::{Color, ColorContext, Length, LengthSpec, ValueType};

/// The feature kinds of a worksheet, in listing order.
pub(super) const KINDS: [Kind; 9] = [
    Kind::Comment,
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
    Comment(Note),
    Hyperlink(Box<litchi_xls::hyperlinks::Hyperlink>),
    Drawing(Box<sf::Drawing>),
    Cf(Box<ConditionalFormat>),
    Validation(Box<Validation>),
    Autofilter(AutoFilter),
    Break(PageBreak),
}

/// A legacy note.
#[derive(Clone)]
pub(crate) struct Note {
    pub(super) row: u32,
    pub(super) col: u32,
    /// The note's A1 cell, inline.
    pub(super) reference: crate::format::workbook::Label,
    pub(super) author: String,
    pub(super) text: String,
}

/// The A1 label of a rectangle.
pub(super) fn area_label(area: sf::Area) -> String {
    crate::format::workbook::Label::range(
        (u32::from(area.first_column), u32::from(area.first_row)),
        (u32::from(area.last_column), u32::from(area.last_row)),
    )
    .as_str()
    .to_owned()
}

/// The A1 reference of a hyperlink's cell or range.
pub(super) fn link_reference(link: &litchi_xls::hyperlinks::Hyperlink) -> String {
    let range = link.range();
    area_label(sf::Area {
        first_row: range.first_row(),
        last_row: range.last_row(),
        first_column: u16::from(range.first_column()),
        last_column: u16::from(range.last_column()),
    })
}

/// A hyperlink's external target (its moniker's address).
pub(super) fn link_url(link: &litchi_xls::hyperlinks::Hyperlink) -> Option<&str> {
    link.moniker().and(link.address())
}

/// The space-separated ranges of a format or validation.
pub(super) fn ranges(areas: &[sf::Area]) -> String {
    areas
        .iter()
        .map(|area| area_label(*area))
        .collect::<Vec<_>>()
        .join(" ")
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
            Data::Comment(note) => Cow::Borrowed(note.text.as_str()),
            Data::Hyperlink(link) => Cow::Borrowed(
                link.display_name()
                    .or_else(|| link_url(link))
                    .or_else(|| link.location())
                    .unwrap_or(""),
            ),
            Data::Drawing(drawing) => match drawing.kind {
                DrawingKind::Shape => Cow::Borrowed(drawing.text.as_deref()?),
                DrawingKind::Picture => Cow::Borrowed(drawing.description.as_deref()?),
            },
            Data::Cf(cf) => Cow::Owned(ranges(&cf.ranges)),
            Data::Validation(validation) => Cow::Owned(ranges(&validation.ranges)),
            Data::Autofilter(filter) => Cow::Owned(area_label(filter.area?)),
            Data::Break(_) => return None,
        })
    }

    /// The node view's one-line preview.
    pub(super) fn preview(&self) -> Option<Preview<'_>> {
        Some(Preview::Text(match &self.data {
            Data::Comment(note) => {
                Cow::Owned(format!("{}: {}", note.reference.as_str(), note.text))
            },
            Data::Hyperlink(link) => Cow::Owned(link_reference(link)),
            Data::Drawing(drawing) => Cow::Borrowed(drawing.name.as_deref()?),
            Data::Cf(cf) => Cow::Owned(format!("{} {}", rule_type(cf), ranges(&cf.ranges))),
            Data::Validation(validation) => Cow::Owned(ranges(&validation.ranges)),
            Data::Autofilter(filter) => Cow::Owned(area_label(filter.area?)),
            Data::Break(item) => Cow::Owned(item.index.to_string()),
        }))
    }

    /// The property table of the feature's kind.
    pub(super) fn table(&self) -> &'static [PropSpec] {
        props(self.kind)
    }

    /// The value at `index` of the feature's property table; `sheet` is the
    /// name of the feature's worksheet.
    pub(super) fn value(&self, index: usize, sheet: &str) -> Option<Value<'_>> {
        match &self.data {
            Data::Comment(note) => match *COMMENT_KEYS.get(index)? {
                CommentKey::AnchoredTo => {
                    Some(owned(format!("/{sheet}/{}", note.reference.as_str())))
                },
                key => comment_value(note, key),
            },
            Data::Hyperlink(link) => hyperlink_value(link, *HYPERLINK_KEYS.get(index)?),
            Data::Drawing(drawing) => match self.kind {
                Kind::Picture => picture_value(drawing, *PICTURE_KEYS.get(index)?),
                _ => shape_value(drawing, *SHAPE_KEYS.get(index)?),
            },
            Data::Cf(cf) => cf_value(cf, *CF_KEYS.get(index)?),
            Data::Validation(validation) => {
                validation_value(validation, *VALIDATION_KEYS.get(index)?)
            },
            Data::Autofilter(filter) => match *AUTOFILTER_KEYS.get(index)? {
                AutofilterKey::Range => filter.area.map(|area| owned(area_label(area))),
                AutofilterKey::Criteria => None,
            },
            Data::Break(item) => break_value(self.kind, *item, *BREAK_KEYS.get(index)?),
        }
    }
}

/// The path step name of a feature kind.
pub(super) const fn step_name(kind: Kind) -> &'static str {
    match kind {
        Kind::Comment => "comment",
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
        Kind::Comment => COMMENT_PROPS,
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

fn rgb(color: u32) -> Value<'static> {
    let [_, red, green, blue] = color.to_be_bytes();
    Value::Color(Color::rgb(red, green, blue))
}

fn points(twips: u32) -> Value<'static> {
    // One twip is 635 EMU.
    Value::Length(Length::from_emu(i64::from(twips) * 635))
}

// ---------------------------------------------------------------------------
// Listing

/// The selection that reads `kind`.
fn selection(kind: Kind) -> FeatureSet {
    match kind {
        Kind::Comment => FeatureSet::NOTES,
        Kind::Hyperlink => FeatureSet::HYPERLINKS,
        Kind::Picture | Kind::Shape => FeatureSet::DRAWINGS,
        Kind::Cf => FeatureSet::FORMATS,
        Kind::DataValidation => FeatureSet::VALIDATIONS,
        Kind::Autofilter => FeatureSet::AUTO_FILTER,
        _ => FeatureSet::PAGE_BREAKS,
    }
}

/// The notes of a worksheet in row-major order of their cells.
pub(super) fn notes(records: Vec<sf::Note>) -> Vec<Note> {
    let mut notes: Vec<Note> = records
        .into_iter()
        .map(|record| Note {
            row: u32::from(record.row),
            col: u32::from(record.column),
            reference: crate::format::workbook::Label::cell(
                u32::from(record.column),
                u32::from(record.row),
            ),
            author: record.author,
            text: record.text,
        })
        .collect();
    notes.sort_by_key(|note| (note.row, note.col));
    notes
}

/// The features of `kind` on a worksheet, in position order.
pub(super) fn list(file: &Xls, sheet: &Sheet<'_>, kind: Kind) -> Result<Vec<Feature>> {
    list_handle(file, &sheet.handle, kind)
}

/// [`list`] over a worksheet handle.
pub(super) fn list_handle(
    file: &Xls,
    handle: &SheetHandle<'_>,
    kind: Kind,
) -> Result<Vec<Feature>> {
    if !handle.is_grid() {
        return Ok(Vec::new());
    }
    let features = file
        .owner()
        .sheet_features(handle.position(), selection(kind))?;
    Ok(project(features, kind))
}

/// The features of `kind` among `features`.
pub(super) fn project(features: sf::Features, kind: Kind) -> Vec<Feature> {
    let data: Vec<Data> = match kind {
        Kind::Comment => notes(features.notes)
            .into_iter()
            .map(Data::Comment)
            .collect(),
        Kind::Picture | Kind::Shape => {
            let wanted = if kind == Kind::Picture {
                DrawingKind::Picture
            } else {
                DrawingKind::Shape
            };
            features
                .drawings
                .into_iter()
                .filter(|drawing| drawing.kind == wanted)
                .map(|drawing| Data::Drawing(Box::new(drawing)))
                .collect()
        },
        Kind::Hyperlink => features
            .hyperlinks
            .into_iter()
            .map(|link| Data::Hyperlink(Box::new(link)))
            .collect(),
        Kind::Cf => features
            .formats
            .into_iter()
            .map(|format| Data::Cf(Box::new(format)))
            .collect(),
        Kind::DataValidation => features
            .validations
            .into_iter()
            .map(|validation| Data::Validation(Box::new(validation)))
            .collect(),
        Kind::Autofilter => features
            .auto_filter
            .map(Data::Autofilter)
            .into_iter()
            .collect(),
        Kind::RowBreak => features.row_breaks.into_iter().map(Data::Break).collect(),
        Kind::ColBreak => features
            .column_breaks
            .into_iter()
            .map(Data::Break)
            .collect(),
        _ => Vec::new(),
    };
    data.into_iter()
        .enumerate()
        .map(|(index, data)| Feature {
            kind,
            index: u32::try_from(index).unwrap_or(u32::MAX),
            data,
        })
        .collect()
}

/// Resolve one feature step under a worksheet.
pub(super) fn lookup<'f>(
    file: &'f Xls,
    sheet: &Sheet<'f>,
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
        Some(index) => Lookup::Found(Node::new_feature(
            file,
            sheet.position,
            features.swap_remove(index),
        )),
        None => Lookup::Missing,
    })
}

impl<'f> Node<'f> {
    /// A worksheet feature.
    pub(super) fn new_feature(file: &'f Xls, sheet: usize, feature: Feature) -> Self {
        Self::in_sheet_item(file, sheet, Item::Feature(Arc::new(feature)))
    }
}

// ---------------------------------------------------------------------------
// Property tables

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
const fn xls_gap(
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

const fn bar_gap(key: &'static str) -> GapSpec {
    xls_gap(
        "conditionalFormatting",
        key,
        "record 0327 follow-up: XLS data-bar extensions",
        "a BIFF8 CF12 data bar has a color, thresholds and lengths; negative and axis styling exist only in Office 2010 formats",
    )
}

static NEGATIVE_COLOR_GAP: GapSpec = bar_gap("negativeColor");
static AXIS_COLOR_GAP: GapSpec = bar_gap("axisColor");
static AXIS_POSITION_GAP: GapSpec = bar_gap("axisPosition");
static DIRECTION_GAP: GapSpec = bar_gap("direction");
static CRITERIA_GAP: GapSpec = xls_gap(
    "autofilter",
    "criteria0",
    "record 0327 follow-up: XLS auto-filter column criteria",
    "XLS auto-filter column criteria (AutoFilter/AutoFilter12 records) are kept but not authored; the filter range is",
);

const fn effect_gap(element: &'static str, key: &'static str) -> GapSpec {
    xls_gap(
        element,
        key,
        "record 0327 follow-up: XLS drawing effects, crops and gradient fills",
        "the XLS drawing editor authors solid fills, outlines, text, names and transforms; effects, crops, transparency and gradient fills are not authored",
    )
}

const fn text_gap(key: &'static str) -> GapSpec {
    xls_gap(
        "shape",
        key,
        "record 0327 follow-up: XLS shape text formatting",
        "a legacy shape's text runs name entries of the workbook font table, which the drawing editor does not author; text is written in the default font",
    )
}

static PICTURE_TITLE_GAP: GapSpec = xls_gap(
    "picture",
    "title",
    "record 0327: wontfix (format)",
    "a legacy OfficeArt picture has a name and alternative text but no title",
);
static PICTURE_CROP_GAP: GapSpec = effect_gap("picture", "crop");
static PICTURE_OPACITY_GAP: GapSpec = effect_gap("picture", "opacity");
static PICTURE_SHADOW_GAP: GapSpec = effect_gap("picture", "shadow");
static PICTURE_GLOW_GAP: GapSpec = effect_gap("picture", "glow");
static PICTURE_REFLECTION_GAP: GapSpec = effect_gap("picture", "reflection");
static PICTURE_SOFT_EDGE_GAP: GapSpec = effect_gap("picture", "softEdge");
static SHAPE_GRADIENT_GAP: GapSpec = effect_gap("shape", "gradientFill");
static SHAPE_SHADOW_GAP: GapSpec = effect_gap("shape", "shadow");
static SHAPE_GLOW_GAP: GapSpec = effect_gap("shape", "glow");
static SHAPE_REFLECTION_GAP: GapSpec = effect_gap("shape", "reflection");
static SHAPE_SOFT_EDGE_GAP: GapSpec = effect_gap("shape", "softEdge");
static SHAPE_FONT_GAP: GapSpec = text_gap("font");
static SHAPE_SIZE_GAP: GapSpec = text_gap("size");
static SHAPE_BOLD_GAP: GapSpec = text_gap("bold");
static SHAPE_ITALIC_GAP: GapSpec = text_gap("italic");
static SHAPE_UNDERLINE_GAP: GapSpec = text_gap("underline");
static SHAPE_COLOR_GAP: GapSpec = text_gap("color");
static SHAPE_ALIGN_GAP: GapSpec = text_gap("align");
static SHAPE_VALIGN_GAP: GapSpec = text_gap("valign");
static SHAPE_MARGIN_GAP: GapSpec = text_gap("margin");

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

/// Note keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CommentKey {
    Ref,
    Author,
    Text,
    AnchoredTo,
}

keyed!(CommentKey, CommentKey, COMMENT_TABLE, COMMENT_PROPS, COMMENT_KEYS, [
    entry(read("ref", TEXT), Access::SetAndAdd(CommentKey::Ref)) => Ref,
    entry(read("author", TEXT), Access::SetAndAdd(CommentKey::Author)) => Author,
    entry(read("text", TEXT), Access::SetAndAdd(CommentKey::Text)) => Text,
    entry(read("anchoredTo", TEXT), Access::ReadOnly) => AnchoredTo,
], []);

fn comment_value(note: &Note, key: CommentKey) -> Option<Value<'_>> {
    match key {
        CommentKey::Ref => Some(text(note.reference.as_str())),
        CommentKey::Author => Some(text(&note.author)),
        CommentKey::Text => Some(text(&note.text)),
        // Read through `Feature::value`, which knows the sheet's name.
        CommentKey::AnchoredTo => None,
    }
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

fn hyperlink_value(
    link: &litchi_xls::hyperlinks::Hyperlink,
    key: HyperlinkKey,
) -> Option<Value<'_>> {
    match key {
        HyperlinkKey::Ref => Some(owned(link_reference(link))),
        HyperlinkKey::Url => link_url(link).map(text),
        HyperlinkKey::Location => link.location().map(text),
        HyperlinkKey::Display => link.display_name().map(text),
        HyperlinkKey::Tooltip => link.tooltip().map(text),
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
    Rotation,
    Flip,
    ContentType,
    FileSize,
    Src,
    Anchor,
    FlipH,
    FlipV,
    FlipBoth,
    Pending,
}

keyed!(PictureKey, PictureKey, PICTURE_TABLE, PICTURE_PROPS, PICTURE_KEYS, [
    entry(read("name", TEXT), Access::SetAndAdd(PictureKey::Name)) => Name,
    clearable(read("alt", TEXT), Access::SetAndAdd(PictureKey::Alt)) => Alt,
    entry(write_only("title", TEXT), Access::Gap(&PICTURE_TITLE_GAP)) => Title,
    entry(read("decorative", BOOL), Access::ReadOnly) => Decorative,
    entry(read("hyperlink", TEXT), Access::ReadOnly) => Hyperlink,
    entry(read("x", SCALAR), Access::SetAndAdd(PictureKey::X)) => X,
    entry(read("y", SCALAR), Access::SetAndAdd(PictureKey::Y)) => Y,
    entry(read("width", SCALAR), Access::SetAndAdd(PictureKey::Width)) => Width,
    entry(read("height", SCALAR), Access::SetAndAdd(PictureKey::Height)) => Height,
    entry(read("anchorMode", TEXT), Access::Add(PictureKey::AnchorMode)) => AnchorMode,
    entry(read("rotation", NUMBER), Access::SetAndAdd(PictureKey::Rotation)) => Rotation,
    entry(read("flip", TEXT), Access::SetAndAdd(PictureKey::Flip)) => Flip,
    entry(read("contentType", TEXT), Access::ReadOnly) => ContentType,
    entry(read("fileSize", INT), Access::ReadOnly) => FileSize,
    entry(write_only("src", TEXT), Access::Add(PictureKey::Src)) => Src,
    entry(write_only("anchor", TEXT), Access::Add(PictureKey::Anchor)) => Anchor,
    entry(write_only("flipH", BOOL), Access::SetAndAdd(PictureKey::FlipH)) => FlipH,
    entry(write_only("flipV", BOOL), Access::SetAndAdd(PictureKey::FlipV)) => FlipV,
    entry(write_only("flipBoth", BOOL), Access::SetAndAdd(PictureKey::FlipBoth)) => FlipBoth,
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
    Font,
    Size,
    Bold,
    Italic,
    Underline,
    Color,
    Align,
    Valign,
    Rotation,
    Flip,
    Hyperlink,
    Alt,
    Anchor,
    FlipH,
    FlipV,
    FlipBoth,
    Margin,
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
    entry(write_only("font", TEXT), Access::Gap(&SHAPE_FONT_GAP)) => Font,
    entry(write_only("size", FONT_SIZE), Access::Gap(&SHAPE_SIZE_GAP)) => Size,
    entry(write_only("bold", BOOL), Access::Gap(&SHAPE_BOLD_GAP)) => Bold,
    entry(write_only("italic", BOOL), Access::Gap(&SHAPE_ITALIC_GAP)) => Italic,
    entry(write_only("underline", TEXT), Access::Gap(&SHAPE_UNDERLINE_GAP)) => Underline,
    entry(write_only("color", DRAWING_COLOR), Access::Gap(&SHAPE_COLOR_GAP)) => Color,
    entry(write_only("align", TEXT), Access::Gap(&SHAPE_ALIGN_GAP)) => Align,
    entry(write_only("valign", TEXT), Access::Gap(&SHAPE_VALIGN_GAP)) => Valign,
    entry(read("rotation", NUMBER), Access::SetAndAdd(ShapeKey::Rotation)) => Rotation,
    entry(read("flip", TEXT), Access::SetAndAdd(ShapeKey::Flip)) => Flip,
    entry(read("hyperlink", TEXT), Access::ReadOnly) => Hyperlink,
    clearable(read("alt", TEXT), Access::SetAndAdd(ShapeKey::Alt)) => Alt,
    entry(write_only("anchor", TEXT), Access::Add(ShapeKey::Anchor)) => Anchor,
    entry(write_only("flipH", BOOL), Access::SetAndAdd(ShapeKey::FlipH)) => FlipH,
    entry(write_only("flipV", BOOL), Access::SetAndAdd(ShapeKey::FlipV)) => FlipV,
    entry(write_only("flipBoth", BOOL), Access::SetAndAdd(ShapeKey::FlipBoth)) => FlipBoth,
    entry(write_only("margin", TEXT), Access::Gap(&SHAPE_MARGIN_GAP)) => Margin,
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

fn flip(h: bool, v: bool) -> Option<Value<'static>> {
    match (h, v) {
        (true, true) => Some(Value::Token("both")),
        (true, false) => Some(Value::Token("h")),
        (false, true) => Some(Value::Token("v")),
        (false, false) => None,
    }
}

fn rotation(drawing: &sf::Drawing) -> Option<Value<'static>> {
    (drawing.rotation != 0.0)
        .then(|| Value::number(drawing.rotation))
        .flatten()
}

/// OfficeCLI's cell geometry of an anchor: its origin and its span in
/// cells.
fn geometry(drawing: &sf::Drawing, axis: usize) -> Option<Value<'static>> {
    let anchor = drawing.anchor?;
    let (from, to) = (anchor.top_left(), anchor.bottom_right());
    Some(int(match axis {
        0 => from.column(),
        1 => from.row(),
        2 => to.column().saturating_sub(from.column()),
        _ => to.row().saturating_sub(from.row()),
    }))
}

fn anchor_mode(drawing: &sf::Drawing) -> Option<Value<'static>> {
    match drawing.anchor?.behavior() {
        litchi_xls::AnchorBehavior::MoveWithoutResize => Some(Value::Token("oneCell")),
        litchi_xls::AnchorBehavior::Absolute => Some(Value::Token("absolute")),
        _ => None,
    }
}

/// The preset name of an OfficeArt shape type.
pub(super) const fn geometry_name(shape_type: u16) -> Option<&'static str> {
    Some(match shape_type {
        1 | 202 => "rect",
        2 => "roundRect",
        3 => "ellipse",
        4 => "diamond",
        5 => "triangle",
        6 => "rtTriangle",
        7 => "parallelogram",
        8 => "trapezoid",
        9 => "hexagon",
        10 => "octagon",
        12 => "star5",
        13 => "rightArrow",
        20 | 32 => "line",
        _ => return None,
    })
}

fn picture_value(drawing: &sf::Drawing, key: PictureKey) -> Option<Value<'_>> {
    match key {
        PictureKey::Name => drawing.name.as_deref().map(text),
        PictureKey::Alt => drawing.description.as_deref().map(text),
        PictureKey::X => geometry(drawing, 0),
        PictureKey::Y => geometry(drawing, 1),
        PictureKey::Width => geometry(drawing, 2),
        PictureKey::Height => geometry(drawing, 3),
        PictureKey::AnchorMode => anchor_mode(drawing),
        PictureKey::Rotation => rotation(drawing),
        PictureKey::Flip => flip(drawing.flip_h, drawing.flip_v),
        PictureKey::ContentType => drawing.image.map(|(content, _)| Value::Token(content)),
        PictureKey::FileSize => drawing.image.map(|(_, size)| int(size)),
        PictureKey::Title
        | PictureKey::Decorative
        | PictureKey::Hyperlink
        | PictureKey::Src
        | PictureKey::Anchor
        | PictureKey::FlipH
        | PictureKey::FlipV
        | PictureKey::FlipBoth
        | PictureKey::Pending => None,
    }
}

fn shape_value(drawing: &sf::Drawing, key: ShapeKey) -> Option<Value<'_>> {
    match key {
        ShapeKey::Name => drawing.name.as_deref().map(text),
        ShapeKey::Text => drawing.text.as_deref().map(text),
        ShapeKey::Geometry => geometry_name(drawing.shape_type).map(Value::Token),
        ShapeKey::X => geometry(drawing, 0),
        ShapeKey::Y => geometry(drawing, 1),
        ShapeKey::Width => geometry(drawing, 2),
        ShapeKey::Height => geometry(drawing, 3),
        ShapeKey::Fill => match drawing.fill? {
            Fill::None => Some(Value::Token("none")),
            Fill::Solid(color) => Some(owned(format!("#{:06X}", color & 0x00FF_FFFF))),
        },
        ShapeKey::Line => match drawing.line? {
            Line::None => Some(Value::Token("none")),
            Line::Solid { color, width } => {
                #[allow(clippy::cast_precision_loss, reason = "line widths are small")]
                let points = f64::from(width) / 12_700.0;
                Some(owned(format!(
                    "#{:06X}:{}",
                    color & 0x00FF_FFFF,
                    crate::value::DotNetDouble(points)
                )))
            },
        },
        ShapeKey::Rotation => rotation(drawing),
        ShapeKey::Flip => flip(drawing.flip_h, drawing.flip_v),
        ShapeKey::Alt => drawing.description.as_deref().map(text),
        ShapeKey::Font
        | ShapeKey::Size
        | ShapeKey::Bold
        | ShapeKey::Italic
        | ShapeKey::Underline
        | ShapeKey::Color
        | ShapeKey::Align
        | ShapeKey::Valign
        | ShapeKey::Hyperlink
        | ShapeKey::Margin
        | ShapeKey::Anchor
        | ShapeKey::FlipH
        | ShapeKey::FlipV
        | ShapeKey::FlipBoth
        | ShapeKey::Pending => None,
    }
}

/// Conditional-format keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CfKey {
    Type,
    Ref,
    Priority,
    StopIfTrue,
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
    StdDev,
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
    Reverse,
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
    Direction,
}

keyed!(CfKey, CfKey, CF_TABLE, CF_PROPS, CF_KEYS, [
    entry(read("type", TEXT), Access::Add(CfKey::Type)) => Type,
    entry(read("ref", TEXT), Access::SetAndAdd(CfKey::Ref)) => Ref,
    entry(read("priority", INT), Access::ReadOnly) => Priority,
    entry(read("stopIfTrue", BOOL), Access::SetAndAdd(CfKey::StopIfTrue)) => StopIfTrue,
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
    entry(read("stdDev", INT), Access::SetAndAdd(CfKey::StdDev)) => StdDev,
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
    entry(read("reverse", BOOL), Access::SetAndAdd(CfKey::Reverse)) => Reverse,
    entry(read("fill", SHEET_COLOR), Access::SetAndAdd(CfKey::Fill)) => Fill,
    entry(read("font.color", SHEET_COLOR), Access::SetAndAdd(CfKey::FontColor)) => FontColor,
    entry(read("font.bold", BOOL), Access::SetAndAdd(CfKey::FontBold)) => FontBold,
    entry(read("font.italic", BOOL), Access::SetAndAdd(CfKey::FontItalic)) => FontItalic,
    entry(read("font.strike", BOOL), Access::SetAndAdd(CfKey::FontStrike)) => FontStrike,
    entry(read("font.underline", TEXT), Access::SetAndAdd(CfKey::FontUnderline)) => FontUnderline,
    entry(read("font.size", FONT_SIZE), Access::SetAndAdd(CfKey::FontSize)) => FontSize,
    entry(read("font.name", TEXT), Access::SetAndAdd(CfKey::FontName)) => FontName,
    entry(write_only("negativeColor", SHEET_COLOR), Access::Gap(&NEGATIVE_COLOR_GAP)) => NegativeColor,
    entry(write_only("axisColor", SHEET_COLOR), Access::Gap(&AXIS_COLOR_GAP)) => AxisColor,
    entry(write_only("axisPosition", TEXT), Access::Gap(&AXIS_POSITION_GAP)) => AxisPosition,
    entry(write_only("direction", TEXT), Access::Gap(&DIRECTION_GAP)) => Direction,
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

/// The OfficeCLI name of an icon set (`iIconSet`, [MS-XLS] 2.5.36).
pub(super) const ICON_SETS: [&str; 17] = [
    "3Arrows",
    "3ArrowsGray",
    "3Flags",
    "3TrafficLights1",
    "3TrafficLights2",
    "3Signs",
    "3Symbols",
    "3Symbols2",
    "4Arrows",
    "4ArrowsGray",
    "4RedToBlack",
    "4Rating",
    "4TrafficLights",
    "5Arrows",
    "5ArrowsGray",
    "5Rating",
    "5Quarters",
];

/// The `CFOper` names of a cellIs rule, by wire value 1–8.
pub(super) const CF_OPERATORS: [&str; 8] = [
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual",
];

/// The `CFTextOper` names of a text rule, by wire value 0–3.
pub(super) const TEXT_OPERATORS: [&str; 4] =
    ["containsText", "notContains", "beginsWith", "endsWith"];

/// The time periods of the `CFTemp` date templates 15–24, in template
/// order, with their `CFDateOper` parameter.
pub(super) const PERIODS: [(&str, u32); 10] = [
    ("today", 0),
    ("tomorrow", 6),
    ("yesterday", 1),
    ("last7Days", 2),
    ("lastMonth", 5),
    ("nextMonth", 8),
    ("thisWeek", 3),
    ("nextWeek", 7),
    ("lastWeek", 4),
    ("thisMonth", 9),
];

/// A threshold's value: a number, percent, percentile or formula.
fn threshold(value: &Threshold) -> Option<Value<'static>> {
    matches!(value.kind, 1 | 4 | 5 | 7)
        .then(|| value.value.clone().map(owned))
        .flatten()
}

/// OfficeCLI's `type` of a conditional format's first rule.
pub(super) fn rule_type(format: &ConditionalFormat) -> &'static str {
    let Some(rule) = format.rule.as_ref() else {
        return "";
    };
    match rule.rule_type {
        RuleType::CellIs => "cellIs",
        RuleType::ColorScale => "colorScale",
        RuleType::DataBar => "dataBar",
        RuleType::IconSet => "iconSet",
        RuleType::TopN => "topN",
        RuleType::Expression => match rule.template {
            7 => "uniqueValues",
            8 => match rule.text_operator {
                1 => "notContainsText",
                2 => "beginsWith",
                3 => "endsWith",
                _ => "containsText",
            },
            9 => "containsBlanks",
            10 => "notContainsBlanks",
            11 => "containsErrors",
            12 => "notContainsErrors",
            15..=24 => "timePeriod",
            25 | 26 | 29 | 30 => "aboveAverage",
            27 => "duplicateValues",
            _ => "formula",
        },
    }
}

fn cf_value(cf: &ConditionalFormat, key: CfKey) -> Option<Value<'_>> {
    let rule = cf.rule.as_ref()?;
    let style = &rule.style;
    let flag = |value: bool| value.then_some(Value::Bool(true));
    let average =
        rule.rule_type == RuleType::Expression && matches!(rule.template, 25 | 26 | 29 | 30);
    match key {
        CfKey::Type => Some(Value::Token(rule_type(cf))),
        CfKey::Ref => Some(owned(ranges(&cf.ranges))),
        CfKey::Priority => Some(int(rule.priority)),
        CfKey::StopIfTrue => flag(rule.stop_if_true),
        CfKey::DxfId => None,
        CfKey::Operator => match rule.rule_type {
            RuleType::CellIs => rule
                .operator
                .and_then(|operator| CF_OPERATORS.get(usize::from(operator).checked_sub(1)?))
                .map(|name| Value::Token(name)),
            RuleType::Expression if rule.template == 8 => TEXT_OPERATORS
                .get(usize::from(rule.text_operator))
                .map(|name| Value::Token(name)),
            _ => None,
        },
        CfKey::Value => (rule.rule_type == RuleType::CellIs)
            .then(|| rule.formulas.first().map(|value| text(value)))
            .flatten(),
        CfKey::Value2 => (rule.rule_type == RuleType::CellIs)
            .then(|| rule.formulas.get(1).map(|value| text(value)))
            .flatten(),
        CfKey::Formula => (rule.rule_type == RuleType::Expression && rule.template == 1)
            .then(|| rule.formulas.first().map(|value| text(value)))
            .flatten(),
        CfKey::Text => rule.text.as_deref().map(text),
        CfKey::Rank => (rule.rule_type == RuleType::TopN).then(|| int(rule.parameter)),
        CfKey::Percent => flag(rule.percent),
        CfKey::Bottom => flag(rule.rule_type == RuleType::TopN && rule.bottom),
        CfKey::AboveAverage => {
            (average && matches!(rule.template, 26 | 30)).then_some(Value::Bool(false))
        },
        CfKey::EqualAverage => flag(average && matches!(rule.template, 29 | 30)),
        CfKey::StdDev => (average && rule.parameter > 0).then(|| int(rule.parameter)),
        CfKey::Period => (rule.rule_type == RuleType::Expression)
            .then(|| {
                rule.template
                    .checked_sub(15)
                    .and_then(|index| PERIODS.get(usize::from(index)))
                    .map(|(name, _)| Value::Token(name))
            })
            .flatten(),
        CfKey::Color => rule.data_bar.as_ref().and_then(|bar| bar.color).map(rgb),
        CfKey::Min => rule.data_bar.as_ref().and_then(|bar| threshold(&bar.min)),
        CfKey::Max => rule.data_bar.as_ref().and_then(|bar| threshold(&bar.max)),
        CfKey::MinLength => rule.data_bar.as_ref().map(|bar| int(bar.min_length)),
        CfKey::MaxLength => rule.data_bar.as_ref().map(|bar| int(bar.max_length)),
        CfKey::ShowValue => match (&rule.data_bar, &rule.icon_set) {
            (Some(bar), _) => (!bar.show_value).then_some(Value::Bool(false)),
            (None, Some(icons)) => (!icons.show_value).then_some(Value::Bool(false)),
            (None, None) => None,
        },
        CfKey::MinColor => rule
            .color_scale
            .as_ref()
            .and_then(|scale| scale.points.first())
            .and_then(|point| point.1)
            .map(rgb),
        CfKey::MidColor => rule
            .color_scale
            .as_ref()
            .filter(|scale| scale.points.len() == 3)
            .and_then(|scale| scale.points[1].1)
            .map(rgb),
        CfKey::MaxColor => rule
            .color_scale
            .as_ref()
            .and_then(|scale| scale.points.last())
            .and_then(|point| point.1)
            .map(rgb),
        CfKey::Midpoint => rule
            .color_scale
            .as_ref()
            .filter(|scale| scale.points.len() == 3)
            .and_then(|scale| threshold(&scale.points[1].0)),
        CfKey::Iconset => rule
            .icon_set
            .as_ref()
            .and_then(|icons| ICON_SETS.get(usize::from(icons.icon_set)))
            .map(|name| Value::Token(name)),
        CfKey::Reverse => rule.icon_set.as_ref().and_then(|icons| flag(icons.reverse)),
        CfKey::Fill => style.fill.map(rgb),
        CfKey::FontColor => style.font_color.map(rgb),
        CfKey::FontBold => style.bold.map(Value::Bool),
        CfKey::FontItalic => style.italic.map(Value::Bool),
        CfKey::FontStrike => style.strike.map(Value::Bool),
        CfKey::FontUnderline => style.underline.map(|underline| {
            Value::Token(if matches!(underline, 2 | 0x22) {
                "double"
            } else {
                "single"
            })
        }),
        CfKey::FontSize => style.size_twips.map(points),
        CfKey::FontName => style.font_name.as_deref().map(text),
        CfKey::NegativeColor | CfKey::AxisColor | CfKey::AxisPosition | CfKey::Direction => None,
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

/// The OfficeCLI names of the validation types, by wire value 0–7.
pub(super) const VALIDATION_TYPES: [&str; 8] = [
    "none",
    "whole",
    "decimal",
    "list",
    "date",
    "time",
    "textLength",
    "custom",
];

/// The OfficeCLI names of the validation operators, by wire value 0–7.
pub(super) const VALIDATION_OPERATORS: [&str; 8] = [
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual",
];

/// The OfficeCLI names of the error styles, by wire value 0–2.
pub(super) const ERROR_STYLES: [&str; 3] = ["stop", "warning", "information"];

fn validation_value(validation: &Validation, key: ValidationKey) -> Option<Value<'_>> {
    let kind = validation.kind;
    let typed = kind != 0;
    match key {
        ValidationKey::Ref => Some(owned(ranges(&validation.ranges))),
        ValidationKey::Type => VALIDATION_TYPES
            .get(usize::from(kind))
            .map(|name| Value::Token(name)),
        ValidationKey::Operator => (typed && !matches!(kind, 3 | 7))
            .then(|| VALIDATION_OPERATORS.get(usize::from(validation.operator)))
            .flatten()
            .map(|name| Value::Token(name)),
        ValidationKey::Formula1 => validation.formula1.as_deref().map(text),
        ValidationKey::Formula2 => validation.formula2.as_deref().map(text),
        ValidationKey::AllowBlank => Some(Value::Bool(validation.allow_blank)),
        ValidationKey::ShowError => Some(Value::Bool(validation.show_error)),
        ValidationKey::ShowInput => Some(Value::Bool(validation.show_input)),
        ValidationKey::ErrorTitle => validation.error_title.as_deref().map(text),
        ValidationKey::Error => validation.error.as_deref().map(text),
        ValidationKey::PromptTitle => validation.prompt_title.as_deref().map(text),
        ValidationKey::Prompt => validation.prompt.as_deref().map(text),
        ValidationKey::ErrorStyle => ERROR_STYLES
            .get(usize::from(validation.error_style))
            .map(|name| Value::Token(name)),
        // `fSuppressCombo` hides the in-cell arrow.
        ValidationKey::InCellDropdown => Some(Value::Bool(!validation.suppress_dropdown)),
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
    entry(read("min", INT), Access::SetAndAdd(BreakKey::Min)) => Min,
    entry(read("max", INT), Access::SetAndAdd(BreakKey::Max)) => Max,
], [alias("index", 0)]);

pub(crate) const COL_BREAK_TABLE: Table<BreakKey, 4, 2> = Table::new(
    [
        entry(read("col", SCALAR), Access::SetAndAdd(BreakKey::Index)),
        entry(read("manual", BOOL), Access::SetAndAdd(BreakKey::Manual)),
        entry(read("min", INT), Access::SetAndAdd(BreakKey::Min)),
        entry(read("max", INT), Access::SetAndAdd(BreakKey::Max)),
    ],
    [alias("column", 0), alias("index", 0)],
);
pub(crate) const COL_BREAK_PROPS: &[PropSpec] = &COL_BREAK_TABLE.readable();

fn break_value(kind: Kind, item: PageBreak, key: BreakKey) -> Option<Value<'static>> {
    let full = if kind == Kind::RowBreak { 255 } else { 65_535 };
    let partial = item.min != 0 || item.max != full;
    match key {
        BreakKey::Index => Some(int(item.index)),
        BreakKey::Manual => Some(Value::Bool(true)),
        BreakKey::Min => partial.then(|| int(item.min)),
        BreakKey::Max => partial.then(|| int(item.max)),
    }
}
