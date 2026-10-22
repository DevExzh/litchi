//! XLSB worksheet features as unified nodes (record 0322): legacy notes
//! (`comment`), hyperlinks, drawing pictures and shapes, conditional
//! formats (`cf`), data validations, the auto-filter and manual page
//! breaks, with the XLSX property vocabulary of record 0320.
//!
//! Every feature is a child of its worksheet, addressed by its position
//! among the worksheet's features of its kind (`/Sheet1/comment[2]`,
//! `/Sheet1/cf[1]`; the auto-filter is `/Sheet1/autofilter`). A node holds
//! one [`Feature`] behind one `Arc`. Listing a kind reads only the parts it
//! lives in, and never decodes the worksheet's cells: the worksheet stream's
//! records after the cell table for conditional formats, validations, the
//! auto-filter, hyperlinks and page breaks; the comments part for notes; the
//! drawing part for pictures and shapes; and the styles part's differential
//! formats for a conditional format's formatting.
//!
//! Positions follow the stream order, except notes, which are numbered in
//! row-major order of their cells, as XLSX numbers them. Conditional formats
//! and validations are the classic records; Office 2013 formats and
//! `BrtDVal14` validations are kept but not listed. The write verbs live in
//! `feature_write`.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};
use litchi_xlsb as xlsb;
use xlsb::conditional_formatting::Formatting;
use xlsb::conditional_formatting::edit::{DifferentialStyle, Underline};
use xlsb::drawing::edit::{Align, Element, ElementKind, Fill, Line, Placement, VerticalAlign};

use super::node::{Item, Node, Sheet, SheetHandle};
use super::spec::Kind;
use super::{Xlsb, XlsbOwner};
use crate::error::{Code, Error, Result};
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
    Hyperlink(Box<xlsb::hyperlinks::Hyperlink>),
    Drawing(Box<Element>),
    Cf(Box<CfData>),
    Validation(Box<xlsb::package::data_validation::Validation>),
    Autofilter(xlsb::auto_filter::AutoFilter),
    Break(xlsb::page_breaks::Break),
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

/// A conditional format and the formatting its first rule applies.
#[derive(Clone)]
pub(crate) struct CfData {
    pub(super) formatting: Formatting,
    pub(super) dxf: Option<DifferentialStyle>,
}

/// The A1 label of a zero-based cell.
pub(super) fn a1(row: u32, col: u32) -> String {
    crate::format::workbook::Label::cell(col, row)
        .as_str()
        .to_owned()
}

/// The A1 reference of a hyperlink's cell or range.
pub(super) fn link_reference(link: &xlsb::hyperlinks::Hyperlink) -> String {
    crate::format::workbook::Label::range(
        (link.col_first, link.row_first),
        (link.col_last, link.row_last),
    )
    .as_str()
    .to_owned()
}

/// The A1 range of an auto-filter.
pub(super) fn filter_reference(filter: xlsb::auto_filter::AutoFilter) -> String {
    crate::format::workbook::Label::range(
        (filter.first_column, filter.first_row),
        (filter.last_column, filter.last_row),
    )
    .as_str()
    .to_owned()
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
                link.display
                    .as_deref()
                    .or(link.target.as_deref())
                    .or(link.location.as_deref())
                    .unwrap_or(""),
            ),
            Data::Drawing(element) => match element.kind {
                ElementKind::Shape => Cow::Borrowed(element.text.as_deref()?),
                _ => Cow::Borrowed(element.description.as_deref()?),
            },
            Data::Cf(cf) => Cow::Owned(ranges(&cf.formatting)),
            Data::Validation(validation) => Cow::Borrowed(validation.cell_ranges.as_str()),
            Data::Autofilter(filter) => Cow::Owned(filter_reference(*filter)),
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
            Data::Drawing(element) => Cow::Borrowed(element.name.as_deref()?),
            Data::Cf(cf) => Cow::Owned(format!(
                "{} {}",
                rule_type(&cf.formatting),
                ranges(&cf.formatting)
            )),
            Data::Validation(validation) => Cow::Borrowed(validation.cell_ranges.as_str()),
            Data::Autofilter(filter) => Cow::Owned(filter_reference(*filter)),
            Data::Break(item) => Cow::Owned(item.id.to_string()),
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
            Data::Drawing(element) => match self.kind {
                Kind::Picture => picture_value(element, *PICTURE_KEYS.get(index)?),
                _ => shape_value(element, *SHAPE_KEYS.get(index)?),
            },
            Data::Cf(cf) => cf_value(cf, *CF_KEYS.get(index)?),
            Data::Validation(validation) => {
                validation_value(validation, *VALIDATION_KEYS.get(index)?)
            },
            Data::Autofilter(filter) => match *AUTOFILTER_KEYS.get(index)? {
                AutofilterKey::Range => Some(owned(filter_reference(*filter))),
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

fn points(points: f64) -> Value<'static> {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "font sizes are a few thousand points at most"
    )]
    Value::Length(Length::from_emu((points * 12_700.0).round() as i64))
}

pub(super) fn ranges(formatting: &Formatting) -> String {
    formatting.ranges.join(" ")
}

// ---------------------------------------------------------------------------
// Listing

impl SheetHandle {
    fn eager(&self) -> Result<&xlsb::Workbook> {
        match self {
            Self::Eager { file, .. } => match &file.0.owner {
                XlsbOwner::Eager(package) => super::projection(package),
                XlsbOwner::Source(_) => Err(lost_owner()),
            },
            Self::Source { .. } => Err(lost_owner()),
        }
    }

    pub(super) fn features(
        &self,
        select: xlsb::sheet::FeatureSet,
    ) -> Result<xlsb::sheet::Features> {
        Ok(match self {
            Self::Eager { .. } => self
                .eager()?
                .sheet_features(self.position().get(), select)?,
            Self::Source { sheet, .. } => sheet.features(select)?,
        })
    }

    fn comments(&self) -> Result<Vec<xlsb::comments::Record>> {
        Ok(match self {
            Self::Eager { .. } => self.eager()?.sheet_comments(self.position().get())?,
            Self::Source { sheet, .. } => sheet.comments()?,
        })
    }

    fn drawing_elements(&self) -> Result<Vec<Element>> {
        Ok(match self {
            Self::Eager { .. } => self.eager()?.drawing_elements(self.position().get())?,
            Self::Source { sheet, .. } => sheet.drawing_elements()?,
        })
    }
}

fn lost_owner() -> Error {
    Error::new(
        Code::InternalError,
        "an XLSB worksheet handle lost its owner state",
    )
    .with_format(FileFormat::Xlsb)
}

/// The workbook's differential formats, read only when a conditional
/// format is listed.
fn differential_formats(file: &Xlsb) -> Result<Vec<DifferentialStyle>> {
    Ok(match &file.0.owner {
        XlsbOwner::Eager(package) => super::projection(package)?.differential_formats()?,
        XlsbOwner::Source(workbook) => workbook.workbook().differential_formats()?,
    })
}

/// The notes of a worksheet in row-major order of their cells.
pub(super) fn notes(records: Vec<xlsb::comments::Record>) -> Vec<Note> {
    let mut notes: Vec<Note> = records.into_iter().map(note).collect();
    notes.sort_by_key(|note| (note.row, note.col));
    notes
}

fn note(record: xlsb::comments::Record) -> Note {
    Note {
        row: record.row,
        col: record.col,
        reference: crate::format::workbook::Label::cell(record.col, record.row),
        author: record.author,
        text: record.text,
    }
}

/// The features of `kind` on a worksheet, in position order.
pub(super) fn list(file: &Xlsb, sheet: &Sheet, kind: Kind) -> Result<Vec<Feature>> {
    list_handle(file, &sheet.handle, kind)
}

/// [`list`] over a worksheet handle.
pub(super) fn list_handle(file: &Xlsb, handle: &SheetHandle, kind: Kind) -> Result<Vec<Feature>> {
    if !handle.is_grid() {
        return Ok(Vec::new());
    }
    let data: Vec<Data> = match kind {
        Kind::Comment => notes(handle.comments()?)
            .into_iter()
            .map(Data::Comment)
            .collect(),
        Kind::Picture | Kind::Shape => {
            let wanted = if kind == Kind::Picture {
                ElementKind::Picture
            } else {
                ElementKind::Shape
            };
            handle
                .drawing_elements()?
                .into_iter()
                .filter(|element| element.kind == wanted)
                .map(|element| Data::Drawing(Box::new(element)))
                .collect()
        },
        Kind::Hyperlink => handle
            .features(xlsb::sheet::FeatureSet::HYPERLINKS)?
            .hyperlinks
            .into_iter()
            .map(|link| Data::Hyperlink(Box::new(link)))
            .collect(),
        Kind::Cf => {
            let formats = handle.features(xlsb::sheet::FeatureSet::FORMATS)?.formats;
            let dxfs = if formats.is_empty() {
                Vec::new()
            } else {
                differential_formats(file)?
            };
            formats
                .into_iter()
                .map(|formatting| {
                    let dxf = formatting
                        .rules
                        .first()
                        .and_then(|rule| rule.dxf_id)
                        .and_then(|index| dxfs.get(index as usize).cloned());
                    Data::Cf(Box::new(CfData { formatting, dxf }))
                })
                .collect()
        },
        Kind::DataValidation => handle
            .features(xlsb::sheet::FeatureSet::VALIDATIONS)?
            .validations
            .into_iter()
            .map(|validation| Data::Validation(Box::new(validation)))
            .collect(),
        Kind::Autofilter => handle
            .features(xlsb::sheet::FeatureSet::AUTO_FILTER)?
            .auto_filter
            .map(Data::Autofilter)
            .into_iter()
            .collect(),
        Kind::RowBreak => handle
            .features(xlsb::sheet::FeatureSet::PAGE_BREAKS)?
            .page_breaks
            .rows
            .into_iter()
            .map(Data::Break)
            .collect(),
        Kind::ColBreak => handle
            .features(xlsb::sheet::FeatureSet::PAGE_BREAKS)?
            .page_breaks
            .columns
            .into_iter()
            .map(Data::Break)
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

/// Resolve one feature step under a worksheet.
pub(super) fn lookup<'f>(
    file: &'f Xlsb,
    sheet: &Sheet,
    kind: Kind,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let index = match select {
        Select::Ordinal(position) => Some(position.get()),
        Select::None => Some(0),
        Select::Last => None,
        _ => return Ok(Lookup::NotApplicable),
    };
    if kind == Kind::Comment {
        // One note is labelled, not all of them (record 0326); the stable
        // sort gives `notes`' row-major order.
        let mut records = if sheet.handle.is_grid() {
            sheet.handle.comments()?
        } else {
            Vec::new()
        };
        records.sort_by_key(|record| (record.row, record.col));
        let index = index.or_else(|| records.len().checked_sub(1));
        return Ok(match index.filter(|index| *index < records.len()) {
            Some(index) => Lookup::Found(Node::new_feature(
                file,
                sheet.position,
                Feature {
                    kind,
                    index: u32::try_from(index).unwrap_or(u32::MAX),
                    data: Data::Comment(note(records.swap_remove(index))),
                },
            )),
            None => Lookup::Missing,
        });
    }
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
    pub(super) fn new_feature(file: &'f Xlsb, sheet: usize, feature: Feature) -> Self {
        Self::in_sheet_item(file, sheet, Item::Feature(Arc::new(feature)))
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
const fn x14_gap(element: &'static str, key: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Property { element, key },
        code: Code::UnsupportedProperty,
        follow_up: "record 0322 follow-up: XLSB Office 2010 conditional-format extensions",
        reason: "the XLSB editor authors classic conditional-format records; data-bar negative and axis styling live in BrtBeginCFRule14 future records, which are not authored yet",
    }
}

static NEGATIVE_COLOR_GAP: GapSpec = x14_gap("conditionalFormatting", "negativeColor");
static AXIS_COLOR_GAP: GapSpec = x14_gap("conditionalFormatting", "axisColor");
static AXIS_POSITION_GAP: GapSpec = x14_gap("conditionalFormatting", "axisPosition");
static DIRECTION_GAP: GapSpec = x14_gap("conditionalFormatting", "direction");
static CRITERIA_GAP: GapSpec = GapSpec {
    subject: Subject::Property {
        element: "autofilter",
        key: "criteria0",
    },
    code: Code::UnsupportedProperty,
    follow_up: "record 0322 follow-up: XLSB auto-filter criteria",
    reason: "XLSB auto-filter column criteria (`criteriaN.op`, BrtBeginFilterColumn) are not authored yet; the filter range is",
};

const fn effect_gap(element: &'static str, key: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Property { element, key },
        code: Code::UnsupportedProperty,
        follow_up: "record 0322 follow-up: drawing effects, crops and gradient fills",
        reason: "the drawing writer authors solid fills, outlines, text and transforms; effects, crops, transparency and gradient fills are not authored yet",
    }
}

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

fn hyperlink_value(link: &xlsb::hyperlinks::Hyperlink, key: HyperlinkKey) -> Option<Value<'_>> {
    match key {
        HyperlinkKey::Ref => Some(owned(link_reference(link))),
        HyperlinkKey::Url => link.target.as_deref().map(text),
        HyperlinkKey::Location => link.location.as_deref().map(text),
        HyperlinkKey::Display => link.display.as_deref().map(text),
        HyperlinkKey::Tooltip => link.tooltip.as_deref().map(text),
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
    entry(read("title", TEXT), Access::Add(PictureKey::Title)) => Title,
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
    entry(read("font", TEXT), Access::Add(ShapeKey::Font)) => Font,
    entry(read("size", FONT_SIZE), Access::Add(ShapeKey::Size)) => Size,
    entry(read("bold", BOOL), Access::Add(ShapeKey::Bold)) => Bold,
    entry(read("italic", BOOL), Access::Add(ShapeKey::Italic)) => Italic,
    entry(read("underline", TEXT), Access::Add(ShapeKey::Underline)) => Underline,
    entry(read("color", DRAWING_COLOR), Access::Add(ShapeKey::Color)) => Color,
    entry(read("align", TEXT), Access::Add(ShapeKey::Align)) => Align,
    entry(read("valign", TEXT), Access::Add(ShapeKey::Valign)) => Valign,
    entry(read("rotation", NUMBER), Access::SetAndAdd(ShapeKey::Rotation)) => Rotation,
    entry(read("flip", TEXT), Access::SetAndAdd(ShapeKey::Flip)) => Flip,
    entry(read("hyperlink", TEXT), Access::ReadOnly) => Hyperlink,
    clearable(read("alt", TEXT), Access::SetAndAdd(ShapeKey::Alt)) => Alt,
    entry(write_only("anchor", TEXT), Access::Add(ShapeKey::Anchor)) => Anchor,
    entry(write_only("flipH", BOOL), Access::SetAndAdd(ShapeKey::FlipH)) => FlipH,
    entry(write_only("flipV", BOOL), Access::SetAndAdd(ShapeKey::FlipV)) => FlipV,
    entry(write_only("flipBoth", BOOL), Access::SetAndAdd(ShapeKey::FlipBoth)) => FlipBoth,
    entry(read("margin", TEXT), Access::Add(ShapeKey::Margin)) => Margin,
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

fn rotation(element: &Element) -> Option<Value<'static>> {
    element
        .rotation
        .filter(|rotation| *rotation != 0.0)
        .and_then(Value::number)
}

/// OfficeCLI's cell geometry of a placement: a two-cell anchor's origin and
/// span in cells; other anchors' position and extent in EMU.
fn geometry(element: &Element, axis: usize) -> Option<Value<'static>> {
    match element.placement? {
        Placement::TwoCell { from, to } => Some(int(match axis {
            0 => from.col,
            1 => from.row,
            2 => to.col.saturating_sub(from.col),
            _ => to.row.saturating_sub(from.row),
        })),
        Placement::OneCell { from, cx, cy } => Some(match axis {
            0 => int(from.col),
            1 => int(from.row),
            2 => owned(format!("{cx}emu")),
            _ => owned(format!("{cy}emu")),
        }),
        Placement::Absolute { x, y, cx, cy } => {
            Some(owned(format!("{}emu", [x, y, cx, cy][axis.min(3)])))
        },
        _ => None,
    }
}

fn picture_value(element: &Element, key: PictureKey) -> Option<Value<'_>> {
    match key {
        PictureKey::Name => element.name.as_deref().map(text),
        PictureKey::Alt => element.description.as_deref().map(text),
        PictureKey::Title => element.title.as_deref().map(text),
        PictureKey::Decorative => element.decorative.then_some(Value::Bool(true)),
        PictureKey::Hyperlink => element.hyperlink.as_deref().map(text),
        PictureKey::X => geometry(element, 0),
        PictureKey::Y => geometry(element, 1),
        PictureKey::Width => geometry(element, 2),
        PictureKey::Height => geometry(element, 3),
        PictureKey::AnchorMode => match element.placement? {
            Placement::OneCell { .. } => Some(Value::Token("oneCell")),
            Placement::Absolute { .. } => Some(Value::Token("absolute")),
            _ => None,
        },
        PictureKey::Rotation => rotation(element),
        PictureKey::Flip => flip(element.flip_h, element.flip_v),
        PictureKey::ContentType => element
            .image
            .as_ref()
            .map(|image| text(&image.content_type)),
        PictureKey::FileSize => element
            .image
            .as_ref()
            .map(|image| int(i64::try_from(image.size).unwrap_or(i64::MAX))),
        PictureKey::Src
        | PictureKey::Anchor
        | PictureKey::FlipH
        | PictureKey::FlipV
        | PictureKey::FlipBoth
        | PictureKey::Pending => None,
    }
}

fn shape_value(element: &Element, key: ShapeKey) -> Option<Value<'_>> {
    match key {
        ShapeKey::Name => element.name.as_deref().map(text),
        ShapeKey::Text => element.text.as_deref().map(text),
        ShapeKey::Geometry => element.geometry.as_deref().map(text),
        ShapeKey::X => geometry(element, 0),
        ShapeKey::Y => geometry(element, 1),
        ShapeKey::Width => geometry(element, 2),
        ShapeKey::Height => geometry(element, 3),
        ShapeKey::Fill => match element.fill? {
            Fill::None => Some(Value::Token("none")),
            Fill::Scheme(name) => Some(Value::Token(name)),
            Fill::Solid(color) => Some(owned(format!("#{:06X}", color & 0x00FF_FFFF))),
            _ => None,
        },
        ShapeKey::Line => match element.line? {
            Line::None => Some(Value::Token("none")),
            Line::Solid { color, width, .. } => {
                let mut value = format!("#{:06X}", color & 0x00FF_FFFF);
                if let Some(width) = width {
                    #[allow(clippy::cast_precision_loss, reason = "line widths are small")]
                    let points = width as f64 / 12_700.0;
                    value.push_str(&format!(":{}", crate::value::DotNetDouble(points)));
                }
                Some(owned(value))
            },
            _ => None,
        },
        ShapeKey::Font => element.font.name.as_deref().map(text),
        ShapeKey::Size => element.font.size.map(points),
        ShapeKey::Bold => element.font.bold.map(Value::Bool),
        ShapeKey::Italic => element.font.italic.map(Value::Bool),
        ShapeKey::Underline => element.font.underline.as_deref().map(|value| {
            Value::Token(match value {
                "dbl" => "double",
                "sng" => "single",
                _ => "single",
            })
        }),
        ShapeKey::Color => element.font.color.map(rgb),
        ShapeKey::Align => element.align.map(|align| {
            Value::Token(match align {
                Align::Left => "left",
                Align::Center => "center",
                Align::Right => "right",
                _ => "justify",
            })
        }),
        ShapeKey::Valign => element.valign.map(|valign| {
            Value::Token(match valign {
                VerticalAlign::Top => "top",
                VerticalAlign::Middle => "middle",
                _ => "bottom",
            })
        }),
        ShapeKey::Rotation => rotation(element),
        ShapeKey::Flip => flip(element.flip_h, element.flip_v),
        ShapeKey::Hyperlink => element.hyperlink.as_deref().map(text),
        ShapeKey::Alt => element.description.as_deref().map(text),
        ShapeKey::Margin => element.margin.map(|margin| {
            let points = |emu: i64| {
                #[allow(clippy::cast_precision_loss, reason = "insets are small")]
                let points = emu as f64 / 12_700.0;
                crate::value::DotNetDouble(points).to_string()
            };
            if margin.iter().all(|inset| *inset == margin[0]) {
                owned(format!("{}pt", points(margin[0])))
            } else {
                owned(
                    margin
                        .iter()
                        .map(|inset| format!("{}pt", points(*inset)))
                        .collect::<Vec<_>>()
                        .join(","),
                )
            }
        }),
        ShapeKey::Anchor
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

/// The OfficeCLI name of an XLSB icon set (`KPISets`, [MS-XLSB] 2.5.86).
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

fn argb_value(argb: u32) -> Value<'static> {
    let [_, red, green, blue] = argb.to_be_bytes();
    Value::Color(Color::rgb(red, green, blue))
}

fn color_record(color: &xlsb::conditional_formatting::Color) -> Option<Value<'static>> {
    color.argb.map(argb_value)
}

/// A threshold's value: a number, percent, percentile or formula.
fn threshold(value: &xlsb::conditional_formatting::Value) -> Option<Value<'static>> {
    matches!(value.cfvo_type, 1 | 4 | 5 | 7)
        .then(|| value.value.clone().map(owned))
        .flatten()
}

/// OfficeCLI's `type` of a conditional format's first rule.
pub(super) fn rule_type(formatting: &Formatting) -> &'static str {
    use xlsb::conditional_formatting::RuleType;
    let Some(rule) = formatting.rules.first() else {
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
            8 => match rule.parameter {
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

fn cf_value(cf: &CfData, key: CfKey) -> Option<Value<'_>> {
    use xlsb::conditional_formatting::RuleType;
    let rule = cf.formatting.rules.first()?;
    let dxf = cf.dxf.as_ref();
    let flag = |value: bool| value.then_some(Value::Bool(true));
    let average =
        rule.rule_type == RuleType::Expression && matches!(rule.template, 25 | 26 | 29 | 30);
    match key {
        CfKey::Type => Some(Value::Token(rule_type(&cf.formatting))),
        CfKey::Ref => Some(owned(ranges(&cf.formatting))),
        CfKey::Priority => Some(int(rule.priority)),
        CfKey::StopIfTrue => flag(rule.stop_if_true),
        CfKey::DxfId => rule.dxf_id.map(int),
        CfKey::Operator => match rule.rule_type {
            RuleType::CellIs => rule
                .operator
                .and_then(|operator| CF_OPERATORS.get(usize::from(operator).checked_sub(1)?))
                .map(|name| Value::Token(name)),
            RuleType::Expression if rule.template == 8 => TEXT_OPERATORS
                .get(rule.parameter as usize)
                .map(|name| Value::Token(name)),
            _ => None,
        },
        CfKey::Value => (rule.rule_type == RuleType::CellIs)
            .then(|| rule.formula_texts.first().map(|value| text(value)))
            .flatten(),
        CfKey::Value2 => (rule.rule_type == RuleType::CellIs)
            .then(|| rule.formula_texts.get(1).map(|value| text(value)))
            .flatten(),
        CfKey::Formula => (rule.rule_type == RuleType::Expression && rule.template == 1)
            .then(|| rule.formula_texts.first().map(|value| text(value)))
            .flatten(),
        CfKey::Text => rule.text.as_deref().map(text),
        CfKey::Rank => (rule.rule_type == RuleType::TopN).then(|| int(rule.parameter)),
        CfKey::Percent => flag(rule.percent),
        CfKey::Bottom => flag(rule.bottom),
        CfKey::AboveAverage => {
            (average && matches!(rule.template, 26 | 30)).then_some(Value::Bool(false))
        },
        CfKey::EqualAverage => flag(average && matches!(rule.template, 29 | 30)),
        CfKey::StdDev => (average && rule.parameter > 0).then(|| int(rule.parameter)),
        CfKey::Period => (rule.rule_type == RuleType::Expression)
            .then(|| {
                rule.template
                    .checked_sub(15)
                    .and_then(|index| PERIODS.get(index as usize))
                    .map(|(name, _)| Value::Token(name))
            })
            .flatten(),
        CfKey::Color => rule.data_bar.as_ref().map(|bar| argb_value(bar.color)),
        CfKey::Min => rule
            .data_bar
            .as_ref()
            .and_then(|bar| threshold(&bar.min_cfvo)),
        CfKey::Max => rule
            .data_bar
            .as_ref()
            .and_then(|bar| threshold(&bar.max_cfvo)),
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
            .and_then(|scale| color_record(&scale.min_color_record)),
        CfKey::MidColor => rule
            .color_scale
            .as_ref()
            .and_then(|scale| scale.mid_color_record.as_ref())
            .and_then(color_record),
        CfKey::MaxColor => rule
            .color_scale
            .as_ref()
            .and_then(|scale| color_record(&scale.max_color_record)),
        CfKey::Midpoint => rule
            .color_scale
            .as_ref()
            .and_then(|scale| scale.mid_cfvo.as_ref())
            .and_then(threshold),
        CfKey::Iconset => rule
            .icon_set
            .as_ref()
            .and_then(|icons| ICON_SETS.get(usize::from(icons.icon_set_type)))
            .map(|name| Value::Token(name)),
        CfKey::Reverse => rule.icon_set.as_ref().and_then(|icons| flag(icons.reverse)),
        CfKey::Fill => dxf?.fill.map(rgb),
        CfKey::FontColor => dxf?.font_color.map(rgb),
        CfKey::FontBold => dxf?.bold.map(Value::Bool),
        CfKey::FontItalic => dxf?.italic.map(Value::Bool),
        CfKey::FontStrike => dxf?.strike.map(Value::Bool),
        CfKey::FontUnderline => dxf?.underline.and_then(|underline| match underline {
            Underline::None => None,
            Underline::Double | Underline::DoubleAccounting => Some(Value::Token("double")),
            _ => Some(Value::Token("single")),
        }),
        CfKey::FontSize => dxf?.size.map(points),
        CfKey::FontName => dxf?.name.as_deref().map(text),
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

/// The OfficeCLI names of the XLSB validation types, by wire value 0–7.
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

/// The OfficeCLI names of the XLSB validation operators, by wire value 0–7.
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

/// The OfficeCLI names of the XLSB error styles, by wire value 0–2.
pub(super) const ERROR_STYLES: [&str; 3] = ["stop", "warning", "information"];

fn validation_value(
    validation: &xlsb::package::data_validation::Validation,
    key: ValidationKey,
) -> Option<Value<'_>> {
    let kind = validation.validation_type;
    let typed = kind != 0;
    match key {
        ValidationKey::Ref => Some(text(validation.cell_ranges.as_str())),
        ValidationKey::Type => VALIDATION_TYPES
            .get(usize::from(kind))
            .map(|name| Value::Token(name)),
        ValidationKey::Operator => (typed && !matches!(kind, 3 | 7))
            .then(|| VALIDATION_OPERATORS.get(usize::from(validation.operator)))
            .flatten()
            .map(|name| Value::Token(name)),
        ValidationKey::Formula1 => validation
            .list_formula
            .as_deref()
            .or(validation.formula1.as_deref())
            .map(text),
        ValidationKey::Formula2 => validation.formula2.as_deref().map(text),
        ValidationKey::AllowBlank => Some(Value::Bool(validation.allow_blank)),
        ValidationKey::ShowError => Some(Value::Bool(validation.show_error_message)),
        ValidationKey::ShowInput => Some(Value::Bool(validation.show_input_message)),
        ValidationKey::ErrorTitle => validation.error_title.as_deref().map(text),
        ValidationKey::Error => validation.error_text.as_deref().map(text),
        ValidationKey::PromptTitle => validation.input_title.as_deref().map(text),
        ValidationKey::Prompt => validation.input_text.as_deref().map(text),
        ValidationKey::ErrorStyle => ERROR_STYLES
            .get(usize::from(validation.error_style))
            .map(|name| Value::Token(name)),
        // `show_dropdown` is the in-cell arrow (the record stores it
        // inverted).
        ValidationKey::InCellDropdown => Some(Value::Bool(validation.show_dropdown)),
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

fn break_value(
    kind: Kind,
    item: xlsb::page_breaks::Break,
    key: BreakKey,
) -> Option<Value<'static>> {
    let full = if kind == Kind::RowBreak {
        xlsb::page_breaks::LAST_COLUMN
    } else {
        xlsb::page_breaks::LAST_ROW
    };
    let partial = item.min != 0 || item.max != full;
    match key {
        BreakKey::Index => Some(int(item.id)),
        BreakKey::Manual => Some(Value::Bool(item.manual)),
        BreakKey::Min => partial.then(|| int(item.min)),
        BreakKey::Max => partial.then(|| int(item.max)),
    }
}

/// The canonical names of the feature elements: settable, addable and
/// removable through the unified verbs (`capabilities()`).
pub(crate) const FEATURE_ELEMENTS: &[&str] = &[
    "comment",
    "hyperlink",
    "picture",
    "shape",
    "cf",
    "dataValidation",
    "autofilter",
    "rowbreak",
    "colbreak",
];
