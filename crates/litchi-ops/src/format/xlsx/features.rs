//! Worksheet features as unified nodes (record 0320): legacy notes
//! (`comment`), hyperlinks, drawing pictures and shapes, conditional
//! formats (`cf`), data validations, the auto-filter and manual page
//! breaks.
//!
//! Every feature is a child of its worksheet, addressed by its position
//! among the worksheet's features of its kind (`/Sheet1/comment[2]`,
//! `/Sheet1/cf[1]`; the auto-filter is `/Sheet1/autofilter`). A node holds
//! one [`Feature`]: the kind, the position and the leaf value it was listed
//! from, behind one `Arc`. Listing a kind reads only the worksheet parts that
//! kind lives in: the worksheet's blocks outside `sheetData` for
//! conditional formats, validations, the auto-filter, hyperlinks and page
//! breaks; the comments part for notes; the drawing part for pictures and
//! shapes; and the styles part's differential formats for a conditional
//! format's formatting.
//!
//! Positions follow the native order, except notes, which a comments part
//! keys by cell: they are numbered in row-major order of their cells. The
//! write verbs live in `feature_write`.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};
use litchi_xlsx as xlsx;
use xlsx::conditional_formatting::edit::{DifferentialStyle, Underline};
use xlsx::conditional_formatting::{Formatting, Payload, ValueKind};
use xlsx::drawing::edit::{Align, Element, ElementKind, Fill, Line, Placement, VerticalAlign};

use super::node::{Item, Node, Sheet, SheetHandle};
use super::spec::Kind;
use super::{Xlsx, XlsxOwner};
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
    Hyperlink(xlsx::hyperlinks::Hyperlink),
    Drawing(Box<Element>),
    Cf(Box<CfData>),
    Validation(Box<xlsx::data_validation::Validation>),
    Autofilter(Box<xlsx::auto_filter::Definition>),
    Break(xlsx::page_breaks::Break),
}

/// A legacy note.
#[derive(Clone)]
pub(crate) struct Note {
    pub(super) reference: String,
    pub(super) author: String,
    pub(super) text: String,
    /// The note's rich-text runs; empty for plain text.
    pub(super) runs: Vec<xlsx::workbook::comments::TextRun>,
}

/// A conditional format and the formatting its first rule applies.
#[derive(Clone)]
pub(crate) struct CfData {
    pub(super) formatting: Formatting,
    pub(super) dxf: Option<DifferentialStyle>,
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
                link.display()
                    .or(link.external_target())
                    .or(link.location())
                    .unwrap_or(""),
            ),
            Data::Drawing(element) => match element.kind {
                ElementKind::Shape => Cow::Borrowed(element.text.as_deref()?),
                _ => Cow::Borrowed(element.description.as_deref()?),
            },
            Data::Cf(cf) => Cow::Owned(ranges(&cf.formatting)),
            Data::Validation(validation) => Cow::Owned(sqref(validation)),
            Data::Autofilter(filter) => Cow::Borrowed(filter.reference()?.as_str()),
            Data::Break(_) => return None,
        })
    }

    /// The node view's one-line preview.
    pub(super) fn preview(&self) -> Option<Preview<'_>> {
        Some(Preview::Text(match &self.data {
            Data::Comment(note) => Cow::Owned(format!("{}: {}", note.reference, note.text)),
            Data::Hyperlink(link) => Cow::Borrowed(link.reference().as_str()),
            Data::Drawing(element) => Cow::Borrowed(element.name.as_deref()?),
            Data::Cf(cf) => Cow::Owned(format!(
                "{} {}",
                rule_type(&cf.formatting),
                ranges(&cf.formatting)
            )),
            Data::Validation(validation) => Cow::Owned(sqref(validation)),
            Data::Autofilter(filter) => Cow::Borrowed(filter.reference()?.as_str()),
            Data::Break(item) => Cow::Owned(item.id().to_string()),
        }))
    }

    /// The property table of the feature's kind.
    pub(super) fn table(&self) -> &'static [PropSpec] {
        props(self.kind)
    }

    /// The value at `index` of the feature's property table; `sheet` is the
    /// name of the feature's worksheet.
    pub(super) fn value<'a>(&'a self, index: usize, sheet: &'a str) -> Option<Value<'a>> {
        match &self.data {
            Data::Comment(note) => match *COMMENT_KEYS.get(index)? {
                CommentKey::AnchoredTo => Some(owned(format!("/{sheet}/{}", note.reference))),
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
                AutofilterKey::Range => filter.reference().map(|range| text(range.as_str())),
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

fn rgb(color: xlsx::Rgb) -> Value<'static> {
    Value::Color(Color::rgb(color.red(), color.green(), color.blue()))
}

fn points(points: f64) -> Value<'static> {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "font sizes are a few thousand points at most"
    )]
    Value::Length(Length::from_emu((points * 12_700.0).round() as i64))
}

fn ranges(formatting: &Formatting) -> String {
    formatting
        .ranges
        .iter()
        .map(xlsx::conditional_formatting::Range::as_str)
        .collect::<Vec<_>>()
        .join(" ")
}

fn sqref(validation: &xlsx::data_validation::Validation) -> String {
    validation
        .sqref()
        .ranges()
        .iter()
        .map(xlsx::data_validation::Range::as_str)
        .collect::<Vec<_>>()
        .join(" ")
}

// ---------------------------------------------------------------------------
// Listing

impl SheetHandle {
    fn comments(&self) -> Result<Option<xlsx::workbook::comments::Comments>> {
        Ok(match self {
            Self::Eager { sheet, .. } => sheet.comments()?,
            Self::Source { sheet } => sheet.comments()?,
        })
    }

    pub(super) fn hyperlinks(&self) -> Result<Vec<xlsx::hyperlinks::Hyperlink>> {
        Ok(match self {
            Self::Eager { sheet, .. } => sheet.hyperlinks()?,
            Self::Source { sheet } => sheet.hyperlinks()?,
        })
    }

    fn drawing_elements(&self) -> Result<Vec<Element>> {
        Ok(match self {
            Self::Eager { sheet, workbook } => match &workbook.0.owner {
                XlsxOwner::Eager(package) => package.drawing_elements(sheet.position())?,
                XlsxOwner::Source(_) => return Err(lost_owner()),
            },
            Self::Source { sheet } => sheet.drawing_elements()?,
        })
    }

    fn conditional_formats(&self) -> Result<Vec<Formatting>> {
        Ok(match self {
            Self::Eager { sheet, .. } => sheet.conditional_formattings()?,
            Self::Source { sheet } => sheet.conditional_formattings()?,
        })
    }

    fn validations(&self) -> Result<Vec<xlsx::data_validation::Collection>> {
        Ok(match self {
            Self::Eager { sheet, .. } => sheet.data_validations()?,
            Self::Source { sheet } => sheet.data_validations()?,
        })
    }

    fn auto_filter(&self) -> Result<Option<xlsx::auto_filter::Definition>> {
        Ok(match self {
            Self::Eager { sheet, .. } => sheet.auto_filter()?,
            Self::Source { sheet } => sheet.auto_filter()?,
        })
    }

    fn page_breaks(&self) -> Result<xlsx::page_breaks::PageBreaks> {
        Ok(match self {
            Self::Eager { sheet, .. } => sheet.page_breaks()?,
            Self::Source { sheet } => sheet.page_breaks()?,
        })
    }
}

fn lost_owner() -> Error {
    Error::new(
        Code::InternalError,
        "an eager worksheet handle lost its eager owner",
    )
    .with_format(FileFormat::Xlsx)
}

/// The workbook's differential formats, read only when a conditional
/// format is listed.
fn differential_formats(file: &Xlsx) -> Result<Vec<xlsx::conditional_formatting::Differential>> {
    Ok(match &file.0.owner {
        XlsxOwner::Eager(package) => package.workbook_ref().differential_formats()?,
        XlsxOwner::Source(workbook) => workbook.differential_formats()?,
    })
}

/// The cell of an A1 reference as `(row, column)`, for ordering notes.
pub(super) fn row_major(reference: &str) -> (u32, u32) {
    litchi_core::selector::CellRef::parse(reference)
        .map_or((u32::MAX, u32::MAX), |cell| (cell.row(), cell.column()))
}

/// The features of `kind` on a worksheet, in position order.
pub(super) fn list(file: &Xlsx, sheet: &Sheet, kind: Kind) -> Result<Vec<Feature>> {
    let handle = &sheet.handle;
    if !handle.is_grid() {
        return Ok(Vec::new());
    }
    let data: Vec<Data> = match kind {
        Kind::Comment => {
            let mut notes: Vec<Note> = handle
                .comments()?
                .map(|comments| {
                    comments
                        .iter()
                        .map(|(reference, comment)| note(reference, comment))
                        .collect()
                })
                .unwrap_or_default();
            notes.sort_by_key(|note| row_major(&note.reference));
            notes.into_iter().map(Data::Comment).collect()
        },
        Kind::Hyperlink => handle
            .hyperlinks()?
            .into_iter()
            .map(Data::Hyperlink)
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
        Kind::Cf => {
            let formats = handle.conditional_formats()?;
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
                        .and_then(|rule| match &rule.differential_format {
                            Some(xlsx::conditional_formatting::DifferentialRef::StylesIndex(
                                index,
                            )) => dxfs.get(*index as usize),
                            _ => None,
                        })
                        .map(DifferentialStyle::read)
                        .transpose()?;
                    Ok(Data::Cf(Box::new(CfData { formatting, dxf })))
                })
                .collect::<Result<_>>()?
        },
        Kind::DataValidation => handle
            .validations()?
            .into_iter()
            .flat_map(|collection| collection.validations().to_vec())
            .map(|validation| Data::Validation(Box::new(validation)))
            .collect(),
        Kind::Autofilter => handle
            .auto_filter()?
            .map(|filter| Data::Autofilter(Box::new(filter)))
            .into_iter()
            .collect(),
        Kind::RowBreak | Kind::ColBreak => {
            let breaks = handle.page_breaks()?;
            let collection = if kind == Kind::RowBreak {
                breaks.horizontal()
            } else {
                breaks.vertical()
            };
            collection
                .map(|collection| {
                    collection
                        .breaks()
                        .iter()
                        .copied()
                        .map(Data::Break)
                        .collect()
                })
                .unwrap_or_default()
        },
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
    file: &'f Xlsx,
    sheet: &Sheet,
    kind: Kind,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    if kind == Kind::Comment {
        return comment(file, sheet, select);
    }
    let mut features = list(file, sheet, kind)?;
    let index = match select {
        Select::Ordinal(position) => Some(position.get()),
        Select::None => Some(0),
        Select::Last => features.len().checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(match index.filter(|index| *index < features.len()) {
        Some(index) => Lookup::Found(Node::new_feature(
            file,
            sheet.position,
            features.swap_remove(index),
        )),
        None => Lookup::Missing,
    })
}

/// [`lookup`] of a note: the notes are ordered row-major by their cells
/// alone, and only the selected one is copied (record 0326). The order is
/// [`list`]'s: its stable sort keeps equal cells in part order, as the
/// part index does here.
fn comment<'f>(
    file: &'f Xlsx,
    sheet: &Sheet,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let comments = if sheet.handle.is_grid() {
        sheet.handle.comments()?
    } else {
        None
    };
    let entries: Vec<_> = comments
        .as_ref()
        .map(|comments| comments.iter().collect())
        .unwrap_or_default();
    let index = match select {
        Select::Ordinal(position) => Some(position.get()),
        Select::None => Some(0),
        Select::Last => entries.len().checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    let Some(index) = index.filter(|index| *index < entries.len()) else {
        return Ok(Lookup::Missing);
    };
    let mut order: Vec<((u32, u32), usize)> = entries
        .iter()
        .enumerate()
        .map(|(at, (reference, _))| (row_major(reference), at))
        .collect();
    let (_, &mut (_, at), _) = order.select_nth_unstable(index);
    let (reference, comment) = entries[at];
    Ok(Lookup::Found(Node::new_feature(
        file,
        sheet.position,
        Feature {
            kind: Kind::Comment,
            index: u32::try_from(index).unwrap_or(u32::MAX),
            data: Data::Comment(note(reference, comment)),
        },
    )))
}

/// A note's copy of a comment of the part.
fn note(reference: &str, comment: &xlsx::workbook::comments::Comment) -> Note {
    Note {
        reference: reference.to_owned(),
        author: comment.author.clone(),
        text: comment.text.clone(),
        runs: comment.runs.clone(),
    }
}

impl<'f> Node<'f> {
    /// A worksheet feature.
    pub(super) fn new_feature(file: &'f Xlsx, sheet: usize, feature: Feature) -> Self {
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
        follow_up: "record 0320 follow-up: Office 2010 conditional-format extensions",
        reason: "the core conditional-format writer refuses x14 extension rules, so data-bar negative and axis styling cannot be authored yet",
    }
}

static NEGATIVE_COLOR_GAP: GapSpec = x14_gap("conditionalFormatting", "negativeColor");
static AXIS_COLOR_GAP: GapSpec = x14_gap("conditionalFormatting", "axisColor");
static AXIS_POSITION_GAP: GapSpec = x14_gap("conditionalFormatting", "axisPosition");
static DIRECTION_GAP: GapSpec = x14_gap("conditionalFormatting", "direction");

const fn effect_gap(element: &'static str, key: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Property { element, key },
        code: Code::UnsupportedProperty,
        follow_up: "record 0320 follow-up: drawing effects, crops and gradient fills",
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
    FontBold,
    FontItalic,
    FontUnderline,
    FontSize,
    FontColor,
    FontName,
    Runs,
}

keyed!(CommentKey, CommentKey, COMMENT_TABLE, COMMENT_PROPS, COMMENT_KEYS, [
    entry(read("ref", TEXT), Access::SetAndAdd(CommentKey::Ref)) => Ref,
    entry(read("author", TEXT), Access::SetAndAdd(CommentKey::Author)) => Author,
    entry(read("text", TEXT), Access::SetAndAdd(CommentKey::Text)) => Text,
    entry(read("anchoredTo", TEXT), Access::ReadOnly) => AnchoredTo,
    entry(read("font.bold", BOOL), Access::SetAndAdd(CommentKey::FontBold)) => FontBold,
    entry(read("font.italic", BOOL), Access::SetAndAdd(CommentKey::FontItalic)) => FontItalic,
    entry(read("font.underline", TEXT), Access::SetAndAdd(CommentKey::FontUnderline)) => FontUnderline,
    entry(read("font.size", FONT_SIZE), Access::SetAndAdd(CommentKey::FontSize)) => FontSize,
    entry(read("font.color", SHEET_COLOR), Access::SetAndAdd(CommentKey::FontColor)) => FontColor,
    entry(read("font.name", TEXT), Access::SetAndAdd(CommentKey::FontName)) => FontName,
    entry(read("runs", TEXT), Access::SetAndAdd(CommentKey::Runs)) => Runs,
], []);

/// The size and typeface Excel gives note text, which a note's single-run
/// `font.size` and `font.name` leave unreported (as OfficeCLI does).
pub(super) const NOTE_SIZE: f64 = 9.0;
/// See [`NOTE_SIZE`].
pub(super) const NOTE_FONT: &str = "Tahoma";

/// The `RRGGBB` of an `AARRGGBB` run color.
fn run_rgb(color: Option<&xlsx::workbook::comments::RunColor>) -> Option<xlsx::Rgb> {
    match color? {
        xlsx::workbook::comments::RunColor::Rgb(argb) => {
            xlsx::Rgb::from_hex(argb.get(argb.len().checked_sub(6)?..)?).ok()
        },
        _ => None,
    }
}

/// A note's runs as OfficeCLI's `runs` JSON array.
fn runs_json(runs: &[xlsx::workbook::comments::TextRun]) -> String {
    let items: Vec<serde_json::Value> = runs
        .iter()
        .map(|run| {
            let font = &run.font;
            let mut item = serde_json::Map::new();
            item.insert("text".into(), run.text.clone().into());
            for (key, on) in [
                ("bold", font.bold),
                ("italic", font.italic),
                ("strike", font.strike),
                (
                    "superscript",
                    font.vert_align.as_deref().map(|v| v == "superscript"),
                ),
                (
                    "subscript",
                    font.vert_align.as_deref().map(|v| v == "subscript"),
                ),
            ] {
                if on == Some(true) {
                    item.insert(key.into(), true.into());
                }
            }
            if let Some(underline) = font.underline.as_deref().filter(|u| *u != "none") {
                item.insert("underline".into(), underline.into());
            }
            if let Some(size) = font.size {
                item.insert(
                    "size".into(),
                    format!("{}pt", crate::value::DotNetDouble(size)).into(),
                );
            }
            if let Some(color) = run_rgb(font.color.as_ref()) {
                item.insert(
                    "color".into(),
                    format!(
                        "#{:02X}{:02X}{:02X}",
                        color.red(),
                        color.green(),
                        color.blue()
                    )
                    .into(),
                );
            }
            if let Some(name) = &font.name {
                item.insert("font".into(), name.clone().into());
            }
            serde_json::Value::Object(item)
        })
        .collect();
    serde_json::Value::Array(items).to_string()
}

fn comment_value(note: &Note, key: CommentKey) -> Option<Value<'_>> {
    match key {
        CommentKey::Ref => Some(text(&note.reference)),
        CommentKey::Author => Some(text(&note.author)),
        CommentKey::Text => Some(text(&note.text)),
        // Read through `Feature::value`, which knows the sheet's name.
        CommentKey::AnchoredTo => None,
        // Two or more runs read as `runs`; one run as its non-default
        // `font.*` facets, as OfficeCLI reports them.
        CommentKey::Runs => (note.runs.len() > 1).then(|| owned(runs_json(&note.runs))),
        key => {
            let [run] = note.runs.as_slice() else {
                return None;
            };
            let font = &run.font;
            match key {
                CommentKey::FontBold => (font.bold == Some(true)).then_some(Value::Bool(true)),
                CommentKey::FontItalic => (font.italic == Some(true)).then_some(Value::Bool(true)),
                CommentKey::FontUnderline => {
                    font.underline.as_deref().filter(|u| *u != "none").map(text)
                },
                CommentKey::FontSize => font.size.filter(|size| *size != NOTE_SIZE).map(points),
                CommentKey::FontColor => run_rgb(font.color.as_ref()).map(rgb),
                CommentKey::FontName => font
                    .name
                    .as_deref()
                    .filter(|name| *name != NOTE_FONT)
                    .map(text),
                _ => None,
            }
        },
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

fn hyperlink_value(link: &xlsx::hyperlinks::Hyperlink, key: HyperlinkKey) -> Option<Value<'_>> {
    match key {
        HyperlinkKey::Ref => Some(text(link.reference().as_str())),
        HyperlinkKey::Url => link.external_target().map(text),
        HyperlinkKey::Location => link.location().map(text),
        HyperlinkKey::Display => link.display().map(text),
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
    // `set src` replaces the image (record 0338).
    entry(write_only("src", TEXT), Access::SetAndAdd(PictureKey::Src)) => Src,
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
            Fill::Solid(color) => Some(owned(format!(
                "#{:02X}{:02X}{:02X}",
                color.red(),
                color.green(),
                color.blue()
            ))),
            _ => None,
        },
        ShapeKey::Line => match element.line? {
            Line::None => Some(Value::Token("none")),
            Line::Solid { color, width, .. } => {
                let mut value = format!(
                    "#{:02X}{:02X}{:02X}",
                    color.red(),
                    color.green(),
                    color.blue()
                );
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

fn color_value(color: &xlsx::conditional_formatting::Color) -> Option<Value<'static>> {
    color.rgb.map(rgb)
}

fn threshold(value: &xlsx::conditional_formatting::Value) -> Option<Value<'static>> {
    match value.kind {
        ValueKind::Number | ValueKind::Percent | ValueKind::Percentile | ValueKind::Formula => {
            value
                .value
                .clone()
                .or_else(|| value.formula.clone())
                .map(owned)
        },
        _ => None,
    }
}

/// OfficeCLI's `type` of a conditional format's first rule.
fn rule_type(formatting: &Formatting) -> &'static str {
    use xlsx::conditional_formatting::Kind as RuleKind;
    match formatting.rules.first().and_then(|rule| rule.rule_type) {
        Some(RuleKind::Expression) => "formula",
        Some(RuleKind::Top10) => "topN",
        Some(kind) => kind.as_str(),
        None => "",
    }
}

fn cf_value(cf: &CfData, key: CfKey) -> Option<Value<'_>> {
    let rule = cf.formatting.rules.first()?;
    let dxf = cf.dxf.as_ref();
    let flag = |value: bool| value.then_some(Value::Bool(true));
    let payload = rule.payload.as_ref();
    match key {
        CfKey::Type => Some(Value::Token(rule_type(&cf.formatting))),
        CfKey::Ref => Some(owned(ranges(&cf.formatting))),
        CfKey::Priority => rule.priority.map(int),
        CfKey::StopIfTrue => flag(rule.stop_if_true),
        CfKey::DxfId => match rule.differential_format {
            Some(xlsx::conditional_formatting::DifferentialRef::StylesIndex(index)) => {
                Some(int(index))
            },
            _ => None,
        },
        CfKey::Operator => rule
            .operator
            .map(|operator| Value::Token(operator.as_str())),
        CfKey::Value => match rule.rule_type {
            Some(xlsx::conditional_formatting::Kind::CellIs) => {
                rule.formulas.first().map(|v| text(v))
            },
            _ => None,
        },
        CfKey::Value2 => match rule.rule_type {
            Some(xlsx::conditional_formatting::Kind::CellIs) => {
                rule.formulas.get(1).map(|v| text(v))
            },
            _ => None,
        },
        CfKey::Formula => match rule.rule_type {
            Some(xlsx::conditional_formatting::Kind::Expression) => {
                rule.formulas.first().map(|v| text(v))
            },
            _ => None,
        },
        CfKey::Text => rule.text.as_deref().map(text),
        CfKey::Rank => rule.rank.map(int),
        CfKey::Percent => flag(rule.percent),
        CfKey::Bottom => flag(rule.bottom),
        CfKey::AboveAverage => match rule.rule_type {
            Some(xlsx::conditional_formatting::Kind::AboveAverage) if !rule.above_average => {
                Some(Value::Bool(false))
            },
            _ => None,
        },
        CfKey::EqualAverage => flag(rule.equal_average),
        CfKey::StdDev => rule.standard_deviations.map(int),
        CfKey::Period => rule.time_period.map(|period| Value::Token(period.as_str())),
        CfKey::Color => match payload? {
            Payload::DataBar(bar) => bar
                .colors
                .first()
                .and_then(|named| color_value(&named.color)),
            _ => None,
        },
        CfKey::Min | CfKey::Max => match payload? {
            Payload::DataBar(bar) => bar
                .thresholds
                .get(usize::from(key == CfKey::Max))
                .and_then(threshold),
            _ => None,
        },
        CfKey::MinLength | CfKey::MaxLength => match payload? {
            Payload::DataBar(bar) => Some(int(if key == CfKey::MinLength {
                bar.min_length
            } else {
                bar.max_length
            })),
            _ => None,
        },
        CfKey::ShowValue => match payload? {
            Payload::DataBar(bar) => (!bar.show_value).then_some(Value::Bool(false)),
            Payload::IconSet(icons) => (!icons.show_value).then_some(Value::Bool(false)),
            _ => None,
        },
        CfKey::MinColor | CfKey::MidColor | CfKey::MaxColor => match payload? {
            Payload::ColorScale(scale) => {
                let colors = &scale.colors;
                let slot = match key {
                    CfKey::MinColor => Some(0),
                    CfKey::MaxColor => colors.len().checked_sub(1),
                    _ => (colors.len() == 3).then_some(1),
                }?;
                colors.get(slot).and_then(color_value)
            },
            _ => None,
        },
        CfKey::Midpoint => match payload? {
            Payload::ColorScale(scale) if scale.thresholds.len() == 3 => {
                threshold(&scale.thresholds[1])
            },
            _ => None,
        },
        CfKey::Iconset => match payload? {
            Payload::IconSet(icons) => Some(Value::Token(icons.set.as_str())),
            _ => None,
        },
        CfKey::Reverse => match payload? {
            Payload::IconSet(icons) => flag(icons.reverse),
            _ => None,
        },
        CfKey::Fill => dxf?.fill.map(rgb),
        CfKey::FontColor => dxf?.font_color.map(rgb),
        CfKey::FontBold => dxf?.bold.map(Value::Bool),
        CfKey::FontItalic => dxf?.italic.map(Value::Bool),
        CfKey::FontStrike => dxf?.strike.map(Value::Bool),
        CfKey::FontUnderline => dxf?.underline.map(|underline| {
            Value::Token(match underline {
                Underline::Double => "double",
                _ => "single",
            })
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

fn validation_value(
    validation: &xlsx::data_validation::Validation,
    key: ValidationKey,
) -> Option<Value<'_>> {
    use xlsx::data_validation::{ListSource, ValidationType};
    let typed = validation.validation_type() != ValidationType::None;
    match key {
        ValidationKey::Ref => Some(owned(sqref(validation))),
        ValidationKey::Type => Some(Value::Token(validation.validation_type().as_str())),
        ValidationKey::Operator => (typed
            && !matches!(
                validation.validation_type(),
                ValidationType::List | ValidationType::Custom
            ))
        .then(|| Value::Token(validation.operator().as_str())),
        ValidationKey::Formula1 => validation.formula1().map(|source| match source {
            ListSource::Formula(formula) => text(formula.as_str()),
            ListSource::QuotedList(list) => text(list.as_str()),
        }),
        ValidationKey::Formula2 => validation.formula2().map(|formula| text(formula.as_str())),
        ValidationKey::AllowBlank => Some(Value::Bool(validation.allow_blank())),
        ValidationKey::ShowError => Some(Value::Bool(validation.show_error_message())),
        ValidationKey::ShowInput => Some(Value::Bool(validation.show_input_message())),
        ValidationKey::ErrorTitle => validation.error_title().map(text),
        ValidationKey::Error => validation.error().map(text),
        ValidationKey::PromptTitle => validation.prompt_title().map(text),
        ValidationKey::Prompt => validation.prompt().map(text),
        ValidationKey::ErrorStyle => Some(Value::Token(validation.error_style().as_str())),
        // `showDropDown` is inverted: true hides the in-cell arrow.
        ValidationKey::InCellDropdown => Some(Value::Bool(!validation.show_drop_down())),
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
    entry(write_only("criteria0", TEXT), Access::Add(AutofilterKey::Criteria)) => Criteria,
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
    item: xlsx::page_breaks::Break,
    key: BreakKey,
) -> Option<Value<'static>> {
    let full = if kind == Kind::RowBreak {
        16_383
    } else {
        1_048_575
    };
    let partial = item.minimum() != 0 || item.maximum() != full;
    match key {
        BreakKey::Index => Some(int(item.id())),
        BreakKey::Manual => Some(Value::Bool(item.is_manual())),
        BreakKey::Min => partial.then(|| int(item.minimum())),
        BreakKey::Max => partial.then(|| int(item.maximum())),
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
