//! The RTF parts of record 0328: sections, bookmarks, main-story fields
//! and tables of contents, paragraph tab stops, text boxes and drawing
//! shapes, read from the native model's source-ordered collections.

use std::borrow::Cow;

use litchi_rtf::field::{Field, FieldOwner, FieldType, TableOfContentsOption};
use litchi_rtf::{
    PageNumberFormat, PageOrientation, SectionBreakType, SectionProperties, ShapeType,
    TabAlignment, TabLeader, TabStop, VerticalAlignment, WrapMode,
};

use super::Rtf;
use crate::node::Value;
use crate::value::Length;

/// Whether a field belongs to the main story (a body paragraph or a body
/// table cell).
const fn main_story(field: &Field<'_>) -> bool {
    matches!(field.owner, FieldOwner::Body | FieldOwner::TableCell(_))
}

/// The native indexes of the main-story fields, `TOC` fields only when
/// `toc`.
pub(crate) fn fields(doc: &Rtf, toc: bool) -> impl Iterator<Item = usize> + '_ {
    doc.0
        .package
        .fields()
        .iter()
        .enumerate()
        .filter(move |(_, field)| main_story(field) && (!toc || field.field_type == FieldType::Toc))
        .map(|(index, _)| index)
}

/// A text box: a text-box shape (by its index among the shapes) or a
/// legacy text box (by its index among those).
#[derive(Clone, Copy)]
pub(crate) enum TextBoxRef {
    Shape(usize),
    Legacy(usize),
}

/// Whether a shape is a text box.
fn is_text_box(shape: &litchi_rtf::Shape<'_>) -> bool {
    shape.shape_type == ShapeType::TextBox
}

/// The text boxes of the main story: text-box shapes, then legacy text
/// boxes.
pub(crate) fn text_boxes(doc: &Rtf) -> Vec<TextBoxRef> {
    let package = &doc.0.package;
    let mut out: Vec<TextBoxRef> = package
        .shapes()
        .iter()
        .enumerate()
        .filter(|(_, shape)| is_text_box(shape))
        .map(|(index, _)| TextBoxRef::Shape(index))
        .collect();
    out.extend((0..package.raw().legacy_text_boxes().len()).map(TextBoxRef::Legacy));
    out
}

/// The native indexes of the main-story shapes that are not text boxes.
pub(crate) fn shapes(doc: &Rtf) -> impl Iterator<Item = usize> + '_ {
    doc.0
        .package
        .shapes()
        .iter()
        .enumerate()
        .filter(|(_, shape)| !is_text_box(shape))
        .map(|(index, _)| index)
}

fn twips(value: i32) -> Option<Value<'static>> {
    Some(Value::Length(Length::from_twips(value)))
}

fn extent(value: i32) -> Option<Value<'static>> {
    (value >= 0).then(|| Value::Length(Length::from_twips(value)))
}

/// The DOCX `type` of a section break.
const fn section_type(kind: SectionBreakType) -> &'static str {
    match kind {
        SectionBreakType::Continuous => "continuous",
        SectionBreakType::Column => "nextColumn",
        SectionBreakType::Page => "nextPage",
        SectionBreakType::EvenPage => "evenPage",
        SectionBreakType::OddPage => "oddPage",
    }
}

/// The DOCX `pageNumFmt` of a page-number format.
fn page_number_format(format: PageNumberFormat) -> Cow<'static, str> {
    Cow::Borrowed(match format {
        PageNumberFormat::Decimal => "decimal",
        PageNumberFormat::UpperRoman => "upperRoman",
        PageNumberFormat::LowerRoman => "lowerRoman",
        PageNumberFormat::UpperLetter => "upperLetter",
        PageNumberFormat::LowerLetter => "lowerLetter",
        other => {
            let name = format!("{other:?}");
            let mut chars = name.chars();
            return Cow::Owned(match chars.next() {
                Some(first) => first.to_ascii_lowercase().to_string() + chars.as_str(),
                None => name,
            });
        },
    })
}

/// The value at `index` of [`super::spec::SECTION_PROPS`].
pub(crate) fn section_value(
    properties: &SectionProperties,
    index: usize,
) -> Option<Value<'static>> {
    match index {
        0 => Some(Value::Token(section_type(properties.break_type))),
        1 => extent(properties.page_width),
        2 => extent(properties.page_height),
        3 => Some(Value::Token(match properties.orientation {
            PageOrientation::Landscape => "landscape",
            _ => "portrait",
        })),
        4 => twips(properties.margin_top),
        5 => twips(properties.margin_bottom),
        6 => extent(properties.margin_left),
        7 => extent(properties.margin_right),
        8 => extent(properties.header_distance),
        9 => extent(properties.footer_distance),
        10 => extent(properties.margin_gutter),
        11 => Some(Value::Int(i64::from(properties.columns.count.max(1)))),
        12 => extent(properties.columns.default_spacing),
        13 => Some(Value::Bool(properties.columns.separator)),
        14 => Some(Value::Bool(properties.title_page)),
        15 => Some(Value::Text(page_number_format(
            properties.page_number_format,
        ))),
        16 => Some(Value::Int(i64::from(properties.page_number_start))),
        17 => properties.direction.map(|direction| {
            Value::Token(match direction {
                litchi_rtf::TextDirection::RightToLeft => "rtl",
                litchi_rtf::TextDirection::LeftToRight => "ltr",
            })
        }),
        18 => Some(Value::Token(match properties.vertical_alignment {
            VerticalAlignment::Top => "top",
            VerticalAlignment::Center => "center",
            VerticalAlignment::Justify => "both",
            VerticalAlignment::Bottom => "bottom",
        })),
        _ => None,
    }
}

/// A field's DOCX `fieldType`: its instruction keyword, lowercased.
pub(crate) fn field_type(field: &Field<'_>) -> Option<String> {
    field
        .instruction
        .split_whitespace()
        .next()
        .map(|keyword| keyword.trim_start_matches('\\').to_ascii_lowercase())
        .filter(|keyword| !keyword.is_empty())
}

/// The value at `index` of [`super::spec::FIELD_PROPS`].
pub(crate) fn field_value<'a>(field: &'a Field<'a>, index: usize) -> Option<Value<'a>> {
    match index {
        0 => field_type(field).map(|kind| Value::Text(Cow::Owned(kind))),
        1 => Some(Value::Text(Cow::Borrowed(field.instruction.trim()))),
        2 => Some(Value::Bool(field.status.locked)),
        3 => Some(Value::Bool(!field.result.is_empty())),
        _ => None,
    }
}

/// The value at `index` of [`super::spec::TOC_PROPS`], from the field's
/// switches.
pub(crate) fn toc_value<'a>(field: &'a Field<'a>, index: usize) -> Option<Value<'a>> {
    let toc = field.table_of_contents()?;
    let options = toc.options();
    match index {
        0 => options.iter().find_map(|option| match option {
            TableOfContentsOption::HeadingStyleRange(range) => Some(Value::Text(Cow::Owned(
                range
                    .as_deref()
                    .map_or_else(|| "1-9".to_owned(), str::to_owned),
            ))),
            _ => None,
        }),
        1 => Some(Value::Bool(options.iter().any(|option| {
            matches!(option, TableOfContentsOption::Hyperlinks)
        }))),
        2 => Some(Value::Bool(!options.iter().any(|option| {
            matches!(option, TableOfContentsOption::OmitPageNumbers(None))
        }))),
        3 => Some(Value::Text(Cow::Borrowed(field.instruction.trim()))),
        _ => None,
    }
}

/// The DOCX `val` of a tab alignment.
pub(crate) const fn tab_val(alignment: TabAlignment) -> &'static str {
    match alignment {
        TabAlignment::Left => "left",
        TabAlignment::Center => "center",
        TabAlignment::Right => "right",
        TabAlignment::Decimal => "decimal",
        TabAlignment::Bar => "bar",
    }
}

/// The DOCX `leader` of a tab leader.
pub(crate) const fn tab_leader(leader: TabLeader) -> &'static str {
    match leader {
        TabLeader::None => "none",
        TabLeader::Dot => "dot",
        TabLeader::MiddleDot => "middleDot",
        TabLeader::Hyphen => "hyphen",
        TabLeader::Underscore => "underscore",
        TabLeader::ThickLine => "heavy",
        TabLeader::Equal => "equal",
    }
}

/// The value at `index` of [`super::spec::TAB_PROPS`].
pub(crate) fn tab_value(stop: TabStop, index: usize) -> Option<Value<'static>> {
    match index {
        0 => Some(Value::Int(i64::from(stop.position))),
        1 => Some(Value::Token(tab_val(stop.alignment))),
        2 => Some(Value::Token(tab_leader(stop.leader))),
        _ => None,
    }
}

/// The DOCX preset name of a shape type.
fn geometry(kind: ShapeType) -> Cow<'static, str> {
    Cow::Borrowed(match kind {
        ShapeType::Rectangle => "rect",
        ShapeType::RoundRectangle => "roundRect",
        ShapeType::Ellipse => "ellipse",
        ShapeType::Line => "line",
        ShapeType::Polygon => "polygon",
        ShapeType::Arc => "arc",
        ShapeType::TextBox => "textbox",
        ShapeType::PictureFrame => "pictureFrame",
        ShapeType::Group => "group",
        ShapeType::Custom(number) => return Cow::Owned(format!("shapeType{number}")),
        ShapeType::Unknown => return Cow::Borrowed("unknown"),
    })
}

/// The value at `index` of [`super::spec::SHAPE_PROPS`].
pub(crate) fn shape_value<'a>(shape: &'a litchi_rtf::Shape<'a>, index: usize) -> Option<Value<'a>> {
    let geometry_value = &shape.geometry;
    match index {
        0 => extent(geometry_value.width),
        1 => extent(geometry_value.height),
        2 => twips(geometry_value.x),
        3 => twips(geometry_value.y),
        4 => Some(Value::Token(match shape.wrap_mode {
            WrapMode::None => "none",
            WrapMode::Square => "square",
            WrapMode::Tight => "tight",
            WrapMode::Through => "through",
            WrapMode::TopAndBottom => "topAndBottom",
            WrapMode::Behind => "behind",
            WrapMode::InFront => "front",
        })),
        5 => Some(Value::Bool(shape.behind_doc)),
        6 => Some(Value::Text(geometry(shape.shape_type))),
        7 => (!shape.name.is_empty()).then(|| Value::Text(Cow::Borrowed(shape.name.as_ref()))),
        _ => None,
    }
}
