//! DOC marks under the DOCX vocabulary (record 0324): paragraph
//! `hyperlink[M]` (main-story `HYPERLINK` fields) and `tab[M]` (tab stops),
//! the document's `bookmark[N]`, `field[N]`, `toc[N]` and `section[N]`, and
//! the body's `textbox[N]` and `shape[N]`.
//!
//! Each node is read from the native tables the open parsed: `PlcfFldMom`
//! with the field text, `PlcfBkf`/`PlcfBkl` with `SttbfBkmk`, `PlcfSed` with
//! each `SEPX`, a paragraph's `sprmPChgTabsPapx` stops, `PlcftxbxTxt` and
//! `PlcfSpaMom`.

use std::borrow::Cow;

use litchi_doc::parts::pap::{TabAlignment, TabLeader, TabStop};
use litchi_doc::parts::spa::{
    ShapeAnchor, ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin,
};
use litchi_doc::section::{
    BreakKind, LineNumberRestart, PageOrientation, Section, VerticalJustification,
};
use litchi_doc::{FieldStory, FieldText};

use super::Doc;
use super::spec;
use crate::error::Result;
use crate::node::Value;
use crate::value::{Color, Length};

/// A main-story `HYPERLINK` field and the native values its node reads.
#[derive(Debug)]
pub(crate) struct Link {
    /// The field, its instruction and its cached result.
    pub(crate) field: FieldText,
    /// The external target (the first argument).
    url: Option<String>,
    /// `\l`: the bookmark the link jumps to.
    anchor: Option<String>,
    /// `\o`: the screen tip.
    tooltip: Option<String>,
    /// `\t`: the target frame.
    frame: Option<String>,
    /// The first run of the cached result, whose formatting the node reads.
    run: Option<litchi_doc::Run>,
}

impl Link {
    /// The link's displayed text: its cached result without field marks.
    pub(crate) fn text(&self) -> Cow<'_, str> {
        result_text(self.field.result.as_deref().unwrap_or_default())
    }
}

/// Text without DOC control characters, tabs kept.
pub(super) fn visible(text: &str) -> Cow<'_, str> {
    if text
        .chars()
        .any(|character| character.is_control() && character != '\t')
    {
        Cow::Owned(
            text.chars()
                .filter(|character| !character.is_control() || *character == '\t')
                .collect(),
        )
    } else {
        Cow::Borrowed(text)
    }
}

/// A field result as it displays: without the instructions of nested
/// fields and without field marks or other control characters.
pub(super) fn result_text(text: &str) -> Cow<'_, str> {
    if !text.contains('\u{13}') {
        return visible(text);
    }
    // One entry per open nested field: `true` while in its instruction.
    let mut stack: Vec<bool> = Vec::new();
    let mut output = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '\u{13}' => stack.push(true),
            '\u{14}' => {
                if let Some(top) = stack.last_mut() {
                    *top = false;
                }
            },
            '\u{15}' => {
                stack.pop();
            },
            _ if stack.iter().any(|instruction| *instruction) => {},
            _ if character.is_control() && character != '\t' => {},
            _ => output.push(character),
        }
    }
    Cow::Owned(output)
}

/// The instruction's keyword (`HYPERLINK`, `TOC`), upper-cased.
pub(crate) fn keyword(instruction: &str) -> Option<String> {
    instruction
        .split_whitespace()
        .next()
        .map(str::to_ascii_uppercase)
}

/// Every main-story field, in start order.
pub(super) fn fields(doc: &Doc) -> Result<Vec<FieldText>> {
    let mut fields: Vec<FieldText> = doc
        .reader()?
        .fields()?
        .into_iter()
        .filter(|field| field.field.story == FieldStory::Main)
        .collect();
    fields.sort_by_key(|field| field.field.start_cp);
    Ok(fields)
}

/// Every main-story `TOC` field, in start order.
pub(super) fn tocs(doc: &Doc) -> Result<Vec<FieldText>> {
    Ok(fields(doc)?
        .into_iter()
        .filter(|field| keyword(&field.instruction).as_deref() == Some("TOC"))
        .collect())
}

/// The `HYPERLINK` fields that start in `paragraph`, in order.
pub(super) fn links(doc: &Doc, paragraph: &litchi_doc::Paragraph) -> Result<Vec<Link>> {
    let Some(range) = paragraph.cp_range() else {
        return Ok(Vec::new());
    };
    let mut links = Vec::new();
    let mut runs = None;
    for field in fields(doc)? {
        let start = field.field.start_cp;
        if start < range.start() || start >= range.end() {
            continue;
        }
        if field.field.field_type != litchi_doc::FieldType::Hyperlink {
            continue;
        }
        // Producers leave stray control characters (an object anchor, a
        // revision mark) in the instruction; the leaf parse reads the
        // instruction without them.
        let link = match field.hyperlink_field() {
            Some(link) => Some(link),
            None => {
                let mut clean = field.clone();
                clean.instruction = visible(&field.instruction).into_owned();
                clean.hyperlink_field()
            },
        };
        let Some(link) = link else {
            continue;
        };
        let owned = |value: Option<&str>| value.map(str::to_owned);
        let (url, anchor, tooltip, frame) = (
            owned(link.external_target()),
            owned(link.bookmark()),
            owned(link.screen_tip()),
            owned(link.target_frame()),
        );
        let run = match field.field.result_range() {
            Some((from, to)) if from < to => {
                if runs.is_none() {
                    runs = Some(paragraph.runs()?);
                }
                runs.as_ref().and_then(|runs| {
                    runs.iter()
                        .find(|run| {
                            run.cp_range()
                                .is_some_and(|span| span.start() >= from && span.start() < to)
                                && run.text().is_ok_and(|text| !visible(text).is_empty())
                        })
                        .cloned()
                })
            },
            _ => None,
        };
        links.push(Link {
            field,
            url,
            anchor,
            tooltip,
            frame,
            run,
        });
    }
    Ok(links)
}

/// The value at `index` of the hyperlink table.
pub(super) fn link_value<'a>(doc: &'a Doc, link: &'a Link, index: usize) -> Option<Value<'a>> {
    let text = |value: &'a Option<String>| {
        value
            .as_deref()
            .map(|value| Value::Text(Cow::Borrowed(value)))
    };
    let run = link.run.as_ref();
    match index {
        0 => text(&link.url),
        1 => text(&link.anchor),
        2 => text(&link.tooltip),
        3 => text(&link.frame),
        4 => run.and_then(|run| {
            let name = doc.parsed().font_name(run.font_index()?).ok()??;
            Some(Value::Text(Cow::Borrowed(name)))
        }),
        5 => run
            .and_then(litchi_doc::Run::font_size)
            .map(|size| Value::Length(Length::from_half_points(i32::from(size)))),
        6 => run.and_then(litchi_doc::Run::bold).map(Value::Bool),
        7 => run.and_then(litchi_doc::Run::italic).map(Value::Bool),
        8 => run
            .and_then(litchi_doc::Run::color)
            .map(|(red, green, blue)| Value::Color(Color::rgb(red, green, blue))),
        9 => run.and_then(|run| {
            run.underline()
                .map(|_| Value::Token(spec::underline_name(run.underline_style())))
        }),
        10 => run
            .and_then(litchi_doc::Run::strikethrough)
            .map(Value::Bool),
        11 => run.and_then(|run| super::parts::run_style(doc, run)),
        _ => None,
    }
}

/// The value at `index` of the field table.
pub(super) fn field_value(field: &FieldText, index: usize) -> Option<Value<'_>> {
    match index {
        0 => keyword(&field.instruction)
            .map(|keyword| Value::Text(Cow::Owned(keyword.to_ascii_lowercase()))),
        1 => Some(match visible(&field.instruction) {
            Cow::Borrowed(text) => Value::Text(Cow::Borrowed(text.trim())),
            Cow::Owned(text) => Value::Text(Cow::Owned(text.trim().to_owned())),
        }),
        2 => field.field.end_flags.locked.then_some(Value::Bool(true)),
        3 => Some(Value::Bool(field.result.is_some())),
        _ => None,
    }
}

/// Whether switch `name` is present in `instruction`, and its quoted
/// argument.
fn switch(instruction: &str, name: char) -> (bool, Option<String>) {
    let mut chars = instruction.char_indices();
    while let Some((at, character)) = chars.next() {
        if character != '\\' {
            continue;
        }
        let Some((_, found)) = chars.next() else {
            break;
        };
        if !found.eq_ignore_ascii_case(&name) {
            continue;
        }
        let rest = instruction[at + 2..].trim_start();
        let value = rest
            .strip_prefix('"')
            .and_then(|quoted| quoted.find('"').map(|end| quoted[..end].to_owned()));
        return (true, value);
    }
    (false, None)
}

/// The value at `index` of the table-of-contents table.
pub(super) fn toc_value(field: &FieldText, index: usize) -> Option<Value<'static>> {
    match index {
        0 => switch(&field.instruction, 'o')
            .1
            .map(|levels| Value::Text(Cow::Owned(levels))),
        1 => Some(Value::Bool(switch(&field.instruction, 'h').0)),
        2 => Some(Value::Bool(!switch(&field.instruction, 'n').0)),
        _ => None,
    }
}

/// The value at `index` of the bookmark table.
pub(super) fn bookmark_value(bookmark: &litchi_doc::Bookmark, index: usize) -> Option<Value<'_>> {
    match index {
        0 => Some(Value::Text(Cow::Borrowed(bookmark.name.as_str()))),
        _ => None,
    }
}

fn twips_extent(value: impl Into<i64>) -> Option<Value<'static>> {
    i32::try_from(value.into())
        .ok()
        .map(|value| Value::Length(Length::from_twips(value)))
}

/// The value at `index` of the section table.
pub(super) fn section_value(section: &Section, index: usize) -> Option<Value<'static>> {
    let page = &section.page;
    match index {
        0 => Some(Value::Token(match section.break_kind {
            BreakKind::Continuous => "continuous",
            BreakKind::NewColumn => "nextColumn",
            BreakKind::NewPage => "nextPage",
            BreakKind::EvenPage => "evenPage",
            BreakKind::OddPage => "oddPage",
        })),
        1 => twips_extent(page.width_twips),
        2 => twips_extent(page.height_twips),
        3 => Some(Value::Token(match page.orientation {
            PageOrientation::Portrait => "portrait",
            PageOrientation::Landscape => "landscape",
        })),
        4 => twips_extent(page.margins.top.signed_twips()),
        5 => twips_extent(page.margins.bottom.signed_twips()),
        6 => twips_extent(page.margins.left_twips),
        7 => twips_extent(page.margins.right_twips),
        8 => twips_extent(page.header_distance_twips),
        9 => twips_extent(page.footer_distance_twips),
        10 => twips_extent(page.margins.gutter_twips),
        11 => i64::try_from(section.columns.count()).ok().map(Value::Int),
        12 => section
            .columns
            .equal()
            .and_then(|(_, spacing, _)| twips_extent(spacing)),
        13 => section.columns.line_between().then_some(Value::Bool(true)),
        14 => section
            .behavior
            .different_first_page
            .then_some(Value::Bool(true)),
        15 => Some(Value::Text(Cow::Borrowed(super::parts::number_format(
            section.page_numbering.number_format,
        )))),
        16 => section
            .page_numbering
            .restart
            .then(|| Value::Int(i64::from(section.page_numbering.start_at))),
        17 => (section.line_numbering.interval > 0).then_some(Value::Token(
            match section.line_numbering.restart {
                LineNumberRestart::EachPage => "restartPage",
                LineNumberRestart::EachSection => "restartSection",
                LineNumberRestart::Continuous => "continuous",
            },
        )),
        18 => (section.line_numbering.interval > 0)
            .then(|| Value::Int(i64::from(section.line_numbering.interval))),
        19 => (section.line_numbering.interval > 0)
            .then(|| Value::Int(i64::from(section.line_numbering.distance_twips))),
        20 => section
            .behavior
            .right_to_left
            .then_some(Value::Token("rtl")),
        21 => section
            .behavior
            .right_to_left_gutter
            .then_some(Value::Bool(true)),
        22 => Some(Value::Token(match page.vertical_justification {
            VerticalJustification::Top => "top",
            VerticalJustification::Center => "center",
            VerticalJustification::Justified => "both",
            VerticalJustification::Bottom => "bottom",
        })),
        _ => None,
    }
}

/// The `ST_TabJc` spelling of a tab alignment.
pub(super) const fn tab_val(alignment: TabAlignment) -> &'static str {
    match alignment {
        TabAlignment::Left => "left",
        TabAlignment::Center => "center",
        TabAlignment::Right => "right",
        TabAlignment::Decimal => "decimal",
        TabAlignment::Bar => "bar",
        TabAlignment::List => "num",
    }
}

/// The `ST_TabTlc` spelling of a tab leader.
pub(super) const fn tab_leader(leader: TabLeader) -> &'static str {
    match leader {
        TabLeader::None | TabLeader::DefaultLeader => "none",
        TabLeader::Dots => "dot",
        TabLeader::Hyphens => "hyphen",
        TabLeader::Underline => "underscore",
        TabLeader::Heavy => "heavy",
        TabLeader::MiddleDot => "middleDot",
    }
}

/// The value at `index` of the tab table.
pub(super) fn tab_value(stop: &TabStop, index: usize) -> Option<Value<'static>> {
    match index {
        0 => Some(Value::Int(i64::from(stop.position))),
        1 => Some(Value::Token(tab_val(stop.alignment))),
        2 => (!matches!(stop.leader, TabLeader::None | TabLeader::DefaultLeader))
            .then(|| Value::Token(tab_leader(stop.leader))),
        _ => None,
    }
}

/// A main-story text box and the floating anchor that places it.
#[derive(Debug)]
pub(crate) struct TextBox {
    pub(crate) text_box: litchi_doc::parts::textbox::TextBox,
    anchor: Option<ShapeAnchor>,
}

/// The main story's text boxes, each with its anchor.
pub(super) fn text_boxes(doc: &Doc) -> Result<Vec<TextBox>> {
    let reader = doc.reader()?;
    let anchors = reader.shape_positions();
    Ok(reader
        .text_boxes()
        .into_iter()
        .map(|text_box| TextBox {
            anchor: anchors
                .iter()
                .find(|anchor| anchor.spa.shape_id == text_box.shape_id)
                .cloned(),
            text_box,
        })
        .collect())
}

fn width(anchor: &ShapeAnchor) -> Option<Value<'static>> {
    twips_extent(i64::from(anchor.spa.right) - i64::from(anchor.spa.left))
}

fn height(anchor: &ShapeAnchor) -> Option<Value<'static>> {
    twips_extent(i64::from(anchor.spa.bottom) - i64::from(anchor.spa.top))
}

/// The value at `index` of the text box table.
pub(super) fn text_box_value(text_box: &TextBox, index: usize) -> Option<Value<'static>> {
    let anchor = text_box.anchor.as_ref()?;
    match index {
        0 => width(anchor),
        1 => height(anchor),
        _ => None,
    }
}

/// The value at `index` of the shape table.
pub(super) fn shape_value(anchor: &ShapeAnchor, index: usize) -> Option<Value<'static>> {
    let spa = &anchor.spa;
    match index {
        0 => width(anchor),
        1 => height(anchor),
        2 => twips_extent(spa.left),
        3 => twips_extent(spa.top),
        4 => Some(Value::Token(match spa.wrap {
            ShapeTextWrap::Around | ShapeTextWrap::Square => "square",
            ShapeTextWrap::TopAndBottom => "topAndBottom",
            ShapeTextWrap::None => "none",
            ShapeTextWrap::Tight => "tight",
            ShapeTextWrap::Through => "through",
        })),
        5 => Some(Value::Token(match spa.horizontal_origin {
            ShapeHorizontalOrigin::Margin => "margin",
            ShapeHorizontalOrigin::Page => "page",
            ShapeHorizontalOrigin::Column => "column",
        })),
        6 => Some(Value::Token(match spa.vertical_origin {
            ShapeVerticalOrigin::Margin => "margin",
            ShapeVerticalOrigin::Page => "page",
            ShapeVerticalOrigin::Paragraph => "paragraph",
        })),
        7 => spa.below_text.then_some(Value::Bool(true)),
        _ => None,
    }
}
