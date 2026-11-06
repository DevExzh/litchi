//! The record-0323 element surface of legacy PowerPoint: slide comments,
//! table rows, cells and columns, hyperlinks, text runs, line breaks, the
//! theme and charts.
//!
//! Each element is a value read from the native presentation projection when
//! the node is resolved; the node keeps it so `format` and `text` answer
//! without another parse. Writes go through the native primary edit
//! (`format::ppt::write`).

use std::borrow::Cow;
use std::sync::Arc;

use litchi_ppt::ParsedComment;

use super::node::{Node, ShapeItem, length};
use super::spec::Kind;
use crate::Result;
use crate::node::{Preview, Value};

/// One record-0323 element.
#[derive(Clone)]
pub(crate) enum Element {
    /// A slide comment: its slide, its zero-based position among the slide's
    /// comments, and the parsed comment.
    Comment {
        slide: usize,
        index: usize,
        comment: Arc<ParsedComment>,
    },
    /// A table row: the table shape and the zero-based row.
    Row { shape: ShapeItem, row: usize },
    /// A table cell: the table shape, row and column.
    Cell {
        shape: ShapeItem,
        row: usize,
        column: usize,
    },
    /// A table column: the table shape and the zero-based column.
    Column { shape: ShapeItem, column: usize },
    /// A click hyperlink of a shape (`run: None`) or of a text range (the
    /// UTF-16 range it covers).
    Hyperlink {
        shape: ShapeItem,
        run: Option<(usize, usize)>,
        click: Arc<litchi_ppt::hyperlink::Click>,
    },
    /// A text run: a piece of one paragraph with one character format.
    Run {
        shape: ShapeItem,
        paragraph: usize,
        run: usize,
        body: Arc<litchi_ppt::text_edit::TextRuns>,
        piece: Piece,
    },
    /// A line break (`\u{b}`) of one paragraph.
    LineBreak { shape: ShapeItem },
    /// The presentation theme.
    Theme(Arc<litchi_ppt::theme::Theme>),
    /// An embedded chart of a slide.
    Chart {
        slide: usize,
        chart: Arc<ChartSummary>,
    },
    /// An OLE object or media picture frame read as its element kind
    /// (`ole`, `video`, `audio` or `media`), with an OLE object's `ProgID`
    /// and whether it is linked.
    Frame {
        shape: ShapeItem,
        kind: Kind,
        program: Option<Arc<str>>,
        linked: bool,
    },
}

/// The frames of `kind` (`ole`, `video`, `audio`, `media`) among the direct
/// pictures of a slide or group's shape array, in source order.
pub(crate) fn frames(
    file: &super::Ppt,
    direct: &[litchi_ppt::shapes::ShapeEnum<'static>],
    kind: Kind,
) -> Result<Vec<(usize, Option<Arc<str>>, bool)>> {
    use litchi_ppt::shapes::{PictureFrameKind, ShapeEnum};
    let wanted = match kind {
        Kind::Ole => PictureFrameKind::OleObject,
        _ => PictureFrameKind::Media,
    };
    let mut out = Vec::new();
    let mut objects = None;
    let mut media = None;
    for (index, shape) in direct.iter().enumerate() {
        let ShapeEnum::Picture(picture) = shape else {
            continue;
        };
        if picture.frame_kind() != wanted {
            continue;
        }
        let id = picture.external_object_id();
        match kind {
            Kind::Ole => {
                if objects.is_none() {
                    objects = Some(file.reader().ole_objects()?);
                }
                let definition =
                    id.and_then(|id| {
                        objects
                            .as_ref()?
                            .as_ref()?
                            .objects
                            .iter()
                            .find_map(|object| match object {
                                litchi_ppt::embedded::object::ExternalObject::Object(
                                    definition,
                                ) if object.id() == id => Some(definition),
                                _ => None,
                            })
                    });
                out.push((
                    index,
                    definition
                        .and_then(|definition| definition.program_id.as_deref())
                        .map(Arc::from),
                    definition.is_some_and(|definition| {
                        matches!(
                            definition.kind,
                            litchi_ppt::embedded::object::ContainerKind::Linked(_)
                        )
                    }),
                ));
            },
            Kind::Media => out.push((index, None, false)),
            _ => {
                if media.is_none() {
                    media = Some(file.reader().external_media()?);
                }
                let movie = id.is_some_and(|id| {
                    media
                        .as_ref()
                        .and_then(Option::as_ref)
                        .is_some_and(|collection| {
                            collection.objects.iter().any(|object| {
                                object.id() == id
                                    && matches!(
                                        object,
                                        litchi_ppt::external_media::Object::Movie(_)
                                    )
                            })
                        })
                });
                if movie == (kind == Kind::Video) {
                    out.push((index, None, false));
                }
            },
        }
    }
    Ok(out)
}

/// What a chart node reads: its frame and its first chart substream.
#[derive(Debug)]
pub(crate) struct ChartSummary {
    shape_id: u32,
    name: Option<String>,
    bounds: Option<(i32, i32, i32, i32)>,
    chart_type: Option<&'static str>,
    title: Option<String>,
}

/// The PPTX `chartType` name of an MS-OGRAPH chart group family.
fn chart_type(family: litchi_ppt::chart::semantic::Family) -> &'static str {
    use litchi_ppt::chart::semantic::Family;
    let stacked = |stacked: bool, percent: bool, name: &'static [&'static str; 3]| {
        if percent {
            name[2]
        } else if stacked {
            name[1]
        } else {
            name[0]
        }
    };
    match family {
        Family::Bar { flags, .. } if flags & 0x0001 != 0 => stacked(
            flags & 0x0002 != 0,
            flags & 0x0004 != 0,
            &["bar", "stackedBar", "percentStackedBar"],
        ),
        Family::Bar { flags, .. } => stacked(
            flags & 0x0002 != 0,
            flags & 0x0004 != 0,
            &["column", "stackedColumn", "percentStackedColumn"],
        ),
        Family::Line { flags } => stacked(
            flags & 0x0001 != 0,
            flags & 0x0002 != 0,
            &["line", "stackedLine", "percentStackedLine"],
        ),
        Family::Area { flags } => stacked(
            flags & 0x0001 != 0,
            flags & 0x0002 != 0,
            &["area", "stackedArea", "percentStackedArea"],
        ),
        Family::Pie { hole, .. } if hole > 0 => "doughnut",
        Family::Pie { .. } => "pie",
        Family::OfPie { options } => match options.kind() {
            litchi_ograph::record::of_pie::Kind::Pie => "pieOfPie",
            litchi_ograph::record::of_pie::Kind::Bar => "barOfPie",
        },
        Family::Scatter { flags, .. } if flags & 0x0001 != 0 => "bubble",
        Family::Scatter { .. } => "scatter",
        Family::Radar { .. } => "radar",
        Family::Surface { .. } => "surface",
    }
}

/// The charts of the slide at `slide`, in inventory order.
pub(crate) fn slide_charts(file: &super::Ppt, slide: usize) -> Result<Vec<ChartSummary>> {
    let reader = file.reader();
    let inventory = reader.charts()?;
    let shapes = match reader.slide_at(slide)? {
        Some(selected) => selected.shared_shapes()?,
        None => return Ok(Vec::new()),
    };
    let mut out = Vec::new();
    for chart in inventory.on_slide(slide + 1) {
        let Some(frame) = chart.info().frame() else {
            continue;
        };
        let shape = shapes.iter().find(|shape| shape.id() == frame.shape());
        let bounds = match shape {
            Some(litchi_ppt::shapes::ShapeEnum::Picture(picture)) => {
                let properties = picture.properties();
                Some((
                    properties.x,
                    properties.y,
                    properties.width,
                    properties.height,
                ))
            },
            _ => None,
        };
        let semantic = chart.semantic_chart(0).ok().flatten();
        out.push(ChartSummary {
            shape_id: frame.shape(),
            name: shape.and_then(super::node::shape_name).map(str::to_owned),
            bounds,
            chart_type: semantic
                .as_ref()
                .and_then(|chart| chart.groups().first())
                .map(|group| chart_type(group.family)),
            title: semantic.and_then(|chart| chart.title().map(str::to_owned)),
        });
    }
    Ok(out)
}

/// One run piece: its UTF-16 range in the body and its character format's
/// index among the body's runs.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Piece {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) format: usize,
}

/// The run pieces of one paragraph: its text between line breaks, cut at
/// character-run boundaries. A paragraph is the text between `\r` marks;
/// its mark belongs to no run, as a PPTX paragraph's end properties.
pub(crate) fn pieces(body: &litchi_ppt::text_edit::TextRuns, paragraph: usize) -> Vec<Piece> {
    let units = body.text().encode_utf16().collect::<Vec<_>>();
    let mut start = 0usize;
    for (index, part) in units.split(|unit| *unit == 0x000D).enumerate() {
        if index == paragraph {
            let end = start + part.len();
            let mut out = Vec::new();
            let mut segment = start;
            let breaks = part
                .iter()
                .enumerate()
                .filter(|(_, unit)| **unit == 0x000B)
                .map(|(index, _)| start + index)
                .chain(std::iter::once(end));
            for offset in breaks {
                for (format, run) in body.runs().iter().enumerate() {
                    let piece_start = segment.max(run.start);
                    let piece_end = offset.min(run.start + run.length);
                    if piece_start < piece_end {
                        out.push(Piece {
                            start: piece_start,
                            end: piece_end,
                            format,
                        });
                    }
                }
                segment = offset + 1;
            }
            return out;
        }
        start += part.len() + 1;
    }
    Vec::new()
}

/// The number of line breaks in one paragraph of `text`.
pub(crate) fn line_breaks(text: &str, paragraph: usize) -> usize {
    text.split('\r')
        .nth(paragraph)
        .map_or(0, |part| part.matches('\u{b}').count())
}

/// The values one element's `format` projection answers, in its property
/// table's order.
pub(crate) struct Values<'a> {
    values: Vec<Option<Value<'a>>>,
}

impl<'a> Values<'a> {
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        self.values.get(index).cloned().flatten()
    }
}

impl Element {
    pub(crate) const fn kind(&self) -> Kind {
        match self {
            Self::Comment { .. } => Kind::Comment,
            Self::Row { .. } => Kind::Row,
            Self::Cell { .. } => Kind::Cell,
            Self::Column { .. } => Kind::Col,
            Self::Hyperlink { .. } => Kind::Hyperlink,
            Self::Run { .. } => Kind::Run,
            Self::LineBreak { .. } => Kind::LineBreak,
            Self::Theme(_) => Kind::Theme,
            Self::Chart { .. } => Kind::Chart,
            Self::Frame { kind, .. } => *kind,
        }
    }

    pub(crate) const fn slide(&self) -> Option<usize> {
        match self {
            Self::Comment { slide, .. } => Some(*slide),
            Self::Row { shape, .. }
            | Self::Cell { shape, .. }
            | Self::Column { shape, .. }
            | Self::Hyperlink { shape, .. }
            | Self::Run { shape, .. }
            | Self::LineBreak { shape, .. }
            | Self::Frame { shape, .. } => Some(shape.slide),
            Self::Chart { slide, .. } => Some(*slide),
            Self::Theme(_) => None,
        }
    }

    pub(crate) const fn shape_item(&self) -> Option<&ShapeItem> {
        match self {
            Self::Comment { .. } => None,
            Self::Row { shape, .. }
            | Self::Cell { shape, .. }
            | Self::Column { shape, .. }
            | Self::Hyperlink { shape, .. }
            | Self::Run { shape, .. }
            | Self::LineBreak { shape, .. }
            | Self::Frame { shape, .. } => Some(shape),
            Self::Theme(_) | Self::Chart { .. } => None,
        }
    }

    pub(crate) fn text<'n>(&'n self, node: &'n Node<'_>) -> Result<Option<Cow<'n, str>>> {
        Ok(match self {
            Self::Comment { comment, .. } => Some(Cow::Borrowed(comment.text.as_str())),
            Self::Cell { row, column, .. } => table(node)?
                .cell(*row, *column)
                .map(|text| Cow::Owned(text.to_owned())),
            Self::Run { body, piece, .. } => Some(Cow::Owned(String::from_utf16_lossy(
                &body.text().encode_utf16().collect::<Vec<_>>()[piece.start..piece.end],
            ))),
            Self::Row { .. }
            | Self::Column { .. }
            | Self::Hyperlink { .. }
            | Self::LineBreak { .. }
            | Self::Frame { .. }
            | Self::Theme(_) => None,
            Self::Chart { chart, .. } => chart.title.as_deref().map(Cow::Borrowed),
        })
    }

    pub(crate) fn preview<'n>(&'n self, node: &'n Node<'_>) -> Result<Option<Preview<'n>>> {
        Ok(self.text(node)?.map(Preview::Prefix))
    }

    pub(crate) fn child_count(&self, node: &Node<'_>) -> Result<usize> {
        Ok(match self {
            Self::Row { .. } => table(node)?.columns(),
            Self::Comment { .. }
            | Self::Cell { .. }
            | Self::Column { .. }
            | Self::Hyperlink { .. }
            | Self::LineBreak { .. }
            | Self::Theme(_)
            | Self::Frame { .. }
            | Self::Chart { .. } => 0,
            Self::Run { .. } => usize::from(run_click(node)?.is_some()),
        })
    }

    pub(crate) fn values<'n>(&'n self, node: &'n Node<'_>) -> Result<Values<'n>> {
        let props = self.kind().spec().props;
        let values = match self {
            Self::Comment { comment, .. } => props
                .iter()
                .map(|prop| comment_value(comment, prop.name))
                .collect(),
            Self::Row { row, .. } => {
                let heights = table(node)?.row_heights();
                props
                    .iter()
                    .map(|prop| match prop.name {
                        "height" => heights.get(*row).copied().map(length),
                        _ => None,
                    })
                    .collect()
            },
            Self::Column { column, .. } => {
                let widths = table(node)?.column_widths();
                props
                    .iter()
                    .map(|prop| match prop.name {
                        "width" => widths.get(*column).copied().map(length),
                        _ => None,
                    })
                    .collect()
            },
            Self::Cell { .. } => Vec::new(),
            Self::LineBreak { .. } => Vec::new(),
            Self::Frame {
                program, linked, ..
            } => {
                let picture = node.shape();
                let bounds = node.anchor_bounds();
                props
                    .iter()
                    .map(|prop| match prop.name {
                        "id" => picture.map(|shape| Value::Int(i64::from(shape.id()))),
                        "name" => picture
                            .and_then(super::node::shape_name)
                            .map(|name| Value::Text(Cow::Borrowed(name))),
                        "progId" => program
                            .as_deref()
                            .map(|program| Value::Text(Cow::Borrowed(program))),
                        "linked" => program.as_ref().map(|_| Value::Bool(*linked)),
                        "x" => bounds.map(|bounds| length(bounds.0)),
                        "y" => bounds.map(|bounds| length(bounds.1)),
                        "width" => bounds.map(|bounds| length(bounds.2)),
                        "height" => bounds.map(|bounds| length(bounds.3)),
                        _ => None,
                    })
                    .collect()
            },
            Self::Chart { chart, .. } => props
                .iter()
                .map(|prop| match prop.name {
                    "id" => Some(Value::Int(i64::from(chart.shape_id))),
                    "name" => chart
                        .name
                        .as_deref()
                        .map(|name| Value::Text(Cow::Borrowed(name))),
                    "x" => chart.bounds.map(|bounds| length(bounds.0)),
                    "y" => chart.bounds.map(|bounds| length(bounds.1)),
                    "width" => chart.bounds.map(|bounds| length(bounds.2)),
                    "height" => chart.bounds.map(|bounds| length(bounds.3)),
                    "chartType" => chart.chart_type.map(Value::Token),
                    "title" => chart
                        .title
                        .as_deref()
                        .map(|title| Value::Text(Cow::Borrowed(title))),
                    _ => None,
                })
                .collect(),
            Self::Theme(theme) => props
                .iter()
                .map(|prop| theme_value(theme, prop.name))
                .collect(),
            Self::Run { body, piece, .. } => {
                let format = &body.runs()[piece.format].format;
                let font = match format.font_ref() {
                    Some(index) => node
                        .file()
                        .reader()
                        .fonts()?
                        .get_base(index)
                        .map(|font| font.name.clone()),
                    None => None,
                };
                props
                    .iter()
                    .map(|prop| run_value(format, font.as_deref(), prop.name))
                    .collect()
            },
            Self::Hyperlink { click, .. } => props
                .iter()
                .map(|prop| match prop.name {
                    "link" => Some(Value::Text(link_text(&click.target))),
                    "tooltip" => click
                        .tooltip
                        .as_deref()
                        .map(|tooltip| Value::Text(Cow::Borrowed(tooltip))),
                    _ => None,
                })
                .collect(),
        };
        Ok(Values { values })
    }
}

fn comment_value<'a>(comment: &'a ParsedComment, key: &str) -> Option<Value<'a>> {
    match key {
        "text" => Some(Value::Text(Cow::Borrowed(comment.text.as_str()))),
        "author" => Some(Value::Text(Cow::Borrowed(comment.author.as_str()))),
        "initials" => Some(Value::Text(Cow::Borrowed(comment.initials.as_str()))),
        "index" => Some(Value::Int(i64::from(comment.index))),
        "date" => comment_date(comment).map(|date| Value::Text(Cow::Owned(date))),
        "x" => Some(length(comment.x)),
        "y" => Some(length(comment.y)),
        _ => None,
    }
}

/// The comment's `SYSTEMTIME` as `YYYY-MM-DDTHH:MM:SS.fff`; an all-zero date
/// (no creation time) reads as absent.
fn comment_date(comment: &ParsedComment) -> Option<String> {
    (comment.year != 0).then(|| {
        litchi_ppt::comments::DateTime {
            year: comment.year,
            month: comment.month,
            day: comment.day,
            hour: comment.hour,
            minute: comment.minute,
            second: comment.second,
            millisecond: comment.millisecond,
        }
        .to_string()
    })
}

/// A slide's comments, read from the native slide.
pub(crate) fn slide_comments(file: &super::Ppt, slide: usize) -> Result<Vec<ParsedComment>> {
    let selected = file.reader().slide_at(slide)?.ok_or_else(|| {
        crate::Error::new(
            crate::Code::InternalError,
            format!("slide {} vanished from the catalog", slide + 1),
        )
        .with_format(litchi_core::FileFormat::Ppt)
    })?;
    Ok(selected.comments()?)
}

/// The native table behind a table, row, cell or column node.
pub(crate) fn table<'n>(
    node: &'n Node<'_>,
) -> Result<&'n litchi_ppt::shapes::shape_enum::TableShape> {
    node.shape()
        .and_then(litchi_ppt::shapes::ShapeEnum::as_table)
        .ok_or_else(|| {
            crate::Error::new(
                crate::Code::InternalError,
                "table vanished from the native scene",
            )
            .with_format(litchi_core::FileFormat::Ppt)
        })
}

/// The PPTX `link` readback of a click target.
pub(crate) fn link_text(target: &litchi_ppt::hyperlink::Target) -> Cow<'_, str> {
    use litchi_ppt::hyperlink::{InteractionJump, Target};
    match target {
        Target::Url(url) => Cow::Borrowed(url.as_str()),
        Target::Slide(position) => Cow::Owned(format!("slide[{}]", position.get() + 1)),
        Target::Jump(jump) => Cow::Borrowed(match jump {
            InteractionJump::FirstSlide => "firstslide",
            InteractionJump::LastSlide => "lastslide",
            InteractionJump::NextSlide => "nextslide",
            InteractionJump::PreviousSlide => "previousslide",
            InteractionJump::EndShow => "endshow",
            InteractionJump::LastSlideViewed => "lastslideviewed",
            InteractionJump::None => "none",
        }),
    }
}

/// The click hyperlink of a shape node, or `None`.
pub(crate) fn shape_click(node: &Node<'_>) -> Result<Option<litchi_ppt::hyperlink::Click>> {
    let (Some(shape), Some(item)) = (node.shape(), node.shape_item()) else {
        return Ok(None);
    };
    if item.kind != Kind::Shape {
        return Ok(None);
    }
    Ok(node.file().reader().shape_click(item.slide, shape.id())?)
}

fn run_value<'a>(
    format: &litchi_ppt::text_format_exception::TextCFException,
    font: Option<&str>,
    key: &str,
) -> Option<Value<'a>> {
    let style = format.font_style();
    let flag = |mask: u32, value: fn(&litchi_ppt::text_format_exception::CFStyle) -> bool| {
        (format.masks() & mask != 0).then(|| style.as_ref().is_some_and(value))
    };
    match key {
        "font.latin" => font.map(|font| Value::Text(Cow::Owned(font.to_owned()))),
        "size" => format
            .font_size()
            .map(|size| Value::Length(crate::value::Length::from_points(i32::from(size)))),
        "bold" => flag(0x0001, litchi_ppt::text_format_exception::CFStyle::bold).map(Value::Bool),
        "italic" => {
            flag(0x0002, litchi_ppt::text_format_exception::CFStyle::italic).map(Value::Bool)
        },
        "underline" => flag(
            0x0004,
            litchi_ppt::text_format_exception::CFStyle::underline,
        )
        .map(|on| Value::Token(if on { "single" } else { "none" })),
        "color" => format.color().and_then(|color| {
            (color.kind == litchi_ppt::slide_show_settings::ColorIndexKind::Srgb)
                .then(|| Value::Color(crate::value::Color::rgb(color.red, color.green, color.blue)))
        }),
        _ => None,
    }
}

/// The text runs of a shape node's text body, or `None`.
pub(crate) fn shape_runs(node: &Node<'_>) -> Result<Option<litchi_ppt::text_edit::TextRuns>> {
    let (Some(shape), Some(item)) = (node.shape(), node.shape_item()) else {
        return Ok(None);
    };
    if item.kind != Kind::Shape {
        return Ok(None);
    }
    let Some(slide) = node.file().reader().slide_at(item.slide)? else {
        return Ok(None);
    };
    Ok(slide.shape_text_runs(shape.id())?)
}

/// The click hyperlink of a run node: the text link whose range holds the
/// run's start, with that range.
pub(crate) fn run_click(
    node: &Node<'_>,
) -> Result<Option<((usize, usize), litchi_ppt::hyperlink::Click)>> {
    let Some(Element::Run { body, piece, .. }) = node.element_item() else {
        return Ok(None);
    };
    for link in body.links() {
        if link.start <= piece.start && piece.start < link.end {
            if let Some(click) = node.file().reader().resolve_click(&link.interaction)? {
                return Ok(Some(((link.start, link.end), click)));
            }
        }
    }
    Ok(None)
}

fn theme_value<'a>(theme: &'a litchi_ppt::theme::Theme, key: &str) -> Option<Value<'a>> {
    let slot = match key {
        "hyperlink" => Some("hlink"),
        "followedhyperlink" => Some("folHlink"),
        other if litchi_ppt::theme::COLOR_SLOTS.contains(&other) => Some(other),
        _ => None,
    };
    if let Some(slot) = slot {
        return theme
            .color(slot)
            .map(|[red, green, blue]| Value::Color(crate::value::Color::rgb(red, green, blue)));
    }
    let text = |value: &'a Option<String>| {
        value
            .as_deref()
            .map(|text| Value::Text(Cow::Borrowed(text)))
    };
    match key {
        "headingFont" => text(&theme.major.latin),
        "bodyFont" => text(&theme.minor.latin),
        "headingFont.ea" => text(&theme.major.east_asian),
        "bodyFont.ea" => text(&theme.minor.east_asian),
        "headingFont.cs" => text(&theme.major.complex),
        "bodyFont.cs" => text(&theme.minor.complex),
        "name" => text(&theme.name),
        _ => None,
    }
}
