//! What a DOCX node shows in its view and offers to queries: its text,
//! preview, style, child count and `format` entries (record 0010 §2.3,
//! §2.4, §4.5.3).
//!
//! `format` reads the node's typed direct properties once (L-DOCX-R4:
//! `Paragraph::direct_properties`, `Run::direct_properties`,
//! `Table::properties`, `Row::properties`, `Cell::properties`) and yields
//! the entries by table index. Values are local, never effective: nothing
//! is inherited from styles or document defaults (ADR 0004). An
//! explicitly-off flag is `false`, an unset one absent; theme colours stay
//! theme references. Each value is printed the way OfficeCLI prints its key
//! (`12pt`, `1.15x`, `#1F4E79`, `"2880dxa"`).

use std::borrow::Cow;
use std::fmt::Write as _;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_docx::color::Theme;
use litchi_docx::paragraph::LineSpacingRule;
use litchi_docx::properties::{
    self as direct, HexColor, Indentation, Justification, Merge, TableJustification, TableLayout,
    VerticalAlign, Width,
};
use litchi_docx::{InlineHyperlink, RunUnderlineColor};

use super::nav::Nav;
use super::node::{Node, Value};
use super::spec::{
    CELL_PROPS, DOCUMENT_PROPS, HYPERLINK_PROPS, Kind, PARAGRAPH_PROPS, ROW_PROPS, RUN_PROPS,
    SDT_PROPS, TABLE_PROPS, index_of,
};
use crate::error::Result;
use crate::node::{
    Children, Format, NodePath, Preview, Project, Source, Value as NodeValue, ViewOptions,
};
use crate::path::navigate::{KindSet, Navigate};
use crate::value::{Color, Length, Scheme};

/// The node's text: a paragraph's, run's, cell's, content control's or
/// hyperlink's visible text; `None` for the root, the body, tables, rows and
/// section breaks, as OfficeCLI's node builders leave it.
pub(super) fn text<'n>(node: &'n Node<'_>) -> Result<Option<Cow<'n, str>>> {
    Ok(match &node.value {
        Value::Paragraph { paragraph, .. } => Some(Cow::Owned(paragraph.text()?)),
        Value::Run(run) => Some(Cow::Owned(run.text()?)),
        Value::Cell(cell) => Some(Cow::Owned(cell.text()?)),
        Value::Hyperlink(link) => Some(Cow::Borrowed(link.link().text())),
        Value::Sdt(block) => match block.indexed()?.content_control() {
            Some(control) => Some(Cow::Owned(control.text()?)),
            None => None,
        },
        Value::Equation(block) => equation_latex(block)?.map(Cow::Owned),
        Value::Comment(comment) => Some(Cow::Owned(comment_text(comment)?)),
        Value::SmartArt(smartart) => Some(Cow::Owned(super::smartart::smartart_text(smartart))),
        Value::Part(value) => super::parts::text(node.docx, *value)?.map(Cow::Owned),
        Value::Mark(mark) => super::marks::text(mark).map(Cow::Owned),
        Value::DiagramNode(node) => super::smartart::node_text(node).map(Cow::Owned),
        Value::Picture(_) | Value::Chart(_) | Value::ChartChild(..) | Value::Column(..) => None,
        Value::Document
        | Value::Body(_)
        | Value::Table(_)
        | Value::Row(_)
        | Value::Section(_)
        | Value::Comments => None,
    })
}

/// The LaTeX of a display equation, which OfficeCLI's equation node holds
/// as its text (`FormulaParser.ToLatex`). `None` without the `formula`
/// feature, and for an equation that does not convert: the read does not
/// fail for it.
pub(super) fn equation_latex(block: &super::node::Block) -> Result<Option<String>> {
    #[cfg(feature = "formula")]
    {
        let indexed = block.indexed()?;
        let formulas = match indexed.paragraph() {
            Some(paragraph) => paragraph.paragraph_level_formulas()?,
            None => vec![String::from_utf8_lossy(&indexed.self_contained_xml()?).into_owned()],
        };
        let mut latex = String::new();
        for formula in formulas {
            let Ok(converted) =
                litchi_formula::omml_to_latex(&formula, &litchi_formula::omml::Limits::DEFAULT)
            else {
                return Ok(None);
            };
            // The converter wraps a math paragraph in display delimiters;
            // OfficeCLI's node text is the bare expression.
            let trimmed = converted.trim();
            let bare = trimmed
                .strip_prefix("\\[")
                .and_then(|inner| inner.strip_suffix("\\]"))
                .or_else(|| {
                    trimmed
                        .strip_prefix("$$")
                        .and_then(|inner| inner.strip_suffix("$$"))
                })
                .unwrap_or(trimmed)
                .trim();
            if !latex.is_empty() {
                latex.push(' ');
            }
            latex.push_str(bare);
        }
        Ok((!latex.is_empty()).then_some(latex))
    }
    #[cfg(not(feature = "formula"))]
    {
        let _ = block;
        Ok(None)
    }
}

/// The `preview` member: a paragraph's text cut to 50 characters, as
/// OfficeCLI previews Word paragraphs.
pub(super) fn preview<'n>(node: &'n Node<'_>) -> Result<Option<Preview<'n>>> {
    Ok(match &node.value {
        Value::Paragraph { paragraph, .. } => Some(Preview::Prefix(Cow::Owned(paragraph.text()?))),
        _ => None,
    })
}

/// The `style` member of a paragraph: its style id, or `Normal` when it
/// names none (OfficeCLI prints the style's display name, which needs the
/// styles part; M1-Δ4).
pub(super) fn style<'n>(node: &'n Node<'_>) -> Result<Option<Cow<'n, str>>> {
    Ok(match &node.value {
        Value::Paragraph { paragraph, .. } => Some(Cow::Owned(
            paragraph.style_id()?.unwrap_or_else(|| "Normal".to_owned()),
        )),
        _ => None,
    })
}

/// The `childCount` member: the children the view lists, or `None` for a
/// body whose blocks no request has read (M1-Δ1).
pub(super) fn child_count(node: &Node<'_>) -> Result<Option<u32>> {
    let count = match &node.value {
        Value::Document => {
            1 + super::parts::root_count(node.docx)?
                + usize::from(super::nav::has_comments(node.docx)?)
        },
        Value::Part(value) => super::parts::child_count(node.docx, *value)?,
        Value::Mark(_) => 0,
        Value::Comments => super::nav::comments(node.docx)?.len(),
        Value::Body(lock) => match lock.get() {
            Some(blocks) => listed_blocks(blocks),
            None => return Ok(None),
        },
        Value::Cell(cell) => listed_blocks(&cell.blocks()?),
        Value::Sdt(block) => match block.indexed()?.content_control() {
            Some(control) => listed_blocks(&control.blocks()?),
            None => 0,
        },
        Value::Table(block) => match block.indexed()?.table() {
            Some(table) => table.row_count()?,
            None => 0,
        },
        Value::Row(row) => row.cell_count()?,
        Value::Paragraph { paragraph, .. } => paragraph.addressable_runs()?.len(),
        Value::Chart(chart) => super::chart::child_count(node.docx, chart)?,
        Value::SmartArt(_) | Value::DiagramNode(_) => {
            super::smartart::node_count(node.docx, &node.value)?
        },
        Value::Run(_)
        | Value::Hyperlink(_)
        | Value::Equation(_)
        | Value::Section(_)
        | Value::Comment(_)
        | Value::Picture(_)
        | Value::ChartChild(..)
        | Value::Column(..) => 0,
    };
    Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
}

/// The blocks of a container the view lists: every block but `w:altChunk`
/// anchors and unknown children.
fn listed_blocks(blocks: &litchi_docx::Blocks) -> usize {
    blocks
        .iter()
        .filter(|block| Kind::of_block(block.kind()).is_some())
        .count()
}

/// The node's `format` entries.
pub(super) fn format<'n>(node: &'n Node<'_>) -> Result<Format<'n>> {
    let props = match &node.value {
        Value::Document => Props::Document(
            Box::new(node.docx.metadata()?.unwrap_or_default()),
            node.docx.app_props()?.map(Box::new),
        ),
        Value::Paragraph { paragraph, .. } => Props::Paragraph {
            docx: node.docx,
            props: Box::new(paragraph.direct_properties()?),
            id: litchi_docx::document::raw::paragraph_ids(paragraph)?
                .para_id()
                .map(|id| format!("{:08X}", id.get())),
        },
        Value::Run(run) => Props::Run(Box::new(run.direct_properties()?)),
        Value::Hyperlink(link) => {
            // OfficeCLI reads the formatting of the first run that holds a
            // `w:t`.
            let mut first = None;
            for run in link.runs() {
                let holds_text = run
                    .contents()?
                    .iter()
                    .any(|content| matches!(content, litchi_docx::RunContent::Text(_)));
                if holds_text {
                    first = Some(Box::new(run.direct_properties()?));
                    break;
                }
            }
            Props::Hyperlink { link, first }
        },
        Value::Table(block) => match block.indexed()?.table() {
            Some(table) => {
                let props = table.properties()?;
                let rows = table.rows()?;
                let columns = match &props.grid {
                    Some(grid) => grid.len(),
                    None => match rows.first() {
                        Some(row) => row.cell_count()?,
                        None => 0,
                    },
                };
                Props::Table(Box::new(TableProps {
                    props,
                    rows: rows.len(),
                    columns,
                }))
            },
            None => Props::None,
        },
        Value::Row(row) => Props::Row(Box::new(row.properties()?)),
        Value::Cell(cell) => Props::Cell(Box::new(cell.properties()?)),
        Value::Sdt(block) => match block.indexed()?.content_control() {
            Some(control) => Props::Sdt(Box::new(SdtProps {
                alias: control.alias()?.map(str::to_owned),
                tag: control.tag()?.map(str::to_owned),
                id: litchi_docx::document::raw::content_control_id(&control)?,
            })),
            None => Props::None,
        },
        Value::Equation(_) => Props::Equation,
        Value::Comment(comment) => {
            let (done, parent) = comment_thread(node.docx, comment)?;
            Props::Comment(
                Box::new(CommentProps {
                    id: comment.id(),
                    author: comment.author().to_owned(),
                    initials: comment.initials().map(str::to_owned),
                    date: comment.date().map(str::to_owned),
                    text: comment_text(comment)?,
                    done,
                    parent,
                }),
                node.docx,
            )
        },
        Value::Picture(picture) => Props::Picture(Arc::clone(picture)),
        Value::Chart(chart) => Props::Chart(node.docx, Arc::clone(chart)),
        Value::SmartArt(smartart) => Props::SmartArt(Arc::clone(smartart)),
        Value::ChartChild(chart, child) => {
            Props::ChartChild(node.docx, Arc::clone(chart), node.kind, *child)
        },
        Value::DiagramNode(value) => Props::DiagramNode(Arc::clone(value)),
        Value::Part(value) => Props::Part(super::parts::props(node.docx, *value)?),
        Value::Mark(mark) => Props::Mark(Arc::clone(mark)),
        Value::Column(block, column) => Props::Column(super::column::width(block, *column)?),
        Value::Body(_) | Value::Section(_) | Value::Comments => Props::None,
    };
    let invalid = match &props {
        Props::Paragraph { props, .. } => {
            props.invalid_properties().next().is_some()
                || props
                    .mark
                    .as_ref()
                    .is_some_and(|mark| mark.invalid_properties().next().is_some())
        },
        Props::Run(props) => props.invalid_properties().next().is_some(),
        Props::Hyperlink { first, .. } => first
            .as_ref()
            .is_some_and(|run| run.invalid_properties().next().is_some()),
        Props::Table(table) => table.props.invalid_properties().next().is_some(),
        Props::Row(props) => props.invalid_properties().next().is_some(),
        Props::Cell(props) => props.invalid_properties().next().is_some(),
        _ => false,
    };
    require_valid(invalid)?;
    // A part element's semantic identity (`type`) is an ordinary property;
    // its native identifiers (`id`, `ilvl`) are raw-only, as list ids are
    // for RTF (record 0319).
    if let Props::Mark(_) = &props {
        return Ok(Format::new(
            super::marks::table(node.kind),
            Source::Docx(props),
        ));
    }
    if let Props::Part(_) = &props {
        let hidden: &'static [&'static str] = match node.kind {
            Kind::AbstractNum | Kind::Num => &["id"],
            Kind::Level => &["ilvl"],
            _ => &[],
        };
        return Ok(
            Format::new(super::parts::table(node.kind), Source::Docx(props)).without(hidden),
        );
    }
    let table = node.kind.spec().props;
    // An axis's `role` is semantic: it selects and is shown (record 0335).
    let hidden = if node.kind == Kind::Axis {
        &[]
    } else {
        node.kind.spec().identities
    };
    Ok(Format::new(table, Source::Docx(props)).without(hidden))
}

/// Refuse malformed typed direct properties instead of reporting them as absent.
pub(super) fn require_valid(invalid: bool) -> Result<()> {
    if invalid {
        return Err(crate::Error::new(
            crate::Code::CorruptFile,
            "a DOCX direct property has an invalid value",
        )
        .with_format(litchi_core::FileFormat::Docx));
    }
    Ok(())
}

/// A table's direct properties with its row and column counts.
pub(crate) struct TableProps {
    props: direct::Table,
    rows: usize,
    columns: usize,
}

/// A comment's identity, author and text, and its `commentsExtended`
/// state and reply parent (record 0345).
pub(crate) struct CommentProps {
    id: u32,
    author: String,
    initials: Option<String>,
    date: Option<String>,
    text: String,
    done: bool,
    parent: Option<u32>,
}

/// A content control's identity.
pub(crate) struct SdtProps {
    alias: Option<String>,
    tag: Option<String>,
    id: Option<i32>,
}

/// A DOCX node's parsed property set, read once per `format` request; the
/// node view's and the query walker's source of `format` values.
pub(crate) enum Props<'a> {
    /// No readable properties.
    None,
    /// A display equation: `mode` is `display`.
    Equation,
    /// The core properties and, when the part is present and valid, the
    /// application properties (record 0342).
    Document(
        Box<Metadata>,
        Option<Box<litchi_ooxml_common::properties::app::Props>>,
    ),
    Paragraph {
        /// The handle the paragraph's style and list properties read the
        /// styles and numbering parts through (record 0319).
        docx: &'a super::Docx,
        props: Box<direct::Paragraph>,
        /// Validated native identifier, available only to raw query access.
        id: Option<String>,
    },
    Run(Box<direct::Run>),
    Hyperlink {
        link: &'a InlineHyperlink,
        /// The direct properties of the link's first run that holds text.
        first: Option<Box<direct::Run>>,
    },
    Table(Box<TableProps>),
    Row(Box<direct::Row>),
    Cell(Box<direct::Cell>),
    Sdt(Box<SdtProps>),
    /// A comment, with the handle its anchor is found through.
    Comment(Box<CommentProps>, &'a super::Docx),
    Picture(Arc<litchi_docx::image::InlineImage>),
    /// A chart: its frame, and its chart part read through the handle.
    Chart(&'a super::Docx, Arc<super::chart::ChartValue>),
    /// A SmartArt graphic (record 0295).
    SmartArt(Arc<super::smartart::SmartArtValue>),
    /// A chart's series or axis (record 0335).
    ChartChild(&'a super::Docx, Arc<super::chart::ChartValue>, Kind, u32),
    /// A graphic's content node (record 0335).
    DiagramNode(Arc<super::smartart::DiagramNodeValue>),
    /// A styles, numbering, story or note element (record 0319).
    Part(super::parts::PartProps<'a>),
    /// A tab stop, break, bookmark, field, table of contents, watermark or
    /// text box (record 0319).
    Mark(Arc<super::marks::MarkValue>),
    /// A table grid column's width in twips (record 0345).
    Column(Option<u32>),
}

impl<'a> Props<'a> {
    /// The value of the property at `index` in the node kind's table.
    pub(crate) fn get(&self, index: usize) -> Option<NodeValue<'a>> {
        match self {
            Self::None => None,
            Self::Equation => (index == 0).then_some(NodeValue::Token("display")),
            Self::Document(metadata, app) => {
                document(metadata, index).or_else(|| extended(app.as_deref()?, index))
            },
            Self::Paragraph { docx, props, id } => match index {
                P_STYLE_NAME | P_NUM_FMT | P_LIST_STYLE | P_START => {
                    super::parts::paragraph_part_value(docx, props, index)
                },
                _ => paragraph_value(props, id.as_deref(), index),
            },
            Self::Run(props) => run(props, index),
            Self::Hyperlink { link, first } => hyperlink(link, first.as_deref(), index),
            Self::Table(table) => table_value(table, index),
            Self::Row(props) => row(props, index),
            Self::Cell(props) => cell(props, index),
            Self::Sdt(sdt) => sdt_value(sdt, index),
            Self::Comment(comment, docx) => comment_value(docx, comment, index),
            Self::Picture(picture) => picture_value(picture, index),
            Self::Chart(docx, chart) => super::chart::chart_value(docx, chart, index),
            Self::SmartArt(smartart) => super::smartart::smartart_value(smartart, index),
            Self::ChartChild(docx, chart, Kind::Series, child) => {
                super::chart::series_value(docx, chart, *child, index)
            },
            Self::ChartChild(docx, chart, _, child) => {
                super::chart::axis_value(docx, chart, *child, index)
            },
            Self::DiagramNode(node) => super::smartart::node_value(node, index),
            Self::Part(part) => part.get(index),
            Self::Mark(mark) => super::marks::value(mark, index),
            Self::Column(width) => super::column::column_value(*width, index),
        }
    }
}

impl std::fmt::Debug for Props<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::None => "Props::None",
            Self::Equation => "Props::Equation",
            Self::Document(..) => "Props::Document",
            Self::Paragraph { .. } => "Props::Paragraph",
            Self::Run(_) => "Props::Run",
            Self::Hyperlink { .. } => "Props::Hyperlink",
            Self::Table(_) => "Props::Table",
            Self::Row(_) => "Props::Row",
            Self::Cell(_) => "Props::Cell",
            Self::Sdt(_) => "Props::Sdt",
            Self::Comment(..) => "Props::Comment",
            Self::Picture(_) => "Props::Picture",
            Self::Chart(..) => "Props::Chart",
            Self::SmartArt(_) => "Props::SmartArt",
            Self::ChartChild(..) => "Props::ChartChild",
            Self::DiagramNode(_) => "Props::DiagramNode",
            Self::Part(_) => "Props::Part",
            Self::Mark(_) => "Props::Mark",
            Self::Column(_) => "Props::Column",
        })
    }
}

fn owned<'a>(text: &str) -> NodeValue<'a> {
    NodeValue::Text(Cow::Owned(text.to_owned()))
}

fn text_of<'a>(text: Option<&String>) -> Option<NodeValue<'a>> {
    text.filter(|text| !text.is_empty()).map(|text| owned(text))
}

fn twips(value: i64) -> Option<NodeValue<'static>> {
    i32::try_from(value)
        .ok()
        .map(|twips| NodeValue::Length(Length::from_twips(twips)))
}

fn half_points(value: i64) -> Option<NodeValue<'static>> {
    i32::try_from(value)
        .ok()
        .map(|half| NodeValue::Length(Length::from_half_points(half)))
}

/// A decimal with at most `places` fractional digits, trailing zeros
/// dropped, as .NET's `0.##` and `0.####` formats write it.
fn decimal(value: f64, places: i32) -> String {
    let scale = 10_f64.powi(places);
    let rounded = (value * scale).round() / scale;
    let mut text = format!("{rounded:.*}", usize::try_from(places).unwrap_or(0));
    if text.contains('.') {
        while text.ends_with('0') {
            text.pop();
        }
        if text.ends_with('.') {
            text.pop();
        }
    }
    if text == "-0" {
        text = "0".to_owned();
    }
    text
}

const fn scheme(theme: Theme) -> Option<Scheme> {
    Some(match theme {
        Theme::Dark1 => Scheme::Dark1,
        Theme::Light1 => Scheme::Light1,
        Theme::Dark2 => Scheme::Dark2,
        Theme::Light2 => Scheme::Light2,
        Theme::Accent1 => Scheme::Accent1,
        Theme::Accent2 => Scheme::Accent2,
        Theme::Accent3 => Scheme::Accent3,
        Theme::Accent4 => Scheme::Accent4,
        Theme::Accent5 => Scheme::Accent5,
        Theme::Accent6 => Scheme::Accent6,
        Theme::Hyperlink => Scheme::Hyperlink,
        Theme::FollowedHyperlink => Scheme::FollowedHyperlink,
        Theme::Background1 => Scheme::Background1,
        Theme::Text1 => Scheme::Text1,
        Theme::Background2 => Scheme::Background2,
        Theme::Text2 => Scheme::Text2,
        Theme::None => return None,
    })
}

const fn hex_color(value: HexColor) -> Color {
    match value {
        HexColor::Auto => Color::AUTO,
        HexColor::Rgb([red, green, blue]) => Color::rgb(red, green, blue),
    }
}

/// A run colour as OfficeCLI's `StyleColorWithThemeTail` prints it: the
/// value, a bare theme name when the colour is `auto` with an untransformed
/// theme, or the value with a `;themeColor=…;themeShade=…;themeTint=…`
/// tail when a theme reference accompanies an explicit value or carries a
/// transform.
fn word_color<'a>(color: &direct::Color) -> NodeValue<'a> {
    let theme = color.theme.filter(|theme| *theme != Theme::None);
    let transformed = color.theme_shade.is_some() || color.theme_tint.is_some();
    match (theme, color.value) {
        (None, value) if !transformed => NodeValue::Color(hex_color(value)),
        (Some(theme), HexColor::Auto) if !transformed => match scheme(theme) {
            Some(scheme) => NodeValue::Color(Color::scheme(scheme)),
            None => NodeValue::Color(Color::AUTO),
        },
        (theme, value) => {
            let mut tail = hex_color(value).to_string();
            if let Some(theme) = theme {
                let _ = write!(tail, ";themeColor={}", theme.as_str());
            }
            if let Some(shade) = color.theme_shade {
                let _ = write!(tail, ";themeShade={shade:02X}");
            }
            if let Some(tint) = color.theme_tint {
                let _ = write!(tail, ";themeTint={tint:02X}");
            }
            NodeValue::Text(Cow::Owned(tail))
        },
    }
}

// ---------------------------------------------------------------------------
// Document
// ---------------------------------------------------------------------------

const D_TITLE: usize = index_of(DOCUMENT_PROPS, "title");
const D_AUTHOR: usize = index_of(DOCUMENT_PROPS, "author");
const D_SUBJECT: usize = index_of(DOCUMENT_PROPS, "subject");
const D_KEYWORDS: usize = index_of(DOCUMENT_PROPS, "keywords");
const D_DESCRIPTION: usize = index_of(DOCUMENT_PROPS, "description");
const D_CATEGORY: usize = index_of(DOCUMENT_PROPS, "category");
const D_LAST_MODIFIED_BY: usize = index_of(DOCUMENT_PROPS, "lastModifiedBy");
const D_REVISION: usize = index_of(DOCUMENT_PROPS, "revisionNumber");
const D_CREATED: usize = index_of(DOCUMENT_PROPS, "created");
const D_MODIFIED: usize = index_of(DOCUMENT_PROPS, "modified");

fn document<'a>(metadata: &Metadata, index: usize) -> Option<NodeValue<'a>> {
    let text = |value: &Option<String>| value.as_deref().map(owned);
    match index {
        D_TITLE => text(&metadata.title),
        D_AUTHOR => text(&metadata.author),
        D_SUBJECT => text(&metadata.subject),
        D_KEYWORDS => NodeValue::owned_list(&metadata.keywords),
        D_DESCRIPTION => text(&metadata.description),
        D_CATEGORY => text(&metadata.category),
        D_LAST_MODIFIED_BY => text(&metadata.last_modified_by),
        D_REVISION => text(&metadata.revision),
        D_CREATED => metadata
            .created
            .and_then(litchi_core::metadata::Timestamp::utc)
            .map(|at| NodeValue::Text(Cow::Owned(at.format("%Y-%m-%dT%H:%M:%SZ").to_string()))),
        D_MODIFIED => metadata
            .modified
            .and_then(litchi_core::metadata::Timestamp::utc)
            .map(|at| NodeValue::Text(Cow::Owned(at.format("%Y-%m-%dT%H:%M:%SZ").to_string()))),
        _ => None,
    }
}

/// An `extended.*` key of the root: the application properties, as
/// OfficeCLI's `ExtendedPropertiesHandler` reads them (record 0342).
fn extended<'a>(
    app: &litchi_ooxml_common::properties::app::Props,
    index: usize,
) -> Option<NodeValue<'a>> {
    let text = |value: &Option<String>| value.as_deref().map(owned);
    let int = |value: Option<i32>| value.map(|value| NodeValue::Int(i64::from(value)));
    match index {
        D_EXTENDED_TEMPLATE => text(&app.template),
        D_EXTENDED_MANAGER => text(&app.manager),
        D_EXTENDED_COMPANY => text(&app.company),
        D_EXTENDED_APPLICATION => text(&app.application),
        D_EXTENDED_APPLICATION_VERSION => text(&app.app_version),
        D_EXTENDED_PAGES => int(app.pages),
        D_EXTENDED_WORDS => int(app.words),
        D_EXTENDED_CHARACTERS => int(app.characters),
        D_EXTENDED_LINES => int(app.lines),
        D_EXTENDED_PARAGRAPHS => int(app.paragraphs),
        D_EXTENDED_TOTAL_TIME => int(app.total_time),
        _ => None,
    }
}

const D_EXTENDED_TEMPLATE: usize = index_of(DOCUMENT_PROPS, "extended.template");
const D_EXTENDED_MANAGER: usize = index_of(DOCUMENT_PROPS, "extended.manager");
const D_EXTENDED_COMPANY: usize = index_of(DOCUMENT_PROPS, "extended.company");
const D_EXTENDED_APPLICATION: usize = index_of(DOCUMENT_PROPS, "extended.application");
const D_EXTENDED_APPLICATION_VERSION: usize =
    index_of(DOCUMENT_PROPS, "extended.applicationVersion");
const D_EXTENDED_PAGES: usize = index_of(DOCUMENT_PROPS, "extended.pages");
const D_EXTENDED_WORDS: usize = index_of(DOCUMENT_PROPS, "extended.words");
const D_EXTENDED_CHARACTERS: usize = index_of(DOCUMENT_PROPS, "extended.characters");
const D_EXTENDED_LINES: usize = index_of(DOCUMENT_PROPS, "extended.lines");
const D_EXTENDED_PARAGRAPHS: usize = index_of(DOCUMENT_PROPS, "extended.paragraphs");
const D_EXTENDED_TOTAL_TIME: usize = index_of(DOCUMENT_PROPS, "extended.totalTime");

// ---------------------------------------------------------------------------
// Paragraph
// ---------------------------------------------------------------------------

const P_PARA_ID: usize = index_of(PARAGRAPH_PROPS, "paraId");
const P_STYLE: usize = index_of(PARAGRAPH_PROPS, "style");
const P_STYLE_ID: usize = index_of(PARAGRAPH_PROPS, "styleId");
pub(super) const P_STYLE_NAME: usize = index_of(PARAGRAPH_PROPS, "styleName");
pub(super) const P_NUM_FMT: usize = index_of(PARAGRAPH_PROPS, "numFmt");
pub(super) const P_LIST_STYLE: usize = index_of(PARAGRAPH_PROPS, "listStyle");
pub(super) const P_START: usize = index_of(PARAGRAPH_PROPS, "start");
const P_ALIGN: usize = index_of(PARAGRAPH_PROPS, "align");
const P_SPACE_BEFORE: usize = index_of(PARAGRAPH_PROPS, "spaceBefore");
const P_SPACE_AFTER: usize = index_of(PARAGRAPH_PROPS, "spaceAfter");
const P_LINE_SPACING: usize = index_of(PARAGRAPH_PROPS, "lineSpacing");
const P_LINE_RULE: usize = index_of(PARAGRAPH_PROPS, "lineRule");
const P_SPACE_BEFORE_LINES: usize = index_of(PARAGRAPH_PROPS, "spaceBeforeLines");
const P_SPACE_AFTER_LINES: usize = index_of(PARAGRAPH_PROPS, "spaceAfterLines");
const P_SPACE_BEFORE_AUTO: usize = index_of(PARAGRAPH_PROPS, "spaceBeforeAuto");
const P_SPACE_AFTER_AUTO: usize = index_of(PARAGRAPH_PROPS, "spaceAfterAuto");
const P_FIRST_LINE_INDENT: usize = index_of(PARAGRAPH_PROPS, "firstLineIndent");
const P_HANGING_INDENT: usize = index_of(PARAGRAPH_PROPS, "hangingIndent");
const P_INDENT: usize = index_of(PARAGRAPH_PROPS, "indent");
const P_RIGHT_INDENT: usize = index_of(PARAGRAPH_PROPS, "rightIndent");
const P_FIRST_LINE_CHARS: usize = index_of(PARAGRAPH_PROPS, "firstLineChars");
const P_HANGING_CHARS: usize = index_of(PARAGRAPH_PROPS, "hangingChars");
const P_LEFT_CHARS: usize = index_of(PARAGRAPH_PROPS, "leftChars");
const P_RIGHT_CHARS: usize = index_of(PARAGRAPH_PROPS, "rightChars");
const P_KEEP_NEXT: usize = index_of(PARAGRAPH_PROPS, "keepNext");
const P_KEEP_LINES: usize = index_of(PARAGRAPH_PROPS, "keepLines");
const P_PAGE_BREAK_BEFORE: usize = index_of(PARAGRAPH_PROPS, "pageBreakBefore");
const P_WIDOW_CONTROL: usize = index_of(PARAGRAPH_PROPS, "widowControl");
const P_DIRECTION: usize = index_of(PARAGRAPH_PROPS, "direction");
const P_CONTEXTUAL_SPACING: usize = index_of(PARAGRAPH_PROPS, "contextualSpacing");
const P_NUM_ID: usize = index_of(PARAGRAPH_PROPS, "numId");
const P_NUM_LEVEL: usize = index_of(PARAGRAPH_PROPS, "numLevel");
const P_OUTLINE_LEVEL: usize = index_of(PARAGRAPH_PROPS, "outlineLvl");

pub(super) fn paragraph_value<'a>(
    props: &direct::Paragraph,
    id: Option<&str>,
    index: usize,
) -> Option<NodeValue<'a>> {
    let spacing = props.spacing.as_ref();
    let indentation: Option<&Indentation> = props.indentation.as_ref();
    let flag = |value: Option<bool>| value.map(NodeValue::Bool);
    match index {
        P_PARA_ID => id.map(owned),
        P_STYLE | P_STYLE_ID => props.style.as_deref().map(owned),
        P_ALIGN => props.justification.map(|align| {
            NodeValue::Token(match align {
                Justification::Both => "justify",
                other => other.as_str(),
            })
        }),
        P_SPACE_BEFORE => spacing?
            .before
            .and_then(|before| i64::try_from(before).ok())
            .and_then(twips),
        P_SPACE_AFTER => spacing?
            .after
            .and_then(|after| i64::try_from(after).ok())
            .and_then(twips),
        P_LINE_SPACING => {
            let spacing = spacing?;
            let line = f64::from(spacing.line?);
            Some(NodeValue::Text(Cow::Owned(match spacing.line_rule {
                None | Some(LineSpacingRule::Auto) => format!("{}x", decimal(line / 240.0, 4)),
                Some(_) => format!("{}pt", decimal(line / 20.0, 2)),
            })))
        },
        P_LINE_RULE => spacing?
            .line_rule
            .map(|rule| NodeValue::Token(rule.as_str())),
        P_SPACE_BEFORE_LINES => spacing?.before_lines.map(NodeValue::from),
        P_SPACE_AFTER_LINES => spacing?.after_lines.map(NodeValue::from),
        P_SPACE_BEFORE_AUTO => flag(spacing?.before_auto_spacing),
        P_SPACE_AFTER_AUTO => flag(spacing?.after_auto_spacing),
        P_FIRST_LINE_INDENT => indentation?.first_line.and_then(|value| twips(value.0)),
        P_HANGING_INDENT => indentation?.hanging.and_then(|value| twips(value.0)),
        P_INDENT => indentation?.leading().and_then(|value| twips(value.0)),
        P_RIGHT_INDENT => indentation?.trailing().and_then(|value| twips(value.0)),
        P_FIRST_LINE_CHARS => indentation?.first_line_chars.map(NodeValue::Int),
        P_HANGING_CHARS => indentation?.hanging_chars.map(NodeValue::Int),
        P_LEFT_CHARS => {
            let indentation = indentation?;
            indentation
                .left_chars
                .or(indentation.start_chars)
                .map(NodeValue::Int)
        },
        P_RIGHT_CHARS => {
            let indentation = indentation?;
            indentation
                .right_chars
                .or(indentation.end_chars)
                .map(NodeValue::Int)
        },
        P_KEEP_NEXT => flag(props.keep_next),
        P_KEEP_LINES => flag(props.keep_lines),
        P_PAGE_BREAK_BEFORE => flag(props.page_break_before),
        P_WIDOW_CONTROL => flag(props.widow_control),
        P_DIRECTION => props
            .bidi
            .map(|rtl| NodeValue::Token(if rtl { "rtl" } else { "ltr" })),
        P_CONTEXTUAL_SPACING => flag(props.contextual_spacing),
        // A direct `w:numPr` names an instance and a level (0 when absent);
        // one without `w:numId` overrides the level alone. Numbering
        // inherited from the style is not resolved (local properties).
        P_NUM_ID => props
            .numbering?
            .id
            .map(|id| NodeValue::Text(Cow::Owned(id.to_string()))),
        P_NUM_LEVEL => {
            let numbering = props.numbering?;
            let level = match numbering.id {
                Some(_) => Some(numbering.level.unwrap_or(0)),
                None => numbering.level,
            };
            level.map(|level| NodeValue::Text(Cow::Owned(level.to_string())))
        },
        P_OUTLINE_LEVEL => props
            .outline_level
            .map(|level| NodeValue::Int(i64::from(level))),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------

const R_FONT_LATIN: usize = index_of(RUN_PROPS, "font.latin");
const R_FONT_ASCII: usize = index_of(RUN_PROPS, "font.ascii");
const R_FONT_HANSI: usize = index_of(RUN_PROPS, "font.hAnsi");
const R_FONT_EA: usize = index_of(RUN_PROPS, "font.ea");
const R_FONT_ASCII_THEME: usize = index_of(RUN_PROPS, "font.asciiTheme");
const R_FONT_HANSI_THEME: usize = index_of(RUN_PROPS, "font.hAnsiTheme");
const R_FONT_EA_THEME: usize = index_of(RUN_PROPS, "font.eaTheme");
const R_FONT_CS_THEME: usize = index_of(RUN_PROPS, "font.csTheme");
const R_FONT_HINT: usize = index_of(RUN_PROPS, "font.hint");
const R_LANG_LATIN: usize = index_of(RUN_PROPS, "lang.latin");
const R_LANG_EA: usize = index_of(RUN_PROPS, "lang.ea");
const R_LANG_CS: usize = index_of(RUN_PROPS, "lang.cs");
const R_SIZE: usize = index_of(RUN_PROPS, "size");
const R_BOLD: usize = index_of(RUN_PROPS, "bold");
const R_ITALIC: usize = index_of(RUN_PROPS, "italic");
const R_FONT_CS: usize = index_of(RUN_PROPS, "font.cs");
const R_SIZE_CS: usize = index_of(RUN_PROPS, "size.cs");
const R_BOLD_CS: usize = index_of(RUN_PROPS, "bold.cs");
const R_ITALIC_CS: usize = index_of(RUN_PROPS, "italic.cs");
const R_COLOR: usize = index_of(RUN_PROPS, "color");
const R_UNDERLINE: usize = index_of(RUN_PROPS, "underline");
const R_UNDERLINE_COLOR: usize = index_of(RUN_PROPS, "underline.color");
const R_STRIKE: usize = index_of(RUN_PROPS, "strike");
const R_HIGHLIGHT: usize = index_of(RUN_PROPS, "highlight");
const R_CAPS: usize = index_of(RUN_PROPS, "caps");
const R_SMALL_CAPS: usize = index_of(RUN_PROPS, "smallcaps");
const R_DOUBLE_STRIKE: usize = index_of(RUN_PROPS, "dstrike");
const R_VANISH: usize = index_of(RUN_PROPS, "vanish");
const R_OUTLINE: usize = index_of(RUN_PROPS, "outline");
const R_SHADOW: usize = index_of(RUN_PROPS, "shadow");
const R_EMBOSS: usize = index_of(RUN_PROPS, "emboss");
const R_IMPRINT: usize = index_of(RUN_PROPS, "imprint");
const R_NO_PROOF: usize = index_of(RUN_PROPS, "noproof");
const R_DIRECTION: usize = index_of(RUN_PROPS, "direction");
const R_SUPERSCRIPT: usize = index_of(RUN_PROPS, "superscript");
const R_SUBSCRIPT: usize = index_of(RUN_PROPS, "subscript");
const R_POSITION: usize = index_of(RUN_PROPS, "position");
const R_CHAR_SPACING: usize = index_of(RUN_PROPS, "charSpacing");
const R_STYLE: usize = index_of(RUN_PROPS, "rStyle");

pub(super) fn run<'a>(props: &direct::Run, index: usize) -> Option<NodeValue<'a>> {
    let fonts = props.fonts.as_ref();
    let flag = |value: Option<bool>| value.map(NodeValue::Bool);
    // OfficeCLI collapses equal ASCII and high-ANSI slots into `font.latin`
    // and keeps differing or single slots apart.
    let latin = || {
        let fonts = fonts?;
        let ascii = fonts.ascii.as_deref().filter(|name| !name.is_empty());
        let h_ansi = fonts.h_ansi.as_deref().filter(|name| !name.is_empty());
        Some((ascii, h_ansi))
    };
    let theme = |slot: Option<litchi_docx::properties::ThemeFont>| {
        slot.map(|theme| NodeValue::Token(theme.as_str()))
    };
    let language = props.language.as_ref();
    match index {
        R_FONT_LATIN => match latin()? {
            (Some(ascii), Some(h_ansi)) if ascii == h_ansi => Some(owned(ascii)),
            _ => None,
        },
        R_FONT_ASCII => match latin()? {
            (Some(ascii), Some(h_ansi)) if ascii == h_ansi => None,
            (Some(ascii), _) => Some(owned(ascii)),
            _ => None,
        },
        R_FONT_HANSI => match latin()? {
            (Some(ascii), Some(h_ansi)) if ascii == h_ansi => None,
            (_, Some(h_ansi)) => Some(owned(h_ansi)),
            _ => None,
        },
        R_FONT_EA => text_of(fonts?.east_asia.as_ref()),
        R_FONT_ASCII_THEME => theme(fonts?.ascii_theme),
        R_FONT_HANSI_THEME => theme(fonts?.h_ansi_theme),
        R_FONT_EA_THEME => theme(fonts?.east_asia_theme),
        R_FONT_CS_THEME => theme(fonts?.complex_script_theme),
        R_FONT_HINT => fonts?.hint.map(|hint| NodeValue::Token(hint.as_str())),
        R_LANG_LATIN => language?.latin.as_deref().map(owned),
        R_LANG_EA => language?.east_asia.as_deref().map(owned),
        R_LANG_CS => language?.bidi.as_deref().map(owned),
        R_SIZE => props.size.and_then(|size| half_points(size.0)),
        R_BOLD => flag(props.bold),
        R_ITALIC => flag(props.italic),
        R_FONT_CS => text_of(fonts?.complex_script.as_ref()),
        R_SIZE_CS => props
            .size_complex_script
            .and_then(|size| half_points(size.0)),
        R_BOLD_CS => flag(props.bold_complex_script),
        R_ITALIC_CS => flag(props.italic_complex_script),
        R_COLOR => props.color.as_ref().map(word_color),
        R_UNDERLINE => props
            .underline
            .map(|underline| NodeValue::Token(underline.style.as_str())),
        R_UNDERLINE_COLOR => props.underline?.color.map(|color| {
            NodeValue::Color(match color {
                RunUnderlineColor::Auto => Color::AUTO,
                RunUnderlineColor::Rgb([red, green, blue]) => Color::rgb(red, green, blue),
            })
        }),
        R_STRIKE => flag(props.strike),
        R_HIGHLIGHT => props
            .highlight
            .map(|highlight| NodeValue::Token(highlight.as_str())),
        R_CAPS => flag(props.caps),
        R_SMALL_CAPS => flag(props.small_caps),
        R_DOUBLE_STRIKE => flag(props.double_strike),
        R_VANISH => flag(props.vanish),
        R_OUTLINE => flag(props.outline),
        R_SHADOW => flag(props.shadow),
        R_EMBOSS => flag(props.emboss),
        R_IMPRINT => flag(props.imprint),
        R_NO_PROOF => flag(props.no_proof),
        R_DIRECTION => props
            .right_to_left
            .map(|rtl| NodeValue::Token(if rtl { "rtl" } else { "ltr" })),
        // OfficeCLI writes only the position that is set.
        R_SUPERSCRIPT => (props.vertical_align == Some(VerticalAlign::Superscript))
            .then_some(NodeValue::Bool(true)),
        R_SUBSCRIPT => (props.vertical_align == Some(VerticalAlign::Subscript))
            .then_some(NodeValue::Bool(true)),
        R_POSITION => props.position.and_then(|position| half_points(position.0)),
        R_CHAR_SPACING => props.character_spacing.and_then(|spacing| twips(spacing.0)),
        R_STYLE => props.style.as_deref().map(owned),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Picture
// ---------------------------------------------------------------------------

fn picture_value<'a>(
    picture: &litchi_docx::image::InlineImage,
    index: usize,
) -> Option<NodeValue<'a>> {
    const WIDTH: usize = index_of(super::spec::PICTURE_PROPS, "width");
    const HEIGHT: usize = index_of(super::spec::PICTURE_PROPS, "height");
    const ALT: usize = index_of(super::spec::PICTURE_PROPS, "alt");
    const NAME: usize = index_of(super::spec::PICTURE_PROPS, "name");
    const REL_ID: usize = index_of(super::spec::PICTURE_PROPS, "relId");
    match index {
        WIDTH => Some(NodeValue::Length(Length::from_emu(picture.width_emu()))),
        HEIGHT => Some(NodeValue::Length(Length::from_emu(picture.height_emu()))),
        ALT => (!picture.description().is_empty()).then(|| owned(picture.description())),
        NAME => (!picture.name().is_empty()).then(|| owned(picture.name())),
        REL_ID => (!picture.r_embed().is_empty()).then(|| owned(picture.r_embed())),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Comment
// ---------------------------------------------------------------------------

/// A comment's text: its paragraphs' text, one per line, as the document
/// text separates body paragraphs.
fn comment_text(comment: &litchi_docx::comment::Comment) -> Result<String> {
    let blocks = comment.blocks()?;
    let mut text = String::new();
    for (index, paragraph) in blocks
        .iter()
        .filter_map(|block| block.paragraph())
        .enumerate()
    {
        if index > 0 {
            text.push('\n');
        }
        text.push_str(&paragraph.text()?);
    }
    Ok(text)
}

fn comment_value<'a>(
    docx: &super::Docx,
    comment: &CommentProps,
    index: usize,
) -> Option<NodeValue<'a>> {
    const ID: usize = index_of(super::spec::COMMENT_PROPS, "id");
    const AUTHOR: usize = index_of(super::spec::COMMENT_PROPS, "author");
    const INITIALS: usize = index_of(super::spec::COMMENT_PROPS, "initials");
    const DATE: usize = index_of(super::spec::COMMENT_PROPS, "date");
    const TEXT: usize = index_of(super::spec::COMMENT_PROPS, "text");
    const ANCHORED_TO: usize = index_of(super::spec::COMMENT_PROPS, "anchoredTo");
    const DONE: usize = index_of(super::spec::COMMENT_PROPS, "done");
    const PARENT: usize = index_of(super::spec::COMMENT_PROPS, "parentId");
    match index {
        ID => Some(NodeValue::Int(i64::from(comment.id))),
        AUTHOR => text_of(Some(&comment.author)),
        INITIALS => text_of(comment.initials.as_ref()),
        DATE => text_of(comment.date.as_ref()),
        TEXT => Some(owned(&comment.text)),
        // Read on demand: finding the anchor walks the main story.
        ANCHORED_TO => comment_anchor(docx, comment.id)
            .ok()
            .flatten()
            .map(|path| NodeValue::Text(Cow::Owned(path))),
        // OfficeCLI reports `done` on every comment, `parentId` on replies.
        DONE => Some(NodeValue::Bool(comment.done)),
        PARENT => comment
            .parent
            .map(|parent| NodeValue::Int(i64::from(parent))),
        _ => None,
    }
}

/// A comment's resolved state and reply parent (the parent comment's id),
/// from its last paragraph's `commentsExtended` entry (record 0345):
/// unresolved and no parent without one.
fn comment_thread(
    docx: &super::Docx,
    comment: &litchi_docx::comment::Comment,
) -> Result<(bool, Option<u32>)> {
    let Some(paragraph) = comment.paragraph_id() else {
        return Ok((false, None));
    };
    let entries = super::nav::comments_extended(docx)?;
    let Some(entry) = entries.iter().find(|entry| entry.paragraph_id == paragraph) else {
        return Ok((false, None));
    };
    let parent = match entry.parent_paragraph_id {
        Some(parent) => super::nav::comments(docx)?
            .iter()
            .find(|other| other.paragraph_id() == Some(parent))
            .map(litchi_docx::comment::Comment::id),
        None => None,
    };
    Ok((entry.done, parent))
}

/// The path of the paragraph comment `id` is anchored to, as OfficeCLI's
/// `FindCommentAnchorPath` finds it: the paragraph holding its range start,
/// or failing that its reference (a point comment), anywhere in the main
/// story's body, tables and content controls (record 0345). One walk of
/// the story finds every comment's anchor; the handle keeps the map.
pub(super) fn comment_anchor(docx: &super::Docx, id: u32) -> Result<Option<String>> {
    if let Some(anchors) = docx.parts().comment_anchors.get() {
        return Ok(anchors.get(&id).cloned());
    }
    let body = super::nav::load_body(docx)?;
    let mut starts = std::collections::HashMap::new();
    let mut references = std::collections::HashMap::new();
    let finders = [
        memchr::memmem::Finder::new(b"commentRangeStart"),
        memchr::memmem::Finder::new(b"commentReference"),
    ];
    anchors_in(&body, "/body", &finders, &mut starts, &mut references)?;
    for (id, path) in references {
        starts.entry(id).or_insert(path);
    }
    let anchors = docx
        .parts()
        .comment_anchors
        .get_or_init(|| Arc::new(starts));
    Ok(anchors.get(&id).cloned())
}

/// Record, for every comment id, the first paragraph of `blocks` (in
/// document order) holding its range start and its reference.
fn anchors_in(
    blocks: &litchi_docx::Blocks,
    prefix: &str,
    finders: &[memchr::memmem::Finder<'_>; 2],
    starts: &mut std::collections::HashMap<u32, String>,
    references: &mut std::collections::HashMap<u32, String>,
) -> Result<()> {
    for block in blocks.iter() {
        let Some(kind) = Kind::of_block(block.kind()) else {
            continue;
        };
        let xml = block.xml_bytes();
        if finders.iter().all(|finder| finder.find(xml).is_none()) {
            continue;
        }
        let step = format!(
            "{prefix}/{}[{}]",
            kind.step_name(),
            block.ordinal().get() + 1
        );
        match kind {
            Kind::Paragraph | Kind::Equation => {
                for (marker, map) in [
                    (finders[0].needle(), &mut *starts),
                    (finders[1].needle(), &mut *references),
                ] {
                    for id in marker_ids(xml, marker) {
                        map.entry(id).or_insert_with(|| step.clone());
                    }
                }
            },
            Kind::Table => {
                let Some(table) = block.table() else { continue };
                for (row_index, row) in table.rows()?.iter().enumerate() {
                    for (cell_index, cell) in row.cells()?.iter().enumerate() {
                        let cell_prefix =
                            format!("{step}/tr[{}]/tc[{}]", row_index + 1, cell_index + 1);
                        anchors_in(&cell.blocks()?, &cell_prefix, finders, starts, references)?;
                    }
                }
            },
            Kind::Sdt => {
                if let Some(control) = block.content_control() {
                    anchors_in(&control.blocks()?, &step, finders, starts, references)?;
                }
            },
            _ => {},
        }
    }
    Ok(())
}

/// The `id` of every `marker` element of `xml`, in document order.
fn marker_ids(xml: &[u8], marker: &[u8]) -> Vec<u32> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut ids = Vec::new();
    loop {
        match reader.read_event() {
            Ok(
                quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element),
            ) if element.local_name().as_ref() == marker => {
                if let Some(id) = element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .find(|attribute| attribute.key.local_name().as_ref() == b"id")
                    .and_then(|attribute| {
                        std::str::from_utf8(&attribute.value)
                            .ok()?
                            .trim()
                            .parse::<u32>()
                            .ok()
                    })
                {
                    ids.push(id);
                }
            },
            Ok(quick_xml::events::Event::Eof) | Err(_) => return ids,
            _ => {},
        }
    }
}

// ---------------------------------------------------------------------------
// Hyperlink
// ---------------------------------------------------------------------------

const H_URL: usize = index_of(HYPERLINK_PROPS, "url");
const H_ANCHOR: usize = index_of(HYPERLINK_PROPS, "anchor");
const H_DOC_LOCATION: usize = index_of(HYPERLINK_PROPS, "docLocation");
const H_TOOLTIP: usize = index_of(HYPERLINK_PROPS, "tooltip");
const H_TARGET_FRAME: usize = index_of(HYPERLINK_PROPS, "tgtFrame");
const H_FONT: usize = index_of(HYPERLINK_PROPS, "font");
const H_FONT_CS: usize = index_of(HYPERLINK_PROPS, "font.cs");
const H_SIZE: usize = index_of(HYPERLINK_PROPS, "size");
const H_BOLD: usize = index_of(HYPERLINK_PROPS, "bold");
const H_ITALIC: usize = index_of(HYPERLINK_PROPS, "italic");
const H_COLOR: usize = index_of(HYPERLINK_PROPS, "color");
const H_UNDERLINE: usize = index_of(HYPERLINK_PROPS, "underline");
const H_UNDERLINE_COLOR: usize = index_of(HYPERLINK_PROPS, "underline.color");
const H_STRIKE: usize = index_of(HYPERLINK_PROPS, "strike");
const H_HIGHLIGHT: usize = index_of(HYPERLINK_PROPS, "highlight");
const H_STYLE: usize = index_of(HYPERLINK_PROPS, "rStyle");

fn hyperlink<'a>(
    link: &'a InlineHyperlink,
    first: Option<&direct::Run>,
    index: usize,
) -> Option<NodeValue<'a>> {
    let text = |value: Option<&'a str>| value.map(|text| NodeValue::Text(Cow::Borrowed(text)));
    match index {
        H_URL => text(link.link().url()),
        H_ANCHOR => text(link.link().anchor()),
        H_DOC_LOCATION => text(link.document_location()),
        H_TOOLTIP => text(link.link().tooltip()),
        H_TARGET_FRAME => text(link.target_frame()),
        H_FONT => first?.fonts.as_ref()?.ascii.as_deref().map(owned),
        H_FONT_CS => first?.fonts.as_ref()?.complex_script.as_deref().map(owned),
        // The theme colour wins over the stored value here, as
        // `HyperlinkToNode` reads it, without a transform tail.
        H_COLOR => {
            let color = first?.color.as_ref()?;
            Some(NodeValue::Color(
                match color
                    .theme
                    .filter(|theme| *theme != Theme::None)
                    .and_then(scheme)
                {
                    Some(scheme) => Color::scheme(scheme),
                    None => hex_color(color.value),
                },
            ))
        },
        H_SIZE => run(first?, R_SIZE),
        H_BOLD => run(first?, R_BOLD),
        H_ITALIC => run(first?, R_ITALIC),
        H_UNDERLINE => run(first?, R_UNDERLINE),
        H_UNDERLINE_COLOR => run(first?, R_UNDERLINE_COLOR),
        H_STRIKE => run(first?, R_STRIKE),
        H_HIGHLIGHT => run(first?, R_HIGHLIGHT),
        H_STYLE => run(first?, R_STYLE),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

/// A table, cell or row width as OfficeCLI's `FormatTableWidth` writes it.
fn width_text<'a>(width: Width) -> NodeValue<'a> {
    NodeValue::Text(Cow::Owned(match width {
        Width::Nil => "nil".to_owned(),
        Width::Auto => "auto".to_owned(),
        Width::Twips(twips) => format!("{}dxa", twips.0),
        Width::Percent(fiftieths) => percent(fiftieths),
    }))
}

/// A `pct` width in fiftieths of a percent, as `FormatPctWidth` writes it.
fn percent(fiftieths: i64) -> String {
    #[allow(
        clippy::cast_precision_loss,
        reason = "widths are far below 2^53 fiftieths"
    )]
    let value = fiftieths as f64 / 50.0;
    format!("{}%", decimal(value, 2))
}

/// The twips of a `dxa` width, for keys OfficeCLI writes as numbers.
fn twips_number<'a>(width: Option<Width>) -> Option<NodeValue<'a>> {
    match width? {
        Width::Twips(twips) => Some(NodeValue::Int(twips.0)),
        _ => None,
    }
}

const fn table_align(value: TableJustification) -> &'static str {
    match value {
        TableJustification::Center => "center",
        TableJustification::End => "end",
        TableJustification::Left => "left",
        TableJustification::Right => "right",
        TableJustification::Start => "start",
    }
}

const T_COLS: usize = index_of(TABLE_PROPS, "cols");
const T_ROWS: usize = index_of(TABLE_PROPS, "rows");
const T_STYLE: usize = index_of(TABLE_PROPS, "style");
const T_ROW_BAND: usize = index_of(TABLE_PROPS, "rowBandSize");
const T_COL_BAND: usize = index_of(TABLE_PROPS, "colBandSize");
const T_WIDTH: usize = index_of(TABLE_PROPS, "width");
const T_ALIGN: usize = index_of(TABLE_PROPS, "align");
const T_INDENT: usize = index_of(TABLE_PROPS, "indent");
const T_CELL_SPACING: usize = index_of(TABLE_PROPS, "cellSpacing");
const T_LAYOUT: usize = index_of(TABLE_PROPS, "layout");
const T_LOOK: usize = index_of(TABLE_PROPS, "tblLook");
const T_DIRECTION: usize = index_of(TABLE_PROPS, "direction");
const T_PADDING_TOP: usize = index_of(TABLE_PROPS, "padding.top");
const T_PADDING_BOTTOM: usize = index_of(TABLE_PROPS, "padding.bottom");
const T_PADDING_LEFT: usize = index_of(TABLE_PROPS, "padding.left");
const T_PADDING_RIGHT: usize = index_of(TABLE_PROPS, "padding.right");
const T_FIRST_ROW: usize = index_of(TABLE_PROPS, "firstRow");
const T_LAST_ROW: usize = index_of(TABLE_PROPS, "lastRow");
const T_FIRST_COL: usize = index_of(TABLE_PROPS, "firstCol");
const T_LAST_COL: usize = index_of(TABLE_PROPS, "lastCol");
const T_BANDED_ROWS: usize = index_of(TABLE_PROPS, "bandedRows");
const T_BANDED_COLS: usize = index_of(TABLE_PROPS, "bandedCols");
const T_CAPTION: usize = index_of(TABLE_PROPS, "caption");
const T_DESCRIPTION: usize = index_of(TABLE_PROPS, "description");
const T_COL_WIDTHS: usize = index_of(TABLE_PROPS, "colWidths");

fn table_value<'a>(table: &TableProps, index: usize) -> Option<NodeValue<'a>> {
    let props = &table.props;
    let look = props.look.as_ref();
    // OfficeCLI writes a look flag only when it is on, and banding only
    // when `noHBand`/`noVBand` is explicitly off.
    let on = |value: Option<bool>| (value == Some(true)).then_some(NodeValue::Bool(true));
    let off = |value: Option<bool>| (value == Some(false)).then_some(NodeValue::Bool(true));
    match index {
        T_COLS => Some(NodeValue::Text(Cow::Owned(table.columns.to_string()))),
        T_ROWS => Some(NodeValue::Text(Cow::Owned(table.rows.to_string()))),
        T_STYLE => text_of(props.style.as_ref()),
        T_ROW_BAND => props.row_band_size.map(NodeValue::from),
        T_COL_BAND => props.column_band_size.map(NodeValue::from),
        T_WIDTH => props.width.map(|width| {
            NodeValue::Text(Cow::Owned(match width {
                Width::Twips(twips) => twips.0.to_string(),
                Width::Percent(fiftieths) => percent(fiftieths),
                Width::Auto => "auto".to_owned(),
                Width::Nil => "0".to_owned(),
            }))
        }),
        T_ALIGN => props
            .justification
            .map(|align| NodeValue::Token(table_align(align))),
        T_INDENT => match props.indent? {
            Width::Twips(twips) => Some(NodeValue::Int(twips.0)),
            Width::Percent(fiftieths) => Some(NodeValue::Text(Cow::Owned(percent(fiftieths)))),
            Width::Auto | Width::Nil => None,
        },
        T_CELL_SPACING => twips_number(props.cell_spacing),
        T_LAYOUT => props.layout.map(|layout| {
            NodeValue::Token(match layout {
                TableLayout::Fixed => "fixed",
                TableLayout::Autofit => "autofit",
            })
        }),
        // The conditional-formatting mask as four hexadecimal digits.
        T_LOOK => look.map(|look| NodeValue::Text(Cow::Owned(format!("{:04X}", look.mask())))),
        T_DIRECTION => (props.bidi_visual == Some(true)).then_some(NodeValue::Token("rtl")),
        T_PADDING_TOP => twips_number(props.cell_margins.as_ref()?.top),
        T_PADDING_BOTTOM => twips_number(props.cell_margins.as_ref()?.bottom),
        T_PADDING_LEFT => {
            let margins = props.cell_margins.as_ref()?;
            twips_number(margins.left.or(margins.start))
        },
        T_PADDING_RIGHT => {
            let margins = props.cell_margins.as_ref()?;
            twips_number(margins.right.or(margins.end))
        },
        T_FIRST_ROW => on(look?.first_row),
        T_LAST_ROW => on(look?.last_row),
        T_FIRST_COL => on(look?.first_column),
        T_LAST_COL => on(look?.last_column),
        T_BANDED_ROWS => off(look?.no_horizontal_band),
        T_BANDED_COLS => off(look?.no_vertical_band),
        T_CAPTION => text_of(props.caption.as_ref()),
        T_DESCRIPTION => text_of(props.description.as_ref()),
        T_COL_WIDTHS => {
            let grid = props.grid.as_ref().filter(|grid| !grid.is_empty())?;
            let mut widths = String::new();
            for (position, width) in grid.iter().enumerate() {
                if position > 0 {
                    widths.push(',');
                }
                let _ = write!(widths, "{}dxa", width.map_or(0, |width| width.0));
            }
            Some(NodeValue::Text(Cow::Owned(widths)))
        },
        _ => None,
    }
}

const W_HEIGHT: usize = index_of(ROW_PROPS, "height");
const W_HEIGHT_RULE: usize = index_of(ROW_PROPS, "height.rule");
const W_HEADER: usize = index_of(ROW_PROPS, "header");
const W_CANT_SPLIT: usize = index_of(ROW_PROPS, "cantSplit");
const W_HIDDEN: usize = index_of(ROW_PROPS, "hidden");
const W_ALIGN: usize = index_of(ROW_PROPS, "rowAlign");
const W_GRID_BEFORE: usize = index_of(ROW_PROPS, "gridBefore");
const W_WIDTH_BEFORE: usize = index_of(ROW_PROPS, "wBefore");
const W_GRID_AFTER: usize = index_of(ROW_PROPS, "gridAfter");
const W_WIDTH_AFTER: usize = index_of(ROW_PROPS, "wAfter");
const W_CELL_SPACING: usize = index_of(ROW_PROPS, "cellSpacing");

fn row<'a>(props: &direct::Row, index: usize) -> Option<NodeValue<'a>> {
    let on = |value: Option<bool>| (value == Some(true)).then_some(NodeValue::Bool(true));
    match index {
        W_HEIGHT => props
            .height
            .map(|height| NodeValue::Text(Cow::Owned(format!("{}dxa", height.value.0)))),
        W_HEIGHT_RULE => props.height?.rule.and_then(|rule| match rule {
            direct::HeightRule::Exact => Some(NodeValue::Token("exact")),
            direct::HeightRule::AtLeast => Some(NodeValue::Token("atLeast")),
            _ => None,
        }),
        W_HEADER => on(props.header),
        W_CANT_SPLIT => on(props.cant_split),
        W_HIDDEN => on(props.hidden),
        W_ALIGN => props.justification.and_then(|align| match align {
            TableJustification::Center => Some(NodeValue::Token("center")),
            TableJustification::Right => Some(NodeValue::Token("right")),
            TableJustification::Left => Some(NodeValue::Token("left")),
            _ => None,
        }),
        W_GRID_BEFORE => props
            .grid_before
            .map(|count| NodeValue::Text(Cow::Owned(count.to_string()))),
        W_WIDTH_BEFORE => props.width_before.map(width_text),
        W_GRID_AFTER => props
            .grid_after
            .map(|count| NodeValue::Text(Cow::Owned(count.to_string()))),
        W_WIDTH_AFTER => props.width_after.map(width_text),
        W_CELL_SPACING => twips_number(props.cell_spacing),
        _ => None,
    }
}

const C_WIDTH: usize = index_of(CELL_PROPS, "width");
const C_VALIGN: usize = index_of(CELL_PROPS, "valign");
const C_VMERGE: usize = index_of(CELL_PROPS, "vmerge");
const C_HMERGE: usize = index_of(CELL_PROPS, "hmerge");
const C_COLSPAN: usize = index_of(CELL_PROPS, "colspan");
const C_PADDING_TOP: usize = index_of(CELL_PROPS, "padding.top");
const C_PADDING_BOTTOM: usize = index_of(CELL_PROPS, "padding.bottom");
const C_PADDING_LEFT: usize = index_of(CELL_PROPS, "padding.left");
const C_PADDING_RIGHT: usize = index_of(CELL_PROPS, "padding.right");
const C_TEXT_DIRECTION: usize = index_of(CELL_PROPS, "textDirection");
const C_NO_WRAP: usize = index_of(CELL_PROPS, "nowrap");
const C_HIDE_MARK: usize = index_of(CELL_PROPS, "hideMark");
const C_FIT_TEXT: usize = index_of(CELL_PROPS, "tcFitText");

fn cell<'a>(props: &direct::Cell, index: usize) -> Option<NodeValue<'a>> {
    let on = |value: Option<bool>| (value == Some(true)).then_some(NodeValue::Bool(true));
    let merge = |value: Option<Merge>| {
        value.map(|merge| {
            NodeValue::Token(match merge {
                Merge::Restart => "restart",
                Merge::Continue => "continue",
            })
        })
    };
    let margins = props.margins.as_ref();
    match index {
        C_WIDTH => props.width.map(|width| match width {
            Width::Twips(twips) => NodeValue::Text(Cow::Owned(format!("{}dxa", twips.0))),
            other => width_text(other),
        }),
        C_VALIGN => props
            .vertical_align
            .map(|align| NodeValue::Token(align.as_str())),
        C_VMERGE => merge(props.vertical_merge),
        C_HMERGE => merge(props.horizontal_merge),
        C_COLSPAN => props.grid_span.map(NodeValue::from),
        C_PADDING_TOP => twips_number(margins?.top),
        C_PADDING_BOTTOM => twips_number(margins?.bottom),
        C_PADDING_LEFT => twips_number(margins?.left.or(margins?.start)),
        C_PADDING_RIGHT => twips_number(margins?.right.or(margins?.end)),
        C_TEXT_DIRECTION => props
            .text_direction
            .map(|direction| NodeValue::Token(direction.as_str())),
        C_NO_WRAP => on(props.no_wrap),
        C_HIDE_MARK => on(props.hide_mark),
        C_FIT_TEXT => on(props.fit_text),
        _ => None,
    }
}

const S_ALIAS: usize = index_of(SDT_PROPS, "alias");
const S_TAG: usize = index_of(SDT_PROPS, "tag");
const S_ID: usize = index_of(SDT_PROPS, "sdtId");

fn sdt_value<'a>(sdt: &SdtProps, index: usize) -> Option<NodeValue<'a>> {
    match index {
        S_ALIAS => sdt.alias.as_deref().map(owned),
        S_TAG => sdt.tag.as_deref().map(owned),
        S_ID => sdt.id.map(NodeValue::from),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// The node view
// ---------------------------------------------------------------------------

/// The kinds the node view lists under a node: a paragraph lists its runs
/// only, as OfficeCLI does (its hyperlinks are reachable by path and query);
/// every other node lists all its children.
fn view_kinds(kind: Kind) -> KindSet<Kind> {
    match kind {
        Kind::Paragraph => KindSet::of(Kind::Run),
        // Bookmarks, fields, tables of contents, the watermark and text
        // boxes are reached by path and query, not listed (record 0319).
        Kind::Document => KindSet::of(Kind::Body)
            .with(Kind::Styles)
            .with(Kind::Header)
            .with(Kind::Footer)
            .with(Kind::Numbering)
            .with(Kind::Footnotes)
            .with(Kind::Endnotes)
            .with(Kind::Comments)
            .with(Kind::SmartArt),
        Kind::Body => KindSet::of(Kind::Paragraph)
            .with(Kind::Table)
            .with(Kind::Sdt)
            .with(Kind::Equation)
            .with(Kind::Section),
        _ => KindSet::ALL,
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        self.kind.type_name()
    }

    fn file_format(&self) -> litchi_core::FileFormat {
        litchi_core::FileFormat::Docx
    }

    fn path(&self) -> NodePath<'_> {
        self.canonical_path()
    }

    fn positional_path(&self) -> NodePath<'_> {
        self.positional()
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        self.docx.owner().check_source()?;
        self.docx.owner().finish(text(self))
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.docx.owner().check_source()?;
        self.docx.owner().finish(preview(self))
    }

    fn style(&self) -> Result<Option<Cow<'_, str>>> {
        self.docx.owner().check_source()?;
        self.docx.owner().finish(style(self))
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        self.docx.owner().check_source()?;
        self.docx.owner().finish(child_count(self))
    }

    fn format(&self) -> Result<Format<'_>> {
        self.docx.owner().check_source()?;
        self.docx.owner().finish(format(self))
    }

    fn children(&self, _: &ViewOptions) -> Children<'_, 'f> {
        match Nav::new(self.docx).children(self, view_kinds(self.kind)) {
            Ok(children) => Children::docx(children),
            Err(error) => Children::failed(error),
        }
    }
}
