//! DOCX views (record 0010 §6): the text, annotated, outline and stats
//! modes, with OfficeCLI's content and JSON shapes (`WordHandler.View.cs`).
//!
//! The views walk the body in visible order through the adapter's own
//! children, so every line names the element by the path `get` resolves:
//! block content controls are flattened (their blocks get the control's
//! label), `w:customXml` wrappers are flattened by the leaf's index, and
//! display equations are `oMathPara[N]`. The walk keeps an explicit stack,
//! so nested content controls cannot grow the call stack, and the text and
//! annotated views stop at `end` and `max_lines`. Like OfficeCLI, trailing
//! white space of the whole output is dropped.
//!
//! | Mode | Content |
//! |---|---|
//! | text | `[path] text` per paragraph, `[path] [Table: N rows]`, `[path] [Equation] text`; `start`/`end` count listed elements, section breaks included |
//! | annotated | per run `[path] 「text」 ← style \| format`, `[path] [Image: …] ← style`, `[path] [] <- style \| empty paragraph`, `[path] [Table: R×C]`, `[path] [Equation: "text"] ← display` |
//! | outline | the counts line, then the heading tree: `├── [N] "text" (style)` |
//! | stats | paragraph, word and character counts, then the style, font and font-size distributions, empty paragraphs and consecutive spaces |
//!
//! Departures from OfficeCLI, all from reading local properties only
//! (ADR 0004) or from parts the adapter does not read yet:
//!
//! - styles are style ids, not display names (M1-Δ4); run formats, fonts
//!   and sizes are the run's own properties, not the effective ones (M1-Δ2);
//! - list markers, field results, form fields, OLE placeholders, the
//!   watermark and the file name are not shown; equations show their
//!   visible text, not LaTeX;
//! - paragraph paths inside a content control count within the control
//!   (`/body/sdt[1]/p[2]`), where OfficeCLI's view counts body-wide and
//!   prints paths that do not resolve;
//! - the annotated view lists the paragraph's addressable runs (the `r[N]`
//!   set), not every descendant run.

// `File::view`, `view_to` and `view_json` (M1 step R5) are the production
// callers; until they land, the tests are.
#![cfg_attr(
    not(test),
    allow(dead_code, reason = "M1 step R5 exposes the views through File::view")
)]

use std::collections::HashMap;
use std::fmt::{self, Write as _};
use std::ops::ControlFlow;

use litchi_docx::properties::{self as direct, HalfPoints};
use litchi_docx::{BlockKind, Blocks, Run, RunContent};
use serde::Serialize;
use serde::ser::{SerializeMap, Serializer};

use super::Docx;
use super::children::Children;
use super::nav;
use super::node::{Node, Value};
use super::props;
use super::spec::Kind;
use crate::error::{Code, Error, Result};
use crate::path::navigate::KindSet;
use crate::view::mode::{Mode, Options};

/// Writes the `mode` view of the document into `out`.
///
/// # Errors
///
/// `unsupported_capability` for the modes Litchi does not render,
/// `invalid_argument` for workbook-only options, the leaf's error for a part
/// that fails to read, and `io_error` when the sink refuses output.
pub(crate) fn view_to<W: fmt::Write>(
    docx: &Docx,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    docx.owner().check_source()?;
    let result = view_to_checked(docx, mode, options, out);
    docx.owner().finish(result)
}

fn view_to_checked<W: fmt::Write>(
    docx: &Docx,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    check_options(options)?;
    let mut out = TrimEnd::new(out);
    match mode {
        Mode::Text => text(docx, options, &mut out),
        Mode::Annotated => annotated(docx, options, &mut out),
        Mode::Outline => outline(docx)?.write(&mut out).map_err(sink),
        Mode::Stats => stats(docx)?.write(&mut out).map_err(sink),
        _ => Err(mode.unsupported()),
    }
}

/// The `mode` view of the document as OfficeCLI's JSON.
///
/// # Errors
///
/// As [`view_to`].
pub(crate) fn view_json(docx: &Docx, mode: Mode, options: &Options) -> Result<Json> {
    docx.owner().check_source()?;
    let result = view_json_checked(docx, mode, options);
    docx.owner().finish(result)
}

fn view_json_checked(docx: &Docx, mode: Mode, options: &Options) -> Result<Json> {
    check_options(options)?;
    Ok(match mode {
        Mode::Text => Json::Text(text_json(docx, options)?),
        Mode::Annotated => {
            let mut content = String::new();
            view_to(docx, mode, options, &mut content)?;
            Json::Annotated(Annotated {
                view: "annotated",
                content,
            })
        },
        Mode::Outline => Json::Outline(outline(docx)?.json()),
        Mode::Stats => Json::Stats(stats(docx)?.json()),
        _ => return Err(mode.unsupported()),
    })
}

/// Refuses the workbook-only options (`--cols`, `--range`), as OfficeCLI
/// does for documents.
fn check_options(options: &Options) -> Result<()> {
    if options.columns.is_some() || options.range.is_some() {
        return Err(Error::new(
            Code::InvalidArgument,
            "columns and ranges select worksheet cells; DOCX views do not take them",
        )
        .with_format(litchi_core::FileFormat::Docx)
        .with_suggestion("use start, end or max_lines to bound a document view"));
    }
    Ok(())
}

fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the view's output sink refused to write")
}

// ---------------------------------------------------------------------------
// The body, flattened
// ---------------------------------------------------------------------------

/// The kinds a view lists in a block container.
fn listed_kinds() -> KindSet<Kind> {
    KindSet::of(Kind::Paragraph)
        .with(Kind::Table)
        .with(Kind::Sdt)
        .with(Kind::Equation)
        .with(Kind::Section)
}

/// One listed element of the flattened body and the label of the content
/// control holding it.
struct Line<'f> {
    node: Node<'f>,
    /// `[sdt:alias] `, `[sdt:tag] ` or `[sdt] `, when a content control
    /// holds the element; the innermost control names it.
    label: Option<Label>,
}

#[derive(Clone)]
enum Label {
    Named(String),
    Unnamed,
}

impl Label {
    /// The content control's label: its alias, else its tag.
    fn of(node: &Node<'_>) -> Result<Self> {
        let Some(control) = node.content_control()? else {
            return Ok(Self::Unnamed);
        };
        let name = match control.alias()?.filter(|alias| !alias.is_empty()) {
            Some(alias) => Some(alias.to_owned()),
            None => control
                .tag()?
                .filter(|tag| !tag.is_empty())
                .map(str::to_owned),
        };
        Ok(name.map_or(Self::Unnamed, Self::Named))
    }

    /// The JSON `sdt` member: the name, or `true`.
    fn json(&self) -> SdtJson {
        match self {
            Self::Named(name) => SdtJson::Name(name.clone()),
            Self::Unnamed => SdtJson::Present(true),
        }
    }
}

impl fmt::Display for Label {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Named(name) => write!(f, "[sdt:{name}]"),
            Self::Unnamed => f.write_str("[sdt]"),
        }
    }
}

/// The body's node.
fn body(docx: &Docx) -> Result<Node<'_>> {
    nav::body_node(docx)
}

/// Visits each listed element of the flattened body in visible order:
/// paragraphs, tables, display equations and section breaks, with the
/// blocks of each content control in its place. Stops when `visit` breaks.
fn for_each_line<'f>(
    docx: &'f Docx,
    mut visit: impl FnMut(Line<'f>) -> Result<ControlFlow<()>>,
) -> Result<()> {
    let body = body(docx)?;
    let mut stack: Vec<(Children<'f>, Option<Label>)> =
        vec![(Children::of(&body, listed_kinds())?, None)];
    while let Some((children, label)) = stack.last_mut() {
        let Some(child) = children.next() else {
            stack.pop();
            continue;
        };
        let node = child?;
        if node.element_kind() == Kind::Sdt {
            let label = Some(Label::of(&node)?);
            let children = Children::of(&node, listed_kinds())?;
            stack.push((children, label));
            continue;
        }
        let label = label.clone();
        if visit(Line { node, label })?.is_break() {
            break;
        }
    }
    Ok(())
}

/// The number of listed elements of the flattened body: OfficeCLI's
/// `totalElements`. Reads the block indexes only.
fn line_count(docx: &Docx) -> Result<usize> {
    let mut count = 0;
    let mut stack: Vec<Blocks> = vec![nav::load_body(docx)?];
    while let Some(blocks) = stack.pop() {
        for block in blocks.iter() {
            match block.kind() {
                BlockKind::ContentControl => {
                    if let Some(control) = block.content_control() {
                        stack.push(control.blocks()?);
                    }
                },
                kind if Kind::of_block(kind).is_some() => count += 1,
                _ => {},
            }
        }
    }
    Ok(count)
}

/// The rows of a table node and the cells of its first row.
fn table_shape(node: &Node<'_>) -> Result<(usize, usize)> {
    let Some(table) = node.table()? else {
        return Ok((0, 0));
    };
    let rows = table.rows()?;
    let columns = match rows.first() {
        Some(row) => row.cell_count()?,
        None => 0,
    };
    Ok((rows.len(), columns))
}

/// Applies `start`, `end` and `max_lines` to a stream of lines: `Skip` the
/// line, `Show` it, or `Stop` (with whether the cap was reached).
struct Window {
    start: Option<u32>,
    end: Option<u32>,
    max: Option<u32>,
    line: u32,
    emitted: u32,
}

enum Step {
    Skip,
    Show,
    End,
    Capped,
}

impl Window {
    fn new(options: &Options) -> Self {
        Self {
            start: options.start,
            end: options.end,
            max: options.max_lines,
            line: 0,
            emitted: 0,
        }
    }

    /// The next line's fate; counts it.
    fn next(&mut self) -> Step {
        self.line = self.line.saturating_add(1);
        if self.start.is_some_and(|start| self.line < start) {
            return Step::Skip;
        }
        if self.end.is_some_and(|end| self.line > end) {
            return Step::End;
        }
        if self.max.is_some_and(|max| self.emitted >= max) {
            return Step::Capped;
        }
        Step::Show
    }
}

// ---------------------------------------------------------------------------
// text
// ---------------------------------------------------------------------------

fn text(docx: &Docx, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let mut window = Window::new(options);
    let mut capped = false;
    for_each_line(docx, |line| {
        let kind = line.node.element_kind();
        match window.next() {
            Step::Skip => return Ok(ControlFlow::Continue(())),
            Step::End => return Ok(ControlFlow::Break(())),
            Step::Capped => {
                capped = true;
                return Ok(ControlFlow::Break(()));
            },
            Step::Show => {},
        }
        // Section breaks are layout, not content: counted, never shown.
        if kind == Kind::Section {
            return Ok(ControlFlow::Continue(()));
        }
        write!(out, "[{}] ", line.node.canonical_path()).map_err(sink)?;
        if let Some(label) = &line.label {
            write!(out, "{label} ").map_err(sink)?;
        }
        match kind {
            Kind::Table => {
                let (rows, _) = table_shape(&line.node)?;
                writeln!(out, "[Table: {rows} rows]").map_err(sink)?;
            },
            Kind::Equation => match props::text(&line.node)? {
                Some(latex) => writeln!(out, "[Equation] {latex}").map_err(sink)?,
                None => writeln!(out, "[Equation]").map_err(sink)?,
            },
            _ => {
                let text = props::text(&line.node)?.unwrap_or_default();
                writeln!(out, "{text}").map_err(sink)?;
            },
        }
        window.emitted += 1;
        Ok(ControlFlow::Continue(()))
    })?;
    if capped {
        writeln!(
            out,
            "... (showed {} rows, {} total, use --start/--end to view more)",
            window.emitted,
            line_count(docx)?
        )
        .map_err(sink)?;
    }
    Ok(())
}

fn text_json(docx: &Docx, options: &Options) -> Result<TextJson> {
    let mut window = Window::new(options);
    let mut elements = Vec::new();
    for_each_line(docx, |line| {
        let kind = line.node.element_kind();
        match window.next() {
            Step::Skip => return Ok(ControlFlow::Continue(())),
            Step::End | Step::Capped => return Ok(ControlFlow::Break(())),
            Step::Show => {},
        }
        if kind == Kind::Section {
            return Ok(ControlFlow::Continue(()));
        }
        let text = match kind {
            Kind::Table => Some(format!("[Table: {} rows]", table_shape(&line.node)?.0)),
            _ => props::text(&line.node)?.map(|text| text.into_owned()),
        };
        elements.push(TextElement {
            path: line.node.canonical_path().to_string(),
            kind: kind.type_name(),
            text,
            sdt: line.label.as_ref().map(Label::json),
        });
        window.emitted += 1;
        Ok(ControlFlow::Continue(()))
    })?;
    Ok(TextJson {
        total_elements: line_count(docx)?,
        elements,
    })
}

// ---------------------------------------------------------------------------
// annotated
// ---------------------------------------------------------------------------

fn annotated(docx: &Docx, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let mut window = Window::new(options);
    let mut capped = false;
    for_each_line(docx, |line| {
        let kind = line.node.element_kind();
        match window.next() {
            Step::Skip => return Ok(ControlFlow::Continue(())),
            Step::End => return Ok(ControlFlow::Break(())),
            Step::Capped => {
                capped = true;
                return Ok(ControlFlow::Break(()));
            },
            Step::Show => {},
        }
        let path = line.node.canonical_path();
        match kind {
            Kind::Section => return Ok(ControlFlow::Continue(())),
            Kind::Table => {
                let (rows, columns) = table_shape(&line.node)?;
                writeln!(out, "[{path}] [Table: {rows}\u{d7}{columns}]").map_err(sink)?;
            },
            Kind::Equation => match props::text(&line.node)? {
                Some(latex) => {
                    writeln!(out, "[{path}] [Equation: \"{latex}\"] \u{2190} display")
                        .map_err(sink)?;
                },
                None => writeln!(out, "[{path}] [Equation] \u{2190} display").map_err(sink)?,
            },
            _ => annotate_paragraph(&line, out)?,
        }
        window.emitted += 1;
        Ok(ControlFlow::Continue(()))
    })?;
    if capped {
        writeln!(
            out,
            "... (showed {} rows, {} total, use --start/--end to view more)",
            window.emitted,
            line_count(docx)?
        )
        .map_err(sink)?;
    }
    Ok(())
}

fn annotate_paragraph(line: &Line<'_>, out: &mut impl fmt::Write) -> Result<()> {
    let Value::Paragraph { paragraph, .. } = &line.node.value else {
        return Ok(());
    };
    let path = line.node.canonical_path();
    let style = props::style(&line.node)?.unwrap_or_default();
    let runs = paragraph.addressable_runs()?;
    if runs.is_empty() {
        let suffix = match &line.label {
            Some(label) => format!(" | {label}"),
            None => String::new(),
        };
        writeln!(out, "[{path}] [] <- {style} | empty paragraph{suffix}").map_err(sink)?;
        return Ok(());
    }
    for run in &runs {
        if let Some(image) = first_image(run)? {
            writeln!(out, "[{path}] [Image: {image}] \u{2190} {style}").map_err(sink)?;
            continue;
        }
        let text = run.text()?;
        let direct = run.direct_properties()?;
        props::require_valid(direct.invalid_properties().next().is_some())?;
        let format = run_format(&direct);
        write!(
            out,
            "[{path}] \u{300c}{text}\u{300d} \u{2190} {style} | {format}"
        )
        .map_err(sink)?;
        if let Some(label) = &line.label {
            write!(out, " | {label}").map_err(sink)?;
        }
        writeln!(out).map_err(sink)?;
    }
    Ok(())
}

/// The first picture of a run, described as OfficeCLI's `GetDrawingInfo`
/// does: `alt="…"` or `name="…"`, then the extent in centimetres.
fn first_image(run: &Run) -> Result<Option<String>> {
    for content in run.contents()? {
        if let RunContent::Image(image) = content {
            let mut info = String::new();
            if !image.description().is_empty() {
                let _ = write!(info, "alt=\"{}\"", image.description());
            } else if !image.name().is_empty() {
                let _ = write!(info, "name=\"{}\"", image.name());
            }
            if !info.is_empty() {
                info.push_str(", ");
            }
            #[allow(
                clippy::cast_precision_loss,
                reason = "picture extents are far below 2^53 EMU"
            )]
            let (width, height) = (
                image.width_emu() as f64 / 360_000.0,
                image.height_emu() as f64 / 360_000.0,
            );
            let _ = write!(info, "{width:.1}cm\u{d7}{height:.1}cm");
            return Ok(Some(info));
        }
    }
    Ok(None)
}

/// A run's own formatting as OfficeCLI's `GetRunFormatDescription` lists
/// it: font, size, then `bold`, `italic`, `underline`, `strikethrough`, or
/// `(default)`.
fn run_format(props: &direct::Run) -> String {
    let mut parts: Vec<String> = Vec::new();
    if let Some(font) = run_font(props) {
        parts.push(font.to_owned());
    }
    if let Some(size) = props.size {
        parts.push(points(size));
    }
    if props.bold == Some(true) {
        parts.push("bold".to_owned());
    }
    if props.italic == Some(true) {
        parts.push("italic".to_owned());
    }
    if props
        .underline
        .is_some_and(|underline| underline.style != litchi_docx::UnderlineStyle::None)
    {
        parts.push("underline".to_owned());
    }
    if props.strike == Some(true) {
        parts.push("strikethrough".to_owned());
    }
    if parts.is_empty() {
        "(default)".to_owned()
    } else {
        parts.join(" ")
    }
}

/// OfficeCLI's `GetFontFromProperties`: East Asian, then ASCII, then
/// high-ANSI.
fn run_font(props: &direct::Run) -> Option<&str> {
    let fonts = props.fonts.as_ref()?;
    [&fonts.east_asia, &fonts.ascii, &fonts.h_ansi]
        .into_iter()
        .find_map(|font| font.as_deref().filter(|font| !font.is_empty()))
}

/// A size in half points as `"10.5pt"`.
fn points(size: HalfPoints) -> String {
    if size.0 % 2 == 0 {
        format!("{}pt", size.0 / 2)
    } else {
        format!("{}.5pt", size.0 / 2)
    }
}

// ---------------------------------------------------------------------------
// The body's paragraphs, for outline and stats
// ---------------------------------------------------------------------------

/// Visits each paragraph of the flattened body, display-equation paragraphs
/// included (OfficeCLI's `GetBodyElements(...).OfType<Paragraph>()`), with
/// its 1-based number.
fn for_each_paragraph<'f>(
    docx: &'f Docx,
    mut visit: impl FnMut(u32, &Node<'f>) -> Result<()>,
) -> Result<(u32, u32, u32)> {
    let (mut paragraphs, mut tables, mut equations) = (0_u32, 0_u32, 0_u32);
    for_each_line(docx, |line| {
        match line.node.element_kind() {
            Kind::Table => tables += 1,
            Kind::Equation => {
                equations += 1;
                // A display equation held in a `w:p` is a paragraph to
                // OfficeCLI's counts; a bare `m:oMathPara` is not.
                if line
                    .node
                    .block()?
                    .is_some_and(|block| block.paragraph().is_some())
                {
                    paragraphs += 1;
                    visit(paragraphs, &line.node)?;
                }
            },
            Kind::Paragraph => {
                paragraphs += 1;
                visit(paragraphs, &line.node)?;
            },
            _ => {},
        }
        Ok(ControlFlow::Continue(()))
    })?;
    Ok((paragraphs, tables, equations))
}

/// The style of a paragraph node, or of the `w:p` holding a display
/// equation: its style id, or `Normal` (OfficeCLI's `GetStyleName`, by id).
fn paragraph_style(node: &Node<'_>) -> Result<String> {
    Ok(match &node.value {
        Value::Paragraph { paragraph, .. } => paragraph.style_id()?,
        _ => match node.block()?.and_then(|block| block.paragraph()) {
            Some(paragraph) => paragraph.style_id()?,
            None => None,
        },
    }
    .unwrap_or_else(|| "Normal".to_owned()))
}

/// The visible text of a paragraph node, or of the `w:p` holding a display
/// equation (whose math is not text, as OfficeCLI's `GetParagraphText`
/// reads it).
fn paragraph_text(node: &Node<'_>) -> Result<String> {
    Ok(match &node.value {
        Value::Paragraph { paragraph, .. } => paragraph.text()?,
        _ => match node.block()?.and_then(|block| block.paragraph()) {
            Some(paragraph) => paragraph.text()?,
            None => String::new(),
        },
    })
}

// ---------------------------------------------------------------------------
// outline
// ---------------------------------------------------------------------------

/// The outline view: counts and the heading tree.
#[derive(Debug, Clone)]
struct Outline {
    paragraphs: u32,
    tables: u32,
    images: usize,
    equations: u32,
    content_controls: usize,
    headings: Vec<Heading>,
}

#[derive(Debug, Clone, Serialize)]
struct Heading {
    line: u32,
    text: String,
    style: String,
    level: u8,
}

fn outline(docx: &Docx) -> Result<Outline> {
    let mut headings = Vec::new();
    let (paragraphs, tables, equations) = for_each_paragraph(docx, |line, node| {
        let Some(level) = heading_level(node)? else {
            return Ok(());
        };
        let text = paragraph_text(node)?;
        if text.trim().is_empty() {
            return Ok(());
        }
        headings.push(Heading {
            line,
            text,
            style: paragraph_style(node)?,
            level,
        });
        Ok(())
    })?;
    let (images, content_controls) = main_story_counts(docx)?;
    Ok(Outline {
        paragraphs,
        tables,
        images,
        equations,
        content_controls,
        headings,
    })
}

/// The outline level OfficeCLI's `GetParagraphOutlineLevel` gives a
/// paragraph: its own `w:outlineLvl` (0-8, as level 1-9), else a level
/// from its style: a `Heading`/`heading`/`标题` style's first digit,
/// `Title` 0, `Subtitle` 1. Styles are read by id (M1-Δ4).
fn heading_level(node: &Node<'_>) -> Result<Option<u8>> {
    let Value::Paragraph { paragraph, .. } = &node.value else {
        return Ok(None);
    };
    let props = paragraph.direct_properties()?;
    props::require_valid(
        props.invalid_properties().next().is_some()
            || props
                .mark
                .as_ref()
                .is_some_and(|mark| mark.invalid_properties().next().is_some()),
    )?;
    if let Some(level) = props.outline_level.filter(|level| *level <= 8) {
        return Ok(Some(level + 1));
    }
    let Some(style) = props.style.as_deref() else {
        return Ok(None);
    };
    if style.contains("Heading")
        || style.contains("\u{6807}\u{9898}")
        || style
            .get(..7)
            .is_some_and(|head| head.eq_ignore_ascii_case("heading"))
    {
        // `GetHeadingLevel`: the first digit, else 1.
        let digit = style.chars().find_map(|c| c.to_digit(10));
        return Ok(Some(
            digit.map_or(1, |digit| u8::try_from(digit).unwrap_or(1)),
        ));
    }
    Ok(match style {
        "Title" => Some(0),
        "Subtitle" => Some(1),
        _ => None,
    })
}

/// The pictures and content controls of the whole main story, tables
/// included (OfficeCLI counts every `w:drawing` and every block and inline
/// `w:sdt` of the body).
fn main_story_counts(docx: &Docx) -> Result<(usize, usize)> {
    let paragraphs = match docx.owner() {
        super::DocxOwner::Eager(package) => package.document()?.paragraphs()?,
        super::DocxOwner::Source(package) => package.document()?.paragraphs()?.into_vec(),
    };
    let mut images = 0;
    let mut inline_controls = 0;
    for paragraph in &paragraphs {
        images += paragraph.images()?.len();
        inline_controls += paragraph
            .inlines()?
            .iter()
            .filter(|inline| {
                matches!(inline, litchi_docx::Inline::Unknown(opaque)
                    if opaque.kind() == litchi_docx::InlineKind::ContentControl)
            })
            .count();
    }
    Ok((images, inline_controls + block_controls(docx)?))
}

/// The block content controls of the body, nested ones included.
fn block_controls(docx: &Docx) -> Result<usize> {
    let mut count = 0;
    let mut stack: Vec<Blocks> = vec![nav::load_body(docx)?];
    while let Some(blocks) = stack.pop() {
        for block in blocks.iter() {
            match block.kind() {
                BlockKind::ContentControl => {
                    count += 1;
                    if let Some(control) = block.content_control() {
                        stack.push(control.blocks()?);
                    }
                },
                BlockKind::Table => {
                    if let Some(table) = block.table() {
                        for row in table.rows()? {
                            for cell in row.cells()? {
                                stack.push(cell.blocks()?);
                            }
                        }
                    }
                },
                _ => {},
            }
        }
    }
    Ok(count)
}

impl Outline {
    fn write(&self, out: &mut impl fmt::Write) -> fmt::Result {
        write!(
            out,
            "{} paragraphs | {} tables | {} images",
            self.paragraphs, self.tables, self.images
        )?;
        if self.equations > 0 {
            write!(out, " | {} equations", self.equations)?;
        }
        if self.content_controls > 0 {
            write!(out, " | {} content controls", self.content_controls)?;
        }
        writeln!(out)?;
        writeln!(out)?;
        for heading in &self.headings {
            if heading.level > 1 {
                for _ in 0..(heading.level - 1) * 2 {
                    out.write_char(' ')?;
                }
            }
            let prefix = if heading.level == 0 {
                "\u{25a0}"
            } else {
                "\u{251c}\u{2500}\u{2500}"
            };
            writeln!(
                out,
                "{prefix} [{}] \"{}\" ({})",
                heading.line, heading.text, heading.style
            )?;
        }
        Ok(())
    }

    fn json(self) -> OutlineJson {
        OutlineJson {
            paragraphs: self.paragraphs,
            tables: self.tables,
            images: self.images,
            equations: self.equations,
            content_controls: (self.content_controls > 0).then_some(self.content_controls),
            headings: self.headings,
        }
    }
}

// ---------------------------------------------------------------------------
// stats
// ---------------------------------------------------------------------------

/// The stats view's counts and distributions.
#[derive(Debug, Clone, Default)]
struct Stats {
    paragraphs: u32,
    words: usize,
    characters: usize,
    empty: usize,
    double_spaces: usize,
    styles: Distribution,
    fonts: Distribution,
    sizes: Distribution,
}

fn stats(docx: &Docx) -> Result<Stats> {
    let mut stats = Stats::default();
    let (paragraphs, _, _) = for_each_paragraph(docx, |_, node| {
        stats.styles.add(paragraph_style(node)?);
        let text = paragraph_text(node)?;
        stats.words += text.split_whitespace().count();
        let runs = match &node.value {
            Value::Paragraph { paragraph, .. } => paragraph.addressable_runs()?.into_vec(),
            _ => Vec::new(),
        };
        // An equation's paragraph holds math, so it is not empty.
        let equation = node.element_kind() == Kind::Equation;
        if !equation && runs.is_empty() && text.trim().is_empty() {
            stats.empty += 1;
            return Ok(());
        }
        for run in &runs {
            let text = run.text()?;
            // OfficeCLI counts .NET string lengths: UTF-16 code units.
            stats.characters += text.encode_utf16().count();
            if text.contains("  ") {
                stats.double_spaces += 1;
            }
            let props = run.direct_properties()?;
            props::require_valid(props.invalid_properties().next().is_some())?;
            stats
                .fonts
                .add(run_font(&props).unwrap_or("(default)").to_owned());
            stats
                .sizes
                .add(props.size.map_or_else(|| "(default)".to_owned(), points));
        }
        Ok(())
    })?;
    stats.paragraphs = paragraphs;
    stats.styles.sort();
    stats.fonts.sort();
    stats.sizes.sort();
    Ok(stats)
}

impl Stats {
    fn write(&self, out: &mut impl fmt::Write) -> fmt::Result {
        writeln!(
            out,
            "Paragraphs: {} | Words: {} | Total Characters: {}",
            self.paragraphs, self.words, self.characters
        )?;
        for (title, distribution) in [
            ("Style Distribution:", &self.styles),
            ("Font Usage:", &self.fonts),
            ("Font Size Usage:", &self.sizes),
        ] {
            writeln!(out)?;
            writeln!(out, "{title}")?;
            for (name, count) in &distribution.entries {
                writeln!(out, "  {name}: {count}")?;
            }
        }
        writeln!(out)?;
        writeln!(out, "Empty Paragraphs: {}", self.empty)?;
        writeln!(out, "Consecutive Spaces: {}", self.double_spaces)
    }

    fn json(self) -> StatsJson {
        StatsJson {
            paragraphs: self.paragraphs,
            words: self.words,
            total_characters: self.characters,
            empty_paragraphs: self.empty,
            consecutive_spaces: self.double_spaces,
            style_distribution: self.styles,
            font_usage: self.fonts,
            font_size_usage: self.sizes,
        }
    }
}

/// Names and their counts, in first-seen order until sorted by count,
/// descending and stable (LINQ's `OrderByDescending`).
#[derive(Debug, Clone, Default)]
struct Distribution {
    entries: Vec<(String, usize)>,
    index: HashMap<String, usize>,
}

impl Distribution {
    fn add(&mut self, name: String) {
        match self.index.get(&name) {
            Some(&at) => self.entries[at].1 += 1,
            None => {
                self.index.insert(name.clone(), self.entries.len());
                self.entries.push((name, 1));
            },
        }
    }

    fn sort(&mut self) {
        self.entries.sort_by_key(|entry| std::cmp::Reverse(entry.1));
        self.index.clear();
    }
}

impl Serialize for Distribution {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(self.entries.len()))?;
        for (name, count) in &self.entries {
            map.serialize_entry(name, count)?;
        }
        map.end()
    }
}

// ---------------------------------------------------------------------------
// JSON
// ---------------------------------------------------------------------------

/// A DOCX view as OfficeCLI's JSON (record 0010 §6.1).
#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
pub(crate) enum Json {
    /// `{"totalElements":N,"elements":[{"path","type","text","sdt"}]}`.
    Text(TextJson),
    /// `{"view":"annotated","content":"..."}`.
    Annotated(Annotated),
    /// `{"paragraphs","tables","images","equations","headings":[...]}`.
    Outline(OutlineJson),
    /// `{"paragraphs","words","totalCharacters",...,"styleDistribution",...}`.
    Stats(StatsJson),
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct TextJson {
    #[serde(rename = "totalElements")]
    total_elements: usize,
    elements: Vec<TextElement>,
}

#[derive(Debug, Clone, Serialize)]
struct TextElement {
    path: String,
    #[serde(rename = "type")]
    kind: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    sdt: Option<SdtJson>,
}

/// The `sdt` member: the control's alias or tag, else `true`.
#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
enum SdtJson {
    Name(String),
    Present(bool),
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct Annotated {
    view: &'static str,
    content: String,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct OutlineJson {
    paragraphs: u32,
    tables: u32,
    images: usize,
    equations: u32,
    #[serde(rename = "contentControls", skip_serializing_if = "Option::is_none")]
    content_controls: Option<usize>,
    headings: Vec<Heading>,
}

#[derive(Debug, Clone, Serialize)]
pub(crate) struct StatsJson {
    paragraphs: u32,
    words: usize,
    #[serde(rename = "totalCharacters")]
    total_characters: usize,
    #[serde(rename = "emptyParagraphs")]
    empty_paragraphs: usize,
    #[serde(rename = "consecutiveSpaces")]
    consecutive_spaces: usize,
    #[serde(rename = "styleDistribution")]
    style_distribution: Distribution,
    #[serde(rename = "fontUsage")]
    font_usage: Distribution,
    #[serde(rename = "fontSizeUsage")]
    font_size_usage: Distribution,
}

// ---------------------------------------------------------------------------
// Output
// ---------------------------------------------------------------------------

/// A sink that drops trailing white space of the whole output, as
/// OfficeCLI's `TrimEnd()` does, without buffering more than the pending
/// white space.
struct TrimEnd<'w, W> {
    out: &'w mut W,
    pending: String,
}

impl<'w, W: fmt::Write> TrimEnd<'w, W> {
    fn new(out: &'w mut W) -> Self {
        Self {
            out,
            pending: String::new(),
        }
    }
}

impl<W: fmt::Write> fmt::Write for TrimEnd<'_, W> {
    fn write_str(&mut self, text: &str) -> fmt::Result {
        let kept = text.trim_end();
        if kept.is_empty() {
            self.pending.push_str(text);
            return Ok(());
        }
        if !self.pending.is_empty() {
            self.out.write_str(&self.pending)?;
            self.pending.clear();
        }
        self.out.write_str(kept)?;
        self.pending.push_str(&text[kept.len()..]);
        Ok(())
    }
}
