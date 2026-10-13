//! Legacy PPT semantic views (record 0010 §6).
//!
//! The binary PowerPoint leaf exposes a lazy slide and shape projection rather
//! than a PresentationML scene.  These views therefore walk that projection
//! directly.  They preserve the four semantic modes shared by the OOXML
//! presentation adapter, while reporting facts only where the PPT leaf has a
//! typed value: pictures report their `wzDescription` alternative text, and
//! the legacy leaf has no chart model.

use std::fmt;

use litchi_ppt::TextBox;
use litchi_ppt::shapes::{PictureFrameKind, ShapeEnum};
use serde::Serialize;

use super::Ppt;
use crate::error::{Code, Error, Result};
use crate::view::mode::{Mode, Options};

/// Writes the `mode` view of a legacy PowerPoint presentation into `out`.
///
/// Text and annotated views apply the slide selectors in `options`; outline
/// and stats inspect every slide.  The source is fenced before and after each
/// selected slide so a positional source changed during a view is reported as
/// a typed I/O error by the leaf owner.
///
/// # Errors
///
/// Returns `invalid_argument` for workbook-only options, the leaf's typed
/// error for a failed slide or shape read, `unsupported_capability` for a
/// rendering mode outside the semantic view surface, and `io_error` when the
/// output sink refuses a write.
pub(crate) fn view_to<W: fmt::Write>(
    file: &Ppt,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    file.check_source()?;
    check_options(options)?;
    let mut out = TrimEnd::new(out);
    match mode {
        Mode::Text => text_view(file, options, &mut out)?,
        Mode::Annotated => annotated(file, options, &mut out)?,
        Mode::Outline => outline(file, &mut out)?,
        Mode::Stats => stats(file)?.write(&mut out).map_err(sink)?,
        _ => return Err(unsupported(mode)),
    }
    file.check_source()?;
    Ok(())
}

/// The `mode` view as its OfficeCLI-shaped JSON projection.
///
/// # Errors
///
/// As [`view_to`].
pub(crate) fn view_json(file: &Ppt, mode: Mode, options: &Options) -> Result<Json> {
    file.check_source()?;
    check_options(options)?;
    let json = match mode {
        Mode::Text => Json::Text(text_json(file, options)?),
        Mode::Annotated => {
            let mut content = String::new();
            view_to(file, mode, options, &mut content)?;
            Json::Annotated(Annotated {
                view: "annotated",
                content,
            })
        },
        Mode::Outline => Json::Outline(outline_json(file)?),
        Mode::Stats => Json::Stats(stats(file)?.json()),
        _ => return Err(unsupported(mode)),
    };
    file.check_source()?;
    Ok(json)
}

fn unsupported(mode: Mode) -> Error {
    mode.unsupported().with_format(litchi_core::FileFormat::Ppt)
}

/// `columns` and `range` select workbook cells, and have no PPT meaning.
fn check_options(options: &Options) -> Result<()> {
    if options.columns.is_some() || options.range.is_some() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the columns and range view options apply to workbooks only",
        )
        .with_format(litchi_core::FileFormat::Ppt)
        .with_suggestion("select slides with start, end and max_lines"));
    }
    Ok(())
}

fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the PPT view output sink refused to write")
        .with_format(litchi_core::FileFormat::Ppt)
}

// Keep recursive walks finite even when a legacy shape tree is deeply nested
// or contains an unexpectedly large group. The budget is per slide, so a
// large presentation remains inspectable without allowing one slide to
// monopolize the view operation.
const MAX_VIEW_SHAPES: usize = 1_000_000;
const MAX_VIEW_DEPTH: usize = 256;

fn enter_shape(depth: usize, seen: &mut usize) -> Result<()> {
    if depth > MAX_VIEW_DEPTH {
        return Err(view_limit("PPT view shape nesting exceeded"));
    }
    *seen = seen
        .checked_add(1)
        .ok_or_else(|| view_limit("PPT view shape count overflowed"))?;
    if *seen > MAX_VIEW_SHAPES {
        return Err(view_limit("PPT view contains too many shapes"));
    }
    Ok(())
}

fn view_limit(message: &'static str) -> Error {
    Error::new(Code::LimitExceeded, message).with_format(litchi_core::FileFormat::Ppt)
}

/// Visit selected slides without materializing the complete slide catalog.
fn each_slide<W: fmt::Write>(
    file: &Ppt,
    options: &Options,
    out: &mut W,
    mut visit: impl FnMut(u32, &[ShapeEnum<'static>], &mut W) -> Result<()>,
) -> Result<()> {
    let total = file.reader().slide_count();
    let start = options.start.unwrap_or(1);
    for position in 0..total {
        let number = u32::try_from(position + 1).unwrap_or(u32::MAX);
        if number < start {
            continue;
        }
        if options.end.is_some_and(|end| number > end) {
            break;
        }
        if options
            .max_lines
            .is_some_and(|max| number.saturating_sub(start) >= max)
        {
            writeln!(
                out,
                "... (showed {shown} of {total} slides, use --start/--end to see more)",
                shown = options.max_lines.unwrap_or(0)
            )
            .map_err(sink)?;
            break;
        }
        let slide = file
            .reader()
            .slide_at(position)?
            .ok_or_else(|| missing_slide(position))?;
        let shapes = slide.shapes()?;
        visit(number, shapes, out)?;
        file.check_source()?;
    }
    Ok(())
}

fn missing_slide(position: usize) -> Error {
    Error::new(
        Code::InternalError,
        format!("slide {} vanished from the catalog", position + 1),
    )
    .with_format(litchi_core::FileFormat::Ppt)
}

fn is_blank(text: &str) -> bool {
    text.trim().is_empty()
}

fn shape_text(shape: &ShapeEnum<'static>) -> Result<String> {
    shape.text().map_err(Error::from)
}

fn is_title(shape: &ShapeEnum<'_>) -> bool {
    matches!(shape, ShapeEnum::Placeholder(placeholder) if placeholder.is_title())
}

fn table_cells(table: &litchi_ppt::shapes::shape_enum::TableShape) -> impl Iterator<Item = &str> {
    (0..table.rows())
        .flat_map(move |row| (0..table.columns()).filter_map(move |column| table.cell(row, column)))
}

/// Emit the text-bearing leaves in the same order as the native shape tree.
/// Groups are containers, so their aggregate text is not emitted a second
/// time after their children.
fn emit_text_shape<W: fmt::Write>(
    shape: &ShapeEnum<'static>,
    depth: usize,
    seen: &mut usize,
    out: &mut W,
) -> Result<()> {
    enter_shape(depth, seen)?;
    match shape {
        shape @ (ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_)) => {
            let text = shape_text(shape)?;
            if !is_blank(&text) {
                writeln!(out, "{text}").map_err(sink)?;
            }
        },
        ShapeEnum::Table(table) => {
            for cell in table_cells(table) {
                if !is_blank(cell) {
                    writeln!(out, "{cell}").map_err(sink)?;
                }
            }
        },
        ShapeEnum::Group(group) => {
            for child in group.children() {
                emit_text_shape(child, depth + 1, seen, out)?;
            }
        },
        ShapeEnum::Picture(_) | ShapeEnum::Line(_) => {},
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// text
// ---------------------------------------------------------------------------

fn text_view(file: &Ppt, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    each_slide(file, options, out, |number, shapes, out| {
        writeln!(out, "=== /slide[{number}] ===").map_err(sink)?;
        let mut seen = 0;
        for shape in shapes {
            emit_text_shape(shape, 0, &mut seen, out)?;
        }
        writeln!(out).map_err(sink)
    })
}

/// `{"totalSlides":N,"slides":[{"index":k,"path":"/slide[k]","texts":[…]}]}`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextJson {
    total_slides: usize,
    slides: Vec<SlideTexts>,
}

#[derive(Debug, Clone, Serialize)]
struct SlideTexts {
    index: u32,
    path: String,
    texts: Vec<String>,
}

fn collect_texts(
    shape: &ShapeEnum<'static>,
    depth: usize,
    seen: &mut usize,
    texts: &mut Vec<String>,
) -> Result<()> {
    enter_shape(depth, seen)?;
    match shape {
        shape @ (ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_)) => {
            let text = shape_text(shape)?;
            if !is_blank(&text) {
                texts.push(text);
            }
        },
        ShapeEnum::Table(table) => {
            for cell in table_cells(table) {
                if !is_blank(cell) {
                    texts.push(cell.to_owned());
                }
            }
        },
        ShapeEnum::Group(group) => {
            for child in group.children() {
                collect_texts(child, depth + 1, seen, texts)?;
            }
        },
        ShapeEnum::Picture(_) | ShapeEnum::Line(_) => {},
    }
    Ok(())
}

fn text_json(file: &Ppt, options: &Options) -> Result<TextJson> {
    file.check_source()?;
    let total = file.reader().slide_count();
    let start = options.start.unwrap_or(1);
    let mut slides = Vec::new();
    for position in 0..total {
        let number = u32::try_from(position + 1).unwrap_or(u32::MAX);
        if number < start {
            continue;
        }
        if options.end.is_some_and(|end| number > end) {
            break;
        }
        if options
            .max_lines
            .is_some_and(|max| slides.len() >= max as usize)
        {
            break;
        }
        let slide = file
            .reader()
            .slide_at(position)?
            .ok_or_else(|| missing_slide(position))?;
        let mut texts = Vec::new();
        let mut seen = 0;
        for shape in slide.shapes()? {
            collect_texts(shape, 0, &mut seen, &mut texts)?;
        }
        slides.push(SlideTexts {
            index: number,
            path: format!("/slide[{number}]"),
            texts,
        });
        file.check_source()?;
    }
    Ok(TextJson {
        total_slides: total,
        slides,
    })
}

// ---------------------------------------------------------------------------
// annotated
// ---------------------------------------------------------------------------

fn annotated(file: &Ppt, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    each_slide(file, options, out, |number, shapes, out| {
        writeln!(out, "[/slide[{number}]]").map_err(sink)?;
        let mut seen = 0;
        annotate(shapes, 1, &mut seen, out)?;
        writeln!(out).map_err(sink)
    })
}

fn text_style(shape: &ShapeEnum<'_>) -> (Option<String>, Option<u16>) {
    let ShapeEnum::TextBox(textbox) = shape else {
        return (None, None);
    };
    let font = textbox
        .runs()
        .iter()
        .find_map(|run| run.formatting.font_name.clone());
    (font, textbox.font_size())
}

fn annotate<W: fmt::Write>(
    shapes: &[ShapeEnum<'static>],
    indent: usize,
    seen: &mut usize,
    out: &mut W,
) -> Result<()> {
    let pad = "  ".repeat(indent);
    for shape in shapes {
        annotate_shape(shape, indent, 0, seen, &pad, out)?;
    }
    Ok(())
}

fn annotate_shape<W: fmt::Write>(
    shape: &ShapeEnum<'static>,
    indent: usize,
    depth: usize,
    seen: &mut usize,
    pad: &str,
    out: &mut W,
) -> Result<()> {
    enter_shape(depth, seen)?;
    match shape {
        shape @ (ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_)) => {
            let text = shape_text(shape)?;
            if is_blank(&text) {
                return Ok(());
            }
            let kind = if is_title(shape) { "Title" } else { "Text Box" };
            let (font, size) = text_style(shape);
            let font = font.as_deref().unwrap_or("(unavailable)");
            let size = size.map_or_else(String::new, |size| format!("{size}pt"));
            writeln!(out, "{pad}[{kind}] \"{text}\" ← {font} {size}").map_err(sink)?;
        },
        ShapeEnum::Table(table) => {
            writeln!(
                out,
                "{pad}[Table] \"Table\" ← {}x{}",
                table.rows(),
                table.columns()
            )
            .map_err(sink)?;
        },
        ShapeEnum::Picture(picture) => {
            let named = super::node::shape_name(shape);
            let (kind, name) = match picture.frame_kind() {
                PictureFrameKind::Picture => ("Picture", named.unwrap_or("Picture")),
                PictureFrameKind::OleObject => ("OLE", named.unwrap_or("OLE")),
                PictureFrameKind::Media => ("Media", named.unwrap_or("Media")),
            };
            let alt = shape.description().filter(|alt| !alt.is_empty());
            match (picture.frame_kind(), alt) {
                (PictureFrameKind::Picture, Some(alt)) => {
                    writeln!(out, "{pad}[{kind}] \"{name}\" ← alt=\"{alt}\"")
                },
                (PictureFrameKind::Picture, None) => {
                    writeln!(out, "{pad}[{kind}] \"{name}\" ← ⚠ no alt text")
                },
                _ => writeln!(out, "{pad}[{kind}] \"{name}\" ← ⚠ preview metadata only"),
            }
            .map_err(sink)?;
        },
        ShapeEnum::Group(group) => {
            let items = group.children().len();
            writeln!(out, "{pad}[Group] \"Group\" ← {items} item(s)").map_err(sink)?;
            let child_pad = "  ".repeat(indent + 1);
            for child in group.children() {
                annotate_shape(child, indent + 1, depth + 1, seen, &child_pad, out)?;
            }
        },
        ShapeEnum::Line(_) => {},
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// outline
// ---------------------------------------------------------------------------

/// One slide of the outline.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct SlideOutline {
    index: u32,
    title: Option<String>,
    text_boxes: usize,
    pictures: usize,
    ole_objects: usize,
}

fn outline_shape(
    shape: &ShapeEnum<'static>,
    depth: usize,
    seen: &mut usize,
    title: &mut Option<String>,
    text_boxes: &mut usize,
    pictures: &mut usize,
    ole_objects: &mut usize,
) -> Result<()> {
    enter_shape(depth, seen)?;
    match shape {
        shape @ (ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_)) => {
            let text = shape_text(shape)?;
            if is_title(shape) {
                if title.is_none() && !is_blank(&text) {
                    *title = Some(text);
                }
            } else if !is_blank(&text) {
                *text_boxes += 1;
            }
        },
        ShapeEnum::Picture(picture) => match picture.frame_kind() {
            PictureFrameKind::Picture | PictureFrameKind::Media => *pictures += 1,
            PictureFrameKind::OleObject => *ole_objects += 1,
        },
        ShapeEnum::Group(group) => {
            for child in group.children() {
                outline_shape(
                    child,
                    depth + 1,
                    seen,
                    title,
                    text_boxes,
                    pictures,
                    ole_objects,
                )?;
            }
        },
        ShapeEnum::Table(_) | ShapeEnum::Line(_) => {},
    }
    Ok(())
}

fn slide_outlines(file: &Ppt) -> Result<Vec<SlideOutline>> {
    file.check_source()?;
    let total = file.reader().slide_count();
    let mut slides = Vec::with_capacity(total);
    for position in 0..total {
        let slide = file
            .reader()
            .slide_at(position)?
            .ok_or_else(|| missing_slide(position))?;
        let mut title = None;
        let (mut text_boxes, mut pictures, mut ole_objects) = (0, 0, 0);
        let mut seen = 0;
        for shape in slide.shapes()? {
            outline_shape(
                shape,
                0,
                &mut seen,
                &mut title,
                &mut text_boxes,
                &mut pictures,
                &mut ole_objects,
            )?;
        }
        slides.push(SlideOutline {
            index: u32::try_from(position + 1).unwrap_or(u32::MAX),
            title,
            text_boxes,
            pictures,
            ole_objects,
        });
        file.check_source()?;
    }
    Ok(slides)
}

fn outline(file: &Ppt, out: &mut impl fmt::Write) -> Result<()> {
    let slides = slide_outlines(file)?;
    writeln!(out, "{} slides", slides.len()).map_err(sink)?;
    for slide in &slides {
        let mut details = Vec::new();
        if slide.text_boxes > 0 {
            details.push(format!("{} text box(es)", slide.text_boxes));
        }
        if slide.pictures > 0 {
            details.push(format!("{} picture(s)", slide.pictures));
        }
        if slide.ole_objects > 0 {
            details.push(format!("{} ole object(s)", slide.ole_objects));
        }
        let details = if details.is_empty() {
            String::new()
        } else {
            format!(" - {}", details.join(", "))
        };
        writeln!(
            out,
            "├── Slide {}: \"{}\"{details}",
            slide.index,
            slide.title.as_deref().unwrap_or("(untitled)")
        )
        .map_err(sink)?;
    }
    Ok(())
}

/// `{"totalSlides":N,"slides":[{"index","title","textBoxes","pictures","oleObjects"}]}`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OutlineJson {
    total_slides: usize,
    slides: Vec<SlideOutline>,
}

fn outline_json(file: &Ppt) -> Result<OutlineJson> {
    let slides = slide_outlines(file)?;
    Ok(OutlineJson {
        total_slides: slides.len(),
        slides,
    })
}

// ---------------------------------------------------------------------------
// stats
// ---------------------------------------------------------------------------

/// Counts exposed by the stats view. The chart count is omitted because the
/// PPT adapter has no typed chart projection; that gap is advertised by its
/// format specification rather than inferred from a generic frame. Pictures
/// without alternate text count media frames too, as in the PPTX adapter's
/// picture count.
#[derive(Debug, Clone, Default)]
struct Stats {
    slides: usize,
    shapes: usize,
    text_boxes: usize,
    pictures: usize,
    ole_objects: usize,
    words: usize,
    slides_without_title: usize,
    pictures_without_alt: usize,
    fonts: Vec<(String, usize)>,
}

fn words(text: &str) -> usize {
    text.split_whitespace().count()
}

fn record_font(stats: &mut Stats, textbox: &TextBox<'_>) {
    for run in textbox.runs() {
        let Some(font) = run.formatting.font_name.as_deref() else {
            continue;
        };
        match stats.fonts.iter_mut().find(|(name, _)| name == font) {
            Some((_, count)) => *count += 1,
            None => stats.fonts.push((font.to_owned(), 1)),
        }
    }
}

fn walk_shape(
    shape: &ShapeEnum<'static>,
    depth: usize,
    seen: &mut usize,
    stats: &mut Stats,
    titled: &mut bool,
) -> Result<()> {
    enter_shape(depth, seen)?;
    stats.shapes += 1;
    match shape {
        shape @ (ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_)) => {
            stats.text_boxes += 1;
            let text = shape_text(shape)?;
            stats.words += words(&text);
            if is_title(shape) {
                *titled = true;
            }
            if let ShapeEnum::TextBox(textbox) = shape {
                record_font(stats, textbox);
            }
        },
        ShapeEnum::Table(table) => {
            for cell in table_cells(table) {
                stats.words += words(cell);
            }
        },
        ShapeEnum::Picture(picture) => match picture.frame_kind() {
            PictureFrameKind::Picture | PictureFrameKind::Media => {
                stats.pictures += 1;
                if shape.description().is_none_or(str::is_empty) {
                    stats.pictures_without_alt += 1;
                }
            },
            PictureFrameKind::OleObject => stats.ole_objects += 1,
        },
        ShapeEnum::Group(group) => {
            for child in group.children() {
                walk_shape(child, depth + 1, seen, stats, titled)?;
            }
        },
        ShapeEnum::Line(_) => {},
    }
    Ok(())
}

fn stats(file: &Ppt) -> Result<Stats> {
    file.check_source()?;
    let mut stats = Stats {
        slides: file.reader().slide_count(),
        ..Stats::default()
    };
    for position in 0..stats.slides {
        let slide = file
            .reader()
            .slide_at(position)?
            .ok_or_else(|| missing_slide(position))?;
        let mut titled = false;
        let mut seen = 0;
        for shape in slide.shapes()? {
            walk_shape(shape, 0, &mut seen, &mut stats, &mut titled)?;
        }
        if !titled {
            stats.slides_without_title += 1;
        }
        file.check_source()?;
    }
    stats
        .fonts
        .sort_by_key(|(_, count)| std::cmp::Reverse(*count));
    Ok(stats)
}

impl Stats {
    fn write(&self, out: &mut impl fmt::Write) -> fmt::Result {
        writeln!(out, "Slides: {}", self.slides)?;
        writeln!(out, "Total shapes: {}", self.shapes)?;
        writeln!(out, "Text boxes: {}", self.text_boxes)?;
        writeln!(out, "Pictures: {}", self.pictures)?;
        if self.ole_objects > 0 {
            writeln!(out, "OLE Objects: {}", self.ole_objects)?;
        }
        writeln!(out, "Words: {}", self.words)?;
        writeln!(out, "Slides without title: {}", self.slides_without_title)?;
        writeln!(
            out,
            "Pictures without alt text: {}",
            self.pictures_without_alt
        )?;
        if !self.fonts.is_empty() {
            writeln!(out)?;
            writeln!(out, "Font usage:")?;
            for (font, count) in &self.fonts {
                writeln!(out, "  {font}: {count} occurrence(s)")?;
            }
        }
        Ok(())
    }

    fn json(self) -> StatsJson {
        StatsJson {
            slides: self.slides,
            total_shapes: self.shapes,
            text_boxes: self.text_boxes,
            pictures: self.pictures,
            ole_objects: self.ole_objects,
            words: self.words,
            slides_without_title: self.slides_without_title,
            pictures_without_alt: self.pictures_without_alt,
            font_usage: (!self.fonts.is_empty()).then_some(FontUsage(self.fonts)),
        }
    }
}

/// `{"slides","totalShapes","textBoxes","pictures","oleObjects","words","slidesWithoutTitle","picturesWithoutAlt","fontUsage"}`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct StatsJson {
    slides: usize,
    total_shapes: usize,
    text_boxes: usize,
    pictures: usize,
    ole_objects: usize,
    words: usize,
    slides_without_title: usize,
    pictures_without_alt: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    font_usage: Option<FontUsage>,
}

/// Font names with their counts, in first-use order after count sorting.
#[derive(Debug, Clone)]
struct FontUsage(Vec<(String, usize)>);

impl Serialize for FontUsage {
    fn serialize<S: serde::Serializer>(
        &self,
        serializer: S,
    ) -> std::result::Result<S::Ok, S::Error> {
        use serde::ser::SerializeMap;
        let mut map = serializer.serialize_map(Some(self.0.len()))?;
        for (font, count) in &self.0 {
            map.serialize_entry(font, count)?;
        }
        map.end()
    }
}

/// `{"view":"annotated","content":"..."}`.
#[derive(Debug, Clone, Serialize)]
pub(crate) struct Annotated {
    view: &'static str,
    content: String,
}

/// A view as OfficeCLI's JSON.
#[derive(Debug, Clone, Serialize)]
#[serde(untagged)]
pub(crate) enum Json {
    /// The text view.
    Text(TextJson),
    /// The annotated view.
    Annotated(Annotated),
    /// The outline view.
    Outline(OutlineJson),
    /// The stats view.
    Stats(StatsJson),
}

/// A sink wrapper that drops trailing whitespace from the whole view, as
/// OfficeCLI does.  Whitespace is held until non-whitespace text follows it.
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
        let content = text.trim_end();
        if content.is_empty() {
            self.pending.push_str(text);
            return Ok(());
        }
        if !self.pending.is_empty() {
            self.out.write_str(&self.pending)?;
            self.pending.clear();
        }
        self.out.write_str(content)?;
        self.pending.push_str(&text[content.len()..]);
        Ok(())
    }
}
