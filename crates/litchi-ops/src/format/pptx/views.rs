//! PPTX views (record 0010 §6): the text, annotated, outline and stats
//! modes, with OfficeCLI's content and JSON shapes
//! (`PowerPointHandler.View.cs`).
//!
//! Views read each slide's shape scene and write straight into the caller's
//! sink. The text and annotated views count slides for `start`, `end` and
//! `max_lines`, skip the slides before `start` without reading them and stop
//! at `end`; the outline and stats views read every slide. Like OfficeCLI,
//! trailing white space of the whole output is dropped.
//!
//! | Mode | Content |
//! |---|---|
//! | text | `=== /slide[N] ===`, then the text of each shape (groups included), then each table cell's text |
//! | annotated | `[/slide[N]]`, then a tree of titles, text boxes (`← font size`), tables (`← RxC`), pictures (`← alt="…"` or `⚠ no alt text`) and groups |
//! | outline | `N slides`, then `├── Slide k: "title" - N text box(es), M picture(s), K ole object(s)` |
//! | stats | slides, shapes, text boxes, pictures, charts, OLE objects, words, slides without title, pictures without alt text, then font usage |
//!
//! OfficeCLI also prints the file name in the outline; the handle knows no
//! file name, so the first line counts the slides only. Equation shapes are
//! listed as text boxes, because the leaf does not detect Office math yet
//! (spec gap `equation`), and line breaks read as line feeds, the leaf's
//! shape text, where OfficeCLI writes vertical tabs. OLE objects are the
//! OLE frames of the shape trees; OfficeCLI also counts embedded parts that
//! no frame references.

use std::fmt;

use litchi_pptx::shape::Shape;
use litchi_pptx::shape::owned::{Handle as ShapeHandle, Handles, Scene as SharedScene};
use serde::Serialize;

use super::Pptx;
use super::node::is_title;
use super::text;
use crate::error::{Code, Error, Result};
use crate::view::mode::{Mode, Options};

/// Writes the `mode` view of the presentation into `out`.
///
/// # Errors
///
/// `unsupported_capability` for the modes Litchi does not render;
/// `invalid_argument` for workbook-only options (`columns`, `range`); the
/// leaf's error for a slide that fails to read; and `io_error` when the sink
/// refuses output.
pub(crate) fn view_to<W: fmt::Write>(
    file: &Pptx,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    file.finish(|| view_to_unfenced(file, mode, options, out))
}

fn view_to_unfenced<W: fmt::Write>(
    file: &Pptx,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    check_options(options)?;
    let mut out = TrimEnd::new(out);
    match mode {
        Mode::Text => text_view(file, options, &mut out)?,
        Mode::Annotated => annotated(file, options, &mut out)?,
        Mode::Outline => outline(file, &mut out)?,
        Mode::Stats => stats(file)?.write(&mut out).map_err(sink)?,
        _ => {
            return Err(mode
                .unsupported()
                .with_format(litchi_core::FileFormat::Pptx));
        },
    }
    Ok(())
}

/// The `mode` view of the presentation as OfficeCLI's JSON.
///
/// # Errors
///
/// As [`view_to`].
pub(crate) fn view_json(file: &Pptx, mode: Mode, options: &Options) -> Result<Json> {
    file.finish(|| view_json_unfenced(file, mode, options))
}

fn view_json_unfenced(file: &Pptx, mode: Mode, options: &Options) -> Result<Json> {
    check_options(options)?;
    Ok(match mode {
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
        _ => {
            return Err(mode
                .unsupported()
                .with_format(litchi_core::FileFormat::Pptx));
        },
    })
}

/// `columns` and `range` select workbook cells (OfficeCLI
/// `ViewRangeGuard.RejectTextRange`).
fn check_options(options: &Options) -> Result<()> {
    if options.columns.is_some() || options.range.is_some() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the columns and range view options apply to workbooks only",
        )
        .with_format(litchi_core::FileFormat::Pptx)
        .with_suggestion("select slides with start, end and max_lines"));
    }
    Ok(())
}

fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the view's output sink refused to write")
}

/// The slides a text or annotated view shows: 1-based numbers from
/// `start` through `end`, at most `max_lines` of them. Calls `visit` for
/// each with its scene; writes OfficeCLI's truncation line when
/// `max_lines` cuts the list.
fn each_slide<W: fmt::Write>(
    file: &Pptx,
    options: &Options,
    out: &mut W,
    mut visit: impl FnMut(u32, &SharedScene, &mut W) -> Result<()>,
) -> Result<()> {
    let total = file.slides_in_catalog()?;
    let start = options.start.unwrap_or(1);
    for position in 0..total {
        let number = u32::try_from(position + 1).unwrap_or(u32::MAX);
        if number < start {
            continue;
        }
        if options.end.is_some_and(|end| number > end) {
            break;
        }
        if let Some(max) = options.max_lines {
            if number - start >= max {
                writeln!(
                    out,
                    "... (showed {max} of {total} slides, use --start/--end to see more)"
                )
                .map_err(sink)?;
                break;
            }
        }
        let scene = scene(file, position)?;
        visit(number, &scene, out)?;
    }
    Ok(())
}

fn scene(file: &Pptx, position: usize) -> Result<SharedScene> {
    file.slide_scene(position)?.ok_or_else(|| {
        Error::new(
            Code::InternalError,
            format!("slide {} vanished from the catalog", position + 1),
        )
    })
}

/// The shape text OfficeCLI reads (`GetShapeText`): `""` without text.
fn shape_text(handle: &ShapeHandle) -> &str {
    handle.view().text().unwrap_or("")
}

fn is_blank(text: &str) -> bool {
    text.trim().is_empty()
}

// ---------------------------------------------------------------------------
// text
// ---------------------------------------------------------------------------

fn text_view(file: &Pptx, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    each_slide(file, options, out, |number, scene, out| {
        writeln!(out, "=== /slide[{number}] ===").map_err(sink)?;
        for handle in scene.iter() {
            if matches!(handle.view(), Shape::Auto(_)) {
                let text = shape_text(&handle);
                if !is_blank(text) {
                    writeln!(out, "{text}").map_err(sink)?;
                }
            }
        }
        for handle in scene.iter() {
            if matches!(handle.view(), Shape::Table(_)) {
                for cell in text::table_cells(&handle)? {
                    if !is_blank(&cell) {
                        writeln!(out, "{cell}").map_err(sink)?;
                    }
                }
            }
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

fn text_json(file: &Pptx, options: &Options) -> Result<TextJson> {
    let total = file.slides_in_catalog()?;
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
        let scene = scene(file, position)?;
        let texts = scene
            .iter()
            .filter(|handle| matches!(handle.view(), Shape::Auto(_)))
            .map(|handle| shape_text(&handle).to_owned())
            .filter(|text| !is_blank(text))
            .collect();
        slides.push(SlideTexts {
            index: number,
            path: format!("/slide[{number}]"),
            texts,
        });
    }
    Ok(TextJson {
        total_slides: total,
        slides,
    })
}

// ---------------------------------------------------------------------------
// annotated
// ---------------------------------------------------------------------------

fn annotated(file: &Pptx, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    each_slide(file, options, out, |number, scene, out| {
        writeln!(out, "[/slide[{number}]]").map_err(sink)?;
        annotate(scene.roots(), 1, out)?;
        writeln!(out).map_err(sink)
    })
}

fn annotate<W: fmt::Write>(handles: Handles, indent: usize, out: &mut W) -> Result<()> {
    let pad = "  ".repeat(indent);
    for handle in handles {
        match handle.view() {
            Shape::Auto(_) => {
                let text = shape_text(&handle);
                if is_blank(text) {
                    continue;
                }
                let kind = if is_title(&handle) {
                    "Title"
                } else {
                    "Text Box"
                };
                let body = handle.view().text_body()?;
                let first = body.as_ref().and_then(|body| {
                    body.paragraphs
                        .iter()
                        .flat_map(|paragraph| paragraph.runs.iter())
                        .find(|run| run.is_text())
                });
                let font = first
                    .and_then(|run| {
                        run.properties
                            .latin
                            .as_ref()
                            .or(run.properties.east_asian.as_ref())
                    })
                    .map_or("(default)", |font| font.typeface.as_str());
                let size = first
                    .and_then(|run| run.properties.size)
                    .map(|size| format!("{}pt", text::decimal(f64::from(size.get()) / 100.0, 2)))
                    .unwrap_or_default();
                writeln!(out, "{pad}[{kind}] \"{text}\" \u{2190} {font} {size}").map_err(sink)?;
            },
            Shape::Table(_) => {
                let (rows, cols) = text::table_shape(&handle)?;
                let name = handle.view().name().unwrap_or("Table");
                writeln!(out, "{pad}[Table] \"{name}\" \u{2190} {rows}x{cols}").map_err(sink)?;
            },
            Shape::Picture(_) => {
                let name = handle.view().name().unwrap_or("Picture");
                let alt = handle.view().description()?;
                match alt.as_deref().filter(|alt| !alt.is_empty()) {
                    Some(alt) => writeln!(out, "{pad}[Picture] \"{name}\" \u{2190} alt=\"{alt}\""),
                    None => writeln!(
                        out,
                        "{pad}[Picture] \"{name}\" \u{2190} \u{26a0} no alt text"
                    ),
                }
                .map_err(sink)?;
            },
            Shape::Group(_) => {
                let name = handle.view().name().unwrap_or("Group");
                let items = handle.children().filter(is_content).count();
                writeln!(out, "{pad}[Group] \"{name}\" \u{2190} {items} item(s)").map_err(sink)?;
                annotate(handle.children(), indent + 1, out)?;
            },
            _ => {},
        }
    }
    Ok(())
}

/// The elements OfficeCLI counts as a group's items: shapes, groups,
/// pictures, graphic frames and connectors.
fn is_content(handle: &ShapeHandle) -> bool {
    matches!(
        handle.view(),
        Shape::Auto(_)
            | Shape::Picture(_)
            | Shape::Table(_)
            | Shape::Chart(_)
            | Shape::Diagram(_)
            | Shape::Ole(_)
            | Shape::Frame(_)
            | Shape::Group(_)
            | Shape::Connector(_)
    )
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

fn slide_outlines(file: &Pptx) -> Result<Vec<SlideOutline>> {
    let total = file.slides_in_catalog()?;
    let mut slides = Vec::with_capacity(total);
    for position in 0..total {
        let scene = scene(file, position)?;
        let mut title = None;
        let (mut text_boxes, mut pictures, mut ole_objects) = (0, 0, 0);
        for handle in scene.iter() {
            match handle.view() {
                Shape::Auto(_) => {
                    let text = shape_text(&handle);
                    if is_title(&handle) {
                        if title.is_none() && !is_blank(text) {
                            title = Some(text.to_owned());
                        }
                    } else if !is_blank(text) {
                        text_boxes += 1;
                    }
                },
                Shape::Picture(_) => pictures += 1,
                Shape::Ole(_) => ole_objects += 1,
                _ => {},
            }
        }
        slides.push(SlideOutline {
            index: u32::try_from(position + 1).unwrap_or(u32::MAX),
            title,
            text_boxes,
            pictures,
            ole_objects,
        });
    }
    Ok(slides)
}

fn outline(file: &Pptx, out: &mut impl fmt::Write) -> Result<()> {
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
            "\u{251c}\u{2500}\u{2500} Slide {}: \"{}\"{details}",
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

fn outline_json(file: &Pptx) -> Result<OutlineJson> {
    let slides = slide_outlines(file)?;
    Ok(OutlineJson {
        total_slides: slides.len(),
        slides,
    })
}

// ---------------------------------------------------------------------------
// stats
// ---------------------------------------------------------------------------

/// The counts of the stats view.
#[derive(Debug, Clone, Default)]
struct Stats {
    slides: usize,
    shapes: usize,
    text_boxes: usize,
    pictures: usize,
    charts: usize,
    ole_objects: usize,
    words: usize,
    slides_without_title: usize,
    pictures_without_alt: usize,
    /// Fonts in order of first use, with their run counts.
    fonts: Vec<(String, usize)>,
}

fn words(text: &str) -> usize {
    text.split_whitespace().count()
}

fn stats(file: &Pptx) -> Result<Stats> {
    let mut stats = Stats {
        slides: file.slides_in_catalog()?,
        ..Stats::default()
    };
    for position in 0..stats.slides {
        let scene = scene(file, position)?;
        let mut titled = false;
        for handle in scene.iter() {
            match handle.view() {
                Shape::Auto(_) => {
                    stats.shapes += 1;
                    if is_title(&handle) {
                        titled = true;
                    } else {
                        stats.text_boxes += 1;
                    }
                    stats.words += words(shape_text(&handle));
                    if let Some(body) = handle.view().text_body()? {
                        let runs = body
                            .paragraphs
                            .iter()
                            .flat_map(|paragraph| paragraph.runs.iter())
                            .filter(|run| run.is_text());
                        for run in runs {
                            let font = run
                                .properties
                                .latin
                                .as_ref()
                                .or(run.properties.east_asian.as_ref());
                            if let Some(font) = font {
                                match stats
                                    .fonts
                                    .iter_mut()
                                    .find(|(name, _)| *name == font.typeface)
                                {
                                    Some((_, count)) => *count += 1,
                                    None => stats.fonts.push((font.typeface.clone(), 1)),
                                }
                            }
                        }
                    }
                },
                Shape::Picture(_) => {
                    stats.pictures += 1;
                    let alt = handle.view().description()?;
                    if alt.as_deref().is_none_or(str::is_empty) {
                        stats.pictures_without_alt += 1;
                    }
                },
                Shape::Table(_) => {
                    stats.shapes += 1;
                    for cell in text::table_cells(&handle)? {
                        stats.words += words(&cell);
                    }
                },
                Shape::Chart(_) => stats.charts += 1,
                Shape::Ole(_) => stats.ole_objects += 1,
                _ => {},
            }
        }
        if !titled {
            stats.slides_without_title += 1;
        }
    }
    // OfficeCLI orders by count, descending; the sort is stable.
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
        if self.charts > 0 {
            writeln!(out, "Charts: {}", self.charts)?;
        }
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
            charts: self.charts,
            ole_objects: self.ole_objects,
            words: self.words,
            slides_without_title: self.slides_without_title,
            pictures_without_alt: self.pictures_without_alt,
            font_usage: (!self.fonts.is_empty()).then_some(FontUsage(self.fonts)),
        }
    }
}

/// `{"slides","totalShapes","textBoxes","pictures","charts","oleObjects","words","slidesWithoutTitle","picturesWithoutAlt","fontUsage"}`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct StatsJson {
    slides: usize,
    total_shapes: usize,
    text_boxes: usize,
    pictures: usize,
    charts: usize,
    ole_objects: usize,
    words: usize,
    slides_without_title: usize,
    pictures_without_alt: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    font_usage: Option<FontUsage>,
}

/// Font names with their counts, as one JSON object in count order.
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

/// A sink wrapper that drops the output's trailing white space, as
/// OfficeCLI's views do (`TrimEnd`): white space is held back until text
/// follows it.
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
