//! OpenDocument presentation semantic views (record 0010 §6).
//!
//! The views walk the scene index the navigator reads, in the four semantic
//! modes shared by the presentation adapters. Character formatting lives
//! in automatic styles the adapter does not resolve, so the annotated view
//! names a text frame's style instead of a font, and the stats view has no
//! font usage (a registered view gap).

use std::fmt;

use litchi_odp::scene::{Object, Page};
use serde::Serialize;

use super::Odp;
use super::node::object_kind;
use super::spec::Kind;
use crate::error::{Code, Error, Result};
use crate::view::mode::{Mode, Options};

const MAX_VIEW_DEPTH: usize = 256;

/// Writes the `mode` view into `out`.
///
/// # Errors
///
/// Returns `invalid_argument` for workbook-only options, the leaf's typed
/// error when the scene cannot be built, `unsupported_capability` for a
/// rendering mode, and `io_error` when the sink refuses a write.
pub(crate) fn view_to<W: fmt::Write>(
    file: &Odp,
    mode: Mode,
    options: &Options,
    out: &mut W,
) -> Result<()> {
    file.check_source()?;
    check_options(options)?;
    match mode {
        Mode::Text => text_view(file, options, out)?,
        Mode::Annotated => annotated(file, options, out)?,
        Mode::Outline => outline(file, out)?,
        Mode::Stats => stats(file)?.write(out).map_err(sink)?,
        _ => return Err(unsupported(mode)),
    }
    file.check_source()
}

/// The `mode` view as its OfficeCLI-shaped JSON projection.
///
/// # Errors
///
/// As [`view_to`].
pub(crate) fn view_json(file: &Odp, mode: Mode, options: &Options) -> Result<Json> {
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
    mode.unsupported().with_format(litchi_core::FileFormat::Odp)
}

fn check_options(options: &Options) -> Result<()> {
    if options.columns.is_some() || options.range.is_some() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the columns and range view options apply to workbooks only",
        )
        .with_format(litchi_core::FileFormat::Odp)
        .with_suggestion("select slides with start, end and max_lines"));
    }
    Ok(())
}

fn sink(_: fmt::Error) -> Error {
    Error::new(Code::IoError, "the ODP view output sink refused to write")
        .with_format(litchi_core::FileFormat::Odp)
}

fn depth_limit() -> Error {
    Error::new(Code::LimitExceeded, "ODP view group nesting exceeded")
        .with_format(litchi_core::FileFormat::Odp)
}

/// The selected slides, as `(number, page)`, with the truncation note the
/// text views print.
fn selected<'s>(
    pages: &'s [Page],
    options: &Options,
) -> (Vec<(u32, &'s Page)>, Option<(u32, usize)>) {
    let start = options.start.unwrap_or(1);
    let mut chosen = Vec::new();
    for (position, page) in pages.iter().enumerate() {
        let number = u32::try_from(position + 1).unwrap_or(u32::MAX);
        if number < start {
            continue;
        }
        if options.end.is_some_and(|end| number > end) {
            break;
        }
        if let Some(max) = options.max_lines
            && number.saturating_sub(start) >= max
        {
            return (chosen, Some((max, pages.len())));
        }
        chosen.push((number, page));
    }
    (chosen, None)
}

/// The text-bearing objects in document order, groups flattened.
fn collect_texts(object: &Object, depth: usize, texts: &mut Vec<String>) -> Result<()> {
    if depth > MAX_VIEW_DEPTH {
        return Err(depth_limit());
    }
    match object_kind(object) {
        Kind::Group => {
            for child in object.children() {
                collect_texts(child, depth + 1, texts)?;
            }
        },
        Kind::Table => {
            if let Some(table) = object.table() {
                for row in table.rows() {
                    for cell in row.cells() {
                        let text = cell.text();
                        if !text.trim().is_empty() {
                            texts.push(text);
                        }
                    }
                }
            }
        },
        _ => {
            let text = object.text();
            if !text.trim().is_empty() {
                texts.push(text);
            }
        },
    }
    Ok(())
}

fn text_view(file: &Odp, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let scene = file.scene()?;
    let (pages, truncated) = selected(scene.pages(), options);
    let mut first = true;
    for (number, page) in pages {
        if !first {
            writeln!(out).map_err(sink)?;
        }
        first = false;
        write!(out, "=== /slide[{number}] ===").map_err(sink)?;
        let mut texts = Vec::new();
        for object in page.objects() {
            collect_texts(object, 0, &mut texts)?;
        }
        for text in texts {
            write!(out, "\n{text}").map_err(sink)?;
        }
        writeln!(out).map_err(sink)?;
    }
    if let Some((shown, total)) = truncated {
        writeln!(
            out,
            "... (showed {shown} of {total} slides, use --start/--end to see more)"
        )
        .map_err(sink)?;
    }
    Ok(())
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

fn text_json(file: &Odp, options: &Options) -> Result<TextJson> {
    let scene = file.scene()?;
    let (pages, _) = selected(scene.pages(), options);
    let mut slides = Vec::with_capacity(pages.len());
    for (number, page) in pages {
        let mut texts = Vec::new();
        for object in page.objects() {
            collect_texts(object, 0, &mut texts)?;
        }
        slides.push(SlideTexts {
            index: number,
            path: format!("/slide[{number}]"),
            texts,
        });
    }
    Ok(TextJson {
        total_slides: scene.pages().len(),
        slides,
    })
}

fn annotated(file: &Odp, options: &Options, out: &mut impl fmt::Write) -> Result<()> {
    let scene = file.scene()?;
    let (pages, truncated) = selected(scene.pages(), options);
    let mut first = true;
    for (number, page) in pages {
        if !first {
            writeln!(out).map_err(sink)?;
        }
        first = false;
        writeln!(out, "[/slide[{number}]]").map_err(sink)?;
        for object in page.objects() {
            annotate(object, 1, 0, out)?;
        }
    }
    if let Some((shown, total)) = truncated {
        writeln!(
            out,
            "... (showed {shown} of {total} slides, use --start/--end to see more)"
        )
        .map_err(sink)?;
    }
    Ok(())
}

fn annotate(object: &Object, indent: usize, depth: usize, out: &mut impl fmt::Write) -> Result<()> {
    if depth > MAX_VIEW_DEPTH {
        return Err(depth_limit());
    }
    let pad = "  ".repeat(indent);
    let name = |fallback: &'static str| object.name().unwrap_or(fallback).to_owned();
    match object_kind(object) {
        Kind::Group => {
            writeln!(
                out,
                "{pad}[Group] \"{}\" ← {} item(s)",
                name("Group"),
                object.children().len()
            )
            .map_err(sink)?;
            for child in object.children() {
                annotate(child, indent + 1, depth + 1, out)?;
            }
        },
        Kind::Picture => {
            let alt = object
                .title()
                .or(object.description())
                .filter(|alt| !alt.is_empty());
            match alt {
                Some(alt) => writeln!(
                    out,
                    "{pad}[Picture] \"{}\" ← alt=\"{alt}\"",
                    name("Picture")
                ),
                None => writeln!(
                    out,
                    "{pad}[Picture] \"{}\" ← ⚠ no alt text",
                    name("Picture")
                ),
            }
            .map_err(sink)?;
        },
        Kind::Table => {
            let (rows, columns) = object
                .table()
                .map_or((0, 0), |table| (table.rows().len(), table.columns()));
            writeln!(out, "{pad}[Table] \"{}\" ← {rows}x{columns}", name("Table")).map_err(sink)?;
        },
        Kind::Chart => writeln!(out, "{pad}[Chart] \"{}\"", name("Chart")).map_err(sink)?,
        Kind::Ole => writeln!(out, "{pad}[OLE] \"{}\"", name("OLE")).map_err(sink)?,
        Kind::Media => writeln!(out, "{pad}[Media] \"{}\"", name("Media")).map_err(sink)?,
        Kind::Connector => {},
        _ => {
            let text = object.text();
            if text.trim().is_empty() {
                return Ok(());
            }
            let kind = if object.class() == Some("title") {
                "Title"
            } else {
                "Text Box"
            };
            let style = object.style_name().map_or_else(
                || "(default style)".to_owned(),
                |style| format!("style={style}"),
            );
            writeln!(out, "{pad}[{kind}] \"{text}\" ← {style}").map_err(sink)?;
        },
    }
    Ok(())
}

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

fn count_objects(
    object: &Object,
    depth: usize,
    title: &mut Option<String>,
    counts: &mut [usize; 3],
) -> Result<()> {
    if depth > MAX_VIEW_DEPTH {
        return Err(depth_limit());
    }
    match object_kind(object) {
        Kind::Group => {
            for child in object.children() {
                count_objects(child, depth + 1, title, counts)?;
            }
        },
        Kind::Picture | Kind::Media => counts[1] += 1,
        Kind::Ole | Kind::Chart => counts[2] += 1,
        Kind::Shape => {
            let text = object.text();
            if object.class() == Some("title") {
                if title.is_none() && !text.trim().is_empty() {
                    *title = Some(text);
                }
            } else if !text.trim().is_empty() {
                counts[0] += 1;
            }
        },
        _ => {},
    }
    Ok(())
}

fn slide_outlines(file: &Odp) -> Result<Vec<SlideOutline>> {
    let scene = file.scene()?;
    let mut slides = Vec::with_capacity(scene.pages().len());
    for (position, page) in scene.pages().iter().enumerate() {
        let mut title = None;
        let mut counts = [0; 3];
        for object in page.objects() {
            count_objects(object, 0, &mut title, &mut counts)?;
        }
        slides.push(SlideOutline {
            index: u32::try_from(position + 1).unwrap_or(u32::MAX),
            title,
            text_boxes: counts[0],
            pictures: counts[1],
            ole_objects: counts[2],
        });
    }
    Ok(slides)
}

fn outline(file: &Odp, out: &mut impl fmt::Write) -> Result<()> {
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

fn outline_json(file: &Odp) -> Result<OutlineJson> {
    let slides = slide_outlines(file)?;
    Ok(OutlineJson {
        total_slides: slides.len(),
        slides,
    })
}

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
    comments: usize,
}

fn walk(object: &Object, depth: usize, stats: &mut Stats, titled: &mut bool) -> Result<()> {
    if depth > MAX_VIEW_DEPTH {
        return Err(depth_limit());
    }
    stats.shapes += 1;
    match object_kind(object) {
        Kind::Group => {
            for child in object.children() {
                walk(child, depth + 1, stats, titled)?;
            }
        },
        Kind::Table => {
            if let Some(table) = object.table() {
                for row in table.rows() {
                    for cell in row.cells() {
                        stats.words += cell.text().split_whitespace().count();
                    }
                }
            }
        },
        Kind::Picture | Kind::Media => {
            stats.pictures += 1;
            if object
                .title()
                .or(object.description())
                .is_none_or(str::is_empty)
            {
                stats.pictures_without_alt += 1;
            }
        },
        Kind::Ole | Kind::Chart => stats.ole_objects += 1,
        Kind::Shape => {
            if object.story().is_some() {
                stats.text_boxes += 1;
            }
            stats.words += object.text().split_whitespace().count();
            if object.class() == Some("title") {
                *titled = true;
            }
        },
        _ => {},
    }
    Ok(())
}

fn stats(file: &Odp) -> Result<Stats> {
    let scene = file.scene()?;
    let mut stats = Stats {
        slides: scene.pages().len(),
        comments: file.annotations()?.len(),
        ..Stats::default()
    };
    for page in scene.pages() {
        let mut titled = false;
        for object in page.objects() {
            walk(object, 0, &mut stats, &mut titled)?;
        }
        if !titled {
            stats.slides_without_title += 1;
        }
    }
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
        if self.comments > 0 {
            writeln!(out, "Comments: {}", self.comments)?;
        }
        writeln!(out, "Slides without title: {}", self.slides_without_title)?;
        write!(
            out,
            "Pictures without alt text: {}",
            self.pictures_without_alt
        )
    }

    fn json(self) -> StatsJson {
        StatsJson {
            slides: self.slides,
            total_shapes: self.shapes,
            text_boxes: self.text_boxes,
            pictures: self.pictures,
            ole_objects: self.ole_objects,
            words: self.words,
            comments: self.comments,
            slides_without_title: self.slides_without_title,
            pictures_without_alt: self.pictures_without_alt,
        }
    }
}

/// `{"slides","totalShapes","textBoxes","pictures","oleObjects","words","comments","slidesWithoutTitle","picturesWithoutAlt"}`.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct StatsJson {
    slides: usize,
    total_shapes: usize,
    text_boxes: usize,
    pictures: usize,
    ole_objects: usize,
    words: usize,
    comments: usize,
    slides_without_title: usize,
    pictures_without_alt: usize,
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
