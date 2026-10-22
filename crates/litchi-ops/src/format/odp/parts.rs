//! The ODP elements of record 0328: shape and text hyperlinks, line
//! breaks, table columns, media typing, formula objects and chart series
//! and axes.
//!
//! Hyperlinks, line breaks, columns and media read the bytes of their
//! indexed node through `litchi_odp::scene::parts` on demand. Formula
//! objects and charts live in package subdocuments; they are read once per
//! handle, on the first read that needs them, into [`Index`].

use litchi_odf_common::chart::{Dimension, Element, Kind as ChartKind};
use litchi_odp::scene::{FrameContent, Object, ObjectRef, Scene, Scope};

use super::Odp;
use crate::Result;

/// The formula objects and charts of one handle.
pub(crate) struct Index {
    equations: Vec<(ObjectRef, String)>,
    charts: Vec<(ObjectRef, Chart)>,
}

/// The series and axes of one chart object.
#[derive(Default)]
pub(crate) struct Chart {
    pub(crate) series: Vec<Series>,
    pub(crate) axes: Vec<Axis>,
}

/// One `chart:series`.
pub(crate) struct Series {
    pub(crate) values: Option<String>,
    pub(crate) label: Option<String>,
    pub(crate) class: Option<String>,
    pub(crate) axis: Option<String>,
    pub(crate) points: usize,
}

/// One `chart:axis`.
pub(crate) struct Axis {
    pub(crate) dimension: Option<&'static str>,
    pub(crate) name: Option<String>,
    pub(crate) title: Option<String>,
    pub(crate) categories: Option<String>,
    pub(crate) style: Option<String>,
}

impl Axis {
    /// The DOCX/PPTX role of the axis: `category` (x), `value` (y) or
    /// `series` (z).
    pub(crate) fn role(&self) -> Option<&'static str> {
        match self.dimension {
            Some("x") => Some("category"),
            Some("y") => Some("value"),
            Some("z") => Some("series"),
            _ => None,
        }
    }
}

/// The content part of the package subdocument `href` names (`./Object 1`
/// → `Object 1/content.xml`), or `None` for a link outside the package.
fn subdocument(href: &str) -> Option<String> {
    let path = href
        .strip_prefix("./")
        .unwrap_or(href)
        .trim_end_matches('/');
    if path.is_empty() || path.contains(':') || path.starts_with('/') || path.starts_with("..") {
        return None;
    }
    Some(format!("{path}/content.xml"))
}

/// Every embedded-object frame of the slides, groups included, in document
/// order with its reference.
fn object_frames(scene: &Scene) -> Vec<(ObjectRef, &Object)> {
    fn walk<'s>(
        page: usize,
        path: &mut Vec<usize>,
        objects: &'s [Object],
        out: &mut Vec<(ObjectRef, &'s Object)>,
    ) {
        for (index, object) in objects.iter().enumerate() {
            path.push(index);
            if object.content() == FrameContent::Object {
                out.push((
                    ObjectRef {
                        page,
                        scope: Scope::Page,
                        path: path.clone(),
                    },
                    object,
                ));
            }
            walk(page, path, object.children(), out);
            path.pop();
        }
    }
    let mut out = Vec::new();
    for (page, slide) in scene.pages().iter().enumerate() {
        walk(page, &mut Vec::new(), slide.objects(), &mut out);
    }
    out
}

impl Index {
    /// Read the formula objects and charts of `file`.
    pub(super) fn build(file: &Odp) -> Result<Self> {
        let scene = file.scene()?;
        let xml = file.content_xml()?;
        let mut equations = Vec::new();
        let mut charts = Vec::new();
        for (reference, object) in object_frames(scene) {
            let content = || -> Result<Option<String>> {
                let Some(path) = object.href().and_then(subdocument) else {
                    return Ok(None);
                };
                Ok(file
                    .member(&path)?
                    .and_then(|bytes| String::from_utf8(bytes).ok()))
            };
            if object.class() == Some("chart") {
                let chart = match content()? {
                    // The ODP owner retains no raw XML policy; read the chart
                    // part under the owner default, as the leaf does.
                    Some(xml) => match litchi_odf_common::chart::read(
                        &xml,
                        litchi_odf_common::XmlStreamLimits::owner_default_value(),
                    ) {
                        Ok(element) => summarize(&element),
                        Err(_) => Chart::default(),
                    },
                    None => Chart::default(),
                };
                charts.push((reference, chart));
                continue;
            }
            let formula = match litchi_odp::scene::parts::inline_math(xml, object) {
                Some(math) => litchi_odp::scene::parts::formula_text(math),
                None => content()?.and_then(|xml| litchi_odp::scene::parts::formula_text(&xml)),
            };
            if let Some(formula) = formula {
                equations.push((reference, formula));
            }
        }
        Ok(Self { equations, charts })
    }

    /// The formula text of an object, when the object is a formula.
    pub(crate) fn formula(&self, object: &ObjectRef) -> Option<&str> {
        self.equations
            .iter()
            .find(|(held, _)| held == object)
            .map(|(_, formula)| formula.as_str())
    }

    /// The series and axes of a chart object.
    pub(crate) fn chart(&self, object: &ObjectRef) -> Option<&Chart> {
        self.charts
            .iter()
            .find(|(held, _)| held == object)
            .map(|(_, chart)| chart)
    }
}

/// The series and axes of a chart's plot area.
fn summarize(element: &Element) -> Chart {
    let Some(plot) = element.plot_area() else {
        return Chart::default();
    };
    let series = plot
        .series()
        .map(|series| Series {
            values: series.values_range().map(str::to_owned),
            label: series.label_cell().map(str::to_owned),
            class: series.class_name().map(str::to_owned),
            axis: series.attached_axis().map(str::to_owned),
            points: series
                .data_points()
                .map(|point| point.repeated().unwrap_or(1) as usize)
                .sum(),
        })
        .collect();
    let axes = plot
        .axes()
        .map(|axis| Axis {
            dimension: match axis.dimension() {
                Ok(Dimension::X) => Some("x"),
                Ok(Dimension::Y) => Some("y"),
                Ok(Dimension::Z) => Some("z"),
                _ => None,
            },
            name: axis.name().map(str::to_owned),
            title: axis
                .element()
                .children_of_kind(ChartKind::Title)
                .next()
                .map(Element::all_text),
            categories: axis.categories_range().map(str::to_owned),
            style: axis.style_name().map(str::to_owned),
        })
        .collect();
    Chart { series, axes }
}

/// The DOCX/PPTX `link` value of a shape's click action: a URL,
/// `slide[N]` for a page jump, or the PresentationML jump names.
pub(crate) fn click_link(scene: &Scene, click: &litchi_odp::scene::parts::Click) -> Option<String> {
    match click.action() {
        None | Some("show") => {
            let href = click.href()?;
            if let Some(name) = href.strip_prefix('#') {
                let page = scene
                    .pages()
                    .iter()
                    .position(|page| page.name() == Some(name));
                return Some(match page {
                    Some(page) => format!("slide[{}]", page + 1),
                    None => href.to_owned(),
                });
            }
            Some(href.to_owned())
        },
        Some("first-page") => Some("firstslide".to_owned()),
        Some("last-page") => Some("lastslide".to_owned()),
        Some("next-page") => Some("nextslide".to_owned()),
        Some("previous-page") => Some("previousslide".to_owned()),
        Some("stop") => Some("endshow".to_owned()),
        Some(other) => Some(other.to_owned()),
    }
}

/// A unified `link` value as an ODP click action.
pub(crate) fn shape_action(
    scene: &Scene,
    link: &str,
) -> std::result::Result<litchi_odp::scene::ShapeAction, &'static str> {
    use litchi_odp::scene::ShapeAction;
    let lower = link.to_ascii_lowercase();
    Ok(match lower.as_str() {
        "firstslide" => ShapeAction::FirstPage,
        "lastslide" => ShapeAction::LastPage,
        "nextslide" => ShapeAction::NextPage,
        "previousslide" | "prevslide" => ShapeAction::PreviousPage,
        "endshow" => ShapeAction::Stop,
        _ => {
            if let Some(number) = lower
                .strip_prefix("slide[")
                .and_then(|rest| rest.strip_suffix(']'))
            {
                let index = number
                    .parse::<usize>()
                    .ok()
                    .and_then(|number| number.checked_sub(1))
                    .ok_or("a slide link names a slide by its 1-based number")?;
                let page = scene
                    .pages()
                    .get(index)
                    .ok_or("the slide link names a slide past the end")?;
                let name = page
                    .name()
                    .ok_or("the target slide has no draw:name to link to")?;
                ShapeAction::Page(name.to_owned())
            } else {
                check_url(link)?;
                ShapeAction::Url(link.to_owned())
            }
        },
    })
}

/// Refuse script, data, file and archive links (OfficeCLI's allow-list,
/// record 0321).
pub(crate) fn check_url(link: &str) -> std::result::Result<(), &'static str> {
    let lower = link.trim().to_ascii_lowercase();
    if ["javascript:", "vbscript:", "data:", "file:", "jar:"]
        .iter()
        .any(|scheme| lower.starts_with(scheme))
    {
        return Err("script, data, file and archive links are refused");
    }
    if link.trim().is_empty() {
        return Err("a hyperlink needs a non-empty target");
    }
    Ok(())
}

/// The semantic kind an object takes among slide objects: `Video` or
/// `Audio` for a plugin frame of that media, `Equation` for a formula
/// object; `None` for every other object.
pub(crate) fn semantic_kind(
    file: &Odp,
    reference: &ObjectRef,
    object: &Object,
) -> Result<Option<super::spec::Kind>> {
    use super::spec::Kind;
    Ok(match object.content() {
        FrameContent::Plugin => {
            let xml = file.content_xml()?;
            match litchi_odp::scene::parts::plugin(xml, object)
                .and_then(|plugin| plugin.media_kind())
            {
                Some(litchi_odp::scene::parts::MediaKind::Video) => Some(Kind::Video),
                Some(litchi_odp::scene::parts::MediaKind::Audio) => Some(Kind::Audio),
                _ => None,
            }
        },
        FrameContent::Object if object.class() != Some("chart") => {
            file.parts()?.formula(reference).map(|_| Kind::Equation)
        },
        _ => None,
    })
}

/// Whether a shape has a hyperlink: a click action or a link run.
pub(crate) fn shape_has_link(file: &Odp, scene: &Scene, object: &ObjectRef) -> Result<bool> {
    let Some(shape) = scene.object(object) else {
        return Ok(false);
    };
    let xml = file.content_xml()?;
    Ok(litchi_odp::scene::parts::shape_click(xml, shape).is_some()
        || shape
            .story()
            .and_then(super::node::first_link_run)
            .is_some())
}

/// The number of line breaks of a paragraph.
pub(crate) fn line_break_count(
    file: &Odp,
    scene: &Scene,
    story: &litchi_odp::scene::StoryRef,
    paragraph: usize,
) -> Result<usize> {
    let xml = file.content_xml()?;
    Ok(scene
        .story(story)
        .and_then(|story| story.paragraphs().get(paragraph))
        .map_or(0, |paragraph| {
            litchi_odp::scene::parts::line_breaks(xml, paragraph).len()
        }))
}

/// The number of declared columns of a table frame.
pub(crate) fn column_count(file: &Odp, scene: &Scene, object: &ObjectRef) -> Result<usize> {
    let xml = file.content_xml()?;
    Ok(scene
        .object(object)
        .and_then(Object::table)
        .map_or(0, |table| {
            litchi_odp::scene::parts::table_columns(xml, table)
                .iter()
                .map(|column| column.repeat() as usize)
                .sum()
        }))
}
