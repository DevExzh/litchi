//! Placeholders of a slide inserted from a layout (record 0252; ported from
//! the M2 L-PPTX-W4 reference work and adapted to the immutable editors).
//!
//! A new slide gets the layout's placeholders as empty placeholder shapes, as
//! `PowerPoint` creates them: one `p:sp` per placeholder holding a `p:ph`
//! with the layout placeholder's type, orientation, size and index, an empty
//! `p:spPr` (position and formatting inherit from the layout), and an empty
//! text body for the types that take text. Date, footer, slide-number,
//! header and slide-image placeholders are not instantiated. A layout that
//! repeats one placeholder identity gives every repeat after the first its
//! layout bounds, which inheritance cannot supply.

use std::collections::HashSet;
use std::fmt::Write as _;

use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::master_layout::PlaceholderKind;
use crate::{Error, Result};

/// The most placeholders one new slide instantiates.
const MAX_PLACEHOLDERS: usize = 256;

fn unsupported(reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "insert_slide",
        reason,
    )
}

/// One placeholder a new slide instantiates, with the `p:ph` attributes the
/// layout wrote (normalized tokens; `None` where it wrote none).
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Planned {
    kind: PlaceholderKind,
    kind_written: bool,
    orientation: Option<&'static str>,
    size: Option<&'static str>,
    index: Option<u32>,
    bounds: Option<(i64, i64, i64, i64)>,
}

impl Planned {
    const fn effective_index(&self) -> u32 {
        match self.index {
            Some(index) => index,
            None => 0,
        }
    }
}

/// The placeholders a new slide based on the layout XML instantiates.
pub(crate) fn plan(layout_xml: &[u8], shape_limits: crate::shape::Limits) -> Result<Vec<Planned>> {
    let scene = crate::shape::Scene::read(layout_xml, shape_limits)?;
    let mut plan = Vec::new();
    let mut identities: HashSet<(u32, Option<PlaceholderKind>)> = HashSet::new();
    for shape in scene.roots() {
        let Some(placeholder) = shape.placeholder() else {
            continue;
        };
        let attributes = attributes(&shape.self_contained_xml()?)?
            .ok_or_else(|| unsupported("a layout placeholder's p:ph cannot be located"))?;
        if attributes.effective_index() != placeholder.index()
            || attributes.kind_written != placeholder.kind().is_some()
        {
            return Err(unsupported(
                "a layout placeholder's identity reads ambiguously",
            ));
        }
        if matches!(
            attributes.kind,
            PlaceholderKind::DateTime
                | PlaceholderKind::Footer
                | PlaceholderKind::SlideNumber
                | PlaceholderKind::Header
                | PlaceholderKind::SlideImage
        ) {
            continue;
        }
        let identity = if attributes.effective_index() == 0 {
            (0, Some(inheritance_kind(attributes.kind)))
        } else {
            (attributes.effective_index(), None)
        };
        let repeated = !identities.insert(identity);
        let bounds = if repeated {
            shape
                .common()
                .bounds()
                .map(|bounds| (bounds.x(), bounds.y(), bounds.width(), bounds.height()))
        } else {
            None
        };
        if plan.len() >= MAX_PLACEHOLDERS {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Placeholders),
                observed: plan.len().checked_add(1).unwrap_or(usize::MAX),
                maximum: MAX_PLACEHOLDERS,
            });
        }
        plan.push(Planned {
            bounds,
            ..attributes
        });
    }
    Ok(plan)
}

/// Title placeholders inherit from either title type.
const fn inheritance_kind(kind: PlaceholderKind) -> PlaceholderKind {
    match kind {
        PlaceholderKind::CenteredTitle => PlaceholderKind::Title,
        other => other,
    }
}

/// The attributes of a shape fragment's `p:nv*Pr/p:nvPr/p:ph`.
fn attributes(fragment: &[u8]) -> Result<Option<Planned>> {
    let mut reader = NsReader::from_reader(fragment);
    let mut open: Vec<Vec<u8>> = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                open.pop();
                continue;
            },
            Event::Eof => return Ok(None),
            Event::DocType(_) => return Err(Error::Invalid("DTDs are rejected".into())),
            _ => continue,
        };
        let presentation = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == crate::namespace::PRESENTATIONML_NAMESPACE
                    || value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE
        );
        let local = element.local_name().as_ref().to_vec();
        let on_path = open.len() == 3
            && open[1].starts_with(b"nv")
            && open[1].ends_with(b"Pr")
            && open[2] == b"nvPr";
        if on_path && presentation && local == b"ph" {
            let decoder = reader.decoder();
            let value = |name: &[u8]| {
                litchi_ooxml_common::xml::unqualified_attribute_value(&element, name, decoder)
            };
            let kind_token = value(b"type")?;
            let kind = match kind_token.as_deref() {
                None => PlaceholderKind::Object,
                Some(token) => PlaceholderKind::from_token(token)
                    .ok_or_else(|| unsupported("unknown layout placeholder type"))?,
            };
            let orientation = match value(b"orient")?.as_deref().map(str::trim) {
                None => None,
                Some("horz") => Some("horz"),
                Some("vert") => Some("vert"),
                Some(_) => return Err(unsupported("unknown layout placeholder orientation")),
            };
            let size = match value(b"sz")?.as_deref().map(str::trim) {
                None => None,
                Some("full") => Some("full"),
                Some("half") => Some("half"),
                Some("quarter") => Some("quarter"),
                Some(_) => return Err(unsupported("unknown layout placeholder size")),
            };
            let index = match value(b"idx")? {
                None => None,
                Some(text) => Some(
                    text.trim()
                        .parse::<u32>()
                        .map_err(|_error| unsupported("invalid layout placeholder index"))?,
                ),
            };
            return Ok(Some(Planned {
                kind,
                kind_written: kind_token.is_some(),
                orientation,
                size,
                index,
                bounds: None,
            }));
        }
        if !empty {
            open.push(local);
        }
    }
}

/// `PowerPoint`'s name for a new placeholder.
fn name(planned: &Planned, shape_id: u32) -> String {
    placeholder_name(planned.kind, planned.orientation == Some("vert"), shape_id)
}

/// `PowerPoint`'s name for a new placeholder of `kind` with non-visual
/// identifier `shape_id` (record 0346 shares it with placeholder addition).
pub(crate) fn placeholder_name(kind: PlaceholderKind, vertical: bool, shape_id: u32) -> String {
    let label = match kind {
        PlaceholderKind::Title | PlaceholderKind::CenteredTitle => "Title",
        PlaceholderKind::Subtitle => "Subtitle",
        PlaceholderKind::Body => "Text Placeholder",
        PlaceholderKind::Object => "Content Placeholder",
        PlaceholderKind::Chart => "Chart Placeholder",
        PlaceholderKind::Table => "Table Placeholder",
        PlaceholderKind::ClipArt => "ClipArt Placeholder",
        PlaceholderKind::Diagram => "SmartArt Placeholder",
        PlaceholderKind::Media => "Media Placeholder",
        PlaceholderKind::Picture => "Picture Placeholder",
        PlaceholderKind::SlideImage => "Slide Image Placeholder",
        PlaceholderKind::DateTime => "Date Placeholder",
        PlaceholderKind::Footer => "Footer Placeholder",
        PlaceholderKind::SlideNumber => "Slide Number Placeholder",
        PlaceholderKind::Header => "Header Placeholder",
    };
    let vertical = if vertical { "Vertical " } else { "" };
    format!("{vertical}{label} {}", shape_id.saturating_sub(1))
}

/// Whether a new placeholder of `kind` holds an empty text body.
pub(crate) const fn takes_text(kind: PlaceholderKind) -> bool {
    matches!(
        kind,
        PlaceholderKind::Title
            | PlaceholderKind::CenteredTitle
            | PlaceholderKind::Subtitle
            | PlaceholderKind::Body
            | PlaceholderKind::Object
    )
}

/// The transitional XML of a new slide, checked against the selected shape
/// output bound before the caller publishes it.
pub(crate) fn slide_xml_with_limits(
    slide_id: u32,
    plan: &[Planned],
    shape_limits: crate::shape::Limits,
) -> Result<String> {
    let mut xml = crate::builder::Slide::new(slide_id).generate_slide_xml()?;
    if plan.is_empty() {
        if xml.len() > shape_limits.output_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: xml.len(),
                maximum: shape_limits.output_bytes(),
            });
        }
        return Ok(xml);
    }
    let anchor = xml
        .rfind("</p:spTree>")
        .ok_or_else(|| Error::Invalid("generated slide lacks its shape tree".into()))?;
    let mut shapes = String::new();
    let format_error = |error: std::fmt::Error| Error::Xml(error.to_string());
    for (offset, planned) in plan.iter().enumerate() {
        let shape_id = u32::try_from(offset)
            .ok()
            .and_then(|offset| offset.checked_add(2))
            .ok_or_else(|| unsupported("the layout has too many placeholders"))?;
        write!(
            shapes,
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph"#,
            name(planned, shape_id)
        )
        .map_err(format_error)?;
        if planned.kind_written {
            write!(shapes, r#" type="{}""#, planned.kind.as_str()).map_err(format_error)?;
        }
        if let Some(orientation) = planned.orientation {
            write!(shapes, r#" orient="{orientation}""#).map_err(format_error)?;
        }
        if let Some(size) = planned.size {
            write!(shapes, r#" sz="{size}""#).map_err(format_error)?;
        }
        if let Some(index) = planned.index {
            write!(shapes, r#" idx="{index}""#).map_err(format_error)?;
        }
        shapes.push_str("/></p:nvPr></p:nvSpPr>");
        match planned.bounds {
            Some((x, y, width, height)) => write!(
                shapes,
                r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm></p:spPr>"#
            )
            .map_err(format_error)?,
            None => shapes.push_str("<p:spPr/>"),
        }
        if takes_text(planned.kind) {
            shapes.push_str("<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>");
        }
        shapes.push_str("</p:sp>");
    }
    xml.insert_str(anchor, &shapes);
    if xml.len() > shape_limits.output_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: xml.len(),
            maximum: shape_limits.output_bytes(),
        });
    }
    Ok(xml)
}

/// Check a staged new slide under the selected shape policy.
pub(crate) fn read_back_with_limits(
    slide_xml: &[u8],
    plan: &[Planned],
    shape_limits: crate::shape::Limits,
) -> Result<()> {
    let scene = crate::shape::Scene::read(slide_xml, shape_limits)?;
    let read: Vec<_> = scene
        .iter()
        .map(|shape| {
            shape.placeholder().map(|placeholder| {
                (
                    placeholder
                        .kind()
                        .map_or(Some(PlaceholderKind::Object), PlaceholderKind::from_token),
                    placeholder.index(),
                )
            })
        })
        .collect();
    let expected: Vec<_> = plan
        .iter()
        .map(|planned| Some((Some(planned.kind), planned.effective_index())))
        .collect();
    if read != expected {
        return Err(Error::Invalid(
            "new slide placeholders did not read back".into(),
        ));
    }
    Ok(())
}
