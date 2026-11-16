//! Office Math equation shapes: add, replace and remove.
//!
//! PowerPoint 2010 and later write a shape whose text holds an equation as
//! one `mc:AlternateContent` block ([MS-ODRAWXML] §2.3.3.41 for `a14:m`,
//! §3.5 for the example; ECMA-376 Part 3 for markup compatibility): the
//! `mc:Choice Requires="a14"` branch holds the shape with the equation's
//! OMML in `p:txBody/a:p/a14:m`, and the `mc:Fallback` branch the same shape
//! with locked text, filled with a picture of the equation, for readers
//! without Office Math. [`Edit::add_equation`] writes that form for a new
//! text box, [`Edit::set_equation`] replaces the OMML and the picture of an
//! existing one, and [`Edit::remove_equation`] removes the whole block.
//! Every other byte of the slide stays as it was.
//!
//! Only a shape whose text is exactly one equation is rewritten: the
//! fallback picture shows the whole text body, which the picture renderer
//! cannot draw when it mixes equations with other text.

use std::ops::Range;

use litchi_equation::{Justification, Limits, OmmlHost, PngOptions};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::{BlobPart, PackURI};
use quick_xml::Reader;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::picture::attribute_value;
use super::{
    Edit, conformance_image_relationship, conformance_namespaces, ensure_shape_id, free_part_name,
    invalid, max_shape_reference_id, remove_unreferenced_dependencies, shape_relationship_ids,
};
use crate::shape::math::sites::{self, Block, DEFAULT_INSETS, Site};
use crate::shape::math::{EquationSource, EquationWrite};
use crate::{Error, Result};

const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const MATH: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const TRANSITIONAL_DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// The relationship namespaces (transitional and strict).
const RELATIONSHIPS: [&[u8]; 2] = [
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships",
];

/// The text size of a new text box when the caller names none, in points:
/// PowerPoint's default for text boxes.
const DEFAULT_FONT_SIZE_PT: f64 = 18.0;
/// Pixels per CSS pixel of the fallback picture (384 pixels per inch).
const PICTURE_SCALE: f64 = 4.0;
/// The most pixels of a fallback picture; a larger frame is drawn at a
/// lower resolution.
const MAX_FALLBACK_PIXELS: f64 = 4_194_304.0;
/// EMU per point (ECMA-376 Part 1, §20.1.10.16).
const EMU_PER_POINT: f64 = 12_700.0;
/// The font PowerPoint names on every math run.
const DEFAULT_RUN_FONTS: &str = r#"<a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>"#;
/// The picture-fill locks of a fallback shape, as PowerPoint writes them.
const FALLBACK_LOCKS: &str = r#"<a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/>"#;

fn refused(operation: &'static str, reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        operation,
        reason,
    )
}

/// The host run properties every new math run gets: the attributes and the
/// child elements of an `a:rPr` (ECMA-376 Part 1, §21.1.2.3.9
/// `CT_TextCharacterProperties`), as raw XML with the `a` prefix.
#[derive(Debug, Clone)]
struct HostRun {
    attributes: Vec<(String, String)>,
    /// Each child's local name and raw XML, in order.
    children: Vec<(String, String)>,
}

/// The fill choices of `CT_TextCharacterProperties`.
const FILLS: [&str; 6] = [
    "noFill",
    "solidFill",
    "gradFill",
    "blipFill",
    "pattFill",
    "grpFill",
];

impl HostRun {
    /// What PowerPoint writes, with `size` in hundredths of a point.
    fn native(size: Option<u32>) -> Self {
        let mut attributes = vec![
            ("lang".to_owned(), "en-US".to_owned()),
            ("i".to_owned(), "1".to_owned()),
        ];
        if let Some(size) = size {
            attributes.insert(1, ("sz".to_owned(), size.to_string()));
        }
        Self {
            attributes,
            children: vec![("latin".to_owned(), DEFAULT_RUN_FONTS.to_owned())],
        }
    }

    /// The text size, in points.
    fn font_size_pt(&self) -> Option<f64> {
        self.attributes
            .iter()
            .find(|(key, _)| key == "sz")
            .and_then(|(_, value)| value.parse::<u32>().ok())
            .filter(|size| (100..=400_000).contains(size))
            .map(|size| f64::from(size) / 100.0)
    }

    /// The first host `a:rPr` of a math run in `omml`; `None` when there is
    /// none or it, or one of its children, uses another prefix than `a`.
    fn of(omml: &str) -> Result<Option<Self>> {
        let mut reader = Reader::from_reader(omml.as_bytes());
        let mut math_prefix: Option<Vec<u8>> = None;
        let mut path: Vec<Vec<u8>> = Vec::new();
        // The template found, the path depth of its element while it is
        // open, and the start of its open child.
        let mut found: Option<(Self, usize)> = None;
        let mut child: Option<(usize, String)> = None;
        loop {
            let start = position(&reader)?;
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let end = position(&reader)?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let name = element.name();
                    let prefix = name.prefix().map(|prefix| prefix.as_ref().to_vec());
                    let local = String::from_utf8_lossy(name.local_name().as_ref()).into_owned();
                    let empty = matches!(event, Event::Empty(_));
                    if math_prefix.is_none() {
                        math_prefix = Some(prefix.clone().unwrap_or_default());
                    }
                    if let Some((run, level)) = found.as_mut() {
                        if path.len() == *level + 1 {
                            if prefix.as_deref() != Some(b"a") {
                                return Ok(None);
                            }
                            if empty {
                                run.children.push((local, omml[start..end].to_owned()));
                            } else {
                                child = Some((start, local));
                            }
                        }
                    } else if local == "rPr"
                        && prefix.as_deref() == Some(b"a")
                        && path.last().is_some_and(|parent| {
                            math_prefix
                                .as_ref()
                                .is_some_and(|math| parent.as_slice() == qualified(math, b"r"))
                        })
                    {
                        let mut attributes = Vec::new();
                        for attribute in element.checked_attributes() {
                            let attribute =
                                attribute.map_err(|error| Error::Xml(error.to_string()))?;
                            let key = std::str::from_utf8(attribute.key.as_ref())
                                .map_err(|error| Error::Xml(error.to_string()))?;
                            let value = std::str::from_utf8(&attribute.value)
                                .map_err(|error| Error::Xml(error.to_string()))?;
                            attributes.push((key.to_owned(), value.to_owned()));
                        }
                        let run = Self {
                            attributes,
                            children: Vec::new(),
                        };
                        if empty {
                            return Ok(Some(run));
                        }
                        found = Some((run, path.len()));
                    }
                    if !empty {
                        path.push(name.as_ref().to_vec());
                    }
                },
                Event::End(_) => {
                    path.pop();
                    if let Some((run, level)) = found.as_mut() {
                        if path.len() == *level + 1
                            && let Some((child_start, local)) = child.take()
                        {
                            run.children
                                .push((local, omml[child_start..end].to_owned()));
                        } else if path.len() == *level {
                            return Ok(found.map(|(run, _)| run));
                        }
                    }
                },
                Event::Eof => return Ok(None),
                _ => {},
            }
        }
    }

    /// An `a:rPr` element: these properties with `overrides` (attributes
    /// the converter wrote) on top, and `fills` (the converter's fill) in
    /// place of this template's fill. `upright` drops the italic switch.
    fn element(&self, overrides: &[(String, String)], fills: &str, upright: bool) -> String {
        let mut out = String::from("<a:rPr");
        for (key, value) in &self.attributes {
            if overrides.iter().any(|(name, _)| name == key) || (upright && key == "i") {
                continue;
            }
            push_attribute(&mut out, key, value);
        }
        for (key, value) in overrides {
            push_attribute(&mut out, key, value);
        }
        let mut children = String::new();
        let mut filled = fills.is_empty();
        for (local, raw) in &self.children {
            let is_fill = FILLS.contains(&local.as_str());
            // The fill follows only the outline (`a:ln`).
            if !filled && local != "ln" {
                children.push_str(fills);
                filled = true;
            }
            if is_fill && !fills.is_empty() {
                continue;
            }
            children.push_str(raw);
        }
        if !filled {
            children.push_str(fills);
        }
        if children.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            out.push_str(&children);
            out.push_str("</a:rPr>");
        }
        out
    }
}

fn push_attribute(out: &mut String, key: &str, value: &str) {
    out.push(' ');
    out.push_str(key);
    out.push_str("=\"");
    // Values come from parsed XML and are still escaped; only a quote of
    // the other kind can need escaping here.
    out.push_str(&value.replace('"', "&quot;"));
    out.push('"');
}

fn qualified(prefix: &[u8], local: &[u8]) -> Vec<u8> {
    if prefix.is_empty() {
        local.to_vec()
    } else {
        let mut out = prefix.to_vec();
        out.push(b':');
        out.extend_from_slice(local);
        out
    }
}

fn position<R>(reader: &Reader<R>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_error| invalid("the equation is too large"))
}

/// Give every run of converter-written OMML the host run properties
/// PowerPoint writes: `run` for runs without any, merged under the
/// converter's own `a:rPr` (colour, bold, italic) where it wrote one.
fn apply_host_runs(omml: &str, run: &HostRun) -> Result<String> {
    let mut reader = Reader::from_reader(omml.as_bytes());
    let mut out = String::with_capacity(omml.len() * 2);
    let mut copied = 0usize;
    // For each open element: its name, and for an `m:r`, whether its host
    // properties are written and whether its `m:rPr` makes it upright.
    let mut stack: Vec<(Vec<u8>, bool, bool)> = Vec::new();
    // An open converter `a:rPr`: its start, its attributes and where its
    // children start.
    let mut pending: Option<(usize, Vec<(String, String)>, usize, bool)> = None;
    loop {
        let start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = position(&reader)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let name = element.name().as_ref().to_vec();
                let empty = matches!(event, Event::Empty(_));
                if let Some((parent, written, upright)) = stack.last_mut()
                    && parent.as_slice() == b"m:r"
                    && !*written
                {
                    match name.as_slice() {
                        b"m:rPr" => {
                            let tail = &omml[start..];
                            let close = if empty {
                                end - start
                            } else {
                                tail.find("</m:rPr>").map_or(0, |at| at + "</m:rPr>".len())
                            };
                            let properties = &tail[..close];
                            *upright = properties.contains("<m:nor/>")
                                || properties.contains(r#"m:val="p""#)
                                || properties.contains(r#"m:val="b""#);
                        },
                        b"a:rPr" => {
                            *written = true;
                            let mut attributes = Vec::new();
                            for attribute in element.checked_attributes() {
                                let attribute =
                                    attribute.map_err(|error| Error::Xml(error.to_string()))?;
                                attributes.push((
                                    String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                                    String::from_utf8_lossy(&attribute.value).into_owned(),
                                ));
                            }
                            if empty {
                                out.push_str(&omml[copied..start]);
                                out.push_str(&run.element(&attributes, "", *upright));
                                copied = end;
                            } else {
                                pending = Some((start, attributes, end, *upright));
                            }
                        },
                        _ => {
                            *written = true;
                            out.push_str(&omml[copied..start]);
                            out.push_str(&run.element(&[], "", *upright));
                            copied = start;
                        },
                    }
                }
                if !empty {
                    stack.push((name, false, false));
                }
            },
            Event::End(ref element) => {
                stack.pop();
                if element.name().as_ref() == b"a:rPr"
                    && let Some((open, attributes, children, upright)) = pending.take()
                {
                    out.push_str(&omml[copied..open]);
                    out.push_str(&run.element(&attributes, &omml[children..start], upright));
                    copied = end;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out.push_str(&omml[copied..]);
    Ok(out)
}

/// The namespace prefixes the slide's root element binds, and their URIs.
fn root_bindings(xml: &[u8]) -> Result<Bindings> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(ref element) | Event::Empty(ref element) => {
                return Ok(element
                    .checked_attributes()
                    .flatten()
                    .filter_map(|attribute| {
                        attribute
                            .key
                            .as_ref()
                            .strip_prefix(b"xmlns:")
                            .map(|prefix| (prefix.to_vec(), attribute.value.into_owned()))
                    })
                    .collect());
            },
            Event::Eof => return Ok(Vec::new()),
            _ => {},
        }
    }
}

/// Namespace prefixes in scope and their URIs.
type Bindings = Vec<(Vec<u8>, Vec<u8>)>;

/// Check caller OMML and declare the `m` and `a` prefixes it leaves
/// undeclared and the slide does not bind (`bound`) to the same namespace,
/// so it can stand inside the slide.
fn admit_omml(omml: &str, dml: &str, bound: &Bindings) -> Result<String> {
    let omml = omml.trim();
    let mut reader = NsReader::from_reader(omml.as_bytes());
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut root_name_end = None;
    let mut undeclared: Vec<Vec<u8>> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the equation is too large"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(format!("malformed OMML: {error}")))?;
        let note = |undeclared: &mut Vec<Vec<u8>>, namespace: &ResolveResult<'_>| {
            if let ResolveResult::Unknown(prefix) = namespace
                && !undeclared.contains(prefix)
            {
                undeclared.push(prefix.clone());
            }
        };
        let math = matches!(&namespace, ResolveResult::Bound(Namespace(value)) if *value == MATH.as_bytes())
            || matches!(&namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"m");
        note(&mut undeclared, &namespace);
        drop(namespace);
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                for attribute in element.checked_attributes() {
                    let attribute = attribute
                        .map_err(|error| Error::Xml(format!("malformed OMML: {error}")))?;
                    let key = attribute.key;
                    if key.as_ref() == b"xmlns" || key.as_ref().starts_with(b"xmlns:") {
                        continue;
                    }
                    let (namespace, _) = reader.resolver().resolve_attribute(key);
                    let uri = match &namespace {
                        ResolveResult::Bound(Namespace(uri)) => Some(uri.to_vec()),
                        ResolveResult::Unknown(prefix) => bound
                            .iter()
                            .find(|(name, _)| name == prefix)
                            .map(|(_, uri)| uri.clone()),
                        ResolveResult::Unbound => None,
                    };
                    if uri.is_some_and(|uri| RELATIONSHIPS.contains(&uri.as_slice())) {
                        return Err(invalid(
                            "equation OMML cannot refer to relationships (r:id, r:embed and the like): the slide has none for it",
                        ));
                    }
                    note(&mut undeclared, &namespace);
                }
                if depth == 0 {
                    roots += 1;
                    let local = element.local_name();
                    if !matches!(local.as_ref(), b"oMath" | b"oMathPara") {
                        return Err(invalid(
                            "equation OMML must be one m:oMath or m:oMathPara element",
                        ));
                    }
                    if !math {
                        return Err(invalid("equation OMML must use the Office Math namespace"));
                    }
                    root_name_end = Some(start + 1 + element.name().as_ref().len());
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Text(ref text) if depth == 0 => {
                if text.iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(invalid("equation OMML has text outside its root"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::Decl(_) => {
                return Err(invalid(
                    "equation OMML cannot hold declarations or processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let at = root_name_end.ok_or_else(|| invalid("equation OMML is empty"))?;
    if depth != 0 {
        return Err(Error::Xml(
            "malformed OMML: an element is not closed".to_owned(),
        ));
    }
    if roots != 1 {
        return Err(invalid(
            "equation OMML must be one m:oMath or m:oMathPara element",
        ));
    }
    let mut declarations = String::new();
    for prefix in undeclared {
        let inherited = bound
            .iter()
            .find(|(name, _)| *name == prefix)
            .map(|(_, uri)| uri.as_slice());
        match (prefix.as_slice(), inherited) {
            (b"m", Some(uri)) if uri == MATH.as_bytes() => {},
            (b"a", Some(uri)) if uri == dml.as_bytes() => {},
            (b"m", _) => declarations.push_str(&format!(r#" xmlns:m="{MATH}""#)),
            (b"a", _) => declarations.push_str(&format!(r#" xmlns:a="{dml}""#)),
            (_, Some(_)) => {},
            (_, None) => {
                return Err(invalid(
                    "equation OMML uses a namespace prefix it does not declare",
                ));
            },
        }
    }
    let mut out = String::with_capacity(omml.len() + declarations.len());
    out.push_str(&omml[..at]);
    out.push_str(&declarations);
    out.push_str(&omml[at..]);
    Ok(out)
}

/// The OMML form of an equation being replaced: a display paragraph with
/// its justification, or an inline `m:oMath`.
#[derive(Debug, Clone, Copy)]
enum Form {
    Inline,
    Display(Option<Justification>),
}

impl Form {
    fn of(omml: &str) -> Result<Self> {
        let mut reader = Reader::from_reader(omml.as_bytes());
        let mut root = true;
        loop {
            match reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
            {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let local = element.local_name();
                    if root {
                        if local.as_ref() != b"oMathPara" {
                            return Ok(Self::Inline);
                        }
                        root = false;
                    } else if local.as_ref() == b"jc" {
                        let value = element
                            .checked_attributes()
                            .flatten()
                            .find(|attribute| attribute.key.local_name().as_ref() == b"val")
                            .map(|attribute| attribute.value.into_owned());
                        return Ok(Self::Display(match value.as_deref() {
                            Some(b"left") => Some(Justification::Left),
                            Some(b"right") => Some(Justification::Right),
                            Some(b"center") => Some(Justification::Center),
                            Some(b"centerGroup") => Some(Justification::CenterGroup),
                            _ => None,
                        }));
                    } else if local.as_ref() == b"oMath" {
                        return Ok(Self::Display(None));
                    }
                },
                Event::Eof => {
                    return Ok(if root {
                        Self::Inline
                    } else {
                        Self::Display(None)
                    });
                },
                _ => {},
            }
        }
    }
}

/// OMML ready for a slide, with what its conversion lost.
struct Prepared {
    omml: String,
    losses: Vec<String>,
}

fn prepare(
    source: EquationSource<'_>,
    form: Form,
    run: &HostRun,
    dml: &str,
    bound: &Bindings,
    max_bytes: usize,
) -> Result<Prepared> {
    let input = match source {
        EquationSource::Latex(text) | EquationSource::Omml(text) => text,
    };
    if input.len() > max_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::PatchBytes,
            observed: input.len(),
            maximum: max_bytes,
        });
    }
    let limits = Limits::DEFAULT;
    let (omml, losses) = match source {
        EquationSource::Latex(latex) => {
            let converted = match form {
                Form::Inline => {
                    litchi_equation::latex_to_omml(latex, OmmlHost::DrawingMl, &limits)?
                },
                Form::Display(justification) => litchi_equation::latex_to_omml_paragraph(
                    latex,
                    OmmlHost::DrawingMl,
                    justification,
                    &limits,
                )?,
            };
            let losses = converted.losses().to_vec();
            let mut omml = apply_host_runs(converted.xml(), run)?;
            if dml != TRANSITIONAL_DML {
                omml = omml.replacen(
                    &format!(r#" xmlns:a="{TRANSITIONAL_DML}""#),
                    &format!(r#" xmlns:a="{dml}""#),
                    1,
                );
            }
            (omml, losses)
        },
        EquationSource::Omml(omml) => (admit_omml(omml, dml, bound)?, Vec::new()),
    };
    // The converter must read back what is written, so the reader of the
    // slide and every later edit can.
    litchi_equation::omml_to_latex(&omml, &limits)?;
    Ok(Prepared { omml, losses })
}

/// A picture of the equation, the size it covers on the slide, in EMU,
/// and the scale it was drawn at.
struct Picture {
    png: litchi_equation::Png,
    width: i64,
    height: i64,
    scale: f64,
}

/// The pixels of a rendered equation, at most.
const MAX_PICTURE_PIXELS: f64 = 16_000_000.0;
/// The widest or tallest rendered equation, in pixels (the renderer's own
/// bound is 8192).
const MAX_PICTURE_SIDE: f64 = 8_000.0;

fn render(omml: &str, font_size_pt: f64) -> Result<Picture> {
    let options = PngOptions::new(font_size_pt)?;
    // Draw at PICTURE_SCALE, or lower when that would exceed the pixel
    // bounds: a long equation is drawn coarser instead of refused.
    let (width_pt, height_pt) = litchi_equation::omml_size_pt(omml, &options, &Limits::DEFAULT)?;
    let pixels_per_point = 96.0 / 72.0;
    let (width_px, height_px) = (
        (width_pt * pixels_per_point).max(1.0),
        (height_pt * pixels_per_point).max(1.0),
    );
    let scale = PICTURE_SCALE
        .min(MAX_PICTURE_SIDE / width_px)
        .min(MAX_PICTURE_SIDE / height_px)
        .min((MAX_PICTURE_PIXELS / (width_px * height_px)).sqrt());
    let options = options
        .with_scale(scale)?
        .with_max_pixels(MAX_PICTURE_PIXELS as u64 + 1_000_000)?;
    let png = litchi_equation::omml_to_png(omml, &options, &Limits::DEFAULT)?;
    let (width, height) = png.size_pt();
    let emu = |points: f64| -> i64 {
        // Bounded by the renderer's pixel limit, far below i64::MAX.
        #[allow(
            clippy::cast_possible_truncation,
            reason = "the picture is at most a few thousand points wide"
        )]
        let value = (points * EMU_PER_POINT).ceil() as i64;
        value.max(1)
    };
    Ok(Picture {
        width: emu(width),
        height: emu(height),
        png,
        scale,
    })
}

/// The frame size of a shape that fits a picture inside its text insets.
fn fitted(picture: &Picture, insets: [i64; 4]) -> Result<(i64, i64)> {
    let [left, top, right, bottom] = insets.map(|inset| inset.max(0));
    let width = picture
        .width
        .checked_add(left)
        .and_then(|value| value.checked_add(right))
        .ok_or_else(|| invalid("the equation frame is too large"))?;
    let height = picture
        .height
        .checked_add(top)
        .and_then(|value| value.checked_add(bottom))
        .ok_or_else(|| invalid("the equation frame is too large"))?;
    Ok((width, height))
}

/// The fill of a fallback picture: the whole frame, as PowerPoint writes it
/// for a picture of the shape's text body.
const FALLBACK_FILL: &str = "<a:stretch><a:fillRect/></a:stretch>";

/// The fallback picture of a shape: its whole frame, transparent, with the
/// equation at its own size inside the text insets, placed by `place`
/// (horizontal and vertical, 0 at the start, 1 at the end) and scaled down
/// only to fit, as PowerPoint draws its picture of the text body. Readers
/// stretch the picture over the frame, so the equation keeps its
/// proportions only when the picture has the frame's.
fn fallback_picture(
    picture: &Picture,
    (width, height): (i64, i64),
    insets: [i64; 4],
    place: (f64, f64),
) -> Result<Vec<u8>> {
    #![allow(
        clippy::cast_precision_loss,
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "frames and insets are far below 2^52 EMU, canvases a few thousand pixels"
    )]
    let [left, top, right, bottom] = insets.map(|inset| inset.max(0) as f64);
    let (width, height) = (width.max(1) as f64, height.max(1) as f64);
    // Pixels per EMU at the rendering scale (96 pixels per inch at 1x).
    let natural = picture.scale * 96.0 / 914_400.0;
    let mut density = natural;
    let pixels = (width * density).ceil() * (height * density).ceil();
    if pixels > MAX_FALLBACK_PIXELS {
        density *= (MAX_FALLBACK_PIXELS / pixels).sqrt();
    }
    let room = (
        (width - left - right).max(1.0),
        (height - top - bottom).max(1.0),
    );
    let fit = (room.0 / picture.width as f64)
        .min(room.1 / picture.height as f64)
        .min(1.0);
    let zoom = (density / natural * fit).clamp(f64::MIN_POSITIVE, 1.0);
    let shown = (picture.width as f64 * fit, picture.height as f64 * fit);
    let x = left + (room.0 - shown.0).max(0.0) * place.0.clamp(0.0, 1.0);
    let y = top + (room.1 - shown.1).max(0.0) * place.1.clamp(0.0, 1.0);
    let canvas = |extent: f64| ((extent * density).ceil().max(1.0)) as u32;
    Ok(picture
        .png
        .placed(
            canvas(width),
            canvas(height),
            x * density,
            y * density,
            zoom,
        )?
        .into_bytes())
}

/// What a placeholder inherits; see [`Edit::placeholder_inheritance`].
#[derive(Debug, Default)]
struct Inherited {
    extent: Option<(i64, i64)>,
    text_size_pt: Option<f64>,
}

/// The first-level text size, in points, of a master's `p:txStyles` style
/// (`titleStyle` or `bodyStyle`): its `a:lvl1pPr/a:defRPr@sz`.
fn master_text_size(xml: &[u8], style: &[u8]) -> Result<Option<f64>> {
    let mut reader = Reader::from_reader(xml);
    // How far into `p:txStyles`, the style and its first level we are.
    let mut path: Vec<Vec<u8>> = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                let inside = path.len() >= 3
                    && path[path.len() - 3] == b"txStyles"
                    && path[path.len() - 2] == style
                    && path[path.len() - 1] == b"lvl1pPr";
                if inside && local == b"defRPr" {
                    return Ok(element
                        .checked_attributes()
                        .flatten()
                        .find(|attribute| attribute.key.as_ref() == b"sz")
                        .and_then(|attribute| {
                            std::str::from_utf8(&attribute.value)
                                .ok()?
                                .trim()
                                .parse::<u32>()
                                .ok()
                        })
                        .filter(|size| (100..=400_000).contains(size))
                        .map(|size| f64::from(size) / 100.0));
                }
                if matches!(event, Event::Start(_)) {
                    path.push(local);
                }
            },
            Event::End(_) => {
                path.pop();
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// The `cx` and `cy` of an `a:ext` start tag.
fn extent(tag: &[u8]) -> Option<(i64, i64)> {
    let value = |name: &[u8]| -> Option<i64> {
        let (span, _, _) = attribute_value(tag, name)?;
        std::str::from_utf8(&tag[span]).ok()?.trim().parse().ok()
    };
    Some((value(b"cx")?, value(b"cy")?)).filter(|(width, height)| *width > 0 && *height > 0)
}

/// The largest `cNvPr` ID anywhere in the slide, both branches of every
/// markup-compatibility block included.
fn max_raw_shape_id(xml: &[u8]) -> Result<u32> {
    let mut reader = Reader::from_reader(xml);
    let mut maximum = 1u32;
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(ref element) | Event::Empty(ref element)
                if element.local_name().as_ref() == b"cNvPr" =>
            {
                for attribute in element.checked_attributes().flatten() {
                    if attribute.key.as_ref() == b"id"
                        && let Some(id) = std::str::from_utf8(&attribute.value)
                            .ok()
                            .and_then(|value| value.trim().parse::<u32>().ok())
                    {
                        maximum = maximum.max(id);
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(maximum)
}

/// Replacements of byte ranges of one slide, applied in one pass.
#[derive(Default)]
struct Splice {
    edits: Vec<(Range<usize>, String)>,
}

impl Splice {
    fn replace(&mut self, range: Range<usize>, with: impl Into<String>) {
        self.edits.push((range, with.into()));
    }

    fn apply(mut self, xml: &[u8], max_bytes: usize) -> Result<Vec<u8>> {
        self.edits.sort_by_key(|(range, _)| range.start);
        let mut out = Vec::new();
        let mut at = 0usize;
        let size = self.edits.iter().fold(xml.len(), |size, (range, with)| {
            size.saturating_sub(range.len()).saturating_add(with.len())
        });
        if size > max_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: size,
                maximum: max_bytes,
            });
        }
        out.try_reserve_exact(size)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation equation XML",
                source,
            })?;
        for (range, with) in &self.edits {
            if range.start < at || range.end > xml.len() {
                return Err(invalid("equation edits overlap"));
            }
            out.extend_from_slice(&xml[at..range.start]);
            out.extend_from_slice(with.as_bytes());
            at = range.end;
        }
        out.extend_from_slice(&xml[at..]);
        Ok(out)
    }
}

/// Replace the `cx` and `cy` values of an `a:ext` start tag at `tag`.
fn resize(splice: &mut Splice, xml: &[u8], tag: &Range<usize>, size: (i64, i64)) -> Result<()> {
    let bytes = &xml[tag.clone()];
    for (name, value) in [(b"cx".as_slice(), size.0), (b"cy", size.1)] {
        let (span, _, _) = attribute_value(bytes, name)
            .ok_or_else(|| invalid("the equation frame has no extent"))?;
        splice.replace(
            tag.start + span.start..tag.start + span.end,
            value.to_string(),
        );
    }
    Ok(())
}

/// The equation at `index` and the block holding it, refused when the
/// equation is not in the text of a shape-level markup-compatibility block
/// with one Choice and no markup compatibility of its own.
fn located(
    sites: &[Site],
    blocks: &[Block],
    index: usize,
    operation: &'static str,
) -> Result<(Site, usize)> {
    let site = sites.get(index).ok_or(Error::IndexOutOfBounds {
        index,
        len: sites.len(),
    })?;
    let block = site
        .block
        .filter(|block| blocks.get(*block).is_some_and(|block| block.shape_level))
        .ok_or_else(|| {
            refused(
                operation,
                "only an equation in the text of a shape-level markup-compatibility block is edited",
            )
        })?;
    editable_block(&blocks[block], operation)?;
    Ok((site.clone(), block))
}

fn editable_block(block: &Block, operation: &'static str) -> Result<()> {
    if block.nested || block.choices != 1 || block.choice_shape.is_none() {
        return Err(refused(
            operation,
            "the equation's shape holds markup-compatibility content of its own",
        ));
    }
    Ok(())
}

/// Whether the block's shape holds this one equation and no other text.
const fn equation_only(block: &Block) -> bool {
    block.equations == 1 && !block.other_content
}

/// The horizontal position of an equation in its room, from 0 (left) to 1
/// (right): a display paragraph's justification (centred as a group when it
/// names none), otherwise its paragraph's alignment.
fn horizontal(omml: &str, align: Option<&str>) -> Result<f64> {
    Ok(match Form::of(omml)? {
        Form::Display(Some(Justification::Left)) => 0.0,
        Form::Display(Some(Justification::Right)) => 1.0,
        Form::Display(_) => 0.5,
        Form::Inline => match align {
            Some("ctr") => 0.5,
            Some("r") => 1.0,
            _ => 0.0,
        },
    })
}

/// The vertical position of the text in its room, from 0 (top) to 1
/// (bottom), from the text anchor (`a:bodyPr@anchor`, top when absent).
fn vertical(anchor: Option<&str>) -> f64 {
    match anchor {
        Some("ctr") => 0.5,
        Some("b") => 1.0,
        _ => 0.0,
    }
}

/// The text of a run that stands for an equation in a text fallback: its
/// linear form (UnicodeMath, else LaTeX), escaped for character data.
fn linear_run(omml: &str) -> String {
    let limits = Limits::DEFAULT;
    let linear = litchi_equation::omml_to_unicode_math(omml, &limits)
        .or_else(|_error| litchi_equation::omml_to_latex(omml, &limits))
        .unwrap_or_default();
    let properties = HostRun::of(omml).ok().flatten().map_or_else(
        || r#"<a:rPr lang="en-US"/>"#.to_owned(),
        |run| run.element(&[], "", true),
    );
    format!(
        "<a:r>{properties}<a:t>{}</a:t></a:r>",
        quick_xml::escape::escape(linear.as_str())
    )
}

/// The qualified name of the element whose start tag begins `tag`.
fn tag_name(tag: &[u8]) -> Result<&str> {
    let tag = tag
        .strip_prefix(b"<")
        .ok_or_else(|| invalid("the equation block is malformed"))?;
    let end = tag
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')
        .unwrap_or(tag.len());
    std::str::from_utf8(&tag[..end]).map_err(|_error| invalid("the equation block is malformed"))
}

/// The edit that gives block `index` of `xml` a text Fallback: a copy of
/// its Choice shape in which every equation is a run of its linear form,
/// with no picture fill. Office ignores the Fallback; other readers show
/// the text.
fn text_fallback(xml: &[u8], index: usize) -> Result<(Range<usize>, String)> {
    let (sites, blocks) = sites::scan(xml)?;
    let block = blocks
        .get(index)
        .ok_or_else(|| invalid("the equation block disappeared"))?;
    let shape = block
        .choice_shape
        .clone()
        .ok_or_else(|| invalid("the equation block has no Choice shape"))?;
    let mut copy = Splice::default();
    for site in sites
        .iter()
        .filter(|site| shape.start <= site.element.start && site.element.end <= shape.end)
    {
        let omml = std::str::from_utf8(&xml[site.content.clone()])
            .map_err(|error| Error::Xml(error.to_string()))?;
        copy.replace(
            site.element.start - shape.start..site.element.end - shape.start,
            linear_run(omml),
        );
    }
    let copy = String::from_utf8(copy.apply(&xml[shape], usize::MAX)?)
        .map_err(|error| Error::Xml(error.to_string()))?;
    let block_name = tag_name(&xml[block.range.clone()])?;
    match (&block.fallback, &block.fallback_tag) {
        (Some((range, empty)), Some(tag)) => {
            let name = tag_name(&xml[tag.clone()])?;
            let open = std::str::from_utf8(&xml[tag.clone()])
                .map_err(|error| Error::Xml(error.to_string()))?;
            let open = if *empty {
                let trimmed = open.trim_end_matches('>').trim_end().trim_end_matches('/');
                format!("{trimmed}>")
            } else {
                open.to_owned()
            };
            Ok((range.clone(), format!("{open}{copy}</{name}>")))
        },
        _ => {
            // No Fallback yet: one goes before the block's end tag, under
            // the block's own prefix.
            let prefix = block_name
                .split_once(':')
                .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
            let close = format!("</{block_name}>");
            let at = block
                .range
                .end
                .checked_sub(close.len())
                .filter(|at| xml[*at..block.range.end] == *close.as_bytes())
                .ok_or_else(|| invalid("the equation block is malformed"))?;
            Ok((
                at..at,
                format!("<{prefix}Fallback>{copy}</{prefix}Fallback>"),
            ))
        },
    }
}

/// A block's Choice shape standing on its own: `shape` (its bytes) with
/// the namespace declarations of the start tags before it in the block
/// (`mc:AlternateContent`, `mc:Choice`) that it does not make itself, so
/// every prefix it uses stays declared.
fn unwrapped_shape(wrappers: &[u8], shape: &[u8]) -> Result<String> {
    let declarations = |tags: &[u8], stop_after_first: bool| -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
        let mut reader = Reader::from_reader(tags);
        let mut out = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                    for attribute in element.checked_attributes().flatten() {
                        let key = attribute.key.as_ref();
                        if key == b"xmlns" || key.starts_with(b"xmlns:") {
                            out.push((key.to_vec(), attribute.value.into_owned()));
                        }
                    }
                    if stop_after_first {
                        return Ok(out);
                    }
                },
                // The wrappers are start tags only; their end is not here.
                Ok(Event::Eof) | Err(_) => return Ok(out),
                Ok(_) => {},
            }
        }
    };
    let own = declarations(shape, true)?;
    let name = tag_name(shape)?;
    let mut carried = String::new();
    let mut seen: Vec<Vec<u8>> = own.iter().map(|(key, _)| key.clone()).collect();
    // The innermost declaration of a prefix wins: walk from the Choice out.
    for (key, value) in declarations(wrappers, false)?.into_iter().rev() {
        if seen.contains(&key) {
            continue;
        }
        let key_text =
            String::from_utf8(key.clone()).map_err(|error| Error::Xml(error.to_string()))?;
        let value_text = String::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))?;
        push_attribute(&mut carried, &key_text, &value_text);
        seen.push(key);
    }
    let text = std::str::from_utf8(shape).map_err(|error| Error::Xml(error.to_string()))?;
    let at = 1 + name.len();
    Ok(format!("{}{carried}{}", &text[..at], &text[at..]))
}

/// Refuse XML that uses a namespace prefix it does not declare, before it
/// is staged: a commit refuses such a slide, with every other edit.
fn require_declared_prefixes(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let unknown = matches!(namespace, ResolveResult::Unknown(_));
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if unknown {
                    return Err(invalid(
                        "the equation edit leaves a namespace prefix undeclared",
                    ));
                }
                for attribute in element.checked_attributes().flatten() {
                    let key = attribute.key;
                    if key.as_ref().starts_with(b"xmlns") || key.prefix().is_none() {
                        continue;
                    }
                    if key.as_ref().starts_with(b"xml:") {
                        continue;
                    }
                    if matches!(
                        reader.resolver().resolve_attribute(key).0,
                        ResolveResult::Unknown(_)
                    ) {
                        return Err(invalid(
                            "the equation edit leaves a namespace prefix undeclared",
                        ));
                    }
                }
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

fn cnvpr_id(xml: &[u8]) -> Result<Option<u32>> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(ref element) | Event::Empty(ref element)
                if element.local_name().as_ref() == b"cNvPr" =>
            {
                return Ok(element
                    .checked_attributes()
                    .flatten()
                    .find(|attribute| attribute.key.as_ref() == b"id")
                    .and_then(|attribute| {
                        std::str::from_utf8(&attribute.value)
                            .ok()
                            .and_then(|value| value.trim().parse().ok())
                    }));
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

impl Edit {
    /// What a slide placeholder shape (`shape_id`) inherits from its layout
    /// and master (ECMA-376 Part 1, §19.3.1.36: matched by index on the
    /// layout and by type on the master): its frame size, when it has none
    /// of its own, and the master's first-level text size for its kind
    /// (§19.3.1.51 `txStyles`). Both `None` for a shape that is no
    /// placeholder.
    fn placeholder_inheritance(
        &self,
        slide: &PackURI,
        xml: &[u8],
        shape_id: u32,
    ) -> Result<Inherited> {
        let limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(xml, limits)?;
        let Some(placeholder) = scene
            .iter()
            .find(|shape| shape.id() == Some(shape_id))
            .and_then(crate::shape::Shape::placeholder)
        else {
            return Ok(Inherited::default());
        };
        // An omitted type is `obj`, whose master counterpart is the body.
        let kind = placeholder.kind().unwrap_or("obj");
        let master_kind = match kind {
            "obj" | "subTitle" => "body",
            "ctrTitle" => "title",
            other => other,
        };
        let related = |owner: &PackURI, kind: &str| -> Result<Option<PackURI>> {
            let part = self.working.get_part(owner)?;
            part.rels()
                .iter()
                .filter(|relationship| !relationship.is_external())
                .find(|relationship| relationship.reltype().ends_with(kind))
                .map(litchi_opc::Relationship::target_partname)
                .transpose()
                .map_err(Error::Opc)
        };
        let mut inherited = Inherited::default();
        let mut owner = slide.clone();
        for (relationship, by_index) in [("/slideLayout", true), ("/slideMaster", false)] {
            let Some(next) = related(&owner, relationship)? else {
                break;
            };
            owner = next;
            let blob = self.working.get_part(&owner)?.blob();
            if inherited.extent.is_none() {
                let scene = crate::shape::Scene::read(blob, limits)?;
                let found = scene.iter().find_map(|shape| {
                    let candidate = shape.placeholder()?;
                    let candidate_kind = candidate.kind().unwrap_or("obj");
                    let same = if by_index && placeholder.index() != 0 {
                        candidate.index() == placeholder.index()
                    } else if by_index {
                        candidate_kind == kind
                    } else {
                        candidate_kind == master_kind
                    };
                    same.then(|| shape.bounds())
                });
                if let Some(Some(bounds)) = found {
                    inherited.extent = Some((bounds.width(), bounds.height()))
                        .filter(|(width, height)| *width > 0 && *height > 0);
                }
            }
            if !by_index {
                let style: &[u8] = if master_kind == "title" {
                    b"titleStyle"
                } else {
                    b"bodyStyle"
                };
                inherited.text_size_pt = master_text_size(blob, style)?;
            }
        }
        Ok(inherited)
    }

    /// The conformance-specific DrawingML namespace of a slide.
    fn drawing_namespace(&self, xml: &[u8]) -> Result<crate::media_parts::Conformance> {
        crate::media_parts::document_conformance_with_limits(
            xml,
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )
    }

    /// Stage `next` as the XML of `slide` in `candidate` (which may hold a
    /// new picture part), after checking that every shape but the edited
    /// one (`shape`, which a reader without Office Math may only now see)
    /// keeps its place and ID, that a `removed` shape is gone, and that
    /// `equations` remain; the relationships `before` referenced and `next`
    /// no longer does are removed, with the parts nothing else uses.
    #[allow(
        clippy::too_many_arguments,
        reason = "one staging step; the arguments are its inputs and its checks"
    )]
    fn publish_equation_edit(
        &mut self,
        slide: &PackURI,
        mut candidate: super::Working,
        before: &[u8],
        next: Vec<u8>,
        equations: usize,
        shape: u32,
        removed: bool,
    ) -> Result<()> {
        let shape_limits = self.semantic_limits().shape();
        require_declared_prefixes(&next)?;
        if removed {
            crate::opened::xml::validate_shape_removal_dependencies(
                &next,
                &std::collections::BTreeSet::from([shape]),
            )?;
        }
        let others = |xml: &[u8]| -> Result<Vec<Option<u32>>> {
            Ok(crate::shape::Scene::read(xml, shape_limits)?
                .iter()
                .map(crate::shape::Shape::id)
                .filter(|id| *id != Some(shape))
                .collect())
        };
        let gone = !crate::shape::Scene::read(&next, shape_limits)?
            .iter()
            .any(|staged| staged.id() == Some(shape));
        if others(before)? != others(&next)?
            || (removed && !gone)
            || sites::scan(&next)?.0.len() != equations
        {
            return Err(invalid("the equation edit did not round-trip"));
        }
        let owner = candidate.get_part(slide)?;
        let referenced = shape_relationship_ids(before, owner.rels())?;
        let retained = shape_relationship_ids(&next, owner.rels())?;
        let mut roots = Vec::new();
        {
            let owner = candidate.get_part_mut(slide)?;
            owner.set_blob(next);
            for id in referenced.difference(&retained) {
                if let Some(relationship) = owner.rels_mut().remove(id)
                    && !relationship.is_external()
                {
                    roots.push(relationship.target_partname()?);
                }
            }
        }
        remove_unreferenced_dependencies(&mut candidate, roots)?;
        self.working = candidate;
        self.note_scene_read(slide);
        Ok(())
    }

    /// Add an equation as a new text box at `origin` (`x`, `y` in EMU), in
    /// the form PowerPoint writes: a markup-compatibility block whose Choice
    /// holds the OMML in a display paragraph and whose Fallback shows a PNG
    /// picture of it, with the picture part and its relationship.
    ///
    /// The text box sizes itself to the equation at `font_size_pt` (18 when
    /// `None`); a given size is also written on the equation's runs.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Equation`] for LaTeX or OMML that does not convert
    /// or render, an error for malformed OMML, a missing slide, malformed
    /// slide XML or exceeded bounds. A failure leaves the edit unchanged.
    pub fn add_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        source: EquationSource<'_>,
        origin: (i64, i64),
        font_size_pt: Option<f64>,
    ) -> Result<EquationWrite> {
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let owner = self.working.get_part(&selected.part_name)?;
        let xml = owner.blob();
        let conformance = self.drawing_namespace(xml)?;
        let (pml, dml, rel) = conformance_namespaces(conformance);
        let size = match font_size_pt {
            Some(points) => {
                // ST_TextFontSize: 1 to 4000 points in hundredths.
                if !(points.is_finite() && (1.0..=4000.0).contains(&points)) {
                    return Err(invalid("the equation text size must be 1 to 4000 points"));
                }
                #[allow(
                    clippy::cast_possible_truncation,
                    clippy::cast_sign_loss,
                    reason = "the size is checked to lie in 1..=4000 points"
                )]
                Some((points * 100.0).round() as u32)
            },
            None => None,
        };
        let run = HostRun::native(size);
        let prepared = prepare(
            source,
            Form::Display(Some(Justification::CenterGroup)),
            &run,
            dml,
            &root_bindings(xml)?,
            max_patch_bytes,
        )?;
        let picture = render(&prepared.omml, font_size_pt.unwrap_or(DEFAULT_FONT_SIZE_PT))?;
        let (width, height) = fitted(&picture, DEFAULT_INSETS)?;
        let png = fallback_picture(&picture, (width, height), DEFAULT_INSETS, (0.5, 0.0))?;
        let (x, y) = origin;
        if x.checked_add(width).is_none() || y.checked_add(height).is_none() {
            return Err(invalid(
                "the equation frame lies outside the coordinate space",
            ));
        }
        let shape_id = max_raw_shape_id(xml)?
            .max(max_shape_reference_id(xml)?.unwrap_or(1))
            .checked_add(1)
            .ok_or_else(|| invalid("opened-presentation equation shape ID overflow"))?;
        let before = sites::scan(xml)?.0.len();
        let part_name = free_part_name(&self.working, "/ppt/media/image", ".png")?;
        let mut candidate = self.working.clone();
        candidate.try_add_part(Box::new(BlobPart::new(
            part_name.clone(),
            "image/png".to_owned(),
            png,
        )))?;
        let target = part_name.relative_ref(selected.part_name.base_uri());
        let relationship = candidate
            .get_part_mut(&selected.part_name)?
            .relate_to(&target, conformance_image_relationship(conformance));
        let name = format!("TextBox {}", shape_id - 1);
        let end_size = size.map_or_else(String::new, |size| format!(r#" sz="{size}""#));
        let geometry = format!(
            r#"<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#
        );
        let fragment = format!(
            concat!(
                r#"<mc:AlternateContent xmlns:mc="{mce}" xmlns:a14="{a14}" xmlns:p="{pml}" xmlns:a="{dml}" xmlns:r="{rel}">"#,
                r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
                r#"<p:spPr>{geometry}<a:noFill/></p:spPr>"#,
                r#"<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a14:m>{omml}</a14:m><a:endParaRPr lang="en-US"{end_size} dirty="0"/></a:p></p:txBody></p:sp></mc:Choice>"#,
                r#"<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr txBox="1">{locks}</p:cNvSpPr><p:nvPr/></p:nvSpPr>"#,
                r#"<p:spPr>{geometry}<a:blipFill rotWithShape="1"><a:blip r:embed="{relationship}"/>{fill}</a:blipFill></p:spPr>"#,
                "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"><a:noFill/></a:rPr><a:t>\u{a0}</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>",
            ),
            mce = MCE,
            a14 = A14,
            pml = pml,
            dml = dml,
            rel = rel,
            id = shape_id,
            name = name,
            geometry = geometry,
            omml = prepared.omml,
            end_size = end_size,
            locks = FALLBACK_LOCKS,
            relationship = relationship,
            fill = FALLBACK_FILL,
        );
        let next = crate::opened::xml::append_shape_with_limit(
            candidate.get_part(&selected.part_name)?.blob(),
            fragment.as_bytes(),
            shape_limits.output_bytes(),
        )?;
        ensure_shape_id(&next, shape_id, "equation", shape_limits)?;
        let (after, _) = sites::scan(&next)?;
        if after.len() != before + 1 {
            return Err(invalid("the equation shape did not round-trip"));
        }
        candidate.get_part_mut(&selected.part_name)?.set_blob(next);
        self.working = candidate;
        self.note_scene_read(&selected.part_name);
        Ok(EquationWrite {
            shape_id,
            index: before,
            losses: prepared.losses,
        })
    }

    /// Add an equation at the end of the text of the shape `shape_id` (its
    /// `cNvPr` ID): inline, after the last paragraph's runs. A shape that
    /// holds no equation yet is wrapped in the markup-compatibility block
    /// PowerPoint writes; its Fallback (here, and whenever a shape's text
    /// mixes equations with other text) is a copy of the shape in which each
    /// equation is a run of its linear form (UnicodeMath), for readers
    /// without Office Math.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing shape or one without a text paragraph,
    /// a refusal for a shape with markup-compatibility content of its own,
    /// [`Error::Equation`] for input that does not convert, and an error for
    /// malformed OMML or slide XML. A failure leaves the edit unchanged.
    pub fn add_equation_to_shape<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape_id: u32,
        source: EquationSource<'_>,
    ) -> Result<EquationWrite> {
        const OPERATION: &str = "add_equation_to_shape";
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let (_, dml, _) = conformance_namespaces(self.drawing_namespace(&xml)?);
        let (sites, blocks) = sites::scan(&xml)?;
        let mut held = None;
        for (index, block) in blocks.iter().enumerate() {
            if let Some(shape) = &block.choice_shape
                && cnvpr_id(&xml[shape.clone()])? == Some(shape_id)
            {
                editable_block(block, OPERATION)?;
                held = Some((index, shape.clone()));
                break;
            }
        }
        let (shape, wrapped) = match held {
            Some((_, ref shape)) => (shape.clone(), false),
            None => (
                sites::plain_shape(&xml, shape_id)?
                    .ok_or_else(|| refused(OPERATION, "the target is not a text shape"))?,
                true,
            ),
        };
        let end = sites::text_end(&xml[shape.clone()])?
            .ok_or_else(|| refused(OPERATION, "the shape has no text paragraph"))?;
        let prepared = prepare(
            source,
            Form::Inline,
            // The equation takes the text size of the paragraph it ends.
            &HostRun::native(end.size),
            dml,
            &root_bindings(&xml)?,
            max_patch_bytes,
        )?;
        let math = format!(r#"<a14:m xmlns:a14="{A14}">{}</a14:m>"#, prepared.omml);
        // An empty paragraph (`<a:p/>`) opens to take the equation.
        let (at, math) = match end.place {
            sites::TextPlace::Before(at) => (at..at, math),
            sites::TextPlace::Empty(tag) => {
                let name = tag_name(&xml[shape.start + tag.start..shape.start + tag.end])?;
                let open =
                    std::str::from_utf8(&xml[shape.start + tag.start..shape.start + tag.end])
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .trim_end_matches('>')
                        .trim_end()
                        .trim_end_matches('/')
                        .to_owned();
                (tag, format!("{open}>{math}</{name}>"))
            },
        };
        let mut splice = Splice::default();
        if wrapped {
            let mut inner = Splice::default();
            inner.replace(at, math);
            let inner = String::from_utf8(inner.apply(&xml[shape.clone()], usize::MAX)?)
                .map_err(|error| Error::Xml(error.to_string()))?;
            splice.replace(
                shape.clone(),
                format!(
                    r#"<mc:AlternateContent xmlns:mc="{MCE}" xmlns:a14="{A14}"><mc:Choice Requires="a14">{inner}</mc:Choice></mc:AlternateContent>"#
                ),
            );
        } else {
            splice.replace(shape.start + at.start..shape.start + at.end, math);
        }
        let staged = splice.apply(&xml, shape_limits.output_bytes())?;
        let (_, staged_blocks) = sites::scan(&staged)?;
        let block = if wrapped {
            staged_blocks
                .iter()
                .position(|block| block.range.start == shape.start)
        } else {
            held.as_ref().map(|(index, _)| *index)
        }
        .ok_or_else(|| invalid("the equation block disappeared"))?;
        let (range, fallback) = text_fallback(&staged, block)?;
        let mut splice = Splice::default();
        splice.replace(range, fallback);
        let next = splice.apply(&staged, shape_limits.output_bytes())?;
        let (after, after_blocks) = sites::scan(&next)?;
        let choice = after_blocks
            .get(block)
            .and_then(|block| block.choice_shape.clone())
            .ok_or_else(|| invalid("the equation block disappeared"))?;
        let index = after
            .iter()
            .rposition(|site| choice.start <= site.element.start && site.element.end <= choice.end)
            .ok_or_else(|| invalid("the added equation did not round-trip"))?;
        let candidate = self.working.clone();
        self.publish_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len() + 1,
            shape_id,
            false,
        )?;
        Ok(EquationWrite {
            shape_id,
            index,
            losses: prepared.losses,
        })
    }

    /// Replace the content of the slide's equation at `equation` (its
    /// position among [`Slide::equations`](crate::slide::Slide::equations)).
    /// The Choice branch gets the new OMML. When the shape holds only this
    /// equation and its Fallback shows a picture, the picture is redrawn (a
    /// shape that resizes to its text, `a:spAutoFit`, takes the new size);
    /// otherwise the Fallback becomes a copy of the shape with every
    /// equation as a run of its linear form. The shape keeps its ID,
    /// position and other properties; pictures and relationships nothing
    /// else uses are removed. LaTeX keeps the equation's form (inline or
    /// display, with its justification) and its first run's host properties
    /// (text size, fonts). Returns `None` when the OMML is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing equation, a
    /// refusal for an equation outside a shape-level markup-compatibility
    /// block, [`Error::Equation`] for input that does not convert or render,
    /// and an error for malformed OMML or slide XML. A failure leaves the
    /// edit unchanged.
    pub fn set_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        equation: usize,
        source: EquationSource<'_>,
    ) -> Result<Option<EquationWrite>> {
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let (sites, blocks) = sites::scan(&xml)?;
        let (site, index) = located(&sites, &blocks, equation, "set_equation")?;
        let block = blocks[index].clone();
        let conformance = self.drawing_namespace(&xml)?;
        let (_, dml, _) = conformance_namespaces(conformance);
        let old = std::str::from_utf8(&xml[site.content.clone()])
            .map_err(|error| Error::Xml(error.to_string()))?;
        let run = HostRun::of(old)?.unwrap_or_else(|| HostRun::native(None));
        let prepared = prepare(
            source,
            Form::of(old)?,
            &run,
            dml,
            &root_bindings(&xml)?,
            max_patch_bytes,
        )?;
        if prepared.omml == old.trim() {
            return Ok(None);
        }
        let shape_id = cnvpr_id(&xml[block.range.clone()])?
            .ok_or_else(|| invalid("the equation shape has no ID"))?;
        let mut splice = Splice::default();
        let candidate = if equation_only(&block) && block.blip.is_some() {
            let inherited = self.placeholder_inheritance(&selected.part_name, &xml, shape_id)?;
            let picture = render(
                &prepared.omml,
                run.font_size_pt()
                    .or(inherited.text_size_pt)
                    .unwrap_or(DEFAULT_FONT_SIZE_PT),
            )?;
            let blip = block
                .blip
                .clone()
                .ok_or_else(|| invalid("the equation fallback has no picture"))?;
            let (embed, _, _) = attribute_value(&xml[blip.clone()], b"embed")
                .ok_or_else(|| invalid("the equation picture is not embedded"))?;
            let embed = blip.start + embed.start..blip.start + embed.end;
            // The frame the picture covers: resized to the new picture when
            // the shape fits its text, otherwise its own or its
            // placeholder's; a placeholder frame that cannot be found is
            // taken as the picture's.
            let frame = match (&block.choice_extent, &block.fallback_extent) {
                (Some(choice), Some(fallback)) if block.autofit => {
                    let size = fitted(&picture, block.insets)?;
                    resize(&mut splice, &xml, choice, size)?;
                    resize(&mut splice, &xml, fallback, size)?;
                    Some(size)
                },
                (Some(choice), _) => extent(&xml[choice.clone()]),
                (None, _) => inherited.extent,
            };
            let frame = match frame {
                Some(frame) => frame,
                None => fitted(&picture, block.insets)?,
            };
            if let Some(stretch) = block.stretch.clone() {
                splice.replace(stretch, FALLBACK_FILL);
            }
            let place = (
                horizontal(&prepared.omml, site.align.as_deref())?,
                vertical(block.anchor.as_deref()),
            );
            let png = fallback_picture(&picture, frame, block.insets, place)?;
            let part_name = free_part_name(&self.working, "/ppt/media/image", ".png")?;
            let mut candidate = self.working.clone();
            candidate.try_add_part(Box::new(BlobPart::new(
                part_name.clone(),
                "image/png".to_owned(),
                png,
            )))?;
            let target = part_name.relative_ref(selected.part_name.base_uri());
            let new_id = candidate
                .get_part_mut(&selected.part_name)?
                .relate_to(&target, conformance_image_relationship(conformance));
            splice.replace(embed, new_id);
            splice.replace(site.content.clone(), prepared.omml);
            candidate
        } else {
            splice.replace(site.content.clone(), prepared.omml);
            self.working.clone()
        };
        let mut next = splice.apply(&xml, shape_limits.output_bytes())?;
        if !(equation_only(&block) && block.blip.is_some()) {
            let (range, fallback) = text_fallback(&next, index)?;
            let mut splice = Splice::default();
            splice.replace(range, fallback);
            next = splice.apply(&next, shape_limits.output_bytes())?;
        }
        self.publish_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len(),
            shape_id,
            false,
        )?;
        Ok(Some(EquationWrite {
            shape_id,
            index: equation,
            losses: prepared.losses,
        }))
    }

    /// Remove the slide's equation at `equation` (its position among
    /// [`Slide::equations`](crate::slide::Slide::equations)). A shape whose
    /// text is only this equation leaves with it: the whole
    /// markup-compatibility block, with or without a picture Fallback. In a
    /// shape whose text holds more, only the equation leaves: the Fallback
    /// becomes a copy of the shape with every remaining equation as a run of
    /// its linear form, and a shape left without equations is unwrapped to
    /// its plain Choice shape. Pictures and relationships nothing else uses
    /// are removed. Returns the ID of the shape that held the equation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing equation, a
    /// refusal for an equation outside a shape-level markup-compatibility
    /// block or a removed shape other markup refers to (connectors,
    /// animations), and an error for malformed slide XML. A failure leaves
    /// the edit unchanged.
    pub fn remove_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        equation: usize,
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let (sites, blocks) = sites::scan(&xml)?;
        let (site, index) = located(&sites, &blocks, equation, "remove_equation")?;
        let block = &blocks[index];
        let shape_id = cnvpr_id(&xml[block.range.clone()])?
            .ok_or_else(|| invalid("the equation shape has no ID"))?;
        let mut splice = Splice::default();
        if equation_only(block) {
            splice.replace(block.range.clone(), String::new());
            let next = splice.apply(&xml, shape_limits.output_bytes())?;
            let candidate = self.working.clone();
            self.publish_equation_edit(
                &selected.part_name,
                candidate,
                &xml,
                next,
                sites.len() - 1,
                shape_id,
                true,
            )?;
            return Ok(shape_id);
        }
        splice.replace(site.element.clone(), String::new());
        let staged = splice.apply(&xml, shape_limits.output_bytes())?;
        let (_, staged_blocks) = sites::scan(&staged)?;
        let staged_block = staged_blocks
            .get(index)
            .ok_or_else(|| invalid("the equation block disappeared"))?;
        let choice_shape = staged_block
            .choice_shape
            .clone()
            .ok_or_else(|| invalid("the equation block has no Choice shape"))?;
        let mut splice = Splice::default();
        if staged_block.equations == 0 {
            // No equation is left: the shape stands without the block, with
            // the block's declarations it relies on.
            let shape = unwrapped_shape(
                &staged[staged_block.range.start..choice_shape.start],
                &staged[choice_shape],
            )?;
            splice.replace(staged_block.range.clone(), shape);
        } else {
            let (range, fallback) = text_fallback(&staged, index)?;
            splice.replace(range, fallback);
        }
        let next = splice.apply(&staged, shape_limits.output_bytes())?;
        let candidate = self.working.clone();
        self.publish_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len() - 1,
            shape_id,
            false,
        )?;
        Ok(shape_id)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const DML_BINDING: (&[u8], &[u8]) = (b"a", TRANSITIONAL_DML.as_bytes());

    fn bound() -> Bindings {
        vec![(DML_BINDING.0.to_vec(), DML_BINDING.1.to_vec())]
    }

    #[test]
    fn converter_runs_get_powerpoints_host_properties() {
        let run = HostRun::native(Some(2400));
        let omml = r#"<m:oMath xmlns:m="m" xmlns:a="a"><m:r><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r><m:r><a:rPr b="1"/><m:t>B</m:t></m:r><m:r><a:rPr i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><m:t>c</m:t></m:r></m:oMath>"#;
        let written = apply_host_runs(omml, &run).unwrap();
        let fonts = DEFAULT_RUN_FONTS;
        assert_eq!(
            written,
            format!(
                concat!(
                    r#"<m:oMath xmlns:m="m" xmlns:a="a">"#,
                    r#"<m:r><a:rPr lang="en-US" sz="2400" i="1">{fonts}</a:rPr><m:t>x</m:t></m:r>"#,
                    r#"<m:r><m:rPr><m:sty m:val="p"/></m:rPr><a:rPr lang="en-US" sz="2400">{fonts}</a:rPr><m:t>sin</m:t></m:r>"#,
                    r#"<m:r><a:rPr lang="en-US" sz="2400" i="1" b="1">{fonts}</a:rPr><m:t>B</m:t></m:r>"#,
                    r#"<m:r><a:rPr lang="en-US" sz="2400" i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>{fonts}</a:rPr><m:t>c</m:t></m:r>"#,
                    "</m:oMath>"
                ),
                fonts = fonts
            )
        );
    }

    #[test]
    fn the_first_math_run_properties_are_the_template() {
        let omml = r#"<m:oMath><m:r><a:rPr lang="de-DE" sz="3200" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>x</m:t></m:r></m:oMath>"#;
        let run = HostRun::of(omml).unwrap().unwrap();
        assert_eq!(run.font_size_pt(), Some(32.0));
        assert_eq!(
            run.children,
            [(
                "latin".to_owned(),
                r#"<a:latin typeface="Cambria Math"/>"#.to_owned()
            )]
        );
        assert_eq!(
            run.element(&[], "", false),
            r#"<a:rPr lang="de-DE" sz="3200" i="1"><a:latin typeface="Cambria Math"/></a:rPr>"#
        );
        // Fonts after a fill are still found.
        let filled = r#"<m:oMath><m:r><a:rPr i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr><m:t>x</m:t></m:r></m:oMath>"#;
        let template = HostRun::of(filled).unwrap().unwrap();
        assert_eq!(template.children.len(), 2);
        // The template's fill is kept, or replaced by the converter's.
        assert_eq!(
            template.element(&[], "", false),
            r#"<a:rPr i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Cambria Math"/></a:rPr>"#
        );
        assert_eq!(
            template.element(&[], "<a:noFill/>", true),
            r#"<a:rPr><a:noFill/><a:latin typeface="Cambria Math"/></a:rPr>"#
        );
        // A run without host properties, or another prefix, is no template.
        assert!(
            HostRun::of("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")
                .unwrap()
                .is_none()
        );
        assert!(
            HostRun::of(r#"<m:oMath><m:r><w:rPr/><m:t>x</m:t></m:r></m:oMath>"#)
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn caller_omml_is_admitted_with_its_missing_prefixes() {
        let dml = TRANSITIONAL_DML;
        assert_eq!(
            admit_omml(
                "  <m:oMath><m:r><a:rPr/><m:t>x</m:t></m:r></m:oMath> ",
                dml,
                &bound()
            )
            .unwrap(),
            format!(r#"<m:oMath xmlns:m="{MATH}"><m:r><a:rPr/><m:t>x</m:t></m:r></m:oMath>"#)
        );
        assert_eq!(
            admit_omml("<m:oMath><a:rPr/></m:oMath>", dml, &Vec::new()).unwrap(),
            format!(r#"<m:oMath xmlns:m="{MATH}" xmlns:a="{dml}"><a:rPr/></m:oMath>"#)
        );
        let declared = format!(r#"<q:oMathPara xmlns:q="{MATH}"/>"#);
        assert_eq!(admit_omml(&declared, dml, &bound()).unwrap(), declared);
        for bad in [
            "",
            "x",
            "<m:r/>",
            "<m:oMath/><m:oMath/>",
            "<m:oMath><z:r/></m:oMath>",
            "<m:oMath>",
            r#"<?xml version="1.0"?><m:oMath/>"#,
            r#"<m:oMath xmlns:m="urn:other"/>"#,
        ] {
            assert!(admit_omml(bad, dml, &bound()).is_err(), "{bad}");
        }
    }

    #[test]
    fn the_form_of_an_equation_is_kept() {
        assert!(matches!(Form::of("<m:oMath/>").unwrap(), Form::Inline));
        assert!(matches!(
            Form::of(r#"<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath/></m:oMathPara>"#).unwrap(),
            Form::Display(Some(Justification::Left))
        ));
        assert!(matches!(
            Form::of("<m:oMathPara><m:oMath/></m:oMathPara>").unwrap(),
            Form::Display(None)
        ));
    }

    #[test]
    fn master_text_sizes_are_read_per_style() {
        let master = br#"<p:sldMaster xmlns:p="p" xmlns:a="a"><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="0"><a:defRPr sz="2800"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2400"/></a:lvl2pPr></p:bodyStyle></p:txStyles></p:sldMaster>"#;
        assert_eq!(master_text_size(master, b"titleStyle").unwrap(), Some(44.0));
        assert_eq!(master_text_size(master, b"bodyStyle").unwrap(), Some(28.0));
        assert_eq!(master_text_size(master, b"otherStyle").unwrap(), None);
    }

    #[test]
    fn frames_and_extents() {
        assert_eq!(extent(br#"<a:ext cx="10" cy="20"/>"#), Some((10, 20)));
        assert_eq!(extent(br#"<a:ext cx="0" cy="20"/>"#), None);
        assert_eq!(extent(br#"<a:ext cy="20"/>"#), None);
        assert_eq!(
            max_raw_shape_id(br#"<r><p:cNvPr id="7"/><x><a:cNvPr id="12"/></x></r>"#).unwrap(),
            12
        );
    }
}
