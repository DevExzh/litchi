#![cfg_attr(
    not(feature = "equation"),
    allow(
        dead_code,
        reason = "equation objects are written with the equation feature"
    )
)]
//! The markup of a paragraph's OLE equation objects: insertion, rewriting
//! and removal.
//!
//! An object is counted as [`crate::Paragraph::equation_objects`] counts
//! it. Each edit rewrites only the object's own run (or, in a run that holds
//! other content, its `w:object`) and records it as one exact element
//! replacement of the paragraph; the parts the markup names are the
//! package editor's to stage.
//!
//! The picture shape of an object takes the VML picture-frame shape type,
//! `#_x0000_t75`, which Word defines once per part, in the first object or
//! picture that uses it. A new object references the part's definition and
//! writes one only when the part has none. A removed object that holds the
//! part's definition while other shapes of the part still use it leaves the
//! definition behind in a `w:pict` of its own, as Word writes shape-type
//! definitions, so the remaining shapes keep their type.

use std::ops::Range;
use std::sync::Arc;

use litchi_core::Position;
use quick_xml::events::Event;

use super::super::{buffer, staging_add};
use super::{Edit, ElementPath, Error, Refusal, Result, bindings_at, locate, not_found, refusal};
use crate::object::{EquationSpan, equation_spans};
use crate::picture_spans::{attribute_named, with_attribute};

const VML: &str = "urn:schemas-microsoft-com:vml";
const OFFICE: &str = "urn:schemas-microsoft-com:office:office";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// The VML picture-frame shape type as Word writes it beside the first
/// picture or object of a part.
const PICTURE_SHAPETYPE: &str = concat!(
    r#"<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" "#,
    r#"path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/>"#,
    r#"<v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/>"#,
    r#"<v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/>"#,
    r#"<v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/>"#,
    r#"<v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/>"#,
    r#"<v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>"#,
    r#"<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>"#,
    r#"<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>"#
);

/// The run properties `EG_RPrBase` (ECMA-376 Part 1, §17.3.2.28) orders
/// after `w:position`, and the change marker that closes a `w:rPr`.
const AFTER_POSITION: [&[u8]; 17] = [
    b"sz",
    b"szCs",
    b"highlight",
    b"u",
    b"effect",
    b"bdr",
    b"shd",
    b"fitText",
    b"vertAlign",
    b"rtl",
    b"cs",
    b"em",
    b"lang",
    b"eastAsianLayout",
    b"specVanish",
    b"oMath",
    b"rPrChange",
];

/// The markup of a new equation object.
pub(crate) struct ObjectMarkup<'a> {
    /// The object's ProgID.
    pub(crate) prog_id: &'a str,
    /// The number of the VML shape, `_x0000_i{shape}`.
    pub(crate) shape: u32,
    /// The number of the object, `_{object}`.
    pub(crate) object: u64,
    /// The relationship of the embedded compound file.
    pub(crate) object_relationship: &'a str,
    /// The relationship of the preview image.
    pub(crate) preview_relationship: &'a str,
    /// The shown size in points.
    pub(crate) size_pt: (f64, f64),
    /// How far the run is lowered, in half-points.
    pub(crate) lowering: u32,
}

/// A change of an existing object's markup: its size, its lowering and,
/// when given, the preview image and the compound file it names.
pub(crate) struct ObjectRewrite<'a> {
    pub(crate) size_pt: (f64, f64),
    pub(crate) lowering: u32,
    pub(crate) preview_relationship: Option<&'a str>,
    pub(crate) object_relationship: Option<&'a str>,
}

/// A size in points as VML writes it: at most two decimals.
fn points(value: f64) -> String {
    let text = format!("{value:.2}");
    let text = text.trim_end_matches('0').trim_end_matches('.');
    if text.is_empty() || text == "-" {
        "0".to_owned()
    } else {
        text.to_owned()
    }
}

/// Twentieths of a point, as `w:dxaOrig` and `w:dyaOrig` take them.
#[expect(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "the value is clamped to the attribute's range first"
)]
fn twips(value: f64) -> u32 {
    (value * 20.0).round().clamp(1.0, f64::from(i32::MAX)) as u32
}

/// The `style` of an object's shape with its width and height replaced (or
/// appended), every other declaration kept.
fn restyled(style: &str, (width, height): (f64, f64)) -> String {
    let mut declarations: Vec<String> = Vec::new();
    let (mut wrote_width, mut wrote_height) = (false, false);
    for declaration in style.split(';') {
        let name = declaration.split(':').next().unwrap_or("").trim();
        if name == "width" {
            declarations.push(format!("width:{}pt", points(width)));
            wrote_width = true;
        } else if name == "height" {
            declarations.push(format!("height:{}pt", points(height)));
            wrote_height = true;
        } else if !declaration.trim().is_empty() {
            declarations.push(declaration.to_owned());
        }
    }
    if !wrote_width {
        declarations.push(format!("width:{}pt", points(width)));
    }
    if !wrote_height {
        declarations.push(format!("height:{}pt", points(height)));
    }
    declarations.join(";")
}

/// Whether `xml` defines the VML picture-frame shape type.
fn defines_picture_shapetype(xml: &[u8]) -> Result<bool> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid Word XML".into())))?
        {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"shapetype" =>
            {
                let id = attribute_value(&element, b"id");
                if id.as_deref() == Some("_x0000_t75") {
                    return Ok(true);
                }
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
    }
}

fn attribute_value(element: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    element
        .unchecked_attributes()
        .filter_map(std::result::Result::ok)
        .find(|attribute| attribute.key.as_ref() == key)
        .and_then(|attribute| String::from_utf8(attribute.value.into_owned()).ok())
}

/// The largest VML shape number (`_x0000_iN`) and OLE object number
/// (`ObjectID="_N"`) `xml` uses.
pub(crate) fn object_numbers(xml: &[u8]) -> (Option<u32>, Option<u64>) {
    fn digits_after(xml: &[u8], marker: &[u8]) -> Vec<u64> {
        let mut found = Vec::new();
        let mut at = 0;
        while let Some(offset) = xml.get(at..).and_then(|tail| {
            tail.windows(marker.len())
                .position(|window| window == marker)
        }) {
            let start = at + offset + marker.len();
            let end = xml[start..]
                .iter()
                .position(|byte| !byte.is_ascii_digit())
                .map_or(xml.len(), |length| start + length);
            if end > start
                && let Some(number) = std::str::from_utf8(&xml[start..end])
                    .ok()
                    .and_then(|text| text.parse::<u64>().ok())
            {
                found.push(number);
            }
            at = start;
        }
        found
    }
    let shape = digits_after(xml, b"_x0000_i")
        .into_iter()
        .filter_map(|number| u32::try_from(number).ok())
        .max();
    let object = digits_after(xml, b"ObjectID=\"_")
        .into_iter()
        .chain(digits_after(xml, b"ObjectID='_"))
        .max();
    (shape, object)
}

/// The namespace declarations (`xmlns`, `xmlns:*`) of a start tag, as
/// written.
fn declarations_of(tag: &[u8]) -> Vec<u8> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    let mut reader = quick_xml::Reader::from_reader(tag);
    let mut out = Vec::new();
    if let Ok(Event::Start(element) | Event::Empty(element)) = reader.read_event() {
        for attribute in element.unchecked_attributes().flatten() {
            let key = attribute.key.as_ref();
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                out.push(b' ');
                out.extend_from_slice(key);
                out.extend_from_slice(b"=\"");
                out.extend_from_slice(&attribute.value);
                out.push(b'"');
            }
        }
    }
    out
}

/// The qualified name of the element starting `tag` (`w:object`), and its
/// prefix (`w`, or empty).
fn element_prefix(tag: &[u8]) -> Vec<u8> {
    let name_end = tag
        .iter()
        .skip(1)
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .map_or(tag.len(), |position| position + 1);
    let name = tag.get(1..name_end).unwrap_or_default();
    name.iter()
        .position(|byte| *byte == b':')
        .map(|colon| name[..colon].to_vec())
        .unwrap_or_default()
}

fn qualified(prefix: &[u8], local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{}:{local}", String::from_utf8_lossy(prefix))
    }
}

/// What decides the text size of a run of a paragraph: the paragraph's
/// style, the run's character style and its own size (`w:sz`, in
/// half-points).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct FontHint {
    pub(crate) paragraph_style: Option<String>,
    pub(crate) run_style: Option<String>,
    pub(crate) size: Option<u32>,
}

/// The paragraph style of the paragraph `xml` and, for each of its runs in
/// document order (direct or nested), its start, character style and size.
fn run_hints(xml: &[u8]) -> Result<(Option<String>, Vec<(usize, FontHint)>)> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut paragraph_style = None;
    let mut runs: Vec<(usize, FontHint)> = Vec::new();
    // Open elements by local name.
    let mut stack: Vec<Vec<u8>> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| super::invalid("position overflows"))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid paragraph XML".into())))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                let parents: Vec<&[u8]> = stack.iter().rev().take(2).map(Vec::as_slice).collect();
                let value = || attribute_local(element, b"val");
                match (local.as_slice(), parents.as_slice()) {
                    (b"pStyle", [b"pPr", b"p", ..]) => paragraph_style = value(),
                    (b"r", _) if stack.last().map(Vec::as_slice) != Some(b"rPr".as_slice()) => {
                        runs.push((start, FontHint::default()));
                    },
                    (b"rStyle", [b"rPr", b"r", ..]) => {
                        if let Some((_, hint)) = runs.last_mut() {
                            hint.run_style = value();
                        }
                    },
                    (b"sz", [b"rPr", b"r", ..]) => {
                        if let Some((_, hint)) = runs.last_mut() {
                            hint.size = value().and_then(|text| text.trim().parse().ok());
                        }
                    },
                    _ => {},
                }
                if matches!(event, Event::Start(_)) {
                    stack.push(local);
                }
            },
            Event::End(_) => {
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
    }
    for (_, hint) in &mut runs {
        hint.paragraph_style.clone_from(&paragraph_style);
    }
    Ok((paragraph_style, runs))
}

/// The value of the attribute whose local name is `local`.
fn attribute_local(element: &quick_xml::events::BytesStart<'_>, local: &[u8]) -> Option<String> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    element
        .unchecked_attributes()
        .filter_map(std::result::Result::ok)
        .find(|attribute| attribute.key.local_name().as_ref() == local)
        .and_then(|attribute| String::from_utf8(attribute.value.into_owned()).ok())
}

impl Edit {
    /// What decides the text size of a new object of the paragraph at
    /// `paragraph` inserted before its direct run `at`: that run's, or for
    /// `None` the paragraph's last run's, or the paragraph's style alone.
    pub(crate) fn insertion_font_hint(
        &self,
        paragraph: &ElementPath,
        at: Option<Position>,
    ) -> Result<FontHint> {
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let located = locate(&self.projected, paragraph.steps())?;
        let bytes = self.element_bytes(located)?;
        let (paragraph_style, runs) = run_hints(&bytes)?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        // An empty paragraph (`<w:p/>`) has no runs.
        let starts: Vec<usize> = match super::element_children(&bytes, b"p", &inherited) {
            Ok(direct) => (0..direct.count(b"r"))
                .filter_map(|index| direct.index_of(b"r", Position::new(index)))
                .filter_map(|child| direct.insertion_offset(child))
                .collect(),
            Err(_) if super::super::root_tag(&bytes)?.1 => Vec::new(),
            Err(error) => return Err(error),
        };
        let chosen = match at {
            Some(position) => starts.get(position.get()).copied(),
            None => starts.last().copied(),
        };
        Ok(chosen
            .and_then(|start| runs.iter().find(|(run, _)| *run == start))
            .map_or_else(
                || FontHint {
                    paragraph_style: paragraph_style.clone(),
                    ..FontHint::default()
                },
                |(_, hint)| hint.clone(),
            ))
    }

    /// What decides the text size of the equation object at `index` among
    /// the paragraph's: its own run's.
    pub(crate) fn object_font_hint(
        &self,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<FontHint> {
        let (bytes, spans) = self.paragraph_objects(paragraph)?;
        let span = spans
            .get(index.get())
            .ok_or_else(|| refusal(Refusal::EquationNotFound))?;
        let (paragraph_style, runs) = run_hints(&bytes)?;
        Ok(span
            .run
            .as_ref()
            .and_then(|run| runs.iter().find(|(start, _)| *start == run.start))
            .map_or_else(
                || FontHint {
                    paragraph_style,
                    ..FontHint::default()
                },
                |(_, hint)| hint.clone(),
            ))
    }

    /// The paragraph at `paragraph` and its equation objects.
    fn paragraph_objects(
        &self,
        paragraph: &ElementPath,
    ) -> Result<(Arc<Vec<u8>>, Vec<EquationSpan>)> {
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let located = locate(&self.projected, paragraph.steps())?;
        let before = self.element_bytes(located)?;
        let spans = equation_spans(&before).map_err(Error::Document)?;
        Ok((before, spans))
    }

    /// The equation object at `index` among the paragraph's.
    pub(crate) fn equation_object_span(
        &self,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<EquationSpan> {
        let (_, spans) = self.paragraph_objects(paragraph)?;
        let span = spans
            .into_iter()
            .nth(index.get())
            .ok_or_else(|| refusal(Refusal::EquationNotFound))?;
        if span.alternate {
            return Err(refusal(Refusal::ComplexContent));
        }
        Ok(span)
    }

    /// Insert a run holding the equation object `markup` into the paragraph
    /// at `paragraph`, before its direct run `at` (after its last child for
    /// `None`), and return the object's position among the paragraph's
    /// equation objects.
    pub(crate) fn insert_equation_object_markup_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        markup: &ObjectMarkup<'_>,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_equation_object_at")?;
        if self.is_strict() {
            // Strict WordprocessingML has no VML (ISO/IEC 29500-1 holds it
            // in Part 4, Transitional Migration Features).
            return Err(refusal(Refusal::TransitionalValue));
        }
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let located = locate(&self.projected, paragraph.steps())?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        let bound = |prefix: &str, namespace: &str| {
            inherited
                .iter()
                .rev()
                .find(|(bound, _)| bound.as_deref() == Some(prefix.as_bytes()))
                .is_some_and(|(_, uri)| uri.as_slice() == namespace.as_bytes())
        };
        let mut declarations = String::new();
        for (prefix, namespace) in [("v", VML), ("o", OFFICE), ("r", RELATIONSHIPS)] {
            if !bound(prefix, namespace) {
                declarations.push_str(&format!(" xmlns:{prefix}=\"{namespace}\""));
            }
        }
        let shapetype = if defines_picture_shapetype(self.projected.xml_bytes())? {
            ""
        } else {
            PICTURE_SHAPETYPE
        };
        let shape = format!("_x0000_i{}", markup.shape);
        let escape = |text: &str| {
            let mut escaped = Vec::new();
            super::super::super::append_literal_escaped_xml(&mut escaped, text);
            String::from_utf8(escaped).map_err(|_| refusal(Refusal::InvalidPropertyValue))
        };
        let prog_id = escape(markup.prog_id)?;
        let object_relationship = escape(markup.object_relationship)?;
        let preview_relationship = escape(markup.preview_relationship)?;
        let (width, height) = markup.size_pt;
        let body = format!(
            concat!(
                "{shapetype}",
                r##"<v:shape id="{shape}" type="#_x0000_t75" style="width:{width}pt;height:{height}pt" o:ole="">"##,
                r#"<v:imagedata r:id="{preview}" o:title=""/></v:shape>"#,
                r#"<o:OLEObject Type="Embed" ProgID="{prog_id}" ShapeID="{shape}" DrawAspect="Content" "#,
                r#"ObjectID="_{object}" r:id="{relationship}"/>"#
            ),
            shapetype = shapetype,
            shape = shape,
            width = points(width),
            height = points(height),
            preview = preview_relationship,
            prog_id = prog_id,
            object = markup.object,
            relationship = object_relationship,
        );
        let (dxa, dya) = (twips(width), twips(height));
        let lowering = markup.lowering;
        let shape_id = shape.clone();
        self.insert_inline_counted(paragraph, at, b"r", false, |prefix| {
            let (prefix, declare) = if prefix.is_empty() {
                (b"w".as_slice(), true)
            } else {
                (prefix, false)
            };
            let w = String::from_utf8_lossy(prefix).into_owned();
            let mut xml = buffer(staging_add(body.len(), 512)?)?;
            xml.extend_from_slice(format!("<{w}:r").as_bytes());
            if declare {
                xml.extend_from_slice(
                    format!(
                        " xmlns:{w}=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                    )
                    .as_bytes(),
                );
            }
            xml.push(b'>');
            if lowering > 0 {
                xml.extend_from_slice(
                    format!("<{w}:rPr><{w}:position {w}:val=\"-{lowering}\"/></{w}:rPr>").as_bytes(),
                );
            }
            xml.extend_from_slice(
                format!(
                    "<{w}:object {w}:dxaOrig=\"{dxa}\" {w}:dyaOrig=\"{dya}\"{declarations}>"
                )
                .as_bytes(),
            );
            xml.extend_from_slice(body.as_bytes());
            xml.extend_from_slice(format!("</{w}:object></{w}:r>").as_bytes());
            Ok(xml)
        })?;
        let (_, spans) = self.paragraph_objects(paragraph)?;
        spans
            .iter()
            .position(|span| span.shape_id.as_deref() == Some(shape_id.as_str()))
            .map(Position::new)
            .ok_or_else(|| super::invalid("the inserted equation object did not read back"))
    }

    /// Rewrite the size, lowering and preview of the equation object at
    /// `index` among the paragraph's, and return whether anything changed.
    /// The lowering is written only on a run the object is the only content
    /// of.
    pub(crate) fn rewrite_equation_object_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
        rewrite: &ObjectRewrite<'_>,
    ) -> Result<bool> {
        self.ensure_unmanaged("set_equation_object_at")?;
        let (before, spans) = self.paragraph_objects(paragraph)?;
        let span = spans
            .get(index.get())
            .ok_or_else(|| refusal(Refusal::EquationNotFound))?;
        if span.alternate {
            return Err(refusal(Refusal::ComplexContent));
        }
        let unreadable = || super::invalid("the equation object has an unreadable tag");
        let slice = |range: &Range<usize>| before.get(range.clone()).ok_or_else(unreadable);
        let mut edits: Vec<(Range<usize>, Vec<u8>)> = Vec::new();
        // `w:object`: its original size in twentieths of a point.
        let tag = slice(&span.object_tag)?;
        let prefix = element_prefix(tag);
        let (width, height) = rewrite.size_pt;
        let mut object_tag = tag.to_vec();
        for (local, value) in [("dxaOrig", twips(width)), ("dyaOrig", twips(height))] {
            let key =
                attribute_named(&object_tag, local).unwrap_or_else(|| qualified(&prefix, local));
            object_tag =
                with_attribute(&object_tag, &key, &value.to_string()).ok_or_else(unreadable)?;
        }
        edits.push((span.object_tag.clone(), object_tag));
        // `v:shape`: its shown size.
        if let Some(range) = &span.shape_tag {
            let style = restyled(span.style.as_deref().unwrap_or(""), rewrite.size_pt);
            let tag = with_attribute(slice(range)?, "style", &style).ok_or_else(unreadable)?;
            edits.push((range.clone(), tag));
        }
        // `v:imagedata`: the preview it names.
        if let (Some(range), Some(preview)) = (&span.imagedata_tag, rewrite.preview_relationship) {
            let tag = slice(range)?;
            let key = attribute_named(tag, "id").ok_or_else(unreadable)?;
            edits.push((
                range.clone(),
                with_attribute(tag, &key, preview).ok_or_else(unreadable)?,
            ));
        }
        // `o:OLEObject`: the compound file it names.
        if let Some(relationship) = rewrite.object_relationship {
            let range = ole_object_tag(&before, &span.object)?;
            let tag = slice(&range)?;
            let key = attribute_named(tag, "id").ok_or_else(unreadable)?;
            edits.push((
                range,
                with_attribute(tag, &key, relationship).ok_or_else(unreadable)?,
            ));
        }
        // The run's lowering.
        if span.sole
            && let Some(run) = &span.run
        {
            let lowering = format!("-{}", rewrite.lowering);
            match (&span.position, &span.run_properties, rewrite.lowering) {
                (Some(position), _, 0) => edits.push((position.clone(), Vec::new())),
                (Some(position), _, _) => {
                    let tag = slice(position)?;
                    let key =
                        attribute_named(tag, "val").unwrap_or_else(|| qualified(&prefix, "val"));
                    edits.push((
                        position.clone(),
                        with_attribute(tag, &key, &lowering).ok_or_else(unreadable)?,
                    ));
                },
                (None, _, 0) => {},
                (None, Some(properties), _) => {
                    let element = format!(
                        "<{} {}=\"{lowering}\"/>",
                        qualified(&prefix, "position"),
                        qualified(&prefix, "val")
                    );
                    edits.push(position_insertion(&before, properties, &prefix, element)?);
                },
                (None, None, _) => {
                    let run_tag_end = element_start_end(&before, run.start)?;
                    let element = format!(
                        "<{rpr}><{position} {val}=\"{lowering}\"/></{rpr}>",
                        rpr = qualified(&prefix, "rPr"),
                        position = qualified(&prefix, "position"),
                        val = qualified(&prefix, "val"),
                    );
                    edits.push((run_tag_end..run_tag_end, element.into_bytes()));
                },
            }
        }
        edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
        let mut after = before.as_ref().clone();
        for (range, bytes) in edits {
            after.splice(range, bytes);
        }
        if after.as_slice() == before.as_slice() {
            return Ok(false);
        }
        self.replace_element(paragraph, &before, &Arc::new(after))?;
        Ok(true)
    }

    /// Remove the equation object at `index` among the paragraph's: its run
    /// when the run holds nothing else, its `w:object` otherwise. Returns
    /// the removed object's span, read before the removal.
    pub(crate) fn remove_equation_object_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<EquationSpan> {
        self.ensure_unmanaged("remove_equation_object_at")?;
        let (before, spans) = self.paragraph_objects(paragraph)?;
        let span = spans
            .get(index.get())
            .ok_or_else(|| refusal(Refusal::EquationNotFound))?
            .clone();
        if span.alternate {
            return Err(refusal(Refusal::ComplexContent));
        }
        let unreadable = || super::invalid("the equation object has an unreadable tag");
        let (range, whole_run) = match (&span.run, span.sole) {
            (Some(run), true) => (run.clone(), true),
            _ => (span.object.clone(), false),
        };
        // The part's picture shape type stays while another shape uses it.
        let mut replacement = Vec::new();
        if let Some(shapetype) = &span.shapetype {
            let object = before.get(span.object.clone()).ok_or_else(unreadable)?;
            let reference = b"#_x0000_t75";
            let count = |xml: &[u8]| {
                xml.windows(reference.len())
                    .filter(|window| *window == reference)
                    .count()
            };
            let remaining = count(self.projected.xml_bytes()).saturating_sub(count(object));
            if remaining > 0 {
                let tag = before.get(span.object_tag.clone()).ok_or_else(unreadable)?;
                let prefix = element_prefix(tag);
                let pict = qualified(&prefix, "pict");
                let mut carrier = format!("<{pict}").into_bytes();
                carrier.extend_from_slice(&declarations_of(tag));
                carrier.push(b'>');
                carrier.extend_from_slice(before.get(shapetype.clone()).ok_or_else(unreadable)?);
                carrier.extend_from_slice(format!("</{pict}>").as_bytes());
                if whole_run {
                    let run = qualified(&prefix, "r");
                    replacement = format!("<{run}>").into_bytes();
                    replacement.extend_from_slice(&carrier);
                    replacement.extend_from_slice(format!("</{run}>").as_bytes());
                } else {
                    replacement = carrier;
                }
            }
        }
        let mut after = buffer(before.len())?;
        after.extend_from_slice(before.get(..range.start).ok_or_else(unreadable)?);
        after.extend_from_slice(&replacement);
        after.extend_from_slice(before.get(range.end..).ok_or_else(unreadable)?);
        self.replace_element(paragraph, &before, &Arc::new(after))?;
        Ok(span)
    }
}

/// The `o:OLEObject` start tag of the object at `object` of `xml`.
fn ole_object_tag(xml: &[u8], object: &Range<usize>) -> Result<Range<usize>> {
    let bytes = xml
        .get(object.clone())
        .ok_or_else(|| super::invalid("equation object outside its XML"))?;
    let mut reader = quick_xml::Reader::from_reader(bytes);
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| super::invalid("position overflows"))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid equation object".into())))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| super::invalid("position overflows"))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"OLEObject" =>
            {
                return Ok(object.start + start..object.start + end);
            },
            Event::Eof => return Err(super::invalid("the equation object has no OLEObject")),
            _ => {},
        }
    }
}

/// The end of the start tag of the element starting at `start` of `xml`.
fn element_start_end(xml: &[u8], start: usize) -> Result<usize> {
    let tail = xml
        .get(start..)
        .ok_or_else(|| super::invalid("element range outside its XML"))?;
    let mut reader = quick_xml::Reader::from_reader(tail);
    match reader.read_event() {
        Ok(Event::Start(_)) => usize::try_from(reader.buffer_position())
            .map(|length| start + length)
            .map_err(|_| super::invalid("position overflows")),
        _ => Err(super::invalid("the run has no start tag")),
    }
}

/// Where a new `w:position` goes in the run properties `properties` of
/// `xml`, in schema order: before the first property ordered after it, or
/// at the end. An empty `w:rPr/` is opened.
fn position_insertion(
    xml: &[u8],
    properties: &Range<usize>,
    prefix: &[u8],
    element: String,
) -> Result<(Range<usize>, Vec<u8>)> {
    let bytes = xml
        .get(properties.clone())
        .ok_or_else(|| super::invalid("run properties outside their XML"))?;
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut depth = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| super::invalid("position overflows"))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid run properties".into())))?;
        match event {
            Event::Empty(_) if depth == 0 => {
                // `<w:rPr/>`.
                let rpr = qualified(prefix, "rPr");
                return Ok((
                    properties.clone(),
                    format!("<{rpr}>{element}</{rpr}>").into_bytes(),
                ));
            },
            Event::Start(ref child) | Event::Empty(ref child) if depth == 1 => {
                if AFTER_POSITION.contains(&child.local_name().as_ref()) {
                    let at = properties.start + start;
                    return Ok((at..at, element.into_bytes()));
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            },
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                if depth == 1 {
                    let at = properties.start + start;
                    return Ok((at..at, element.into_bytes()));
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => return Err(super::invalid("unterminated run properties")),
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sizes_and_styles_are_written_as_word_writes_them() {
        assert_eq!(points(30.849_999), "30.85");
        assert_eq!(points(12.0), "12");
        assert_eq!(twips(30.85), 617);
        assert_eq!(
            restyled("width:1pt;height:2pt;mso-position:x", (3.5, 4.25)),
            "width:3.5pt;height:4.25pt;mso-position:x"
        );
        assert_eq!(restyled("", (1.0, 2.0)), "width:1pt;height:2pt");
    }

    #[test]
    fn shape_and_object_numbers_are_found() {
        let xml = br#"<v:shape id="_x0000_i1027"/><o:OLEObject ShapeID="_x0000_i1025" ObjectID="_1490538856"/><v:shapetype id="_x0000_t75"/>"#;
        assert_eq!(object_numbers(xml), (Some(1027), Some(1_490_538_856)));
        assert_eq!(object_numbers(b"<w:p/>"), (None, None));
        assert!(defines_picture_shapetype(xml).unwrap());
    }

    #[test]
    fn a_position_goes_where_the_schema_orders_it() {
        let xml = br#"<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>"#;
        let (range, bytes) =
            position_insertion(xml, &(0..xml.len()), b"w", "<w:position/>".into()).unwrap();
        assert_eq!(range, 13..13);
        assert_eq!(bytes, b"<w:position/>");
        let xml = br#"<w:rPr><w:b/></w:rPr>"#;
        let (range, _) =
            position_insertion(xml, &(0..xml.len()), b"w", "<w:position/>".into()).unwrap();
        assert_eq!(range, 13..13);
        let xml = br#"<w:rPr/>"#;
        let (range, bytes) =
            position_insertion(xml, &(0..xml.len()), b"w", "<w:position/>".into()).unwrap();
        assert_eq!(range, 0..8);
        assert_eq!(bytes, b"<w:rPr><w:position/></w:rPr>");
    }

    #[test]
    fn declarations_and_prefixes_are_read_from_tags() {
        assert_eq!(
            declarations_of(br#"<w:object w:dxaOrig="1" xmlns:v="urn:v">"#),
            br#" xmlns:v="urn:v""#
        );
        assert_eq!(element_prefix(b"<w:object w:dxaOrig=\"1\">"), b"w");
        assert_eq!(element_prefix(b"<object>"), b"");
    }
}

#[cfg(test)]
mod hint_tests {
    use super::*;

    #[test]
    fn run_hints_read_styles_and_sizes() {
        let xml = br#"<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Big"/><w:sz w:val="28"/></w:rPr><w:t>x</w:t></w:r><w:r><w:t>y</w:t></w:r></w:p>"#;
        let (style, runs) = run_hints(xml).unwrap();
        assert_eq!(style.as_deref(), Some("Body"));
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].1.run_style.as_deref(), Some("Big"));
        assert_eq!(runs[0].1.size, Some(28));
        assert_eq!(runs[1].1.size, None);
        assert_eq!(runs[1].1.paragraph_style.as_deref(), Some("Body"));
    }
}
