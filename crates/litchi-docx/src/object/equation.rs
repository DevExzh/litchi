//! OLE equation objects: the MathType (`Equation.DSMT4`) and Equation Editor
//! 3.0 (`Equation.3`) objects Word shows inline in a run.
//!
//! Word writes each one as a `w:object` run child ([MS-OI29500] 2.1.1641,
//! ECMA-376 Part 1 §17.3.3.19) holding a VML picture shape whose
//! `v:imagedata` names the preview image part, and an `o:OLEObject`
//! (ECMA-376 Part 4 §19.2.2.17) whose `r:id` names the embedded compound
//! file (`word/embeddings/oleObjectN.bin`), the object's ProgID and the
//! shape it draws in:
//!
//! ```xml
//! <w:r><w:rPr><w:position w:val="-6"/></w:rPr>
//!   <w:object w:dxaOrig="620" w:dyaOrig="380">
//!     <v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:31pt;height:19pt" o:ole="">
//!       <v:imagedata r:id="rId9" o:title=""/>
//!     </v:shape>
//!     <o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025"
//!       DrawAspect="Content" ObjectID="_1700000001" r:id="rId8"/>
//!   </w:object>
//! </w:r>
//! ```
//!
//! [`equation_spans`] finds them in a paragraph's XML in document order; the
//! reader ([`crate::Paragraph::equation_objects`]) and the edits count them
//! the same way.

use std::ops::Range;

use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::error::{Error, Result};

/// The deepest element nesting and the most elements a scan accepts.
const MAX_SCAN_DEPTH: usize = 128;
const MAX_SCAN_NODES: usize = 1_000_000;

/// Whether `prog_id` names an equation object: `Equation.3` (Equation
/// Editor 3.0) or an `Equation.DSMT*` (MathType) class, ignoring case as
/// COM does.
///
/// ```
/// use litchi_docx::object::is_equation_prog_id;
///
/// assert!(is_equation_prog_id("Equation.DSMT4"));
/// assert!(is_equation_prog_id("equation.3"));
/// assert!(!is_equation_prog_id("Excel.Sheet.12"));
/// ```
#[must_use]
pub fn is_equation_prog_id(prog_id: &str) -> bool {
    prog_id.eq_ignore_ascii_case("Equation.3")
        || prog_id
            .get(..13)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("Equation.DSMT"))
}

/// An OLE equation object of a paragraph, as its markup describes it.
///
/// Its embedded compound file is the internal part the story part relates
/// through [`Self::object_relationship_id`]; its `Equation Native` stream
/// holds the equation as MTEF.
#[derive(Debug, Clone, PartialEq)]
pub struct EquationObject {
    prog_id: String,
    object_type: Option<String>,
    object_relationship_id: Option<String>,
    preview_relationship_id: Option<String>,
    shape_id: Option<String>,
    size_pt: Option<(f64, f64)>,
}

impl EquationObject {
    /// The object's ProgID: `Equation.DSMT4` (MathType) or `Equation.3`
    /// (Equation Editor 3.0).
    #[must_use]
    pub fn prog_id(&self) -> &str {
        &self.prog_id
    }

    /// The embedding type (`o:OLEObject/@Type`): `Embed`, or `Link` for an
    /// object whose data lives in another file.
    #[must_use]
    pub fn object_type(&self) -> Option<&str> {
        self.object_type.as_deref()
    }

    /// The relationship of the embedded compound file
    /// (`o:OLEObject/@r:id`).
    #[must_use]
    pub fn object_relationship_id(&self) -> Option<&str> {
        self.object_relationship_id.as_deref()
    }

    /// The relationship of the preview image (`v:imagedata/@r:id`).
    #[must_use]
    pub fn preview_relationship_id(&self) -> Option<&str> {
        self.preview_relationship_id.as_deref()
    }

    /// The VML shape identifier (`v:shape/@id`, `_x0000_iNNNN`).
    #[must_use]
    pub fn shape_id(&self) -> Option<&str> {
        self.shape_id.as_deref()
    }

    /// The shown width and height in points, from the shape's `style`.
    #[must_use]
    pub const fn size_pt(&self) -> Option<(f64, f64)> {
        self.size_pt
    }
}

/// One equation object of a paragraph's XML, as byte ranges of it.
#[derive(Debug, Clone, Default)]
pub(crate) struct EquationSpan {
    /// The `w:object` element.
    pub(crate) object: Range<usize>,
    /// The `w:object` start tag.
    pub(crate) object_tag: Range<usize>,
    /// The enclosing `w:r`.
    pub(crate) run: Option<Range<usize>>,
    /// Whether the object is the run's only content besides its
    /// properties.
    pub(crate) sole: bool,
    /// The run's `w:rPr`, when it is the run's direct child.
    pub(crate) run_properties: Option<Range<usize>>,
    /// The `w:position` of that `w:rPr`.
    pub(crate) position: Option<Range<usize>>,
    /// A `v:shapetype` child of the object defining `_x0000_t75`.
    pub(crate) shapetype: Option<Range<usize>>,
    /// The first `v:shape` start tag.
    pub(crate) shape_tag: Option<Range<usize>>,
    /// The first `v:imagedata` start tag of that shape.
    pub(crate) imagedata_tag: Option<Range<usize>>,
    /// Whether the object sits in markup-compatibility alternate content.
    pub(crate) alternate: bool,
    pub(crate) prog_id: String,
    pub(crate) object_type: Option<String>,
    pub(crate) object_relationship_id: Option<String>,
    pub(crate) preview_relationship_id: Option<String>,
    pub(crate) shape_id: Option<String>,
    pub(crate) style: Option<String>,
}

impl EquationSpan {
    /// The reader's view of the object.
    pub(crate) fn object(&self) -> EquationObject {
        EquationObject {
            prog_id: self.prog_id.clone(),
            object_type: self.object_type.clone(),
            object_relationship_id: self.object_relationship_id.clone(),
            preview_relationship_id: self.preview_relationship_id.clone(),
            shape_id: self.shape_id.clone(),
            size_pt: self.style.as_deref().and_then(style_size_pt),
        }
    }
}

/// The value of an attribute whose qualified name ends with `:local` (or is
/// `local`).
fn attribute(element: &quick_xml::events::BytesStart<'_>, local: &[u8]) -> Option<String> {
    first_wins(element).flatten().find_map(|attribute| {
        let key = attribute.key.as_ref();
        let matches = key == local
            || (key.len() > local.len()
                && key.ends_with(local)
                && key[key.len() - local.len() - 1] == b':');
        matches
            .then(|| String::from_utf8(attribute.value.into_owned()).ok())
            .flatten()
    })
}

/// What an open element is, for the scan.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Open {
    Run,
    RunProperties,
    Object,
    Alternate,
    Other,
}

/// Every equation object of `xml` (a paragraph, or any fragment holding
/// paragraphs), in document order: each `w:object` whose `o:OLEObject`
/// names an equation ProgID.
pub(crate) fn equation_spans(xml: &[u8]) -> Result<Vec<EquationSpan>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let invalid = |message: &str| Error::InvalidFormat(format!("equation object scan: {message}"));
    let mut found = Vec::new();
    // Open elements: what each is, its start and its non-property children.
    let mut stack: Vec<(Open, usize, usize)> = Vec::new();
    // The open run: its stack depth, its `w:rPr` and that `w:rPr`'s
    // `w:position`.
    let mut runs: Vec<(usize, Option<Range<usize>>, Option<Range<usize>>)> = Vec::new();
    // The open object.
    let mut current: Option<(usize, EquationSpan)> = None;
    // Objects waiting for their run to close: index into `found`, run depth.
    let mut awaiting: Vec<(usize, usize)> = Vec::new();
    let mut alternates = 0usize;
    let mut nodes = 0usize;
    loop {
        let start =
            usize::try_from(reader.buffer_position()).map_err(|_| invalid("position overflows"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("paragraph XML: {error}")))?;
        let end =
            usize::try_from(reader.buffer_position()).map_err(|_| invalid("position overflows"))?;
        let (element, empty) = match &event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                let depth = stack.len();
                let (open, opened, children) =
                    stack.pop().ok_or_else(|| invalid("unbalanced XML"))?;
                match open {
                    Open::Object => {
                        if let Some((at, mut span)) = current.take()
                            && at == depth
                        {
                            span.object = opened..end;
                            if is_equation_prog_id(&span.prog_id) {
                                if let Some((run_depth, properties, position)) = runs.last() {
                                    span.run_properties.clone_from(properties);
                                    span.position.clone_from(position);
                                    awaiting.push((found.len(), *run_depth));
                                }
                                found.push(span);
                            }
                        }
                    },
                    Open::Run => {
                        runs.pop();
                        awaiting.retain(|(index, run_depth)| {
                            if *run_depth != depth {
                                return true;
                            }
                            if let Some(span) = found.get_mut(*index) {
                                span.run = Some(opened..end);
                                span.sole = children == 1;
                            }
                            false
                        });
                    },
                    Open::Alternate => alternates = alternates.saturating_sub(1),
                    Open::RunProperties | Open::Other => {},
                }
                continue;
            },
            Event::Eof => {
                if !stack.is_empty() {
                    return Err(invalid("unterminated XML"));
                }
                break;
            },
            _ => continue,
        };
        nodes += 1;
        if nodes > MAX_SCAN_NODES || stack.len() >= MAX_SCAN_DEPTH {
            return Err(invalid("scan limit exceeded"));
        }
        let local = element.local_name();
        let local = local.as_ref();
        let parent = stack.last().map(|(open, _, _)| *open);
        if let Some((_, _, children)) = stack.last_mut()
            && local != b"rPr"
        {
            *children += 1;
        }
        let depth = stack.len() + 1;
        let open = match (local, parent) {
            (b"r", _) if current.is_none() => Open::Run,
            (b"rPr", Some(Open::Run)) => Open::RunProperties,
            (b"position", Some(Open::RunProperties)) => {
                // The run's own properties are the innermost `w:rPr` of the
                // stack; a non-empty `w:position` ends where it closes.
                if let Some(run) = runs.last_mut() {
                    run.2 = Some(start..if empty { end } else { usize::MAX });
                }
                Open::Other
            },
            (b"object", _) if current.is_none() => Open::Object,
            (b"AlternateContent", _) => Open::Alternate,
            _ => Open::Other,
        };
        if let Some((at, span)) = current.as_mut() {
            let in_object = depth == *at + 1;
            match local {
                b"shapetype" if in_object => {
                    if attribute(element, b"id").as_deref() == Some("_x0000_t75") {
                        span.shapetype = Some(start..end);
                        if !empty {
                            // Completed when it closes, below.
                            span.shapetype = Some(start..usize::MAX);
                        }
                    }
                },
                b"shape" if span.shape_tag.is_none() => {
                    span.shape_tag = Some(start..end);
                    span.shape_id = attribute(element, b"id");
                    span.style = attribute(element, b"style");
                },
                b"imagedata" if span.imagedata_tag.is_none() => {
                    span.imagedata_tag = Some(start..end);
                    span.preview_relationship_id =
                        attribute(element, b"id").or_else(|| attribute(element, b"embed"));
                },
                b"OLEObject" if span.prog_id.is_empty() => {
                    span.prog_id = attribute(element, b"ProgID").unwrap_or_default();
                    span.object_type = attribute(element, b"Type");
                    span.object_relationship_id = first_wins(element).flatten().find_map(|a| {
                        let key = a.key.as_ref();
                        (key.ends_with(b":id"))
                            .then(|| String::from_utf8(a.value.into_owned()).ok())
                            .flatten()
                    });
                },
                _ => {},
            }
        }
        if open == Open::Object {
            let span = EquationSpan {
                object_tag: start..end,
                alternate: alternates > 0,
                ..EquationSpan::default()
            };
            if empty {
                // An empty object holds nothing to show.
                continue;
            }
            current = Some((depth, span));
        }
        if empty {
            continue;
        }
        match open {
            Open::Run => runs.push((depth, None, None)),
            Open::Alternate => alternates += 1,
            Open::RunProperties => {
                if let Some(run) = runs.last_mut() {
                    run.1 = Some(start..usize::MAX);
                }
            },
            Open::Object | Open::Other => {},
        }
        stack.push((open, start, 0));
    }
    complete_ranges(xml, &mut found)?;
    Ok(found)
}

/// Fill the ends of the `w:rPr` and `v:shapetype` ranges the scan opened.
fn complete_ranges(xml: &[u8], found: &mut [EquationSpan]) -> Result<()> {
    for span in found {
        if let Some(range) = &mut span.run_properties
            && range.end == usize::MAX
        {
            range.end = element_end(xml, range.start)?;
        }
        if let Some(range) = &mut span.position
            && range.end == usize::MAX
        {
            range.end = element_end(xml, range.start)?;
        }
        if let Some(range) = &mut span.shapetype
            && range.end == usize::MAX
        {
            range.end = element_end(xml, range.start)?;
        }
    }
    Ok(())
}

/// The end of the element starting at `start`.
pub(crate) fn element_end(xml: &[u8], start: usize) -> Result<usize> {
    let tail = xml
        .get(start..)
        .ok_or_else(|| Error::InvalidFormat("element range outside its XML".into()))?;
    let mut reader = Reader::from_reader(tail);
    let mut depth = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("paragraph XML: {error}")))?;
        match event {
            Event::Start(_) => depth += 1,
            Event::Empty(_) if depth == 0 => break,
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("unbalanced XML".into()))?;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => return Err(Error::InvalidFormat("unterminated element".into())),
            _ => {},
        }
    }
    let read = usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("position overflows".into()))?;
    Ok(start + read)
}

/// The `width` and `height` of a VML `style`, in points.
pub(crate) fn style_size_pt(style: &str) -> Option<(f64, f64)> {
    let mut width = None;
    let mut height = None;
    for declaration in style.split(';') {
        let Some((name, value)) = declaration.split_once(':') else {
            continue;
        };
        let value = value.trim();
        let unit_start = value
            .find(|c: char| !(c.is_ascii_digit() || matches!(c, '.' | '-' | '+')))
            .unwrap_or(value.len());
        let Ok(number) = value[..unit_start].parse::<f64>() else {
            continue;
        };
        let scale = match value[unit_start..].trim() {
            "" | "pt" => 1.0,
            "in" => 72.0,
            "cm" => 72.0 / 2.54,
            "mm" => 72.0 / 25.4,
            "px" => 0.75,
            _ => continue,
        };
        let points = number * scale;
        if !points.is_finite() {
            continue;
        }
        match name.trim() {
            "width" => width = Some(points),
            "height" => height = Some(points),
            _ => {},
        }
    }
    Some((width?, height?))
}

#[cfg(test)]
mod tests {
    use super::*;

    const PARAGRAPH: &str = concat!(
        r#"<w:p xmlns:w="w" xmlns:v="v" xmlns:o="o" xmlns:r="r">"#,
        r#"<w:r><w:t>x</w:t></w:r>"#,
        r#"<w:r><w:rPr><w:position w:val="-6"/></w:rPr><w:object w:dxaOrig="620" w:dyaOrig="380">"#,
        r#"<v:shapetype id="_x0000_t75" coordsize="21600,21600"><v:stroke joinstyle="miter"/></v:shapetype>"#,
        r##"<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:31pt;height:19pt" o:ole="">"##,
        r#"<v:imagedata r:id="rId9" o:title=""/></v:shape>"#,
        r#"<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1" r:id="rId8"/>"#,
        r#"</w:object></w:r>"#,
        r#"<w:r><w:t>y</w:t><w:object><v:shape id="_x0000_i1026" style="width:1in;height:10pt"><v:imagedata r:id="rId3"/></v:shape>"#,
        r#"<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rId4"/></w:object></w:r>"#,
        r#"<w:r><w:t>z</w:t><w:object><v:shape id="_x0000_i1027" style="width:2pt;height:3pt"><v:imagedata r:id="rId5"/></v:shape>"#,
        r#"<o:OLEObject Type="Embed" ProgID="Equation.3" r:id="rId6"/></w:object></w:r>"#,
        "</w:p>"
    );

    #[test]
    fn equation_objects_are_found_with_their_parts_and_runs() {
        let xml = PARAGRAPH.as_bytes();
        let spans = equation_spans(xml).unwrap();
        assert_eq!(spans.len(), 2);
        let text = |range: &Range<usize>| std::str::from_utf8(&xml[range.clone()]).unwrap();
        let first = &spans[0];
        assert_eq!(first.prog_id, "Equation.DSMT4");
        assert_eq!(first.object_relationship_id.as_deref(), Some("rId8"));
        assert_eq!(first.preview_relationship_id.as_deref(), Some("rId9"));
        assert_eq!(first.shape_id.as_deref(), Some("_x0000_i1025"));
        assert!(first.sole);
        assert_eq!(
            text(first.position.as_ref().unwrap()),
            r#"<w:position w:val="-6"/>"#
        );
        assert_eq!(
            text(first.run_properties.as_ref().unwrap()),
            r#"<w:rPr><w:position w:val="-6"/></w:rPr>"#
        );
        assert!(text(first.shapetype.as_ref().unwrap()).ends_with("</v:shapetype>"));
        assert!(text(&first.object).starts_with("<w:object "));
        assert!(text(&first.object).ends_with("</w:object>"));
        assert!(text(first.run.as_ref().unwrap()).starts_with("<w:r><w:rPr>"));
        assert_eq!(first.object().size_pt(), Some((31.0, 19.0)));
        let second = &spans[1];
        assert_eq!(second.prog_id, "Equation.3");
        assert!(!second.sole);
        assert!(second.run_properties.is_none());
        assert!(second.shapetype.is_none());
    }

    #[test]
    fn prog_ids_and_styles_are_read() {
        assert!(is_equation_prog_id("Equation.DSMT6"));
        assert!(!is_equation_prog_id("Equation"));
        assert_eq!(style_size_pt("width:1in;height:96px"), Some((72.0, 72.0)));
        assert_eq!(style_size_pt("width:1in"), None);
    }

    #[test]
    fn malformed_xml_is_an_error() {
        assert!(equation_spans(b"<w:p><w:r><w:object>").is_err());
    }
}
