//! Independent XML structure gate for the 0437 fresh ODP buffered corpus.
//!
//! Integration assumptions (kept deliberately explicit): this file is merged
//! beside the fixture generator and calls the sibling
//! `odp_buffered_slide_count`,
//! `odp_buffered_title`, and `odp_buffered_body` functions.  It does not
//! regenerate or derive expected text from `content_xml`.
//!
//! This gate is for the fixed, no-transition, title/body fresh corpus.  It
//! rejects optional pages, shapes, notes, animations, declarations, controls,
//! and foreign active XML.  Package member/manifest checks belong to the
//! surrounding archive oracle.

use super::{SemanticShape, odp_buffered_body, odp_buffered_slide_count, odp_buffered_title};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::error::Error;
use std::io;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const FO_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const XLINK_NS: &[u8] = b"http://www.w3.org/1999/xlink";
const DC_NS: &[u8] = b"http://purl.org/dc/elements/1.1/";
const META_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const NUMBER_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const ANIM_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
const SMIL_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const CHART_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const DR3D_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const MATH_NS: &[u8] = b"http://www.w3.org/1998/Math/MathML";
const FORM_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const SCRIPT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const OOO_NS: &[u8] = b"http://openoffice.org/2004/office";

// These are the namespace declarations emitted by Builder::generate_content_xml
// for a fresh no-animation presentation.  Prefix order is not significant;
// presence, value, and absence of extra declarations are significant.
const ROOT_NAMESPACE_BINDINGS: &[(&[u8], &[u8])] = &[
    (b"xmlns:office", OFFICE_NS),
    (b"xmlns:style", STYLE_NS),
    (b"xmlns:text", TEXT_NS),
    (
        b"xmlns:table",
        b"urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    ),
    (b"xmlns:draw", DRAW_NS),
    (b"xmlns:fo", FO_NS),
    (b"xmlns:xlink", XLINK_NS),
    (b"xmlns:dc", DC_NS),
    (b"xmlns:meta", META_NS),
    (b"xmlns:number", NUMBER_NS),
    (b"xmlns:presentation", PRESENTATION_NS),
    (b"xmlns:anim", ANIM_NS),
    (b"xmlns:smil", SMIL_NS),
    (b"xmlns:svg", SVG_NS),
    (b"xmlns:chart", CHART_NS),
    (b"xmlns:dr3d", DR3D_NS),
    (b"xmlns:math", MATH_NS),
    (b"xmlns:form", FORM_NS),
    (b"xmlns:script", SCRIPT_NS),
    (b"xmlns:ooo", OOO_NS),
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Element {
    Document,
    Scripts,
    FontFaceDecls,
    AutomaticStyles,
    Style,
    DrawingPageProperties,
    Body,
    Presentation,
    Page,
    Frame,
    TextBox,
    Paragraph,
    TitleFrame,
    BodyFrame,
    TitleTextBox,
    BodyTextBox,
    TitleParagraph,
    BodyParagraph,
}

#[derive(Clone, Copy)]
struct ExpectedAttribute {
    namespace: &'static [u8],
    local: &'static [u8],
    value: &'static str,
}

fn invalid(message: impl Into<String>) -> Box<dyn Error> {
    io::Error::new(io::ErrorKind::InvalidData, message.into()).into()
}

fn element(namespace: &ResolveResult<'_>, local: &[u8]) -> Option<Element> {
    let ResolveResult::Bound(Namespace(uri)) = namespace else {
        return None;
    };
    let kind = |expected: &[u8]| *uri == expected;
    Some(
        match (
            kind(OFFICE_NS),
            kind(STYLE_NS),
            kind(TEXT_NS),
            kind(DRAW_NS),
            kind(PRESENTATION_NS),
            kind(SVG_NS),
            local,
        ) {
            (true, _, _, _, _, _, b"document-content") => Element::Document,
            (true, _, _, _, _, _, b"scripts") => Element::Scripts,
            (true, _, _, _, _, _, b"font-face-decls") => Element::FontFaceDecls,
            (true, _, _, _, _, _, b"automatic-styles") => Element::AutomaticStyles,
            (_, true, _, _, _, _, b"style") => Element::Style,
            (_, true, _, _, _, _, b"drawing-page-properties") => Element::DrawingPageProperties,
            (true, _, _, _, _, _, b"body") => Element::Body,
            (true, _, _, _, _, _, b"presentation") => Element::Presentation,
            (_, _, _, true, _, _, b"page") => Element::Page,
            (_, _, _, true, _, _, b"frame") => Element::Frame,
            (_, _, _, true, _, _, b"text-box") => Element::TextBox,
            (_, _, true, _, _, _, b"p") => Element::Paragraph,
            _ => return None,
        },
    )
}

fn classify_start(parent: Option<Element>, raw: Element, page_stage: u8) -> Option<Element> {
    match (parent, raw) {
        (Some(Element::Page), Element::Frame) => match page_stage {
            0 => Some(Element::TitleFrame),
            1 => Some(Element::BodyFrame),
            _ => None,
        },
        (Some(Element::TitleFrame), Element::TextBox) => Some(Element::TitleTextBox),
        (Some(Element::BodyFrame), Element::TextBox) => Some(Element::BodyTextBox),
        (Some(Element::TitleTextBox), Element::Paragraph) => Some(Element::TitleParagraph),
        (Some(Element::BodyTextBox), Element::Paragraph) => Some(Element::BodyParagraph),
        _ => Some(raw),
    }
}

fn classify_end(top: Option<Element>, raw: Element) -> Option<Element> {
    match (top, raw) {
        (Some(Element::TitleFrame | Element::BodyFrame), Element::Frame) => top,
        (Some(Element::TitleTextBox | Element::BodyTextBox), Element::TextBox) => top,
        (Some(Element::TitleParagraph | Element::BodyParagraph), Element::Paragraph) => top,
        _ => Some(raw),
    }
}

fn same_namespace(actual: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(actual, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn decode_attribute(
    reader: &NsReader<&[u8]>,
    attribute: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String, Box<dyn Error>> {
    attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| invalid(format!("invalid ODP attribute value: {error}")))
}

fn require_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[ExpectedAttribute],
    label: &str,
) -> Result<(), Box<dyn Error>> {
    let mut seen = vec![false; expected.len()];
    for raw in element.attributes().with_checks(true) {
        let attribute =
            raw.map_err(|error| invalid(format!("{label}: invalid attribute: {error}")))?;
        if attribute.key.as_namespace_binding().is_some() {
            return Err(invalid(format!(
                "{label}: unexpected namespace declaration"
            )));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let Some(index) = expected.iter().position(|candidate| {
            same_namespace(&namespace, candidate.namespace) && local.as_ref() == candidate.local
        }) else {
            return Err(invalid(format!(
                "{label}: unexpected attribute {:?}",
                attribute.key.as_ref()
            )));
        };
        drop(namespace);
        if seen[index] {
            return Err(invalid(format!("{label}: duplicate attribute")));
        }
        seen[index] = true;
        let value = decode_attribute(reader, &attribute)?;
        if value != expected[index].value {
            return Err(invalid(format!(
                "{label}: attribute {:?} has value {:?}, expected {:?}",
                attribute.key.as_ref(),
                value,
                expected[index].value
            )));
        }
    }
    if let Some(index) = seen.iter().position(|value| !value) {
        return Err(invalid(format!(
            "{label}: missing required attribute {:?}",
            expected[index].local
        )));
    }
    Ok(())
}

fn require_no_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    label: &str,
) -> Result<(), Box<dyn Error>> {
    require_attributes(reader, element, &[], label)
}

fn require_default_page_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    page_name: &str,
) -> Result<(), Box<dyn Error>> {
    let mut seen_name = false;
    let mut seen_style = false;
    let mut seen_master = false;
    for raw in element.attributes().with_checks(true) {
        let attribute =
            raw.map_err(|error| invalid(format!("draw:page: invalid attribute: {error}")))?;
        if attribute.key.as_namespace_binding().is_some() {
            return Err(invalid("draw:page: unexpected namespace declaration"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !same_namespace(&namespace, DRAW_NS) {
            return Err(invalid(format!(
                "draw:page: unexpected attribute {:?}",
                attribute.key.as_ref()
            )));
        }
        drop(namespace);
        let value = decode_attribute(reader, &attribute)?;
        match local.as_ref() {
            b"name" if !seen_name => {
                seen_name = true;
                if value != page_name {
                    return Err(invalid(format!(
                        "draw:page name {value:?} differs from expected {page_name:?}"
                    )));
                }
            },
            b"style-name" if !seen_style => {
                seen_style = true;
                if value != "dp1" {
                    return Err(invalid("draw:page style-name must be dp1"));
                }
            },
            b"master-page-name" if !seen_master => {
                seen_master = true;
                if value != "Default" {
                    return Err(invalid("draw:page master-page-name must be Default"));
                }
            },
            _ => {
                return Err(invalid(format!(
                    "draw:page: unexpected or duplicate attribute {:?}",
                    attribute.key.as_ref()
                )));
            },
        }
    }
    if !seen_name || !seen_style || !seen_master {
        return Err(invalid(
            "draw:page: default metadata attributes are incomplete",
        ));
    }
    Ok(())
}

fn require_root_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<(), Box<dyn Error>> {
    let mut seen_namespaces = vec![false; ROOT_NAMESPACE_BINDINGS.len()];
    let mut seen_version = false;
    for raw in element.attributes().with_checks(true) {
        let attribute =
            raw.map_err(|error| invalid(format!("document root: invalid attribute: {error}")))?;
        if attribute.key.as_namespace_binding().is_some() {
            let raw_name = attribute.key.as_ref();
            let Some(index) = ROOT_NAMESPACE_BINDINGS
                .iter()
                .position(|(name, _)| *name == raw_name)
            else {
                return Err(invalid(format!(
                    "document root: unexpected namespace declaration {raw_name:?}"
                )));
            };
            if seen_namespaces[index] {
                return Err(invalid("document root: duplicate namespace declaration"));
            }
            seen_namespaces[index] = true;
            let value = decode_attribute(reader, &attribute)?;
            if value.as_bytes() != ROOT_NAMESPACE_BINDINGS[index].1 {
                return Err(invalid(format!(
                    "document root: namespace {:?} differs from the fixed ODP binding",
                    raw_name
                )));
            }
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !same_namespace(&namespace, OFFICE_NS) || local.as_ref() != b"version" {
            return Err(invalid(format!(
                "document root: unexpected attribute {:?}",
                attribute.key.as_ref()
            )));
        }
        drop(namespace);
        if seen_version {
            return Err(invalid("document root: duplicate office:version"));
        }
        seen_version = true;
        if decode_attribute(reader, &attribute)? != "1.3" {
            return Err(invalid("document root: office:version must be 1.3"));
        }
    }
    if !seen_version || seen_namespaces.iter().any(|seen| !seen) {
        return Err(invalid(
            "document root: fixed namespace/version prelude is incomplete",
        ));
    }
    Ok(())
}

fn decode_text(text: &quick_xml::events::BytesText<'_>) -> Result<String, Box<dyn Error>> {
    let encoded = text
        .xml_content(XmlVersion::Explicit1_0)
        .map_err(|error| invalid(format!("invalid ODP text: {error}")))?;
    quick_xml::escape::unescape(&encoded)
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| invalid(format!("invalid ODP text entity: {error}")))
}

fn decode_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String, Box<dyn Error>> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| invalid(format!("invalid ODP character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| invalid(format!("invalid ODP entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_owned()),
        "lt" => Ok("<".to_owned()),
        "gt" => Ok(">".to_owned()),
        "quot" => Ok("\"".to_owned()),
        "apos" => Ok("'".to_owned()),
        _ => Err(invalid(format!(
            "unsupported ODP entity reference &{name};"
        ))),
    }
}

fn expected_frame_attributes(kind: Element) -> &'static [ExpectedAttribute] {
    match kind {
        Element::TitleFrame => &[
            ExpectedAttribute {
                namespace: DRAW_NS,
                local: b"style-name",
                value: "gr1",
            },
            ExpectedAttribute {
                namespace: DRAW_NS,
                local: b"text-style-name",
                value: "P1",
            },
            ExpectedAttribute {
                namespace: DRAW_NS,
                local: b"layer",
                value: "layout",
            },
            ExpectedAttribute {
                namespace: PRESENTATION_NS,
                local: b"class",
                value: "title",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"width",
                value: "25.199cm",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"height",
                value: "3.506cm",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"x",
                value: "1.4cm",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"y",
                value: "0.962cm",
            },
        ],
        Element::BodyFrame => &[
            ExpectedAttribute {
                namespace: DRAW_NS,
                local: b"style-name",
                value: "gr2",
            },
            ExpectedAttribute {
                namespace: DRAW_NS,
                local: b"text-style-name",
                value: "P2",
            },
            ExpectedAttribute {
                namespace: DRAW_NS,
                local: b"layer",
                value: "layout",
            },
            ExpectedAttribute {
                namespace: PRESENTATION_NS,
                local: b"class",
                value: "object",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"width",
                value: "25.199cm",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"height",
                value: "10cm",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"x",
                value: "1.4cm",
            },
            ExpectedAttribute {
                namespace: SVG_NS,
                local: b"y",
                value: "5.0cm",
            },
        ],
        _ => &[],
    }
}

/// Validate the exact no-transition title/body `content.xml` grammar emitted
/// by the fresh buffered ODP corpus producer.
///
/// Expected count and text come from the independently specified fixture.
pub(super) fn verify_odp_buffered_content_structure(
    content_xml: &[u8],
    shape: SemanticShape,
) -> Result<(), Box<dyn Error>> {
    let xml = std::str::from_utf8(content_xml)
        .map_err(|error| invalid(format!("ODP content.xml is not UTF-8: {error}")))?;
    let mut reader = NsReader::from_str(xml);
    let expected_count = odp_buffered_slide_count(shape);
    let mut buffer = Vec::new();
    let mut stack = Vec::<Element>::new();
    let mut declaration_seen = false;
    let mut root_closed = false;
    let mut document_stage = 0u8;
    let mut style_count = 0usize;
    let mut style_property_seen = false;
    let mut page_index = 0usize;
    let mut presentation_seen = false;
    let mut page_stage = 0u8;
    let mut frame_box_seen = false;
    let mut textbox_paragraph_seen = false;
    let mut paragraph: Option<(Element, String)> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid ODP content.xml: {error}")))?;
        match event {
            Event::Decl(declaration) => {
                if declaration.version()?.as_ref() != b"1.0"
                    || declaration.encoding().transpose()?.as_deref() != Some(b"UTF-8".as_slice())
                    || declaration.standalone().is_some()
                {
                    return Err(invalid(
                        "ODP declaration differs from the UTF-8 XML 1.0 contract",
                    ));
                }
                if declaration_seen || !stack.is_empty() || root_closed {
                    return Err(invalid("ODP XML declaration is duplicated or misplaced"));
                }
                declaration_seen = true;
            },
            Event::Start(start) => {
                let raw_kind =
                    element(&namespace, start.local_name().as_ref()).ok_or_else(|| {
                        invalid(format!(
                            "unexpected ODP element {:?}",
                            start.name().as_ref()
                        ))
                    })?;
                let parent = stack.last().copied();
                let kind = classify_start(parent, raw_kind, page_stage).ok_or_else(|| {
                    invalid(format!(
                        "unexpected ODP parent for {:?}",
                        start.name().as_ref()
                    ))
                })?;
                // `read_resolved_event_into` returns a namespace view borrowed
                // from the reader.  End that borrow before attribute helpers
                // call reader.resolver()/reader.decoder().
                match kind {
                    Element::Document => {
                        if !stack.is_empty() || root_closed {
                            return Err(invalid("ODP document root is duplicated or nested"));
                        }
                        require_root_attributes(&reader, &start)?;
                    },
                    Element::AutomaticStyles => {
                        if parent != Some(Element::Document) || document_stage != 2 {
                            return Err(invalid("ODP automatic-styles is out of order"));
                        }
                        require_no_attributes(&reader, &start, "office:automatic-styles")?;
                    },
                    Element::Style => {
                        if parent != Some(Element::AutomaticStyles) || style_count != 0 {
                            return Err(invalid("ODP automatic-styles has an unexpected style"));
                        }
                        require_attributes(
                            &reader,
                            &start,
                            &[
                                ExpectedAttribute {
                                    namespace: STYLE_NS,
                                    local: b"name",
                                    value: "dp1",
                                },
                                ExpectedAttribute {
                                    namespace: STYLE_NS,
                                    local: b"family",
                                    value: "drawing-page",
                                },
                            ],
                            "style:style",
                        )?;
                        style_property_seen = false;
                    },
                    Element::Body => {
                        if parent != Some(Element::Document) || document_stage != 3 {
                            return Err(invalid("ODP office:body is out of order"));
                        }
                        require_no_attributes(&reader, &start, "office:body")?;
                    },
                    Element::Presentation => {
                        if parent != Some(Element::Body) || presentation_seen {
                            return Err(invalid("ODP office:presentation has the wrong parent"));
                        }
                        require_no_attributes(&reader, &start, "office:presentation")?;
                        presentation_seen = true;
                    },
                    Element::Page => {
                        if parent != Some(Element::Presentation) || page_index >= expected_count {
                            return Err(invalid("ODP page count/order differs from the fixture"));
                        }
                        let page_name = format!("page{}", page_index + 1);
                        require_default_page_attributes(&reader, &start, &page_name)?;
                        page_stage = 0;
                    },
                    Element::TitleFrame | Element::BodyFrame => {
                        if parent != Some(Element::Page) {
                            return Err(invalid("ODP frame has the wrong parent"));
                        }
                        let expected_stage = if kind == Element::TitleFrame { 0 } else { 1 };
                        if page_stage != expected_stage {
                            return Err(invalid(
                                "ODP title/body frame order differs from the fixture",
                            ));
                        }
                        require_attributes(
                            &reader,
                            &start,
                            expected_frame_attributes(kind),
                            if kind == Element::TitleFrame {
                                "title frame"
                            } else {
                                "body frame"
                            },
                        )?;
                        page_stage += 1;
                        frame_box_seen = false;
                    },
                    Element::TitleTextBox | Element::BodyTextBox => {
                        let expected_parent = if kind == Element::TitleTextBox {
                            Element::TitleFrame
                        } else {
                            Element::BodyFrame
                        };
                        if parent != Some(expected_parent) || frame_box_seen {
                            return Err(invalid("ODP frame must contain exactly one text-box"));
                        }
                        require_no_attributes(&reader, &start, "draw:text-box")?;
                        frame_box_seen = true;
                        textbox_paragraph_seen = false;
                    },
                    Element::TitleParagraph | Element::BodyParagraph => {
                        let expected_parent = if kind == Element::TitleParagraph {
                            Element::TitleTextBox
                        } else {
                            Element::BodyTextBox
                        };
                        if parent != Some(expected_parent) || textbox_paragraph_seen {
                            return Err(invalid("ODP text-box must contain exactly one paragraph"));
                        }
                        let style = if kind == Element::TitleParagraph {
                            "P1"
                        } else {
                            "P2"
                        };
                        require_attributes(
                            &reader,
                            &start,
                            &[ExpectedAttribute {
                                namespace: TEXT_NS,
                                local: b"style-name",
                                value: style,
                            }],
                            "text:p",
                        )?;
                        textbox_paragraph_seen = true;
                        paragraph = Some((kind, String::new()));
                    },
                    Element::Scripts | Element::FontFaceDecls | Element::DrawingPageProperties => {
                        return Err(invalid(
                            "ODP fixed empty element was emitted as a start element",
                        ));
                    },
                    Element::Frame | Element::TextBox | Element::Paragraph => {
                        return Err(invalid("ODP drawing/text element has no valid parent"));
                    },
                }
                stack.push(kind);
            },
            Event::Empty(empty) => {
                let raw_kind =
                    element(&namespace, empty.local_name().as_ref()).ok_or_else(|| {
                        invalid(format!(
                            "unexpected ODP empty element {:?}",
                            empty.name().as_ref()
                        ))
                    })?;
                let parent = stack.last().copied();
                let kind = classify_start(parent, raw_kind, page_stage).ok_or_else(|| {
                    invalid(format!(
                        "unexpected ODP empty parent for {:?}",
                        empty.name().as_ref()
                    ))
                })?;
                match kind {
                    Element::Scripts => {
                        if parent != Some(Element::Document) || document_stage != 0 {
                            return Err(invalid("ODP office:scripts is out of order"));
                        }
                        require_no_attributes(&reader, &empty, "office:scripts")?;
                        document_stage = 1;
                    },
                    Element::FontFaceDecls => {
                        if parent != Some(Element::Document) || document_stage != 1 {
                            return Err(invalid("ODP office:font-face-decls is out of order"));
                        }
                        require_no_attributes(&reader, &empty, "office:font-face-decls")?;
                        document_stage = 2;
                    },
                    Element::DrawingPageProperties => {
                        if parent != Some(Element::Style) || style_property_seen {
                            return Err(invalid(
                                "ODP dp1 style must contain one drawing-page-properties",
                            ));
                        }
                        require_no_attributes(&reader, &empty, "style:drawing-page-properties")?;
                        style_property_seen = true;
                    },
                    _ => return Err(invalid("unexpected non-empty ODP element")),
                }
            },
            Event::Text(text) => {
                let Some((kind, value)) = paragraph.as_mut() else {
                    return Err(invalid("ODP content has text outside a text:p"));
                };
                if stack.last().copied() != Some(*kind) {
                    return Err(invalid("ODP paragraph text state is inconsistent"));
                }
                value.push_str(&decode_text(&text)?);
            },
            Event::GeneralRef(reference) => {
                let Some((kind, value)) = paragraph.as_mut() else {
                    return Err(invalid("ODP entity reference occurs outside a text:p"));
                };
                if stack.last().copied() != Some(*kind) {
                    return Err(invalid("ODP paragraph reference state is inconsistent"));
                }
                value.push_str(&decode_reference(&reference)?);
            },
            Event::End(end) => {
                let raw_kind = element(&namespace, end.local_name().as_ref()).ok_or_else(|| {
                    invalid(format!(
                        "unexpected ODP end element {:?}",
                        end.name().as_ref()
                    ))
                })?;
                let kind = classify_end(stack.last().copied(), raw_kind)
                    .ok_or_else(|| invalid("ODP end element has the wrong context"))?;
                if stack.last().copied() != Some(kind) {
                    return Err(invalid("ODP element nesting/order is invalid"));
                }
                match kind {
                    Element::TitleParagraph | Element::BodyParagraph => {
                        let (paragraph_kind, actual) = paragraph
                            .take()
                            .ok_or_else(|| invalid("ODP paragraph closed without text state"))?;
                        let expected = if paragraph_kind == Element::TitleParagraph {
                            odp_buffered_title(page_index)
                        } else {
                            odp_buffered_body(page_index)
                        };
                        if actual != expected {
                            return Err(invalid(format!(
                                "ODP slide {page_index} {:?} text differs from fixture",
                                paragraph_kind
                            )));
                        }
                    },
                    Element::TitleTextBox | Element::BodyTextBox => {
                        if paragraph.is_some() || !textbox_paragraph_seen {
                            return Err(invalid("ODP text-box paragraph contract is incomplete"));
                        }
                    },
                    Element::TitleFrame | Element::BodyFrame => {
                        if paragraph.is_some() || !frame_box_seen {
                            return Err(invalid("ODP frame text-box contract is incomplete"));
                        }
                    },
                    Element::Page => {
                        if page_stage != 2 || paragraph.is_some() {
                            return Err(invalid(
                                "ODP page does not contain exactly title then body",
                            ));
                        }
                        page_index = page_index
                            .checked_add(1)
                            .ok_or_else(|| invalid("ODP page count overflow"))?;
                    },
                    Element::Presentation => {
                        if page_index != expected_count {
                            return Err(invalid(
                                "ODP presentation slide count differs from fixture",
                            ));
                        }
                    },
                    Element::Style => {
                        if !style_property_seen {
                            return Err(invalid("ODP dp1 style lacks drawing-page-properties"));
                        }
                        style_count = style_count
                            .checked_add(1)
                            .ok_or_else(|| invalid("ODP style count overflow"))?;
                    },
                    Element::AutomaticStyles => {
                        if style_count != 1 {
                            return Err(invalid("ODP automatic-styles must contain exactly dp1"));
                        }
                        document_stage = 3;
                    },
                    Element::Body => {
                        if !presentation_seen {
                            return Err(invalid("ODP body lacks its presentation"));
                        }
                        document_stage = 4;
                    },
                    Element::Document => {
                        if document_stage != 4 || page_index != expected_count {
                            return Err(invalid("ODP document prelude/body is incomplete"));
                        }
                        root_closed = true;
                    },
                    Element::Scripts | Element::FontFaceDecls | Element::DrawingPageProperties => {
                        return Err(invalid("unexpected end for fixed empty ODP element"));
                    },
                    Element::Frame | Element::TextBox | Element::Paragraph => {
                        return Err(invalid("ODP generic drawing/text end has no valid context"));
                    },
                }
                stack.pop();
            },
            Event::Eof => {
                if !stack.is_empty() || paragraph.is_some() || !root_closed {
                    return Err(invalid(
                        "ODP content.xml ended before the fixed structure closed",
                    ));
                }
                break;
            },
            Event::CData(_) | Event::Comment(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "ODP content contains an unsupported active/non-fixed XML event",
                ));
            },
        }
        buffer.clear();
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use soapberry_zip::office::ArchiveReader;

    #[test]
    fn fixed_content_gate_rejects_structural_and_semantic_mutations() {
        let archive = super::super::odp_buffered_bytes(SemanticShape::Tiny).unwrap();
        let content = ArchiveReader::new(&archive)
            .unwrap()
            .read("content.xml")
            .unwrap();
        verify_odp_buffered_content_structure(&content, SemanticShape::Tiny).unwrap();
        let xml = String::from_utf8(content).unwrap();
        for (original, replacement) in [
            ("draw:name=\"page1\"", "draw:name=\"page2\""),
            (
                "draw:master-page-name=\"Default\"",
                "draw:master-page-name=\"Other\"",
            ),
            ("svg:width=\"25.199cm\"", "svg:width=\"25.198cm\""),
            (
                "presentation:class=\"title\"",
                "presentation:class=\"object\"",
            ),
            ("text:style-name=\"P1\"", "text:style-name=\"P2\""),
            ("style:name=\"dp1\"", "style:name=\"dp2\""),
            (
                "litchi-perf-odp-buffered-title-00000",
                "litchi-perf-odp-buffered-title-00001",
            ),
            ("</draw:page>", "<presentation:notes/></draw:page>"),
            (
                "</draw:text-box>",
                "<text:p text:style-name=\"P1\">extra</text:p></draw:text-box>",
            ),
            (
                "</office:presentation>",
                "</office:presentation><office:presentation></office:presentation>",
            ),
            ("<office:scripts/>", "<office:scripts/><office:scripts/>"),
            ("<office:body>", "<office:body extra=\"value\">"),
            ("<office:body>", "<!--unexpected--><office:body>"),
            ("<office:body>", "<?unexpected value?><office:body>"),
            ("plain slide", "plain<text:tab/>slide"),
            ("plain slide", "<![CDATA[plain slide]]>"),
            ("&amp;", "&unbound;"),
            ("version=\"1.0\"", "version=\"1.1\""),
            ("office:version=\"1.3\"", "office:version=\"1.2\""),
            (
                "xmlns:ooo=\"http://openoffice.org/2004/office\"",
                "xmlns:ooo=\"urn:unexpected\"",
            ),
        ] {
            let mutated = xml.replacen(original, replacement, 1);
            assert_ne!(mutated, xml, "mutation target must exist: {original}");
            assert!(
                verify_odp_buffered_content_structure(mutated.as_bytes(), SemanticShape::Tiny)
                    .is_err(),
                "accepted content mutation: {original} -> {replacement}"
            );
        }
    }
}
