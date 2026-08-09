//! Lossless XML discovery and serialization for shape design elements.

use std::ops::Range;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Opaque, Snapshot};
use super::validation::{
    MAX_EXTENSION_BYTES, MAX_SHAPE_BYTES, MAX_SHAPE_DEPTH, count_node, invalid, push_unknown,
    xml_error,
};
use crate::tag::{Conformance, replace_xml};
use crate::{Error, Result};

const PML_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const PML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const P15: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2015/main";

/// A checked element range in the selected raw shape.
#[derive(Debug, Clone)]
pub(super) struct Element {
    pub(super) span: Range<usize>,
    pub(super) close_start: usize,
    pub(super) empty: bool,
}

#[derive(Debug, Clone)]
pub(super) struct Extension {
    pub(super) element: Element,
    pub(super) design: Option<Element>,
    pub(super) other_content: bool,
}

#[derive(Debug, Clone)]
pub(super) struct ExtensionList {
    pub(super) element: Element,
    pub(super) child_elements: usize,
    pub(super) other_content: bool,
}

/// Validated shape-local ranges needed by an atomic transaction.
#[derive(Debug, Clone)]
pub(super) struct Layout {
    pub(super) conformance: Conformance,
    pub(super) nv_pr: Element,
    pub(super) ext_lst: Option<ExtensionList>,
    pub(super) known: Option<Extension>,
}

/// Parsed design-element source and its raw placement.
#[derive(Debug, Clone)]
pub(super) struct Source {
    pub(super) snapshot: Option<Snapshot>,
    pub(super) layout: Layout,
}

#[derive(Debug, Clone, Copy)]
enum Kind {
    Root,
    NonVisual,
    NvPr,
    ExtList,
    Extension { known: bool },
    DesignElement,
    Other,
}

#[derive(Debug, Clone, Copy)]
enum NamespaceKind {
    Pml(Conformance),
    P15,
    Other,
}

#[derive(Debug)]
struct Frame {
    kind: Kind,
    start: usize,
    other_content: bool,
    design_seen: bool,
}

/// Parse one bounded selected shape into a transaction-ready source.
pub(super) fn read(xml: &[u8]) -> Result<Source> {
    if xml.is_empty() {
        return Err(invalid("designer shape XML is empty"));
    }
    if xml.len() > MAX_SHAPE_BYTES {
        return Err(Error::Limit {
            resource: "designer shape XML bytes",
            limit: MAX_SHAPE_BYTES,
        });
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut frames = Vec::<Frame>::new();
    let mut nodes = 0usize;
    let mut nv_pr = None;
    let mut ext_lst = None;
    let mut known = None;
    let mut unknown_extensions = Vec::new();
    let mut design = None;
    let mut extension_children = 0usize;
    let mut conformance = None;
    let mut root_closed = false;

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let namespace = namespace_kind(namespace);
        let event = event.into_owned();
        let end = position(&reader)?;

        match event {
            Event::Start(element) => {
                count_node(&mut nodes)?;
                if frames.is_empty() {
                    let profile = pml_profile(namespace)
                        .ok_or_else(|| invalid("designer shape root is not PresentationML"))?;
                    if !is_shape_root(element.local_name().as_ref()) {
                        return Err(invalid("designer owner is not a supported shape"));
                    }
                    conformance = Some(profile);
                    frames.push(Frame {
                        kind: Kind::Root,
                        start,
                        other_content: false,
                        design_seen: false,
                    });
                    continue;
                }

                let kind = classify_start(
                    namespace,
                    &element,
                    decoder,
                    frames
                        .last_mut()
                        .ok_or_else(|| invalid("designer shape frame stack became empty"))?,
                    &mut nv_pr,
                    &mut ext_lst,
                    &mut known,
                    &mut conformance,
                )?;
                if frames.len() >= MAX_SHAPE_DEPTH {
                    return Err(Error::Limit {
                        resource: "designer shape XML depth",
                        limit: MAX_SHAPE_DEPTH,
                    });
                }
                frames.push(Frame {
                    kind,
                    start,
                    other_content: false,
                    design_seen: false,
                });
            },
            Event::Empty(element) => {
                count_node(&mut nodes)?;
                if frames.is_empty() {
                    let profile = pml_profile(namespace)
                        .ok_or_else(|| invalid("designer shape root is not PresentationML"))?;
                    if !is_shape_root(element.local_name().as_ref()) {
                        return Err(invalid("designer owner is not a supported shape"));
                    }
                    conformance = Some(profile);
                    frames.push(Frame {
                        kind: Kind::Root,
                        start,
                        other_content: false,
                        design_seen: false,
                    });
                    finish_frame(
                        xml,
                        frames
                            .pop()
                            .ok_or_else(|| invalid("missing empty designer shape root"))?,
                        end,
                        true,
                        &mut nv_pr,
                        &mut ext_lst,
                        &mut known,
                        &mut unknown_extensions,
                        &mut design,
                        &mut extension_children,
                    )?;
                    root_closed = true;
                    continue;
                }

                let kind = classify_start(
                    namespace,
                    &element,
                    decoder,
                    frames
                        .last_mut()
                        .ok_or_else(|| invalid("designer shape frame stack became empty"))?,
                    &mut nv_pr,
                    &mut ext_lst,
                    &mut known,
                    &mut conformance,
                )?;
                finish_frame(
                    xml,
                    Frame {
                        kind,
                        start,
                        other_content: false,
                        design_seen: false,
                    },
                    end,
                    true,
                    &mut nv_pr,
                    &mut ext_lst,
                    &mut known,
                    &mut unknown_extensions,
                    &mut design,
                    &mut extension_children,
                )?;
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("designer shape XML stack underflow"))?;
                finish_frame(
                    xml,
                    frame,
                    end,
                    false,
                    &mut nv_pr,
                    &mut ext_lst,
                    &mut known,
                    &mut unknown_extensions,
                    &mut design,
                    &mut extension_children,
                )?;
                if frames.is_empty() {
                    if end != xml.len() {
                        return Err(invalid("designer shape root does not cover its XML"));
                    }
                    root_closed = true;
                    break;
                }
            },
            Event::Text(text) => {
                let value = text.decode().map_err(xml_error)?;
                if !value.trim().is_empty()
                    && matches!(
                        frames.last().map(|frame| frame.kind),
                        Some(Kind::NvPr | Kind::ExtList | Kind::DesignElement)
                    )
                {
                    return Err(invalid(
                        "designer structural metadata contains non-whitespace text",
                    ));
                }
                if !value.trim().is_empty()
                    && matches!(
                        frames.last().map(|frame| frame.kind),
                        Some(Kind::Extension { .. })
                    )
                    && let Some(frame) = frames.last_mut()
                {
                    frame.other_content = true;
                }
            },
            Event::CData(text) => {
                if matches!(
                    frames.last().map(|frame| frame.kind),
                    Some(Kind::DesignElement)
                ) {
                    return Err(invalid("designElem cannot contain CDATA"));
                }
                if !text.decode().map_err(xml_error)?.trim().is_empty()
                    && matches!(
                        frames.last().map(|frame| frame.kind),
                        Some(Kind::NvPr | Kind::ExtList | Kind::DesignElement)
                    )
                {
                    return Err(invalid("designer structural metadata contains CDATA"));
                }
                if matches!(
                    frames.last().map(|frame| frame.kind),
                    Some(Kind::Extension { .. })
                ) && let Some(frame) = frames.last_mut()
                {
                    frame.other_content = true;
                }
            },
            Event::Comment(_) => {
                if matches!(
                    frames.last().map(|frame| frame.kind),
                    Some(Kind::DesignElement)
                ) {
                    return Err(invalid("designElem cannot contain comments"));
                }
                if matches!(
                    frames.last().map(|frame| frame.kind),
                    Some(Kind::ExtList | Kind::Extension { .. })
                ) && let Some(frame) = frames.last_mut()
                {
                    frame.other_content = true;
                }
            },
            Event::GeneralRef(_)
                if matches!(
                    frames.last().map(|frame| frame.kind),
                    Some(Kind::NvPr | Kind::ExtList | Kind::DesignElement)
                ) =>
            {
                return Err(invalid(
                    "designer structural metadata contains an entity reference",
                ));
            },
            Event::Decl(_) if frames.is_empty() => {},
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) if !frames.is_empty() => {
                return Err(invalid("designer shape contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !root_closed || !frames.is_empty() {
        return Err(invalid("designer shape XML is unterminated"));
    }
    let conformance = conformance.ok_or_else(|| invalid("designer shape profile is missing"))?;
    let nv_pr = nv_pr.ok_or_else(|| invalid("designer shape has no direct p:nvPr"))?;
    let snapshot = known
        .as_ref()
        .and_then(|extension| extension.design.as_ref())
        .map(|element| parse_snapshot(xml, element, &unknown_extensions))
        .transpose()?
        .flatten();

    Ok(Source {
        snapshot,
        layout: Layout {
            conformance,
            nv_pr,
            ext_lst,
            known,
        },
    })
}

fn classify_start(
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    parent: &mut Frame,
    nv_pr: &mut Option<Element>,
    ext_lst: &mut Option<ExtensionList>,
    known: &mut Option<Extension>,
    conformance: &mut Option<Conformance>,
) -> Result<Kind> {
    let local = element.local_name();
    let pml = pml_profile(namespace);
    let kind = match parent.kind {
        Kind::Root => {
            if pml.is_some() && is_non_visual(local.as_ref()) {
                Kind::NonVisual
            } else {
                mark_other(parent);
                Kind::Other
            }
        },
        Kind::NonVisual => {
            if pml.is_some() && local.as_ref() == b"nvPr" {
                Kind::NvPr
            } else {
                mark_other(parent);
                Kind::Other
            }
        },
        Kind::NvPr => {
            if pml.is_some() && local.as_ref() == b"extLst" {
                if ext_lst.is_some() {
                    return Err(invalid("designer shape has duplicate p:extLst"));
                }
                if let Some(profile) = pml {
                    if let Some(existing) = conformance
                        && *existing != profile
                    {
                        return Err(invalid(
                            "designer shape mixes Strict and Transitional PresentationML",
                        ));
                    }
                    *conformance = Some(profile);
                }
                Kind::ExtList
            } else {
                mark_other(parent);
                Kind::Other
            }
        },
        Kind::ExtList => {
            if pml.is_some() && local.as_ref() == b"ext" {
                let uri = unqualified_attribute_value(element, b"uri", decoder)?
                    .ok_or_else(|| invalid("designer p:ext has no uri"))?;
                if uri.is_empty() {
                    return Err(invalid("designer p:ext has an empty uri"));
                }
                let is_known = uri == super::EXTENSION_URI;
                if is_known && known.is_some() {
                    return Err(invalid(
                        "designer shape has duplicate design-element extensions",
                    ));
                }
                Kind::Extension { known: is_known }
            } else {
                mark_other(parent);
                Kind::Other
            }
        },
        Kind::Extension { known: is_known } => {
            if is_known
                && matches!(namespace, NamespaceKind::P15)
                && local.as_ref() == b"designElem"
            {
                if parent.design_seen {
                    return Err(invalid(
                        "designer extension has duplicate designElem elements",
                    ));
                }
                validate_design_attributes(element, decoder)?;
                parent.design_seen = true;
                Kind::DesignElement
            } else {
                mark_other(parent);
                Kind::Other
            }
        },
        Kind::DesignElement => {
            return Err(invalid("designElem cannot contain child elements"));
        },
        Kind::Other => Kind::Other,
    };

    if let Kind::NvPr = kind
        && nv_pr.is_some()
    {
        return Err(invalid("designer shape has duplicate p:nvPr elements"));
    }
    Ok(kind)
}

fn finish_frame(
    xml: &[u8],
    frame: Frame,
    end: usize,
    empty: bool,
    nv_pr: &mut Option<Element>,
    ext_lst: &mut Option<ExtensionList>,
    known: &mut Option<Extension>,
    unknown_extensions: &mut Vec<Opaque>,
    design: &mut Option<Element>,
    extension_children: &mut usize,
) -> Result<()> {
    if end > xml.len() || frame.start >= end {
        return Err(invalid("designer element range is outside its shape"));
    }
    match frame.kind {
        Kind::NvPr => {
            if nv_pr.is_some() {
                return Err(invalid("designer shape has duplicate p:nvPr elements"));
            }
            *nv_pr = Some(Element {
                span: frame.start..end,
                close_start: if empty {
                    end
                } else {
                    frame_close_start(end, xml)
                },
                empty,
            });
        },
        Kind::ExtList => {
            if ext_lst.is_some() {
                return Err(invalid("designer shape has duplicate p:extLst elements"));
            }
            *ext_lst = Some(ExtensionList {
                element: Element {
                    span: frame.start..end,
                    close_start: if empty {
                        end
                    } else {
                        frame_close_start(end, xml)
                    },
                    empty,
                },
                child_elements: *extension_children,
                other_content: frame.other_content,
            });
        },
        Kind::Extension { known: is_known } => {
            let close_start = if empty {
                end
            } else {
                frame_close_start(end, xml)
            };
            let element = Element {
                span: frame.start..end,
                close_start,
                empty,
            };
            *extension_children = extension_children.checked_add(1).ok_or(Error::Limit {
                resource: "designer extension count",
                limit: super::validation::MAX_EXTENSION_COUNT,
            })?;
            if is_known {
                if known.is_some() {
                    return Err(invalid(
                        "designer shape has duplicate design-element extensions",
                    ));
                }
                *known = Some(Extension {
                    element,
                    design: design.take(),
                    other_content: frame.other_content,
                });
            } else {
                let raw = xml
                    .get(frame.start..end)
                    .ok_or_else(|| invalid("designer opaque extension range is invalid"))?
                    .to_vec();
                push_unknown(unknown_extensions, raw)?;
            }
        },
        Kind::DesignElement => {
            if design.is_some() {
                return Err(invalid(
                    "designer extension has duplicate designElem elements",
                ));
            }
            *design = Some(Element {
                span: frame.start..end,
                close_start: if empty {
                    end
                } else {
                    frame_close_start(end, xml)
                },
                empty,
            });
        },
        Kind::Root | Kind::NonVisual | Kind::Other => {},
    }
    Ok(())
}

fn parse_snapshot(
    xml: &[u8],
    design: &Element,
    unknown_extensions: &[Opaque],
) -> Result<Option<Snapshot>> {
    let bytes = xml
        .get(design.span.clone())
        .ok_or_else(|| invalid("designElem range is invalid"))?;
    let mut reader = NsReader::from_reader(bytes);
    let decoder = reader.decoder();
    let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
    let is_p15 = match namespace {
        ResolveResult::Bound(Namespace(value)) => value == P15,
        ResolveResult::Unknown(value) => value.as_slice() == b"p15",
        _ => false,
    };
    if !is_p15 {
        return Err(invalid("designElem has the wrong namespace"));
    }
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return Err(invalid("designElem range does not begin with an element")),
    };
    let value = unqualified_attribute_value(&element, b"val", decoder)?
        .map(|value| parse_bool(&value))
        .transpose()?;
    Ok(Some(Snapshot::from_wire(
        value,
        unknown_extensions.to_vec(),
    )?))
}

fn validate_design_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let mut seen_value = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name == b"val" {
            if seen_value {
                return Err(invalid("designElem has duplicate val attributes"));
            }
            seen_value = true;
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            let _ = parse_bool(&value)?;
        } else if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            return Err(invalid("designElem has an unsupported attribute"));
        }
    }
    Ok(())
}

fn parse_bool(value: &str) -> Result<bool> {
    match value.trim() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("unsupported designElem boolean '{value}'"))),
    }
}

fn is_shape_root(local: &[u8]) -> bool {
    matches!(
        local,
        b"sp" | b"pic" | b"cxnSp" | b"graphicFrame" | b"grpSp"
    )
}

fn is_non_visual(local: &[u8]) -> bool {
    matches!(
        local,
        b"nvSpPr" | b"nvPicPr" | b"nvCxnSpPr" | b"nvGraphicFramePr" | b"nvGrpSpPr"
    )
}

fn namespace_kind(namespace: ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if value == PML_TRANSITIONAL => {
            NamespaceKind::Pml(Conformance::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if value == PML_STRICT => {
            NamespaceKind::Pml(Conformance::Strict)
        },
        ResolveResult::Bound(Namespace(value)) if value == P15 => NamespaceKind::P15,
        // A selected shape is often decoded as a standalone span even though
        // its PresentationML declarations live on the containing slide. Keep
        // conventional producer prefixes meaningful in that bounded context.
        ResolveResult::Unknown(prefix) if prefix.as_slice() == b"p" => {
            NamespaceKind::Pml(Conformance::Transitional)
        },
        ResolveResult::Unknown(prefix) if prefix.as_slice() == b"p15" => NamespaceKind::P15,
        _ => NamespaceKind::Other,
    }
}

fn pml_profile(namespace: NamespaceKind) -> Option<Conformance> {
    match namespace {
        NamespaceKind::Pml(profile) => Some(profile),
        NamespaceKind::P15 | NamespaceKind::Other => None,
    }
}

fn mark_other(frame: &mut Frame) {
    if matches!(frame.kind, Kind::ExtList | Kind::Extension { .. }) {
        frame.other_content = true;
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("designer XML offset does not fit usize"))
}

fn frame_close_start(end: usize, xml: &[u8]) -> usize {
    xml[..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .unwrap_or(end)
}

/// Serialize a fresh known design element using a locally bound p15 prefix.
pub(super) fn write_design_element(value: bool) -> Vec<u8> {
    format!(
        "<p15:designElem xmlns:p15=\"{}\" val=\"{}\"/>",
        super::NAMESPACE,
        if value { "true" } else { "false" }
    )
    .into_bytes()
}

pub(super) fn write_extension(conformance: Conformance, value: bool) -> Result<Vec<u8>> {
    let design = write_design_element(value);
    let xml = format!(
        "<p:ext xmlns:p=\"{}\" uri=\"{}\">{}</p:ext>",
        conformance.namespace(),
        super::EXTENSION_URI,
        String::from_utf8_lossy(&design)
    )
    .into_bytes();
    if xml.len() > MAX_EXTENSION_BYTES {
        return Err(Error::Limit {
            resource: "designer extension bytes",
            limit: MAX_EXTENSION_BYTES,
        });
    }
    Ok(xml)
}

pub(super) fn write_extension_list(conformance: Conformance, value: bool) -> Result<Vec<u8>> {
    let extension = write_extension(conformance, value)?;
    let mut xml = Vec::new();
    xml.try_reserve_exact(extension.len() + 64)
        .map_err(|source| Error::Allocation {
            resource: "designer extension list",
            source,
        })?;
    xml.extend_from_slice(format!("<p:extLst xmlns:p=\"{}\">", conformance.namespace()).as_bytes());
    xml.extend_from_slice(&extension);
    xml.extend_from_slice(b"</p:extLst>");
    Ok(xml)
}

pub(super) fn replace(xml: &[u8], range: Range<usize>, value: &[u8]) -> Result<Vec<u8>> {
    replace_xml(xml, range, value)
}
