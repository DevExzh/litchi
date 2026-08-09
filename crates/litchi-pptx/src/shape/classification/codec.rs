//! Lossless XML discovery and serialization for shape classification.

use std::ops::Range;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Opaque, Outcome, Snapshot};
use super::validation::{
    MAX_EXTENSION_BYTES, MAX_SHAPE_BYTES, MAX_SHAPE_DEPTH, count_node, invalid, push_unknown,
    xml_error,
};
use crate::tag::{Conformance, replace_xml};
use crate::{Error, Result};

const PML_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const PML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const P184: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2018/4/main";

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
    pub(super) classification: Option<Element>,
    pub(super) other_content: bool,
}

#[derive(Debug, Clone)]
pub(super) struct ExtensionList {
    pub(super) element: Element,
    pub(super) child_elements: usize,
    pub(super) other_content: bool,
}

/// The validated shape-local ranges needed by a transaction.
#[derive(Debug, Clone)]
pub(super) struct Layout {
    pub(super) conformance: Conformance,
    pub(super) nv_pr: Element,
    pub(super) ext_lst: Option<ExtensionList>,
    pub(super) known: Option<Extension>,
}

/// Parsed classification source plus its raw placement.
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
    Classification,
    Other,
}

#[derive(Debug, Clone, Copy)]
enum NamespaceKind {
    Pml(Conformance),
    P184,
    Other,
}

#[derive(Debug)]
struct Frame {
    kind: Kind,
    start: usize,
    other_content: bool,
    classification_seen: bool,
}

/// Parse a selected raw shape into a bounded classification source.
pub(super) fn read(xml: &[u8]) -> Result<Source> {
    if xml.is_empty() {
        return Err(invalid("classification shape XML is empty"));
    }
    if xml.len() > MAX_SHAPE_BYTES {
        return Err(Error::Limit {
            resource: "classification shape XML bytes",
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
    let mut classification = None;
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
                    let profile = pml_profile(namespace).ok_or_else(|| {
                        invalid("classification shape root is not PresentationML")
                    })?;
                    if !is_shape_root(element.local_name().as_ref()) {
                        return Err(invalid("classification owner is not a supported shape"));
                    }
                    conformance = Some(profile);
                    frames.push(Frame {
                        kind: Kind::Root,
                        start,
                        other_content: false,
                        classification_seen: false,
                    });
                    continue;
                }

                let kind = classify_start(
                    namespace,
                    &element,
                    decoder,
                    frames
                        .last_mut()
                        .ok_or_else(|| invalid("classification shape frame stack became empty"))?,
                    &mut nv_pr,
                    &mut ext_lst,
                    &mut known,
                    &mut conformance,
                )?;
                if frames.len() >= MAX_SHAPE_DEPTH {
                    return Err(Error::Limit {
                        resource: "classification shape XML depth",
                        limit: MAX_SHAPE_DEPTH,
                    });
                }
                frames.push(Frame {
                    kind,
                    start,
                    other_content: false,
                    classification_seen: false,
                });
            },
            Event::Empty(element) => {
                count_node(&mut nodes)?;
                if frames.is_empty() {
                    let profile = pml_profile(namespace).ok_or_else(|| {
                        invalid("classification shape root is not PresentationML")
                    })?;
                    if !is_shape_root(element.local_name().as_ref()) {
                        return Err(invalid("classification owner is not a supported shape"));
                    }
                    conformance = Some(profile);
                    frames.push(Frame {
                        kind: Kind::Root,
                        start,
                        other_content: false,
                        classification_seen: false,
                    });
                    finish_frame(
                        xml,
                        frames
                            .pop()
                            .ok_or_else(|| invalid("missing empty shape root"))?,
                        end,
                        true,
                        &mut nv_pr,
                        &mut ext_lst,
                        &mut known,
                        &mut unknown_extensions,
                        &mut classification,
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
                        .ok_or_else(|| invalid("classification shape frame stack became empty"))?,
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
                        classification_seen: false,
                    },
                    end,
                    true,
                    &mut nv_pr,
                    &mut ext_lst,
                    &mut known,
                    &mut unknown_extensions,
                    &mut classification,
                    &mut extension_children,
                )?;
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("classification shape XML stack underflow"))?;
                finish_frame(
                    xml,
                    frame,
                    end,
                    false,
                    &mut nv_pr,
                    &mut ext_lst,
                    &mut known,
                    &mut unknown_extensions,
                    &mut classification,
                    &mut extension_children,
                )?;
                if frames.is_empty() {
                    if end != xml.len() {
                        return Err(invalid("classification shape root does not cover its XML"));
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
                        Some(Kind::NvPr | Kind::ExtList)
                    )
                {
                    return Err(invalid(
                        "classification structural metadata contains non-whitespace text",
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
                if !text.decode().map_err(xml_error)?.trim().is_empty()
                    && matches!(
                        frames.last().map(|frame| frame.kind),
                        Some(Kind::NvPr | Kind::ExtList)
                    )
                {
                    return Err(invalid("classification structural metadata contains CDATA"));
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
                    Some(Kind::ExtList | Kind::Extension { .. })
                ) && let Some(frame) = frames.last_mut()
                {
                    frame.other_content = true;
                }
            },
            Event::GeneralRef(_)
                if matches!(
                    frames.last().map(|frame| frame.kind),
                    Some(Kind::NvPr | Kind::ExtList)
                ) =>
            {
                return Err(invalid(
                    "classification structural metadata contains an entity reference",
                ));
            },
            Event::Decl(_) if frames.is_empty() => {},
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) if !frames.is_empty() => {
                return Err(invalid("classification shape contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !root_closed || !frames.is_empty() {
        return Err(invalid("classification shape XML is unterminated"));
    }
    let conformance =
        conformance.ok_or_else(|| invalid("classification shape profile is missing"))?;
    let nv_pr = nv_pr.ok_or_else(|| invalid("classification shape has no direct p:nvPr"))?;
    let snapshot = known
        .as_ref()
        .and_then(|extension| extension.classification.as_ref())
        .map(|classification| parse_snapshot(xml, classification, &unknown_extensions))
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
                    return Err(invalid("classification shape has duplicate p:extLst"));
                }
                if let Some(profile) = pml {
                    if let Some(existing) = conformance
                        && *existing != profile
                    {
                        return Err(invalid(
                            "classification shape mixes Strict and Transitional PresentationML",
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
                    .ok_or_else(|| invalid("classification p:ext has no uri"))?;
                if uri.is_empty() {
                    return Err(invalid("classification p:ext has an empty uri"));
                }
                let is_known = uri == super::EXTENSION_URI;
                if is_known && known.is_some() {
                    return Err(invalid(
                        "classification shape has duplicate classification extensions",
                    ));
                }
                Kind::Extension { known: is_known }
            } else {
                mark_other(parent);
                Kind::Other
            }
        },
        Kind::Extension { known: is_known } => {
            if is_known && is_p184(namespace, element.local_name().as_ref()) {
                if local.as_ref() == b"classification" {
                    if parent.classification_seen {
                        return Err(invalid(
                            "classification extension has duplicate classification elements",
                        ));
                    }
                    validate_classification_attributes(element, decoder)?;
                    parent.classification_seen = true;
                    Kind::Classification
                } else {
                    mark_other(parent);
                    Kind::Other
                }
            } else {
                mark_other(parent);
                Kind::Other
            }
        },
        Kind::Classification => {
            return Err(invalid(
                "classification element cannot contain child elements",
            ));
        },
        Kind::Other => Kind::Other,
    };

    if let Kind::NvPr = kind
        && nv_pr.is_some()
    {
        return Err(invalid(
            "classification shape has duplicate p:nvPr elements",
        ));
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
    classification: &mut Option<Element>,
    extension_children: &mut usize,
) -> Result<()> {
    if end > xml.len() || frame.start >= end {
        return Err(invalid("classification element range is outside its shape"));
    }
    match frame.kind {
        Kind::NvPr => {
            if nv_pr.is_some() {
                return Err(invalid(
                    "classification shape has duplicate p:nvPr elements",
                ));
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
                return Err(invalid(
                    "classification shape has duplicate p:extLst elements",
                ));
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
                resource: "classification extension count",
                limit: super::validation::MAX_EXTENSION_COUNT,
            })?;
            if is_known {
                if known.is_some() {
                    return Err(invalid(
                        "classification shape has duplicate classification extensions",
                    ));
                }
                *known = Some(Extension {
                    element,
                    classification: classification.take(),
                    other_content: frame.other_content,
                });
            } else {
                let raw = xml
                    .get(frame.start..end)
                    .ok_or_else(|| invalid("classification opaque extension range is invalid"))?
                    .to_vec();
                push_unknown(unknown_extensions, raw)?;
            }
        },
        Kind::Classification => {
            if classification.is_some() {
                return Err(invalid(
                    "classification extension has duplicate classification elements",
                ));
            }
            *classification = Some(Element {
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
    classification: &Element,
    unknown_extensions: &[Opaque],
) -> Result<Option<Snapshot>> {
    let bytes = xml
        .get(classification.span.clone())
        .ok_or_else(|| invalid("classification element range is invalid"))?;
    let mut reader = NsReader::from_reader(bytes);
    let decoder = reader.decoder();
    let (_, event) = reader.read_resolved_event().map_err(xml_error)?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => {
            return Err(invalid(
                "classification range does not begin with an element",
            ));
        },
    };
    let outcome = unqualified_attribute_value(&element, b"val", decoder)?
        .map(|value| Outcome::parse(&value))
        .transpose()?;
    let unknown_extensions = unknown_extensions.to_vec();
    Ok(Some(Snapshot::from_wire(outcome, unknown_extensions)?))
}

fn validate_classification_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let mut seen_val = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name == b"val" {
            if seen_val {
                return Err(invalid(
                    "classification element has duplicate val attributes",
                ));
            }
            seen_val = true;
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            let _ = Outcome::parse(&value)?;
        } else if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            return Err(invalid(
                "classification element has an unsupported attribute",
            ));
        }
    }
    Ok(())
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
        ResolveResult::Bound(Namespace(value)) if value == P184 => NamespaceKind::P184,
        _ => NamespaceKind::Other,
    }
}

fn pml_profile(namespace: NamespaceKind) -> Option<Conformance> {
    match namespace {
        NamespaceKind::Pml(profile) => Some(profile),
        NamespaceKind::P184 | NamespaceKind::Other => None,
    }
}

fn is_p184(namespace: NamespaceKind, local: &[u8]) -> bool {
    local == b"classification" && matches!(namespace, NamespaceKind::P184)
}

fn mark_other(frame: &mut Frame) {
    if matches!(frame.kind, Kind::ExtList | Kind::Extension { .. }) {
        frame.other_content = true;
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("classification XML offset does not fit usize"))
}

fn frame_close_start(end: usize, xml: &[u8]) -> usize {
    xml[..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .unwrap_or(end)
}

/// Serialize one fresh known extension using the selected `PresentationML`
/// namespace profile.
pub(super) fn write_extension(conformance: Conformance, outcome: Outcome) -> Result<Vec<u8>> {
    let xml = format!(
        "<p:ext xmlns:p=\"{}\" uri=\"{}\"><p184:classification xmlns:p184=\"{}\" val=\"{}\"/></p:ext>",
        conformance.namespace(),
        super::EXTENSION_URI,
        super::NAMESPACE,
        outcome.wire(),
    )
    .into_bytes();
    if xml.len() > MAX_EXTENSION_BYTES {
        return Err(Error::Limit {
            resource: "classification extension bytes",
            limit: MAX_EXTENSION_BYTES,
        });
    }
    Ok(xml)
}

pub(super) fn write_extension_list(conformance: Conformance, outcome: Outcome) -> Result<Vec<u8>> {
    let extension = write_extension(conformance, outcome)?;
    let mut xml = Vec::new();
    xml.try_reserve_exact(extension.len() + 64)
        .map_err(|source| Error::Allocation {
            resource: "classification extension list",
            source,
        })?;
    xml.extend_from_slice(format!("<p:extLst xmlns:p=\"{}\">", conformance.namespace()).as_bytes());
    xml.extend_from_slice(&extension);
    xml.extend_from_slice(b"</p:extLst>");
    Ok(xml)
}

pub(super) fn write_classification(outcome: Outcome) -> Vec<u8> {
    format!(
        "<p184:classification xmlns:p184=\"{}\" val=\"{}\"/>",
        super::NAMESPACE,
        outcome.wire()
    )
    .into_bytes()
}

pub(super) fn replace(xml: &[u8], range: Range<usize>, value: &[u8]) -> Result<Vec<u8>> {
    replace_xml(xml, range, value)
}
