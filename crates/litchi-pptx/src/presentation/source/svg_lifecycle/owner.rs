//! Bounded, source-position preserving scanner for the native SVG owner.
//!
//! This module deliberately operates on the original slide bytes.  The shape
//! model can be useful for discovering ordinary pictures, but it is not a
//! source-range index: namespace declarations may be inherited and a model
//! round trip may normalize unrelated markup.  The lifecycle therefore uses
//! this scanner for both selection and every byte range which is later edited.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesRef, BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::reader::Reader;

use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DRAWINGML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const SVG_BLIP_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_EXTENSION_URI: &[u8] = b"{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

// These are intentionally local to the owner scanner.  The source package
// loader already limits the slide member, but keeping the parser's own cap
// makes this helper safe when called directly by a lifecycle test or future
// package reader.
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;
const MAX_NAME_BYTES: usize = 4 * 1024;
const MAX_NAMESPACE_LEXICAL_BYTES: usize = MAX_NAME_BYTES * 4;
const MAX_ATTRIBUTE_COUNT: usize = 256;
const MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT: usize = 256;
const MAX_ACTIVE_NAMESPACE_BINDINGS: usize = 16 * 1024;

/// A half-open byte range in the original slide source.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct ByteRange {
    pub(super) start: usize,
    pub(super) end: usize,
}

/// A source range for an element whose open and close tags may be edited
/// independently.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct ElementRange {
    pub(super) range: ByteRange,
    pub(super) start_end: usize,
    pub(super) close_start: Option<usize>,
    pub(super) prefix: Vec<u8>,
}

/// Exact source ranges needed by the SVG lifecycle writer.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct PictureLayout {
    pub(super) picture: ByteRange,
    pub(super) blip: ElementRange,
    pub(super) ext_list: Option<ElementRange>,
    pub(super) svg_extension: Option<ByteRange>,
    namespace_context: Arc<NamespaceContext>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Slide,
    CommonSlide,
    ShapeTree,
    Group,
    Picture,
    BlipFill,
    Blip,
    ExtList,
    UnknownExt,
    SvgExt,
    Mce,
    Other,
}

#[derive(Clone, Debug)]
struct Frame {
    kind: Kind,
    selected_context: bool,
    selected_picture: bool,
    namespace_declarations: usize,
    context_before: Arc<NamespaceContext>,
    source_start: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Binding {
    prefix: Vec<u8>,
    uri: Vec<u8>,
    depth: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct NamespaceContext {
    parent: Option<Arc<NamespaceContext>>,
    declarations: Arc<[Arc<Binding>]>,
}

#[derive(Debug)]
struct PictureCapture {
    picture: ByteRange,
    blip: Option<ElementRange>,
    ext_list: Option<ElementRange>,
    svg_extension: Option<ByteRange>,
    namespace_context: Arc<NamespaceContext>,
}

/// A small namespace resolver which performs no full active-scope cloning.
/// Every declaration is copied once, and a closing element drops only the
/// declarations introduced by that element.
struct Namespaces {
    bindings: Vec<Arc<Binding>>,
    by_prefix: HashMap<Vec<u8>, Vec<Arc<Binding>>>,
    active: usize,
    context: Arc<NamespaceContext>,
}

impl Default for Namespaces {
    fn default() -> Self {
        Self {
            bindings: Vec::new(),
            by_prefix: HashMap::new(),
            active: 0,
            context: Arc::new(NamespaceContext {
                parent: None,
                declarations: Arc::<[Arc<Binding>]>::from([]),
            }),
        }
    }
}

impl Namespaces {
    fn preflight(&self, element: &BytesStart<'_>, decoder: Decoder) -> Result<(usize, usize)> {
        let mut attributes = 0usize;
        let mut declarations = 0usize;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(xml_error)?;
            attributes = attributes
                .checked_add(1)
                .ok_or_else(|| limit("XML attributes", MAX_ATTRIBUTE_COUNT))?;
            if attributes > MAX_ATTRIBUTE_COUNT {
                return Err(limit("XML attributes", MAX_ATTRIBUTE_COUNT));
            }
            validate_qname(attribute.key.as_ref(), "attribute")?;
            validate_raw_attribute(attribute.value.as_ref())?;
            if let Some(prefix) = attribute.key.as_namespace_binding() {
                declarations = declarations.checked_add(1).ok_or_else(|| {
                    limit(
                        "namespace declarations",
                        MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT,
                    )
                })?;
                if declarations > MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT {
                    return Err(limit(
                        "namespace declarations per element",
                        MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT,
                    ));
                }
                if attribute.value.len() > MAX_NAMESPACE_LEXICAL_BYTES {
                    return Err(limit(
                        "namespace declaration lexical bytes",
                        MAX_NAMESPACE_LEXICAL_BYTES,
                    ));
                }
                let uri = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?;
                validate_xml_characters(uri.as_bytes())?;
                validate_namespace_binding(prefix, uri.as_bytes())?;
            }
        }
        let active = self
            .active
            .checked_add(declarations)
            .ok_or_else(|| limit("active namespace bindings", MAX_ACTIVE_NAMESPACE_BINDINGS))?;
        if active > MAX_ACTIVE_NAMESPACE_BINDINGS {
            return Err(limit(
                "active namespace bindings",
                MAX_ACTIVE_NAMESPACE_BINDINGS,
            ));
        }
        Ok((attributes, declarations))
    }

    fn push(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        depth: usize,
        declarations: usize,
    ) -> Result<()> {
        if declarations == 0 {
            return Ok(());
        }
        self.bindings
            .try_reserve(declarations)
            .map_err(|source| Error::Allocation {
                resource: "SVG owner namespace bindings",
                source,
            })?;
        self.by_prefix
            .try_reserve(declarations)
            .map_err(|source| Error::Allocation {
                resource: "SVG owner namespace prefix index",
                source,
            })?;
        let mut local = Vec::<Arc<Binding>>::new();
        local
            .try_reserve_exact(declarations)
            .map_err(|source| Error::Allocation {
                resource: "SVG owner namespace context",
                source,
            })?;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let Some(prefix) = attribute.key.as_namespace_binding() else {
                continue;
            };
            let uri = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?;
            let prefix = match prefix {
                quick_xml::name::PrefixDeclaration::Default => &[][..],
                quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
            };
            let mut prefix_copy = Vec::new();
            prefix_copy
                .try_reserve_exact(prefix.len())
                .map_err(|source| Error::Allocation {
                    resource: "SVG owner namespace prefix",
                    source,
                })?;
            prefix_copy.extend_from_slice(prefix);
            let mut uri_copy = Vec::new();
            uri_copy
                .try_reserve_exact(uri.len())
                .map_err(|source| Error::Allocation {
                    resource: "SVG owner namespace URI",
                    source,
                })?;
            uri_copy.extend_from_slice(uri.as_bytes());
            let binding = Arc::new(Binding {
                prefix: prefix_copy,
                uri: uri_copy,
                depth,
            });
            self.bindings.push(Arc::clone(&binding));
            let stack = self.by_prefix.entry(binding.prefix.clone()).or_default();
            stack.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "SVG owner namespace prefix stack",
                source,
            })?;
            stack.push(Arc::clone(&binding));
            local.push(binding);
        }
        let local_declarations: Arc<[Arc<Binding>]> = Arc::from(local.into_boxed_slice());
        self.context = Arc::new(NamespaceContext {
            parent: Some(Arc::clone(&self.context)),
            declarations: local_declarations,
        });
        self.active = self
            .active
            .checked_add(declarations)
            .ok_or_else(|| limit("active namespace bindings", MAX_ACTIVE_NAMESPACE_BINDINGS))?;
        Ok(())
    }

    fn pop(&mut self, depth: usize, declarations: usize) {
        if declarations == 0 {
            return;
        }
        let mut removed = 0usize;
        while removed < declarations {
            let Some(binding) = self.bindings.pop() else {
                break;
            };
            debug_assert_eq!(binding.depth, depth);
            let empty = if let Some(stack) = self.by_prefix.get_mut(&binding.prefix) {
                let _ = stack.pop();
                stack.is_empty()
            } else {
                false
            };
            if empty {
                self.by_prefix.remove(&binding.prefix);
            }
            removed += 1;
        }
        self.active = self.active.saturating_sub(removed);
    }

    fn resolve_element(&self, name: QName<'_>) -> Result<Option<&[u8]>> {
        let (local, prefix) = name.decompose();
        let _ = local;
        match prefix {
            Some(prefix) => self.resolve_prefix(prefix.as_ref(), true),
            None => Ok(self
                .by_prefix
                .get(&[][..])
                .and_then(|stack| stack.last())
                .and_then(|binding| (!binding.uri.is_empty()).then_some(binding.uri.as_slice()))),
        }
    }

    fn resolve_attribute(&self, name: QName<'_>) -> Result<Option<&[u8]>> {
        let (_, prefix) = name.decompose();
        match prefix {
            None => Ok(None),
            Some(prefix) => self.resolve_prefix(prefix.as_ref(), false),
        }
    }

    fn resolve_prefix(&self, prefix: &[u8], prefixed: bool) -> Result<Option<&[u8]>> {
        if prefix == b"xml" {
            return Ok(Some(XML_NAMESPACE));
        }
        let Some(binding) = self.by_prefix.get(prefix).and_then(|stack| stack.last()) else {
            return Err(invalid("XML name uses an unbound namespace prefix"));
        };
        if binding.uri.is_empty() && prefixed {
            return Err(invalid(
                "a prefixed XML name cannot use an empty namespace binding",
            ));
        }
        Ok((!binding.uri.is_empty()).then_some(binding.uri.as_slice()))
    }
}

fn namespace_context_bindings(context: &NamespaceContext) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut contexts = Vec::<&NamespaceContext>::new();
    let mut cursor = Some(context);
    let mut declaration_count = 0usize;
    while let Some(current) = cursor {
        declaration_count = declaration_count
            .checked_add(current.declarations.len())
            .ok_or_else(|| limit("active namespace bindings", MAX_ACTIVE_NAMESPACE_BINDINGS))?;
        if declaration_count > MAX_ACTIVE_NAMESPACE_BINDINGS {
            return Err(limit(
                "active namespace bindings",
                MAX_ACTIVE_NAMESPACE_BINDINGS,
            ));
        }
        contexts
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "SVG owner namespace context walk",
                source,
            })?;
        contexts.push(current);
        cursor = current.parent.as_deref();
    }
    let mut output = Vec::new();
    output
        .try_reserve(declaration_count)
        .map_err(|source| Error::Allocation {
            resource: "SVG owner inherited namespace bindings",
            source,
        })?;
    let mut seen = HashSet::<&[u8]>::new();
    seen.try_reserve(declaration_count)
        .map_err(|source| Error::Allocation {
            resource: "SVG owner inherited namespace prefix index",
            source,
        })?;
    for current in contexts {
        for binding in current.declarations.iter().rev() {
            if !seen.insert(binding.prefix.as_slice()) {
                continue;
            }
            let mut prefix = Vec::new();
            prefix
                .try_reserve_exact(binding.prefix.len())
                .map_err(|source| Error::Allocation {
                    resource: "SVG owner inherited namespace prefix",
                    source,
                })?;
            prefix.extend_from_slice(&binding.prefix);
            let mut uri = Vec::new();
            uri.try_reserve_exact(binding.uri.len())
                .map_err(|source| Error::Allocation {
                    resource: "SVG owner inherited namespace URI",
                    source,
                })?;
            uri.extend_from_slice(binding.uri.as_slice());
            output.push((prefix, uri));
        }
    }
    Ok(output)
}

fn root_fragment_namespace_info(fragment: &[u8]) -> Result<(usize, Vec<Vec<u8>>)> {
    let mut reader = Reader::from_reader(fragment);
    let origin = ReaderOrigin::of(fragment);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let end = origin
                    .offset(reader.buffer_position())
                    .ok_or_else(|| invalid("picture namespace root end exceeds usize"))?;
                let mut declared = Vec::<Vec<u8>>::new();
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(xml_error)?;
                    let Some(prefix) = attribute.key.as_namespace_binding() else {
                        continue;
                    };
                    let prefix = match prefix {
                        quick_xml::name::PrefixDeclaration::Default => &[][..],
                        quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
                    };
                    declared
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "SVG owner declared namespace prefixes",
                            source,
                        })?;
                    let mut prefix_copy = Vec::new();
                    prefix_copy
                        .try_reserve_exact(prefix.len())
                        .map_err(|source| Error::Allocation {
                            resource: "SVG owner declared namespace prefix",
                            source,
                        })?;
                    prefix_copy.extend_from_slice(prefix);
                    declared.push(prefix_copy);
                }
                return Ok((end, declared));
            },
            Event::Eof => return Err(invalid("picture fragment has no root element")),
            _ => {},
        }
        buffer.clear();
    }
}

fn namespace_context_additions_len(
    fragment: &[u8],
    context: &NamespaceContext,
    max_output_bytes: u64,
) -> Result<()> {
    if fragment.len() as u64 > max_output_bytes {
        return Err(limit(
            "SVG namespace-complete fragment bytes",
            usize::try_from(max_output_bytes).unwrap_or(usize::MAX),
        ));
    }
    let (_root_end, declared) = root_fragment_namespace_info(fragment)?;
    let mut contexts = Vec::<&NamespaceContext>::new();
    let mut cursor = Some(context);
    let mut declaration_count = 0usize;
    while let Some(current) = cursor {
        declaration_count = declaration_count
            .checked_add(current.declarations.len())
            .ok_or_else(|| limit("active namespace bindings", MAX_ACTIVE_NAMESPACE_BINDINGS))?;
        if declaration_count > MAX_ACTIVE_NAMESPACE_BINDINGS {
            return Err(limit(
                "active namespace bindings",
                MAX_ACTIVE_NAMESPACE_BINDINGS,
            ));
        }
        contexts
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "SVG owner namespace context walk",
                source,
            })?;
        contexts.push(current);
        cursor = current.parent.as_deref();
    }
    let mut seen = HashSet::<&[u8]>::new();
    seen.try_reserve(declaration_count)
        .map_err(|source| Error::Allocation {
            resource: "SVG owner inherited namespace prefix index",
            source,
        })?;
    let mut additions_len = 0usize;
    for current in contexts {
        for binding in current.declarations.iter().rev() {
            if declared
                .iter()
                .any(|prefix| prefix.as_slice() == binding.prefix.as_slice())
                || !seen.insert(binding.prefix.as_slice())
            {
                continue;
            }
            let uri = std::str::from_utf8(binding.uri.as_slice())
                .map_err(|error| Error::Xml(error.to_string()))?;
            let escaped_len = namespace_uri_escaped_len(uri)?;
            let (fixed_len, prefix_len) = if binding.prefix.is_empty() {
                (9usize, 0usize)
            } else {
                (10usize, binding.prefix.len())
            };
            additions_len = additions_len
                .checked_add(fixed_len)
                .and_then(|length| length.checked_add(prefix_len))
                .and_then(|length| length.checked_add(escaped_len))
                .ok_or_else(|| invalid("picture namespace declaration size overflows"))?;
        }
    }
    let output_len = fragment
        .len()
        .checked_add(additions_len)
        .ok_or_else(|| invalid("namespace-complete picture size overflows"))?;
    if output_len as u64 > max_output_bytes {
        return Err(limit(
            "SVG namespace-complete fragment bytes",
            usize::try_from(max_output_bytes).unwrap_or(usize::MAX),
        ));
    }
    Ok(())
}

/// Locate the exact raw ranges for one scene-order `p:pic`.
pub(super) fn locate(xml: &[u8], image_position: usize) -> Result<PictureLayout> {
    let layouts = locate_all(xml)?;
    let len = layouts.len();
    layouts
        .into_iter()
        .nth(image_position)
        .ok_or(Error::IndexOutOfBounds {
            index: image_position,
            len,
        })
}

/// Scan a slide once and retain exact source ranges for every direct
/// scene-order picture.  Callers which already have the scene model can then
/// parse each sliced picture without rescanning the complete slide for every
/// image.
pub(super) fn locate_all(xml: &[u8]) -> Result<Vec<PictureLayout>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("SVG owner slide XML bytes", MAX_XML_BYTES));
    }

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let decoder = reader.decoder();
    let origin = ReaderOrigin::of(xml);
    let mut namespaces = Namespaces::default();
    let mut frames = Vec::<Frame>::new();
    frames.try_reserve(32).map_err(|source| Error::Allocation {
        resource: "SVG owner XML stack",
        source,
    })?;
    let mut buffer = Vec::new();
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prolog = true;
    let mut picture_count = 0usize;
    let mut captures = Vec::<PictureCapture>::new();
    let mut active_capture = None::<usize>;

    loop {
        let event_start = position(&reader, origin)?;
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let event_end = position(&reader, origin)?;
        if event_end < event_start || event_end > xml.len() {
            return Err(invalid("XML event range is outside the source bytes"));
        }
        match event {
            Event::Start(element) => {
                nodes = add_count(nodes, MAX_XML_NODES, "SVG owner XML nodes")?;
                let depth = frames
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("SVG owner XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("SVG owner XML depth", MAX_XML_DEPTH));
                }
                if root_closed {
                    return Err(invalid("slide XML contains a second root element"));
                }
                let (_, declaration_count) = namespaces.preflight(&element, decoder)?;
                let context_before = Arc::clone(&namespaces.context);
                namespaces.push(&element, decoder, depth, declaration_count)?;
                let resolved = namespaces.resolve_element(element.name())?;
                validate_start_attributes(&element, &namespaces, decoder)?;
                let selected_context = frames.last().is_some_and(|frame| frame.selected_context);
                let mce_ancestor = frames.iter().any(|frame| frame.kind == Kind::Mce);
                let opaque_ancestor = frames.iter().any(|frame| frame.kind == Kind::UnknownExt);
                let mut kind = classify(
                    resolved,
                    element.name(),
                    frames.last().map(|frame| frame.kind),
                    selected_context,
                    &frames,
                );
                if kind == Kind::Other {
                    kind = classify_ext(
                        resolved,
                        element.name(),
                        frames.last().map(|frame| frame.kind),
                        selected_context,
                        &element,
                        decoder,
                    )?;
                }
                if frames.is_empty() {
                    if root_seen || !is_name(resolved, element.name(), b"sld", PML, STRICT_PML) {
                        return Err(invalid("slide XML must contain exactly one p:sld root"));
                    }
                    root_seen = true;
                }
                if is_name(resolved, element.name(), b"pic", PML, STRICT_PML)
                    && mce_ancestor
                    && !opaque_ancestor
                    && frames
                        .iter()
                        .any(|frame| matches!(frame.kind, Kind::ShapeTree | Kind::Group))
                {
                    return Err(Error::UnsafeEdit {
                        operation: "source-backed picture inventory",
                        reason: "picture ownership inside an MCE branch is ambiguous",
                    });
                }
                if kind == Kind::Mce && selected_context && !opaque_ancestor {
                    return Err(Error::UnsafeEdit {
                        operation: "source-backed picture inventory",
                        reason: "picture markup compatibility is unsupported by the SVG owner",
                    });
                }
                if selected_context
                    && frames
                        .last()
                        .is_some_and(|frame| frame.kind == Kind::ExtList)
                    && !matches!(kind, Kind::UnknownExt | Kind::SvgExt)
                {
                    return Err(invalid("a:extLst direct children must be a:ext"));
                }
                if selected_context
                    && frames
                        .last()
                        .is_some_and(|frame| frame.kind == Kind::SvgExt)
                    && !is_name(
                        resolved,
                        element.name(),
                        b"svgBlip",
                        SVG_BLIP_NAMESPACE,
                        SVG_BLIP_NAMESPACE,
                    )
                {
                    return Err(invalid(
                        "recognized SVG extension has a non-SVG svgBlip child",
                    ));
                }
                let selected_picture = kind == Kind::Picture
                    && !mce_ancestor
                    && !opaque_ancestor
                    && matches!(
                        frames.last().map(|frame| frame.kind),
                        Some(Kind::ShapeTree | Kind::Group)
                    );
                let selected_picture = if selected_picture {
                    picture_count = picture_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("picture count overflows usize"))?;
                    if active_capture.is_some() {
                        return Err(invalid("scene-order pictures cannot be nested"));
                    }
                    captures
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "SVG owner picture captures",
                            source,
                        })?;
                    captures.push(PictureCapture {
                        picture: ByteRange {
                            start: event_start,
                            end: event_end,
                        },
                        blip: None,
                        ext_list: None,
                        svg_extension: None,
                        namespace_context: Arc::clone(&namespaces.context),
                    });
                    active_capture = Some(captures.len() - 1);
                    true
                } else {
                    false
                };
                let selected_context = selected_context || selected_picture;
                let kind = if selected_picture {
                    Kind::Picture
                } else {
                    kind
                };
                if selected_context && kind == Kind::ExtList {
                    let capture = active_capture
                        .and_then(|index| captures.get_mut(index))
                        .ok_or_else(|| invalid("extLst has no active picture owner"))?;
                    if capture.ext_list.is_none() {
                        capture.ext_list = Some(ElementRange {
                            range: ByteRange {
                                start: event_start,
                                end: event_end,
                            },
                            start_end: event_end,
                            close_start: None,
                            prefix: qname_prefix(element.name().as_ref())?,
                        });
                    }
                }
                if selected_context && kind == Kind::Blip {
                    let capture = active_capture
                        .and_then(|index| captures.get_mut(index))
                        .ok_or_else(|| invalid("blip has no active picture owner"))?;
                    if capture.blip.is_none() {
                        capture.blip = Some(ElementRange {
                            range: ByteRange {
                                start: event_start,
                                end: event_end,
                            },
                            start_end: event_end,
                            close_start: None,
                            prefix: qname_prefix(element.name().as_ref())?,
                        });
                    }
                }
                if selected_context && kind == Kind::SvgExt {
                    let capture = active_capture
                        .and_then(|index| captures.get_mut(index))
                        .ok_or_else(|| invalid("SVG extension has no active picture owner"))?;
                    if capture.svg_extension.is_none() {
                        capture.svg_extension = Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        });
                    }
                }
                frames.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "SVG owner XML stack",
                    source,
                })?;
                frames.push(Frame {
                    kind,
                    selected_context,
                    selected_picture,
                    namespace_declarations: declaration_count,
                    context_before,
                    source_start: event_start,
                });
                prolog = false;
            },
            Event::Empty(element) => {
                nodes = add_count(nodes, MAX_XML_NODES, "SVG owner XML nodes")?;
                let depth = frames
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("SVG owner XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("SVG owner XML depth", MAX_XML_DEPTH));
                }
                if root_closed {
                    return Err(invalid("slide XML contains a second root element"));
                }
                let (_, declaration_count) = namespaces.preflight(&element, decoder)?;
                let context_before = Arc::clone(&namespaces.context);
                namespaces.push(&element, decoder, depth, declaration_count)?;
                let resolved = namespaces.resolve_element(element.name())?;
                validate_start_attributes(&element, &namespaces, decoder)?;
                let selected_context = frames.last().is_some_and(|frame| frame.selected_context);
                let mce_ancestor = frames.iter().any(|frame| frame.kind == Kind::Mce);
                let opaque_ancestor = frames.iter().any(|frame| frame.kind == Kind::UnknownExt);
                let mut kind = classify(
                    resolved,
                    element.name(),
                    frames.last().map(|frame| frame.kind),
                    selected_context,
                    &frames,
                );
                if kind == Kind::Other {
                    kind = classify_ext(
                        resolved,
                        element.name(),
                        frames.last().map(|frame| frame.kind),
                        selected_context,
                        &element,
                        decoder,
                    )?;
                }
                if frames.is_empty() {
                    if root_seen || !is_name(resolved, element.name(), b"sld", PML, STRICT_PML) {
                        return Err(invalid("slide XML must contain exactly one p:sld root"));
                    }
                    return Err(invalid("slide root p:sld cannot be empty"));
                }
                if is_name(resolved, element.name(), b"pic", PML, STRICT_PML)
                    && mce_ancestor
                    && !opaque_ancestor
                    && frames
                        .iter()
                        .any(|frame| matches!(frame.kind, Kind::ShapeTree | Kind::Group))
                {
                    return Err(Error::UnsafeEdit {
                        operation: "source-backed picture inventory",
                        reason: "picture ownership inside an MCE branch is ambiguous",
                    });
                }
                if kind == Kind::Mce && selected_context && !opaque_ancestor {
                    return Err(Error::UnsafeEdit {
                        operation: "source-backed picture inventory",
                        reason: "picture markup compatibility is unsupported by the SVG owner",
                    });
                }
                if selected_context
                    && frames
                        .last()
                        .is_some_and(|frame| frame.kind == Kind::ExtList)
                    && !matches!(kind, Kind::UnknownExt | Kind::SvgExt)
                {
                    return Err(invalid("a:extLst direct children must be a:ext"));
                }
                if selected_context
                    && frames
                        .last()
                        .is_some_and(|frame| frame.kind == Kind::SvgExt)
                    && !is_name(
                        resolved,
                        element.name(),
                        b"svgBlip",
                        SVG_BLIP_NAMESPACE,
                        SVG_BLIP_NAMESPACE,
                    )
                {
                    return Err(invalid(
                        "recognized SVG extension has a non-SVG svgBlip child",
                    ));
                }
                let scene_picture = kind == Kind::Picture
                    && !mce_ancestor
                    && !opaque_ancestor
                    && matches!(
                        frames.last().map(|frame| frame.kind),
                        Some(Kind::ShapeTree | Kind::Group)
                    );
                let selected_picture = if scene_picture {
                    picture_count = picture_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("picture count overflows usize"))?;
                    if active_capture.is_some() {
                        return Err(invalid("scene-order pictures cannot be nested"));
                    }
                    captures
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "SVG owner picture captures",
                            source,
                        })?;
                    captures.push(PictureCapture {
                        picture: ByteRange {
                            start: event_start,
                            end: event_end,
                        },
                        blip: None,
                        ext_list: None,
                        svg_extension: None,
                        namespace_context: Arc::clone(&namespaces.context),
                    });
                    active_capture = Some(captures.len() - 1);
                    true
                } else {
                    false
                };
                let selected_context = selected_context || selected_picture;
                let kind = if selected_picture {
                    Kind::Picture
                } else {
                    kind
                };
                if selected_context && kind == Kind::ExtList {
                    let capture = active_capture
                        .and_then(|index| captures.get_mut(index))
                        .ok_or_else(|| invalid("extLst has no active picture owner"))?;
                    if capture.ext_list.is_none() {
                        capture.ext_list = Some(ElementRange {
                            range: ByteRange {
                                start: event_start,
                                end: event_end,
                            },
                            start_end: event_end,
                            close_start: None,
                            prefix: qname_prefix(element.name().as_ref())?,
                        });
                    }
                }
                if selected_context && kind == Kind::Blip {
                    let capture = active_capture
                        .and_then(|index| captures.get_mut(index))
                        .ok_or_else(|| invalid("blip has no active picture owner"))?;
                    if capture.blip.is_none() {
                        capture.blip = Some(ElementRange {
                            range: ByteRange {
                                start: event_start,
                                end: event_end,
                            },
                            start_end: event_end,
                            close_start: None,
                            prefix: qname_prefix(element.name().as_ref())?,
                        });
                    }
                }
                if selected_context && kind == Kind::SvgExt {
                    let capture = active_capture
                        .and_then(|index| captures.get_mut(index))
                        .ok_or_else(|| invalid("SVG extension has no active picture owner"))?;
                    if capture.svg_extension.is_none() {
                        capture.svg_extension = Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        });
                    }
                }
                if selected_picture {
                    active_capture = None;
                }
                namespaces.pop(depth, declaration_count);
                namespaces.context = context_before;
                prolog = false;
            },
            Event::End(element) => {
                if frames.is_empty() {
                    return Err(invalid("slide XML contains an unmatched end element"));
                }
                let depth = frames.len();
                validate_qname(element.name().as_ref(), "end element")?;
                let resolved = namespaces.resolve_element(element.name())?;
                let frame = frames.pop().ok_or_else(|| invalid("XML stack underflow"))?;
                if frame.kind == Kind::Picture && frame.selected_picture {
                    let index = active_capture
                        .take()
                        .ok_or_else(|| invalid("picture close has no active capture"))?;
                    let capture = captures
                        .get_mut(index)
                        .ok_or_else(|| invalid("picture capture index is outside the scan"))?;
                    capture.picture.end = event_end;
                }
                if frame.selected_context {
                    match frame.kind {
                        Kind::Blip => {
                            let index = active_capture
                                .ok_or_else(|| invalid("blip close has no active picture"))?;
                            let value = captures
                                .get_mut(index)
                                .and_then(|capture| capture.blip.as_mut())
                                .ok_or_else(|| {
                                    invalid("selected blip close has no opening range")
                                })?;
                            if value.range.start == frame.source_start {
                                value.range.end = event_end;
                                value.close_start = Some(event_start);
                            }
                        },
                        Kind::ExtList => {
                            let index = active_capture
                                .ok_or_else(|| invalid("extLst close has no active picture"))?;
                            let value = captures
                                .get_mut(index)
                                .and_then(|capture| capture.ext_list.as_mut())
                                .ok_or_else(|| {
                                    invalid("selected extLst close has no opening range")
                                })?;
                            if value.range.start == frame.source_start {
                                value.range.end = event_end;
                                value.close_start = Some(event_start);
                            }
                        },
                        Kind::SvgExt => {
                            let index = active_capture
                                .ok_or_else(|| invalid("SVG extension close has no picture"))?;
                            if let Some(value) = captures
                                .get_mut(index)
                                .and_then(|capture| capture.svg_extension.as_mut())
                            {
                                if value.start == frame.source_start {
                                    value.end = event_end;
                                }
                            }
                        },
                        _ => {},
                    }
                }
                if frame.kind == Kind::Slide {
                    if !is_name(resolved, element.name(), b"sld", PML, STRICT_PML) {
                        return Err(invalid("slide root close does not match p:sld"));
                    }
                    root_closed = true;
                }
                namespaces.pop(depth, frame.namespace_declarations);
                namespaces.context = frame.context_before;
                prolog = false;
            },
            Event::Decl(declaration) => {
                if declaration_seen || !prolog || root_seen {
                    return Err(invalid(
                        "XML declaration is not at the beginning of the slide",
                    ));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is forbidden in slide XML")),
            Event::Text(text) => {
                validate_xml_text(text.as_ref())?;
                let outside = frames.is_empty();
                if outside && (!root_seen || root_closed) && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid(
                        "non-whitespace text appears outside the slide root",
                    ));
                }
                validate_text_context(&frames, outside, !is_xml_whitespace(text.as_ref()))?;
                prolog = false;
            },
            Event::CData(data) => {
                validate_xml_text(data.as_ref())?;
                if frames.is_empty() {
                    return Err(invalid("CDATA appears outside the slide root"));
                }
                validate_text_context(&frames, false, true)?;
                prolog = false;
            },
            Event::GeneralRef(reference) => {
                validate_reference(&reference)?;
                if frames.is_empty() {
                    return Err(invalid("entity reference appears outside the slide root"));
                }
                validate_text_context(&frames, false, true)?;
                prolog = false;
            },
            Event::Comment(comment) => {
                validate_xml_text(comment.as_ref())?;
                prolog = false;
            },
            Event::PI(pi) => {
                validate_processing_instruction(pi.as_ref())?;
                prolog = false;
            },
            Event::Eof => break,
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || !frames.is_empty() {
        return Err(invalid("slide XML has an incomplete p:sld root"));
    }
    if active_capture.is_some() {
        return Err(invalid("selected picture is not closed"));
    }
    let mut layouts = Vec::new();
    layouts
        .try_reserve_exact(captures.len())
        .map_err(|source| Error::Allocation {
            resource: "SVG owner picture layout index",
            source,
        })?;
    for capture in captures {
        let blip = capture
            .blip
            .ok_or_else(|| Error::Relationship("picture lacks a:blip".into()))?;
        layouts.push(PictureLayout {
            picture: capture.picture,
            blip,
            ext_list: capture.ext_list,
            svg_extension: capture.svg_extension,
            namespace_context: capture.namespace_context,
        });
    }
    Ok(layouts)
}

/// Return a parser-only copy of the selected picture with ancestor namespace
/// bindings declared on its root.  The lifecycle edits the original source
/// ranges; this copy exists because the legacy relationship grammar parses a
/// sliced `p:pic` fragment and therefore cannot see declarations inherited
/// from `p:sld`/`p:spTree`.
pub(super) fn namespace_complete_picture(
    layout: &PictureLayout,
    fragment: &[u8],
    max_output_bytes: u64,
) -> Result<Vec<u8>> {
    let max_output_bytes = max_output_bytes.min(MAX_XML_BYTES as u64);
    namespace_context_additions_len(fragment, &layout.namespace_context, max_output_bytes)?;
    let inherited = namespace_context_bindings(&layout.namespace_context)?;
    namespace_complete_element_fragment(fragment, &inherited, max_output_bytes)
}

pub(super) fn namespace_complete_element_fragment(
    fragment: &[u8],
    inherited_namespaces: &[(Vec<u8>, Vec<u8>)],
    max_output_bytes: u64,
) -> Result<Vec<u8>> {
    let max_output_bytes = max_output_bytes.min(MAX_XML_BYTES as u64);
    if fragment.len() as u64 > max_output_bytes {
        return Err(Error::Limit {
            resource: "SVG namespace-complete fragment bytes",
            limit: usize::try_from(max_output_bytes).unwrap_or(usize::MAX),
        });
    }
    if inherited_namespaces.is_empty() {
        let mut output = Vec::new();
        output
            .try_reserve_exact(fragment.len())
            .map_err(|source| Error::Allocation {
                resource: "SVG namespace-complete fragment",
                source,
            })?;
        output.extend_from_slice(fragment);
        return Ok(output);
    }
    let mut reader = Reader::from_reader(fragment);
    let origin = ReaderOrigin::of(fragment);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let end = origin
                    .offset(reader.buffer_position())
                    .ok_or_else(|| invalid("picture namespace root end exceeds usize"))?;
                let mut declared = Vec::<Vec<u8>>::new();
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(xml_error)?;
                    let Some(prefix) = attribute.key.as_namespace_binding() else {
                        continue;
                    };
                    let prefix = match prefix {
                        quick_xml::name::PrefixDeclaration::Default => &[][..],
                        quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
                    };
                    declared
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "SVG owner declared namespace prefixes",
                            source,
                        })?;
                    let mut prefix_copy = Vec::new();
                    prefix_copy
                        .try_reserve_exact(prefix.len())
                        .map_err(|source| Error::Allocation {
                            resource: "SVG owner declared namespace prefix",
                            source,
                        })?;
                    prefix_copy.extend_from_slice(prefix);
                    declared.push(prefix_copy);
                }
                let insertion = if fragment
                    .get(..end)
                    .is_some_and(|opening| opening.ends_with(b"/>"))
                {
                    end.checked_sub(2)
                        .ok_or_else(|| invalid("picture self-closing root is truncated"))?
                } else {
                    end.checked_sub(1)
                        .ok_or_else(|| invalid("picture root is truncated"))?
                };
                let mut additions_len = 0usize;
                for (prefix, uri) in inherited_namespaces {
                    if declared
                        .iter()
                        .any(|old| old.as_slice() == prefix.as_slice())
                    {
                        continue;
                    }
                    let uri =
                        std::str::from_utf8(uri).map_err(|error| Error::Xml(error.to_string()))?;
                    let escaped_len = namespace_uri_escaped_len(uri)?;
                    let (fixed_len, prefix_len) = if prefix.is_empty() {
                        (9usize, 0usize)
                    } else {
                        (10usize, prefix.len())
                    };
                    additions_len = additions_len
                        .checked_add(fixed_len)
                        .and_then(|length| length.checked_add(prefix_len))
                        .and_then(|length| length.checked_add(escaped_len))
                        .ok_or_else(|| invalid("picture namespace declaration size overflows"))?;
                }
                let output_len = fragment
                    .len()
                    .checked_add(additions_len)
                    .ok_or_else(|| invalid("namespace-complete picture size overflows"))?;
                if output_len as u64 > max_output_bytes {
                    return Err(Error::Limit {
                        resource: "SVG namespace-complete fragment bytes",
                        limit: usize::try_from(max_output_bytes).unwrap_or(usize::MAX),
                    });
                }
                let mut output = Vec::new();
                output
                    .try_reserve_exact(output_len)
                    .map_err(|source| Error::Allocation {
                        resource: "namespace-complete SVG picture fragment",
                        source,
                    })?;
                output.extend_from_slice(&fragment[..insertion]);
                for (prefix, uri) in inherited_namespaces {
                    if declared
                        .iter()
                        .any(|old| old.as_slice() == prefix.as_slice())
                    {
                        continue;
                    }
                    if prefix.is_empty() {
                        output.extend_from_slice(b" xmlns=\"");
                    } else {
                        output.extend_from_slice(b" xmlns:");
                        output.extend_from_slice(prefix);
                        output.extend_from_slice(b"=\"");
                    }
                    let uri =
                        std::str::from_utf8(uri).map_err(|error| Error::Xml(error.to_string()))?;
                    push_namespace_uri(&mut output, uri)?;
                    output.push(b'"');
                }
                output.extend_from_slice(&fragment[insertion..]);
                return Ok(output);
            },
            Event::Eof => {
                return Err(invalid("picture fragment has no root element"));
            },
            _ => {},
        }
        buffer.clear();
    }
}

fn classify(
    namespace: Option<&[u8]>,
    name: QName<'_>,
    parent: Option<Kind>,
    selected_context: bool,
    frames: &[Frame],
) -> Kind {
    if namespace.is_some_and(|value| value == MCE) {
        return Kind::Mce;
    }
    if is_name(namespace, name, b"sld", PML, STRICT_PML) && parent.is_none() {
        return Kind::Slide;
    }
    if is_name(namespace, name, b"cSld", PML, STRICT_PML) && parent == Some(Kind::Slide) {
        return Kind::CommonSlide;
    }
    if is_name(namespace, name, b"spTree", PML, STRICT_PML) && parent == Some(Kind::CommonSlide) {
        return Kind::ShapeTree;
    }
    if is_name(namespace, name, b"grpSp", PML, STRICT_PML)
        && matches!(parent, Some(Kind::ShapeTree | Kind::Group))
        && !frames.iter().any(|frame| frame.kind == Kind::Mce)
    {
        return Kind::Group;
    }
    if is_name(namespace, name, b"pic", PML, STRICT_PML)
        && matches!(parent, Some(Kind::ShapeTree | Kind::Group))
    {
        return Kind::Picture;
    }
    if selected_context
        && is_name(namespace, name, b"blipFill", PML, STRICT_PML)
        && parent == Some(Kind::Picture)
    {
        return Kind::BlipFill;
    }
    if selected_context
        && is_name(namespace, name, b"blip", DRAWINGML, STRICT_DRAWINGML)
        && parent == Some(Kind::BlipFill)
    {
        return Kind::Blip;
    }
    if selected_context
        && is_name(namespace, name, b"extLst", DRAWINGML, STRICT_DRAWINGML)
        && parent == Some(Kind::Blip)
    {
        return Kind::ExtList;
    }
    if selected_context
        && is_name(namespace, name, b"ext", DRAWINGML, STRICT_DRAWINGML)
        && parent == Some(Kind::ExtList)
    {
        return Kind::Other;
    }
    Kind::Other
}

fn classify_ext(
    namespace: Option<&[u8]>,
    name: QName<'_>,
    parent: Option<Kind>,
    selected_context: bool,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Kind> {
    if selected_context
        && parent == Some(Kind::ExtList)
        && is_name(namespace, name, b"ext", DRAWINGML, STRICT_DRAWINGML)
    {
        if extension_uri_is_supported(element, decoder)? {
            return Ok(Kind::SvgExt);
        }
        return Ok(Kind::UnknownExt);
    }
    Ok(Kind::Other)
}

fn is_name(
    namespace: Option<&[u8]>,
    name: QName<'_>,
    local: &[u8],
    expected: &[u8],
    strict: &[u8],
) -> bool {
    name.local_name().as_ref() == local
        && namespace.is_some_and(|value| value == expected || value == strict)
}

fn extension_uri_is_supported(element: &BytesStart<'_>, decoder: Decoder) -> Result<bool> {
    let mut uri = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.key.prefix().is_some() || attribute.key.as_ref() != b"uri" {
            return Err(invalid("a:ext has an unsupported non-namespace attribute"));
        }
        if uri.is_some() {
            return Err(invalid("a:ext has duplicate uri attributes"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        if value.is_empty() {
            return Err(invalid("a:ext uri is empty"));
        }
        uri = Some(value);
    }
    let value = uri.ok_or_else(|| invalid("a:ext is missing its uri attribute"))?;
    Ok(xsd_token_is(value.as_bytes(), SVG_EXTENSION_URI))
}

fn xsd_token_is(value: &[u8], expected: &[u8]) -> bool {
    let mut atoms = value
        .split(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
        .filter(|atom| !atom.is_empty());
    atoms.next() == Some(expected) && atoms.next().is_none()
}

fn validate_start_attributes(
    element: &BytesStart<'_>,
    namespaces: &Namespaces,
    decoder: Decoder,
) -> Result<()> {
    let mut seen: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let namespace = namespaces.resolve_attribute(attribute.key)?;
        let local = attribute.key.local_name();
        let uri = namespace.unwrap_or(&[]);
        if attribute.key.prefix().is_some() && uri.is_empty() {
            return Err(invalid("a prefixed attribute has no namespace binding"));
        }
        let duplicate = seen.iter().any(|(old_uri, old_local)| {
            old_uri.as_slice() == uri && old_local.as_slice() == local.as_ref()
        });
        if duplicate {
            return Err(invalid("element has duplicate expanded attributes"));
        }
        seen.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "SVG owner expanded attribute set",
            source,
        })?;
        seen.push((uri.to_vec(), local.as_ref().to_vec()));
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        validate_xml_characters(value.as_bytes())?;
    }
    Ok(())
}

fn validate_text_context(frames: &[Frame], outside: bool, non_whitespace: bool) -> Result<()> {
    if outside || !non_whitespace {
        return Ok(());
    }
    let Some(frame) = frames.last() else {
        return Ok(());
    };
    if matches!(
        frame.kind,
        Kind::Slide
            | Kind::CommonSlide
            | Kind::ShapeTree
            | Kind::Group
            | Kind::Picture
            | Kind::BlipFill
            | Kind::Blip
            | Kind::ExtList
            | Kind::SvgExt
    ) {
        return Err(invalid(
            "known slide/picture container contains non-whitespace character data",
        ));
    }
    Ok(())
}

fn validate_xml_text(bytes: &[u8]) -> Result<()> {
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Xml(error.to_string()))?;
    if bytes.windows(3).any(|window| window == b"]]>") {
        return Err(invalid("XML text contains the forbidden ]]> delimiter"));
    }
    validate_xml_characters(value.as_bytes())
}

fn validate_xml_characters(bytes: &[u8]) -> Result<()> {
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Xml(error.to_string()))?;
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid("XML value contains an invalid XML character"))
    }
}

fn validate_raw_attribute(bytes: &[u8]) -> Result<()> {
    if bytes.contains(&b'<') {
        return Err(invalid("XML attribute contains a forbidden raw delimiter"));
    }
    Ok(())
}

fn validate_reference(reference: &BytesRef<'_>) -> Result<()> {
    let bytes = reference.as_ref();
    if bytes == b"amp" || bytes == b"lt" || bytes == b"gt" || bytes == b"quot" || bytes == b"apos" {
        return Ok(());
    }
    let character = reference
        .resolve_char_ref()
        .map_err(xml_error)?
        .ok_or_else(|| invalid("unknown XML entity reference"))?;
    if !is_xml_char(character) {
        return Err(invalid(
            "XML entity reference resolves to an invalid XML character",
        ));
    }
    Ok(())
}

fn validate_namespace_binding(
    prefix: quick_xml::name::PrefixDeclaration<'_>,
    uri: &[u8],
) -> Result<()> {
    let prefix = match prefix {
        quick_xml::name::PrefixDeclaration::Default => &[][..],
        quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
    };
    if prefix.len() > MAX_NAME_BYTES || uri.len() > MAX_NAME_BYTES {
        return Err(limit("namespace prefix or URI bytes", MAX_NAME_BYTES));
    }
    if prefix.is_empty() {
        if uri == XML_NAMESPACE || uri == XMLNS_NAMESPACE {
            return Err(invalid(
                "reserved XML namespace URI is not a default binding",
            ));
        }
        return Ok(());
    }
    if prefix == b"xmlns" {
        return Err(invalid("xmlns is not a bindable namespace prefix"));
    }
    if prefix == b"xml" {
        if uri != XML_NAMESPACE {
            return Err(invalid("xml prefix is bound to the wrong namespace URI"));
        }
        return Ok(());
    }
    if uri.is_empty() {
        return Err(invalid("a non-default namespace prefix cannot be empty"));
    }
    if uri == XML_NAMESPACE || uri == XMLNS_NAMESPACE {
        return Err(invalid(
            "reserved XML namespace URI has a non-reserved prefix",
        ));
    }
    Ok(())
}

fn validate_qname(bytes: &[u8], kind: &str) -> Result<()> {
    if bytes.is_empty() || bytes.len() > MAX_NAME_BYTES {
        return Err(invalid("XML qualified name exceeds its bound"));
    }
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Xml(error.to_string()))?;
    if !litchi_ooxml_common::xml_name::is_qualified_name(value) {
        return Err(Error::Invalid(format!("invalid XML {kind} qualified name")));
    }
    if bytes
        .split(|byte| *byte == b':')
        .any(|component| component.len() > MAX_NAME_BYTES)
    {
        return Err(limit(
            "XML namespace prefix or local name bytes",
            MAX_NAME_BYTES,
        ));
    }
    Ok(())
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let version = declaration.version().map_err(xml_error)?;
    if version.as_ref() != b"1.0" {
        return Err(invalid("XML declaration must use version 1.0"));
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(xml_error)?;
        if !encoding.as_ref().eq_ignore_ascii_case(b"utf-8") {
            return Err(invalid("slide XML declaration must use UTF-8 encoding"));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(xml_error)?;
        if standalone.as_ref() != b"yes" && standalone.as_ref() != b"no" {
            return Err(invalid("XML declaration standalone must be yes or no"));
        }
    }
    Ok(())
}

fn validate_processing_instruction(bytes: &[u8]) -> Result<()> {
    validate_xml_characters(bytes)?;
    let end = bytes
        .iter()
        .position(u8::is_ascii_whitespace)
        .unwrap_or(bytes.len());
    let target = bytes.get(..end).ok_or_else(|| invalid("empty PI target"))?;
    let target = std::str::from_utf8(target).map_err(|error| Error::Xml(error.to_string()))?;
    if !litchi_ooxml_common::xml_name::is_xml_name(target) || target.eq_ignore_ascii_case("xml") {
        return Err(invalid("invalid XML processing-instruction target"));
    }
    Ok(())
}

fn is_xml_char(value: char) -> bool {
    matches!(
        value,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

fn is_xml_whitespace(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn namespace_uri_escaped_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let added = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        length
            .checked_add(added)
            .ok_or_else(|| invalid("namespace URI escaped size overflows"))
    })
}

fn push_namespace_uri(output: &mut Vec<u8>, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
    Ok(())
}

fn qname_prefix(name: &[u8]) -> Result<Vec<u8>> {
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &name[..index]);
    if prefix.len() > MAX_NAME_BYTES {
        return Err(limit("XML namespace prefix bytes", MAX_NAME_BYTES));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(prefix.len())
        .map_err(|source| Error::Allocation {
            resource: "SVG owner element prefix",
            source,
        })?;
    output.extend_from_slice(prefix);
    Ok(output)
}

fn add_count(value: usize, limit_value: usize, resource: &'static str) -> Result<usize> {
    let value = value
        .checked_add(1)
        .ok_or_else(|| limit(resource, limit_value))?;
    if value > limit_value {
        Err(limit(resource, limit_value))
    } else {
        Ok(value)
    }
}

fn position(reader: &Reader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("XML source position exceeds usize"))
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: &str) -> Error {
    Error::Invalid(format!("source-backed SVG owner: {message}"))
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

#[cfg(test)]
mod tests {
    use super::*;

    const FOREIGN_MCE: &[u8] = b"http://purl.oclc.org/ooxml/markup-compatibility/2006";

    #[test]
    fn only_canonical_mce_namespace_classifies_as_mce() {
        let name = QName(b"AlternateContent");
        assert_eq!(classify(Some(MCE), name, None, false, &[]), Kind::Mce);
        assert_eq!(
            classify(Some(FOREIGN_MCE), name, None, false, &[]),
            Kind::Other
        );
    }
}
