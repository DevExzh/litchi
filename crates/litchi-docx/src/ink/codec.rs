#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "the scanner keeps its XML vocabulary and helpers together"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are refined after each bounded validation step"
)]

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, OffsetLimits, active_offsets};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::package::story::StoryDialect;
use crate::{Error, Result};

const TRANSITIONAL_WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIPS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

const TRANSITIONAL_DRAWINGML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const TRANSITIONAL_WORDPROCESSING_DRAWING: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const STRICT_WORDPROCESSING_DRAWING: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const WORD_2010_WORDML: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const WORDPROCESSING_INK: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
const WORDPROCESSING_CANVAS: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const WORDPROCESSING_GROUP: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_MCE_MARKED_BYTES: usize = 128 * 1024 * 1024;
const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1024 * 1024;
const MAX_QNAME_BYTES: usize = 4096;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_COMMENT_BYTES: usize = 1024 * 1024;
const MAX_REFERENCE_BYTES: usize = 4096;
const MAX_DECLARATION_BYTES: usize = 4096;
const MAX_MCE_MARKER_BYTES_PER_ANCHOR: usize = 64;

/// The syntactic host used by one discovered relationship anchor.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Form {
    /// Part 1 §17.3.3.2 `w:contentPart` directly in a run.
    Base,
    /// [MS-ODRAWXML] Word 2010 `w14:contentPart` in DrawingML.
    Drawing,
    /// A Word 2010 canvas/group `w14:contentPart` carrying generic XML.
    GenericDrawing,
}

/// One bounded relationship identity discovered in a story XML part.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Anchor {
    pub relationship_id: String,
    pub form: Form,
}

#[cfg(test)]
#[path = "codec_tests.rs"]
mod tests;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Frame {
    Other,
    Run,
    Drawing,
    Inline,
    Anchor,
    Graphic,
    GraphicDataInk,
    GraphicDataCanvas,
    GraphicDataGroup,
    Canvas,
    Group,
    MceAlternateContent,
    MceChoice,
    MceFallback,
}

#[derive(Clone, Copy, Debug)]
struct MceContext {
    direct_run: bool,
    choice_count: usize,
    fallback_count: usize,
    choice_requires: Option<GraphicDataKind>,
    choice_requires_valid: bool,
    fallback_root_seen: bool,
    fallback_root_valid: bool,
    selected_kind: Option<GraphicDataKind>,
}

#[derive(Clone, Copy, Debug)]
struct PendingAnchor {
    offset: u32,
    form: Form,
}

/// Inventory relationship anchors in one bounded WordprocessingML story.
///
/// This scanner deliberately stops at the story-side grammar. Package
/// integration validates the owning relationship, target part, and target
/// content type/root. A generic base `contentPart` is therefore retained in
/// this inventory even when its target is not InkML.
pub(crate) fn scan(
    xml: &[u8],
    dialect: StoryDialect,
    max_nodes: usize,
    max_depth: usize,
    max_anchors: usize,
) -> Result<Vec<Anchor>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("story XML bytes", xml.len(), MAX_XML_BYTES));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);

    let word_namespace = match dialect {
        StoryDialect::Transitional => TRANSITIONAL_WORD,
        StoryDialect::Strict => STRICT_WORD,
    };
    let relationship_namespace = match dialect {
        StoryDialect::Transitional => TRANSITIONAL_RELATIONSHIPS,
        StoryDialect::Strict => STRICT_RELATIONSHIPS,
    };

    let mut stack = Vec::new();
    let mut pending = Vec::new();
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prolog_started = false;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_position = reader.buffer_position();
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                observe_node(&mut nodes, max_nodes)?;
                observe_start_envelope(
                    &element,
                    &namespace,
                    resolver,
                    &stack,
                    max_depth,
                    &mut root_seen,
                    root_closed,
                    &mut prolog_started,
                )?;
                let parent = semantic_parent(&stack);
                let frame = classify_frame(&namespace, &element, parent, resolver, dialect);
                if is_mce_frame(frame) {
                    prolog_started = true;
                }
                if let Some(form) = classify_anchor(&namespace, &element, word_namespace, parent) {
                    let offset = element_offset(event_position, &element, false, xml)?;
                    push_pending(&mut pending, form, offset, max_nodes)?;
                }
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "DOCX ink XML frames",
                    source,
                })?;
                stack.push(frame);
            },
            Event::Empty(element) => {
                observe_node(&mut nodes, max_nodes)?;
                observe_empty_envelope(
                    &element,
                    &namespace,
                    resolver,
                    &stack,
                    max_depth,
                    &mut root_seen,
                    &mut root_closed,
                    &mut prolog_started,
                )?;
                let parent = semantic_parent(&stack);
                if let Some(form) = classify_anchor(&namespace, &element, word_namespace, parent) {
                    let offset = element_offset(event_position, &element, true, xml)?;
                    push_pending(&mut pending, form, offset, max_nodes)?;
                }
            },
            Event::End(element) => {
                if !root_seen || stack.is_empty() {
                    return Err(invalid("DOCX story has an unexpected end element"));
                }
                validate_qname(element.name().as_ref(), "XML end QName")?;
                if matches!(namespace, ResolveResult::Unknown(_)) {
                    return Err(invalid(
                        "DOCX story end element uses an unknown namespace prefix",
                    ));
                }
                stack.pop();
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                bounded(text.as_ref(), MAX_TEXT_BYTES, "XML text")?;
                let text = super::xml::text(text.as_ref())?;
                if !root_seen {
                    prolog_started = true;
                }
                if !root_seen || root_closed {
                    if !text.as_bytes().iter().all(u8::is_ascii_whitespace) {
                        return Err(invalid(
                            "DOCX story has non-whitespace text outside its root",
                        ));
                    }
                }
            },
            Event::CData(data) => {
                bounded(data.as_ref(), MAX_TEXT_BYTES, "XML CDATA")?;
                super::xml::text(data.as_ref())?;
                if !root_seen || root_closed {
                    return Err(invalid("DOCX story has CDATA outside its root"));
                }
            },
            Event::Comment(comment) => {
                bounded(comment.as_ref(), MAX_COMMENT_BYTES, "XML comment")?;
                super::xml::text(comment.as_ref())?;
                if !root_seen {
                    prolog_started = true;
                }
            },
            Event::GeneralRef(reference) => {
                bounded(reference.as_ref(), MAX_REFERENCE_BYTES, "XML reference")?;
                super::xml::reference(reference.as_ref())?;
                if !root_seen || root_closed {
                    return Err(invalid("DOCX story has a reference outside its root"));
                }
            },
            Event::Decl(declaration) => {
                bounded(
                    declaration.as_ref(),
                    MAX_DECLARATION_BYTES,
                    "XML declaration",
                )?;
                super::xml::declaration(&declaration)?;
                if declaration_seen || prolog_started || root_seen {
                    return Err(invalid(
                        "DOCX story has an XML declaration outside its prolog",
                    ));
                }
                declaration_seen = true;
            },
            Event::DocType(doctype) => {
                bounded(doctype.as_ref(), MAX_DECLARATION_BYTES, "XML DTD")?;
                return Err(invalid("DOCX story rejects DTDs"));
            },
            Event::PI(instruction) => {
                bounded(
                    instruction.as_ref(),
                    MAX_DECLARATION_BYTES,
                    "XML processing instruction",
                )?;
                return Err(invalid("DOCX story rejects processing instructions"));
            },
            Event::Eof => {
                if !root_seen || !root_closed || !stack.is_empty() {
                    return Err(invalid("DOCX story has an unterminated XML root"));
                }
                break;
            },
        }
    }

    let selected_offsets = select_active_offsets(xml, &pending, max_nodes, max_depth)?;
    if selected_offsets.len() > max_anchors {
        return Err(limit("ink anchors", selected_offsets.len(), max_anchors));
    }
    collect_active_anchors(
        xml,
        dialect,
        max_nodes,
        max_depth,
        &pending,
        &selected_offsets,
        relationship_namespace,
    )
}

fn collect_active_anchors(
    xml: &[u8],
    dialect: StoryDialect,
    max_nodes: usize,
    max_depth: usize,
    candidates: &[PendingAnchor],
    selected_offsets: &[u32],
    relationship_namespace: &[u8],
) -> Result<Vec<Anchor>> {
    let mut anchors = Vec::new();
    anchors
        .try_reserve_exact(selected_offsets.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink anchors",
            source,
        })?;
    if selected_offsets.is_empty() {
        return Ok(anchors);
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    let word_namespace = dialect_word_namespace(dialect);
    let mut stack = Vec::new();
    let mut nodes = 0usize;
    let mut candidate_index = 0usize;
    let mut selected_index = 0usize;
    let mut active_base_depth = None;
    let mut mce_contexts = Vec::new();

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_position = reader.buffer_position();
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if active_base_depth.is_some() {
                    return Err(invalid(
                        "DOCX base contentPart has a non-empty child element",
                    ));
                }
                observe_node(&mut nodes, max_nodes)?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("XML depth", usize::MAX, max_depth))?;
                if depth > max_depth {
                    return Err(limit("XML depth", depth, max_depth));
                }
                let parent = semantic_parent(&stack);
                let frame = classify_frame(&namespace, &element, parent, resolver, dialect);
                if frame == Frame::MceAlternateContent {
                    push_mce_context(&mut mce_contexts, &stack)?;
                } else {
                    observe_mce_start(
                        &mut mce_contexts,
                        &stack,
                        frame,
                        &element,
                        resolver,
                        dialect,
                    )?;
                }
                if let Some(form) = classify_anchor(&namespace, &element, word_namespace, parent) {
                    let offset = element_offset(event_position, &element, false, xml)?;
                    let candidate = candidates
                        .get(candidate_index)
                        .ok_or_else(|| invalid("DOCX ink candidate inventory diverged"))?;
                    if candidate.offset != offset || candidate.form != form {
                        return Err(invalid("DOCX ink candidate inventory diverged"));
                    }
                    if selected_offsets.get(selected_index) == Some(&offset) {
                        validate_anchor_placement(form, &stack)?;
                        match form {
                            Form::Drawing => {
                                mark_selected_drawing(&mut mce_contexts, &stack)?;
                            },
                            Form::GenericDrawing => {
                                let kind = generic_drawing_kind(&stack).ok_or_else(|| {
                                    invalid("DOCX generic contentPart has no drawing kind")
                                })?;
                                mark_selected_extension(&mut mce_contexts, &stack, kind)?;
                            },
                            Form::Base => {},
                        }
                        anchors.push(read_anchor(
                            form,
                            &element,
                            resolver,
                            match form {
                                Form::Base => relationship_namespace,
                                Form::Drawing | Form::GenericDrawing => TRANSITIONAL_RELATIONSHIPS,
                            },
                        )?);
                        if form == Form::Base {
                            active_base_depth = Some(depth);
                        }
                        selected_index += 1;
                    }
                    candidate_index += 1;
                }
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "DOCX ink XML frames",
                    source,
                })?;
                stack.push(frame);
            },
            Event::Empty(element) => {
                if active_base_depth.is_some() {
                    return Err(invalid(
                        "DOCX base contentPart has a non-empty child element",
                    ));
                }
                observe_node(&mut nodes, max_nodes)?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("XML depth", usize::MAX, max_depth))?;
                if depth > max_depth {
                    return Err(limit("XML depth", depth, max_depth));
                }
                let parent = semantic_parent(&stack);
                let frame = classify_frame(&namespace, &element, parent, resolver, dialect);
                observe_mce_empty(
                    &mut mce_contexts,
                    &stack,
                    frame,
                    &element,
                    resolver,
                    dialect,
                )?;
                if let Some(form) = classify_anchor(&namespace, &element, word_namespace, parent) {
                    let offset = element_offset(event_position, &element, true, xml)?;
                    let candidate = candidates
                        .get(candidate_index)
                        .ok_or_else(|| invalid("DOCX ink candidate inventory diverged"))?;
                    if candidate.offset != offset || candidate.form != form {
                        return Err(invalid("DOCX ink candidate inventory diverged"));
                    }
                    if selected_offsets.get(selected_index) == Some(&offset) {
                        validate_anchor_placement(form, &stack)?;
                        match form {
                            Form::Drawing => {
                                mark_selected_drawing(&mut mce_contexts, &stack)?;
                            },
                            Form::GenericDrawing => {
                                let kind = generic_drawing_kind(&stack).ok_or_else(|| {
                                    invalid("DOCX generic contentPart has no drawing kind")
                                })?;
                                mark_selected_extension(&mut mce_contexts, &stack, kind)?;
                            },
                            Form::Base => {},
                        }
                        anchors.push(read_anchor(
                            form,
                            &element,
                            resolver,
                            match form {
                                Form::Base => relationship_namespace,
                                Form::Drawing | Form::GenericDrawing => TRANSITIONAL_RELATIONSHIPS,
                            },
                        )?);
                        selected_index += 1;
                    }
                    candidate_index += 1;
                }
            },
            Event::End(_) => {
                if let Some(base_depth) = active_base_depth {
                    if stack.len() < base_depth {
                        return Err(invalid("DOCX base contentPart has an invalid end"));
                    }
                    if stack.len() == base_depth {
                        active_base_depth = None;
                    }
                }
                if stack.last() == Some(&Frame::MceAlternateContent) {
                    let context = mce_contexts
                        .pop()
                        .ok_or_else(|| invalid("DOCX MCE context inventory diverged"))?;
                    validate_mce_context(context)?;
                }
                if stack.pop().is_none() {
                    return Err(invalid("DOCX ink active scan has an unexpected end"));
                }
            },
            Event::Text(text) => {
                if active_base_depth.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "DOCX base contentPart has non-whitespace character content",
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                if active_base_depth.is_some() {
                    return Err(invalid(
                        "DOCX base contentPart has non-whitespace character content",
                    ));
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if candidate_index != candidates.len() || selected_index != selected_offsets.len() {
        return Err(invalid("DOCX ink candidate inventory diverged"));
    }
    Ok(anchors)
}

fn validate_anchor_placement(form: Form, stack: &[Frame]) -> Result<()> {
    let valid = match form {
        Form::Base => semantic_parent(stack) == Some(Frame::Run),
        Form::Drawing => drawing_ink_suffix(stack),
        Form::GenericDrawing => generic_drawing_suffix(stack),
    };
    valid
        .then_some(())
        .ok_or_else(|| invalid("DOCX ink contentPart is outside its legal host placement"))
}

fn push_mce_context(contexts: &mut Vec<MceContext>, stack: &[Frame]) -> Result<()> {
    contexts
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink MCE contexts",
            source,
        })?;
    contexts.push(MceContext {
        direct_run: stack.last() == Some(&Frame::Run),
        choice_count: 0,
        fallback_count: 0,
        choice_requires: None,
        choice_requires_valid: false,
        fallback_root_seen: false,
        fallback_root_valid: false,
        selected_kind: None,
    });
    Ok(())
}

fn observe_mce_start(
    contexts: &mut [MceContext],
    stack: &[Frame],
    frame: Frame,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    dialect: StoryDialect,
) -> Result<()> {
    observe_mce_child(contexts, stack, frame, element, resolver, dialect)
}

fn observe_mce_empty(
    contexts: &mut [MceContext],
    stack: &[Frame],
    frame: Frame,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    dialect: StoryDialect,
) -> Result<()> {
    observe_mce_child(contexts, stack, frame, element, resolver, dialect)
}

fn observe_mce_child(
    contexts: &mut [MceContext],
    stack: &[Frame],
    frame: Frame,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    dialect: StoryDialect,
) -> Result<()> {
    let Some(context) = contexts.last_mut() else {
        return Ok(());
    };
    if stack.last() == Some(&Frame::MceAlternateContent) {
        match frame {
            Frame::MceChoice => {
                context.choice_count = context
                    .choice_count
                    .checked_add(1)
                    .ok_or_else(|| limit("MCE choices", usize::MAX, usize::MAX))?;
                let requires = choice_requires_kind(element, resolver);
                if context.choice_count == 1 {
                    context.choice_requires = requires;
                    context.choice_requires_valid = requires.is_some();
                } else {
                    context.choice_requires_valid = false;
                }
            },
            Frame::MceFallback => {
                context.fallback_count = context
                    .fallback_count
                    .checked_add(1)
                    .ok_or_else(|| limit("MCE fallbacks", usize::MAX, usize::MAX))?;
            },
            _ => {},
        }
    }
    if stack.last() == Some(&Frame::MceFallback) {
        if context.fallback_root_seen {
            context.fallback_root_valid = false;
        } else {
            context.fallback_root_seen = true;
            let (namespace, local) = resolver.resolve_element(element.name());
            context.fallback_root_valid = is_namespace(&namespace, dialect_word_namespace(dialect))
                && local.as_ref() == b"pict";
        }
    }
    Ok(())
}

fn mark_selected_drawing(contexts: &mut [MceContext], stack: &[Frame]) -> Result<()> {
    mark_selected_extension(contexts, stack, GraphicDataKind::Ink)
}

fn mark_selected_extension(
    contexts: &mut [MceContext],
    stack: &[Frame],
    kind: GraphicDataKind,
) -> Result<()> {
    if contexts.len() != 1 {
        return Err(invalid(
            "DOCX drawing contentPart requires one direct AlternateContent host",
        ));
    }
    let mut branch = None;
    for frame in stack.iter().rev() {
        match frame {
            Frame::MceChoice => {
                branch = Some(Frame::MceChoice);
                break;
            },
            Frame::MceFallback => {
                branch = Some(Frame::MceFallback);
                break;
            },
            Frame::MceAlternateContent => break,
            _ => {},
        }
    }
    if branch != Some(Frame::MceChoice) {
        return Err(invalid(
            "DOCX drawing contentPart must be in the active MCE Choice",
        ));
    }
    let context = contexts
        .last_mut()
        .ok_or_else(|| invalid("DOCX ink MCE context inventory diverged"))?;
    if context
        .selected_kind
        .is_some_and(|selected| selected != kind)
    {
        return Err(invalid(
            "DOCX drawing contentPart has multiple active host branches",
        ));
    }
    context.selected_kind = Some(kind);
    Ok(())
}

fn validate_mce_context(context: MceContext) -> Result<()> {
    if context.selected_kind.is_none() {
        return Ok(());
    }
    if !context.direct_run
        || context.choice_count != 1
        || context.fallback_count != 1
        || !context.choice_requires_valid
        || context.choice_requires != context.selected_kind
        || !context.fallback_root_seen
        || !context.fallback_root_valid
    {
        return Err(invalid(
            "DOCX drawing Ink AlternateContent does not match its normative Choice/Fallback host",
        ));
    }
    Ok(())
}

fn choice_requires_kind(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Option<GraphicDataKind> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.ok()?;
        if attribute.key.as_ref() != b"Requires" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) || value.is_some() {
            return None;
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
                .ok()?,
        );
    }
    let value = value?;
    let mut tokens = value.split_whitespace();
    let token = tokens.next()?;
    if tokens.next().is_some() {
        return None;
    }
    let prefix = token.as_bytes();
    if prefix.is_empty() || prefix.len() + 2 > MAX_QNAME_BYTES {
        return None;
    }
    let mut qualified = [0u8; MAX_QNAME_BYTES];
    qualified[..prefix.len()].copy_from_slice(prefix);
    qualified[prefix.len()] = b':';
    qualified[prefix.len() + 1] = b'x';
    let (namespace, _) =
        resolver.resolve_element(quick_xml::name::QName(&qualified[..prefix.len() + 2]));
    match namespace {
        ResolveResult::Bound(Namespace(value)) if value == WORDPROCESSING_INK => {
            Some(GraphicDataKind::Ink)
        },
        ResolveResult::Bound(Namespace(value)) if value == WORDPROCESSING_CANVAS => {
            Some(GraphicDataKind::Canvas)
        },
        ResolveResult::Bound(Namespace(value)) if value == WORDPROCESSING_GROUP => {
            Some(GraphicDataKind::Group)
        },
        _ => None,
    }
}

fn drawing_ink_suffix(stack: &[Frame]) -> bool {
    let mut frames = stack
        .iter()
        .rev()
        .copied()
        .filter(|frame| !is_mce_frame(*frame));
    matches!(frames.next(), Some(Frame::GraphicDataInk))
        && matches!(frames.next(), Some(Frame::Graphic))
        && matches!(frames.next(), Some(Frame::Inline | Frame::Anchor))
        && matches!(frames.next(), Some(Frame::Drawing))
        && matches!(frames.next(), Some(Frame::Run))
}

fn generic_drawing_suffix(stack: &[Frame]) -> bool {
    let mut frames = stack
        .iter()
        .rev()
        .copied()
        .filter(|frame| !is_mce_frame(*frame));
    let terminal = frames.next();
    if terminal == Some(Frame::Canvas) {
        return semantic_suffix_matches(
            stack,
            &[
                Frame::Run,
                Frame::Drawing,
                Frame::Inline,
                Frame::Graphic,
                Frame::GraphicDataCanvas,
                Frame::Canvas,
            ],
        );
    }
    if terminal != Some(Frame::Group) {
        return false;
    }

    // A group can contain another group shape, so consume one or more Group
    // frames before checking the fixed outer DrawingML path.
    let mut reverse = stack
        .iter()
        .rev()
        .copied()
        .filter(|frame| !is_mce_frame(*frame));
    let mut groups = 0usize;
    loop {
        match reverse.next() {
            Some(Frame::Group) => groups = groups.saturating_add(1),
            Some(Frame::GraphicDataGroup) => break,
            Some(Frame::Canvas) => {
                if !matches!(reverse.next(), Some(Frame::GraphicDataCanvas)) {
                    return false;
                }
                break;
            },
            _ => return false,
        }
    }
    (groups > 0)
        && matches!(reverse.next(), Some(Frame::Graphic))
        && matches!(reverse.next(), Some(Frame::Inline | Frame::Anchor))
        && matches!(reverse.next(), Some(Frame::Drawing))
        && matches!(reverse.next(), Some(Frame::Run))
}

fn generic_drawing_kind(stack: &[Frame]) -> Option<GraphicDataKind> {
    stack.iter().rev().find_map(|frame| match frame {
        Frame::GraphicDataCanvas => Some(GraphicDataKind::Canvas),
        Frame::GraphicDataGroup => Some(GraphicDataKind::Group),
        _ => None,
    })
}

fn semantic_suffix_matches(stack: &[Frame], expected: &[Frame]) -> bool {
    let mut frames = stack
        .iter()
        .rev()
        .copied()
        .filter(|frame| !is_mce_frame(*frame));
    expected.iter().rev().all(|wanted| {
        let actual = frames.next();
        actual == Some(*wanted) || (*wanted == Frame::Inline && actual == Some(Frame::Anchor))
    })
}

fn read_anchor(
    form: Form,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    relationship_namespace: &[u8],
) -> Result<Anchor> {
    let mut relationship_id = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_namespace(&namespace, relationship_namespace) {
            continue;
        }
        if relationship_id.is_some() {
            return Err(invalid(
                "DOCX ink contentPart has duplicate relationship IDs",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(invalid(
                "DOCX ink contentPart has an invalid relationship ID",
            ));
        }
        if value
            .chars()
            .any(|character| character.is_control() || character.is_whitespace())
        {
            return Err(invalid(
                "DOCX ink contentPart relationship ID contains whitespace",
            ));
        }
        let mut owned = String::new();
        owned
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX ink relationship ID",
                source,
            })?;
        owned.push_str(&value);
        relationship_id = Some(owned);
    }
    let relationship_id = relationship_id.ok_or_else(|| {
        invalid("DOCX ink contentPart is missing a bound dialect relationship ID")
    })?;
    Ok(Anchor {
        relationship_id,
        form,
    })
}

fn observe_start_envelope(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    stack: &[Frame],
    max_depth: usize,
    root_seen: &mut bool,
    root_closed: bool,
    prolog_started: &mut bool,
) -> Result<()> {
    validate_element_envelope(element, namespace, resolver)?;
    let depth = stack
        .len()
        .checked_add(1)
        .ok_or_else(|| limit("XML depth", usize::MAX, max_depth))?;
    if depth > max_depth {
        return Err(limit("XML depth", depth, max_depth));
    }
    if *root_seen && root_closed {
        return Err(invalid("DOCX story has more than one XML root"));
    }
    if stack.is_empty() {
        if *root_seen {
            return Err(invalid("DOCX story has more than one XML root"));
        }
        *root_seen = true;
    }
    *prolog_started = true;
    Ok(())
}

fn observe_empty_envelope(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    stack: &[Frame],
    max_depth: usize,
    root_seen: &mut bool,
    root_closed: &mut bool,
    prolog_started: &mut bool,
) -> Result<()> {
    validate_element_envelope(element, namespace, resolver)?;
    let depth = stack
        .len()
        .checked_add(1)
        .ok_or_else(|| limit("XML depth", usize::MAX, max_depth))?;
    if depth > max_depth {
        return Err(limit("XML depth", depth, max_depth));
    }
    if stack.is_empty() {
        if *root_seen {
            return Err(invalid("DOCX story has more than one XML root"));
        }
        *root_seen = true;
        *root_closed = true;
    } else if *root_closed {
        return Err(invalid("DOCX story has markup after its XML root"));
    }
    *prolog_started = true;
    Ok(())
}

fn validate_element_envelope(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
) -> Result<()> {
    super::xml::element(element, resolver)?;
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid(
            "DOCX story element uses an unknown namespace prefix",
        ));
    }
    Ok(())
}

fn classify_frame(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    parent: Option<Frame>,
    resolver: &NamespaceResolver,
    dialect: StoryDialect,
) -> Frame {
    let local = element.local_name();
    if is_namespace(namespace, MCE_NAMESPACE) {
        return match local.as_ref() {
            b"AlternateContent" => Frame::MceAlternateContent,
            b"Choice" => Frame::MceChoice,
            b"Fallback" => Frame::MceFallback,
            _ => Frame::Other,
        };
    }
    if is_namespace(namespace, dialect_word_namespace(dialect)) {
        if local.as_ref() == b"r" {
            return Frame::Run;
        }
        if local.as_ref() == b"drawing" {
            return Frame::Drawing;
        }
    }
    if is_wordprocessing_drawing_namespace(namespace) {
        if local.as_ref() == b"inline" && matches!(parent, Some(Frame::Drawing)) {
            return Frame::Inline;
        }
        if local.as_ref() == b"anchor" && matches!(parent, Some(Frame::Drawing)) {
            return Frame::Anchor;
        }
    }
    if is_drawing_namespace(namespace) && local.as_ref() == b"graphic" {
        if matches!(parent, Some(Frame::Inline | Frame::Anchor)) {
            return Frame::Graphic;
        }
    }
    if is_drawing_namespace(namespace) && local.as_ref() == b"graphicData" {
        return match graphic_data_kind(element, resolver) {
            Some(GraphicDataKind::Ink) => Frame::GraphicDataInk,
            Some(GraphicDataKind::Canvas) => Frame::GraphicDataCanvas,
            Some(GraphicDataKind::Group) => Frame::GraphicDataGroup,
            _ => Frame::Other,
        };
    }
    if is_namespace(namespace, WORDPROCESSING_CANVAS)
        && local.as_ref() == b"wpc"
        && matches!(parent, Some(Frame::GraphicDataCanvas))
    {
        return Frame::Canvas;
    }
    if is_namespace(namespace, WORDPROCESSING_GROUP)
        && matches!(local.as_ref(), b"wgp" | b"grpSp")
        && matches!(
            parent,
            Some(Frame::GraphicDataGroup | Frame::Canvas | Frame::Group)
        )
    {
        return Frame::Group;
    }
    Frame::Other
}

fn classify_anchor(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    word_namespace: &[u8],
    parent: Option<Frame>,
) -> Option<Form> {
    let local = element.local_name();
    if local.as_ref() != b"contentPart" {
        return None;
    }
    if is_namespace(namespace, word_namespace) {
        return Some(Form::Base);
    }
    if is_namespace(namespace, WORD_2010_WORDML) {
        if matches!(parent, Some(Frame::Canvas | Frame::Group)) {
            return Some(Form::GenericDrawing);
        }
        return Some(Form::Drawing);
    }
    None
}

fn semantic_parent(stack: &[Frame]) -> Option<Frame> {
    stack
        .iter()
        .rev()
        .copied()
        .find(|frame| !is_mce_frame(*frame))
}

fn is_mce_frame(frame: Frame) -> bool {
    matches!(
        frame,
        Frame::MceAlternateContent | Frame::MceChoice | Frame::MceFallback
    )
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum GraphicDataKind {
    Ink,
    Canvas,
    Group,
}

fn graphic_data_kind(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Option<GraphicDataKind> {
    let mut uri = None;
    for attribute in element.attributes() {
        let attribute = attribute.ok()?;
        if attribute.key.as_ref() != b"uri" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            continue;
        }
        if uri.is_some() {
            return None;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .ok()?;
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return None;
        }
        uri = Some(match value.as_ref().as_bytes() {
            value if value == WORDPROCESSING_INK => GraphicDataKind::Ink,
            value if value == WORDPROCESSING_CANVAS => GraphicDataKind::Canvas,
            value if value == WORDPROCESSING_GROUP => GraphicDataKind::Group,
            _ => return None,
        });
    }
    uri
}

fn push_pending(
    pending: &mut Vec<PendingAnchor>,
    form: Form,
    offset: u32,
    max_candidates: usize,
) -> Result<()> {
    if pending.len() >= max_candidates {
        return Err(limit(
            "ink candidate anchors",
            pending.len().saturating_add(1),
            max_candidates,
        ));
    }
    pending.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "DOCX ink pending anchors",
        source,
    })?;
    pending.push(PendingAnchor { offset, form });
    Ok(())
}

fn select_active_offsets(
    xml: &[u8],
    pending: &[PendingAnchor],
    max_nodes: usize,
    max_depth: usize,
) -> Result<Vec<u32>> {
    let mut offsets = Vec::new();
    offsets
        .try_reserve_exact(pending.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX ink anchor offsets",
            source,
        })?;
    offsets.extend(pending.iter().map(|candidate| candidate.offset));
    if offsets.is_empty() {
        return Ok(offsets);
    }
    let marked_extra = pending
        .len()
        .checked_mul(MAX_MCE_MARKER_BYTES_PER_ANCHOR)
        .ok_or_else(|| limit("MCE marked XML bytes", usize::MAX, MAX_MCE_MARKED_BYTES))?;
    let marked_bytes = xml
        .len()
        .checked_add(marked_extra)
        .ok_or_else(|| limit("MCE marked XML bytes", usize::MAX, MAX_MCE_MARKED_BYTES))?;
    if marked_bytes > MAX_MCE_MARKED_BYTES {
        return Err(limit(
            "MCE marked XML bytes",
            marked_bytes,
            MAX_MCE_MARKED_BYTES,
        ));
    }
    let mut capabilities = Capabilities::default();
    for namespace in [
        WORD_2010_WORDML,
        WORDPROCESSING_INK,
        WORDPROCESSING_CANVAS,
        WORDPROCESSING_GROUP,
    ] {
        let namespace = std::str::from_utf8(namespace)
            .map_err(|error| Error::Invalid(format!("invalid fixed MCE namespace: {error}")))?;
        capabilities.understand_namespace(namespace);
    }
    let processing = MceLimits {
        max_input_bytes: marked_bytes,
        max_output_bytes: mce_output_limit(marked_bytes, max_nodes, max_depth)?,
        max_depth,
        max_namespace_bindings: MAX_NAMESPACE_DECLARATIONS,
        max_directive_tokens: max_nodes,
        max_choices_per_alternate: max_nodes.max(1),
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    };
    let limits = OffsetLimits {
        max_source_bytes: MAX_XML_BYTES,
        max_offsets: pending.len(),
        max_marked_bytes: MAX_MCE_MARKED_BYTES,
        processing,
    };
    active_offsets(xml, &offsets, &capabilities, &limits).map_err(Error::from)
}

fn mce_output_limit(marked_bytes: usize, max_nodes: usize, max_depth: usize) -> Result<usize> {
    const MAX_OUTPUT_BYTES: usize = 128 * 1024 * 1024;
    const MIN_GROWTH_FACTOR: usize = 8;
    // The MCE rewriter may repeat effective ancestor namespace declarations on
    // every emitted descendant. Charge a checked source-size allowance based
    // on the caller's independent node/depth bounds, while retaining one hard
    // output ceiling for hostile namespace closure. The minimum covers the
    // marker and the small, shallow Word fixtures without assuming source+64.
    let growth_factor = max_nodes.max(max_depth).max(MIN_GROWTH_FACTOR);
    let expanded = marked_bytes
        .checked_mul(growth_factor)
        .ok_or_else(|| limit("MCE output bytes", usize::MAX, MAX_OUTPUT_BYTES))?;
    Ok(expanded.min(MAX_OUTPUT_BYTES))
}

/// The offset of the start tag that ended at reader position
/// `event_position`; `xml` is the reader's input, whose [`ReaderOrigin`]
/// converts the position to a byte offset.
fn element_offset(
    event_position: u64,
    element: &BytesStart<'_>,
    empty: bool,
    xml: &[u8],
) -> Result<u32> {
    let end = ReaderOrigin::of(xml)
        .offset(event_position)
        .ok_or_else(|| invalid("XML position does not fit usize"))?;
    let suffix = if empty { 3 } else { 2 };
    let consumed = element
        .as_ref()
        .len()
        .checked_add(suffix)
        .ok_or_else(|| invalid("DOCX story element position overflowed"))?;
    let start = end
        .checked_sub(consumed)
        .ok_or_else(|| invalid("DOCX story element position underflowed"))?;
    if xml.get(start).copied() != Some(b'<') {
        return Err(invalid("DOCX story element position is not an opening tag"));
    }
    u32::try_from(start)
        .map_err(|error| Error::Invalid(format!("DOCX story position exceeds u32: {error}")))
}

fn observe_node(nodes: &mut usize, maximum: usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("XML nodes", usize::MAX, maximum))?;
    if *nodes > maximum {
        return Err(limit("XML nodes", *nodes, maximum));
    }
    Ok(())
}

fn validate_qname(name: &[u8], resource: &'static str) -> Result<()> {
    if name.is_empty() || name.len() > MAX_QNAME_BYTES {
        return Err(limit(resource, name.len(), MAX_QNAME_BYTES));
    }
    let mut colons = 0usize;
    for byte in name {
        if *byte == b':' {
            colons = colons
                .checked_add(1)
                .ok_or_else(|| limit(resource, usize::MAX, MAX_QNAME_BYTES))?;
        }
        if byte.is_ascii_whitespace() || matches!(*byte, b'<' | b'>' | b'&' | b'"' | b'\'' | b'=') {
            return Err(invalid("DOCX story contains an invalid XML QName"));
        }
    }
    if colons > 1 || name.first() == Some(&b':') || name.last() == Some(&b':') {
        return Err(invalid("DOCX story contains an invalid XML QName"));
    }
    Ok(())
}

fn bounded(bytes: &[u8], maximum: usize, resource: &'static str) -> Result<()> {
    if bytes.len() > maximum {
        return Err(limit(resource, bytes.len(), maximum));
    }
    Ok(())
}

fn is_drawing_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace(namespace, TRANSITIONAL_DRAWINGML) || is_namespace(namespace, STRICT_DRAWINGML)
}

fn is_wordprocessing_drawing_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace(namespace, TRANSITIONAL_WORDPROCESSING_DRAWING)
        || is_namespace(namespace, STRICT_WORDPROCESSING_DRAWING)
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

const fn dialect_word_namespace(dialect: StoryDialect) -> &'static [u8] {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_WORD,
        StoryDialect::Strict => STRICT_WORD,
    }
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InkLimit {
        resource,
        actual,
        maximum,
    }
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}
