//! Lossless, namespace-aware discovery of the PresentationML InkAction
//! extension.  This scanner deliberately keeps the owning XML untouched: the
//! package transaction replaces only an already classified target part.

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::events::{BytesDecl, BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{AnchorFingerprint, Branch, Candidate, Dialect};
use crate::presentation::embedded::{MAX_XML_DEPTH, increment_nodes, invalid, limit};
use crate::{Error, Result};
use litchi_ooxml_common::xml_name::{is_ncname, is_qualified_name};

pub(crate) const MC: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
pub(crate) const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const REL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
pub(crate) const ACTION: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2014/inkAction";
const UNKNOWN_REQUIRES: &[u8] = b"__ink_action_unresolved_requires__";

pub(crate) const REQUIRED_CONTENT_TYPE: &str = "text/xml";
pub(crate) const MAX_ATTRIBUTES: usize = 256;
pub(crate) const MAX_NAMESPACE_DECLARATIONS: usize = 256;
pub(crate) const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ElementNamespace {
    Mc,
    PmlTransitional,
    PmlStrict,
    Other,
    Unbound,
    Unknown,
}

impl ElementNamespace {
    fn from_resolved(value: &ResolveResult<'_>) -> Self {
        match value {
            ResolveResult::Bound(Namespace(uri)) if namespace_matches(uri, MC) => Self::Mc,
            ResolveResult::Bound(Namespace(uri)) if namespace_matches(uri, PML) => {
                Self::PmlTransitional
            },
            ResolveResult::Bound(Namespace(uri)) if namespace_matches(uri, STRICT_PML) => {
                Self::PmlStrict
            },
            ResolveResult::Bound(_) => Self::Other,
            ResolveResult::Unbound => Self::Unbound,
            ResolveResult::Unknown(_) => Self::Unknown,
        }
    }

    pub(crate) fn is_pml(self) -> bool {
        matches!(self, Self::PmlTransitional | Self::PmlStrict)
    }

    fn is_mc(self) -> bool {
        matches!(self, Self::Mc)
    }

    fn dialect(self) -> Option<Dialect> {
        match self {
            Self::PmlTransitional => Some(Dialect::Transitional),
            Self::PmlStrict => Some(Dialect::Strict),
            _ => None,
        }
    }
}

fn namespace_matches(actual: &[u8], expected: &[u8]) -> bool {
    if actual == expected {
        return true;
    }
    if actual.len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
        return false;
    }
    let Ok(actual) = std::str::from_utf8(actual) else {
        return false;
    };
    quick_xml::escape::unescape(actual).is_ok_and(|decoded| decoded.as_bytes() == expected)
}

fn normalize_namespace(actual: &[u8]) -> Result<Vec<u8>> {
    if actual.len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
        return Err(limit(
            "ink-action namespace URI bytes",
            crate::presentation::embedded::MAX_ATTRIBUTE_BYTES,
        ));
    }
    if !actual.contains(&b'&') {
        return Ok(actual.to_vec());
    }
    let actual = std::str::from_utf8(actual)
        .map_err(|_| invalid("ink-action namespace URI is not UTF-8"))?;
    quick_xml::escape::unescape(actual)
        .map(|decoded| decoded.into_owned().into_bytes())
        .map_err(|error| Error::Xml(error.to_string()))
}

#[derive(Debug, Clone)]
struct ElementFrame {
    start: usize,
    depth: usize,
    namespace: ElementNamespace,
    local: Vec<u8>,
    alternate: Option<AlternateState>,
}

#[derive(Debug, Clone)]
struct AlternateState {
    start: usize,
    depth: usize,
    choices: Vec<BranchState>,
    fallback: Option<BranchState>,
}

#[derive(Debug, Clone)]
struct BranchState {
    branch: Branch,
    start: usize,
    depth: usize,
    end: Option<usize>,
    requires: Option<Vec<Vec<u8>>>,
    element_children: usize,
    picture_children: usize,
    content_start: Option<usize>,
    content_end: Option<usize>,
    relationship_attributes: Vec<RelationshipAttribute>,
}

#[derive(Debug, Clone)]
struct RelationshipAttribute {
    namespace: Vec<u8>,
    value: Vec<u8>,
}

impl BranchState {
    fn new(branch: Branch, start: usize, depth: usize, requires: Option<Vec<Vec<u8>>>) -> Self {
        Self {
            branch,
            start,
            depth,
            end: None,
            requires,
            element_children: 0,
            picture_children: 0,
            content_start: None,
            content_end: None,
            relationship_attributes: Vec::new(),
        }
    }

    fn action_capable(&self) -> bool {
        self.requires.as_ref().is_some_and(|requires| {
            requires
                .iter()
                .any(|uri| uri.as_slice() == ACTION || uri.as_slice() == UNKNOWN_REQUIRES)
        })
    }
}

/// Scan one slide and retain the complete source spans needed to classify
/// action anchors. Unsupported or generic content-part MCE branches are
/// ignored as opaque source; a branch that advertises the exact action
/// capability but has an invalid closure is reported as a typed error.
pub(crate) fn scan_slide(xml: &[u8], maximum: usize) -> Result<Vec<Candidate>> {
    if xml.len() > crate::presentation::embedded::MAX_XML_BYTES {
        return Err(limit(
            "ink-action owner XML bytes",
            crate::presentation::embedded::MAX_XML_BYTES,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;

    let mut stack = Vec::<ElementFrame>::new();
    let mut candidates = Vec::new();
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut slide_dialect = None;
    let mut raw_ordinal = 0usize;
    let mut declaration_seen = false;
    let mut non_declaration_event_seen = false;

    loop {
        let before = position(&reader, origin)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let after = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (resolved, event) = resolver.resolve_event(event);
        if !matches!(&event, Event::Decl(_) | Event::Eof) {
            non_declaration_event_seen = true;
        }

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                validate_element_name(element.name())?;
                validate_attributes(&element, &resolver, reader.decoder())?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("ink-action owner XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("ink-action owner XML depth", MAX_XML_DEPTH));
                }
                let namespace = ElementNamespace::from_resolved(&resolved);
                if namespace == ElementNamespace::Unknown {
                    return Err(invalid(
                        "ink-action owner element uses an undeclared namespace prefix",
                    ));
                }
                let local = element.name().local_name().as_ref().to_vec();
                if depth == 1 {
                    if root_seen || root_closed {
                        return Err(invalid("ink-action slide has multiple roots"));
                    }
                    if !namespace.is_pml() || local.as_slice() != b"sld" {
                        return Err(invalid("ink-action owner must have one p:sld root"));
                    }
                    slide_dialect = namespace.dialect();
                    root_seen = true;
                }
                if let Some(dialect) = slide_dialect {
                    validate_dialect(namespace, dialect)?;
                }

                let parent_is_shape_tree = stack.last().is_some_and(|frame| {
                    frame.namespace.is_pml()
                        && matches!(frame.local.as_slice(), b"spTree" | b"grpSp")
                });
                let alternate = if namespace.is_mc()
                    && local.as_slice() == b"AlternateContent"
                    && parent_is_shape_tree
                {
                    Some(AlternateState {
                        start: before,
                        depth,
                        choices: Vec::new(),
                        fallback: None,
                    })
                } else {
                    None
                };

                if let Some(alt_index) = stack.iter().rposition(|frame| frame.alternate.is_some()) {
                    let parent_depth = depth.saturating_sub(1);
                    let alt = stack
                        .get_mut(alt_index)
                        .and_then(|frame| frame.alternate.as_mut())
                        .ok_or_else(|| invalid("ink-action AlternateContent state is missing"))?;
                    if alt.depth.saturating_add(1) == depth {
                        if namespace.is_mc() && local.as_slice() == b"Choice" {
                            let requires = requires_value(&element, &resolver, reader.decoder())?;
                            if alt.fallback.is_some() {
                                return Err(invalid(
                                    "ink-action AlternateContent has Choice after Fallback",
                                ));
                            }
                            alt.choices
                                .try_reserve(1)
                                .map_err(|source| Error::Allocation {
                                    resource: "ink-action MCE choices",
                                    source,
                                })?;
                            alt.choices.push(BranchState::new(
                                Branch::Choice,
                                before,
                                depth,
                                Some(requires),
                            ));
                        } else if namespace.is_mc() && local.as_slice() == b"Fallback" {
                            if alt.fallback.is_some() {
                                return Err(invalid(
                                    "ink-action AlternateContent has duplicate Fallback branches",
                                ));
                            }
                            alt.fallback =
                                Some(BranchState::new(Branch::Fallback, before, depth, None));
                        } else if parent_depth == alt.depth {
                            return Err(invalid(
                                "ink-action AlternateContent has an unexpected direct child",
                            ));
                        }
                    }

                    if let Some(branch) = direct_branch_mut(alt, depth) {
                        branch.element_children = branch
                            .element_children
                            .checked_add(1)
                            .ok_or_else(|| limit("ink-action MCE branch elements", maximum))?;
                        let is_content = namespace.is_pml() && local.as_slice() == b"contentPart";
                        let is_picture = namespace.is_pml() && local.as_slice() == b"pic";
                        match branch.branch {
                            Branch::Choice if is_content => {
                                if branch.content_start.is_some() {
                                    return Err(invalid(
                                        "ink-action Choice has duplicate p:contentPart children",
                                    ));
                                }
                                branch.content_start = Some(before);
                                capture_relationship_attributes(
                                    &mut branch.relationship_attributes,
                                    &element,
                                    &resolver,
                                )?;
                            },
                            Branch::Fallback if is_picture => {
                                branch.picture_children =
                                    branch.picture_children.checked_add(1).ok_or_else(|| {
                                        limit("ink-action fallback picture children", maximum)
                                    })?;
                                // The complete fallback source is retained.  Its exact
                                // p:pic shape is checked when the branch closes.
                            },
                            _ => {},
                        }
                    }
                }

                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "ink-action XML element frames",
                    source,
                })?;
                stack.push(ElementFrame {
                    start: before,
                    depth,
                    namespace,
                    local,
                    alternate,
                });
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                validate_element_name(element.name())?;
                validate_attributes(&element, &resolver, reader.decoder())?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("ink-action owner XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("ink-action owner XML depth", MAX_XML_DEPTH));
                }
                let namespace = ElementNamespace::from_resolved(&resolved);
                if namespace == ElementNamespace::Unknown {
                    return Err(invalid(
                        "ink-action owner element uses an undeclared namespace prefix",
                    ));
                }
                let local = element.name().local_name();
                if depth == 1 {
                    if root_seen || root_closed {
                        return Err(invalid("ink-action slide has multiple roots"));
                    }
                    if !namespace.is_pml() || local.as_ref() != b"sld" {
                        return Err(invalid("ink-action owner must have one p:sld root"));
                    }
                    slide_dialect = namespace.dialect();
                    root_seen = true;
                    root_closed = true;
                }
                if let Some(dialect) = slide_dialect {
                    validate_dialect(namespace, dialect)?;
                }
                if let Some(alt_index) = stack.iter().rposition(|frame| frame.alternate.is_some()) {
                    let alt = stack
                        .get_mut(alt_index)
                        .and_then(|frame| frame.alternate.as_mut())
                        .ok_or_else(|| invalid("ink-action AlternateContent state is missing"))?;
                    if let Some(branch) = direct_branch_mut(alt, depth) {
                        branch.element_children = branch
                            .element_children
                            .checked_add(1)
                            .ok_or_else(|| limit("ink-action MCE branch elements", maximum))?;
                        if branch.branch == Branch::Choice
                            && namespace.is_pml()
                            && local.as_ref() == b"contentPart"
                        {
                            if branch.content_start.is_some() {
                                return Err(invalid(
                                    "ink-action Choice has duplicate p:contentPart children",
                                ));
                            }
                            branch.content_start = Some(before);
                            branch.content_end = Some(after);
                            capture_relationship_attributes(
                                &mut branch.relationship_attributes,
                                &element,
                                &resolver,
                            )?;
                        } else if branch.branch == Branch::Fallback
                            && namespace.is_pml()
                            && local.as_ref() == b"pic"
                        {
                            branch.picture_children =
                                branch.picture_children.checked_add(1).ok_or_else(|| {
                                    limit("ink-action fallback picture children", maximum)
                                })?;
                        }
                    }
                }
            },
            Event::End(element) => {
                validate_end_name(element.name())?;
                let namespace = ElementNamespace::from_resolved(&resolved);
                if namespace == ElementNamespace::Unknown {
                    return Err(invalid(
                        "ink-action owner closing element uses an undeclared namespace prefix",
                    ));
                }
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("ink-action slide has an unmatched end element"))?;
                if frame.depth != stack.len().saturating_add(1)
                    || frame.local != element.name().local_name().as_ref()
                    || frame.namespace != namespace
                {
                    return Err(invalid("ink-action slide XML nesting is inconsistent"));
                }
                if let Some(alt_index) = stack.iter().rposition(|item| item.alternate.is_some()) {
                    let alt = stack
                        .get_mut(alt_index)
                        .and_then(|item| item.alternate.as_mut())
                        .ok_or_else(|| invalid("ink-action AlternateContent state is missing"))?;
                    for branch in alt.choices.iter_mut().chain(alt.fallback.iter_mut()) {
                        if branch.content_start == Some(frame.start) {
                            branch.content_end = Some(after);
                        }
                        if branch.start == frame.start {
                            branch.end = Some(after);
                        }
                    }
                }
                if frame.depth == 1 {
                    root_closed = true;
                }
                if let Some(mut alternate) = frame.alternate {
                    let span = alternate.start..after;
                    if let Some(candidate) = finalize_alternate(
                        &mut alternate,
                        span,
                        xml,
                        &mut raw_ordinal,
                        slide_dialect
                            .ok_or_else(|| invalid("ink-action slide dialect is unresolved"))?,
                        reader.decoder(),
                    )? {
                        if candidates.len() >= maximum {
                            return Err(limit("ink-action anchor count", maximum));
                        }
                        candidates
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "ink-action anchor candidates",
                                source,
                            })?;
                        candidates.push(candidate);
                    }
                }
            },
            Event::Text(value) => {
                validate_text(value.as_ref(), "ink-action slide text")?;
                let whitespace = is_xml_whitespace(value.as_ref());
                if stack.is_empty() && !whitespace {
                    return Err(invalid("ink-action slide has text outside its root"));
                }
                if !whitespace && active_alternate_branch(&stack).is_some() {
                    return Err(invalid(
                        "ink-action MCE branch contains non-whitespace text",
                    ));
                }
            },
            Event::CData(value) => {
                validate_text(value.as_ref(), "ink-action slide CDATA")?;
                let whitespace = is_xml_whitespace(value.as_ref());
                if stack.is_empty() && !whitespace {
                    return Err(invalid("ink-action slide has CDATA outside its root"));
                }
                if !whitespace && active_alternate_branch(&stack).is_some() {
                    return Err(invalid("ink-action MCE branch contains CDATA"));
                }
            },
            Event::Comment(value) => {
                validate_comment(value.as_ref())?;
            },
            Event::GeneralRef(reference) => {
                let character = validate_general_ref(&reference)?;
                if stack.is_empty() {
                    return Err(invalid(
                        "ink-action slide has a general reference outside its root",
                    ));
                }
                if active_alternate_branch(&stack).is_some()
                    && !character.is_some_and(is_xml_whitespace_character)
                {
                    return Err(invalid(
                        "ink-action MCE branch contains a non-whitespace reference",
                    ));
                }
            },
            Event::Decl(declaration) => {
                if declaration_seen || non_declaration_event_seen {
                    return Err(invalid(
                        "ink-action slide XML declaration is misplaced or duplicated",
                    ));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "ink-action slide rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || !stack.is_empty() {
                    return Err(invalid("ink-action slide root is absent or unterminated"));
                }
                return Ok(candidates);
            },
        }
    }
}

fn direct_branch_mut(
    alternate: &mut AlternateState,
    child_depth: usize,
) -> Option<&mut BranchState> {
    alternate
        .choices
        .iter_mut()
        .rev()
        .find(|branch| branch.end.is_none() && branch.depth.saturating_add(1) == child_depth)
        .or_else(|| {
            alternate.fallback.as_mut().filter(|branch| {
                branch.end.is_none() && branch.depth.saturating_add(1) == child_depth
            })
        })
}

fn finalize_alternate(
    alternate: &mut AlternateState,
    span: Range<usize>,
    source: &[u8],
    raw_ordinal: &mut usize,
    dialect: Dialect,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<Candidate>> {
    let has_action_capability = alternate.choices.iter().any(BranchState::action_capable);
    if !has_action_capability {
        return Ok(None);
    }
    let source_ordinal = *raw_ordinal;
    *raw_ordinal = raw_ordinal
        .checked_add(1)
        .ok_or_else(|| limit("ink-action source ordinal", usize::MAX))?;
    if alternate.choices.len() != 1 {
        return Err(invalid(
            "ink-action AlternateContent has duplicate or ambiguous Choice branches",
        ));
    }
    let choice = alternate
        .choices
        .first_mut()
        .ok_or_else(|| invalid("ink-action AlternateContent is missing its Choice branch"))?;
    let Some(requires) = choice.requires.as_ref() else {
        return Err(invalid("ink-action Choice is missing Requires"));
    };
    if requires.len() != 2
        || !requires.iter().any(|uri| uri.as_slice() == P14)
        || !requires.iter().any(|uri| uri.as_slice() == ACTION)
    {
        // Keep an action-capable but unsupported branch in the raw source
        // catalog. The typed owner never promotes it to an action target.
        let choice_span = choice.start..choice.end.unwrap_or(span.end);
        let fingerprint = AnchorFingerprint::from_closure(
            source.get(span.clone()).unwrap_or_default(),
            source.get(choice_span.clone()).unwrap_or_default(),
            choice
                .relationship_attributes
                .first()
                .map_or(&[][..], |attribute| attribute.value.as_slice()),
            &[],
            &[],
            &[],
            &[],
        );
        return Ok(Some(Candidate {
            source_ordinal,
            span: span.clone(),
            choice_span: choice_span.clone(),
            fallback_span: alternate
                .fallback
                .as_ref()
                .map(|branch| branch.start..branch.end.unwrap_or(span.end))
                .unwrap_or_else(|| span.clone()),
            content_span: choice
                .content_start
                .zip(choice.content_end)
                .map_or(choice_span, |(start, end)| start..end),
            relationship_id: None,
            fingerprint,
            typed: false,
            dialect,
        }));
    }
    let fallback = alternate
        .fallback
        .as_ref()
        .ok_or_else(|| invalid("ink-action AlternateContent is missing its p:pic fallback"))?;
    if choice.element_children != 1 || choice.content_start.is_none() {
        return Err(invalid(
            "ink-action Choice must contain exactly one p:contentPart child",
        ));
    }
    if fallback.element_children != 1 || fallback.picture_children != 1 {
        return Err(invalid(
            "ink-action Fallback must contain exactly one p:pic child",
        ));
    }
    let content_start = choice
        .content_start
        .ok_or_else(|| invalid("ink-action Choice contentPart span is missing"))?;
    let content_end = choice
        .content_end
        .ok_or_else(|| invalid("ink-action Choice contentPart span is unterminated"))?;
    let relationship_id = relationship_id(&choice.relationship_attributes, decoder, dialect)?;
    source
        .get(span.clone())
        .ok_or_else(|| invalid("ink-action source span is invalid"))?;
    Ok(Some(Candidate {
        source_ordinal,
        span: span.clone(),
        choice_span: choice.start..choice.end.unwrap_or(content_end),
        fallback_span: fallback.start..fallback.end.unwrap_or(content_end),
        content_span: content_start..content_end,
        relationship_id: Some(relationship_id),
        fingerprint: AnchorFingerprint::from_closure(
            source.get(span.clone()).unwrap_or_default(),
            source
                .get(choice.start..choice.end.unwrap_or(content_end))
                .unwrap_or_default(),
            choice
                .relationship_attributes
                .first()
                .map_or(&[][..], |attribute| attribute.value.as_slice()),
            &[],
            &[],
            &[],
            &[],
        ),
        typed: true,
        dialect,
    }))
}

fn requires_value(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<Vec<u8>>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != b"Requires" {
            continue;
        }
        if value.is_some() {
            return Err(invalid(
                "ink-action Choice has duplicate Requires attributes",
            ));
        }
        let text = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let mut uris = Vec::new();
        for token in text.split_ascii_whitespace() {
            if !is_ncname(token) {
                return Err(invalid(
                    "ink-action Choice Requires token is not a namespace prefix",
                ));
            }
            if token.len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
                return Err(limit(
                    "ink-action Requires token",
                    crate::presentation::embedded::MAX_ATTRIBUTE_BYTES,
                ));
            }
            let mut qualified = Vec::new();
            qualified
                .try_reserve(token.len().saturating_add(2))
                .map_err(|source| Error::Allocation {
                    resource: "ink-action Requires token",
                    source,
                })?;
            qualified.extend_from_slice(token.as_bytes());
            qualified.extend_from_slice(b":x");
            let prefix = QName(qualified.as_slice())
                .prefix()
                .ok_or_else(|| invalid("ink-action Choice Requires token is not a prefix"))?;
            let resolved = resolver.resolve_prefix(Some(prefix), false);
            let uri = match resolved {
                ResolveResult::Bound(Namespace(uri)) => normalize_namespace(uri)?,
                ResolveResult::Unbound | ResolveResult::Unknown(_) => UNKNOWN_REQUIRES.to_vec(),
            };
            uris.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "ink-action Requires namespaces",
                source,
            })?;
            uris.push(uri);
        }
        value = Some(uris);
    }
    let value = value.ok_or_else(|| invalid("ink-action Choice is missing Requires"))?;
    if value.is_empty() {
        return Err(invalid("ink-action Choice Requires must contain a prefix"));
    }
    Ok(value)
}

fn capture_relationship_attributes(
    output: &mut Vec<RelationshipAttribute>,
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let namespace = match resolver.resolve_attribute(attribute.key).0 {
            ResolveResult::Bound(Namespace(uri)) => normalize_namespace(uri)?,
            ResolveResult::Unbound => Vec::new(),
            ResolveResult::Unknown(_) => {
                return Err(invalid(
                    "ink-action contentPart r:id uses an undeclared namespace prefix",
                ));
            },
        };
        output.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ink-action contentPart relationship attributes",
            source,
        })?;
        let mut value = Vec::new();
        value
            .try_reserve(attribute.value.len())
            .map_err(|source| Error::Allocation {
                resource: "ink-action contentPart relationship attribute",
                source,
            })?;
        value.extend_from_slice(attribute.value.as_ref());
        output.push(RelationshipAttribute { namespace, value });
    }
    Ok(())
}

fn relationship_id(
    attributes: &[RelationshipAttribute],
    decoder: quick_xml::encoding::Decoder,
    dialect: Dialect,
) -> Result<String> {
    let mut value = None;
    let expected = match dialect {
        Dialect::Transitional => REL,
        Dialect::Strict => STRICT_REL,
    };
    for attribute in attributes {
        let is_relationship = namespace_matches(&attribute.namespace, expected);
        let is_other_relationship = (namespace_matches(&attribute.namespace, REL)
            || namespace_matches(&attribute.namespace, STRICT_REL))
            && !is_relationship;
        if is_other_relationship {
            return Err(invalid(
                "ink-action contentPart r:id namespace does not match the slide dialect",
            ));
        }
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(invalid(
                "ink-action contentPart has duplicate r:id attributes",
            ));
        }
        let attribute = quick_xml::events::attributes::Attribute {
            key: QName(b"id"),
            value: std::borrow::Cow::Borrowed(attribute.value.as_slice()),
        };
        let text = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if text.is_empty() || text.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(invalid("ink-action contentPart has an invalid r:id"));
        }
        value = Some(text.into_owned());
    }
    value.ok_or_else(|| invalid("ink-action contentPart is missing r:id"))
}

fn validate_attributes(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let mut attributes = 0usize;
    let mut namespaces = 0usize;
    let mut seen = Vec::<(Vec<u8>, Vec<u8>)>::new();
    let mut seen_namespace_prefixes = Vec::<Vec<u8>>::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        attributes = attributes
            .checked_add(1)
            .ok_or_else(|| limit("ink-action owner XML attributes", MAX_ATTRIBUTES))?;
        if attributes > MAX_ATTRIBUTES {
            return Err(limit("ink-action owner XML attributes", MAX_ATTRIBUTES));
        }
        if attribute.key.as_ref().len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "ink-action owner XML attribute name bytes",
                crate::presentation::embedded::MAX_ATTRIBUTE_BYTES,
            ));
        }
        validate_qname(attribute.key, "ink-action owner XML attribute name")?;
        if attribute.value.len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "ink-action owner XML attribute bytes",
                crate::presentation::embedded::MAX_ATTRIBUTE_BYTES,
            ));
        }
        if attribute.value.contains(&b'<') {
            return Err(invalid(
                "ink-action owner XML attribute contains a raw '<' delimiter",
            ));
        }
        let is_namespace =
            attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:");
        if is_namespace {
            namespaces = namespaces.checked_add(1).ok_or_else(|| {
                limit(
                    "ink-action namespace declarations",
                    MAX_NAMESPACE_DECLARATIONS,
                )
            })?;
            if namespaces > MAX_NAMESPACE_DECLARATIONS {
                return Err(limit(
                    "ink-action namespace declarations",
                    MAX_NAMESPACE_DECLARATIONS,
                ));
            }
            let prefix = attribute
                .key
                .as_ref()
                .strip_prefix(b"xmlns:")
                .unwrap_or_default();
            if !prefix.is_empty()
                && !is_ncname(
                    std::str::from_utf8(prefix).map_err(|error| Error::Xml(error.to_string()))?,
                )
            {
                return Err(invalid(
                    "ink-action owner XML namespace declaration has an invalid prefix",
                ));
            }
            if seen_namespace_prefixes
                .iter()
                .any(|known| known.as_slice() == prefix)
            {
                return Err(invalid(
                    "ink-action owner XML has duplicate namespace declarations",
                ));
            }
            seen_namespace_prefixes
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ink-action namespace declaration keys",
                    source,
                })?;
            seen_namespace_prefixes.push(prefix.to_vec());
            let value = decode_attribute_value(&attribute, decoder)?;
            validate_namespace_binding(prefix, value.as_bytes())?;
            continue;
        }

        let value = decode_attribute_value(&attribute, decoder)?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(uri)) => normalize_namespace(uri)?,
            ResolveResult::Unbound => Vec::new(),
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "ink-action owner XML attribute uses undeclared namespace prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        let local = local.as_ref();
        if seen.iter().any(|(known_namespace, known_local)| {
            known_namespace.as_slice() == namespace.as_slice() && known_local.as_slice() == local
        }) {
            return Err(invalid(
                "ink-action owner XML has duplicate expanded attributes",
            ));
        }
        seen.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ink-action expanded attribute keys",
            source,
        })?;
        seen.push((namespace, local.to_vec()));
        if value.len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "ink-action owner XML decoded attribute bytes",
                crate::presentation::embedded::MAX_ATTRIBUTE_BYTES,
            ));
        }
    }
    Ok(())
}

fn decode_attribute_value(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<String> {
    let value = attribute
        .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
        .map_err(|error| Error::Xml(error.to_string()))?;
    validate_xml_characters(value.as_bytes(), "ink-action XML attribute")?;
    Ok(value.into_owned())
}

fn validate_element_name(name: QName<'_>) -> Result<()> {
    validate_qname(name, "ink-action owner XML element name")
}

fn validate_end_name(name: QName<'_>) -> Result<()> {
    validate_qname(name, "ink-action owner XML closing element name")
}

fn validate_qname(name: QName<'_>, field: &'static str) -> Result<()> {
    if name.as_ref().len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
        return Err(limit(
            "ink-action owner XML qualified name bytes",
            crate::presentation::embedded::MAX_ATTRIBUTE_BYTES,
        ));
    }
    let name = std::str::from_utf8(name.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    if !is_qualified_name(name) {
        return Err(invalid(format!("{field} is not a valid XML QName")));
    }
    Ok(())
}

fn validate_dialect(namespace: ElementNamespace, dialect: Dialect) -> Result<()> {
    if namespace.is_pml() && namespace.dialect() != Some(dialect) {
        return Err(invalid(
            "ink-action owner XML mixes Strict and Transitional PML namespaces",
        ));
    }
    Ok(())
}

fn validate_namespace_binding(prefix: &[u8], value: &[u8]) -> Result<()> {
    if prefix == b"xmlns" || value == XMLNS_NAMESPACE {
        return Err(invalid(
            "ink-action owner XML uses the reserved XMLNS namespace",
        ));
    }
    if value == XML_NAMESPACE && prefix != b"xml" {
        return Err(invalid(
            "ink-action owner XML binds the XML namespace to a non-xml prefix",
        ));
    }
    if prefix == b"xml" && value != XML_NAMESPACE {
        return Err(invalid("ink-action owner XML rebinds the xml prefix"));
    }
    if prefix.is_empty() && value == XML_NAMESPACE {
        return Err(invalid(
            "ink-action owner XML binds the XML namespace as default",
        ));
    }
    Ok(())
}

fn active_alternate_branch(stack: &[ElementFrame]) -> Option<&BranchState> {
    stack.iter().rev().find_map(|frame| {
        frame.alternate.as_ref().and_then(|alternate| {
            alternate
                .choices
                .iter()
                .rev()
                .find(|branch| branch.end.is_none() && branch.depth <= stack.len())
                .or_else(|| {
                    alternate
                        .fallback
                        .as_ref()
                        .filter(|branch| branch.end.is_none() && branch.depth <= stack.len())
                })
        })
    })
}

fn validate_text(value: &[u8], field: &'static str) -> Result<()> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))?;
    if value.contains("]]>") {
        return Err(invalid(format!(
            "{field} contains the forbidden ']]>' delimiter"
        )));
    }
    validate_xml_characters(value.as_bytes(), field)
}

fn validate_xml_characters(value: &[u8], field: &'static str) -> Result<()> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))?;
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid(format!(
            "{field} contains an invalid XML character"
        )))
    }
}

fn validate_comment(value: &[u8]) -> Result<()> {
    validate_xml_characters(value, "ink-action XML comment")?;
    if value.windows(2).any(|window| window == b"--") || value.ends_with(b"-") {
        return Err(invalid("ink-action XML comment is malformed"));
    }
    Ok(())
}

fn validate_general_ref(reference: &BytesRef<'_>) -> Result<Option<char>> {
    if reference.as_ref().len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES {
        return Err(limit(
            "ink-action XML general reference bytes",
            crate::presentation::embedded::MAX_ATTRIBUTE_BYTES,
        ));
    }
    if reference.is_char_ref() {
        let character = reference
            .resolve_char_ref()
            .map_err(|error| Error::Xml(error.to_string()))?
            .ok_or_else(|| invalid("ink-action XML character reference is invalid"))?;
        if !is_xml_char(character) {
            return Err(invalid(
                "ink-action XML character reference is not an XML character",
            ));
        }
        return Ok(Some(character));
    }
    let name = reference
        .decode()
        .map_err(|error| Error::Xml(error.to_string()))?;
    let replacement = quick_xml::escape::resolve_xml_entity(name.as_ref())
        .ok_or_else(|| invalid("ink-action XML has an undefined general reference"))?;
    validate_xml_characters(replacement.as_bytes(), "ink-action XML general reference")?;
    let mut characters = replacement.chars();
    let character = characters.next();
    if characters.next().is_some() {
        return Ok(None);
    }
    Ok(character)
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn is_xml_whitespace_character(value: char) -> bool {
    matches!(value, ' ' | '\t' | '\r' | '\n')
}

fn is_xml_char(value: char) -> bool {
    matches!(
        value,
        '\u{9}'
            | '\u{A}'
            | '\u{D}'
            | '\u{20}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
    )
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let source = declaration.as_ref();
    if source.len() > crate::presentation::embedded::MAX_ATTRIBUTE_BYTES
        || !source.starts_with(b"xml")
    {
        return Err(invalid("ink-action XML declaration is malformed"));
    }
    let mut cursor = 3usize;
    if source
        .get(cursor)
        .is_none_or(|byte| !is_xml_whitespace_byte(*byte))
    {
        return Err(invalid("ink-action XML declaration lacks whitespace"));
    }
    let mut seen_version = false;
    let mut seen_encoding = false;
    let mut seen_standalone = false;
    while cursor < source.len() {
        while cursor < source.len() && is_xml_whitespace_byte(source[cursor]) {
            cursor += 1;
        }
        if cursor == source.len() {
            break;
        }
        let key_start = cursor;
        while cursor < source.len()
            && !is_xml_whitespace_byte(source[cursor])
            && source[cursor] != b'='
        {
            cursor += 1;
        }
        if key_start == cursor {
            return Err(invalid(
                "ink-action XML declaration has an invalid attribute",
            ));
        }
        let key = &source[key_start..cursor];
        while cursor < source.len() && is_xml_whitespace_byte(source[cursor]) {
            cursor += 1;
        }
        if source.get(cursor) != Some(&b'=') {
            return Err(invalid("ink-action XML declaration attribute lacks '='"));
        }
        cursor += 1;
        while cursor < source.len() && is_xml_whitespace_byte(source[cursor]) {
            cursor += 1;
        }
        let quote = match source.get(cursor).copied() {
            Some(quote @ (b'\'' | b'"')) => quote,
            _ => return Err(invalid("ink-action XML declaration value is not quoted")),
        };
        cursor += 1;
        let value_start = cursor;
        while cursor < source.len() && source[cursor] != quote {
            cursor += 1;
        }
        if cursor == source.len() {
            return Err(invalid("ink-action XML declaration value is unterminated"));
        }
        let value = &source[value_start..cursor];
        cursor += 1;
        if cursor < source.len() && !is_xml_whitespace_byte(source[cursor]) {
            return Err(invalid(
                "ink-action XML declaration attributes are not separated",
            ));
        }
        match key {
            b"version" if !seen_version && !seen_encoding && !seen_standalone => {
                if value != b"1.0" {
                    return Err(invalid("ink-action XML declaration version is not 1.0"));
                }
                seen_version = true;
            },
            b"encoding" if seen_version && !seen_encoding && !seen_standalone => {
                if !value.eq_ignore_ascii_case(b"UTF-8") {
                    return Err(invalid("ink-action XML declaration encoding is not UTF-8"));
                }
                seen_encoding = true;
            },
            b"standalone" if seen_version && !seen_standalone => {
                if value != b"yes" && value != b"no" {
                    return Err(invalid(
                        "ink-action XML declaration standalone value is invalid",
                    ));
                }
                seen_standalone = true;
            },
            b"version" | b"encoding" | b"standalone" => {
                return Err(invalid(
                    "ink-action XML declaration has a duplicate or misplaced attribute",
                ));
            },
            _ => {
                return Err(invalid(
                    "ink-action XML declaration has an unsupported attribute",
                ));
            },
        }
    }
    if !seen_version {
        return Err(invalid("ink-action XML declaration lacks version"));
    }
    Ok(())
}

fn is_xml_whitespace_byte(value: u8) -> bool {
    matches!(value, b' ' | b'\t' | b'\r' | b'\n')
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("ink-action owner XML offset does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn slide(body: &str) -> Vec<u8> {
        format!(
            r#"<p:sld xmlns:p="{pml}" xmlns:r="{rel}" xmlns:mc="{mc}" xmlns:p14="{p14}" xmlns:ia="{action}"><p:spTree>{body}</p:spTree></p:sld>"#,
            pml = std::str::from_utf8(PML).unwrap(),
            rel = std::str::from_utf8(REL).unwrap(),
            mc = std::str::from_utf8(MC).unwrap(),
            p14 = std::str::from_utf8(P14).unwrap(),
            action = std::str::from_utf8(ACTION).unwrap(),
        )
        .into_bytes()
    }

    fn alternate(requires: &str, content: &str) -> String {
        format!(
            r#"<mc:AlternateContent><mc:Choice Requires="{requires}">{content}</mc:Choice><mc:Fallback><p:pic/></mc:Fallback></mc:AlternateContent>"#
        )
    }

    #[test]
    fn rejects_late_duplicate_declarations_comments_and_undefined_references() {
        let body = alternate("p14 ia", r#"<p:contentPart r:id="rIdAction"/>"#);
        let source = slide(&body);
        assert!(scan_slide(b"<?x", 8).is_err());
        let late = format!(
            "<!--before--><?xml version=\"1.0\"?>{}",
            String::from_utf8(source.clone()).unwrap()
        );
        assert!(scan_slide(late.as_bytes(), 8).is_err());
        let duplicate = format!(
            "<?xml version=\"1.0\"?><?xml version=\"1.0\"?>{}",
            String::from_utf8(source.clone()).unwrap()
        );
        assert!(scan_slide(duplicate.as_bytes(), 8).is_err());
        let malformed_comment = String::from_utf8(source.clone())
            .unwrap()
            .replace("<p:spTree>", "<p:spTree><!--bad--comment-->");
        assert!(scan_slide(malformed_comment.as_bytes(), 8).is_err());
        let legal_comment = String::from_utf8(source.clone())
            .unwrap()
            .replace("<p:spTree>", "<p:spTree><!--keep ]]>-->");
        assert_eq!(scan_slide(legal_comment.as_bytes(), 8).unwrap().len(), 1);
        let undefined_reference = String::from_utf8(source)
            .unwrap()
            .replace("<p:spTree>", "<p:spTree>&undefined;");
        assert!(scan_slide(undefined_reference.as_bytes(), 8).is_err());
    }

    #[test]
    fn rejects_duplicate_expanded_attributes_and_uses_one_mce_namespace() {
        let duplicate = slide(
            r#"<p:shape xmlns:a="urn:duplicate" xmlns:b="urn:duplicate" a:id="1" b:id="2"/>"#,
        );
        assert!(scan_slide(&duplicate, 8).is_err());

        let strict = format!(
            r#"<p:sld xmlns:p="{pml}" xmlns:r="{rel}" xmlns:mc="{mc}" xmlns:p14="{p14}" xmlns:ia="{action}"><p:spTree>{body}</p:spTree></p:sld>"#,
            pml = std::str::from_utf8(STRICT_PML).unwrap(),
            mc = std::str::from_utf8(MC).unwrap(),
            rel = std::str::from_utf8(STRICT_REL).unwrap(),
            p14 = std::str::from_utf8(P14).unwrap(),
            action = std::str::from_utf8(ACTION).unwrap(),
            body = alternate("p14 ia", r#"<p:contentPart r:id="rIdAction"/>"#),
        );
        assert_eq!(scan_slide(strict.as_bytes(), 8).unwrap().len(), 1);
        let foreign_mc = strict.replace(
            std::str::from_utf8(MC).unwrap(),
            "http://purl.oclc.org/ooxml/markup-compatibility/2006",
        );
        assert!(scan_slide(foreign_mc.as_bytes(), 8).unwrap().is_empty());
    }

    #[test]
    fn requires_tokens_are_prefixes_and_branch_text_is_rejected() {
        let malformed_prefix = slide(&alternate(
            "p14:main ia",
            r#"<p:contentPart r:id="rIdAction"/>"#,
        ));
        assert!(scan_slide(&malformed_prefix, 8).is_err());

        let text = slide(&alternate(
            "p14 ia",
            r#"text<p:contentPart r:id="rIdAction"/>"#,
        ));
        assert!(scan_slide(&text, 8).is_err());
        let cdata = slide(&alternate(
            "p14 ia",
            r#"<![CDATA[text]]><p:contentPart r:id="rIdAction"/>"#,
        ));
        assert!(scan_slide(&cdata, 8).is_err());
        let whitespace_cdata = slide(&alternate(
            "p14 ia",
            r#"<![CDATA[ ]]><p:contentPart r:id="rIdAction"/>"#,
        ));
        assert_eq!(scan_slide(&whitespace_cdata, 8).unwrap().len(), 1);
    }

    #[test]
    fn defers_malformed_relationship_id_for_unsupported_action_branches() {
        let opaque = slide(&alternate("p14 ia future", r#"<p:contentPart r:id=""/>"#));
        let candidates = scan_slide(&opaque, 8).unwrap();
        assert_eq!(candidates.len(), 1);
        assert!(!candidates[0].typed);

        let typed = slide(&alternate("p14 ia", r#"<p:contentPart r:id=""/>"#));
        assert!(scan_slide(&typed, 8).is_err());
    }
}
