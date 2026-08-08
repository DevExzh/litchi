//! Exact-source snapshots for active WordprocessingML content controls.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_ooxml_common::mce::OffsetLimits;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

use super::{
    BindingFlavor, FORMATTING_ALLOWED_NAMESPACE, Inventory, Limits, Lock,
    STORE_ITEM_CHECKSUM_NAMESPACE, content_control_capabilities,
};

const WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const WORD_2012: &[u8] = b"http://schemas.microsoft.com/office/word/2012/wordml";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_ATTRIBUTES: usize = 1_024;

/// A checked half-open byte range into a [`Snapshot`]'s exact source.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    start: usize,
    end: usize,
}

impl Span {
    pub(crate) fn new(start: usize, end: usize) -> Result<Self> {
        if start > end {
            return Err(invalid("content-control source span is reversed"));
        }
        Ok(Self { start, end })
    }

    /// Inclusive byte offset.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// Exclusive byte offset.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }

    /// Number of bytes in the range.
    #[must_use]
    pub const fn len(self) -> usize {
        self.end - self.start
    }
}

/// Exact lexical extent of an XML attribute and its unquoted value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AttributeSpan {
    /// Complete attribute, excluding preceding whitespace.
    pub attribute: Span,
    /// Attribute value, excluding quotes.
    pub value: Span,
}

/// Exact source location of one active binding element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BindingSpan {
    pub(crate) flavor: BindingFlavor,
    pub(crate) start_tag: Span,
    pub(crate) checksum: Option<AttributeSpan>,
    pub(crate) checksum_count: usize,
    pub(crate) ignorable: Option<AttributeSpan>,
    pub(crate) ignorable_count: usize,
}

impl BindingSpan {
    /// Exact owner vocabulary of this binding element.
    #[must_use]
    pub const fn flavor(&self) -> BindingFlavor {
        self.flavor
    }

    /// Opening or empty-element tag span.
    #[must_use]
    pub const fn start_tag(&self) -> Span {
        self.start_tag
    }

    /// Exact checksum attribute when it occurs once.
    #[must_use]
    pub const fn checksum(&self) -> Option<AttributeSpan> {
        self.checksum
    }

    /// Number of checksum attributes with the exact expanded name.
    #[must_use]
    pub const fn checksum_count(&self) -> usize {
        self.checksum_count
    }
}

/// Exact source location of one active lock element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LockSpan {
    pub(crate) lock: Lock,
    pub(crate) start_tag: Span,
    pub(crate) formatting_allowed: Option<AttributeSpan>,
    pub(crate) formatting_count: usize,
    pub(crate) ignorable: Option<AttributeSpan>,
    pub(crate) ignorable_count: usize,
}

impl LockSpan {
    /// Semantic lock value owned by this exact element.
    #[must_use]
    pub const fn lock(&self) -> Lock {
        self.lock
    }
    /// Opening or empty-element tag span.
    #[must_use]
    pub const fn start_tag(&self) -> Span {
        self.start_tag
    }

    /// Exact formatting attribute when it occurs once.
    #[must_use]
    pub const fn formatting_allowed(&self) -> Option<AttributeSpan> {
        self.formatting_allowed
    }

    /// Number of formatting attributes with the exact expanded name.
    #[must_use]
    pub const fn formatting_count(&self) -> usize {
        self.formatting_count
    }
}

/// Exact coordinates for one source-order active `w:sdtPr` occurrence.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceOccurrence {
    ordinal: usize,
    id: Option<u32>,
    id_count: usize,
    properties: Span,
    start_tag: Span,
    bindings: Arc<[BindingSpan]>,
    locks: Arc<[LockSpan]>,
}

impl SourceOccurrence {
    /// Source-order identity. It remains distinct when IDs are absent or repeated.
    #[must_use]
    pub const fn ordinal(&self) -> usize {
        self.ordinal
    }

    /// Optional semantic `w:id`.
    #[must_use]
    pub const fn id(&self) -> Option<u32> {
        self.id
    }

    /// Number of direct active `w:id` properties.
    #[must_use]
    pub const fn id_count(&self) -> usize {
        self.id_count
    }

    /// Complete `w:sdtPr` extent.
    #[must_use]
    pub const fn properties(&self) -> Span {
        self.properties
    }

    /// Opening `w:sdtPr` tag.
    #[must_use]
    pub const fn start_tag(&self) -> Span {
        self.start_tag
    }

    /// Active core or Word 2012 binding elements owned by this occurrence.
    #[must_use]
    pub fn bindings(&self) -> &[BindingSpan] {
        &self.bindings
    }

    /// Active lock elements owned by this occurrence.
    #[must_use]
    pub fn locks(&self) -> &[LockSpan] {
        &self.locks
    }
}

#[derive(Debug, Clone)]
pub(crate) enum Source {
    Detached(Arc<Vec<u8>>),
    Package(Arc<Vec<u8>>),
}

impl Source {
    pub(crate) fn as_slice(&self) -> &[u8] {
        match self {
            Self::Detached(value) => value,
            Self::Package(value) => value,
        }
    }
}

#[derive(Debug)]
struct Inner {
    source: Source,
    inventory: Arc<Inventory>,
    occurrences: Arc<[SourceOccurrence]>,
    limits: Limits,
}

/// Immutable semantic and exact-source view of active content controls.
#[derive(Debug, Clone)]
pub struct Snapshot(Arc<Inner>);

impl Snapshot {
    /// Parse owned XML with production resource limits.
    pub fn from_xml(source: impl Into<Vec<u8>>) -> Result<Self> {
        Self::from_xml_with_limits(source, Limits::default())
    }

    /// Parse owned XML with explicit resource limits.
    pub fn from_xml_with_limits(source: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        limits.validate()?;
        let source = source.into();
        if source.len() > limits.max_input_bytes {
            return Err(invalid(
                "content-control source exceeds the input byte limit",
            ));
        }
        Self::from_source(Source::Detached(Arc::new(source)), limits)
    }

    pub(crate) fn from_package(source: Arc<Vec<u8>>, limits: Limits) -> Result<Self> {
        Self::from_source(Source::Package(source), limits)
    }

    pub(crate) fn from_source(source: Source, limits: Limits) -> Result<Self> {
        let inventory = Inventory::parse_with_limits(source.as_slice(), &limits)?;
        let mut occurrences = scan(source.as_slice(), &limits)?;
        if occurrences.len() != inventory.occurrences().len() {
            return Err(invalid(format!(
                "semantic and exact-source content-control inventories disagree (semantic {}, source {})",
                inventory.occurrences().len(),
                occurrences.len()
            )));
        }
        for (source_occurrence, semantic) in occurrences.iter_mut().zip(inventory.occurrences()) {
            let semantic_bindings = semantic.control().data_bindings();
            if source_occurrence.bindings.len() != semantic_bindings.len()
                || source_occurrence
                    .bindings
                    .iter()
                    .zip(semantic_bindings)
                    .any(|(source, semantic)| source.flavor() != semantic.flavor())
            {
                return Err(invalid(
                    "semantic and exact-source binding inventories disagree",
                ));
            }
            if source_occurrence.locks.len() > 1
                || source_occurrence
                    .locks
                    .first()
                    .is_some_and(|source| source.lock() != semantic.control().lock())
                || (source_occurrence.locks.is_empty()
                    && semantic.control().lock() != Lock::Unlocked)
            {
                return Err(invalid(
                    "semantic and exact-source lock inventories disagree",
                ));
            }
            source_occurrence.ordinal = semantic.ordinal();
            source_occurrence.id = semantic.id();
        }
        Ok(Self(Arc::new(Inner {
            source,
            inventory: Arc::new(inventory),
            occurrences: Arc::from(occurrences.into_boxed_slice()),
            limits,
        })))
    }

    /// Exact retained XML bytes.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        self.0.source.as_slice()
    }

    pub(crate) fn source_owner(&self) -> Source {
        self.0.source.clone()
    }

    /// Share package-owned source bytes without materializing another story copy.
    pub(crate) fn package_source_arc(&self) -> Option<Arc<Vec<u8>>> {
        match &self.0.source {
            Source::Package(value) => Some(Arc::clone(value)),
            Source::Detached(_) => None,
        }
    }

    /// Bounded semantic inventory.
    #[must_use]
    pub fn inventory(&self) -> &Inventory {
        &self.0.inventory
    }

    /// Exact active occurrences aligned one-to-one with [`Inventory::occurrences`].
    #[must_use]
    pub fn occurrences(&self) -> &[SourceOccurrence] {
        &self.0.occurrences
    }

    /// Limits retained for edits and reparsing.
    #[must_use]
    pub fn limits(&self) -> &Limits {
        &self.0.limits
    }

    /// Begin a failure-atomic exact-source edit.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }
}

#[derive(Debug)]
struct Open {
    depth: usize,
    begin: usize,
    start_tag: Span,
    id_count: usize,
    bindings: Vec<BindingSpan>,
    locks: Vec<LockSpan>,
}

#[derive(Debug)]
enum ProcessTarget {
    Exact { namespace: Vec<u8>, local: Vec<u8> },
    Namespace(Vec<u8>),
}

impl ProcessTarget {
    fn matches(&self, namespace: &[u8], local: &[u8]) -> bool {
        match self {
            Self::Exact {
                namespace: target_namespace,
                local: target_local,
            } => target_namespace == namespace && target_local == local,
            Self::Namespace(target_namespace) => target_namespace == namespace,
        }
    }
}

#[derive(Debug)]
struct MceFrame {
    introduced_ignorable: Vec<Vec<u8>>,
    process: Vec<ProcessTarget>,
    transparent: bool,
}

fn scan(source: &[u8], limits: &Limits) -> Result<Vec<SourceOccurrence>> {
    let active = active_offsets(source, limits)?;
    let mut reader = NsReader::from_reader(source);
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut open = None::<Open>;
    let capabilities = content_control_capabilities();
    let mut mce_path = Vec::<MceFrame>::new();
    let mut occurrences = Vec::new();
    occurrences
        .try_reserve(limits.max_content_controls.min(1_024))
        .map_err(alloc("content-control source inventory"))?;

    loop {
        let begin = pos(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let end = pos(&reader)?;
        events = checked_add(events, 1, "content-control source event count")?;
        if events > limits.max_events {
            return Err(invalid("content-control source event limit exceeded"));
        }
        match event {
            Event::Start(element) => {
                depth = checked_add(depth, 1, "content-control XML depth")?;
                if depth > limits.max_depth {
                    return Err(invalid("content-control XML depth limit exceeded"));
                }
                if is_word(&namespace) && element.local_name().as_ref() == b"sdtPr" {
                    // `active_offsets` deliberately drops comments in the XML
                    // prolog. A standalone `w:sdtPr` is itself the document
                    // element, so its marker is in that prolog even though the
                    // element is necessarily active. Its selected child
                    // markers still provide ordinary MCE filtering.
                    if depth == 1 || active.contains(&begin) {
                        if open.is_some() {
                            return Err(invalid("nested content-control properties are invalid"));
                        }
                        open = Some(Open {
                            depth,
                            begin,
                            start_tag: Span::new(begin, end)?,
                            id_count: 0,
                            bindings: Vec::new(),
                            locks: Vec::new(),
                        });
                    }
                } else if active.contains(&begin) {
                    inspect_child(
                        source,
                        begin,
                        end,
                        depth,
                        &namespace,
                        &element,
                        &mce_path,
                        decoder,
                        &resolver,
                        open.as_mut(),
                    )?;
                }
                let frame = mce_frame(
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    &mce_path,
                    &capabilities,
                    limits,
                )?;
                mce_path
                    .try_reserve(1)
                    .map_err(alloc("content-control MCE ancestry"))?;
                mce_path.push(frame);
            },
            Event::Empty(element) => {
                let child_depth = checked_add(depth, 1, "content-control XML depth")?;
                if child_depth > limits.max_depth {
                    return Err(invalid("content-control XML depth limit exceeded"));
                }
                if is_word(&namespace) && element.local_name().as_ref() == b"sdtPr" {
                    if child_depth == 1 || active.contains(&begin) {
                        push_occurrence(
                            &mut occurrences,
                            Open {
                                depth: child_depth,
                                begin,
                                start_tag: Span::new(begin, end)?,
                                id_count: 0,
                                bindings: Vec::new(),
                                locks: Vec::new(),
                            },
                            end,
                            limits,
                        )?;
                    }
                } else if active.contains(&begin) {
                    inspect_child(
                        source,
                        begin,
                        end,
                        child_depth,
                        &namespace,
                        &element,
                        &mce_path,
                        decoder,
                        &resolver,
                        open.as_mut(),
                    )?;
                }
            },
            Event::End(element) => {
                if open.as_ref().is_some_and(|value| {
                    value.depth == depth
                        && is_word(&namespace)
                        && element.local_name().as_ref() == b"sdtPr"
                }) {
                    let value = open.take().ok_or_else(|| invalid("missing open sdtPr"))?;
                    push_occurrence(&mut occurrences, value, end, limits)?;
                }
                mce_path
                    .pop()
                    .ok_or_else(|| invalid("content-control MCE ancestry underflow"))?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("content-control XML depth underflow"))?;
            },
            Event::Eof => {
                if open.is_some() || depth != 0 || !mce_path.is_empty() {
                    return Err(invalid("unterminated content-control XML"));
                }
                break;
            },
            _ => {},
        }
    }
    Ok(occurrences)
}

#[allow(clippy::too_many_arguments)]
fn inspect_child(
    source: &[u8],
    begin: usize,
    end: usize,
    depth: usize,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    mce_path: &[MceFrame],
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    open: Option<&mut Open>,
) -> Result<()> {
    let Some(open) = open else { return Ok(()) };
    if depth <= open.depth
        || !mce_path
            .get(open.depth..)
            .is_some_and(|wrappers| wrappers.iter().all(|wrapper| wrapper.transparent))
    {
        return Ok(());
    }
    let local = element.local_name();
    if is_word(namespace) && local.as_ref() == b"id" {
        open.id_count = checked_add(open.id_count, 1, "content-control ID count")?;
        return Ok(());
    }
    if (is_word(namespace) || is_namespace(namespace, WORD_2012))
        && local.as_ref() == b"dataBinding"
    {
        let flavor = if is_word(namespace) {
            BindingFlavor::Core
        } else {
            BindingFlavor::Word2012
        };
        let (checksum, checksum_count) = exact_attribute(
            source,
            begin,
            end,
            element,
            decoder,
            resolver,
            b"storeItemChecksum",
            STORE_ITEM_CHECKSUM_NAMESPACE.as_bytes(),
        )?;
        let (ignorable, ignorable_count) = exact_attribute(
            source,
            begin,
            end,
            element,
            decoder,
            resolver,
            b"Ignorable",
            b"http://schemas.openxmlformats.org/markup-compatibility/2006",
        )?;
        open.bindings
            .try_reserve(1)
            .map_err(alloc("content-control binding spans"))?;
        open.bindings.push(BindingSpan {
            flavor,
            start_tag: Span::new(begin, end)?,
            checksum,
            checksum_count,
            ignorable,
            ignorable_count,
        });
    } else if is_word(namespace) && local.as_ref() == b"lock" {
        let lock = lock_value(element, decoder, resolver)?;
        let (formatting_allowed, formatting_count) = exact_attribute(
            source,
            begin,
            end,
            element,
            decoder,
            resolver,
            b"formattingAllowed",
            FORMATTING_ALLOWED_NAMESPACE.as_bytes(),
        )?;
        let (ignorable, ignorable_count) = exact_attribute(
            source,
            begin,
            end,
            element,
            decoder,
            resolver,
            b"Ignorable",
            b"http://schemas.openxmlformats.org/markup-compatibility/2006",
        )?;
        open.locks
            .try_reserve(1)
            .map_err(alloc("content-control lock spans"))?;
        open.locks.push(LockSpan {
            lock,
            start_tag: Span::new(begin, end)?,
            formatting_allowed,
            formatting_count,
            ignorable,
            ignorable_count,
        });
    }
    Ok(())
}

fn lock_value(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<Lock> {
    let mut value = None;
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if attribute.key.local_name().as_ref() != b"val" || !is_word(&namespace) {
            continue;
        }
        if value.is_some() {
            return Err(invalid("duplicate content-control lock value"));
        }
        let lexical = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        value = Some(match lexical.as_ref() {
            "unlocked" => Lock::Unlocked,
            "sdtLocked" => Lock::SdtLocked,
            "contentLocked" => Lock::ContentLocked,
            "sdtContentLocked" => Lock::SdtContentLocked,
            _ => return Err(invalid("invalid content-control lock value")),
        });
    }
    value.ok_or_else(|| invalid("content-control lock has no value"))
}

fn is_mce_wrapper(namespace: &ResolveResult<'_>, local: &[u8]) -> bool {
    is_namespace(namespace, MCE) && matches!(local, b"AlternateContent" | b"Choice" | b"Fallback")
}

fn mce_frame(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    ancestors: &[MceFrame],
    capabilities: &litchi_ooxml_common::mce::Capabilities,
    limits: &Limits,
) -> Result<MceFrame> {
    let (introduced_ignorable, process) =
        mce_directives(element, decoder, resolver, ancestors, limits)?;
    let mut frame = MceFrame {
        introduced_ignorable,
        process,
        transparent: is_mce_wrapper(namespace, element.local_name().as_ref()),
    };
    if frame.transparent {
        return Ok(frame);
    }
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return Ok(frame);
    };
    let namespace_text =
        std::str::from_utf8(namespace).map_err(|error| Error::Xml(error.to_string()))?;
    let local = element.local_name();
    frame.transparent = !capabilities.understands(namespace_text)
        && is_effectively_ignorable(ancestors, &frame, namespace)
        && processes(ancestors, &frame, namespace, local.as_ref());
    Ok(frame)
}

fn mce_directives(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    ancestors: &[MceFrame],
    limits: &Limits,
) -> Result<(Vec<Vec<u8>>, Vec<ProcessTarget>)> {
    let mut local_ignorable = Vec::<Vec<u8>>::new();
    let mut process = Vec::<ProcessTarget>::new();
    let mut tokens = 0usize;
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_namespace(&namespace, MCE) {
            continue;
        }
        let local = attribute.key.local_name();
        if !matches!(local.as_ref(), b"Ignorable" | b"ProcessContent") {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        tokens = checked_add(
            tokens,
            value.split_whitespace().count(),
            "content-control MCE directive token count",
        )?;
        if tokens > limits.max_bindings {
            return Err(invalid(
                "content-control MCE directive token limit exceeded",
            ));
        }
        if local.as_ref() == b"Ignorable" {
            for prefix in value.split_whitespace() {
                let namespace = resolve_prefix_namespace(resolver, prefix)?;
                if !local_ignorable
                    .iter()
                    .any(|candidate| candidate == &namespace)
                {
                    local_ignorable
                        .try_reserve(1)
                        .map_err(alloc("content-control MCE Ignorable directives"))?;
                    local_ignorable.push(namespace);
                }
            }
        } else {
            for token in value.split_whitespace() {
                let (_, local) = token
                    .split_once(':')
                    .ok_or_else(|| invalid("invalid content-control ProcessContent target"))?;
                let (namespace, _) = resolver.resolve_element(QName(token.as_bytes()));
                let namespace = copy_namespace(
                    namespace,
                    "content-control ProcessContent target is unbound",
                )?;
                let target = if local == "*" {
                    ProcessTarget::Namespace(namespace)
                } else {
                    ProcessTarget::Exact {
                        namespace,
                        local: copy_bytes(
                            local.as_bytes(),
                            "content-control MCE ProcessContent local name",
                        )?,
                    }
                };
                process
                    .try_reserve(1)
                    .map_err(alloc("content-control MCE ProcessContent directives"))?;
                process.push(target);
            }
        }
    }

    let mut introduced_ignorable = Vec::new();
    for namespace in local_ignorable {
        if ancestors.iter().any(|frame| {
            frame
                .introduced_ignorable
                .iter()
                .any(|candidate| candidate == &namespace)
        }) {
            continue;
        }
        introduced_ignorable
            .try_reserve(1)
            .map_err(alloc("content-control effective MCE Ignorable directives"))?;
        introduced_ignorable.push(namespace);
    }
    Ok((introduced_ignorable, process))
}

fn resolve_prefix_namespace(resolver: &NamespaceResolver, prefix: &str) -> Result<Vec<u8>> {
    let capacity = checked_add(
        prefix.len(),
        2,
        "content-control MCE namespace probe length",
    )?;
    let mut qualified = Vec::new();
    qualified
        .try_reserve(capacity)
        .map_err(alloc("content-control MCE namespace probe"))?;
    qualified.extend_from_slice(prefix.as_bytes());
    qualified.extend_from_slice(b":_");
    let (namespace, _) = resolver.resolve_element(QName(&qualified));
    copy_namespace(namespace, "content-control Ignorable prefix is unbound")
}

fn copy_namespace(namespace: ResolveResult<'_>, message: &str) -> Result<Vec<u8>> {
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return Err(invalid(message));
    };
    copy_bytes(namespace, "content-control MCE namespace")
}

fn copy_bytes(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copied = Vec::new();
    copied.try_reserve(value.len()).map_err(alloc(resource))?;
    copied.extend_from_slice(value);
    Ok(copied)
}

fn is_effectively_ignorable(ancestors: &[MceFrame], frame: &MceFrame, namespace: &[u8]) -> bool {
    frame
        .introduced_ignorable
        .iter()
        .any(|candidate| candidate == namespace)
        || ancestors.iter().any(|ancestor| {
            ancestor
                .introduced_ignorable
                .iter()
                .any(|candidate| candidate == namespace)
        })
}

fn processes(ancestors: &[MceFrame], frame: &MceFrame, namespace: &[u8], local: &[u8]) -> bool {
    for scope in std::iter::once(frame).chain(ancestors.iter().rev()) {
        if scope
            .process
            .iter()
            .any(|target| target.matches(namespace, local))
        {
            return true;
        }
        if scope
            .introduced_ignorable
            .iter()
            .any(|candidate| candidate == namespace)
        {
            return false;
        }
    }
    false
}

fn push_occurrence(
    occurrences: &mut Vec<SourceOccurrence>,
    open: Open,
    end: usize,
    limits: &Limits,
) -> Result<()> {
    if occurrences.len() >= limits.max_content_controls {
        return Err(invalid("content-control source count limit exceeded"));
    }
    occurrences
        .try_reserve(1)
        .map_err(alloc("content-control source inventory"))?;
    occurrences.push(SourceOccurrence {
        ordinal: occurrences.len(),
        id: None,
        id_count: open.id_count,
        properties: Span::new(open.begin, end)?,
        start_tag: open.start_tag,
        bindings: Arc::from(open.bindings.into_boxed_slice()),
        locks: Arc::from(open.locks.into_boxed_slice()),
    });
    Ok(())
}

fn exact_attribute(
    source: &[u8],
    begin: usize,
    end: usize,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    local: &[u8],
    namespace: &[u8],
) -> Result<(Option<AttributeSpan>, usize)> {
    let mut found = None;
    let mut count = 0usize;
    let mut attributes = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        attributes = checked_add(attributes, 1, "content-control attribute count")?;
        if attributes > MAX_ATTRIBUTES {
            return Err(invalid("content-control attribute count limit exceeded"));
        }
        let (resolved, _) = resolver.resolve_attribute(attribute.key);
        if attribute.key.local_name().as_ref() == local && is_namespace(&resolved, namespace) {
            // Force strict XML 1.0 decoding even though coordinates retain the
            // original lexical bytes.
            let _ = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            count = checked_add(count, 1, "content-control extension attribute count")?;
            let span = find_attr(source, begin, end, attribute.key.as_ref())?;
            if found.is_none() {
                found = Some(span);
            }
        }
    }
    Ok((found, count))
}

fn active_offsets(source: &[u8], limits: &Limits) -> Result<HashSet<usize>> {
    let mut reader = NsReader::from_reader(source);
    let mut offsets = Vec::new();
    offsets
        .try_reserve(limits.max_content_controls.min(1_024))
        .map_err(alloc("content-control MCE offsets"))?;
    let mut events = 0usize;
    let mut depth = 0usize;
    loop {
        let begin = pos(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        events = checked_add(events, 1, "content-control MCE event count")?;
        if events > limits.max_events {
            return Err(invalid("content-control MCE event limit exceeded"));
        }
        match event {
            Event::Start(element) => {
                depth = checked_add(depth, 1, "content-control MCE depth")?;
                if depth > limits.max_depth {
                    return Err(invalid("content-control MCE depth limit exceeded"));
                }
                if is_relevant(&namespace, &element) {
                    push_offset(&mut offsets, begin, limits.max_events)?;
                }
            },
            Event::Empty(element) => {
                if is_relevant(&namespace, &element) {
                    push_offset(&mut offsets, begin, limits.max_events)?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("content-control MCE depth underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if offsets.is_empty() {
        return Ok(HashSet::new());
    }
    let capabilities = content_control_capabilities();
    let selected = litchi_ooxml_common::mce::active_offsets(
        source,
        &offsets,
        &capabilities,
        &OffsetLimits {
            max_source_bytes: limits.max_input_bytes,
            max_offsets: limits.max_events,
            max_marked_bytes: limits.max_mce_marked_bytes,
            processing: litchi_ooxml_common::mce::Limits {
                max_input_bytes: limits.max_mce_marked_bytes,
                max_output_bytes: limits.max_mce_output_bytes,
                max_depth: limits.max_depth,
                max_namespace_bindings: limits.max_bindings,
                max_directive_tokens: limits.max_bindings,
                max_choices_per_alternate: limits.max_content_controls,
            },
        },
    )?;
    let mut active = HashSet::new();
    active
        .try_reserve(selected.len())
        .map_err(alloc("active content-control offsets"))?;
    for offset in selected {
        active.insert(offset as usize);
    }
    Ok(active)
}

fn is_relevant(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    let local = element.local_name();
    (is_word(namespace) && matches!(local.as_ref(), b"sdtPr" | b"id" | b"lock"))
        || ((is_word(namespace) || is_namespace(namespace, WORD_2012))
            && local.as_ref() == b"dataBinding")
}

fn is_word(namespace: &ResolveResult<'_>) -> bool {
    is_namespace(namespace, WORD) || is_namespace(namespace, WORD_STRICT)
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn push_offset(offsets: &mut Vec<u32>, offset: usize, limit: usize) -> Result<()> {
    if offsets.len() >= limit {
        return Err(invalid("content-control MCE offset limit exceeded"));
    }
    offsets
        .try_reserve(1)
        .map_err(alloc("content-control MCE offsets"))?;
    offsets.push(
        u32::try_from(offset)
            .map_err(|_| invalid("content-control source offset does not fit u32"))?,
    );
    Ok(())
}

fn find_attr(source: &[u8], start: usize, end: usize, wanted: &[u8]) -> Result<AttributeSpan> {
    let tag = source
        .get(start..end)
        .ok_or_else(|| invalid("content-control attribute span is out of bounds"))?;
    let mut cursor = tag
        .iter()
        .position(u8::is_ascii_whitespace)
        .unwrap_or(tag.len());
    while cursor < tag.len() {
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if cursor == tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let attribute_start = cursor;
        while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'=' {
            cursor += 1;
        }
        let name = &tag[attribute_start..cursor];
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("malformed content-control attribute"));
        }
        cursor += 1;
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .ok_or_else(|| invalid("unterminated content-control attribute"))?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(invalid("content-control attribute is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while tag.get(cursor).is_some_and(|value| *value != quote) {
            cursor += 1;
        }
        if cursor == tag.len() {
            return Err(invalid("unterminated content-control attribute value"));
        }
        let value_end = cursor;
        cursor += 1;
        if name == wanted {
            return Ok(AttributeSpan {
                attribute: Span::new(start + attribute_start, start + cursor)?,
                value: Span::new(start + value_start, start + value_end)?,
            });
        }
    }
    Err(invalid(
        "resolved content-control attribute has no lexical source span",
    ))
}

fn pos(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("content-control XML offset does not fit usize"))
}

fn checked_add(left: usize, right: usize, resource: &str) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid(format!("{resource} overflow")))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn alloc(resource: &'static str) -> impl FnOnce(std::collections::TryReserveError) -> Error {
    move |source| Error::Allocation { resource, source }
}
