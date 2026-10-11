//! Lossless package-part surgery for local worksheet-name dependencies.
//!
//! Known formula and reference carriers are rewritten together. Unknown
//! formula-like carriers and markup-compatibility alternatives are rejected
//! before publication instead of being guessed at.

use std::collections::{BTreeSet, TryReserveError};
use std::str::Utf8Error;

use litchi_core::xml::ReaderOrigin;
use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::encoding::{Decoder, EncodingError};
use quick_xml::events::attributes::AttrError;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use thiserror::Error;

use crate::relationships::attribute_value as relationship_attribute_value;
use crate::xml::XmlError;
use crate::xml::attributes::BytesStartExt as _;
use crate::xml::decode_xml_reference;

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const CHART: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const EXCEL_FORMULA: &[u8] = b"http://schemas.microsoft.com/office/excel/2006/main";
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const EXTENDED_PROPERTIES: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const STRICT_EXTENDED_PROPERTIES: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/extended-properties";
const PROPERTY_TYPES: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
const STRICT_PROPERTY_TYPES: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes";
const MAX_REFERENCE_DEPTH: usize = 256;
const MAX_CAPTURED_TEXT_BYTES: usize = 8 * 1024 * 1024;

/// One simultaneous source-to-final sheet spelling.
#[derive(Debug, Clone, Copy)]
pub struct Rename<'a> {
    /// Source sheet spelling, compared with the caller's identity policy.
    pub before: &'a str,
    /// Final sheet spelling written to recognized carriers.
    pub after: &'a str,
}

/// Why a local sheet-name dependency could not be edited safely.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
pub enum BlockReason {
    /// The matching carrier is inside an MCE `AlternateContent` branch.
    #[error("markup-compatibility alternatives are not editable")]
    MarkupCompatibility,
    /// A formula-like or name-like carrier is not modeled by this scanner.
    #[error("the reference carrier is not modeled")]
    UnmodeledReference,
}

/// A bounded scanner resource was exhausted before a rewrite could be
/// published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
pub enum Limit {
    /// Element nesting exceeded the scanner's structural bound.
    #[error("sheet-reference XML nesting exceeds {limit} levels")]
    NestingDepth {
        /// Maximum admitted element depth.
        limit: usize,
    },
    /// Captured formula/title text exceeded the scanner's memory bound.
    #[error("sheet-reference text exceeds the {limit}-byte edit budget")]
    CapturedTextBytes {
        /// Maximum admitted captured bytes.
        limit: usize,
    },
}

/// Failure while scanning or rewriting a shared OOXML part.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The XML decoder rejected a character or entity reference.
    #[error(transparent)]
    Xml(#[from] XmlError),
    /// The XML event reader rejected the part's syntax or namespace stream.
    #[error("XML error in {context}: {source}")]
    Parse {
        /// Scanner operation that was decoding the event stream.
        context: &'static str,
        /// Underlying quick-xml parse failure.
        #[source]
        source: quick_xml::Error,
    },
    /// A start-tag attribute could not be decoded or was duplicated.
    #[error("attribute error in {context}: {source}")]
    Attribute {
        /// Scanner operation that was reading the start tag.
        context: &'static str,
        /// Underlying quick-xml attribute failure.
        #[source]
        source: AttrError,
    },
    /// Event text could not be decoded using the XML reader's encoding.
    #[error("encoding error in {context}: {source}")]
    Encoding {
        /// Text carrier being decoded.
        context: &'static str,
        /// Underlying quick-xml encoding failure.
        #[source]
        source: EncodingError,
    },
    /// A namespace or local name was not valid UTF-8.
    #[error("UTF-8 error in {context}: {source}")]
    Utf8 {
        /// Name being decoded.
        context: &'static str,
        /// Underlying UTF-8 failure.
        #[source]
        source: Utf8Error,
    },
    /// The part is well-formed enough to scan but violates an edit invariant.
    #[error("invalid sheet-reference XML: {0}")]
    Invalid(String),
    /// A recognized local dependency cannot be rewritten under the safe policy.
    #[error("cannot rewrite sheet-name dependency: {0}")]
    Blocked(#[from] BlockReason),
    /// A bounded scanner resource was exhausted.
    #[error(transparent)]
    Limit(#[from] Limit),
    /// A bounded replacement buffer could not be reserved.
    #[error("could not reserve memory for OOXML sheet-reference {resource}: {source}")]
    Allocation {
        /// Resource whose bounded allocation failed.
        resource: &'static str,
        /// Underlying allocator failure.
        #[source]
        source: TryReserveError,
    },
}

/// Result of a shared sheet-name rewrite.
pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Mode {
    Formula,
    Suspicious,
    SheetTitle(usize),
    NamedRangeTitle,
    Other,
}

#[derive(Debug)]
struct Frame {
    mode: Mode,
    inner_start: usize,
    text: String,
    has_markup: bool,
    alternate: bool,
    titles: bool,
}

#[derive(Debug, Clone, Copy)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Debug)]
struct Replacement {
    span: Span,
    bytes: Vec<u8>,
}

fn apply_replacements(
    content: &[u8],
    mut replacements: Vec<Replacement>,
    resource: &'static str,
) -> Result<Option<Vec<u8>>> {
    if replacements.is_empty() {
        return Ok(None);
    }
    replacements.sort_unstable_by_key(|replacement| (replacement.span.start, replacement.span.end));
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping extended-properties replacements"));
    }
    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.span.end - replacement.span.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("extended-properties replacement output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation { resource, source })?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(Some(output))
}

#[derive(Debug)]
struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug)]
struct Tag {
    name: Box<str>,
    attributes: Box<[Attribute]>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TitleReorderKind {
    Properties,
    TitlesOfParts,
    TitleVector,
    Title(usize),
    Other,
}

#[derive(Debug)]
struct TitleReorderFrame {
    kind: TitleReorderKind,
    name: Box<[u8]>,
    inner_start: usize,
    text: String,
    has_markup: bool,
}

/// Rewrite one XML part, returning `None` without allocating an output buffer
/// when no recognized local reference changes.
///
/// `sheet_titles` is the complete catalog spelling used to validate
/// `docProps/app.xml` title entries. `equivalent` supplies the owning format's
/// sheet identity policy; the scanner never assumes that case folding rules
/// are shared by every spreadsheet format.
///
/// # Example
///
/// ```
/// use litchi_ooxml_common::xml::sheet_name::{rewrite, Rename};
/// let xml = br#"<c:f xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">Data!A1</c:f>"#;
/// let changed = rewrite(xml, &[Rename { before: "Data", after: "Input" }],
///     &[], |left, right| left.eq_ignore_ascii_case(right))?;
/// assert!(String::from_utf8(changed.expect("changed reference"))?.contains("Input!A1"));
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
///
/// # Errors
///
/// Returns a typed [`enum@Error`] when the part is malformed, exceeds a bounded
/// scan, or contains an unmodeled formula-like carrier or MCE alternative.
pub fn rewrite<F>(
    content: &[u8],
    renames: &[Rename<'_>],
    sheet_titles: &[&str],
    equivalent: F,
) -> Result<Option<Vec<u8>>>
where
    F: Fn(&str, &str) -> bool,
{
    if renames.is_empty() {
        return Ok(None);
    }
    // Branches are recognized on namespace values as written; a binding
    // spelled with references could hide a reference inside one (record 0011).
    if crate::mce::binds_namespace_by_reference(content) {
        return Err(Error::Blocked(BlockReason::MarkupCompatibility));
    }
    let mut reader = NsReader::from_reader(content);
    let origin = ReaderOrigin::of(content);
    let mut stack = Vec::<Frame>::new();
    let mut replacements = Vec::new();
    let mut saw_titles = false;
    let mut title_entries = 0usize;
    let mut captured_text_bytes = 0usize;

    loop {
        let event_start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|source| Error::Parse {
                context: "sheet-reference XML scan",
                source,
            })?
            .into_owned();
        let event_end = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_REFERENCE_DEPTH {
                    return Err(Error::Limit(Limit::NestingDepth {
                        limit: MAX_REFERENCE_DEPTH,
                    }));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let alternate = stack.last().is_some_and(|frame| frame.alternate)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                let name = element.name();
                let local_name = name.local_name();
                let local =
                    std::str::from_utf8(local_name.as_ref()).map_err(|source| Error::Utf8 {
                        context: "sheet-reference element name",
                        source,
                    })?;
                let starts_titles =
                    local == "TitlesOfParts" && is_namespace(&namespace, &[EXTENDED_PROPERTIES]);
                if local == "TitlesOfParts" && !starts_titles && !sheet_titles.is_empty() {
                    return Err(unsupported(alternate));
                }
                if starts_titles && std::mem::replace(&mut saw_titles, true) {
                    return Err(unsupported(alternate));
                }
                let titles = stack.last().is_some_and(|frame| frame.titles) || starts_titles;
                let title_index =
                    (titles && local == "lpstr" && is_namespace(&namespace, &[PROPERTY_TYPES]))
                        .then(|| {
                            let index = title_entries;
                            title_entries = title_entries.saturating_add(1);
                            index
                        });
                let element_mode = mode(local, title_index, &namespace, sheet_titles.len());
                let spreadsheet = is_namespace(&namespace, &[SML, STRICT_SML]);
                attribute_replacement(
                    &element,
                    decoder,
                    &resolver,
                    local,
                    spreadsheet,
                    false,
                    alternate,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                    renames,
                    &equivalent,
                    &mut replacements,
                )?;
                stack.push(Frame {
                    mode: element_mode,
                    inner_start: event_end,
                    text: String::new(),
                    has_markup: false,
                    alternate,
                    titles,
                });
            },
            Event::Empty(element) => {
                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let alternate = stack.last().is_some_and(|frame| frame.alternate)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                let name = element.name();
                let local_name = name.local_name();
                let local =
                    std::str::from_utf8(local_name.as_ref()).map_err(|source| Error::Utf8 {
                        context: "sheet-reference element name",
                        source,
                    })?;
                let starts_titles =
                    local == "TitlesOfParts" && is_namespace(&namespace, &[EXTENDED_PROPERTIES]);
                if local == "TitlesOfParts" && !starts_titles && !sheet_titles.is_empty() {
                    return Err(unsupported(alternate));
                }
                if starts_titles && std::mem::replace(&mut saw_titles, true) {
                    return Err(unsupported(alternate));
                }
                let titles = stack.last().is_some_and(|frame| frame.titles) || starts_titles;
                let title_index =
                    (titles && local == "lpstr" && is_namespace(&namespace, &[PROPERTY_TYPES]))
                        .then(|| {
                            let index = title_entries;
                            title_entries = title_entries.saturating_add(1);
                            index
                        });
                let element_mode = mode(local, title_index, &namespace, sheet_titles.len());
                let spreadsheet = is_namespace(&namespace, &[SML, STRICT_SML]);
                attribute_replacement(
                    &element,
                    decoder,
                    &resolver,
                    local,
                    spreadsheet,
                    true,
                    alternate,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                    renames,
                    &equivalent,
                    &mut replacements,
                )?;
                if matches!(element_mode, Mode::SheetTitle(_)) {
                    finish_frame(
                        Frame {
                            mode: element_mode,
                            inner_start: event_end,
                            text: String::new(),
                            has_markup: false,
                            alternate,
                            titles,
                        },
                        event_end,
                        renames,
                        sheet_titles,
                        &equivalent,
                        &mut replacements,
                    )?;
                }
            },
            Event::Text(text) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "formula text",
                        source,
                    })?;
                    push_text(&mut stack, &text, &mut captured_text_bytes)?;
                }
            },
            Event::CData(text) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "formula CDATA",
                        source,
                    })?;
                    push_text(&mut stack, &text, &mut captured_text_bytes)?;
                }
            },
            Event::GeneralRef(reference) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = decode_xml_reference(&reference)?;
                    push_text(&mut stack, &text, &mut captured_text_bytes)?;
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::DocType(_) | Event::Decl(_) => {
                if let Some(parent) = stack.last_mut()
                    && parent.mode != Mode::Other
                {
                    parent.has_markup = true;
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    invalid("sheet-reference XML scan found an unmatched closing tag")
                })?;
                captured_text_bytes = captured_text_bytes
                    .checked_sub(frame.text.len())
                    .ok_or_else(|| invalid("sheet-reference text budget underflow"))?;
                finish_frame(
                    frame,
                    event_start,
                    renames,
                    sheet_titles,
                    &equivalent,
                    &mut replacements,
                )?;
            },
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("sheet-reference XML scan ended inside an element"));
    }
    if saw_titles && title_entries < sheet_titles.len() {
        return Err(unsupported(false));
    }
    if replacements.is_empty() {
        return Ok(None);
    }
    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping sheet-reference replacements"));
    }
    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.span.end - replacement.span.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("sheet-reference output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "sheet-reference output",
            source,
        })?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(Some(output))
}

/// Reorder the leading catalog titles in an extended-properties part.
///
/// `original_titles` is the complete catalog in its source order, while each
/// entry in `order` names the source catalog slot that should occupy the final
/// position. Only the first `original_titles.len()` `vt:lpstr` entries in the
/// one recognized `TitlesOfParts` vector are changed. Following entries are
/// named-range metadata and remain byte-exact, including formula-like text.
/// The XML namespace URIs are matched rather than their prefixes, so producer
/// chosen prefixes remain intact.
///
/// # Example
///
/// ```
/// use litchi_ooxml_common::xml::sheet_name::reorder_titles;
/// let xml = br#"<p:Properties xmlns:p="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:v="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><p:TitlesOfParts><v:vector><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr></v:vector></p:TitlesOfParts></p:Properties>"#;
/// let changed = reorder_titles(xml, &["One", "Two"], &[1, 0], |left, right| {
///     left.eq_ignore_ascii_case(right)
/// })?;
/// let bytes = changed.expect("changed titles");
/// let changed = std::str::from_utf8(&bytes)?;
/// assert!(changed.contains("<v:lpstr>Two</v:lpstr><v:lpstr>One</v:lpstr>"));
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
///
/// # Errors
///
/// Returns a typed [`enum@Error`] when `order` is not a full permutation, the
/// catalog titles do not match the immutable source catalog, the title context
/// is ambiguous, or the XML contains markup-compatibility content.
pub fn reorder_titles<F>(
    content: &[u8],
    original_titles: &[&str],
    order: &[usize],
    equivalent: F,
) -> Result<Option<Vec<u8>>>
where
    F: Fn(&str, &str) -> bool,
{
    let identity = validate_title_order(original_titles.len(), order)?;
    // Preserve the exact no-op contract before looking at the source bytes.
    // This also lets callers keep malformed optional metadata untouched when
    // the workbook order did not change.
    if identity {
        return Ok(None);
    }

    if crate::mce::binds_namespace_by_reference(content) {
        return Err(Error::Blocked(BlockReason::MarkupCompatibility));
    }

    let mut reader = NsReader::from_reader(content);
    let origin = ReaderOrigin::of(content);
    let mut stack = Vec::<TitleReorderFrame>::new();
    let mut replacements = Vec::<Replacement>::new();
    replacements
        .try_reserve_exact(original_titles.len())
        .map_err(|source| Error::Allocation {
            resource: "sheet-title replacements",
            source,
        })?;
    let mut root_seen = false;
    let mut saw_titles = false;
    let mut saw_vector = false;
    let mut title_entries = 0usize;
    let mut captured_text_bytes = 0usize;

    loop {
        let event_start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|source| Error::Parse {
                context: "sheet-title permutation XML scan",
                source,
            })?
            .into_owned();
        let event_end = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_REFERENCE_DEPTH {
                    return Err(Error::Limit(Limit::NestingDepth {
                        limit: MAX_REFERENCE_DEPTH,
                    }));
                }
                if is_mce_namespace(&namespace) {
                    return Err(Error::Blocked(BlockReason::MarkupCompatibility));
                }
                validate_title_attributes(&element, decoder, &resolver)?;
                if stack.is_empty() {
                    if root_seen {
                        return Err(invalid(
                            "sheet-title permutation XML contains multiple roots",
                        ));
                    }
                    root_seen = true;
                    if !is_extended_element(&namespace, &element, b"Properties") {
                        return Err(invalid(
                            "sheet-title permutation XML has no extended-properties root",
                        ));
                    }
                }

                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let parent_kind = stack.last().map(|frame| frame.kind);
                let local_name = element.name().local_name();
                let local = local_name.as_ref();
                let kind = reorder_kind(parent_kind, &namespace, &element, &mut title_entries);
                reject_ambiguous_title_element(local, parent_kind, kind)?;
                if kind == TitleReorderKind::TitlesOfParts {
                    if std::mem::replace(&mut saw_titles, true) {
                        return Err(invalid("duplicate extended-properties TitlesOfParts"));
                    }
                }
                if kind == TitleReorderKind::TitleVector {
                    if std::mem::replace(&mut saw_vector, true) {
                        return Err(invalid(
                            "duplicate extended-properties TitlesOfParts vector",
                        ));
                    }
                }
                stack.push(TitleReorderFrame {
                    kind,
                    name: element.name().as_ref().to_vec().into_boxed_slice(),
                    inner_start: event_end,
                    text: String::new(),
                    has_markup: false,
                });
            },
            Event::Empty(element) => {
                if is_mce_namespace(&namespace) {
                    return Err(Error::Blocked(BlockReason::MarkupCompatibility));
                }
                validate_title_attributes(&element, decoder, &resolver)?;
                if stack.is_empty() {
                    if root_seen {
                        return Err(invalid(
                            "sheet-title permutation XML contains multiple roots",
                        ));
                    }
                    root_seen = true;
                    if !is_extended_element(&namespace, &element, b"Properties") {
                        return Err(invalid(
                            "sheet-title permutation XML has no extended-properties root",
                        ));
                    }
                }
                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let parent_kind = stack.last().map(|frame| frame.kind);
                let local_name = element.name().local_name();
                let local = local_name.as_ref();
                let kind = reorder_kind(parent_kind, &namespace, &element, &mut title_entries);
                reject_ambiguous_title_element(local, parent_kind, kind)?;
                if kind == TitleReorderKind::TitlesOfParts {
                    if std::mem::replace(&mut saw_titles, true) {
                        return Err(invalid("duplicate extended-properties TitlesOfParts"));
                    }
                }
                if kind == TitleReorderKind::TitleVector {
                    if std::mem::replace(&mut saw_vector, true) {
                        return Err(invalid(
                            "duplicate extended-properties TitlesOfParts vector",
                        ));
                    }
                }
                if let TitleReorderKind::Title(index) = kind {
                    let replacement = finish_reordered_title(
                        index,
                        "",
                        true,
                        Span {
                            start: event_end,
                            end: event_end,
                        },
                        content,
                        original_titles,
                        order,
                        &equivalent,
                    )?;
                    if let Some(replacement) = replacement {
                        replacements.push(replacement);
                    }
                }
            },
            Event::Text(text) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(frame.kind, TitleReorderKind::Title(_))
                {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "extended-properties title text",
                        source,
                    })?;
                    captured_text_bytes =
                        captured_text_bytes
                            .checked_add(text.len())
                            .ok_or(Error::Limit(Limit::CapturedTextBytes {
                                limit: MAX_CAPTURED_TEXT_BYTES,
                            }))?;
                    if captured_text_bytes > MAX_CAPTURED_TEXT_BYTES {
                        return Err(Error::Limit(Limit::CapturedTextBytes {
                            limit: MAX_CAPTURED_TEXT_BYTES,
                        }));
                    }
                    frame.text.push_str(&text);
                } else if stack.is_empty() {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "sheet-title XML text outside root",
                        source,
                    })?;
                    if !is_xml_whitespace(&text) {
                        return Err(invalid(
                            "sheet-title permutation XML has non-whitespace text outside its root",
                        ));
                    }
                }
            },
            Event::CData(text) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(frame.kind, TitleReorderKind::Title(_))
                {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "extended-properties title CDATA",
                        source,
                    })?;
                    captured_text_bytes =
                        captured_text_bytes
                            .checked_add(text.len())
                            .ok_or(Error::Limit(Limit::CapturedTextBytes {
                                limit: MAX_CAPTURED_TEXT_BYTES,
                            }))?;
                    if captured_text_bytes > MAX_CAPTURED_TEXT_BYTES {
                        return Err(Error::Limit(Limit::CapturedTextBytes {
                            limit: MAX_CAPTURED_TEXT_BYTES,
                        }));
                    }
                    frame.text.push_str(&text);
                } else if stack.is_empty() {
                    return Err(invalid(
                        "sheet-title permutation XML has CDATA outside its root",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(frame.kind, TitleReorderKind::Title(_))
                {
                    let text = decode_xml_reference(&reference)?;
                    captured_text_bytes =
                        captured_text_bytes
                            .checked_add(text.len())
                            .ok_or(Error::Limit(Limit::CapturedTextBytes {
                                limit: MAX_CAPTURED_TEXT_BYTES,
                            }))?;
                    if captured_text_bytes > MAX_CAPTURED_TEXT_BYTES {
                        return Err(Error::Limit(Limit::CapturedTextBytes {
                            limit: MAX_CAPTURED_TEXT_BYTES,
                        }));
                    }
                    frame.text.push_str(&text);
                } else if stack.is_empty() {
                    return Err(invalid(
                        "sheet-title permutation XML has an entity outside its root",
                    ));
                }
            },
            Event::End(end) => {
                let frame = stack.pop().ok_or_else(|| {
                    invalid("sheet-title permutation XML found an unmatched closing tag")
                })?;
                if frame.name.as_ref() != end.name().as_ref() {
                    return Err(invalid(
                        "sheet-title permutation XML closing tag does not match its start tag",
                    ));
                }
                captured_text_bytes = captured_text_bytes
                    .checked_sub(frame.text.len())
                    .ok_or_else(|| invalid("sheet-title text budget underflow"))?;
                if let TitleReorderKind::Title(index) = frame.kind {
                    let replacement = finish_reordered_title(
                        index,
                        &frame.text,
                        frame.has_markup,
                        Span {
                            start: frame.inner_start,
                            end: event_start,
                        },
                        content,
                        original_titles,
                        order,
                        &equivalent,
                    )?;
                    if let Some(replacement) = replacement {
                        replacements.push(replacement);
                    }
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::DocType(_) | Event::Decl(_) => {
                if let Some(parent) = stack.last_mut()
                    && matches!(parent.kind, TitleReorderKind::Title(_))
                {
                    parent.has_markup = true;
                }
            },
            Event::Eof => break,
        }
    }

    if !root_seen || !stack.is_empty() {
        return Err(invalid(
            "sheet-title permutation XML ended before a complete root",
        ));
    }
    // TitlesOfParts is optional extended-property metadata. Absence means
    // there is no cached catalog order to update; a present but malformed
    // container still refuses the rewrite below.
    if !saw_titles {
        return Ok(None);
    }
    if !saw_vector {
        return Err(invalid(
            "extended-properties TitlesOfParts title vector is missing",
        ));
    }
    if title_entries < original_titles.len() {
        return Err(invalid(format!(
            "extended-properties TitlesOfParts has {title_entries} titles but the catalog has {}",
            original_titles.len()
        )));
    }
    if replacements.is_empty() {
        return Ok(None);
    }
    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping sheet-title replacements"));
    }
    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.span.end - replacement.span.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("sheet-title permutation output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "reordered extended-properties titles",
            source,
        })?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(Some(output))
}

fn scan_insert_metadata(content: &[u8]) -> Result<Option<InsertMetadata>> {
    if crate::mce::binds_namespace_by_reference(content) {
        return Err(Error::Blocked(BlockReason::MarkupCompatibility));
    }

    let mut reader = NsReader::from_reader(content);
    let origin = ReaderOrigin::of(content);
    let mut stack = Vec::<InsertFrame>::new();
    let mut titles = Vec::<InsertTitle>::new();
    let mut heading_pairs = Vec::<InsertHeadingPair>::new();
    let mut title_vector = None::<InsertVector>;
    let mut heading_vector = None::<InsertVector>;
    let mut saw_titles_container = false;
    let mut saw_heading_container = false;
    let mut pending_heading_label = None::<String>;
    let mut heading_variant_count = 0usize;
    let mut root_seen = false;
    let mut captured_text_bytes = 0usize;
    let mut title_index = 0usize;

    loop {
        let event_start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|source| Error::Parse {
                context: "extended-properties insertion XML scan",
                source,
            })?
            .into_owned();
        let event_end = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_REFERENCE_DEPTH {
                    return Err(Error::Limit(Limit::NestingDepth {
                        limit: MAX_REFERENCE_DEPTH,
                    }));
                }
                if is_mce_namespace(&namespace) {
                    return Err(Error::Blocked(BlockReason::MarkupCompatibility));
                }
                validate_title_attributes(&element, decoder, &resolver)?;
                if stack.is_empty() {
                    if root_seen {
                        return Err(invalid(
                            "extended-properties insertion XML contains multiple roots",
                        ));
                    }
                    root_seen = true;
                    if !is_extended_element(&namespace, &element, b"Properties") {
                        return Err(invalid(
                            "extended-properties insertion XML has no extended-properties root",
                        ));
                    }
                }
                if let Some(parent) = stack.last_mut() {
                    if parent.kind == InsertKind::Title
                        || matches!(
                            parent.kind,
                            InsertKind::HeadingLabel | InsertKind::HeadingCount
                        )
                    {
                        parent.has_markup = true;
                    }
                }
                let parent_kind = stack.last().map(|frame| frame.kind);
                let kind = insert_kind(parent_kind, &namespace, &element, &mut title_index);
                reject_insert_element(parent_kind, kind, element.name().local_name().as_ref())?;
                if kind == InsertKind::Titles {
                    if std::mem::replace(&mut saw_titles_container, true) {
                        return Err(invalid(
                            "duplicate extended-properties TitlesOfParts containers",
                        ));
                    }
                }
                if kind == InsertKind::HeadingPairs {
                    if std::mem::replace(&mut saw_heading_container, true) {
                        return Err(invalid(
                            "duplicate extended-properties HeadingPairs containers",
                        ));
                    }
                }
                if kind == InsertKind::TitleVector {
                    if title_vector.is_some() {
                        return Err(invalid(
                            "duplicate extended-properties TitlesOfParts vectors",
                        ));
                    }
                }
                if kind == InsertKind::HeadingVector {
                    if heading_vector.is_some() {
                        return Err(invalid(
                            "duplicate extended-properties HeadingPairs vectors",
                        ));
                    }
                }
                let vector = matches!(kind, InsertKind::TitleVector | InsertKind::HeadingVector)
                    .then(|| {
                        parse_insert_vector(
                            content,
                            event_start,
                            event_end,
                            &element,
                            &namespace,
                            decoder,
                        )
                    })
                    .transpose()?;
                let title = if kind == InsertKind::Title {
                    Some(insert_title_name(&element, &namespace, &resolver)?)
                } else {
                    None
                };
                stack.push(InsertFrame {
                    kind,
                    name: element.name().as_ref().to_vec().into_boxed_slice(),
                    inner_start: event_end,
                    start: event_start,
                    text: String::new(),
                    has_markup: false,
                    vector,
                    title,
                    heading: (kind == InsertKind::HeadingVariant)
                        .then(InsertHeadingBuilder::default),
                });
            },
            Event::Empty(element) => {
                if is_mce_namespace(&namespace) {
                    return Err(Error::Blocked(BlockReason::MarkupCompatibility));
                }
                validate_title_attributes(&element, decoder, &resolver)?;
                if stack.is_empty() {
                    if root_seen {
                        return Err(invalid(
                            "extended-properties insertion XML contains multiple roots",
                        ));
                    }
                    root_seen = true;
                    if !is_extended_element(&namespace, &element, b"Properties") {
                        return Err(invalid(
                            "extended-properties insertion XML has no extended-properties root",
                        ));
                    }
                }
                if let Some(parent) = stack.last_mut() {
                    if parent.kind == InsertKind::Title
                        || matches!(
                            parent.kind,
                            InsertKind::HeadingLabel | InsertKind::HeadingCount
                        )
                    {
                        parent.has_markup = true;
                    }
                }
                let parent_kind = stack.last().map(|frame| frame.kind);
                let kind = insert_kind(parent_kind, &namespace, &element, &mut title_index);
                reject_insert_element(parent_kind, kind, element.name().local_name().as_ref())?;
                if kind == InsertKind::Titles {
                    if std::mem::replace(&mut saw_titles_container, true) {
                        return Err(invalid(
                            "duplicate extended-properties TitlesOfParts containers",
                        ));
                    }
                }
                if kind == InsertKind::HeadingPairs {
                    if std::mem::replace(&mut saw_heading_container, true) {
                        return Err(invalid(
                            "duplicate extended-properties HeadingPairs containers",
                        ));
                    }
                }
                if matches!(kind, InsertKind::TitleVector | InsertKind::HeadingVector) {
                    return Err(Error::Blocked(BlockReason::UnmodeledReference));
                }
                if kind == InsertKind::Title {
                    let title = insert_title_name(&element, &namespace, &resolver)?;
                    titles.push(InsertTitle {
                        start: event_start,
                        end: event_end,
                        text: String::new(),
                        has_markup: false,
                        name: title.name,
                        needs_binding: title.needs_binding,
                        namespace: title.namespace,
                    });
                } else if matches!(kind, InsertKind::HeadingLabel | InsertKind::HeadingCount) {
                    return Err(invalid(
                        "extended-properties heading pair entries cannot be empty",
                    ));
                } else if parent_kind == Some(InsertKind::HeadingVector)
                    || parent_kind == Some(InsertKind::HeadingPairs)
                    || parent_kind == Some(InsertKind::Titles)
                    || parent_kind == Some(InsertKind::TitleVector)
                {
                    return Err(invalid(
                        "extended-properties metadata container has an unexpected empty child",
                    ));
                }
            },
            Event::Text(text) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(
                        frame.kind,
                        InsertKind::Title | InsertKind::HeadingLabel | InsertKind::HeadingCount
                    )
                {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "extended-properties insertion text",
                        source,
                    })?;
                    push_insert_text(&mut frame.text, &text, &mut captured_text_bytes)?;
                } else if stack.is_empty() {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "extended-properties insertion text outside root",
                        source,
                    })?;
                    if !is_xml_whitespace(&text) {
                        return Err(invalid(
                            "extended-properties insertion XML has non-whitespace text outside its root",
                        ));
                    }
                }
            },
            Event::CData(text) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(
                        frame.kind,
                        InsertKind::Title | InsertKind::HeadingLabel | InsertKind::HeadingCount
                    )
                {
                    let text = text.decode().map_err(|source| Error::Encoding {
                        context: "extended-properties insertion CDATA",
                        source,
                    })?;
                    push_insert_text(&mut frame.text, &text, &mut captured_text_bytes)?;
                } else if stack.is_empty() {
                    return Err(invalid(
                        "extended-properties insertion XML has CDATA outside its root",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(
                        frame.kind,
                        InsertKind::Title | InsertKind::HeadingLabel | InsertKind::HeadingCount
                    )
                {
                    let text = decode_xml_reference(&reference)?;
                    push_insert_text(&mut frame.text, &text, &mut captured_text_bytes)?;
                } else if stack.is_empty() {
                    return Err(invalid(
                        "extended-properties insertion XML has an entity outside its root",
                    ));
                }
            },
            Event::End(end) => {
                let frame = stack.pop().ok_or_else(|| {
                    invalid("extended-properties insertion XML found an unmatched closing tag")
                })?;
                if frame.name.as_ref() != end.name().as_ref() {
                    return Err(invalid(
                        "extended-properties insertion XML closing tag does not match its start tag",
                    ));
                }
                captured_text_bytes = captured_text_bytes
                    .checked_sub(frame.text.len())
                    .ok_or_else(|| {
                        invalid("extended-properties insertion text budget underflow")
                    })?;
                match frame.kind {
                    InsertKind::Title => {
                        let title = frame.title.ok_or_else(|| {
                            invalid("extended-properties title lost its namespace metadata")
                        })?;
                        titles.push(InsertTitle {
                            start: frame.start,
                            end: event_end,
                            text: frame.text,
                            has_markup: frame.has_markup,
                            name: title.name,
                            needs_binding: title.needs_binding,
                            namespace: title.namespace,
                        });
                    },
                    InsertKind::TitleVector => {
                        let mut vector = frame.vector.ok_or_else(|| {
                            invalid("extended-properties title vector lost its attributes")
                        })?;
                        vector.close_start = event_start;
                        title_vector = Some(vector);
                    },
                    InsertKind::HeadingVector => {
                        let mut vector = frame.vector.ok_or_else(|| {
                            invalid("extended-properties heading vector lost its attributes")
                        })?;
                        vector.close_start = event_start;
                        heading_vector = Some(vector);
                    },
                    InsertKind::HeadingLabel => {
                        if frame.has_markup {
                            return Err(invalid(
                                "extended-properties heading label contains nested markup",
                            ));
                        }
                        let parent = stack.last_mut().ok_or_else(|| {
                            invalid("extended-properties heading label has no variant parent")
                        })?;
                        let heading = parent.heading.as_mut().ok_or_else(|| {
                            invalid("extended-properties heading label has no variant parent")
                        })?;
                        if heading.label.replace(frame.text).is_some() {
                            return Err(invalid(
                                "extended-properties heading pair has duplicate labels",
                            ));
                        }
                    },
                    InsertKind::HeadingCount => {
                        if frame.has_markup {
                            return Err(invalid(
                                "extended-properties heading count contains nested markup",
                            ));
                        }
                        let span = decimal_text_span(content, frame.inner_start, event_start)?;
                        let count = frame.text.trim().parse::<usize>().map_err(|_| {
                            invalid("extended-properties heading count is not a decimal integer")
                        })?;
                        let parent = stack.last_mut().ok_or_else(|| {
                            invalid("extended-properties heading count has no variant parent")
                        })?;
                        let heading = parent.heading.as_mut().ok_or_else(|| {
                            invalid("extended-properties heading count has no variant parent")
                        })?;
                        if heading.count.replace((count, span)).is_some() {
                            return Err(invalid(
                                "extended-properties heading pair has duplicate counts",
                            ));
                        }
                    },
                    InsertKind::HeadingVariant => {
                        let heading = frame.heading.ok_or_else(|| {
                            invalid("extended-properties heading variant lost its state")
                        })?;
                        heading_variant_count =
                            heading_variant_count.checked_add(1).ok_or_else(|| {
                                invalid("extended-properties heading variant count overflow")
                            })?;
                        match (heading.label, heading.count) {
                            (Some(label), Some((count, span))) => {
                                heading_pairs.push(InsertHeadingPair { label, count, span });
                            },
                            (Some(label), None) => {
                                if pending_heading_label.replace(label).is_some() {
                                    return Err(invalid(
                                        "extended-properties HeadingPairs has consecutive labels without counts",
                                    ));
                                }
                            },
                            (None, Some((count, span))) => {
                                let label = pending_heading_label.take().ok_or_else(|| {
                                    invalid("extended-properties HeadingPairs count has no label")
                                })?;
                                heading_pairs.push(InsertHeadingPair { label, count, span });
                            },
                            (None, None) => {
                                return Err(invalid(
                                    "extended-properties heading variant has neither label nor count",
                                ));
                            },
                        }
                    },
                    _ => {},
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::DocType(_) | Event::Decl(_) => {
                if let Some(parent) = stack.last_mut()
                    && matches!(
                        parent.kind,
                        InsertKind::Title | InsertKind::HeadingLabel | InsertKind::HeadingCount
                    )
                {
                    parent.has_markup = true;
                }
            },
            Event::Eof => break,
        }
    }
    if !root_seen || !stack.is_empty() {
        return Err(invalid(
            "extended-properties insertion XML ended before a complete root",
        ));
    }
    if pending_heading_label.is_some() {
        return Err(invalid(
            "extended-properties HeadingPairs ends with a label without a count",
        ));
    }

    let Some(title_vector) = title_vector else {
        if !saw_titles_container && !saw_heading_container {
            return Ok(None);
        }
        return Err(invalid(
            "extended-properties HeadingPairs exists without TitlesOfParts",
        ));
    };
    let Some(heading_vector) = heading_vector else {
        return Err(invalid(
            "extended-properties TitlesOfParts exists without HeadingPairs",
        ));
    };
    Ok(Some(InsertMetadata {
        titles,
        heading_pairs,
        title_vector,
        heading_vector,
        heading_variant_count,
    }))
}

/// Insert one title into the leading catalog in an extended-properties part.
///
/// `original_titles` is the immutable catalog before the insertion and
/// `position` is the destination slot in that catalog (`0..=len`). The
/// existing leading title text is checked with `equivalent`; entries after
/// that prefix are treated as named-range metadata and remain byte-exact.
/// Both the `TitlesOfParts` vector size and the matching `HeadingPairs`
/// catalog count are updated. The helper recognizes the aggregate `Sheet`
/// category used by XLSB as well as the standard `Worksheets` category.
///
/// Optional metadata is left untouched when both recognized containers are
/// absent. A partially present or ambiguous layout is rejected so callers do
/// not silently manufacture application metadata.
/// # Example
///
/// ```
/// use litchi_ooxml_common::xml::sheet_name::insert_title;
/// let xml = br#"<p:Properties xmlns:p="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:v="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><p:HeadingPairs><v:vector size="2" baseType="variant"><v:variant><v:lpstr>Worksheets</v:lpstr></v:variant><v:variant><v:i4>2</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="2" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr></v:vector></p:TitlesOfParts></p:Properties>"#;
/// let changed = insert_title(xml, &["One", "Two"], 1, "New", |left, right| left == right)?;
/// let changed = std::str::from_utf8(changed.as_deref().expect("metadata present"))?;
/// assert!(changed.contains("<v:i4>3</v:i4>"));
/// assert!(changed.contains("<v:lpstr>One</v:lpstr><v:lpstr>New</v:lpstr><v:lpstr>Two</v:lpstr>"));
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn insert_title<F>(
    content: &[u8],
    original_titles: &[&str],
    insertion_position: usize,
    name: &str,
    equivalent: F,
) -> Result<Option<Vec<u8>>>
where
    F: Fn(&str, &str) -> bool,
{
    if insertion_position > original_titles.len() {
        return Err(invalid(format!(
            "extended-properties title insertion position {insertion_position} exceeds catalog length {}",
            original_titles.len()
        )));
    }
    let Some(metadata) = scan_insert_metadata(content)? else {
        return Ok(None);
    };
    let InsertMetadata {
        titles,
        heading_pairs,
        title_vector,
        heading_vector,
        heading_variant_count,
    } = metadata;
    validate_insert_layout(
        &title_vector,
        &heading_vector,
        &titles,
        &heading_pairs,
        heading_variant_count,
        original_titles,
        &equivalent,
    )?;

    let title_name = titles.first().map_or_else(
        || qualified_lpstr_name(&title_vector.name),
        |title| title.name.clone(),
    );
    let title_namespace = titles.first().map_or_else(
        || title_vector.namespace.clone(),
        |title| title.namespace.clone(),
    );
    let needs_binding = titles.first().is_some_and(|title| title.needs_binding);
    let new_title = write_insert_title(&title_name, &title_namespace, needs_binding, name)?;
    let insert_at = titles.get(insertion_position).map_or_else(
        || {
            titles
                .last()
                .map_or(title_vector.close_start, |title| title.end)
        },
        |title| title.start,
    );

    let new_size = titles
        .len()
        .checked_add(1)
        .ok_or_else(|| invalid("extended-properties title vector size overflow"))?;
    let size_span = title_vector.size.map(|(_, span)| span).ok_or_else(|| {
        invalid("extended-properties TitlesOfParts vector has no checked size attribute")
    })?;
    let mut replacements = Vec::<Replacement>::new();
    replacements
        .try_reserve_exact(3)
        .map_err(|source| Error::Allocation {
            resource: "inserted extended-properties title replacements",
            source,
        })?;
    replacements.push(Replacement {
        span: size_span,
        bytes: new_size.to_string().into_bytes(),
    });
    let catalog_pair = heading_pairs
        .iter()
        .find(|pair| matches!(pair.label.trim(), "Sheet" | "Worksheets"))
        .ok_or_else(|| invalid("extended-properties heading pairs have no sheet catalog"))?;
    let new_catalog_count = catalog_pair
        .count
        .checked_add(1)
        .ok_or_else(|| invalid("extended-properties worksheet count overflow"))?;
    replacements.push(Replacement {
        span: catalog_pair.span,
        bytes: new_catalog_count.to_string().into_bytes(),
    });
    replacements.push(Replacement {
        span: Span {
            start: insert_at,
            end: insert_at,
        },
        bytes: new_title,
    });
    apply_replacements(content, replacements, "inserted extended-properties titles")
}

/// Remove catalog titles from an extended-properties part.
///
/// `original_titles` is the immutable source catalog and `positions` contains
/// its sorted, unique zero-based slots to remove. The `TitlesOfParts` vector
/// and the matching `Sheet` or `Worksheets` heading count are decremented;
/// trailing named-range entries and all surviving title elements retain their
/// original bytes. If the optional metadata is absent, the part is left
/// untouched.
///
/// # Example
///
/// ```
/// use litchi_ooxml_common::xml::sheet_name::remove_titles;
/// let xml = br#"<p:Properties xmlns:p="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:v="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><p:HeadingPairs><v:vector size="2" baseType="variant"><v:variant><v:lpstr>Worksheets</v:lpstr></v:variant><v:variant><v:i4>3</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="3" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr><v:lpstr>Three</v:lpstr></v:vector></p:TitlesOfParts></p:Properties>"#;
/// let changed = remove_titles(xml, &["One", "Two", "Three"], &[1], |left, right| left == right)?;
/// let changed = std::str::from_utf8(changed.as_deref().expect("metadata present"))?;
/// assert!(changed.contains("<v:i4>2</v:i4>"));
/// assert!(!changed.contains(">Two</v:lpstr>"));
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn remove_titles<F>(
    content: &[u8],
    original_titles: &[&str],
    positions: &[usize],
    equivalent: F,
) -> Result<Option<Vec<u8>>>
where
    F: Fn(&str, &str) -> bool,
{
    validate_remove_positions(original_titles.len(), positions)?;
    if positions.is_empty() {
        return Ok(None);
    }
    let Some(metadata) = scan_insert_metadata(content)? else {
        return Ok(None);
    };
    let InsertMetadata {
        titles,
        heading_pairs,
        title_vector,
        heading_vector,
        heading_variant_count,
    } = metadata;
    validate_insert_layout(
        &title_vector,
        &heading_vector,
        &titles,
        &heading_pairs,
        heading_variant_count,
        original_titles,
        &equivalent,
    )?;

    let size_span = title_vector.size.map(|(_, span)| span).ok_or_else(|| {
        invalid("extended-properties TitlesOfParts vector has no checked size attribute")
    })?;
    let new_size = titles
        .len()
        .checked_sub(positions.len())
        .ok_or_else(|| invalid("extended-properties title vector removal underflow"))?;
    let catalog_pair = heading_pairs
        .iter()
        .find(|pair| matches!(pair.label.trim(), "Sheet" | "Worksheets"))
        .ok_or_else(|| invalid("extended-properties heading pairs have no sheet catalog"))?;
    let new_catalog_count = catalog_pair
        .count
        .checked_sub(positions.len())
        .ok_or_else(|| invalid("extended-properties worksheet count removal underflow"))?;

    let replacement_count = positions
        .len()
        .checked_add(2)
        .ok_or_else(|| invalid("extended-properties title replacement count overflow"))?;
    let mut replacements = Vec::<Replacement>::new();
    replacements
        .try_reserve_exact(replacement_count)
        .map_err(|source| Error::Allocation {
            resource: "removed extended-properties title replacements",
            source,
        })?;
    replacements.push(Replacement {
        span: size_span,
        bytes: new_size.to_string().into_bytes(),
    });
    replacements.push(Replacement {
        span: catalog_pair.span,
        bytes: new_catalog_count.to_string().into_bytes(),
    });
    for &position in positions {
        let title = titles.get(position).ok_or_else(|| {
            invalid(format!(
                "extended-properties title removal position {position} exceeds the source catalog"
            ))
        })?;
        replacements.push(Replacement {
            span: Span {
                start: title.start,
                end: title.end,
            },
            bytes: Vec::new(),
        });
    }
    apply_replacements(content, replacements, "removed extended-properties titles")
}

fn validate_remove_positions(catalog_len: usize, positions: &[usize]) -> Result<()> {
    let mut previous = None;
    for &position in positions {
        if position >= catalog_len {
            return Err(invalid(format!(
                "extended-properties title removal position {position} exceeds catalog length {catalog_len}"
            )));
        }
        if previous.is_some_and(|value| position <= value) {
            return Err(invalid(
                "extended-properties title removal positions must be sorted and unique",
            ));
        }
        previous = Some(position);
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum InsertKind {
    Properties,
    Titles,
    TitleVector,
    Title,
    HeadingPairs,
    HeadingVector,
    HeadingVariant,
    HeadingLabel,
    HeadingCount,
    Other,
}

#[derive(Debug)]
struct InsertMetadata {
    titles: Vec<InsertTitle>,
    heading_pairs: Vec<InsertHeadingPair>,
    title_vector: InsertVector,
    heading_vector: InsertVector,
    heading_variant_count: usize,
}

#[derive(Debug)]
struct InsertFrame {
    kind: InsertKind,
    name: Box<[u8]>,
    start: usize,
    inner_start: usize,
    text: String,
    has_markup: bool,
    vector: Option<InsertVector>,
    title: Option<InsertTitleName>,
    heading: Option<InsertHeadingBuilder>,
}

#[derive(Debug)]
struct InsertVector {
    name: Box<[u8]>,
    namespace: Box<str>,
    size: Option<(usize, Span)>,
    base_type: Option<Box<str>>,
    close_start: usize,
}

#[derive(Debug)]
struct InsertTitle {
    start: usize,
    end: usize,
    text: String,
    has_markup: bool,
    name: Box<[u8]>,
    needs_binding: bool,
    namespace: Box<str>,
}

#[derive(Debug)]
struct InsertTitleName {
    name: Box<[u8]>,
    needs_binding: bool,
    namespace: Box<str>,
}

#[derive(Debug, Default)]
struct InsertHeadingBuilder {
    label: Option<String>,
    count: Option<(usize, Span)>,
}

#[derive(Debug)]
struct InsertHeadingPair {
    label: String,
    count: usize,
    span: Span,
}

#[derive(Debug)]
struct InsertRawAttribute {
    name: Vec<u8>,
    span: Span,
    value: String,
}

fn insert_kind(
    parent: Option<InsertKind>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    title_entries: &mut usize,
) -> InsertKind {
    let local = element.name().local_name();
    let local = local.as_ref();
    if parent.is_none() && is_extended_element(namespace, element, b"Properties") {
        InsertKind::Properties
    } else if parent == Some(InsertKind::Properties)
        && is_extended_namespace(namespace)
        && local == b"TitlesOfParts"
    {
        InsertKind::Titles
    } else if parent == Some(InsertKind::Titles)
        && is_property_types_namespace(namespace)
        && local == b"vector"
    {
        InsertKind::TitleVector
    } else if parent == Some(InsertKind::TitleVector)
        && is_property_types_namespace(namespace)
        && local == b"lpstr"
    {
        *title_entries = title_entries.saturating_add(1);
        InsertKind::Title
    } else if parent == Some(InsertKind::Properties)
        && is_extended_namespace(namespace)
        && local == b"HeadingPairs"
    {
        InsertKind::HeadingPairs
    } else if parent == Some(InsertKind::HeadingPairs)
        && is_property_types_namespace(namespace)
        && local == b"vector"
    {
        InsertKind::HeadingVector
    } else if parent == Some(InsertKind::HeadingVector)
        && is_property_types_namespace(namespace)
        && local == b"variant"
    {
        InsertKind::HeadingVariant
    } else if parent == Some(InsertKind::HeadingVariant)
        && is_property_types_namespace(namespace)
        && local == b"lpstr"
    {
        InsertKind::HeadingLabel
    } else if parent == Some(InsertKind::HeadingVariant)
        && is_property_types_namespace(namespace)
        && local == b"i4"
    {
        InsertKind::HeadingCount
    } else {
        InsertKind::Other
    }
}

fn reject_insert_element(parent: Option<InsertKind>, kind: InsertKind, local: &[u8]) -> Result<()> {
    if local == b"Properties" && kind != InsertKind::Properties {
        return Err(invalid(
            "extended-properties Properties element is outside the root context",
        ));
    }
    if local == b"TitlesOfParts" && kind != InsertKind::Titles {
        return Err(invalid(
            "extended-properties TitlesOfParts is outside its recognized context",
        ));
    }
    if local == b"HeadingPairs" && kind != InsertKind::HeadingPairs {
        return Err(invalid(
            "extended-properties HeadingPairs is outside its recognized context",
        ));
    }
    if local == b"vector"
        && matches!(parent, Some(InsertKind::Titles | InsertKind::HeadingPairs))
        && kind == InsertKind::Other
    {
        return Err(invalid(
            "extended-properties metadata container has a foreign vector child",
        ));
    }
    if local == b"lpstr"
        && matches!(
            parent,
            Some(InsertKind::TitleVector | InsertKind::HeadingVariant)
        )
        && kind == InsertKind::Other
    {
        return Err(invalid(
            "extended-properties metadata container has a foreign lpstr child",
        ));
    }
    if local == b"i4" && parent == Some(InsertKind::HeadingVariant) && kind == InsertKind::Other {
        return Err(invalid(
            "extended-properties heading pair has a foreign count child",
        ));
    }
    if matches!(
        parent,
        Some(
            InsertKind::Titles
                | InsertKind::TitleVector
                | InsertKind::HeadingPairs
                | InsertKind::HeadingVector
                | InsertKind::HeadingVariant
                | InsertKind::Title
                | InsertKind::HeadingLabel
                | InsertKind::HeadingCount
        )
    ) && kind == InsertKind::Other
    {
        return Err(invalid(
            "extended-properties metadata context has an unexpected child",
        ));
    }
    Ok(())
}

fn parse_insert_vector(
    content: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    decoder: Decoder,
) -> Result<InsertVector> {
    let namespace = match namespace {
        ResolveResult::Bound(Namespace(value)) => std::str::from_utf8(value)
            .map_err(|source| Error::Utf8 {
                context: "extended-properties metadata vector namespace",
                source,
            })?
            .to_owned()
            .into_boxed_str(),
        _ => {
            return Err(invalid(
                "extended-properties metadata vector has no bound namespace",
            ));
        },
    };
    let attributes = insert_raw_attributes(content, start, end, element, decoder)?;
    let mut size = None;
    let mut base_type = None;
    for attribute in attributes {
        if attribute.name.as_slice() == b"size" {
            if size.is_some() {
                return Err(invalid(
                    "extended-properties metadata vector has duplicate size attributes",
                ));
            }
            let raw = content
                .get(attribute.span.start..attribute.span.end)
                .ok_or_else(|| {
                    invalid("extended-properties size attribute span is out of bounds")
                })?;
            if raw.is_empty() || !raw.iter().all(u8::is_ascii_digit) {
                return Err(invalid(
                    "extended-properties metadata vector size is not a decimal integer",
                ));
            }
            let value = std::str::from_utf8(raw)
                .map_err(|source| Error::Utf8 {
                    context: "extended-properties metadata vector size",
                    source,
                })?
                .parse::<usize>()
                .map_err(|_| invalid("extended-properties metadata vector size overflows usize"))?;
            size = Some((value, attribute.span));
        } else if attribute.name.as_slice() == b"baseType" {
            if base_type.is_some() {
                return Err(invalid(
                    "extended-properties metadata vector has duplicate baseType attributes",
                ));
            }
            base_type = Some(attribute.value.into_boxed_str());
        }
    }
    Ok(InsertVector {
        name: element.name().as_ref().to_vec().into_boxed_slice(),
        namespace,
        size,
        base_type,
        close_start: end,
    })
}

fn insert_title_name(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
) -> Result<InsertTitleName> {
    let namespace = match namespace {
        ResolveResult::Bound(Namespace(value)) if is_property_types_uri(value) => {
            std::str::from_utf8(value)
                .map_err(|source| Error::Utf8 {
                    context: "extended-properties catalog title namespace",
                    source,
                })?
                .to_owned()
                .into_boxed_str()
        },
        _ => {
            return Err(invalid(
                "extended-properties catalog title has no property-types namespace",
            ));
        },
    };
    let prefix = element.name().prefix();
    let prefix_name = prefix.map(|value| value.into_inner());
    let needs_binding = has_local_namespace_declaration(element, prefix_name)?;
    if matches!(
        resolver.resolve_prefix(prefix, true),
        ResolveResult::Unknown(_)
    ) {
        return Err(invalid(
            "extended-properties catalog title has an unbound namespace prefix",
        ));
    }
    Ok(InsertTitleName {
        name: element.name().as_ref().to_vec().into_boxed_slice(),
        needs_binding,
        namespace,
    })
}

fn is_property_types_uri(value: &[u8]) -> bool {
    value == PROPERTY_TYPES || value == STRICT_PROPERTY_TYPES
}

fn has_local_namespace_declaration(
    element: &BytesStart<'_>,
    prefix: Option<&[u8]>,
) -> Result<bool> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|source| Error::Attribute {
            context: "extended-properties title namespace declaration",
            source,
        })?;
        let key = attribute.key.as_ref();
        if prefix.is_none() {
            if key == b"xmlns" {
                return Ok(true);
            }
        } else if key
            .strip_prefix(b"xmlns:")
            .is_some_and(|value| Some(value) == prefix)
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn insert_raw_attributes(
    content: &[u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Vec<InsertRawAttribute>> {
    let raw = raw_attribute_spans(content, start, end)?;
    let mut decoded = Vec::new();
    decoded
        .try_reserve_exact(raw.len())
        .map_err(|source| Error::Allocation {
            resource: "extended-properties metadata attributes",
            source,
        })?;
    for (index, attribute) in element.checked_attributes().enumerate() {
        let attribute = attribute.map_err(|source| Error::Attribute {
            context: "extended-properties metadata attribute",
            source,
        })?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|source| Error::Parse {
                context: "extended-properties metadata attribute value",
                source,
            })?;
        let raw_attribute = raw.get(index).ok_or_else(|| {
            invalid("extended-properties metadata attribute count is inconsistent")
        })?;
        decoded.push(InsertRawAttribute {
            name: attribute.key.as_ref().to_vec(),
            span: raw_attribute.span,
            value: value.into_owned(),
        });
    }
    if decoded.len() != raw.len() {
        return Err(invalid(
            "extended-properties metadata attribute count is inconsistent",
        ));
    }
    Ok(decoded)
}

fn raw_attribute_spans(
    content: &[u8],
    start: usize,
    end: usize,
) -> Result<Vec<InsertRawAttribute>> {
    let bytes = content
        .get(start..end)
        .ok_or_else(|| invalid("extended-properties tag span is out of bounds"))?;
    if bytes.first() != Some(&b'<') {
        return Err(invalid("extended-properties tag does not start with '<'"));
    }
    let mut cursor = 1usize;
    while cursor < bytes.len()
        && !is_xml_space_byte(bytes[cursor])
        && !matches!(bytes[cursor], b'>')
    {
        cursor += 1;
    }
    let mut output = Vec::new();
    while cursor < bytes.len() {
        while cursor < bytes.len() && is_xml_space_byte(bytes[cursor]) {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] == b'>' {
            break;
        }
        if bytes[cursor] == b'/' {
            if bytes.get(cursor + 1) == Some(&b'>') {
                break;
            }
            return Err(invalid("extended-properties tag has an invalid '/'"));
        }
        let name_start = cursor;
        while cursor < bytes.len()
            && !is_xml_space_byte(bytes[cursor])
            && !matches!(bytes[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        if name_start == cursor {
            return Err(invalid(
                "extended-properties tag has an empty attribute name",
            ));
        }
        let name = bytes[name_start..cursor].to_vec();
        while cursor < bytes.len() && is_xml_space_byte(bytes[cursor]) {
            cursor += 1;
        }
        if bytes.get(cursor) != Some(&b'=') {
            return Err(invalid("extended-properties attribute is missing '='"));
        }
        cursor += 1;
        while cursor < bytes.len() && is_xml_space_byte(bytes[cursor]) {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .ok_or_else(|| invalid("extended-properties attribute is missing its quote"))?;
        if !matches!(quote, b'"' | b'\'') {
            return Err(invalid("extended-properties attribute is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= bytes.len() {
            return Err(invalid("extended-properties attribute quote is not closed"));
        }
        output.push(InsertRawAttribute {
            name,
            span: Span {
                start: start + value_start,
                end: start + value_end,
            },
            value: String::new(),
        });
        cursor += 1;
    }
    Ok(output)
}

fn qualified_lpstr_name(vector_name: &[u8]) -> Box<[u8]> {
    split_prefix(vector_name).map_or_else(
        || b"lpstr".to_vec().into_boxed_slice(),
        |(prefix, _)| {
            let mut name = Vec::with_capacity(prefix.len() + 1 + 5);
            name.extend_from_slice(prefix);
            name.push(b':');
            name.extend_from_slice(b"lpstr");
            name.into_boxed_slice()
        },
    )
}

fn split_prefix(name: &[u8]) -> Option<(&[u8], &[u8])> {
    let separator = name.iter().position(|byte| *byte == b':')?;
    Some((&name[..separator], &name[separator + 1..]))
}

fn write_insert_title(
    title_name: &[u8],
    namespace: &str,
    needs_binding: bool,
    value: &str,
) -> Result<Vec<u8>> {
    if value.len() > MAX_CAPTURED_TEXT_BYTES {
        return Err(Error::Limit(Limit::CapturedTextBytes {
            limit: MAX_CAPTURED_TEXT_BYTES,
        }));
    }
    if value.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{a}' | '\u{d}'
            | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}'
            | '\u{10000}'..='\u{10ffff}')
    }) {
        return Err(invalid(
            "inserted title contains a character forbidden by XML 1.0",
        ));
    }
    // Literal CR is normalized to LF by XML readers; a reference preserves
    // the semantic sheet name for subsequent metadata edits.
    let escaped = escape_xml(value);
    let escaped = if escaped.contains('\r') {
        escaped.replace('\r', "&#13;")
    } else {
        escaped
    };
    let prefix = split_prefix(title_name).map(|(prefix, _)| prefix);
    let binding_len = if needs_binding {
        9usize.saturating_add(prefix.map_or(0, |value| value.len()))
    } else {
        0
    };
    let mut output = Vec::new();
    let capacity = title_name
        .len()
        .checked_mul(2)
        .and_then(|size| size.checked_add(escaped.len()))
        .and_then(|size| size.checked_add(binding_len))
        .and_then(|size| size.checked_add(namespace.len()))
        .and_then(|size| size.checked_add(8))
        .ok_or_else(|| invalid("extended-properties inserted title size overflow"))?;
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "extended-properties inserted title",
            source,
        })?;
    output.extend_from_slice(b"<");
    output.extend_from_slice(title_name);
    if needs_binding {
        if let Some(prefix) = prefix {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix);
        } else {
            output.extend_from_slice(b" xmlns");
        }
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(namespace).as_bytes());
        output.extend_from_slice(b"\"");
    }
    output.extend_from_slice(b">");
    output.extend_from_slice(escaped.as_bytes());
    output.extend_from_slice(b"</");
    output.extend_from_slice(title_name);
    output.extend_from_slice(b">");
    Ok(output)
}

fn push_insert_text(text: &mut String, value: &str, captured: &mut usize) -> Result<()> {
    let next = captured
        .checked_add(value.len())
        .ok_or(Error::Limit(Limit::CapturedTextBytes {
            limit: MAX_CAPTURED_TEXT_BYTES,
        }))?;
    if next > MAX_CAPTURED_TEXT_BYTES {
        return Err(Error::Limit(Limit::CapturedTextBytes {
            limit: MAX_CAPTURED_TEXT_BYTES,
        }));
    }
    text.try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "extended-properties insertion text",
            source,
        })?;
    text.push_str(value);
    *captured = next;
    Ok(())
}

fn decimal_text_span(content: &[u8], start: usize, end: usize) -> Result<Span> {
    let raw = content
        .get(start..end)
        .ok_or_else(|| invalid("extended-properties decimal text span is out of bounds"))?;
    let mut left = 0usize;
    let mut right = raw.len();
    while left < right && is_xml_space_byte(raw[left]) {
        left += 1;
    }
    while right > left && is_xml_space_byte(raw[right - 1]) {
        right -= 1;
    }
    if left == right || !raw[left..right].iter().all(u8::is_ascii_digit) {
        return Err(invalid(
            "extended-properties heading count is not a plain decimal text value",
        ));
    }
    Ok(Span {
        start: start + left,
        end: start + right,
    })
}

fn is_xml_space_byte(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

fn validate_insert_layout<F>(
    title_vector: &InsertVector,
    heading_vector: &InsertVector,
    titles: &[InsertTitle],
    heading_pairs: &[InsertHeadingPair],
    heading_variant_count: usize,
    original_titles: &[&str],
    equivalent: &F,
) -> Result<()>
where
    F: Fn(&str, &str) -> bool,
{
    // Category names are extensible. A valid category outside the modeled
    // catalog/named-range layout is unsupported, not evidence of corruption.
    if heading_pairs.iter().any(|pair| {
        !matches!(
            pair.label.trim(),
            "Sheet" | "Worksheets" | "Named Range" | "Named Ranges"
        )
    }) {
        return Err(Error::Blocked(BlockReason::UnmodeledReference));
    }
    if title_vector
        .base_type
        .as_deref()
        .is_some_and(|value| value != "lpstr")
    {
        return Err(invalid(
            "extended-properties TitlesOfParts vector has a non-lpstr baseType",
        ));
    }
    if heading_vector.base_type.as_deref() != Some("variant") {
        return Err(invalid(
            "extended-properties HeadingPairs vector is missing its variant baseType",
        ));
    }
    if title_vector
        .size
        .is_none_or(|(size, _)| size != titles.len())
    {
        return Err(invalid(
            "extended-properties TitlesOfParts vector size does not match its entries",
        ));
    }
    if heading_vector
        .size
        .is_none_or(|(size, _)| size != heading_variant_count)
    {
        return Err(invalid(
            "extended-properties HeadingPairs vector size does not match its entries",
        ));
    }
    if titles.len() < original_titles.len() {
        return Err(invalid(format!(
            "extended-properties TitlesOfParts has {} titles but the catalog has {}",
            titles.len(),
            original_titles.len()
        )));
    }
    for (index, (actual, expected)) in titles.iter().zip(original_titles).enumerate() {
        if actual.has_markup || !equivalent(&actual.text, expected) {
            return Err(invalid(format!(
                "extended-properties catalog title {index} does not match the source catalog"
            )));
        }
    }
    if titles
        .iter()
        .skip(original_titles.len())
        .any(|title| title.has_markup)
    {
        return Err(invalid(
            "extended-properties trailing named-range title contains nested markup",
        ));
    }
    let trailing = titles.len() - original_titles.len();
    let mut catalog = None::<&InsertHeadingPair>;
    let mut named_ranges = None::<&InsertHeadingPair>;
    let mut group_stage = 0u8;
    for pair in heading_pairs {
        match pair.label.trim() {
            "Sheet" | "Worksheets" => {
                if group_stage != 0 || catalog.replace(pair).is_some() {
                    return Err(invalid(
                        "extended-properties HeadingPairs sheet category is out of order or duplicated",
                    ));
                }
                group_stage = 1;
                if pair.count != original_titles.len() {
                    return Err(invalid(
                        "extended-properties sheet category count does not match the catalog",
                    ));
                }
            },
            "Named Range" | "Named Ranges" => {
                if group_stage != 1 || named_ranges.replace(pair).is_some() {
                    return Err(invalid(
                        "extended-properties HeadingPairs named-range category is out of order or duplicated",
                    ));
                }
                group_stage = 2;
                if pair.count != trailing {
                    return Err(invalid(
                        "extended-properties named-range category count does not match trailing titles",
                    ));
                }
            },
            _ => {
                return Err(invalid(format!(
                    "extended-properties HeadingPairs category {:?} is not a supported sheet group",
                    pair.label
                )));
            },
        }
    }
    let catalog = catalog.ok_or_else(|| {
        invalid("extended-properties HeadingPairs has no Sheet or Worksheets category")
    })?;
    let named_range_count = named_ranges.map_or(0, |pair| pair.count);
    if trailing != named_range_count {
        return Err(invalid(
            "extended-properties trailing titles have no matching named-range category",
        ));
    }
    let total_group_count = catalog
        .count
        .checked_add(named_range_count)
        .ok_or_else(|| invalid("extended-properties HeadingPairs group count overflows usize"))?;
    if total_group_count != titles.len() {
        return Err(invalid(
            "extended-properties HeadingPairs group counts do not match TitlesOfParts",
        ));
    }
    Ok(())
}

fn validate_title_order(catalog_len: usize, order: &[usize]) -> Result<bool> {
    if order.len() != catalog_len {
        return Err(invalid(format!(
            "sheet-title order has {} entries but the catalog has {catalog_len}",
            order.len()
        )));
    }
    let mut seen = Vec::new();
    seen.try_reserve_exact(catalog_len)
        .map_err(|source| Error::Allocation {
            resource: "sheet-title permutation validation",
            source,
        })?;
    seen.resize(catalog_len, false);
    let mut identity = true;
    for (position, &source) in order.iter().enumerate() {
        identity &= position == source;
        let slot = seen
            .get_mut(source)
            .ok_or_else(|| invalid("sheet-title order contains an out-of-range index"))?;
        if std::mem::replace(slot, true) {
            return Err(invalid("sheet-title order contains a duplicate index"));
        }
    }
    if seen.iter().any(|used| !used) {
        return Err(invalid("sheet-title order omits a catalog index"));
    }
    Ok(identity)
}

fn reorder_kind(
    parent: Option<TitleReorderKind>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    title_entries: &mut usize,
) -> TitleReorderKind {
    let local = element.name().local_name();
    let local = local.as_ref();
    if parent.is_none() && is_extended_element(namespace, element, b"Properties") {
        TitleReorderKind::Properties
    } else if parent == Some(TitleReorderKind::Properties)
        && is_extended_namespace(namespace)
        && local == b"TitlesOfParts"
    {
        TitleReorderKind::TitlesOfParts
    } else if parent == Some(TitleReorderKind::TitlesOfParts)
        && is_property_types_namespace(namespace)
        && local == b"vector"
    {
        TitleReorderKind::TitleVector
    } else if parent == Some(TitleReorderKind::TitleVector)
        && is_property_types_namespace(namespace)
        && local == b"lpstr"
    {
        let index = *title_entries;
        *title_entries = (*title_entries).saturating_add(1);
        TitleReorderKind::Title(index)
    } else {
        TitleReorderKind::Other
    }
}

fn reject_ambiguous_title_element(
    local: &[u8],
    parent: Option<TitleReorderKind>,
    kind: TitleReorderKind,
) -> Result<()> {
    if local == b"TitlesOfParts" && kind != TitleReorderKind::TitlesOfParts {
        return Err(invalid(
            "extended-properties TitlesOfParts is outside its recognized context",
        ));
    }
    if local == b"vector"
        && parent == Some(TitleReorderKind::TitlesOfParts)
        && kind != TitleReorderKind::TitleVector
    {
        return Err(invalid(
            "extended-properties TitlesOfParts has a foreign vector child",
        ));
    }
    if local == b"lpstr"
        && parent == Some(TitleReorderKind::TitleVector)
        && kind == TitleReorderKind::Other
    {
        return Err(invalid(
            "extended-properties title vector has a foreign lpstr child",
        ));
    }
    if matches!(
        parent,
        Some(TitleReorderKind::TitlesOfParts | TitleReorderKind::TitleVector)
    ) && kind == TitleReorderKind::Other
    {
        return Err(invalid(
            "extended-properties title context has an unexpected child",
        ));
    }
    Ok(())
}

fn validate_title_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let mut expanded = BTreeSet::<(Vec<u8>, Vec<u8>)>::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|source| Error::Attribute {
            context: "sheet-title XML attribute",
            source,
        })?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|source| Error::Parse {
                context: "sheet-title XML attribute value",
                source,
            })?;
        let key = attribute.key;
        if is_namespace_declaration(key.as_ref()) {
            if value.as_bytes() == MCE {
                return Err(Error::Blocked(BlockReason::MarkupCompatibility));
            }
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => value.to_vec(),
            ResolveResult::Unbound => Vec::new(),
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "sheet-title XML attribute has an unbound namespace prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        if namespace.as_slice() == MCE {
            return Err(Error::Blocked(BlockReason::MarkupCompatibility));
        }
        let expanded_name = (namespace, local.as_ref().to_vec());
        if !expanded.insert(expanded_name) {
            return Err(invalid(format!(
                "duplicate expanded sheet-title XML attribute '{}'",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
    }
    Ok(())
}

fn is_namespace_declaration(key: &[u8]) -> bool {
    key == b"xmlns" || key.starts_with(b"xmlns:")
}

fn is_xml_whitespace(text: &str) -> bool {
    text.bytes()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments carry the complete source and span context for one title edit"
)]
fn finish_reordered_title<F>(
    index: usize,
    actual: &str,
    has_markup: bool,
    span: Span,
    content: &[u8],
    original_titles: &[&str],
    order: &[usize],
    equivalent: &F,
) -> Result<Option<Replacement>>
where
    F: Fn(&str, &str) -> bool,
{
    let Some(expected) = original_titles.get(index) else {
        return Ok(None);
    };
    if has_markup {
        return Err(invalid(format!(
            "extended-properties catalog title {index} contains nested markup"
        )));
    }
    if !equivalent(actual, expected) {
        return Err(invalid(format!(
            "extended-properties catalog title {index} does not match the source catalog"
        )));
    }
    if order[index] == index {
        return Ok(None);
    }
    let target = original_titles[order[index]];
    let escaped = escape_xml(target);
    if span.start == span.end && !escaped.is_empty() {
        return Err(invalid(format!(
            "extended-properties catalog title {index} is empty but its reordered value is not"
        )));
    }
    if content
        .get(span.start..span.end)
        .is_some_and(|bytes| bytes == escaped.as_bytes())
    {
        return Ok(None);
    }
    Ok(Some(Replacement {
        span,
        bytes: escaped.into_bytes(),
    }))
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments describe all axes and bounds of one reference edit"
)]
fn attribute_replacement(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    element_local: &str,
    spreadsheet: bool,
    empty: bool,
    alternate: bool,
    span: Span,
    renames: &[Rename<'_>],
    equivalent: &impl Fn(&str, &str) -> bool,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    let direct_carrier = spreadsheet && matches!(element_local, "worksheetSource" | "dataRef");
    let hyperlink = spreadsheet && element_local == "hyperlink";
    let has_relationship = (direct_carrier || hyperlink)
        && relationship_attribute_value(element, b"id", decoder, resolver)?.is_some();
    let mut changed = Vec::<(Box<str>, String)>::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|source| Error::Attribute {
            context: "sheet-reference attribute",
            source,
        })?;
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(|source| Error::Utf8 {
            context: "sheet-reference attribute",
            source,
        })?;
        let attribute_local_name = attribute.key.local_name();
        let local =
            std::str::from_utf8(attribute_local_name.as_ref()).map_err(|source| Error::Utf8 {
                context: "sheet-reference attribute local name",
                source,
            })?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|source| Error::Parse {
                context: "sheet-reference attribute value",
                source,
            })?;
        let direct_name_attribute = direct_carrier && local == "sheet";
        let known_direct = direct_name_attribute && !has_relationship;
        let known_formula = hyperlink && local == "location" && !has_relationship;
        if direct_name_attribute && has_relationship {
            continue;
        }
        if known_direct {
            if let Some(after) = direct_name(&value, renames, equivalent) {
                ensure_direct(alternate)?;
                changed.push((name.into(), after.to_owned()));
            }
            continue;
        }
        let suspicious_formula = formula_like(element_local) || formula_like(local);
        if known_formula || suspicious_formula {
            let result = rename_formula_sheets(&value, renames, equivalent);
            if known_formula {
                if let Some(text) = result.text {
                    ensure_direct(alternate)?;
                    changed.push((name.into(), text));
                }
            } else if result.matched {
                return Err(unsupported(alternate));
            }
        }
        if !known_direct
            && sheet_name_like(local)
            && direct_name(&value, renames, equivalent).is_some()
        {
            if spreadsheet && element_local == "sheetProtection" && local == "sheet" {
                continue;
            }
            return Err(unsupported(alternate));
        }
    }
    if changed.is_empty() {
        return Ok(());
    }
    replacements.push(Replacement {
        span,
        bytes: write_tag(&tag(element, decoder)?, empty, &changed),
    });
    Ok(())
}

fn finish_frame(
    frame: Frame,
    inner_end: usize,
    renames: &[Rename<'_>],
    sheet_titles: &[&str],
    equivalent: &impl Fn(&str, &str) -> bool,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    if frame.mode == Mode::Other {
        return Ok(());
    }
    let (text, matched) = match frame.mode {
        Mode::Formula => {
            let result = rename_formula_sheets(&frame.text, renames, equivalent);
            (result.text, result.matched)
        },
        Mode::Suspicious => {
            let result = rename_formula_sheets(&frame.text, renames, equivalent);
            if result.matched {
                return Err(unsupported(frame.alternate));
            }
            (None, false)
        },
        Mode::SheetTitle(index) => {
            let expected = sheet_titles.get(index).ok_or_else(|| {
                invalid("extended-property sheet-title index escaped its checked catalog")
            })?;
            if !equivalent(&frame.text, expected) {
                return Err(unsupported(frame.alternate));
            }
            let direct = direct_name(&frame.text, renames, equivalent).map(str::to_owned);
            let matched = direct.is_some();
            (direct, matched)
        },
        Mode::NamedRangeTitle => {
            let result = rename_formula_sheets(&frame.text, renames, equivalent);
            (result.text, result.matched)
        },
        Mode::Other => (None, false),
    };
    if frame.has_markup {
        return if matched {
            Err(unsupported(frame.alternate))
        } else {
            Ok(())
        };
    }
    let Some(text) = text else {
        return Ok(());
    };
    ensure_direct(frame.alternate)?;
    replacements.push(Replacement {
        span: Span {
            start: frame.inner_start,
            end: inner_end,
        },
        bytes: escape_xml(&text).into_bytes(),
    });
    Ok(())
}

fn direct_name<'a>(
    value: &str,
    renames: &'a [Rename<'a>],
    equivalent: &impl Fn(&str, &str) -> bool,
) -> Option<&'a str> {
    renames
        .iter()
        .find(|rename| equivalent(value, rename.before))
        .map(|rename| rename.after)
}

/// Result of scanning one formula-like value. The allocation is deferred
/// until a local prefix actually changes; `matched` still records a
/// case-preserving or identity-equivalent match for conservative carriers.
#[derive(Debug)]
struct FormulaRenameResult {
    text: Option<String>,
    matched: bool,
}

/// Rewrite local single-sheet and 3-D prefixes in one pass.
///
/// String constants, structured-reference brackets, and nonzero/external
/// workbook prefixes remain exact. All mappings observe the source formula,
/// so swaps such as `One <-> Two` do not cascade.
fn rename_formula_sheets(
    formula: &str,
    renames: &[Rename<'_>],
    equivalent: &impl Fn(&str, &str) -> bool,
) -> FormulaRenameResult {
    if formula.is_empty() || renames.is_empty() {
        return FormulaRenameResult {
            text: None,
            matched: false,
        };
    }
    let bytes = formula.as_bytes();
    let mut output = None::<String>;
    let mut copied = 0usize;
    let mut index = 0usize;
    let mut matched = false;

    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = quoted_string_end(bytes, index);
            continue;
        }

        let candidate = if bytes[index] == b'\'' {
            quoted_sheet_prefix_end(bytes, index).map(|end| (end, true))
        } else if is_prefix_start(bytes, index) {
            unquoted_sheet_prefix_end(bytes, index).map(|end| (end, false))
        } else {
            None
        };
        if let Some((end, quoted)) = candidate {
            let prefix_end = end - 1;
            let raw = &formula[index..prefix_end];
            if let Some(rewritten) = rewrite_sheet_prefix(raw, quoted, renames, equivalent) {
                matched = true;
                if rewritten != raw {
                    let target = output.get_or_insert_with(|| {
                        let mut value = String::with_capacity(formula.len());
                        value.push_str(&formula[..index]);
                        value
                    });
                    if copied != 0 {
                        target.push_str(&formula[copied..index]);
                    }
                    target.push_str(&rewritten);
                    target.push('!');
                    copied = end;
                }
            }
            index = end;
            continue;
        }

        let Some(character) = formula[index..].chars().next() else {
            break;
        };
        index += character.len_utf8();
    }

    if let Some(output) = output.as_mut() {
        output.push_str(&formula[copied..]);
    }
    FormulaRenameResult {
        text: output,
        matched,
    }
}

fn quoted_sheet_prefix_end(bytes: &[u8], start: usize) -> Option<usize> {
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == b'\'' {
            if bytes.get(index + 1) == Some(&b'\'') {
                index += 2;
                continue;
            }
            return (bytes.get(index + 1) == Some(&b'!')).then_some(index + 2);
        }
        index += 1;
    }
    None
}

fn unquoted_sheet_prefix_end(bytes: &[u8], start: usize) -> Option<usize> {
    let mut index = start;
    while bytes.get(index).is_some_and(|byte| is_prefix_byte(*byte)) {
        index += 1;
    }
    (index != start && bytes.get(index) == Some(&b'!')).then_some(index + 1)
}

fn is_prefix_start(bytes: &[u8], start: usize) -> bool {
    is_prefix_byte(bytes[start]) && (start == 0 || !is_prefix_byte(bytes[start - 1]))
}

fn is_prefix_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.' | b':' | b'$' | b'[' | b']')
}

fn rewrite_sheet_prefix(
    raw: &str,
    quoted: bool,
    renames: &[Rename<'_>],
    equivalent: &impl Fn(&str, &str) -> bool,
) -> Option<String> {
    let decoded;
    let prefix = if quoted {
        let inner = raw.strip_prefix('\'')?.strip_suffix('\'')?;
        decoded = if inner.contains("''") {
            std::borrow::Cow::Owned(inner.replace("''", "'"))
        } else {
            std::borrow::Cow::Borrowed(inner)
        };
        decoded.as_ref()
    } else {
        raw
    };
    let (book, sheets) = local_workbook_prefix(prefix)?;
    let (first, last) = sheets
        .split_once(':')
        .map_or((sheets, None), |(first, last)| (first, Some(last)));
    if first.is_empty()
        || last.is_some_and(str::is_empty)
        || last.is_some_and(|last| last.contains(':'))
    {
        return None;
    }

    let mut matched = false;
    let first = mapped_name(first, renames, &mut matched, equivalent);
    let last = last.map(|last| mapped_name(last, renames, &mut matched, equivalent));
    if !matched {
        return None;
    }

    let quote = quoted
        || !formula_name_is_unquoted(first)
        || last.is_some_and(|last| !formula_name_is_unquoted(last));
    let mut output = String::with_capacity(raw.len());
    if quote {
        output.push('\'');
    }
    output.push_str(book);
    push_formula_name(&mut output, first, quote);
    if let Some(last) = last {
        output.push(':');
        push_formula_name(&mut output, last, quote);
    }
    if quote {
        output.push('\'');
    }
    Some(output)
}

fn local_workbook_prefix(prefix: &str) -> Option<(&str, &str)> {
    if let Some(rest) = prefix.strip_prefix('[') {
        let close = rest.find(']')?;
        let workbook = &rest[..close];
        if workbook.parse::<u32>().ok()? != 0 {
            return None;
        }
        let book_end = close + 2;
        return Some((&prefix[..book_end], &prefix[book_end..]));
    }
    // Local sheet names cannot contain these characters. Their presence here
    // denotes an external path/book prefix rather than a local sheet.
    if prefix.contains(['[', ']', '\\', '/']) {
        return None;
    }
    Some(("", prefix))
}

fn mapped_name<'a>(
    name: &'a str,
    renames: &'a [Rename<'a>],
    matched: &mut bool,
    equivalent: &impl Fn(&str, &str) -> bool,
) -> &'a str {
    for rename in renames {
        if equivalent(name, rename.before) {
            *matched = true;
            return rename.after;
        }
    }
    name
}

fn formula_name_is_unquoted(name: &str) -> bool {
    !name.is_empty()
        && name
            .chars()
            .all(|character| character.is_alphanumeric() || matches!(character, '_' | '.'))
}

fn push_formula_name(output: &mut String, name: &str, quoted: bool) {
    if quoted {
        for character in name.chars() {
            output.push(character);
            if character == '\'' {
                output.push('\'');
            }
        }
    } else {
        output.push_str(name);
    }
}

fn quoted_string_end(bytes: &[u8], start: usize) -> usize {
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            if bytes.get(index + 1) == Some(&b'"') {
                index += 2;
            } else {
                return index + 1;
            }
        } else {
            index += 1;
        }
    }
    bytes.len()
}

fn mode(
    local: &str,
    title_index: Option<usize>,
    namespace: &ResolveResult<'_>,
    sheet_titles: usize,
) -> Mode {
    if let Some(index) = title_index {
        if index < sheet_titles {
            Mode::SheetTitle(index)
        } else {
            Mode::NamedRangeTitle
        }
    } else if formula_element(local)
        && is_namespace(
            namespace,
            &[SML, STRICT_SML, CHART, STRICT_CHART, EXCEL_FORMULA],
        )
    {
        Mode::Formula
    } else if matches!(local, "formula1" | "formula2") && is_namespace(namespace, &[X14]) {
        // x14 uses these as structural wrappers around either an xm:f formula
        // or an x12ac:list literal. The payload's expanded name decides its
        // semantics, so concatenating wrapper children would be incorrect.
        Mode::Other
    } else if formula_like(local) {
        Mode::Suspicious
    } else {
        Mode::Other
    }
}

fn formula_like(local: &str) -> bool {
    local.eq_ignore_ascii_case("f")
        || contains_ascii_case_insensitive(local, "formula")
        || contains_ascii_case_insensitive(local, "fmla")
        || local.eq_ignore_ascii_case("refersto")
}

fn formula_element(local: &str) -> bool {
    matches!(
        local,
        "f" | "formula"
            | "formula1"
            | "formula2"
            | "definedName"
            | "calculatedColumnFormula"
            | "totalsRowFormula"
    )
}

fn sheet_name_like(local: &str) -> bool {
    local.eq_ignore_ascii_case("sheet")
        || local.eq_ignore_ascii_case("sheetName")
        || local.eq_ignore_ascii_case("worksheet")
}

fn contains_ascii_case_insensitive(value: &str, needle: &str) -> bool {
    value
        .as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle.as_bytes()))
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[&[u8]]) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if expected.contains(value)
    )
}

fn is_extended_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace(
        namespace,
        &[EXTENDED_PROPERTIES, STRICT_EXTENDED_PROPERTIES],
    )
}

fn is_property_types_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace(namespace, &[PROPERTY_TYPES, STRICT_PROPERTY_TYPES])
}

fn is_mce_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace(namespace, &[MCE])
}

fn is_extended_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> bool {
    element.name().local_name().as_ref() == local && is_extended_namespace(namespace)
}

fn ensure_direct(alternate: bool) -> Result<()> {
    if alternate {
        Err(Error::Blocked(BlockReason::MarkupCompatibility))
    } else {
        Ok(())
    }
}

fn unsupported(alternate: bool) -> Error {
    Error::Blocked(if alternate {
        BlockReason::MarkupCompatibility
    } else {
        BlockReason::UnmodeledReference
    })
}

fn push_text(stack: &mut [Frame], text: &str, captured_bytes: &mut usize) -> Result<()> {
    let receivers = stack
        .iter()
        .filter(|frame| frame.mode != Mode::Other)
        .count();
    let added =
        text.len()
            .checked_mul(receivers)
            .ok_or(Error::Limit(Limit::CapturedTextBytes {
                limit: MAX_CAPTURED_TEXT_BYTES,
            }))?;
    let next = captured_bytes
        .checked_add(added)
        .ok_or(Error::Limit(Limit::CapturedTextBytes {
            limit: MAX_CAPTURED_TEXT_BYTES,
        }))?;
    if next > MAX_CAPTURED_TEXT_BYTES {
        return Err(Error::Limit(Limit::CapturedTextBytes {
            limit: MAX_CAPTURED_TEXT_BYTES,
        }));
    }
    *captured_bytes = next;
    for frame in stack.iter_mut().filter(|frame| frame.mode != Mode::Other) {
        frame.text.push_str(text);
    }
    Ok(())
}

fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|source| Error::Utf8 {
            context: "sheet-reference element name",
            source,
        })?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|source| Error::Attribute {
            context: "sheet-reference attribute",
            source,
        })?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|source| Error::Utf8 {
                context: "sheet-reference attribute",
                source,
            })?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|source| Error::Parse {
                context: "sheet-reference attribute value",
                source,
            })?
            .into_owned();
        attributes.push(Attribute {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        });
    }
    Ok(Tag {
        name: name.into_boxed_str(),
        attributes: attributes.into_boxed_slice(),
    })
}

fn write_tag(tag: &Tag, empty: bool, changed: &[(Box<str>, String)]) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        output.extend_from_slice(b" ");
        output.extend_from_slice(attribute.name.as_bytes());
        output.extend_from_slice(b"=\"");
        let value = changed
            .iter()
            .find(|(name, _)| name.as_ref() == attribute.name.as_ref())
            .map_or(attribute.value.as_ref(), |(_, value)| value.as_str());
        output.extend_from_slice(escape_xml(value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    output
}

fn is_mce_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("sheet-reference XML position does not fit usize"))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use caseless::Caseless;
    use unicode_normalization::UnicodeNormalization;

    const SML_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const X14_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    const EXCEL_FORMULA_NAMESPACE: &str = "http://schemas.microsoft.com/office/excel/2006/main";
    const EXTENDED_PROPERTIES_NAMESPACE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    const PROPERTY_TYPES_NAMESPACE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    const MCE_STR: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn equivalent(left: &str, right: &str) -> bool {
        left.chars().nfd().default_case_fold().nfd().eq(right
            .chars()
            .nfd()
            .default_case_fold()
            .nfd())
    }

    fn rename<'a>(before: &'a str, after: &'a str) -> [Rename<'a>; 1] {
        [Rename { before, after }]
    }

    #[test]
    fn rewrites_formula_hyperlink_direct_and_title_dependencies() {
        let source = format!(
            r#"<root xmlns:s="{SML_NAMESPACE}" xmlns:ep="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><s:f>Data!A1</s:f><s:formula1>'Data'!Named</s:formula1><s:hyperlink location="Data!$A$1" keep="x"/><s:worksheetSource sheet="data" ref="A1"/><s:dataRef sheet="Data" ref="A1"/><ep:TitlesOfParts><vt:vector><vt:lpstr>Data</vt:lpstr><vt:lpstr>Data</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></ep:TitlesOfParts></root>"#
        );
        let changed = rewrite(
            source.as_bytes(),
            &rename("Data", "Input 2026"),
            &["Data"],
            equivalent,
        )
        .expect("rewrite")
        .expect("changed output");
        let changed = std::str::from_utf8(&changed).expect("UTF-8");
        assert!(changed.contains("<s:f>&apos;Input 2026&apos;!A1</s:f>"));
        assert!(changed.contains("<s:formula1>&apos;Input 2026&apos;!Named</s:formula1>"));
        assert!(changed.contains("location=\"&apos;Input 2026&apos;!$A$1\""));
        assert_eq!(changed.matches("sheet=\"Input 2026\"").count(), 2);
        assert!(changed.contains("<vt:lpstr>Input 2026</vt:lpstr>"));
        assert!(changed.contains("<vt:lpstr>Data</vt:lpstr>"));
        assert!(changed.contains("<vt:lpstr>&apos;Input 2026&apos;!Print_Area</vt:lpstr>"));
    }

    #[test]
    fn preserves_external_and_inert_formula_values_exactly() {
        let source = format!(
            r#"<root xmlns:s="{SML_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><s:f>[1]Data!A1</s:f><value>Data!A1</value><s:f>Other!A1</s:f><s:worksheetSource sheet="Data" ref="A1" r:id="external-book"/><s:dataRef sheet="Data" ref="A1" r:id="external-data"/><s:hyperlink ref="A1" location="Data!A1" r:id="external-link"/></root>"#
        );
        assert!(
            rewrite(
                source.as_bytes(),
                &rename("Data", "Input"),
                &["Data"],
                equivalent,
            )
            .expect("scan")
            .is_none()
        );
    }

    #[test]
    fn swaps_and_quotes_local_prefixes_without_touching_external_books() {
        let renames = [
            Rename {
                before: "One",
                after: "Two",
            },
            Rename {
                before: "Two",
                after: "One",
            },
        ];
        let source = format!(
            r#"<root xmlns:s="{SML_NAMESPACE}"><s:f>One!A1+Two!LocalName+'One:Two'!$A$1+[0]One!A1+[1]One!A1+'[Book.xlsx]One'!A1+INDIRECT("One!A1")</s:f></root>"#
        );
        let changed = rewrite(source.as_bytes(), &renames, &[], equivalent)
            .expect("rewrite")
            .expect("changed output");
        let changed = std::str::from_utf8(&changed).expect("UTF-8");
        let expected = escape_xml(
            "Two!A1+One!LocalName+'Two:One'!$A$1+[0]Two!A1+[1]One!A1+'[Book.xlsx]One'!A1+INDIRECT(\"One!A1\")",
        );
        assert!(
            changed.contains(&format!("<s:f>{expected}</s:f>")),
            "{changed}"
        );
    }

    #[test]
    fn rewrites_quoted_apostrophes_and_unicode_caseless_names() {
        let source = format!(
            r#"<root xmlns:s="{SML_NAMESPACE}"><s:f>STRASSE!A1+'Straße'!Named</s:f></root>"#
        );
        let changed = rewrite(
            source.as_bytes(),
            &rename("Straße", "O'Brien 2026"),
            &[],
            equivalent,
        )
        .expect("rewrite")
        .expect("changed output");
        assert!(std::str::from_utf8(&changed)
            .expect("UTF-8")
            .contains("<s:f>&apos;O&apos;&apos;Brien 2026&apos;!A1+&apos;O&apos;&apos;Brien 2026&apos;!Named</s:f>"));
    }

    #[test]
    fn x14_formula_namespace_is_modeled_but_spoofed_formula_is_suspicious() {
        let source = format!(
            r#"<root xmlns:x14="{X14_NAMESPACE}" xmlns:xm="{EXCEL_FORMULA_NAMESPACE}"><x14:formula1><xm:f>Data!A1</xm:f></x14:formula1></root>"#
        );
        let changed = rewrite(source.as_bytes(), &rename("Data", "Input"), &[], equivalent)
            .expect("rewrite")
            .expect("changed output");
        assert!(
            std::str::from_utf8(&changed)
                .expect("UTF-8")
                .contains("<xm:f>Input!A1</xm:f>")
        );

        let spoofed = rewrite(
            br#"<root xmlns:x="urn:not-a-formula"><x:f>Data!A1</x:f></root>"#,
            &rename("Data", "Input"),
            &[],
            equivalent,
        );
        assert!(matches!(
            spoofed,
            Err(Error::Blocked(BlockReason::UnmodeledReference))
        ));
    }

    #[test]
    fn no_local_match_returns_exact_noop_without_output_allocation() {
        let source = format!(
            r#"<root xmlns:s="{SML_NAMESPACE}"><s:f>[2]One!A1+Table1[One]+"One!A1"</s:f></root>"#
        );
        assert!(
            rewrite(
                source.as_bytes(),
                &rename("One", "Renamed"),
                &[],
                equivalent,
            )
            .expect("scan")
            .is_none()
        );
    }

    #[test]
    fn reorders_titles_by_expanded_namespace_and_preserves_foreign_markup() {
        let source = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}" xmlns:foreign="urn:foreign"><foreign:note>keep</foreign:note><p:TitlesOfParts><v:vector baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr></v:vector></p:TitlesOfParts></p:Properties>"#
        );
        let changed = reorder_titles(source.as_bytes(), &["One", "Two"], &[1, 0], equivalent)
            .expect("reorder")
            .expect("changed output");
        let changed = std::str::from_utf8(&changed).expect("UTF-8");
        assert!(changed.contains("<v:lpstr>Two</v:lpstr><v:lpstr>One</v:lpstr>"));
        assert!(changed.contains("<foreign:note>keep</foreign:note>"));
        assert!(changed.contains("baseType=\"lpstr\""));
    }

    #[test]
    fn reorders_only_catalog_prefix_and_keeps_named_range_title_exact() {
        let source = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr data-keep="one">One</vt:lpstr><vt:lpstr>Two</vt:lpstr><vt:lpstr data-range="exact">Two!Print_Area</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        let changed = reorder_titles(source.as_bytes(), &["One", "Two"], &[1, 0], equivalent)
            .expect("reorder")
            .expect("changed output");
        let changed = std::str::from_utf8(&changed).expect("UTF-8");
        assert!(changed.contains(concat!(
            "<vt:lpstr data-keep=\"one\">Two</vt:lpstr>",
            "<vt:lpstr>One</vt:lpstr>",
            "<vt:lpstr data-range=\"exact\">Two!Print_Area</vt:lpstr>"
        )));
    }

    #[test]
    fn unchanged_catalog_slot_preserves_its_lexical_text() {
        let source = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><TitlesOfParts><vt:vector><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr><vt:lpstr><![CDATA[Three]]></vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        let changed = reorder_titles(
            source.as_bytes(),
            &["One", "Two", "Three"],
            &[1, 0, 2],
            equivalent,
        )
        .expect("reorder")
        .expect("changed output");
        assert!(
            std::str::from_utf8(&changed)
                .expect("UTF-8")
                .contains("<vt:lpstr><![CDATA[Three]]></vt:lpstr>")
        );
    }

    #[test]
    fn absent_optional_titles_need_no_rewrite() {
        let source = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}"><Application>Producer</Application></Properties>"#
        );
        assert!(
            reorder_titles(source.as_bytes(), &["One", "Two"], &[1, 0], equivalent)
                .expect("optional metadata")
                .is_none()
        );
    }

    #[test]
    fn rejects_non_permutations_before_scanning_source() {
        let malformed = b"<not-xml";
        assert!(matches!(
            reorder_titles(malformed, &["One", "Two"], &[0], equivalent),
            Err(Error::Invalid(_))
        ));
        assert!(matches!(
            reorder_titles(malformed, &["One", "Two"], &[0, 0], equivalent),
            Err(Error::Invalid(_))
        ));
        assert!(matches!(
            reorder_titles(malformed, &["One", "Two"], &[0, 2], equivalent),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn identity_order_is_an_exact_noop_even_for_malformed_xml() {
        assert_eq!(
            reorder_titles(b"<not-xml", &["One", "Two"], &[0, 1], equivalent)
                .expect("identity no-op"),
            None
        );
    }

    #[test]
    fn rejects_stale_nested_and_markup_compatibility_title_contexts() {
        let stale = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><TitlesOfParts><vt:vector><vt:lpstr>Stale</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        assert!(matches!(
            reorder_titles(stale.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Invalid(_))
        ));

        let nested = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><TitlesOfParts><vt:vector><vt:lpstr><vt:rich>One</vt:rich></vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        assert!(matches!(
            reorder_titles(nested.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Invalid(_))
        ));

        let mce = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="x"><TitlesOfParts><vt:vector><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></mc:Choice></mc:AlternateContent></Properties>"#
        );
        assert!(matches!(
            reorder_titles(mce.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Blocked(BlockReason::MarkupCompatibility))
        ));
    }

    #[test]
    fn rejects_excessive_nesting_before_materializing_a_reordered_part() {
        let mut source = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><Application>"#
        );
        source.push_str(&"<x>".repeat(MAX_REFERENCE_DEPTH));
        source.push_str(&"</x>".repeat(MAX_REFERENCE_DEPTH));
        source.push_str(
            r#"</Application><TitlesOfParts><vt:vector><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#,
        );
        assert!(matches!(
            reorder_titles(source.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Limit(Limit::NestingDepth { .. }))
        ));
    }

    #[test]
    fn rejects_duplicate_raw_and_expanded_title_attributes() {
        let raw_duplicate = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><TitlesOfParts><vt:vector a="1" a="2"><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        assert!(matches!(
            reorder_titles(
                raw_duplicate.as_bytes(),
                &["One", "Two"],
                &[1, 0],
                equivalent
            ),
            Err(Error::Attribute { .. })
        ));

        let expanded_duplicate = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}" xmlns:a="urn:duplicate" xmlns:b="urn:duplicate"><TitlesOfParts><vt:vector a:keep="1" b:keep="2"><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        assert!(matches!(
            reorder_titles(
                expanded_duplicate.as_bytes(),
                &["One", "Two"],
                &[1, 0],
                equivalent
            ),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn rejects_markup_compatibility_attributes_even_without_an_alternate_branch() {
        let source = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="future"><TitlesOfParts><vt:vector><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        assert!(matches!(
            reorder_titles(source.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Blocked(BlockReason::MarkupCompatibility))
        ));
    }

    #[test]
    fn rejects_foreign_children_that_make_the_title_prefix_ambiguous() {
        let source = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}" xmlns:foreign="urn:foreign"><TitlesOfParts><foreign:vector><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></foreign:vector></TitlesOfParts></Properties>"#
        );
        assert!(matches!(
            reorder_titles(source.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn rejects_non_whitespace_text_outside_root_and_unclosed_markup() {
        let source = format!(
            r#"junk<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><TitlesOfParts><vt:vector><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#
        );
        assert!(matches!(
            reorder_titles(source.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Invalid(_))
        ));

        let unclosed = format!(
            r#"<Properties xmlns="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:vt="{PROPERTY_TYPES_NAMESPACE}"><TitlesOfParts><vt:vector><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr></vt:vector></TitlesOfParts>"#
        );
        assert!(matches!(
            reorder_titles(unclosed.as_bytes(), &["One", "Two"], &[1, 0], equivalent),
            Err(Error::Invalid(_)) | Err(Error::Parse { .. })
        ));
    }

    fn insertion_source(label: &str) -> String {
        format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}"><p:HeadingPairs><v:vector size="2" baseType="variant"><v:variant><v:lpstr>{label}</v:lpstr></v:variant><v:variant><v:i4>2</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="2" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr></v:vector></p:TitlesOfParts><Company data-keep="yes"/></p:Properties>"#
        )
    }

    #[test]
    fn inserts_catalog_title_at_start_middle_and_end() {
        for (position, expected) in [
            (0, "New &amp; Sheet,One,Two"),
            (1, "One,New &amp; Sheet,Two"),
            (2, "One,Two,New &amp; Sheet"),
        ] {
            let source = insertion_source("Worksheets");
            let output = insert_title(
                &source.into_bytes(),
                &["One", "Two"],
                position,
                "New & Sheet",
                equivalent,
            )
            .expect("insert")
            .expect("metadata present");
            let text = std::str::from_utf8(&output).expect("UTF-8");
            let catalog = expected.replace(',', "</v:lpstr><v:lpstr>");
            assert!(
                text.contains(&format!("<v:lpstr>{catalog}</v:lpstr>")),
                "{text}"
            );
            assert!(text.contains("<v:i4>3</v:i4>"), "{text}");
            assert!(text.contains("size=\"3\""), "{text}");
            assert!(text.contains("<Company data-keep=\"yes\"/>"), "{text}");
        }
    }

    #[test]
    fn inserts_before_trailing_named_range_and_supports_xlsb_sheet_group() {
        let source = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}"><p:HeadingPairs><v:vector size="4" baseType="variant"><v:variant><v:lpstr>Sheet</v:lpstr></v:variant><v:variant><v:i4>2</v:i4></v:variant><v:variant><v:lpstr>Named Ranges</v:lpstr></v:variant><v:variant><v:i4>1</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="3" baseType="lpstr"><v:lpstr data-keep="one">One</v:lpstr><v:lpstr>Two</v:lpstr><v:lpstr data-range="exact">Two!Print_Area</v:lpstr></v:vector></p:TitlesOfParts></p:Properties>"#
        );
        let output = insert_title(source.as_bytes(), &["One", "Two"], 1, "A&B", equivalent)
            .expect("insert")
            .expect("metadata present");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<v:i4>3</v:i4>"), "{text}");
        assert!(text.contains("size=\"4\""), "{text}");
        assert!(text.contains(
            "<v:lpstr data-keep=\"one\">One</v:lpstr><v:lpstr>A&amp;B</v:lpstr><v:lpstr>Two</v:lpstr><v:lpstr data-range=\"exact\">Two!Print_Area</v:lpstr>"
        ));
    }

    #[test]
    fn accepts_padded_catalog_labels_consistently() {
        let source = insertion_source(" Worksheets ");
        let output = insert_title(source.as_bytes(), &["One", "Two"], 2, "Three", equivalent)
            .expect("insert")
            .expect("metadata present");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<v:i4>3</v:i4>"), "{text}");
    }

    #[test]
    fn refuses_reversed_groups_undeclared_trailing_titles_and_missing_heading_metadata() {
        let reversed = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}"><p:HeadingPairs><v:vector size="4" baseType="variant"><v:variant><v:lpstr>Named Ranges</v:lpstr></v:variant><v:variant><v:i4>1</v:i4></v:variant><v:variant><v:lpstr>Sheet</v:lpstr></v:variant><v:variant><v:i4>2</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="3" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr><v:lpstr>Two!Print_Area</v:lpstr></v:vector></p:TitlesOfParts></p:Properties>"#
        );
        assert!(matches!(
            insert_title(reversed.as_bytes(), &["One", "Two"], 0, "New", equivalent),
            Err(Error::Invalid(_))
        ));

        let undeclared_trailing = insertion_source("Worksheets").replace(
            r#"size="2" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr>"#,
            r#"size="3" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr><v:lpstr>Two!Print_Area</v:lpstr>"#,
        );
        assert!(matches!(
            insert_title(
                undeclared_trailing.as_bytes(),
                &["One", "Two"],
                0,
                "New",
                equivalent
            ),
            Err(Error::Invalid(_))
        ));

        let missing_heading_size = insertion_source("Worksheets").replace(
            r#"<v:vector size="2" baseType="variant">"#,
            r#"<v:vector baseType="variant">"#,
        );
        assert!(matches!(
            insert_title(
                missing_heading_size.as_bytes(),
                &["One", "Two"],
                0,
                "New",
                equivalent
            ),
            Err(Error::Invalid(_))
        ));

        let missing_heading_base_type = insertion_source("Worksheets").replace(
            r#"<v:vector size="2" baseType="variant">"#,
            r#"<v:vector size="2">"#,
        );
        assert!(matches!(
            insert_title(
                missing_heading_base_type.as_bytes(),
                &["One", "Two"],
                0,
                "New",
                equivalent
            ),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn absent_optional_metadata_is_an_exact_noop() {
        let source = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}"><p:Application>Producer</p:Application></p:Properties>"#
        );
        assert_eq!(
            insert_title(source.as_bytes(), &["One"], 0, "New", equivalent)
                .expect("optional metadata"),
            None
        );
        let partial = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}"><p:TitlesOfParts/></p:Properties>"#
        );
        assert!(matches!(
            insert_title(partial.as_bytes(), &["One"], 0, "New", equivalent),
            Err(Error::Invalid(_))
        ));
    }

    #[test]
    fn validates_counts_prefix_and_namespace_scope_without_rewriting_other_bytes() {
        let source = format!(
            r#"<weird:Properties xmlns:weird="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:q="{PROPERTY_TYPES_NAMESPACE}"><weird:HeadingPairs><q:vector data-before="x" size="2" baseType="variant"><q:variant><q:lpstr>Worksheets</q:lpstr></q:variant><q:variant><q:i4>2</q:i4></q:variant></q:vector></weird:HeadingPairs><weird:TitlesOfParts><q:vector data-before="y" size="2" baseType="lpstr"><q:lpstr>One</q:lpstr><q:lpstr>Two</q:lpstr></q:vector></weird:TitlesOfParts><weird:Company data-keep="yes"/></weird:Properties>"#
        );
        let output = insert_title(source.as_bytes(), &["One", "Two"], 2, "Three", equivalent)
            .expect("insert")
            .expect("metadata present");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<q:vector data-before=\"x\" size=\"2\" baseType=\"variant\">"));
        assert!(text.contains("<q:vector data-before=\"y\" size=\"3\" baseType=\"lpstr\">"));
        assert!(text.contains("<weird:Company data-keep=\"yes\"/>"));
        assert!(text.contains("<q:lpstr>Three</q:lpstr>"));
    }

    #[test]
    fn inserts_into_an_empty_catalog_using_the_vector_prefix() {
        let source = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}"><p:HeadingPairs><v:vector size="2" baseType="variant"><v:variant><v:lpstr>Sheet</v:lpstr></v:variant><v:variant><v:i4>0</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="0" baseType="lpstr"></v:vector></p:TitlesOfParts></p:Properties>"#
        );
        let output = insert_title(source.as_bytes(), &[], 0, "First", equivalent)
            .expect("insert")
            .expect("metadata present");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<v:i4>1</v:i4>"));
        assert!(text.contains(
            "<v:vector size=\"1\" baseType=\"lpstr\"><v:lpstr>First</v:lpstr></v:vector>"
        ));
    }

    #[test]
    fn inserted_title_preserves_carriage_returns_and_refuses_invalid_xml_characters() {
        let source = insertion_source("Worksheets");
        let name = "New\r\n\t&";
        let output = insert_title(source.as_bytes(), &["One", "Two"], 1, name, equivalent)
            .unwrap()
            .unwrap();
        assert!(
            std::str::from_utf8(&output)
                .unwrap()
                .contains("New&#13;\n\t&amp;")
        );
        insert_title(&output, &["One", name, "Two"], 3, "Last", equivalent)
            .expect("subsequent insert matches the original name");
        assert!(matches!(
            insert_title(
                source.as_bytes(),
                &["One", "Two"],
                1,
                "Bad\u{1}",
                equivalent
            ),
            Err(Error::Invalid(_))
        ));
    }

    fn removal_source(label: &str) -> String {
        format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}"><p:HeadingPairs><v:vector size="4" baseType="variant"><v:variant><v:lpstr>{label}</v:lpstr></v:variant><v:variant><v:i4>4</v:i4></v:variant><v:variant><v:lpstr>Named Ranges</v:lpstr></v:variant><v:variant><v:i4>1</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="5" baseType="lpstr"><v:lpstr data-one="1">One</v:lpstr><v:lpstr>Two</v:lpstr><v:lpstr><![CDATA[Three]]></v:lpstr><v:lpstr>Four</v:lpstr><v:lpstr data-range="exact">Four!Print_Area</v:lpstr></v:vector></p:TitlesOfParts></p:Properties>"#
        )
    }

    #[test]
    fn removes_catalog_title_at_start_middle_and_end() {
        for (positions, expected, size, count) in [
            (&[0usize][..], "Two,Three,Four", 4, 3),
            (&[1usize, 2][..], "One,Four", 3, 2),
            (&[3usize][..], "One,Two,Three", 4, 3),
        ] {
            let source = removal_source("Worksheets");
            let output = remove_titles(
                source.as_bytes(),
                &["One", "Two", "Three", "Four"],
                positions,
                equivalent,
            )
            .expect("remove")
            .expect("metadata present");
            let text = std::str::from_utf8(&output).expect("UTF-8");
            let catalog = expected.replace(',', "</v:lpstr><v:lpstr>");
            if expected == "Two,Three,Four" {
                assert!(
                    text.contains("<v:lpstr>Two</v:lpstr><v:lpstr><![CDATA[Three]]></v:lpstr><v:lpstr>Four</v:lpstr>"),
                    "{text}"
                );
            } else if expected == "One,Four" {
                assert!(
                    text.contains(r#"<v:lpstr data-one="1">One</v:lpstr><v:lpstr>Four</v:lpstr>"#),
                    "{text}"
                );
            } else if expected == "One,Two,Three" {
                assert!(
                    text.contains(
                        r#"<v:lpstr data-one="1">One</v:lpstr><v:lpstr>Two</v:lpstr><v:lpstr><![CDATA[Three]]></v:lpstr>"#
                    ),
                    "{text}"
                );
            } else {
                assert!(
                    text.contains(&format!("<v:lpstr>{catalog}</v:lpstr>")),
                    "{text}"
                );
            }
            assert!(
                text.contains(&format!("<v:vector size=\"{size}\"")),
                "{text}"
            );
            assert!(text.contains(&format!("<v:i4>{count}</v:i4>")), "{text}");
            assert!(
                text.contains(r#"<v:lpstr data-range="exact">Four!Print_Area</v:lpstr>"#),
                "{text}"
            );
            if expected != "One,Four" {
                assert!(text.contains("<![CDATA[Three]]>"), "{text}");
            }
        }
    }

    #[test]
    fn removes_multiple_titles_only_when_positions_are_sorted_and_unique() {
        let source = removal_source("Sheet");
        let output = remove_titles(
            source.as_bytes(),
            &["One", "Two", "Three", "Four"],
            &[0, 2],
            equivalent,
        )
        .expect("remove")
        .expect("metadata present");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<v:lpstr>Two</v:lpstr><v:lpstr>Four</v:lpstr>"));
        assert!(text.contains("<v:vector size=\"3\""));
        assert!(text.contains("<v:i4>2</v:i4>"));

        for positions in [[2, 0], [1, 1], [4, 0]] {
            assert!(matches!(
                remove_titles(
                    source.as_bytes(),
                    &["One", "Two", "Three", "Four"],
                    &positions,
                    equivalent,
                ),
                Err(Error::Invalid(_))
            ));
        }
    }

    #[test]
    fn removes_titles_with_local_namespace_bindings_and_preserves_lexical_survivors() {
        let source = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}"><p:HeadingPairs><v:vector xmlns:v="{PROPERTY_TYPES_NAMESPACE}" size="4" baseType="variant"><v:variant><v:lpstr>Worksheets</v:lpstr></v:variant><v:variant><v:i4>2</v:i4></v:variant><v:variant><v:lpstr>Named Ranges</v:lpstr></v:variant><v:variant><v:i4>1</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector xmlns:v="{PROPERTY_TYPES_NAMESPACE}" data-before="keep" size="3" baseType="lpstr"><v:lpstr data-one="1">One</v:lpstr><v:lpstr data-two="2">Two</v:lpstr><v:lpstr data-range="exact">Two!Print_Area</v:lpstr></v:vector></p:TitlesOfParts><p:Company data-keep="yes"/></p:Properties>"#
        );
        let output = remove_titles(source.as_bytes(), &["One", "Two"], &[0], equivalent)
            .expect("remove")
            .expect("metadata present");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(
            r#"<v:vector xmlns:v="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" data-before="keep" size="2" baseType="lpstr"><v:lpstr data-two="2">Two</v:lpstr><v:lpstr data-range="exact">Two!Print_Area</v:lpstr></v:vector>"#
        ));
        assert!(text.contains(r#"<p:Company data-keep="yes"/>"#));
        assert!(text.contains("<v:i4>1</v:i4>"));
    }

    #[test]
    fn removal_empty_positions_and_absent_metadata_are_exact_noops() {
        assert_eq!(
            remove_titles(b"<not-xml", &["One"], &[], equivalent).expect("empty no-op"),
            None
        );
        let absent = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}"><p:Application>Producer</p:Application></p:Properties>"#
        );
        assert_eq!(
            remove_titles(absent.as_bytes(), &["One"], &[0], equivalent).expect("absent metadata"),
            None
        );
    }

    #[test]
    fn removal_refuses_stale_unknown_and_markup_compatibility_metadata() {
        let stale = removal_source("Worksheets").replace(">One</v:lpstr>", ">Stale</v:lpstr>");
        assert!(matches!(
            remove_titles(
                stale.as_bytes(),
                &["One", "Two", "Three", "Four"],
                &[0],
                equivalent
            ),
            Err(Error::Invalid(_))
        ));
        let unknown = removal_source("Charts");
        assert!(matches!(
            remove_titles(
                unknown.as_bytes(),
                &["One", "Two", "Three", "Four"],
                &[0],
                equivalent
            ),
            Err(Error::Blocked(BlockReason::UnmodeledReference))
        ));
        let mce = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}" xmlns:mc="{MCE_STR}"><p:HeadingPairs><v:vector size="2" baseType="variant"><v:variant><v:lpstr>Worksheets</v:lpstr></v:variant><v:variant><v:i4>1</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="1" baseType="lpstr"><v:lpstr>One</v:lpstr></v:vector></p:TitlesOfParts><mc:AlternateContent/></p:Properties>"#
        );
        assert!(matches!(
            remove_titles(mce.as_bytes(), &["One"], &[0], equivalent),
            Err(Error::Blocked(BlockReason::MarkupCompatibility))
        ));
    }

    #[test]
    fn refuses_malformed_stale_grouped_and_mce_metadata() {
        let stale = insertion_source("Worksheets").replace(">One</v:lpstr>", ">Stale</v:lpstr>");
        assert!(matches!(
            insert_title(stale.as_bytes(), &["One", "Two"], 0, "New", equivalent),
            Err(Error::Invalid(_))
        ));
        let grouped = insertion_source("Charts");
        assert!(matches!(
            insert_title(grouped.as_bytes(), &["One", "Two"], 0, "New", equivalent),
            Err(Error::Blocked(BlockReason::UnmodeledReference))
        ));
        let mce = format!(
            r#"<p:Properties xmlns:p="{EXTENDED_PROPERTIES_NAMESPACE}" xmlns:v="{PROPERTY_TYPES_NAMESPACE}" xmlns:mc="{MCE_STR}"><p:HeadingPairs><v:vector size="2" baseType="variant"><v:variant><v:lpstr>Worksheets</v:lpstr></v:variant><v:variant><v:i4>2</v:i4></v:variant></v:vector></p:HeadingPairs><p:TitlesOfParts><v:vector size="2" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr></v:vector></p:TitlesOfParts><mc:AlternateContent/></p:Properties>"#
        );
        assert!(matches!(
            insert_title(mce.as_bytes(), &["One", "Two"], 0, "New", equivalent),
            Err(Error::Blocked(BlockReason::MarkupCompatibility))
        ));
    }
}
