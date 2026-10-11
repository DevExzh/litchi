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
}
