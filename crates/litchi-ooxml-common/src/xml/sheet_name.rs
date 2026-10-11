//! Lossless package-part surgery for local worksheet-name dependencies.
//!
//! Known formula and reference carriers are rewritten together. Unknown
//! formula-like carriers and markup-compatibility alternatives are rejected
//! before publication instead of being guessed at.

use std::collections::TryReserveError;
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
const PROPERTY_TYPES: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
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
}
