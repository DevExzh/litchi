//! Narrow capture of supported x14ac attributes before MCE preprocessing.
//!
//! Treating the complete extension namespace as understood would make
//! `MustUnderstand` unsound. This scanner instead records only direct
//! `dyDescent` attributes whose core parent structure is already modeled.

use std::{
    cell::RefCell,
    collections::BTreeMap,
    io::{BufRead, Cursor},
};

use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::ROWS;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use litchi_ooxml_common::mce::{
    Capabilities, Name, RawElement, RawElementKind, SemanticElement, SemanticEvent, StreamError,
    StreamLimits, process_markup_compatibility_stream_with_observers,
};

use super::{
    super::namespace::{
        SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE, is_spreadsheetml_name,
    },
    parse_one_based_row,
};
use crate::error::{Result, allocation, invalid};
use crate::layout::Descent;

pub(crate) const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const MAX_XML_DEPTH: usize = 256;
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MARKER: &[u8] = b"litchi_x14ac_dyDescent";

#[derive(Debug, Default)]
pub(crate) struct Values {
    pub(crate) defaults: Option<Descent>,
    pub(crate) rows: BTreeMap<u32, Descent>,
}

/// Diagnostic result for callback-scoped worksheet extension capture.
///
/// The MCE stream error is intentionally retained instead of being collapsed
/// into [`crate::Error`].  The OPC source-backed reader can therefore apply its
/// source, cancellation, and transport precedence before exposing a format
/// error to callers.
pub(crate) type StreamResult<T> = std::result::Result<T, StreamError<crate::Error, crate::Error>>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Worksheet,
    SheetData,
    Row,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum RowMode {
    /// Retain each validated row descent in the legacy map.
    Retain,
    /// Validate row numbers and descent values without retaining row values.
    ValidateOnly,
    /// Preserve the focused-defaults behavior and do not inspect row values.
    Ignore,
}

pub(crate) fn capture(content: &[u8]) -> Result<Values> {
    if has_markup_compatibility(content)
        && has_descent_attribute(content)
        && has_alternate_content(content)
    {
        let mut input = Cursor::new(content);
        return capture_stream_legacy(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
            true,
        );
    }
    capture_inner(content, true)
}

pub(crate) fn capture_defaults(content: &[u8]) -> Result<Option<Descent>> {
    if has_markup_compatibility(content)
        && has_descent_attribute(content)
        && has_alternate_content(content)
    {
        let mut input = Cursor::new(content);
        return capture_stream_defaults_legacy(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
        );
    }
    Ok(capture_inner(content, false)?.defaults)
}

/// Capture supported extension values from a callback-scoped MCE stream.
///
/// Raw observers run for every source `Start` and `Empty` event, including
/// hidden compatibility branches.  The active observer receives only the
/// selected semantic stream and consumes a raw candidate only for its matching
/// element event.  No event or source buffer escapes either callback.
#[expect(
    clippy::result_large_err,
    reason = "The stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
)]
pub(crate) fn capture_stream(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    capture_rows: bool,
) -> StreamResult<Values> {
    let row_mode = if capture_rows {
        RowMode::Retain
    } else {
        RowMode::Ignore
    };
    capture_stream_with_active(input, capabilities, limits, row_mode, |_event| {
        Ok::<(), crate::Error>(())
    })
}

/// Capture x14ac values while composing one active semantic observer.
///
/// The raw observer still runs before MCE selection, and the x14ac active
/// observer still consumes the selected stream before `active` does. This is
/// the internal join point for bounded source readers; both observers share
/// the committed MCE stream and no rewritten XML buffer is materialized.
/// `ValidateOnly` validates row lexical values without retaining the legacy
/// row map. This helper makes no fixed-memory or OOM-safety claim: quick-xml,
/// MCE, and observer allocations remain outside the input-buffer bound.
#[expect(
    clippy::result_large_err,
    reason = "The stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
)]
pub(super) fn capture_stream_with_active<Active>(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    row_mode: RowMode,
    mut active: Active,
) -> StreamResult<Values>
where
    Active: for<'a> FnMut(&SemanticEvent<'a>) -> Result<()>,
{
    let state = RefCell::new(StreamObserverState::new(row_mode));
    let result = process_markup_compatibility_stream_with_observers(
        input,
        capabilities,
        limits,
        |element| {
            let mut state = state
                .try_borrow_mut()
                .map_err(|_| invalid("worksheet extension observers were re-entered"))?;
            state.raw(element)
        },
        |event| {
            {
                let mut state = state
                    .try_borrow_mut()
                    .map_err(|_| invalid("worksheet extension observers were re-entered"))?;
                state.active(&event)?;
            }
            active(&event)
        },
    );
    match result {
        Ok(_) => Ok(state.into_inner().finish()),
        Err(error) => Err(error),
    }
}

/// Capture only the selected worksheet default from a callback-scoped stream.
#[expect(
    clippy::result_large_err,
    reason = "The stream error intentionally retains typed primary plus raw/active callback diagnostics; boxing it would change the established API."
)]
pub(crate) fn capture_stream_defaults(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
) -> StreamResult<Option<Descent>> {
    capture_stream(input, capabilities, limits, false).map(|values| values.defaults)
}

/// Capture extension values through the legacy XLSX error surface.
///
/// This adapter deliberately keeps the historical precedence: a raw observer
/// error wins over an input or MCE error; otherwise the primary input/MCE error
/// wins over an active observer error.  The source-backed OPC layer should use
/// [`capture_stream`] so its source, transport, and cancellation error can
/// outrank this diagnostic callback.
pub(crate) fn capture_stream_legacy(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    capture_rows: bool,
) -> Result<Values> {
    capture_stream(input, capabilities, limits, capture_rows).map_err(map_stream_error)
}

/// Capture only the worksheet default through the legacy XLSX error surface.
pub(crate) fn capture_stream_defaults_legacy(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
) -> Result<Option<Descent>> {
    capture_stream_defaults(input, capabilities, limits).map_err(map_stream_error)
}

fn map_stream_error(error: StreamError<crate::Error, crate::Error>) -> crate::Error {
    match error {
        StreamError::Input {
            raw_error: Some(raw_error),
            ..
        }
        | StreamError::Mce {
            raw_error: Some(raw_error),
            ..
        }
        | StreamError::Callback {
            raw_error: Some(raw_error),
            ..
        } => raw_error,
        StreamError::Input { error, .. } => {
            crate::Error::Package(litchi_opc::OpcError::IoError(error))
        },
        StreamError::Mce {
            error,
            prior_mce_error,
            ..
        } => map_mce_stream_error(error, prior_mce_error),
        StreamError::Callback {
            raw_error: None,
            active_error: Some(active_error),
        } => active_error,
        StreamError::Callback {
            raw_error: None,
            active_error: None,
        } => invalid("MCE stream callback failure without an observer error"),
        _ => invalid("unknown MCE stream error"),
    }
}

fn map_mce_stream_error(
    error: litchi_ooxml_common::mce::Error,
    _prior_mce_error: Option<litchi_ooxml_common::mce::Error>,
) -> crate::Error {
    if let litchi_ooxml_common::mce::Error::Xml(message) = error {
        return invalid(format!("invalid worksheet extension XML: {message}"));
    }
    crate::Error::MarkupCompatibility(error)
}

#[derive(Debug)]
struct PendingCandidate {
    kind: RawElementKind,
    qualified_name: Box<[u8]>,
    expanded_name: Name,
    decoded_value: Box<str>,
    sequence: u64,
}

#[derive(Debug)]
struct StreamObserverState {
    row_mode: RowMode,
    stack: Vec<Context>,
    values: Values,
    previous_row: u32,
    pending: Option<PendingCandidate>,
    raw_sequence: u64,
}

impl StreamObserverState {
    fn new(row_mode: RowMode) -> Self {
        Self {
            row_mode,
            stack: Vec::new(),
            values: Values::default(),
            previous_row: 0,
            pending: None,
            raw_sequence: 0,
        }
    }

    fn finish(self) -> Values {
        self.values
    }

    fn raw(&mut self, element: RawElement<'_>) -> Result<()> {
        // A candidate belongs to exactly one raw element.  Clearing it before
        // inspecting the new event prevents an inactive event from surviving
        // until a later selected event.  The sequence is an additional
        // lifecycle proof: an active callback may consume only this raw
        // invocation's candidate.
        self.pending = None;
        self.raw_sequence = self
            .raw_sequence
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet extension event sequence overflow"))?;

        let mut candidate = None;
        for attribute in &element.attributes {
            if attribute.qualified_name.as_ref() == MARKER {
                return Err(invalid("worksheet XML uses a reserved internal marker"));
            }
            if attribute.expanded_name.local_name == "dyDescent"
                && attribute.expanded_name.namespace.as_bytes() == NAMESPACE
            {
                if candidate.replace(attribute).is_some() {
                    return Err(invalid("duplicate x14ac:dyDescent attribute"));
                }
            }
        }

        let Some(attribute) = candidate else {
            return Ok(());
        };
        self.pending = Some(PendingCandidate {
            kind: element.kind,
            qualified_name: copy_bytes(
                element.qualified_name.as_ref(),
                "worksheet extension pending element name",
            )?,
            expanded_name: copy_name(
                &element.expanded_name,
                "worksheet extension pending expanded element name",
            )?,
            decoded_value: copy_str(
                attribute.decoded(),
                "worksheet extension pending attribute value",
            )?,
            sequence: self.raw_sequence,
        });
        Ok(())
    }

    fn active(&mut self, event: &SemanticEvent<'_>) -> Result<()> {
        match event {
            SemanticEvent::Start(element) => self.selected_start(element, RawElementKind::Start),
            SemanticEvent::Empty(element) => self.selected_start(element, RawElementKind::Empty),
            SemanticEvent::End(_) => {
                // An unwrapped or ignored source element can leave a pending
                // raw candidate without an active start event.  It must never
                // be considered by a later semantic event.
                self.pending = None;
                self.stack
                    .pop()
                    .ok_or_else(|| invalid("worksheet extension XML has an unexpected end"))?;
                Ok(())
            },
            SemanticEvent::Text(_)
            | SemanticEvent::CData(_)
            | SemanticEvent::Comment(_)
            | SemanticEvent::Decl(_)
            | SemanticEvent::GeneralRef(_) => {
                self.pending = None;
                Ok(())
            },
            _ => {
                // Future MCE semantic events are not worksheet structure we
                // understand. Ignore them without changing the selected
                // context or retaining a raw candidate.
                self.pending = None;
                Ok(())
            },
        }
    }

    fn selected_start(
        &mut self,
        element: &SemanticElement<'_>,
        kind: RawElementKind,
    ) -> Result<()> {
        if kind == RawElementKind::Start && self.stack.len() >= MAX_XML_DEPTH {
            return Err(invalid(format!(
                "worksheet extension XML exceeds {MAX_XML_DEPTH} levels"
            )));
        }
        let candidate = self.take_candidate(kind, element);
        let parent = self.stack.last().copied();
        let context = if parent.is_none() && is_spreadsheetml_element(element, "worksheet") {
            Context::Worksheet
        } else if parent == Some(Context::Worksheet)
            && is_spreadsheetml_element(element, "sheetFormatPr")
        {
            if let Some(candidate) = candidate {
                let value = parse_descent_lexical(&candidate.decoded_value)?;
                if self.values.defaults.replace(value).is_some() {
                    return Err(invalid("duplicate worksheet default dyDescent"));
                }
            }
            Context::Other
        } else if parent == Some(Context::Worksheet)
            && is_spreadsheetml_element(element, "sheetData")
        {
            Context::SheetData
        } else if parent == Some(Context::SheetData) && is_spreadsheetml_element(element, "row") {
            if !matches!(self.row_mode, RowMode::Ignore) {
                let number = match unqualified_semantic_attribute(element, "r") {
                    Some(value) => parse_one_based_row(value)?,
                    None => self
                        .previous_row
                        .checked_add(1)
                        .filter(|value| *value <= ROWS)
                        .ok_or_else(|| {
                            invalid("inferred extension row exceeds the spreadsheet grid")
                        })?,
                };
                self.previous_row = number;
                if let Some(candidate) = candidate {
                    let value = parse_descent_lexical(&candidate.decoded_value)?;
                    if matches!(self.row_mode, RowMode::Retain)
                        && self.values.rows.insert(number, value).is_some()
                    {
                        return Err(invalid(format!(
                            "duplicate worksheet row {number} dyDescent"
                        )));
                    }
                }
            }
            Context::Row
        } else {
            // The candidate is deliberately dropped on non-structural
            // elements.  Their x14ac value is not part of the modeled surface.
            Context::Other
        };

        if kind == RawElementKind::Start {
            self.stack
                .try_reserve(1)
                .map_err(|source| allocation("worksheet extension XML element stack", source))?;
            self.stack.push(context);
        }
        Ok(())
    }

    fn take_candidate(
        &mut self,
        kind: RawElementKind,
        element: &SemanticElement<'_>,
    ) -> Option<PendingCandidate> {
        let candidate = self.pending.take()?;
        if candidate.sequence != self.raw_sequence
            || candidate.kind != kind
            || candidate.qualified_name.as_ref() != element.qualified_name.as_ref()
            || candidate.expanded_name != element.expanded_name
        {
            return None;
        }
        Some(candidate)
    }
}

fn copy_bytes(value: &[u8], resource: &'static str) -> Result<Box<[u8]>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    copy.extend_from_slice(value);
    Ok(copy.into_boxed_slice())
}

fn copy_str(value: &str, resource: &'static str) -> Result<Box<str>> {
    Ok(copy_string(value, resource)?.into_boxed_str())
}

fn copy_string(value: &str, resource: &'static str) -> Result<String> {
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    copy.push_str(value);
    Ok(copy)
}

fn copy_name(value: &Name, resource: &'static str) -> Result<Name> {
    Ok(Name {
        namespace: copy_string(value.namespace.as_str(), resource)?,
        local_name: copy_string(value.local_name.as_str(), resource)?,
    })
}

fn is_spreadsheetml_element(element: &SemanticElement<'_>, local_name: &str) -> bool {
    (element.expanded_name.namespace.as_bytes() == SPREADSHEETML_NAMESPACE
        || element.expanded_name.namespace.as_bytes() == STRICT_SPREADSHEETML_NAMESPACE)
        && element.expanded_name.local_name == local_name
}

fn unqualified_semantic_attribute<'a>(
    element: &'a SemanticElement<'_>,
    local_name: &str,
) -> Option<&'a str> {
    element
        .attributes
        .iter()
        .find(|attribute| {
            attribute.expanded_name.namespace.is_empty()
                && attribute.expanded_name.local_name == local_name
        })
        .map(|attribute| attribute.decoded_value.as_ref())
}

fn capture_inner(content: &[u8], capture_rows: bool) -> Result<Values> {
    let mut reader = NsReader::from_reader(content);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::<Context>::new();
    let mut values = Values::default();
    let mut previous_row = 0u32;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid worksheet extension XML: {error}")))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let decoder = reader.decoder();
        let resolver = reader.resolver();
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "worksheet extension XML exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
                let context = start(
                    stack.last().copied(),
                    &namespace,
                    &element,
                    decoder,
                    resolver,
                    &mut previous_row,
                    &mut values,
                    capture_rows,
                )?;
                stack.push(context);
            },
            Event::Empty(element) => {
                start(
                    stack.last().copied(),
                    &namespace,
                    &element,
                    decoder,
                    resolver,
                    &mut previous_row,
                    &mut values,
                    capture_rows,
                )?;
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| invalid("worksheet extension XML has an unexpected end"))?;
            },
            Event::Eof if !stack.is_empty() => {
                return Err(invalid(
                    "worksheet extension XML has an unterminated element",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(values)
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments correspond directly to the x14ac worksheet attributes"
)]
fn start(
    parent: Option<Context>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    previous_row: &mut u32,
    values: &mut Values,
    capture_rows: bool,
) -> Result<Context> {
    if parent.is_none() && is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
        return Ok(Context::Worksheet);
    }
    if parent == Some(Context::Worksheet)
        && is_spreadsheetml_name(namespace, element.name(), b"sheetFormatPr")
    {
        if let Some(value) = descent(element, decoder, resolver)?
            && values.defaults.replace(value).is_some()
        {
            return Err(invalid("duplicate worksheet default dyDescent"));
        }
        return Ok(Context::Other);
    }
    if parent == Some(Context::Worksheet)
        && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
    {
        return Ok(Context::SheetData);
    }
    if parent == Some(Context::SheetData)
        && is_spreadsheetml_name(namespace, element.name(), b"row")
    {
        if !capture_rows {
            return Ok(Context::Row);
        }
        let number = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(value) => parse_one_based_row(&value)?,
            None => previous_row
                .checked_add(1)
                .filter(|value| *value <= ROWS)
                .ok_or_else(|| invalid("inferred extension row exceeds the spreadsheet grid"))?,
        };
        *previous_row = number;
        if let Some(value) = descent(element, decoder, resolver)?
            && values.rows.insert(number, value).is_some()
        {
            return Err(invalid(format!(
                "duplicate worksheet row {number} dyDescent"
            )));
        }
        return Ok(Context::Row);
    }
    Ok(Context::Other)
}

fn has_markup_compatibility(content: &[u8]) -> bool {
    content
        .windows(MCE_NAMESPACE.len())
        .any(|window| window == MCE_NAMESPACE)
}

fn has_descent_attribute(content: &[u8]) -> bool {
    may_contain_descent(content)
}

pub(super) fn may_contain_descent(content: &[u8]) -> bool {
    memchr::memmem::find(content, b"dyDescent").is_some()
}

fn has_alternate_content(content: &[u8]) -> bool {
    content
        .windows(b"AlternateContent".len())
        .any(|window| window == b"AlternateContent")
}

pub(crate) fn descent(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<Descent>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let is_target = local.as_ref() == b"dyDescent"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == NAMESPACE);
        if !is_target {
            continue;
        }
        let lexical = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        let parsed = parse_descent_lexical(&lexical)?;
        if result.replace(parsed).is_some() {
            return Err(invalid("duplicate x14ac:dyDescent attribute"));
        }
    }
    Ok(result)
}

fn parse_descent_lexical(lexical: &str) -> Result<Descent> {
    let parsed = lexical
        .parse::<f64>()
        .map_err(|_source| invalid(format!("invalid x14ac:dyDescent value '{lexical}'")))?;
    Ok(Descent::new(parsed)?)
}

pub(crate) fn attribute_name(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<Option<Box<str>>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() != b"dyDescent"
            || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == NAMESPACE)
        {
            continue;
        }
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("x14ac attribute name is not UTF-8: {error}")))?
            .to_owned()
            .into_boxed_str();
        if result.replace(name).is_some() {
            return Err(invalid("duplicate x14ac:dyDescent attribute"));
        }
    }
    Ok(result)
}
