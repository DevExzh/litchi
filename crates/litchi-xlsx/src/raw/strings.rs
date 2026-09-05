//! Streaming parser for the workbook shared-string table.

use std::io::BufRead;

use litchi_ooxml_common::mce::{
    Capabilities, Name, SemanticElement, SemanticEvent, StreamError, StreamLimits,
    process_markup_compatibility_stream_with_observers,
};
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::cell::Text;
use crate::error::{Result, allocation, invalid};
use crate::raw::namespace::{
    SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE, is_spreadsheetml_name,
};

const MAX_CELL_CHARACTERS: usize = 32_767;
// A supplementary Unicode scalar can occupy two seven-byte `_xHHHH_`
// SpreadsheetML escapes before decoding.
const MAX_ENCODED_TEXT_BYTES: usize = MAX_CELL_CHARACTERS * 14;
const MAX_RUNS: u32 = 32_767;
const MAX_OFFICE_COUNT: u64 = 2_147_483_647;
const MAX_PREALLOCATED_STRINGS: usize = 4_096;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Table,
    Item,
    Run,
    Text,
    Phonetic,
    Other,
}

#[derive(Debug, Default)]
struct Item {
    text: String,
    encoded_bytes: usize,
    runs: u32,
    phonetic_runs: u32,
    saw_simple: bool,
    saw_run: bool,
}

#[derive(Debug, Default)]
struct Parser {
    strings: Vec<Text>,
    item: Option<Item>,
    seen_text_in_run: bool,
}

pub(crate) fn parse(content: &[u8]) -> Result<Box<[Text]>> {
    let processed = litchi_ooxml_common::mce::process_ooxml(content)?;
    let content = std::str::from_utf8(processed.as_ref())
        .map_err(|error| invalid(format!("shared strings XML is not UTF-8: {error}")))?;
    Parser::parse(content)
}

/// Result of a bounded MCE-selected shared-string dependency scan.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Selected {
    /// Number of active direct `si` items in the selected semantic stream.
    pub(crate) count: usize,
    /// Requested plain items that exist and are safely representable as plain
    /// text, in requested-index order.
    pub(crate) requested: Vec<(usize, Text)>,
    /// Whether the active stream contained a rich, extension, or foreign
    /// construct that this dependency scan does not model.
    pub(crate) unsupported_rich: bool,
}

/// Scan one MCE-selected shared-string part without materializing its table.
///
/// The stream is consumed through EOF.  Only requested plain items are
/// retained; every active item is still structurally and text-bound validated.
/// Unsupported rich, phonetic, extension, and foreign constructs are reported
/// in [`Selected::unsupported_rich`] rather than approximated as plain text.
#[allow(clippy::result_large_err)]
pub(crate) fn stream_selected(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
    requested: &[usize],
) -> std::result::Result<Selected, StreamError<crate::Error, crate::Error>> {
    validate_requested_indexes(requested).map_err(|error| StreamError::Callback {
        raw_error: None,
        active_error: Some(error),
    })?;
    let requested_limit = requested
        .len()
        .min(limits.processing.max_input_bytes)
        .min(limits.max_events);
    let mut parser = SelectedParser::new(requested, requested_limit);
    let _report = process_markup_compatibility_stream_with_observers(
        input,
        capabilities,
        limits,
        |_| Ok::<(), crate::Error>(()),
        |event| parser.event(event),
    )?;
    parser.finish().map_err(|error| StreamError::Callback {
        raw_error: None,
        active_error: Some(error),
    })
}

#[derive(Debug)]
struct SelectedParser<'a> {
    stack: Vec<Context>,
    item: Option<SelectedItem>,
    count: usize,
    requested_indexes: &'a [usize],
    next_requested: usize,
    requested_limit: usize,
    requested: Vec<(usize, Text)>,
    root_seen: bool,
    closed_root: bool,
    unsupported_rich: bool,
}

impl<'a> SelectedParser<'a> {
    fn new(requested_indexes: &'a [usize], requested_limit: usize) -> Self {
        Self {
            stack: Vec::new(),
            item: None,
            count: 0,
            requested_indexes,
            next_requested: 0,
            requested_limit,
            requested: Vec::new(),
            root_seen: false,
            closed_root: false,
            unsupported_rich: false,
        }
    }

    fn event(&mut self, event: SemanticEvent<'_>) -> Result<()> {
        match event {
            SemanticEvent::Start(element) => self.start(&element, false),
            SemanticEvent::Empty(element) => self.start(&element, true),
            SemanticEvent::End(element) => self.end(&element),
            SemanticEvent::Text(text) | SemanticEvent::CData(text)
                if self.stack.last() == Some(&Context::Text) =>
            {
                self.push_text(text.text())
            },
            SemanticEvent::GeneralRef(reference) if self.stack.last() == Some(&Context::Text) => {
                let name = std::str::from_utf8(reference.name.as_ref()).map_err(|error| {
                    invalid(format!("shared-string XML reference is not UTF-8: {error}"))
                })?;
                let reference = quick_xml::events::BytesRef::new(name);
                let value = decode_xml_reference(&reference)?;
                self.push_text(&value)
            },
            _ => Ok(()),
        }
    }

    fn start(&mut self, element: &SemanticElement<'_>, empty: bool) -> Result<()> {
        if self.stack.is_empty() {
            if self.root_seen || self.closed_root || !is_spreadsheetml_element(element, b"sst") {
                return Err(invalid(
                    "shared strings XML must have one SpreadsheetML sst root",
                ));
            }
            self.parse_root_hints(element)?;
            self.root_seen = true;
            if empty {
                self.closed_root = true;
            } else {
                self.push_context(Context::Table)?;
            }
            return Ok(());
        }

        let parent = *self
            .stack
            .last()
            .ok_or_else(|| invalid("shared strings XML is missing its root context"))?;
        let child = self.start_context(parent, element)?;
        if marks_unsupported(parent, element) {
            self.mark_unsupported();
        }
        if empty {
            if child == Context::Item {
                self.finish_item()?;
            }
        } else {
            self.push_context(child)?;
        }
        Ok(())
    }

    fn start_context(&mut self, parent: Context, element: &SemanticElement<'_>) -> Result<Context> {
        if parent == Context::Table && is_spreadsheetml_element(element, b"si") {
            if self.item.is_some() {
                return Err(invalid("nested shared-string item"));
            }
            let index = self.count;
            self.count = self
                .count
                .checked_add(1)
                .ok_or_else(|| invalid("shared-string item count overflow"))?;
            let retain_text = self.requested_indexes.get(self.next_requested) == Some(&index);
            if retain_text {
                self.next_requested += 1;
            }
            self.item = Some(SelectedItem::new(index, retain_text));
            return Ok(Context::Item);
        }
        if parent == Context::Item && is_spreadsheetml_element(element, b"t") {
            let item = self
                .item
                .as_mut()
                .ok_or_else(|| invalid("shared-string text outside an item"))?;
            if item.saw_simple || item.saw_run {
                return Err(invalid(
                    "shared-string item mixes or duplicates simple and rich text",
                ));
            }
            item.saw_simple = true;
            return Ok(Context::Text);
        }
        if parent == Context::Item && is_spreadsheetml_element(element, b"r") {
            let item = self
                .item
                .as_mut()
                .ok_or_else(|| invalid("shared-string run outside an item"))?;
            if item.saw_simple {
                return Err(invalid("shared-string item mixes simple and rich text"));
            }
            item.runs = item
                .runs
                .checked_add(1)
                .filter(|count| *count <= MAX_RUNS)
                .ok_or_else(|| invalid("shared string has too many rich-text runs"))?;
            item.saw_run = true;
            item.seen_text_in_run = false;
            return Ok(Context::Run);
        }
        if parent == Context::Run && is_spreadsheetml_element(element, b"t") {
            let item = self
                .item
                .as_mut()
                .ok_or_else(|| invalid("shared-string run text outside an item"))?;
            if item.seen_text_in_run {
                return Err(invalid("shared-string run has duplicate text"));
            }
            item.seen_text_in_run = true;
            return Ok(Context::Text);
        }
        if parent == Context::Item && is_spreadsheetml_element(element, b"rPh") {
            let item = self
                .item
                .as_mut()
                .ok_or_else(|| invalid("phonetic run outside a shared-string item"))?;
            item.phonetic_runs = item
                .phonetic_runs
                .checked_add(1)
                .filter(|count| *count <= MAX_RUNS)
                .ok_or_else(|| invalid("shared string has too many phonetic runs"))?;
            return Ok(Context::Phonetic);
        }
        Ok(Context::Other)
    }

    fn end(&mut self, element: &litchi_ooxml_common::mce::SemanticEnd<'_>) -> Result<()> {
        let context = self
            .stack
            .pop()
            .ok_or_else(|| invalid("shared strings XML has a closing element outside its root"))?;
        match context {
            Context::Item => self.finish_item(),
            Context::Table => {
                if !is_spreadsheetml_element_name(&element.expanded_name, b"sst") {
                    return Err(invalid(
                        "shared strings XML has an invalid root closing element",
                    ));
                }
                if self.item.is_some() {
                    return Err(invalid("shared-string item remained open at the root"));
                }
                self.closed_root = true;
                Ok(())
            },
            Context::Run | Context::Text | Context::Phonetic | Context::Other => Ok(()),
        }
    }

    fn push_text(&mut self, value: &str) -> Result<()> {
        let item = self
            .item
            .as_mut()
            .ok_or_else(|| invalid("shared-string text outside an item"))?;
        item.encoded_bytes = item
            .encoded_bytes
            .checked_add(value.len())
            .filter(|length| *length <= MAX_ENCODED_TEXT_BYTES)
            .ok_or_else(|| invalid("shared-string encoded text is too large"))?;
        item.text_state.push(value)?;
        if let Some(text) = item.text.as_mut() {
            text.try_reserve(value.len())
                .map_err(|source| allocation("shared-string text", source))?;
            text.push_str(value);
        }
        Ok(())
    }

    fn finish_item(&mut self) -> Result<()> {
        let mut item = self
            .item
            .take()
            .ok_or_else(|| invalid("missing shared-string item"))?;
        item.text_state.finish()?;
        if item.text.is_none() || item.unsupported {
            return Ok(());
        }
        let text = item
            .text
            .take()
            .ok_or_else(|| invalid("missing requested shared-string scratch"))?;
        let text = decode_spreadsheet_text(&text)?;
        if text.chars().count() > MAX_CELL_CHARACTERS {
            return Err(invalid(format!(
                "shared string exceeds {MAX_CELL_CHARACTERS} characters"
            )));
        }
        if self.requested.len() >= self.requested_limit {
            return Err(invalid("shared-string requested output limit exceeded"));
        }
        self.requested
            .try_reserve(1)
            .map_err(|source| allocation("shared-string requested items", source))?;
        self.requested.push((item.index, text.into()));
        Ok(())
    }

    fn parse_root_hints(&self, element: &SemanticElement<'_>) -> Result<()> {
        if let Some(value) = root_attribute(element, "uniqueCount") {
            parse_count_hint_value(value, b"uniqueCount")?;
        }
        if let Some(value) = root_attribute(element, "count") {
            parse_count_hint_value(value, b"count")?;
        }
        Ok(())
    }

    fn push_context(&mut self, context: Context) -> Result<()> {
        self.stack
            .try_reserve(1)
            .map_err(|source| allocation("shared-string element stack", source))?;
        self.stack.push(context);
        Ok(())
    }

    fn mark_unsupported(&mut self) {
        self.unsupported_rich = true;
        if let Some(item) = self.item.as_mut() {
            item.unsupported = true;
        }
    }

    fn finish(self) -> Result<Selected> {
        if !self.root_seen || !self.closed_root || !self.stack.is_empty() || self.item.is_some() {
            return Err(invalid(
                "shared strings XML has an invalid root or element stack",
            ));
        }
        Ok(Selected {
            count: self.count,
            requested: self.requested,
            unsupported_rich: self.unsupported_rich,
        })
    }
}

fn validate_requested_indexes(requested: &[usize]) -> Result<()> {
    if requested.windows(2).any(|window| window[0] >= window[1]) {
        return Err(invalid(
            "requested shared-string indexes must be sorted and strictly unique",
        ));
    }
    Ok(())
}

#[derive(Debug)]
struct SelectedItem {
    index: usize,
    text: Option<String>,
    text_state: SpreadsheetTextState,
    encoded_bytes: usize,
    runs: u32,
    phonetic_runs: u32,
    saw_simple: bool,
    saw_run: bool,
    seen_text_in_run: bool,
    unsupported: bool,
}

impl SelectedItem {
    fn new(index: usize, retain_text: bool) -> Self {
        Self {
            index,
            text: retain_text.then(String::new),
            text_state: SpreadsheetTextState::default(),
            encoded_bytes: 0,
            runs: 0,
            phonetic_runs: 0,
            saw_simple: false,
            saw_run: false,
            seen_text_in_run: false,
            unsupported: false,
        }
    }
}

#[derive(Debug, Default)]
struct SpreadsheetTextState {
    candidate: [u8; 7],
    candidate_len: usize,
    high_surrogate: Option<u16>,
    decoded_chars: usize,
}

impl SpreadsheetTextState {
    fn push(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            self.push_character(character)?;
        }
        Ok(())
    }

    fn push_character(&mut self, character: char) -> Result<()> {
        if self.candidate_len != 0 {
            if character.is_ascii() && self.candidate_len < self.candidate.len() {
                self.candidate[self.candidate_len] = character as u8;
                self.candidate_len += 1;
                if self.candidate_len == self.candidate.len() {
                    self.resolve_candidate()?;
                }
                return Ok(());
            }
            self.flush_candidate_literal()?;
        }

        if self.high_surrogate.is_some() {
            if character != '_' {
                return Err(invalid("unpaired high surrogate in SpreadsheetML escape"));
            }
            self.candidate[0] = b'_';
            self.candidate_len = 1;
        } else if character == '_' {
            self.candidate[0] = b'_';
            self.candidate_len = 1;
        } else {
            self.add_decoded_chars(1)?;
        }
        Ok(())
    }

    fn finish(&mut self) -> Result<()> {
        if self.high_surrogate.is_some() {
            return Err(invalid("unpaired high surrogate in SpreadsheetML escape"));
        }
        if self.candidate_len != 0 {
            self.flush_candidate_literal()?;
        }
        Ok(())
    }

    fn resolve_candidate(&mut self) -> Result<()> {
        let candidate = self.candidate;
        let unit = spreadsheet_escape_at(&candidate[..candidate.len()], 0).map(|(unit, _)| unit);
        self.candidate_len = 0;
        let Some(unit) = unit else {
            if self.high_surrogate.is_some() {
                return Err(invalid("unpaired high surrogate in SpreadsheetML escape"));
            }
            self.add_decoded_chars(1)?;
            for byte in &candidate[1..] {
                self.push_character(char::from(*byte))?;
            }
            return Ok(());
        };

        if let Some(high) = self.high_surrogate.take() {
            if !(0xDC00..=0xDFFF).contains(&unit) {
                return Err(invalid(format!(
                    "unpaired high surrogate in SpreadsheetML escape (high {high:04X})"
                )));
            }
            self.add_decoded_chars(1)?;
        } else if (0xD800..=0xDBFF).contains(&unit) {
            self.high_surrogate = Some(unit);
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return Err(invalid("unpaired low surrogate in SpreadsheetML escape"));
        } else {
            self.add_decoded_chars(1)?;
        }
        Ok(())
    }

    fn flush_candidate_literal(&mut self) -> Result<()> {
        let length = self.candidate_len;
        self.candidate_len = 0;
        self.add_decoded_chars(length)
    }

    fn add_decoded_chars(&mut self, count: usize) -> Result<()> {
        self.decoded_chars = self
            .decoded_chars
            .checked_add(count)
            .filter(|count| *count <= MAX_CELL_CHARACTERS)
            .ok_or_else(|| {
                invalid(format!(
                    "shared string exceeds {MAX_CELL_CHARACTERS} characters"
                ))
            })?;
        Ok(())
    }
}

fn root_attribute<'a>(element: &'a SemanticElement<'_>, local_name: &str) -> Option<&'a str> {
    element
        .attrs()
        .iter()
        .find(|attribute| {
            attribute.expanded_name.namespace.is_empty()
                && attribute.expanded_name.local_name == local_name
        })
        .map(|attribute| attribute.value())
}

fn marks_unsupported(parent: Context, element: &SemanticElement<'_>) -> bool {
    if !is_spreadsheetml_element(element, b"") {
        return true;
    }
    match parent {
        Context::Table => element.expanded_name.local_name != "si",
        Context::Item | Context::Run => element.expanded_name.local_name != "t",
        Context::Text | Context::Phonetic | Context::Other => true,
    }
}

fn is_spreadsheetml_element(element: &SemanticElement<'_>, local_name: &[u8]) -> bool {
    is_spreadsheetml_element_name(&element.expanded_name, local_name)
}

fn is_spreadsheetml_element_name(name: &Name, local_name: &[u8]) -> bool {
    (name.namespace.as_bytes() == SPREADSHEETML_NAMESPACE
        || name.namespace.as_bytes() == STRICT_SPREADSHEETML_NAMESPACE)
        && (local_name.is_empty() || name.local_name.as_bytes() == local_name)
}

impl Parser {
    fn parse(content: &str) -> Result<Box<[Text]>> {
        let mut reader = NsReader::from_reader(content.as_bytes());
        let mut parser = Self::default();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| invalid(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"sst") {
                        return Err(invalid(
                            "shared strings XML must have one SpreadsheetML sst root",
                        ));
                    }
                    parser.reserve_hint(&element, decoder)?;
                    stack.push(Context::Table);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"sst") {
                        return Err(invalid(
                            "shared strings XML must have one SpreadsheetML sst root",
                        ));
                    }
                    parser.reserve_hint(&element, decoder)?;
                    closed_root = true;
                },
                Event::Start(element) => {
                    let parent = current(&stack)?;
                    let child = parser.start(parent, &namespace, &element)?;
                    stack.push(child);
                },
                Event::Empty(element) => {
                    let parent = current(&stack)?;
                    let child = parser.start(parent, &namespace, &element)?;
                    parser.finish(child)?;
                },
                Event::Text(value) if stack.last() == Some(&Context::Text) => {
                    parser
                        .push_text(&value.decode().map_err(|error| invalid(error.to_string()))?)?;
                },
                Event::CData(value) if stack.last() == Some(&Context::Text) => {
                    parser
                        .push_text(&value.decode().map_err(|error| invalid(error.to_string()))?)?;
                },
                Event::GeneralRef(value) if stack.last() == Some(&Context::Text) => {
                    parser.push_text(&decode_xml_reference(&value)?)?;
                },
                Event::End(element) => {
                    let ended = stack.pop().ok_or_else(|| {
                        invalid("shared strings XML has a closing element outside its root")
                    })?;
                    parser.finish(ended)?;
                    if ended == Context::Table {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"sst") {
                            return Err(invalid(
                                "shared strings XML has an invalid root closing element",
                            ));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(invalid(
                        "shared strings XML has a missing or unterminated SpreadsheetML sst root",
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
        Ok(parser.strings.into_boxed_slice())
    }

    fn reserve_hint(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let hint = count_hint(element, b"uniqueCount", decoder)?
            .unwrap_or(0)
            .min(MAX_PREALLOCATED_STRINGS);
        count_hint(element, b"count", decoder)?;
        self.strings
            .try_reserve(hint)
            .map_err(|source| allocation("shared-string table", source))
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<Context> {
        if parent == Context::Table && is_spreadsheetml_name(namespace, element.name(), b"si") {
            if self.item.is_some() {
                return Err(invalid("nested shared-string item"));
            }
            self.item = Some(Item::default());
            return Ok(Context::Item);
        }
        if parent == Context::Item && is_spreadsheetml_name(namespace, element.name(), b"t") {
            let item = self
                .item
                .as_mut()
                .ok_or_else(|| invalid("shared-string text outside an item"))?;
            if item.saw_simple || item.saw_run {
                return Err(invalid(
                    "shared-string item mixes or duplicates simple and rich text",
                ));
            }
            item.saw_simple = true;
            return Ok(Context::Text);
        }
        if parent == Context::Item && is_spreadsheetml_name(namespace, element.name(), b"r") {
            let item = self
                .item
                .as_mut()
                .ok_or_else(|| invalid("shared-string run outside an item"))?;
            if item.saw_simple {
                return Err(invalid("shared-string item mixes simple and rich text"));
            }
            item.runs = item
                .runs
                .checked_add(1)
                .filter(|count| *count <= MAX_RUNS)
                .ok_or_else(|| invalid("shared string has too many rich-text runs"))?;
            item.saw_run = true;
            self.seen_text_in_run = false;
            return Ok(Context::Run);
        }
        if parent == Context::Run && is_spreadsheetml_name(namespace, element.name(), b"t") {
            if self.seen_text_in_run {
                return Err(invalid("shared-string run has duplicate text"));
            }
            self.seen_text_in_run = true;
            return Ok(Context::Text);
        }
        if parent == Context::Item && is_spreadsheetml_name(namespace, element.name(), b"rPh") {
            let item = self
                .item
                .as_mut()
                .ok_or_else(|| invalid("phonetic run outside a shared-string item"))?;
            item.phonetic_runs = item
                .phonetic_runs
                .checked_add(1)
                .filter(|count| *count <= MAX_RUNS)
                .ok_or_else(|| invalid("shared string has too many phonetic runs"))?;
            return Ok(Context::Phonetic);
        }
        Ok(Context::Other)
    }

    fn push_text(&mut self, value: &str) -> Result<()> {
        let item = self
            .item
            .as_mut()
            .ok_or_else(|| invalid("shared-string text outside an item"))?;
        item.encoded_bytes = item
            .encoded_bytes
            .checked_add(value.len())
            .filter(|length| *length <= MAX_ENCODED_TEXT_BYTES)
            .ok_or_else(|| invalid("shared-string encoded text is too large"))?;
        item.text
            .try_reserve(value.len())
            .map_err(|source| allocation("shared-string text", source))?;
        item.text.push_str(value);
        Ok(())
    }

    fn finish(&mut self, context: Context) -> Result<()> {
        if context != Context::Item {
            return Ok(());
        }
        let item = self
            .item
            .take()
            .ok_or_else(|| invalid("missing shared-string item"))?;
        let text = decode_spreadsheet_text(&item.text)?;
        if text.chars().count() > MAX_CELL_CHARACTERS {
            return Err(invalid(format!(
                "shared string exceeds {MAX_CELL_CHARACTERS} characters"
            )));
        }
        self.strings
            .try_reserve(1)
            .map_err(|source| allocation("shared-string table", source))?;
        self.strings.push(text.into());
        Ok(())
    }
}

fn current(stack: &[Context]) -> Result<Context> {
    stack
        .last()
        .copied()
        .ok_or_else(|| invalid("shared strings XML is missing its root context"))
}

fn count_hint(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<usize>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| parse_count_hint_value(&value, name))
        .transpose()
}

fn parse_count_hint_value(value: &str, name: &[u8]) -> Result<usize> {
    let count = value.parse::<u64>().map_err(|_source| {
        invalid(format!(
            "invalid shared-string {} '{value}'",
            String::from_utf8_lossy(name)
        ))
    })?;
    if count > MAX_OFFICE_COUNT {
        return Err(invalid(format!(
            "shared-string {} exceeds {MAX_OFFICE_COUNT}",
            String::from_utf8_lossy(name)
        )));
    }
    usize::try_from(count)
        .map_err(|_source| invalid("shared-string count does not fit this platform"))
}

pub(crate) fn decode_spreadsheet_text(value: &str) -> Result<String> {
    let bytes = value.as_bytes();
    let mut decoded = String::new();
    decoded
        .try_reserve(value.len())
        .map_err(|source| allocation("shared-string decoded text", source))?;
    let mut copied_until = 0;
    let mut index = 0;

    while index + 7 <= bytes.len() {
        let Some((unit, end)) = spreadsheet_escape_at(bytes, index) else {
            index += 1;
            continue;
        };
        decoded.push_str(&value[copied_until..index]);
        if (0xD800..=0xDBFF).contains(&unit) {
            let Some((low, pair_end)) = spreadsheet_escape_at(bytes, end) else {
                return Err(invalid(format!(
                    "unpaired high surrogate in SpreadsheetML escape at byte {index}"
                )));
            };
            if !(0xDC00..=0xDFFF).contains(&low) {
                return Err(invalid(format!(
                    "unpaired high surrogate in SpreadsheetML escape at byte {index}"
                )));
            }
            let scalar = 0x1_0000 + ((u32::from(unit) - 0xD800) << 10) + (u32::from(low) - 0xDC00);
            let character = char::from_u32(scalar).ok_or_else(|| {
                invalid(format!(
                    "invalid surrogate pair in SpreadsheetML escape at byte {index}"
                ))
            })?;
            decoded.push(character);
            index = pair_end;
            copied_until = pair_end;
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return Err(invalid(format!(
                "unpaired low surrogate in SpreadsheetML escape at byte {index}"
            )));
        } else {
            let character = char::from_u32(u32::from(unit)).ok_or_else(|| {
                invalid(format!(
                    "invalid code unit in SpreadsheetML escape at byte {index}"
                ))
            })?;
            decoded.push(character);
            index = end;
            copied_until = end;
        }
    }
    decoded.push_str(&value[copied_until..]);
    Ok(decoded)
}

/// Encode XML-illegal control characters and protect literal `_xHHHH_`
/// sequences from `SpreadsheetML`'s second decoding layer.
pub(crate) fn encode_spreadsheet_text(value: &str) -> String {
    let mut encoded = String::with_capacity(value.len());
    for (at, character) in value.char_indices() {
        if character == '_'
            && value
                .as_bytes()
                .get(at..at.saturating_add(7))
                .is_some_and(|bytes| spreadsheet_escape_at(bytes, 0).is_some())
        {
            encoded.push_str("_x005F_");
            continue;
        }
        if matches!(character, '\u{9}' | '\u{A}' | '\u{D}') || character >= '\u{20}' {
            encoded.push(character);
            continue;
        }
        let mut units = [0; 2];
        for unit in character.encode_utf16(&mut units) {
            use std::fmt::Write as _;
            write!(encoded, "_x{unit:04X}_").unwrap_or_else(|error| {
                crate::error::panic_error_invariant("writing to a String is infallible", error)
            });
        }
    }
    encoded
}

fn spreadsheet_escape_at(bytes: &[u8], index: usize) -> Option<(u16, usize)> {
    let escape = bytes.get(index..index.checked_add(7)?)?;
    if escape[0] != b'_' || escape[1] != b'x' || escape[6] != b'_' {
        return None;
    }
    let mut value = 0u16;
    for byte in &escape[2..6] {
        value = value.checked_mul(16)?;
        value = value.checked_add(u16::from(hex(*byte)?))?;
    }
    Some((value, index + 7))
}

fn hex(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use litchi_ooxml_common::mce::{Capabilities, StreamError, StreamLimits};

    use super::*;

    #[allow(clippy::result_large_err)]
    fn streaming_0364_run(
        content: &[u8],
        capabilities: &Capabilities,
        limits: &StreamLimits,
        requested: &[usize],
    ) -> std::result::Result<Selected, StreamError<crate::Error, crate::Error>> {
        let mut input = Cursor::new(content);
        stream_selected(&mut input, capabilities, limits, requested)
    }

    #[allow(clippy::result_large_err)]
    fn streaming_0364_default(
        content: &[u8],
        requested: &[usize],
    ) -> std::result::Result<Selected, StreamError<crate::Error, crate::Error>> {
        streaming_0364_run(
            content,
            &Capabilities::default(),
            &StreamLimits::default(),
            requested,
        )
    }

    #[test]
    fn streaming_0364_selects_plain_requested_indexes() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>first</t></si>
            <si><t>middle</t></si>
            <si><t>last</t></si>
        </sst>"#;

        let first = streaming_0364_default(xml, &[0]).expect("first shared string");
        assert_eq!(first.count, 3);
        assert_eq!(
            first.requested.first().map(|(_, text)| text.as_str()),
            Some("first")
        );
        assert!(!first.unsupported_rich);

        let last = streaming_0364_default(xml, &[2]).expect("last shared string");
        assert_eq!(last.count, 3);
        assert_eq!(
            last.requested.first().map(|(_, text)| text.as_str()),
            Some("last")
        );

        let missing = streaming_0364_default(xml, &[3]).expect("missing shared string");
        assert_eq!(missing.count, 3);
        assert!(missing.requested.is_empty());
    }

    #[test]
    fn streaming_0364_decodes_split_reference_escapes() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>_xD83D&#x5f;_xDE00_</t></si>
            <si><t>A&amp;B&#x21;</t></si>
        </sst>"#;

        let split = streaming_0364_default(xml, &[0]).expect("split SpreadsheetML escape");
        assert_eq!(
            split.requested.first().map(|(_, text)| text.as_str()),
            Some("😀")
        );

        let references = streaming_0364_default(xml, &[1]).expect("XML references");
        assert_eq!(
            references.requested.first().map(|(_, text)| text.as_str()),
            Some("A&B!")
        );
    }

    #[test]
    fn streaming_0364_marks_unsupported_constructs_without_approximation() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="urn:foreign">
            <si><t>safe</t></si>
            <si><r><t>rich</t></r></si>
            <si><rPh><t>phonetic</t></rPh></si>
            <si><extLst><ext/></extLst></si>
            <si><x:foreign/></si>
        </sst>"#;

        for index in 1..5 {
            let selected = streaming_0364_default(xml, &[index]).expect("unsupported item");
            assert_eq!(selected.count, 5);
            assert!(selected.requested.is_empty(), "requested index {index}");
            assert!(selected.unsupported_rich);
        }
        let safe = streaming_0364_default(xml, &[0]).expect("plain item");
        assert_eq!(
            safe.requested.first().map(|(_, text)| text.as_str()),
            Some("safe")
        );
        assert!(safe.unsupported_rich);
    }

    #[test]
    fn streaming_0364_selects_mce_choice_or_fallback() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x="urn:choice">
            <mc:AlternateContent>
                <mc:Choice Requires="x"><si><t>choice</t></si></mc:Choice>
                <mc:Fallback><si><t>fallback</t></si></mc:Fallback>
            </mc:AlternateContent>
        </sst>"#;

        let fallback = streaming_0364_default(xml, &[0]).expect("fallback branch");
        assert_eq!(fallback.count, 1);
        assert_eq!(
            fallback.requested.first().map(|(_, text)| text.as_str()),
            Some("fallback")
        );
        assert!(!fallback.unsupported_rich);

        let mut capabilities = Capabilities::default();
        capabilities.understand_namespace("urn:choice");
        let choice = streaming_0364_run(xml, &capabilities, &StreamLimits::default(), &[0])
            .expect("choice branch");
        assert_eq!(choice.count, 1);
        assert_eq!(
            choice.requested.first().map(|(_, text)| text.as_str()),
            Some("choice")
        );
        assert!(!choice.unsupported_rich);
    }

    #[test]
    fn streaming_0364_drains_malformed_tail_and_rejects_roots() {
        let malformed_tail =
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>requested</t></si><si><t>broken</si></sst>"#;
        assert!(streaming_0364_default(malformed_tail, &[0]).is_err());

        let wrong_root =
            br#"<notSst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        assert!(streaming_0364_default(wrong_root, &[0]).is_err());

        let trailing_root =
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/><tail/>"#;
        assert!(streaming_0364_default(trailing_root, &[]).is_err());
    }

    #[test]
    fn streaming_0364_enforces_exact_stream_and_text_limits() {
        let empty = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        let mut input_exact = StreamLimits::default();
        input_exact.processing.max_input_bytes = empty.len();
        assert!(streaming_0364_run(empty, &Capabilities::default(), &input_exact, &[]).is_ok());
        input_exact.processing.max_input_bytes = empty.len() - 1;
        assert!(streaming_0364_run(empty, &Capabilities::default(), &input_exact, &[]).is_err());

        let event_exact = StreamLimits {
            max_event_bytes: empty.len(),
            ..StreamLimits::default()
        };
        assert!(streaming_0364_run(empty, &Capabilities::default(), &event_exact, &[]).is_ok());
        let event_under = StreamLimits {
            max_event_bytes: empty.len() - 1,
            ..event_exact
        };
        assert!(streaming_0364_run(empty, &Capabilities::default(), &event_under, &[]).is_err());

        let exact_text = "_xD83D__xDE00_".repeat(MAX_CELL_CHARACTERS);
        let exact_xml = format!(
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>{exact_text}</t></si></sst>"#
        );
        assert!(streaming_0364_default(exact_xml.as_bytes(), &[]).is_ok());

        let too_many_chars = "a".repeat(MAX_CELL_CHARACTERS + 1);
        let too_many_xml = format!(
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>{too_many_chars}</t></si></sst>"#
        );
        assert!(streaming_0364_default(too_many_xml.as_bytes(), &[]).is_err());

        let over_encoded = format!("{exact_text}_xD83D__xDE00_");
        let over_encoded_xml = format!(
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>{over_encoded}</t></si></sst>"#
        );
        assert!(streaming_0364_default(over_encoded_xml.as_bytes(), &[]).is_err());
    }

    #[test]
    fn streaming_0364_accepts_advisory_count_mismatches_and_rejects_bad_hints() {
        let mismatch = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="2"><si><t>x</t></si></sst>"#;
        let selected = streaming_0364_default(mismatch, &[]).expect("advisory mismatch");
        assert_eq!(selected.count, 1);

        let invalid = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="NaN"/>"#;
        assert!(streaming_0364_default(invalid, &[]).is_err());

        let excessive = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2147483648"/>"#;
        assert!(streaming_0364_default(excessive, &[]).is_err());
    }

    fn streaming_0365_entries(selected: &Selected) -> Vec<(usize, &str)> {
        selected
            .requested
            .iter()
            .map(|(index, text)| (*index, text.as_str()))
            .collect()
    }

    #[test]
    fn streaming_0365_accepts_empty_request_without_retaining_items() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>first</t></si><si><t>middle</t></si></sst>"#;

        let selected = streaming_0364_default(xml, &[]).expect("empty request");
        assert_eq!(selected.count, 2);
        assert!(selected.requested.is_empty());
        assert!(!selected.unsupported_rich);
    }

    #[test]
    fn streaming_0365_returns_first_middle_last_in_request_order() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>first</t></si><si><t>ignored</t></si><si><t>middle</t></si><si><t>ignored too</t></si><si><t>last</t></si></sst>"#;

        let selected = streaming_0364_default(xml, &[0, 2, 4]).expect("ordered selection");
        assert_eq!(selected.count, 5);
        assert_eq!(
            streaming_0365_entries(&selected),
            vec![(0, "first"), (2, "middle"), (4, "last")]
        );
    }

    #[test]
    fn streaming_0365_omits_missing_and_out_of_range_indexes() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>first</t></si><si><t>middle</t></si><si><t>last</t></si></sst>"#;

        let selected =
            streaming_0364_default(xml, &[0, 3, usize::MAX]).expect("out-of-range selection");
        assert_eq!(selected.count, 3);
        assert_eq!(streaming_0365_entries(&selected), vec![(0, "first")]);
    }

    #[test]
    fn streaming_0365_rejects_unsorted_and_duplicate_requests() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        assert!(streaming_0364_default(xml, &[1, 0]).is_err());
        assert!(streaming_0364_default(xml, &[1, 1]).is_err());
    }

    #[test]
    fn streaming_0365_reports_rich_items_for_requested_and_unrequested_indexes() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>safe</t></si><si><r><t>rich</t></r></si></sst>"#;

        let requested_rich = streaming_0364_default(xml, &[1]).expect("requested rich item");
        assert_eq!(requested_rich.count, 2);
        assert!(requested_rich.requested.is_empty());
        assert!(requested_rich.unsupported_rich);

        let unrequested_rich = streaming_0364_default(xml, &[0]).expect("unrequested rich item");
        assert_eq!(streaming_0365_entries(&unrequested_rich), vec![(0, "safe")]);
        assert!(unrequested_rich.unsupported_rich);
    }

    #[test]
    fn streaming_0365_drains_malformed_tail_after_selected_items() {
        let malformed_tail = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>first</t></si><si><t>middle</t></si><si><t>broken</si></sst>"#;

        assert!(streaming_0364_default(malformed_tail, &[0, 1]).is_err());
    }

    #[test]
    fn streaming_0365_accepts_exact_input_event_and_request_bounds() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si/><si/></sst>"#;
        let limits = StreamLimits {
            processing: litchi_ooxml_common::mce::Limits {
                max_input_bytes: xml.len(),
                ..StreamLimits::default().processing
            },
            max_events: 4,
            ..StreamLimits::default()
        };

        let selected = streaming_0364_run(xml, &Capabilities::default(), &limits, &[0, 1])
            .expect("exact bounded selection");
        assert_eq!(selected.count, 2);
        assert_eq!(selected.requested.len(), 2);
        assert_eq!(streaming_0365_entries(&selected), vec![(0, ""), (1, "")]);
    }

    #[test]
    fn preserves_indexes_and_resolves_plain_rich_and_escaped_text() {
        let xml = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="4">
            <si><t>A &amp; B_x000A__xD83D__xDE00_</t></si>
            <si><t/></si>
            <si><r><rPr><b/></rPr><t>Rich &lt;</t></r><r><t><![CDATA[text>]]></t></r><rPh><t>ignored</t></rPh></si>
            <si/>
        </sst>"#;
        let strings = parse(xml).expect("valid shared strings");
        assert_eq!(strings.len(), 4);
        assert_eq!(strings[0].as_str(), "A & B\n😀");
        assert_eq!(strings[1].as_str(), "");
        assert_eq!(strings[2].as_str(), "Rich <text>");
        assert_eq!(strings[3].as_str(), "");
    }

    #[test]
    fn rejects_mixed_text_and_unpaired_surrogates() {
        let mixed = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>a</t><r><t>b</t></r></si></sst>"#;
        assert!(parse(mixed).is_err());
        assert!(decode_spreadsheet_text("_xD83D_").is_err());
        let excessive = br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="2147483648"/>"#;
        assert!(parse(excessive).is_err());
    }

    #[test]
    fn text_encoding_round_trips_controls_and_literal_escape_syntax() {
        let original = "literal _x0041_\u{1} and 😀";
        let encoded = encode_spreadsheet_text(original);
        assert_eq!(encoded, "literal _x005F_x0041__x0001_ and 😀");
        assert_eq!(
            decode_spreadsheet_text(&encoded).expect("decode encoded text"),
            original
        );
    }
}
