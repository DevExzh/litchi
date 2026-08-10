//! Streaming parser for the workbook shared-string table.

use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::cell::Text;
use crate::error::{Result, allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;

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
        .map(|value| {
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
        })
        .transpose()
}

pub(crate) fn decode_spreadsheet_text(value: &str) -> Result<String> {
    let bytes = value.as_bytes();
    let mut decoded = String::with_capacity(value.len());
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
    use super::*;

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
