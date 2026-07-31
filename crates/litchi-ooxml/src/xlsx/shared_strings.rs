//! Shared-string table support for Excel workbooks.

use std::collections::HashMap;

use litchi_core::sheet::Result as SheetResult;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::RichTextRun;
use super::namespace::is_spreadsheetml_name;
use crate::common::xml::{decode_xml_reference, unqualified_attribute_value};
use crate::error::{OoxmlError, Result};

const MAX_PREALLOCATED_STRINGS: usize = 4096;

/// Shared strings table for efficient string storage.
#[derive(Debug, Default)]
pub struct SharedStrings {
    strings: Vec<String>,
    rich_text: HashMap<usize, Vec<RichTextRun>>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum StringContext {
    SharedStringTable,
    StringItem,
    RichRun,
    RunProperties,
    Text(TextTarget),
    Other,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    Simple,
    RichRun,
}

struct PendingString {
    text: String,
    runs: Vec<RichTextRun>,
    saw_simple_text: bool,
    saw_rich_run: bool,
}

impl PendingString {
    fn new() -> Self {
        Self {
            text: String::new(),
            runs: Vec::new(),
            saw_simple_text: false,
            saw_rich_run: false,
        }
    }
}

struct PendingRun {
    value: RichTextRun,
    saw_text: bool,
    saw_properties: bool,
    seen_properties: u8,
}

impl PendingRun {
    fn new() -> Self {
        Self {
            value: RichTextRun {
                text: String::new(),
                font_name: None,
                font_size: None,
                bold: false,
                italic: false,
                underline: false,
                color: None,
            },
            saw_text: false,
            saw_properties: false,
            seen_properties: 0,
        }
    }
}

struct SharedStringParser {
    strings: Vec<String>,
    rich_text: HashMap<usize, Vec<RichTextRun>>,
    pending_string: Option<PendingString>,
    pending_run: Option<PendingRun>,
    /// `sst/@uniqueCount`, used only to size the string table up front.
    expected_unique_count: Option<u32>,
}

impl SharedStringParser {
    fn new() -> Self {
        Self {
            strings: Vec::new(),
            rich_text: HashMap::new(),
            pending_string: None,
            pending_run: None,
            expected_unique_count: None,
        }
    }

    fn parse(content: &str) -> Result<SharedStrings> {
        let mut reader = NsReader::from_reader(content.as_bytes());
        let mut parser = Self::new();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    if stack.is_empty() {
                        if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"sst")
                        {
                            return Err(invalid(
                                "shared strings XML must have one SpreadsheetML sst root",
                            ));
                        }
                        parser.parse_root_attributes(&element, decoder)?;
                        stack.push(StringContext::SharedStringTable);
                        continue;
                    }
                    let parent = current_context(&stack)?;
                    stack.push(parser.start_element(parent, &namespace, &element, decoder)?);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"sst") {
                        return Err(invalid(
                            "shared strings XML must have one SpreadsheetML sst root",
                        ));
                    }
                    parser.parse_root_attributes(&element, decoder)?;
                    closed_root = true;
                },
                Event::Empty(element) => {
                    let parent = current_context(&stack)?;
                    parser.empty_element(parent, &namespace, &element, decoder)?;
                },
                Event::Text(text) => {
                    if let Some(target) = current_text_target(&stack) {
                        let value = text
                            .decode()
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        parser.push_text(target, &value)?;
                    }
                },
                Event::CData(text) => {
                    if let Some(target) = current_text_target(&stack) {
                        let value = text
                            .decode()
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        parser.push_text(target, &value)?;
                    }
                },
                Event::GeneralRef(reference) => {
                    if let Some(target) = current_text_target(&stack) {
                        parser.push_text(target, &decode_xml_reference(&reference)?)?;
                    }
                },
                Event::End(element) => {
                    let context = stack.pop().ok_or_else(|| {
                        invalid("shared strings XML has a closing element outside its root")
                    })?;
                    parser.finish_context(context)?;
                    if context == StringContext::SharedStringTable {
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
                _ => {},
            }
        }

        Ok(SharedStrings {
            strings: parser.strings,
            rich_text: parser.rich_text,
        })
    }

    /// Read the advisory `count` and `uniqueCount` hints on the `sst` root.
    ///
    /// ECMA-376 declares both as optional, and Excel writes tables whose hints
    /// disagree with the `si` children actually present. The parsed items are
    /// authoritative, so a missing, unparseable, or contradictory hint is
    /// ignored rather than failing the workbook; `uniqueCount` only ever sizes
    /// the initial allocation, itself capped at `MAX_PREALLOCATED_STRINGS`.
    fn parse_root_attributes(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        self.expected_unique_count = optional_advisory_u32(element, b"uniqueCount", decoder)?;
        if let Some(expected) = self.expected_unique_count {
            let capacity = usize::try_from(expected)
                .unwrap_or(usize::MAX)
                .min(MAX_PREALLOCATED_STRINGS);
            self.strings.reserve(capacity);
            self.rich_text.reserve(capacity / 4);
        }
        Ok(())
    }

    fn start_element(
        &mut self,
        parent: StringContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<StringContext> {
        if parent == StringContext::SharedStringTable
            && is_spreadsheetml_name(namespace, element.name(), b"si")
        {
            self.start_string()?;
            return Ok(StringContext::StringItem);
        }
        if parent == StringContext::StringItem
            && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            self.start_text(TextTarget::Simple)?;
            return Ok(StringContext::Text(TextTarget::Simple));
        }
        if parent == StringContext::StringItem
            && is_spreadsheetml_name(namespace, element.name(), b"r")
        {
            self.start_run()?;
            return Ok(StringContext::RichRun);
        }
        if parent == StringContext::RichRun
            && is_spreadsheetml_name(namespace, element.name(), b"rPr")
        {
            self.start_run_properties()?;
            return Ok(StringContext::RunProperties);
        }
        if parent == StringContext::RichRun
            && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            self.start_text(TextTarget::RichRun)?;
            return Ok(StringContext::Text(TextTarget::RichRun));
        }
        if parent == StringContext::RunProperties {
            self.parse_run_property(namespace, element, decoder)?;
        }
        Ok(StringContext::Other)
    }

    fn empty_element(
        &mut self,
        parent: StringContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        if parent == StringContext::SharedStringTable
            && is_spreadsheetml_name(namespace, element.name(), b"si")
        {
            self.start_string()?;
            self.finish_string()?;
        } else if parent == StringContext::StringItem
            && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            self.start_text(TextTarget::Simple)?;
        } else if parent == StringContext::StringItem
            && is_spreadsheetml_name(namespace, element.name(), b"r")
        {
            return Err(invalid("shared-string rich-text run is missing its text"));
        } else if parent == StringContext::RichRun
            && is_spreadsheetml_name(namespace, element.name(), b"rPr")
        {
            self.start_run_properties()?;
        } else if parent == StringContext::RichRun
            && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            self.start_text(TextTarget::RichRun)?;
        } else if parent == StringContext::RunProperties {
            self.parse_run_property(namespace, element, decoder)?;
        }
        Ok(())
    }

    fn start_string(&mut self) -> Result<()> {
        if self.pending_string.is_some() {
            return Err(invalid("nested shared-string item"));
        }
        self.pending_string = Some(PendingString::new());
        Ok(())
    }

    fn start_run(&mut self) -> Result<()> {
        let string = self
            .pending_string
            .as_mut()
            .ok_or_else(|| invalid("rich-text run outside a shared-string item"))?;
        if string.saw_simple_text {
            return Err(invalid(
                "shared-string item mixes simple text and rich-text runs",
            ));
        }
        string.saw_rich_run = true;
        if self.pending_run.is_some() {
            return Err(invalid("nested shared-string rich-text run"));
        }
        self.pending_run = Some(PendingRun::new());
        Ok(())
    }

    fn start_run_properties(&mut self) -> Result<()> {
        let run = self
            .pending_run
            .as_mut()
            .ok_or_else(|| invalid("run properties outside a shared-string rich-text run"))?;
        if run.saw_properties {
            return Err(invalid("duplicate shared-string run properties"));
        }
        run.saw_properties = true;
        Ok(())
    }

    fn start_text(&mut self, target: TextTarget) -> Result<()> {
        match target {
            TextTarget::Simple => {
                let string = self
                    .pending_string
                    .as_mut()
                    .ok_or_else(|| invalid("text outside a shared-string item"))?;
                if string.saw_rich_run {
                    return Err(invalid(
                        "shared-string item mixes simple text and rich-text runs",
                    ));
                }
                if string.saw_simple_text {
                    return Err(invalid("duplicate text in shared-string item"));
                }
                string.saw_simple_text = true;
            },
            TextTarget::RichRun => {
                let run = self
                    .pending_run
                    .as_mut()
                    .ok_or_else(|| invalid("text outside a shared-string rich-text run"))?;
                if run.saw_text {
                    return Err(invalid("duplicate text in shared-string rich-text run"));
                }
                run.saw_text = true;
            },
        }
        Ok(())
    }

    fn push_text(&mut self, target: TextTarget, value: &str) -> Result<()> {
        match target {
            TextTarget::Simple => self
                .pending_string
                .as_mut()
                .ok_or_else(|| invalid("text outside a shared-string item"))?
                .text
                .push_str(value),
            TextTarget::RichRun => self
                .pending_run
                .as_mut()
                .ok_or_else(|| invalid("text outside a shared-string rich-text run"))?
                .value
                .text
                .push_str(value),
        }
        Ok(())
    }

    fn parse_run_property(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let run = self
            .pending_run
            .as_mut()
            .ok_or_else(|| invalid("run property outside a shared-string rich-text run"))?;
        if is_spreadsheetml_name(namespace, element.name(), b"rFont") {
            mark_property(&mut run.seen_properties, 1, "font name")?;
            let value = required_string(element, b"val", decoder, "rich-text font name")?;
            run.value.font_name = Some(value);
        } else if is_spreadsheetml_name(namespace, element.name(), b"sz") {
            mark_property(&mut run.seen_properties, 2, "font size")?;
            let value = required_string(element, b"val", decoder, "rich-text font size")?;
            let size = value
                .parse::<f64>()
                .map_err(|_| invalid(format!("invalid rich-text font size '{value}'")))?;
            if !size.is_finite() || size <= 0.0 {
                return Err(invalid(format!("invalid rich-text font size '{value}'")));
            }
            run.value.font_size = Some(size);
        } else if is_spreadsheetml_name(namespace, element.name(), b"b") {
            mark_property(&mut run.seen_properties, 4, "bold property")?;
            run.value.bold = boolean_property(element, decoder, "rich-text bold")?;
        } else if is_spreadsheetml_name(namespace, element.name(), b"i") {
            mark_property(&mut run.seen_properties, 8, "italic property")?;
            run.value.italic = boolean_property(element, decoder, "rich-text italic")?;
        } else if is_spreadsheetml_name(namespace, element.name(), b"u") {
            mark_property(&mut run.seen_properties, 16, "underline property")?;
            let value = unqualified_attribute_value(element, b"val", decoder)?
                .unwrap_or_else(|| "single".to_string());
            run.value.underline = match value.as_str() {
                "none" => false,
                "single" | "double" | "singleAccounting" | "doubleAccounting" => true,
                _ => {
                    return Err(invalid(format!(
                        "invalid rich-text underline value '{value}'"
                    )));
                },
            };
        } else if is_spreadsheetml_name(namespace, element.name(), b"color") {
            mark_property(&mut run.seen_properties, 32, "color property")?;
            if let Some(rgb) = unqualified_attribute_value(element, b"rgb", decoder)? {
                if !matches!(rgb.len(), 6 | 8) || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit())
                {
                    return Err(invalid(format!("invalid rich-text RGB color '{rgb}'")));
                }
                run.value.color = Some(rgb);
            }
        }
        Ok(())
    }

    fn finish_context(&mut self, context: StringContext) -> Result<()> {
        match context {
            StringContext::RichRun => self.finish_run(),
            StringContext::StringItem => self.finish_string(),
            _ => Ok(()),
        }
    }

    fn finish_run(&mut self) -> Result<()> {
        let mut run = self
            .pending_run
            .take()
            .ok_or_else(|| invalid("missing shared-string rich-text run"))?;
        if !run.saw_text {
            return Err(invalid("shared-string rich-text run is missing its text"));
        }
        run.value.text = decode_spreadsheet_text(&run.value.text)?;
        let string = self
            .pending_string
            .as_mut()
            .ok_or_else(|| invalid("rich-text run outside a shared-string item"))?;
        string.text.push_str(&run.value.text);
        string.runs.push(run.value);
        Ok(())
    }

    fn finish_string(&mut self) -> Result<()> {
        if self.pending_run.is_some() {
            return Err(invalid("unterminated shared-string rich-text run"));
        }
        let mut string = self
            .pending_string
            .take()
            .ok_or_else(|| invalid("missing shared-string item"))?;
        if !string.saw_rich_run {
            string.text = decode_spreadsheet_text(&string.text)?;
        }
        let index = self.strings.len();
        self.strings.push(string.text);
        if !string.runs.is_empty() {
            self.rich_text.insert(index, string.runs);
        }
        Ok(())
    }
}

impl SharedStrings {
    /// Create a new empty shared strings table.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse shared strings from `xl/sharedStrings.xml`.
    pub fn parse(content: &str) -> SheetResult<Self> {
        let content = litchi_ooxml_common::mce::process_str(content)
            .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)?;
        SharedStringParser::parse(content.as_ref())
            .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
    }

    /// Get a string by its index.
    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(String::as_str)
    }

    /// Get the number of strings in the table.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Check if the table is empty.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Get all strings.
    pub fn strings(&self) -> &[String] {
        &self.strings
    }

    /// Get rich text runs for a specific shared string index, if present.
    pub fn rich_text_runs(&self, index: usize) -> Option<&[RichTextRun]> {
        self.rich_text.get(&index).map(Vec::as_slice)
    }
}

fn current_context(stack: &[StringContext]) -> Result<StringContext> {
    stack
        .last()
        .copied()
        .ok_or_else(|| invalid("shared strings XML is missing its root context"))
}

fn current_text_target(stack: &[StringContext]) -> Option<TextTarget> {
    match stack.last() {
        Some(StringContext::Text(target)) => Some(*target),
        _ => None,
    }
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

/// Read an optional attribute that is only a hint, ignoring unusable values.
///
/// Returns `None` when the attribute is absent or does not parse as a `u32`,
/// so a malformed hint never fails the surrounding parse.
fn optional_advisory_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<u32>> {
    Ok(unqualified_attribute_value(element, name, decoder)?
        .and_then(|value| value.parse::<u32>().ok()))
}

fn boolean_property(element: &BytesStart<'_>, decoder: Decoder, description: &str) -> Result<bool> {
    match unqualified_attribute_value(element, b"val", decoder)?.as_deref() {
        None | Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        Some(value) => Err(invalid(format!("invalid {description} value '{value}'"))),
    }
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
            decoded.push(char::from_u32(scalar).ok_or_else(|| {
                invalid(format!(
                    "invalid surrogate pair in SpreadsheetML escape at byte {index}"
                ))
            })?);
            index = pair_end;
            copied_until = pair_end;
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return Err(invalid(format!(
                "unpaired low surrogate in SpreadsheetML escape at byte {index}"
            )));
        } else {
            decoded.push(char::from_u32(u32::from(unit)).ok_or_else(|| {
                invalid(format!(
                    "invalid code unit in SpreadsheetML escape at byte {index}"
                ))
            })?);
            index = end;
            copied_until = end;
        }
    }
    decoded.push_str(&value[copied_until..]);
    Ok(decoded)
}

fn spreadsheet_escape_at(bytes: &[u8], index: usize) -> Option<(u16, usize)> {
    let escape = bytes.get(index..index.checked_add(7)?)?;
    if escape[0] != b'_' || escape[1] != b'x' || escape[6] != b'_' {
        return None;
    }
    let mut value = 0u16;
    for byte in &escape[2..6] {
        value = value.checked_mul(16)?;
        value = value.checked_add(u16::from(hex_value(*byte)?))?;
    }
    Some((value, index + 7))
}

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn mark_property(seen: &mut u8, bit: u8, description: &str) -> Result<()> {
    if *seen & bit != 0 {
        return Err(invalid(format!("duplicate shared-string {description}")));
    }
    *seen |= bit;
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    #[test]
    fn preserves_indexes_and_decodes_plain_and_rich_text() {
        let xml = format!(
            r#"<x:sst xmlns:x="{S}" count="5" uniqueCount="4">
                <x:si><x:t>A &amp; B_x000A__xD83D__xDE00__x005F_x0041_</x:t></x:si>
                <x:si><x:t/></x:si>
                <x:si><x:r><x:rPr><x:rFont val="A &amp; B"/><x:sz val="11.5"/>
                    <x:b val="0"/><x:i/><x:u val="double"/><x:color rgb="FF112233"/></x:rPr>
                    <x:t>Rich &lt;</x:t></x:r><x:r><x:t><![CDATA[text>]]></x:t></x:r>
                    <x:rPh sb="0" eb="1"><x:t>phonetic</x:t></x:rPh></x:si>
                <x:si/>
            </x:sst>"#
        );
        let strings = SharedStrings::parse(&xml).unwrap();
        assert_eq!(
            strings.strings(),
            &["A & B\n😀_x0041_", "", "Rich <text>", ""]
        );
        assert!(strings.rich_text_runs(0).is_none());
        let runs = strings.rich_text_runs(2).unwrap();
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Rich <");
        assert_eq!(runs[0].font_name.as_deref(), Some("A & B"));
        assert_eq!(runs[0].font_size, Some(11.5));
        assert!(!runs[0].bold);
        assert!(runs[0].italic);
        assert!(runs[0].underline);
        assert_eq!(runs[0].color.as_deref(), Some("FF112233"));
        assert_eq!(runs[1].text, "text>");
    }

    #[test]
    fn accepts_strict_namespaces_and_ignores_foreign_lookalikes() {
        let xml = format!(
            r#"<sst xmlns="{STRICT_S}" xmlns:f="urn:foreign" uniqueCount="2">
                <f:si><si><t>Nested</t></si></f:si>
                <si><f:t>Ignored</f:t><t>Strict</t></si><si><r><f:t>Ignored</f:t><t>Run</t></r></si>
            </sst>"#
        );
        let strings = SharedStrings::parse(&xml).unwrap();
        assert_eq!(strings.strings(), &["Strict", "Run"]);
        assert_eq!(strings.rich_text_runs(1).unwrap()[0].text, "Run");
    }

    #[test]
    fn rejects_bad_counts_structure_and_run_properties() {
        for xml in [
            format!(r#"<sst xmlns="{S}"><si><t>one</t><r><t>two</t></r></si></sst>"#),
            format!(r#"<sst xmlns="{S}"><si><r><rPr><sz val="NaN"/></rPr><t>x</t></r></si></sst>"#),
            format!(r#"<sst xmlns="{S}"><si><r><rPr><b/><b/></rPr><t>x</t></r></si></sst>"#),
            format!(r#"<sst xmlns="{S}"><si><t>bad_xD800_</t></si></sst>"#),
            format!(r#"<sst xmlns="{S}"><si><r><rPr/><t>x</t>"#),
        ] {
            assert!(SharedStrings::parse(&xml).is_err(), "accepted {xml}");
        }
    }

    /// `sst/@count` and `sst/@uniqueCount` are optional hints, and Excel emits
    /// tables whose hints disagree with the `si` children present. The parsed
    /// items stay authoritative instead of the workbook being rejected.
    #[test]
    fn treats_contradictory_and_malformed_count_hints_as_advisory() {
        for xml in [
            // uniqueCount too high, too low, and count below the item total.
            format!(r#"<sst xmlns="{S}" uniqueCount="2"><si><t>one</t></si></sst>"#),
            format!(r#"<sst xmlns="{S}" uniqueCount="0"><si><t>one</t></si></sst>"#),
            format!(r#"<sst xmlns="{S}" count="0"><si><t>one</t></si></sst>"#),
            // Non-numeric and negative hints are ignored rather than fatal.
            format!(r#"<sst xmlns="{S}" uniqueCount="NaN"><si><t>one</t></si></sst>"#),
            format!(r#"<sst xmlns="{S}" count="-1"><si><t>one</t></si></sst>"#),
        ] {
            let parsed = SharedStrings::parse(&xml).unwrap_or_else(|e| panic!("{xml}: {e}"));
            assert_eq!(parsed.strings(), ["one"], "{xml}");
        }
    }

    /// An absurd `uniqueCount` must not drive the initial allocation.
    #[test]
    fn oversized_unique_count_hint_does_not_preallocate() {
        let xml = format!(r#"<sst xmlns="{S}" uniqueCount="4294967295"><si><t>one</t></si></sst>"#);
        let parsed = SharedStrings::parse(&xml).unwrap();
        assert_eq!(parsed.strings(), ["one"]);
    }
}
