//! Bounded `SpreadsheetML` ignored-error parsing.

use super::model::{
    IgnoredError, IgnoredErrorRangeReference, IgnoredErrorType, IgnoredErrors,
    IgnoredErrorsExtension,
};
use crate::error::{Error, Result};
use crate::raw::namespace::is_spreadsheetml_name;
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use quick_xml::Writer;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_IGNORED_ERRORS: usize = 9;
const MAX_SQREF_ITEMS: usize = 32_767;
const MAX_SQREF_BYTES: usize = 1024 * 1024;
const MAX_EXTENSIONS: usize = 1024;
const MAX_EXTENSION_BYTES: usize = 1024 * 1024;
const MAX_RETAINED_EXTENSION_BYTES: usize = 4 * 1024 * 1024;
const MAX_EXTENSION_URI_BYTES: usize = 1024;
const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Outside,
    Worksheet,
    Collection,
    IgnoredError,
    ExtensionList,
    Extension,
}

struct Capture {
    depth: usize,
    writer: Writer<Vec<u8>>,
    extension: IgnoredErrorsExtension,
}

struct Parser {
    stack: Vec<Context>,
    collection: Option<IgnoredErrors>,
    capture: Option<Capture>,
    seen_collection: bool,
    seen_extension_list: bool,
    extension_list_start: usize,
    collection_phase: u8,
    retained_extension_bytes: usize,
    root_seen: bool,
    root_closed: bool,
}

/// Parse the worksheet's direct `ignoredErrors` collection.
pub fn parse_worksheet_ignored_errors(xml: &[u8]) -> Result<Option<IgnoredErrors>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("ignoredErrors worksheet XML exceeds size limit"));
    }
    let processed =
        process_markup_compatibility(xml, &Capabilities::default(), &Limits::default())?;
    if processed.xml.len() > MAX_XML_BYTES {
        return Err(invalid(
            "processed ignoredErrors worksheet XML exceeds size limit",
        ));
    }
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut parser = Parser {
        stack: Vec::new(),
        collection: None,
        capture: None,
        seen_collection: false,
        seen_extension_list: false,
        extension_list_start: 0,
        collection_phase: 0,
        retained_extension_bytes: 0,
        root_seen: false,
        root_closed: false,
    };
    let mut events = 0usize;
    let mut declaration_seen = false;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("ignoredErrors XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("ignoredErrors XML exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        if parser.capture.is_some() {
            parser.capture_event(event)?;
            continue;
        }
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => parser.start(&namespace, &element, decoder, &resolver)?,
            Event::Empty(element) => parser.empty(&namespace, &element, decoder, &resolver)?,
            Event::End(element) => parser.end(element.local_name().as_ref())?,
            Event::Text(text)
                if matches!(
                    parser.parent(),
                    Context::Collection | Context::IgnoredError | Context::ExtensionList
                ) && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("unexpected text in ignoredErrors"));
            },
            Event::CData(_)
                if matches!(
                    parser.parent(),
                    Context::Collection | Context::IgnoredError | Context::ExtensionList
                ) =>
            {
                return Err(invalid("unexpected CDATA in ignoredErrors"));
            },
            Event::Text(text)
                if matches!(parser.parent(), Context::Worksheet)
                    && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid(
                    "worksheet cannot contain direct ignoredErrors text",
                ));
            },
            Event::CData(_) if matches!(parser.parent(), Context::Worksheet) => {
                return Err(invalid(
                    "worksheet cannot contain direct ignoredErrors CDATA",
                ));
            },
            Event::Text(text)
                if parser.stack.is_empty()
                    && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("ignoredErrors XML text is outside root"));
            },
            Event::CData(_) if parser.stack.is_empty() => {
                return Err(invalid("ignoredErrors XML CDATA is outside root"));
            },
            Event::Decl(_) => {
                if parser.root_seen || declaration_seen {
                    return Err(invalid("invalid ignoredErrors XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if parser.capture.is_some()
        || !parser.stack.is_empty()
        || !parser.root_seen
        || !parser.root_closed
    {
        return Err(invalid("unterminated ignoredErrors XML"));
    }
    Ok(parser.collection)
}

impl Parser {
    fn parent(&self) -> Context {
        self.stack.last().copied().unwrap_or(Context::Outside)
    }

    fn start(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(namespace, element.name(), local.as_ref());
        if self.stack.is_empty() {
            if self.root_closed || self.root_seen {
                return Err(invalid("ignoredErrors XML contains multiple roots"));
            }
            if !core || local.as_ref() != b"worksheet" {
                return Err(invalid("ignoredErrors parser requires a worksheet root"));
            }
            self.root_seen = true;
        }
        if self.stack.len() >= MAX_DEPTH {
            return Err(invalid("ignoredErrors XML nesting is too deep"));
        }
        match (self.parent(), core, local.as_ref()) {
            (Context::Outside, true, b"worksheet") => self.stack.push(Context::Worksheet),
            (Context::Worksheet, true, b"ignoredErrors") => {
                self.begin_collection(element)?;
                self.stack.push(Context::Collection);
            },
            (Context::Collection, true, b"ignoredError") => {
                self.add_ignored_error(element, decoder, resolver)?;
                self.stack.push(Context::IgnoredError);
            },
            (Context::Collection, true, b"extLst") => {
                self.begin_extension_list(element)?;
                self.stack.push(Context::ExtensionList);
            },
            (Context::ExtensionList, true, b"ext") => {
                let extension = parse_extension(element, decoder, resolver)?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element.clone()))
                    .map_err(xml_error)?;
                self.capture = Some(Capture {
                    depth: 1,
                    writer,
                    extension,
                });
                self.stack.push(Context::Extension);
            },
            (Context::IgnoredError, _, _) => return Err(invalid("ignoredError is a leaf element")),
            (Context::Collection | Context::ExtensionList, _, _) => {
                return Err(invalid(format!(
                    "unexpected ignoredErrors element '{}'",
                    String::from_utf8_lossy(local.as_ref()),
                )));
            },
            _ => self.stack.push(Context::Outside),
        }
        Ok(())
    }

    fn empty(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(namespace, element.name(), local.as_ref());
        if self.stack.is_empty() {
            return Err(if self.root_seen || self.root_closed {
                invalid("ignoredErrors XML contains multiple roots")
            } else {
                invalid("worksheet root cannot be empty")
            });
        }
        match (self.parent(), core, local.as_ref()) {
            (Context::Worksheet, true, b"ignoredErrors") => {
                self.begin_collection(element)?;
                return Err(invalid("ignoredErrors requires at least one ignoredError"));
            },
            (Context::Collection, true, b"ignoredError") => {
                self.add_ignored_error(element, decoder, resolver)?;
            },
            (Context::Collection, true, b"extLst") => {
                self.begin_extension_list(element)?;
                return Err(invalid("ignoredErrors extLst requires at least one ext"));
            },
            (Context::ExtensionList, true, b"ext") => {
                let mut extension = parse_extension(element, decoder, resolver)?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element.clone()))
                    .map_err(xml_error)?;
                extension.markup = writer.into_inner();
                self.add_extension(extension)?;
            },
            (Context::IgnoredError, _, _) => return Err(invalid("ignoredError is a leaf element")),
            (Context::Collection | Context::ExtensionList, _, _) => {
                return Err(invalid(format!(
                    "unexpected ignoredErrors element '{}'",
                    String::from_utf8_lossy(local.as_ref()),
                )));
            },
            _ => {},
        }
        Ok(())
    }

    fn end(&mut self, local: &[u8]) -> Result<()> {
        let context = self
            .stack
            .pop()
            .ok_or_else(|| invalid("unexpected ignoredErrors end element"))?;
        match context {
            Context::Collection if local == b"ignoredErrors" => self.finish_collection(),
            Context::ExtensionList if local == b"extLst" => self.finish_extension_list(),
            Context::IgnoredError if local == b"ignoredError" => Ok(()),
            Context::Worksheet if local == b"worksheet" => {
                self.root_closed = true;
                Ok(())
            },
            Context::Outside => Ok(()),
            Context::Worksheet
            | Context::Collection
            | Context::IgnoredError
            | Context::ExtensionList
            | Context::Extension => Err(invalid("mismatched ignoredErrors end element")),
        }
    }

    fn begin_collection(&mut self, element: &BytesStart<'_>) -> Result<()> {
        if self.seen_collection {
            return Err(invalid("duplicate worksheet ignoredErrors element"));
        }
        reject_attributes(element, "ignoredErrors")?;
        self.seen_collection = true;
        self.collection_phase = 0;
        self.collection = Some(IgnoredErrors {
            entries: Vec::new(),
            extensions: Vec::new(),
        });
        Ok(())
    }

    fn finish_collection(&self) -> Result<()> {
        if self
            .collection
            .as_ref()
            .is_none_or(|value| value.entries.is_empty())
        {
            return Err(invalid("ignoredErrors requires at least one ignoredError"));
        }
        Ok(())
    }

    fn add_ignored_error(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.collection_phase != 0 {
            return Err(invalid("ignoredError appears after ignoredErrors extLst"));
        }
        let collection = self
            .collection
            .as_mut()
            .ok_or_else(|| invalid("ignoredError outside ignoredErrors"))?;
        if collection.entries.len() >= MAX_IGNORED_ERRORS {
            return Err(invalid("Excel permits at most 9 ignoredError entries"));
        }
        collection
            .entries
            .push(parse_ignored_error(element, decoder, resolver)?);
        Ok(())
    }

    fn begin_extension_list(&mut self, element: &BytesStart<'_>) -> Result<()> {
        if self
            .collection
            .as_ref()
            .is_none_or(|value| value.entries.is_empty())
        {
            return Err(invalid("ignoredErrors extLst appears before ignoredError"));
        }
        if self.seen_extension_list {
            return Err(invalid("duplicate ignoredErrors extLst"));
        }
        reject_attributes(element, "extLst")?;
        self.seen_extension_list = true;
        self.collection_phase = 1;
        self.extension_list_start = self
            .collection
            .as_ref()
            .map_or(0, |value| value.extensions.len());
        Ok(())
    }

    fn finish_extension_list(&self) -> Result<()> {
        let count = self
            .collection
            .as_ref()
            .map_or(0, |value| value.extensions.len());
        if count == self.extension_list_start {
            return Err(invalid("ignoredErrors extLst requires at least one ext"));
        }
        Ok(())
    }

    fn capture_event(&mut self, event: Event<'static>) -> Result<()> {
        let capture = self
            .capture
            .as_mut()
            .ok_or_else(|| invalid("missing extension capture"))?;
        match &event {
            Event::Start(_) => {
                capture.depth = capture
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension capture depth overflow"))?;
                if capture.depth > MAX_DEPTH {
                    return Err(invalid("ignoredErrors extension nesting is too deep"));
                }
            },
            Event::End(_) => {
                capture.depth = capture
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("extension capture depth underflow"))?;
            },
            Event::Eof => return Err(invalid("unterminated ignoredErrors extension")),
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        capture.writer.write_event(event).map_err(xml_error)?;
        if capture.writer.get_ref().len() > MAX_EXTENSION_BYTES {
            return Err(invalid("ignoredErrors extension exceeds size limit"));
        }
        if capture.depth == 0 {
            let mut capture = self
                .capture
                .take()
                .ok_or_else(|| invalid("missing extension capture"))?;
            capture.extension.markup = capture.writer.into_inner();
            let context = self
                .stack
                .pop()
                .ok_or_else(|| invalid("missing extension context"))?;
            if context != Context::Extension {
                return Err(invalid("mismatched extension context"));
            }
            self.add_extension(capture.extension)?;
        }
        Ok(())
    }

    fn add_extension(&mut self, extension: IgnoredErrorsExtension) -> Result<()> {
        let collection = self
            .collection
            .as_mut()
            .ok_or_else(|| invalid("ext outside ignoredErrors"))?;
        if collection.extensions.len() >= MAX_EXTENSIONS {
            return Err(invalid("too many ignoredErrors extensions"));
        }
        self.retained_extension_bytes = self
            .retained_extension_bytes
            .checked_add(extension.markup.len())
            .ok_or_else(|| invalid("extension size overflow"))?;
        if self.retained_extension_bytes > MAX_RETAINED_EXTENSION_BYTES {
            return Err(invalid(
                "ignoredErrors retained extensions exceed size limit",
            ));
        }
        collection.extensions.push(extension);
        Ok(())
    }
}

fn parse_ignored_error(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<IgnoredError> {
    let mut sqref = None;
    let mut flags = [false; 9];
    let mut seen_flags = [false; 9];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!(
                "unknown namespaced ignoredError attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        if local.as_ref() == b"sqref" {
            if sqref.is_some() {
                return Err(invalid("duplicate ignoredError sqref attribute"));
            }
            sqref = Some(parse_sqref(&value)?);
            continue;
        }
        let error_type = match local.as_ref() {
            b"calculatedColumn" => IgnoredErrorType::CalculatedColumn,
            b"emptyCellReference" => IgnoredErrorType::EmptyCellReference,
            b"evalError" => IgnoredErrorType::EvaluationError,
            b"formula" => IgnoredErrorType::Formula,
            b"formulaRange" => IgnoredErrorType::FormulaRange,
            b"listDataValidation" => IgnoredErrorType::ListDataValidation,
            b"numberStoredAsText" => IgnoredErrorType::NumberStoredAsText,
            b"twoDigitTextYear" => IgnoredErrorType::TwoDigitTextYear,
            b"unlockedFormula" => IgnoredErrorType::UnlockedFormula,
            name => {
                return Err(invalid(format!(
                    "unknown ignoredError attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        let index = error_type as usize;
        if seen_flags[index] {
            return Err(invalid("duplicate ignoredError flag attribute"));
        }
        seen_flags[index] = true;
        flags[index] = parse_bool(&value, String::from_utf8_lossy(local.as_ref()).as_ref())?;
    }
    let ranges = sqref.ok_or_else(|| invalid("ignoredError requires sqref"))?;
    Ok(IgnoredError { ranges, flags })
}

fn parse_sqref(value: &str) -> Result<Vec<IgnoredErrorRangeReference>> {
    if value.len() > MAX_SQREF_BYTES {
        return Err(invalid("ignoredError sqref exceeds size limit"));
    }
    let mut ranges = Vec::new();
    for token in value.split_whitespace() {
        if ranges.len() >= MAX_SQREF_ITEMS {
            return Err(invalid("too many ignoredError sqref items"));
        }
        validate_range_reference(token)?;
        ranges.push(IgnoredErrorRangeReference(token.to_string()));
    }
    if ranges.is_empty() {
        return Err(invalid("ignoredError sqref cannot be empty"));
    }
    Ok(ranges)
}

fn validate_range_reference(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if parts.next().is_some() || first.is_empty() || second.is_some_and(str::is_empty) {
        return Err(invalid(format!("invalid ignoredError reference '{value}'")));
    }
    validate_cell_reference(first)?;
    if let Some(second) = second {
        validate_cell_reference(second)?;
    }
    Ok(())
}

fn validate_cell_reference(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    if index == column_start || index - column_start > 3 {
        return Err(invalid(format!(
            "invalid ignoredError cell reference '{value}'"
        )));
    }
    let mut column = 0u32;
    for byte in &bytes[column_start..index] {
        column = column * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    if column == 0 || column > MAX_COLUMN {
        return Err(invalid(format!(
            "ignoredError column is out of range in '{value}'"
        )));
    }
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == row_start || index != bytes.len() {
        return Err(invalid(format!(
            "invalid ignoredError cell reference '{value}'"
        )));
    }
    let row = value[row_start..]
        .parse::<u32>()
        .map_err(|_source| invalid(format!("invalid ignoredError row in '{value}'")))?;
    if row == 0 || row > MAX_ROW {
        return Err(invalid(format!(
            "ignoredError row is out of range in '{value}'"
        )));
    }
    Ok(())
}

fn parse_extension(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<IgnoredErrorsExtension> {
    let mut uri = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) || local.as_ref() != b"uri" {
            return Err(invalid(format!(
                "unknown ignoredErrors ext attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
        if uri.is_some() {
            return Err(invalid("duplicate ignoredErrors ext uri"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if value.is_empty() || value.len() > MAX_EXTENSION_URI_BYTES {
            return Err(invalid("ignoredErrors ext uri is empty or too long"));
        }
        uri = Some(value);
    }
    Ok(IgnoredErrorsExtension {
        uri: uri.ok_or_else(|| invalid("ignoredErrors ext requires uri"))?,
        markup: Vec::new(),
    })
}

fn reject_attributes(element: &BytesStart<'_>, name: &str) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err(invalid(format!(
                "unexpected {name} attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!(
            "invalid ignoredError {name} boolean '{value}'"
        ))),
    }
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_) | Event::PI(_)) {
        return Err(invalid("DTD and processing instructions are rejected"));
    }
    if let Event::GeneralRef(reference) = event {
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && !name.starts_with('#')
        {
            return Err(invalid("custom XML entities are rejected"));
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid worksheet ignoredErrors XML: {error}"
    )))
}
