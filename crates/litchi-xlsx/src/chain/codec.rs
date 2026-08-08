//! SpreadsheetML calculation-chain XML codec.

use std::borrow::Cow;
use std::collections::HashSet;

use crate::error::{Result, allocation, invalid};
use litchi_sheet::Cell as Address;
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{
    Cell, Chain, Conformance, Flags, MAX_ATTRIBUTE_BYTES, MAX_CELL_CONTENT_BYTES, MAX_CELLS,
    MAX_EXTENSION_ATTRIBUTES, MAX_EXTENSION_BYTES, MAX_EXTENSION_DEPTH, MAX_OUTPUT_BYTES,
    MAX_REFERENCE_BYTES, MAX_XML_BYTES, STRICT_NS, Sheet, Step, TRANSITIONAL_NS, raw::Attr,
};

/// Serialize a complete calculation chain with bounded allocation.
pub fn write(chain: &Chain, conformance: Conformance) -> Result<Vec<u8>> {
    if chain.cells.len() > MAX_CELLS {
        return Err(invalid("calculation chain has too many cells"));
    }
    let capacity = wire_len(chain, conformance)?;
    let mut xml = String::new();
    xml.try_reserve_exact(capacity)
        .map_err(|source| allocation("calculation-chain output", source))?;
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str("<calcChain xmlns=\"");
    xml.push_str(conformance.namespace());
    xml.push('"');
    for (name, value) in &chain.namespace_declarations {
        if name != "xmlns" {
            xml.push(' ');
            xml.push_str(name);
            xml.push_str("=\"");
            escape_attribute(&mut xml, value);
            xml.push('"');
        }
    }
    write_extension_attributes(&mut xml, &chain.attrs)?;
    xml.push('>');
    for cell in &chain.cells {
        xml.push_str("<c r=\"");
        escape_attribute(&mut xml, &cell.reference);
        xml.push('"');
        if cell.explicit_sheet {
            xml.push_str(" i=\"");
            push_u16(&mut xml, cell.sheet.get());
            xml.push('"');
        }
        match cell.step {
            Step::Same => {},
            Step::Level => write_bool_attribute(&mut xml, "l", true),
            Step::Child => write_bool_attribute(&mut xml, "s", true),
        }
        if cell.flags.contains(Flags::THREAD) {
            write_bool_attribute(&mut xml, "t", true);
        }
        if cell.flags.contains(Flags::ARRAY) {
            write_bool_attribute(&mut xml, "a", true);
        }
        write_extension_attributes(&mut xml, &cell.attrs)?;
        xml.push_str("/>");
    }
    if let Some(extension) = &chain.extension_list_xml {
        xml.push_str(extension);
    }
    xml.push_str("</calcChain>");
    Ok(xml.into_bytes())
}

#[derive(Default)]
struct Builder {
    cells: Vec<Cell>,
    seen_keys: HashSet<(Sheet, Address)>,
    ambiguous_key: Option<(Sheet, Address)>,
    extension_list_xml: Option<String>,
    namespace_declarations: Vec<(String, String)>,
    attrs: Vec<Attr>,
}

impl Builder {
    fn finish(self) -> Result<Chain> {
        if self.cells.is_empty() {
            return Err(invalid("calculation chain must contain at least one cell"));
        }
        let mut chain = Chain {
            cells: self.cells,
            ambiguous_key: self.ambiguous_key,
            extension_list_xml: self.extension_list_xml,
            namespace_declarations: self.namespace_declarations,
            attrs: self.attrs,
        };
        chain.ensure_sheet_boundaries();
        Ok(chain)
    }
}

/// Parse an isolated Calculation Chain part. Formula text is never evaluated.
pub fn read(xml: &[u8]) -> Result<Chain> {
    read_with_projection(xml).map(|(chain, _)| chain)
}

pub(crate) fn read_with_projection(xml: &[u8]) -> Result<(Chain, bool)> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "calculation-chain XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    preflight_raw_attributes(xml)?;
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)
        .map_err(|error| invalid(format!("calculation-chain MCE error: {error}")))?;
    let projected = matches!(&processed, Cow::Owned(_));
    let bytes = processed.as_ref();
    if bytes.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "processed calculation-chain XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut builder = Builder::default();
    let mut current_sheet = None;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut saw_extensions = false;
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid calculation-chain XML: {error}")))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if !saw_root => {
                validate_root(&namespace, &element, closed_root)?;
                saw_root = true;
                parse_root_attributes(&element, decoder, &resolver, &mut builder)?;
            },
            Event::Empty(element) if !saw_root => {
                validate_root(&namespace, &element, closed_root)?;
                saw_root = true;
                closed_root = true;
                parse_root_attributes(&element, decoder, &resolver, &mut builder)?;
            },
            Event::Empty(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"c") =>
            {
                if saw_extensions {
                    return Err(invalid("calculation cells must precede extLst"));
                }
                let cell = parse_cell(&element, decoder, &resolver, current_sheet)?;
                current_sheet = Some(cell.sheet);
                push_cell(&mut builder, cell)?;
            },
            Event::Start(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"c") =>
            {
                if saw_extensions {
                    return Err(invalid("calculation cells must precede extLst"));
                }
                let cell = parse_cell(&element, decoder, &resolver, current_sheet)?;
                let content_start = position(&reader)?;
                consume_leaf(&mut reader, b"c", content_start)?;
                current_sheet = Some(cell.sheet);
                push_cell(&mut builder, cell)?;
            },
            Event::Empty(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"extLst") =>
            {
                if std::mem::replace(&mut saw_extensions, true) {
                    return Err(invalid("duplicate calculation-chain extLst"));
                }
                let end = position(&reader)?;
                builder.extension_list_xml = Some(raw_range(bytes, start, end)?);
            },
            Event::Start(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"extLst") =>
            {
                if std::mem::replace(&mut saw_extensions, true) {
                    return Err(invalid("duplicate calculation-chain extLst"));
                }
                let end = consume_extension_list(&mut reader, start)?;
                builder.extension_list_xml = Some(raw_range(bytes, start, end)?);
            },
            Event::Start(element) | Event::Empty(element) if saw_root && !closed_root => {
                return Err(invalid(format!(
                    "unexpected calculation-chain child '{}'",
                    String::from_utf8_lossy(element.local_name().as_ref())
                )));
            },
            Event::End(element)
                if saw_root && !closed_root && element.local_name().as_ref() == b"calcChain" =>
            {
                closed_root = true
            },
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| invalid(format!("invalid calculation-chain text: {error}")))?
                    .trim()
                    .is_empty() => {},
            Event::Comment(_) | Event::Decl(_) => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTD and processing instructions are rejected in calculation-chain XML",
                ));
            },
            Event::Eof => break,
            _ => return Err(invalid("invalid calculation-chain XML structure")),
        }
    }
    if !saw_root || !closed_root {
        return Err(invalid("calculation-chain XML has no complete root"));
    }
    Ok((builder.finish()?, projected))
}

fn preflight_raw_attributes(xml: &[u8]) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("invalid calculation-chain XML: {error}")))?
        {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        invalid(format!("invalid calculation-chain attribute: {error}"))
                    })?;
                    validate_raw_attribute_size(attribute.key.as_ref(), attribute.value.as_ref())?;
                }
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}
fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    closed: bool,
) -> Result<()> {
    if closed || !is_name(namespace, element, b"calcChain") {
        return Err(invalid(
            "calculation-chain XML has an invalid or trailing root",
        ));
    }
    Ok(())
}

fn is_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TRANSITIONAL_NS.as_bytes() || *value == STRICT_NS.as_bytes())
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    builder: &mut Builder,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid calcChain attribute: {error}")))?;
        validate_raw_attribute_size(attribute.key.as_ref(), attribute.value.as_ref())?;
        let raw = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("calcChain attribute name is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid calcChain attribute value: {error}")))?
            .into_owned();
        validate_attribute_size(&raw, &value)?;
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            if raw != "xmlns" {
                if builder.namespace_declarations.len() >= MAX_EXTENSION_ATTRIBUTES {
                    return Err(invalid("too many calculation-chain namespace declarations"));
                }
                builder.namespace_declarations.push((raw, value));
            }
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if matches!(
            namespace,
            ResolveResult::Unbound | ResolveResult::Unknown(_)
        ) {
            return Err(invalid(format!("unexpected calcChain attribute '{raw}'")));
        }
        push_extension_attribute(&mut builder.attrs, raw, value)?;
    }
    Ok(())
}

fn parse_cell(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    inherited_sheet: Option<Sheet>,
) -> Result<Cell> {
    let mut reference = None;
    let mut sheet = None;
    let mut child = None;
    let mut new_level = None;
    let mut thread = None;
    let mut array = None;
    let mut attrs = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid calculation-cell attribute: {error}")))?;
        validate_raw_attribute_size(attribute.key.as_ref(), attribute.value.as_ref())?;
        let raw = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| {
                invalid(format!(
                    "calculation-cell attribute name is not UTF-8: {error}"
                ))
            })?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid calculation-cell value: {error}")))?
            .into_owned();
        validate_attribute_size(&raw, &value)?;
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) {
            match attribute.key.local_name().as_ref() {
                b"r" => set_once(&mut reference, value, "r")?,
                b"i" => set_once(&mut sheet, parse_sheet(&value)?, "i")?,
                b"s" => set_once(&mut child, parse_bool(&value, "s")?, "s")?,
                b"l" => set_once(&mut new_level, parse_bool(&value, "l")?, "l")?,
                b"t" => set_once(&mut thread, parse_bool(&value, "t")?, "t")?,
                b"a" => set_once(&mut array, parse_bool(&value, "a")?, "a")?,
                _ => {
                    return Err(invalid(format!(
                        "unexpected calculation-cell attribute '{raw}'"
                    )));
                },
            }
        } else if matches!(namespace, ResolveResult::Unknown(_)) {
            return Err(invalid(format!(
                "unbound calculation-cell attribute '{raw}'"
            )));
        } else {
            push_extension_attribute(&mut attrs, raw, value)?;
        }
    }
    let reference = reference.ok_or_else(|| invalid("calculation cell requires r"))?;
    let address = parse_reference(&reference)?;
    let explicit_sheet = sheet.is_some();
    let sheet = sheet.or(inherited_sheet).ok_or_else(|| {
        invalid("the first calculation-chain cell must specify sheet attribute i")
    })?;
    let child = child.unwrap_or(false);
    let new_level = new_level.unwrap_or(false);
    if child && new_level {
        return Err(invalid(
            "calculation-cell attributes l and s are mutually exclusive",
        ));
    }
    let step = if child {
        Step::Child
    } else if new_level {
        Step::Level
    } else {
        Step::Same
    };
    let mut flags = Flags::empty();
    flags.set(Flags::THREAD, thread.unwrap_or(false));
    flags.set(Flags::ARRAY, array.unwrap_or(false));
    Ok(Cell {
        reference: reference.into_boxed_str(),
        address,
        sheet,
        explicit_sheet,
        step,
        flags,
        attrs,
    })
}

fn push_cell(builder: &mut Builder, cell: Cell) -> Result<()> {
    if builder.cells.len() >= MAX_CELLS {
        return Err(invalid("calculation chain has too many cells"));
    }
    let key = (cell.sheet, cell.address);
    let duplicate = builder.seen_keys.contains(&key);
    builder
        .cells
        .try_reserve(1)
        .map_err(|source| allocation("calculation-chain cells", source))?;
    if !duplicate {
        builder
            .seen_keys
            .try_reserve(1)
            .map_err(|source| allocation("calculation-chain key index", source))?;
        builder.seen_keys.insert(key);
    } else if builder.ambiguous_key.is_none() {
        builder.ambiguous_key = Some(key);
    }
    builder.cells.push(cell);
    Ok(())
}

fn consume_leaf(reader: &mut NsReader<&[u8]>, local: &[u8], start: usize) -> Result<()> {
    loop {
        let event_start = position(reader)?;
        enforce_budget(
            start,
            event_start,
            MAX_CELL_CONTENT_BYTES,
            "calculation cell content exceeds its byte limit",
        )?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid calculation-cell XML: {error}")))?;
        let event_end = position(reader)?;
        enforce_budget(
            start,
            event_end,
            MAX_CELL_CONTENT_BYTES,
            "calculation cell content exceeds its byte limit",
        )?;
        match event {
            Event::End(element) if element.local_name().as_ref() == local => return Ok(()),
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| invalid(format!("invalid calculation-cell text: {error}")))?
                    .trim()
                    .is_empty() => {},
            Event::Comment(_) => {},
            Event::Start(_) | Event::Empty(_) | Event::CData(_) => {
                return Err(invalid("calculation cell must be empty"));
            },
            Event::Eof => return Err(invalid("unterminated calculation cell")),
            _ => return Err(invalid("invalid calculation-cell content")),
        }
    }
}

fn consume_extension_list(reader: &mut NsReader<&[u8]>, start: usize) -> Result<usize> {
    let mut depth = 1usize;
    let mut nodes = 0usize;
    while depth != 0 {
        let event_start = position(reader)?;
        enforce_budget(
            start,
            event_start,
            MAX_EXTENSION_BYTES,
            "calculation-chain extension list is too large",
        )?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid extension XML: {error}")))?;
        let event_end = position(reader)?;
        enforce_budget(
            start,
            event_end,
            MAX_EXTENSION_BYTES,
            "calculation-chain extension list is too large",
        )?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension nesting overflow"))?;
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension node count overflow"))?;
                if depth > MAX_EXTENSION_DEPTH || nodes > MAX_CELLS {
                    return Err(invalid("calculation-chain extension is too complex"));
                }
            },
            Event::Empty(_) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension node count overflow"))?;
                if nodes > MAX_CELLS {
                    return Err(invalid("calculation-chain extension has too many nodes"));
                }
            },
            Event::End(_) => depth -= 1,
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTD and processing instructions are rejected in extensions",
                ));
            },
            Event::Eof => return Err(invalid("unterminated calculation-chain extLst")),
            _ => {},
        }
    }
    position(reader)
}

fn parse_reference(value: &str) -> Result<Address> {
    if value.is_empty() || value.len() > MAX_REFERENCE_BYTES {
        return Err(invalid("calculation-cell reference has invalid length"));
    }
    Address::from_a1(value).map_err(Into::into)
}

fn parse_sheet(value: &str) -> Result<Sheet> {
    let value = value
        .parse::<u32>()
        .map_err(|_| invalid("calculation-cell i is not an unsigned integer"))?;
    Sheet::new(value)
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!(
            "invalid calculation-cell {name} boolean '{value}'"
        ))),
    }
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(invalid(format!(
            "duplicate calculation-cell {name} attribute"
        )));
    }
    Ok(())
}

fn push_extension_attribute(attributes: &mut Vec<Attr>, name: String, value: String) -> Result<()> {
    validate_attribute_size(&name, &value)?;
    if attributes.len() >= MAX_EXTENSION_ATTRIBUTES {
        return Err(invalid("too many preserved calculation-chain attributes"));
    }
    if attributes.iter().any(|attribute| attribute.name == name) {
        return Err(invalid(format!("duplicate preserved attribute '{name}'")));
    }
    attributes.push(Attr { name, value });
    Ok(())
}

fn write_extension_attributes(xml: &mut String, attributes: &[Attr]) -> Result<()> {
    if attributes.len() > MAX_EXTENSION_ATTRIBUTES {
        return Err(invalid("too many preserved calculation-chain attributes"));
    }
    for attribute in attributes {
        xml.push(' ');
        xml.push_str(&attribute.name);
        xml.push_str("=\"");
        escape_attribute(xml, &attribute.value);
        xml.push('"');
    }
    Ok(())
}

fn validate_attribute_size(name: &str, value: &str) -> Result<()> {
    if name.len() > MAX_ATTRIBUTE_BYTES || value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(invalid(format!(
            "calculation-chain attribute exceeds {MAX_ATTRIBUTE_BYTES} bytes"
        )));
    }
    Ok(())
}

fn validate_raw_attribute_size(name: &[u8], value: &[u8]) -> Result<()> {
    if name.len() > MAX_ATTRIBUTE_BYTES || value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(invalid(format!(
            "calculation-chain attribute exceeds {MAX_ATTRIBUTE_BYTES} bytes"
        )));
    }
    Ok(())
}

fn write_bool_attribute(xml: &mut String, name: &str, value: bool) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str(if value { "=\"1\"" } else { "=\"0\"" });
}

fn escape_attribute(xml: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '"' => xml.push_str("&quot;"),
            '\t' => xml.push_str("&#x9;"),
            '\n' => xml.push_str("&#xA;"),
            '\r' => xml.push_str("&#xD;"),
            _ => xml.push(character),
        }
    }
}

fn wire_len(chain: &Chain, conformance: Conformance) -> Result<usize> {
    let mut len = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.len();
    add_len(&mut len, "<calcChain xmlns=\"".len())?;
    add_len(&mut len, conformance.namespace().len())?;
    add_len(&mut len, 1)?;
    if chain.namespace_declarations.len() > MAX_EXTENSION_ATTRIBUTES
        || chain.attrs.len() > MAX_EXTENSION_ATTRIBUTES
    {
        return Err(invalid("too many calculation-chain root attributes"));
    }
    for (name, value) in &chain.namespace_declarations {
        validate_attribute_size(name, value)?;
        if name != "xmlns" {
            add_len(&mut len, attribute_len(name, value)?)?;
        }
    }
    for attribute in &chain.attrs {
        add_len(&mut len, attribute_len(&attribute.name, &attribute.value)?)?;
    }
    add_len(&mut len, 1)?;
    for (position, cell) in chain.cells.iter().enumerate() {
        if position == 0 && !cell.explicit_sheet {
            return Err(invalid(
                "the first calculation-chain cell must carry an explicit sheet ID",
            ));
        }
        if cell.reference.len() > MAX_REFERENCE_BYTES
            || parse_reference(&cell.reference)? != cell.address
        {
            return Err(invalid(
                "calculation-cell reference no longer matches its address",
            ));
        }
        if cell.attrs.len() > MAX_EXTENSION_ATTRIBUTES {
            return Err(invalid("too many calculation-cell extension attributes"));
        }
        add_len(&mut len, "<c r=\"".len())?;
        add_len(&mut len, escaped_len(&cell.reference)?)?;
        add_len(&mut len, 1)?;
        if cell.explicit_sheet {
            add_len(&mut len, " i=\"".len())?;
            add_len(&mut len, decimal_len(u32::from(cell.sheet.get())))?;
            add_len(&mut len, 1)?;
        }
        if cell.step != Step::Same {
            add_len(&mut len, 6)?;
        }
        add_len(
            &mut len,
            usize::from(cell.flags.contains(Flags::THREAD)) * 6,
        )?;
        add_len(&mut len, usize::from(cell.flags.contains(Flags::ARRAY)) * 6)?;
        for attribute in &cell.attrs {
            add_len(&mut len, attribute_len(&attribute.name, &attribute.value)?)?;
        }
        add_len(&mut len, 2)?;
    }
    if let Some(extension) = &chain.extension_list_xml {
        if extension.len() > MAX_EXTENSION_BYTES {
            return Err(invalid("calculation-chain extension list is too large"));
        }
        add_len(&mut len, extension.len())?;
    }
    add_len(&mut len, "</calcChain>".len())?;
    if len > MAX_OUTPUT_BYTES {
        return Err(invalid(format!(
            "calculation-chain output exceeds {MAX_OUTPUT_BYTES} bytes"
        )));
    }
    Ok(len)
}

fn attribute_len(name: &str, value: &str) -> Result<usize> {
    validate_attribute_size(name, value)?;
    let mut len = name.len();
    add_len(&mut len, escaped_len(value)?)?;
    add_len(&mut len, 4)?;
    Ok(len)
}

fn escaped_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |mut len, character| {
        let bytes = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        add_len(&mut len, bytes)?;
        Ok(len)
    })
}

fn add_len(total: &mut usize, value: usize) -> Result<()> {
    *total = total
        .checked_add(value)
        .ok_or_else(|| invalid("calculation-chain output length overflow"))?;
    Ok(())
}

fn decimal_len(mut value: u32) -> usize {
    let mut len = 1;
    while value >= 10 {
        value /= 10;
        len += 1;
    }
    len
}

fn push_u16(output: &mut String, mut value: u16) {
    let mut digits = [0u8; 5];
    let mut start = digits.len();
    loop {
        start -= 1;
        digits[start] = b'0' + (value % 10) as u8;
        value /= 10;
        if value == 0 {
            break;
        }
    }
    for digit in &digits[start..] {
        output.push(char::from(*digit));
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("calculation-chain XML offset overflow"))
}

fn raw_range(bytes: &[u8], start: usize, end: usize) -> Result<String> {
    if end < start || end - start > MAX_EXTENSION_BYTES {
        return Err(invalid("calculation-chain extension list is too large"));
    }
    std::str::from_utf8(
        bytes
            .get(start..end)
            .ok_or_else(|| invalid("invalid calculation-chain extension range"))?,
    )
    .map(str::to_owned)
    .map_err(|error| invalid(format!("calculation-chain extension is not UTF-8: {error}")))
}

fn enforce_budget(start: usize, current: usize, limit: usize, message: &'static str) -> Result<()> {
    let consumed = current
        .checked_sub(start)
        .ok_or_else(|| invalid("calculation-chain XML offset moved backwards"))?;
    if consumed > limit {
        return Err(invalid(message));
    }
    Ok(())
}
