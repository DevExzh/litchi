//! Bounded, non-resolving parser for the legacy XLS `XML` stream.

use super::super::model::{DataBinding, Map, MapInfo, OpaqueXml, Schema};
use super::super::validation;
use crate::{Error, Result};

pub(super) const MAX_STREAM_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1 << 20;

#[derive(Debug, Clone)]
pub(super) struct Attribute {
    name: Vec<u8>,
    value: String,
}

#[derive(Debug)]
pub(super) struct StartTag<'a> {
    name: &'a [u8],
    attributes: Vec<Attribute>,
    empty: bool,
    start: usize,
    end: usize,
}

#[derive(Debug)]
pub(super) enum Token<'a> {
    Start(StartTag<'a>),
    End { name: &'a [u8], end: usize },
    Text(&'a [u8]),
    CData,
    Comment,
    ProcessingInstruction,
}

pub(super) struct Parser<'a> {
    input: &'a [u8],
    offset: usize,
    events: usize,
}

impl<'a> Parser<'a> {
    pub(super) fn new(input: &'a [u8]) -> Result<Self> {
        if input.len() > MAX_STREAM_BYTES {
            return Err(limit("XML stream exceeds the 16 MiB limit"));
        }
        let text = std::str::from_utf8(input)
            .map_err(|error| invalid(format!("XML stream is not UTF-8: {error}")))?;
        if text.chars().any(|character| !is_xml_char(character)) {
            return Err(invalid("XML stream contains an XML-forbidden character"));
        }
        Ok(Self {
            input,
            offset: 0,
            events: 0,
        })
    }

    pub(super) fn next(&mut self) -> Result<Option<Token<'a>>> {
        self.events = self
            .events
            .checked_add(1)
            .ok_or_else(|| limit("XML event count overflows"))?;
        if self.events > MAX_EVENTS {
            return Err(limit("XML event count exceeds the safety limit"));
        }
        if self.offset == self.input.len() {
            return Ok(None);
        }
        if self.input[self.offset] != b'<' {
            let start = self.offset;
            while self.offset < self.input.len() && self.input[self.offset] != b'<' {
                self.offset += 1;
            }
            let text = self
                .input
                .get(start..self.offset)
                .ok_or_else(|| invalid("XML text range is invalid"))?;
            decode_entities(text)?;
            return Ok(Some(Token::Text(text)));
        }

        let start = self.offset;
        let rest = self
            .input
            .get(self.offset..)
            .ok_or_else(|| invalid("XML token starts outside the stream"))?;
        if rest.starts_with(b"<!--") {
            self.offset = find_marker(self.input, self.offset + 4, b"-->")?;
            return Ok(Some(Token::Comment));
        }
        if rest.starts_with(b"<?") {
            self.offset = find_marker(self.input, self.offset + 2, b"?>")?;
            return Ok(Some(Token::ProcessingInstruction));
        }
        if rest.starts_with(b"<![CDATA[") {
            self.offset = find_marker(self.input, self.offset + 9, b"]]>")?;
            return Ok(Some(Token::CData));
        }
        if rest.starts_with(b"</") {
            return self.parse_end(start).map(Some);
        }
        if rest.starts_with(b"<!") {
            return Err(invalid("DTD and other declarations are not permitted"));
        }
        self.parse_start(start).map(Some)
    }

    fn parse_start(&mut self, start: usize) -> Result<Token<'a>> {
        self.offset += 1;
        let name_start = self.offset;
        self.skip_name();
        let name = self
            .input
            .get(name_start..self.offset)
            .ok_or_else(|| invalid("XML start-tag name range is invalid"))?;
        validate_name(name)?;
        let mut attributes = Vec::new();
        loop {
            self.skip_space();
            if self
                .input
                .get(self.offset..)
                .is_some_and(|value| value.starts_with(b"/>"))
            {
                self.offset += 2;
                return Ok(Token::Start(StartTag {
                    name,
                    attributes,
                    empty: true,
                    start,
                    end: self.offset,
                }));
            }
            if self.input.get(self.offset) == Some(&b'>') {
                self.offset += 1;
                return Ok(Token::Start(StartTag {
                    name,
                    attributes,
                    empty: false,
                    start,
                    end: self.offset,
                }));
            }
            if attributes.len() >= MAX_ATTRIBUTES {
                return Err(limit("XML attribute count exceeds the safety limit"));
            }
            let attribute_start = self.offset;
            self.skip_name();
            let attribute_name = self
                .input
                .get(attribute_start..self.offset)
                .ok_or_else(|| invalid("XML attribute name range is invalid"))?;
            validate_name(attribute_name)?;
            self.skip_space();
            if self.input.get(self.offset) != Some(&b'=') {
                return Err(invalid("XML attribute lacks an equals sign"));
            }
            self.offset += 1;
            self.skip_space();
            let quote = *self
                .input
                .get(self.offset)
                .ok_or_else(|| invalid("XML attribute value is truncated"))?;
            if quote != b'\'' && quote != b'"' {
                return Err(invalid("XML attribute value is not quoted"));
            }
            self.offset += 1;
            let value_start = self.offset;
            while self.offset < self.input.len() && self.input[self.offset] != quote {
                if self.input[self.offset] == b'<' {
                    return Err(invalid("XML attribute value contains '<'"));
                }
                self.offset += 1;
            }
            let value_end = self.offset;
            if self.input.get(self.offset) != Some(&quote) {
                return Err(invalid("XML attribute value is unterminated"));
            }
            self.offset += 1;
            let raw = self
                .input
                .get(value_start..value_end)
                .ok_or_else(|| invalid("XML attribute value range is invalid"))?;
            if raw.len() > MAX_ATTRIBUTE_BYTES {
                return Err(limit("XML attribute value exceeds the safety limit"));
            }
            if attributes
                .iter()
                .any(|attribute: &Attribute| attribute.name.as_slice() == attribute_name)
            {
                return Err(invalid("XML element contains a duplicate attribute"));
            }
            attributes.push(Attribute {
                name: attribute_name.to_vec(),
                value: decode_entities(raw)?,
            });
        }
    }

    fn parse_end(&mut self, _start: usize) -> Result<Token<'a>> {
        self.offset += 2;
        let name_start = self.offset;
        self.skip_name();
        let name = self
            .input
            .get(name_start..self.offset)
            .ok_or_else(|| invalid("XML end-tag name range is invalid"))?;
        validate_name(name)?;
        self.skip_space();
        if self.input.get(self.offset) != Some(&b'>') {
            return Err(invalid("XML end tag is malformed"));
        }
        self.offset += 1;
        Ok(Token::End {
            name,
            end: self.offset,
        })
    }

    fn skip_name(&mut self) {
        while self.offset < self.input.len() && is_name_byte(self.input[self.offset]) {
            self.offset += 1;
        }
    }

    fn skip_space(&mut self) {
        while self.offset < self.input.len() && is_space(self.input[self.offset]) {
            self.offset += 1;
        }
    }

    fn slice(&self, start: usize, end: usize) -> Result<&'a [u8]> {
        self.input
            .get(start..end)
            .ok_or_else(|| invalid("opaque XML range is invalid"))
    }
}

pub(super) fn capture_element<'a>(parser: &mut Parser<'a>, first: StartTag<'a>) -> Result<Vec<u8>> {
    let start = first.start;
    if first.empty {
        return Ok(parser.slice(start, first.end)?.to_vec());
    }
    let mut names = vec![first.name];
    let mut depth = 1usize;
    loop {
        let token = parser
            .next()?
            .ok_or_else(|| invalid("opaque XML element is not closed"))?;
        match token {
            Token::Start(tag) => {
                if !tag.empty {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("opaque XML depth overflows"))?;
                    if depth > MAX_DEPTH {
                        return Err(limit("opaque XML depth exceeds the safety limit"));
                    }
                    names.push(tag.name);
                }
            },
            Token::End { name, end } => {
                let expected = names
                    .pop()
                    .ok_or_else(|| invalid("opaque XML has an unexpected closing tag"))?;
                if expected != name {
                    return Err(invalid("opaque XML has mismatched element names"));
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opaque XML depth underflows"))?;
                if depth == 0 {
                    return Ok(parser.slice(start, end)?.to_vec());
                }
            },
            Token::Text(text) => {
                decode_entities(text)?;
            },
            Token::CData | Token::Comment | Token::ProcessingInstruction => {},
        }
    }
}

pub(crate) fn validate_opaque_element(input: &[u8]) -> Result<()> {
    let mut parser = Parser::new(input)?;
    let first =
        next_element(&mut parser, false)?.ok_or_else(|| invalid("opaque XML element is empty"))?;
    let captured = capture_element(&mut parser, first)?;
    if captured.as_slice() != input {
        return Err(invalid(
            "opaque XML must contain exactly one complete element",
        ));
    }
    if parser.next()?.is_some() {
        return Err(invalid("opaque XML has trailing content"));
    }
    Ok(())
}

pub(crate) fn parse_stream(input: &[u8]) -> Result<MapInfo> {
    let mut parser = Parser::new(input)?;
    let root = next_element(&mut parser, true)?
        .ok_or_else(|| invalid("XML stream has no MapInfo root"))?;
    if local_name(root.name) != b"MapInfo" {
        return Err(invalid("XML stream root is not MapInfo"));
    }
    let selection = required_attr(&root, b"SelectionNamespaces")?;
    let namespaces = namespace_declarations(&root)?;
    reject_attrs(&root, &[b"SelectionNamespaces"])?;
    let mut schemas = Vec::new();
    let mut maps = Vec::new();
    if !root.empty {
        let mut map_phase = false;
        loop {
            let token =
                next_significant(&mut parser)?.ok_or_else(|| invalid("MapInfo is not closed"))?;
            match token {
                Token::End { name, .. } if local_name(name) == b"MapInfo" => break,
                Token::Start(tag) if local_name(tag.name) == b"Schema" && !map_phase => {
                    schemas.push(parse_schema(&mut parser, tag)?);
                },
                Token::Start(tag) if local_name(tag.name) == b"Map" => {
                    map_phase = true;
                    maps.push(parse_map(&mut parser, tag)?);
                },
                Token::End { .. } => return Err(invalid("MapInfo has a mismatched end tag")),
                Token::Start(_) | Token::CData => {
                    return Err(invalid("MapInfo contains an unexpected child"));
                },
                Token::Text(text) => ensure_whitespace(text)?,
                Token::Comment | Token::ProcessingInstruction => {},
            }
        }
    }
    ensure_trailing(&mut parser)?;
    let value = MapInfo::from_parts(selection, namespaces, schemas, maps)?;
    validation::validate(&value)?;
    Ok(value)
}

fn parse_schema<'a>(parser: &mut Parser<'a>, tag: StartTag<'a>) -> Result<Schema> {
    let id = required_attr(&tag, b"ID")?;
    let schema_ref = optional_attr(&tag, b"SchemaRef")?;
    let namespace = optional_attr(&tag, b"Namespace")?;
    let namespaces = namespace_declarations(&tag)?;
    reject_attrs(&tag, &[b"ID", b"SchemaRef", b"Namespace"])?;
    if tag.empty {
        return Err(invalid("Schema must contain one opaque XSD element"));
    }
    let child = next_significant(parser)?.ok_or_else(|| invalid("Schema is not closed"))?;
    let payload = match child {
        Token::Start(child) => OpaqueXml::from_bytes(capture_element(parser, child)?)?,
        Token::Text(text) => {
            ensure_whitespace(text)?;
            return Err(invalid("Schema must contain one opaque XSD element"));
        },
        _ => return Err(invalid("Schema must contain one opaque XSD element")),
    };
    ensure_end(parser, b"Schema")?;
    Schema::from_parts(id, schema_ref, namespace, namespaces, payload)
}

fn parse_map<'a>(parser: &mut Parser<'a>, tag: StartTag<'a>) -> Result<Map> {
    let id = required_attr(&tag, b"ID")?;
    let name = required_attr(&tag, b"Name")?;
    let root = required_attr(&tag, b"RootElement")?;
    let schema_id = required_attr(&tag, b"SchemaID")?;
    let show = required_bool(&tag, b"ShowImportExportValidationErrors")?;
    let auto_fit = required_bool(&tag, b"AutoFit")?;
    let append = required_bool(&tag, b"Append")?;
    let preserve_sort = required_bool(&tag, b"PreserveSortAFLayout")?;
    let preserve_format = required_bool(&tag, b"PreserveFormat")?;
    let namespaces = namespace_declarations(&tag)?;
    reject_attrs(
        &tag,
        &[
            b"ID",
            b"Name",
            b"RootElement",
            b"SchemaID",
            b"ShowImportExportValidationErrors",
            b"AutoFit",
            b"Append",
            b"PreserveSortAFLayout",
            b"PreserveFormat",
        ],
    )?;
    let binding = if tag.empty {
        None
    } else {
        match next_significant(parser)?.ok_or_else(|| invalid("Map is not closed"))? {
            Token::Start(child) if local_name(child.name) == b"DataBinding" => {
                let value = parse_binding(parser, child)?;
                ensure_end(parser, b"Map")?;
                Some(value)
            },
            Token::End { name, .. } if local_name(name) == b"Map" => None,
            Token::Text(text) => {
                ensure_whitespace(text)?;
                return Err(invalid("Map is not closed"));
            },
            _ => return Err(invalid("Map contains an unexpected child")),
        }
    };
    Map::from_parts(
        id,
        name,
        root,
        schema_id,
        show,
        auto_fit,
        append,
        preserve_sort,
        preserve_format,
        binding,
        namespaces,
    )
}

fn parse_binding<'a>(parser: &mut Parser<'a>, tag: StartTag<'a>) -> Result<DataBinding> {
    let binding_name = optional_attr(&tag, b"DataBindingName")?;
    let file_binding = required_attr(&tag, b"FileBinding")?;
    let file_binding_name = optional_attr(&tag, b"FileBindingName")?;
    let load_mode = required_attr(&tag, b"DataBindingLoadMode")?;
    let namespaces = namespace_declarations(&tag)?;
    reject_attrs(
        &tag,
        &[
            b"DataBindingName",
            b"FileBinding",
            b"FileBindingName",
            b"DataBindingLoadMode",
        ],
    )?;
    let payload = if tag.empty {
        None
    } else {
        match next_significant(parser)?.ok_or_else(|| invalid("DataBinding is not closed"))? {
            Token::Start(child) => {
                let value = OpaqueXml::from_bytes(capture_element(parser, child)?)?;
                ensure_end(parser, b"DataBinding")?;
                Some(value)
            },
            Token::End { name, .. } if local_name(name) == b"DataBinding" => None,
            Token::Text(text) => {
                ensure_whitespace(text)?;
                return Err(invalid("DataBinding contains unexpected text"));
            },
            _ => return Err(invalid("DataBinding contains an unexpected child")),
        }
    };
    DataBinding::from_parts(
        binding_name,
        file_binding,
        file_binding_name,
        load_mode,
        payload,
        namespaces,
    )
}

fn next_element<'a>(parser: &mut Parser<'a>, allow_misc: bool) -> Result<Option<StartTag<'a>>> {
    loop {
        let Some(token) = parser.next()? else {
            return Ok(None);
        };
        match token {
            Token::Start(tag) => return Ok(Some(tag)),
            Token::Text(text) if allow_misc => ensure_whitespace(text)?,
            Token::Comment | Token::ProcessingInstruction if allow_misc => {},
            Token::Text(_) | Token::End { .. } => {
                return Err(invalid("XML has unexpected content before its root"));
            },
            Token::CData => return Err(invalid("CDATA appears outside an element")),
            Token::Comment | Token::ProcessingInstruction => {
                if !allow_misc {
                    return Err(invalid("XML has content before its root"));
                }
            },
        }
    }
}

fn next_significant<'a>(parser: &mut Parser<'a>) -> Result<Option<Token<'a>>> {
    loop {
        let Some(token) = parser.next()? else {
            return Ok(None);
        };
        match token {
            Token::Text(text) => ensure_whitespace(text)?,
            Token::Comment | Token::ProcessingInstruction => {},
            token => return Ok(Some(token)),
        }
    }
}

fn ensure_end(parser: &mut Parser<'_>, expected: &[u8]) -> Result<()> {
    match next_significant(parser)?.ok_or_else(|| invalid("XML element is not closed"))? {
        Token::End { name, .. } if local_name(name) == expected => Ok(()),
        Token::End { .. } => Err(invalid("XML element has a mismatched closing tag")),
        _ => Err(invalid(
            "XML element contains more than its permitted children",
        )),
    }
}

fn ensure_trailing(parser: &mut Parser<'_>) -> Result<()> {
    while let Some(token) = parser.next()? {
        match token {
            Token::Text(text) => ensure_whitespace(text)?,
            Token::Comment | Token::ProcessingInstruction => {},
            _ => return Err(invalid("XML stream contains content after MapInfo")),
        }
    }
    Ok(())
}

fn required_attr(tag: &StartTag<'_>, name: &[u8]) -> Result<String> {
    optional_attr(tag, name)?.ok_or_else(|| invalid(format!("missing {} attribute", display(name))))
}

fn optional_attr(tag: &StartTag<'_>, name: &[u8]) -> Result<Option<String>> {
    Ok(tag
        .attributes
        .iter()
        .find(|attribute| attribute.name.as_slice() == name)
        .map(|attribute| attribute.value.clone()))
}

fn required_bool(tag: &StartTag<'_>, name: &[u8]) -> Result<bool> {
    match required_attr(tag, name)?.as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(invalid(format!(
            "{} must be lowercase true or false",
            display(name)
        ))),
    }
}

fn namespace_declarations(tag: &StartTag<'_>) -> Result<Vec<(String, String)>> {
    let mut result = Vec::new();
    for attribute in &tag.attributes {
        if attribute.name.as_slice() == b"xmlns" {
            result.push((String::new(), attribute.value.clone()));
        } else if let Some(prefix) = attribute.name.strip_prefix(b"xmlns:") {
            validate_name(prefix)?;
            result.push((
                String::from_utf8(prefix.to_vec())
                    .map_err(|_error| invalid("namespace prefix is not UTF-8"))?,
                attribute.value.clone(),
            ));
        }
    }
    Ok(result)
}

fn reject_attrs(tag: &StartTag<'_>, allowed: &[&[u8]]) -> Result<()> {
    for attribute in &tag.attributes {
        if attribute.name.as_slice() == b"xmlns"
            || attribute.name.starts_with(b"xmlns:")
            || allowed.contains(&attribute.name.as_slice())
        {
            continue;
        }
        return Err(invalid(format!(
            "{} has an unexpected attribute",
            display(tag.name)
        )));
    }
    Ok(())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn display(value: &[u8]) -> String {
    String::from_utf8_lossy(value).into_owned()
}

fn find_marker(input: &[u8], start: usize, marker: &[u8]) -> Result<usize> {
    let mut index = start;
    while index
        .checked_add(marker.len())
        .is_some_and(|end| end <= input.len())
    {
        if input[index..index + marker.len()] == *marker {
            return Ok(index + marker.len());
        }
        index += 1;
    }
    Err(invalid("XML construct is not closed"))
}

fn validate_name(name: &[u8]) -> Result<()> {
    let Some(&first) = name.first() else {
        return Err(invalid("XML name is empty"));
    };
    if !is_name_start(first) || name.iter().skip(1).any(|byte| !is_name_byte(*byte)) {
        return Err(invalid("XML name is malformed"));
    }
    if name.starts_with(b":") || name.ends_with(b":") {
        return Err(invalid("XML name has an empty namespace component"));
    }
    Ok(())
}

fn decode_entities(input: &[u8]) -> Result<String> {
    let mut output = String::with_capacity(input.len());
    let mut start = 0usize;
    for index in 0..input.len() {
        if input[index] != b'&' {
            continue;
        }
        output.push_str(
            std::str::from_utf8(
                input
                    .get(start..index)
                    .ok_or_else(|| invalid("XML entity range is invalid"))?,
            )
            .map_err(|_error| invalid("XML text is not UTF-8"))?,
        );
        let end = input
            .get(index + 1..)
            .and_then(|rest| rest.iter().position(|byte| *byte == b';'))
            .and_then(|relative| index.checked_add(relative + 1))
            .ok_or_else(|| invalid("XML entity is unterminated"))?;
        let entity = input
            .get(index + 1..end)
            .ok_or_else(|| invalid("XML entity range is invalid"))?;
        match entity {
            b"amp" => output.push('&'),
            b"lt" => output.push('<'),
            b"gt" => output.push('>'),
            b"quot" => output.push('"'),
            b"apos" => output.push('\''),
            value if value.starts_with(b"#x") || value.starts_with(b"#X") => {
                numeric(&mut output, &value[2..], 16)?;
            },
            value if value.starts_with(b"#") => numeric(&mut output, &value[1..], 10)?,
            _ => return Err(invalid("XML contains an unknown entity")),
        }
        start = end + 1;
    }
    output.push_str(
        std::str::from_utf8(
            input
                .get(start..)
                .ok_or_else(|| invalid("XML text range is invalid"))?,
        )
        .map_err(|_error| invalid("XML text is not UTF-8"))?,
    );
    if output.chars().any(|character| !is_xml_char(character)) {
        return Err(invalid("XML text contains a forbidden character"));
    }
    Ok(output)
}

fn numeric(output: &mut String, digits: &[u8], radix: u32) -> Result<()> {
    let text =
        std::str::from_utf8(digits).map_err(|_error| invalid("XML numeric entity is invalid"))?;
    let value = u32::from_str_radix(text, radix)
        .map_err(|_error| invalid("XML numeric entity is invalid"))?;
    let character = char::from_u32(value).ok_or_else(|| invalid("XML entity is not a scalar"))?;
    if !is_xml_char(character) {
        return Err(invalid("XML entity contains a forbidden character"));
    }
    output.push(character);
    Ok(())
}

fn ensure_whitespace(input: &[u8]) -> Result<()> {
    if decode_entities(input)?
        .chars()
        .any(|character| !character.is_whitespace())
    {
        return Err(invalid("XML grammar contains unexpected text"));
    }
    Ok(())
}

fn is_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

fn is_name_start(byte: u8) -> bool {
    byte.is_ascii_alphabetic() || matches!(byte, b'_' | b':') || byte >= 0x80
}

fn is_name_byte(byte: u8) -> bool {
    is_name_start(byte) || byte.is_ascii_digit() || matches!(byte, b'-' | b'.')
}

fn is_xml_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&character)
        || ('\u{E000}'..='\u{FFFD}').contains(&character)
        || ('\u{10000}'..='\u{10FFFF}').contains(&character)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map: {}", message.into()))
}

fn limit(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map resource limit: {}", message.into()))
}
