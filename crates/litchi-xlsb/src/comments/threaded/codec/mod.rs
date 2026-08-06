//! Bounded XML codec for XLSB threaded-comments and persons parts.

use std::fmt::Write as FmtWrite;

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::custom_xml::valid_guid;
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::semantic::{Comment, Comments, Mention, People, Person, RawAttribute, RawXml};
use super::validation::{Error, Result, validate_comments, validate_people};
use super::{
    MAX_COMMENTS, MAX_EXTENSION_BYTES, MAX_EXTENSIONS, MAX_MENTIONS, MAX_PART_BYTES, MAX_PERSONS,
    MAX_XML_DEPTH,
};

const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const XML_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// Parse a bounded `personList` part.
pub fn parse_persons(xml: impl AsRef<[u8]>) -> Result<People> {
    let xml = xml.as_ref();
    check_size(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut people = People::default();
    let mut root_open = false;
    let mut root_closed = false;

    loop {
        let event = next_event(&mut reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Decl(_) | Event::DocType(_) | Event::Comment(_) => {},
            Event::Start(element) if !root_open => {
                require_root(&namespace, element.name(), b"personList", root_closed)?;
                people.attributes = raw_attributes(&element, &[], reader.decoder())?;
                root_open = true;
            },
            Event::Empty(element) if !root_open => {
                require_root(&namespace, element.name(), b"personList", root_closed)?;
                people.attributes = raw_attributes(&element, &[], reader.decoder())?;
                break;
            },
            Event::Start(element) if root_open && !root_closed => {
                if is_name(&namespace, element.name(), b"person") {
                    people.persons.push(parse_person(&mut reader, element)?);
                } else {
                    people
                        .extensions
                        .push(capture_unknown(&mut reader, Event::Start(element))?);
                }
            },
            Event::Empty(element) if root_open && !root_closed => {
                if is_name(&namespace, element.name(), b"person") {
                    people
                        .persons
                        .push(parse_person_empty(&element, reader.decoder())?);
                } else {
                    people
                        .extensions
                        .push(capture_unknown(&mut reader, Event::Empty(element))?);
                }
            },
            Event::End(element) if root_open && !root_closed => {
                if !is_name(&namespace, element.name(), b"personList") {
                    return Err(invalid("persons part has an invalid root closing element"));
                }
                root_closed = true;
            },
            Event::Text(text) if root_open && !root_closed => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("persons part contains unexpected character data"));
                }
            },
            Event::CData(text) if root_open && !root_closed => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("persons part contains unexpected character data"));
                }
            },
            Event::Eof => {
                if !root_closed {
                    return Err(invalid("persons part has a missing or unterminated root"));
                }
                break;
            },
            _ => return Err(invalid("invalid persons XML structure")),
        }
    }
    if people.persons.len() > MAX_PERSONS {
        return Err(invalid("persons part contains too many people"));
    }
    validate_people(&people)?;
    Ok(people)
}

/// Parse a bounded `ThreadedComments` part.
pub fn parse_comments(xml: impl AsRef<[u8]>) -> Result<Comments> {
    let xml = xml.as_ref();
    check_size(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let mut comments = Comments::default();
    let mut root_open = false;
    let mut root_closed = false;

    loop {
        let event = next_event(&mut reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Decl(_) | Event::DocType(_) | Event::Comment(_) => {},
            Event::Start(element) if !root_open => {
                require_root(&namespace, element.name(), b"ThreadedComments", root_closed)?;
                comments.attributes = raw_attributes(&element, &[], reader.decoder())?;
                root_open = true;
            },
            Event::Empty(element) if !root_open => {
                require_root(&namespace, element.name(), b"ThreadedComments", root_closed)?;
                comments.attributes = raw_attributes(&element, &[], reader.decoder())?;
                break;
            },
            Event::Start(element) if root_open && !root_closed => {
                if is_name(&namespace, element.name(), b"threadedComment") {
                    comments.comments.push(parse_comment(&mut reader, element)?);
                } else {
                    comments
                        .extensions
                        .push(capture_unknown(&mut reader, Event::Start(element))?);
                }
            },
            Event::Empty(element) if root_open && !root_closed => {
                if is_name(&namespace, element.name(), b"threadedComment") {
                    comments
                        .comments
                        .push(parse_comment_empty(&element, reader.decoder())?);
                } else {
                    comments
                        .extensions
                        .push(capture_unknown(&mut reader, Event::Empty(element))?);
                }
            },
            Event::End(element) if root_open && !root_closed => {
                if !is_name(&namespace, element.name(), b"ThreadedComments") {
                    return Err(invalid(
                        "threaded-comments part has an invalid root closing element",
                    ));
                }
                root_closed = true;
            },
            Event::Text(text) if root_open && !root_closed => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid(
                        "threaded-comments part contains unexpected character data",
                    ));
                }
            },
            Event::CData(text) if root_open && !root_closed => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid(
                        "threaded-comments part contains unexpected character data",
                    ));
                }
            },
            Event::Eof => {
                if !root_closed {
                    return Err(invalid(
                        "threaded-comments part has a missing or unterminated root",
                    ));
                }
                break;
            },
            _ => return Err(invalid("invalid threaded-comments XML structure")),
        }
    }
    if comments.comments.len() > MAX_COMMENTS {
        return Err(invalid("threaded-comments part contains too many comments"));
    }
    let mentions = comments
        .comments
        .iter()
        .try_fold(0usize, |count, comment| {
            count
                .checked_add(comment.mentions.len())
                .ok_or_else(|| invalid("mention count overflows"))
        })?;
    if mentions > MAX_MENTIONS {
        return Err(invalid("threaded-comments part contains too many mentions"));
    }
    validate_comments(&comments)?;
    Ok(comments)
}

/// Serialize a persons part after validating its bounded semantic model.
pub fn write_persons(people: &People) -> Result<Vec<u8>> {
    validate_people(people)?;
    let mut xml = String::with_capacity(1024);
    xml.push_str(XML_HEADER);
    xml.push_str("<personList xmlns=\"");
    xml.push_str(std::str::from_utf8(NAMESPACE).expect("static namespace is UTF-8"));
    xml.push('"');
    write_attributes(&mut xml, &people.attributes, &[b"xmlns"])?;
    xml.push('>');
    for person in &people.persons {
        write_person(&mut xml, person)?;
    }
    write_extensions(&mut xml, &people.extensions)?;
    xml.push_str("</personList>");
    Ok(xml.into_bytes())
}

/// Serialize a worksheet threaded-comments part after validating its model.
pub fn write_comments(comments: &Comments) -> Result<Vec<u8>> {
    validate_comments(comments)?;
    let mut xml = String::with_capacity(4096);
    xml.push_str(XML_HEADER);
    xml.push_str("<ThreadedComments xmlns=\"");
    xml.push_str(std::str::from_utf8(NAMESPACE).expect("static namespace is UTF-8"));
    xml.push('"');
    write_attributes(&mut xml, &comments.attributes, &[b"xmlns"])?;
    xml.push('>');
    for comment in &comments.comments {
        write_comment(&mut xml, comment)?;
    }
    write_extensions(&mut xml, &comments.extensions)?;
    xml.push_str("</ThreadedComments>");
    Ok(xml.into_bytes())
}

fn parse_person(reader: &mut NsReader<&[u8]>, element: BytesStart<'static>) -> Result<Person> {
    let mut person = parse_person_empty(&element, reader.decoder())?;
    loop {
        let event = next_event(reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                person
                    .extensions
                    .push(capture_unknown(reader, Event::Start(element))?);
            },
            Event::Empty(element) => {
                person
                    .extensions
                    .push(capture_unknown(reader, Event::Empty(element))?);
            },
            Event::End(end) if is_name(&namespace, end.name(), b"person") => return Ok(person),
            Event::Text(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("person contains unexpected character data"));
                }
            },
            Event::CData(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("person contains unexpected character data"));
                }
            },
            Event::End(_) => return Err(invalid("person has an invalid closing element")),
            Event::Eof => return Err(invalid("person is unterminated")),
            _ => {},
        }
    }
}

fn parse_person_empty(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Person> {
    Ok(Person {
        display_name: required_string(element, b"displayName", decoder, "person display name")?,
        id: required_guid(element, b"id", decoder, "person ID")?,
        user_id: unqualified_attribute_value(element, b"userId", decoder)
            .map_err(|error| Error::Encoding(error.to_string()))?,
        provider_id: unqualified_attribute_value(element, b"providerId", decoder)
            .map_err(|error| Error::Encoding(error.to_string()))?,
        attributes: raw_attributes(
            element,
            &[b"displayName", b"id", b"userId", b"providerId"],
            decoder,
        )?,
        extensions: Vec::new(),
    })
}

fn parse_comment(reader: &mut NsReader<&[u8]>, element: BytesStart<'static>) -> Result<Comment> {
    let mut comment = parse_comment_empty(&element, reader.decoder())?;
    let mut saw_text = false;
    let mut saw_mentions = false;
    loop {
        let event = next_event(reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(child) if is_name(&namespace, child.name(), b"text") => {
                if saw_text || saw_mentions {
                    return Err(invalid(
                        "threaded-comment text is duplicated or out of order",
                    ));
                }
                saw_text = true;
                comment.text = Some(parse_text(reader)?);
            },
            Event::Empty(child) if is_name(&namespace, child.name(), b"text") => {
                if saw_text || saw_mentions {
                    return Err(invalid(
                        "threaded-comment text is duplicated or out of order",
                    ));
                }
                saw_text = true;
                comment.text = Some(String::new());
            },
            Event::Start(child) if is_name(&namespace, child.name(), b"mentions") => {
                if saw_mentions {
                    return Err(invalid("threaded-comment mentions element is duplicated"));
                }
                saw_mentions = true;
                parse_mentions(reader, &mut comment)?;
            },
            Event::Empty(child) if is_name(&namespace, child.name(), b"mentions") => {
                if saw_mentions {
                    return Err(invalid("threaded-comment mentions element is duplicated"));
                }
                saw_mentions = true;
            },
            Event::Start(child) => {
                comment
                    .extensions
                    .push(capture_unknown(reader, Event::Start(child))?);
            },
            Event::Empty(child) => {
                comment
                    .extensions
                    .push(capture_unknown(reader, Event::Empty(child))?);
            },
            Event::End(end) if is_name(&namespace, end.name(), b"threadedComment") => {
                return Ok(comment);
            },
            Event::Text(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid(
                        "threaded comment contains unexpected character data",
                    ));
                }
            },
            Event::CData(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid(
                        "threaded comment contains unexpected character data",
                    ));
                }
            },
            Event::End(_) => {
                return Err(invalid("threaded comment has an invalid closing element"));
            },
            Event::Eof => return Err(invalid("threaded comment is unterminated")),
            _ => {},
        }
    }
}

fn parse_comment_empty(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Comment> {
    Ok(Comment {
        cell_ref: unqualified_attribute_value(element, b"ref", decoder)
            .map_err(|error| Error::Encoding(error.to_string()))?,
        id: required_guid(element, b"id", decoder, "threaded-comment ID")?,
        parent_id: optional_guid(element, b"parentId", decoder, "threaded-comment parent ID")?,
        person_id: required_guid(element, b"personId", decoder, "threaded-comment person ID")?,
        text: None,
        date_time: unqualified_attribute_value(element, b"dT", decoder)
            .map_err(|error| Error::Encoding(error.to_string()))?,
        done: optional_bool(element, b"done", decoder, "threaded-comment done flag")?,
        mentions: Vec::new(),
        attributes: raw_attributes(
            element,
            &[b"ref", b"id", b"parentId", b"personId", b"dT", b"done"],
            decoder,
        )?,
        extensions: Vec::new(),
    })
}

fn parse_text(reader: &mut NsReader<&[u8]>) -> Result<String> {
    let mut text_value = String::new();
    loop {
        let event = next_event(reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Text(text) => {
                text_value.push_str(&text.decode().map_err(xml_error)?);
            },
            Event::CData(text) => {
                text_value.push_str(&text.decode().map_err(xml_error)?);
            },
            Event::GeneralRef(reference) => {
                text_value.push_str(&decode_xml_reference(&reference).map_err(xml_error)?);
            },
            Event::End(end) if is_name(&namespace, end.name(), b"text") => {
                return Ok(text_value);
            },
            Event::End(_) => return Err(invalid("text has an invalid closing element")),
            Event::Start(_) | Event::Empty(_) => {
                return Err(invalid("threaded-comment text contains nested XML"));
            },
            Event::Eof => return Err(invalid("threaded-comment text is unterminated")),
            _ => {},
        }
    }
}

fn parse_mentions(reader: &mut NsReader<&[u8]>, comment: &mut Comment) -> Result<()> {
    loop {
        let event = next_event(reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_name(&namespace, element.name(), b"mention") => {
                comment.mentions.push(parse_mention(reader, element)?);
            },
            Event::Empty(element) if is_name(&namespace, element.name(), b"mention") => {
                comment
                    .mentions
                    .push(parse_mention_empty(&element, reader.decoder())?);
            },
            Event::Start(element) => {
                comment
                    .extensions
                    .push(capture_unknown(reader, Event::Start(element))?);
            },
            Event::Empty(element) => {
                comment
                    .extensions
                    .push(capture_unknown(reader, Event::Empty(element))?);
            },
            Event::End(end) if is_name(&namespace, end.name(), b"mentions") => return Ok(()),
            Event::Text(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("mentions contains unexpected character data"));
                }
            },
            Event::CData(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("mentions contains unexpected character data"));
                }
            },
            Event::End(_) => return Err(invalid("mentions has an invalid closing element")),
            Event::Eof => return Err(invalid("mentions is unterminated")),
            _ => {},
        }
    }
}

fn parse_mention(reader: &mut NsReader<&[u8]>, element: BytesStart<'static>) -> Result<Mention> {
    let mut mention = parse_mention_empty(&element, reader.decoder())?;
    loop {
        let event = next_event(reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                mention
                    .extensions
                    .push(capture_unknown(reader, Event::Start(element))?);
            },
            Event::Empty(element) => {
                mention
                    .extensions
                    .push(capture_unknown(reader, Event::Empty(element))?);
            },
            Event::End(end) if is_name(&namespace, end.name(), b"mention") => return Ok(mention),
            Event::Text(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("mention contains unexpected character data"));
                }
            },
            Event::CData(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("mention contains unexpected character data"));
                }
            },
            Event::End(_) => return Err(invalid("mention has an invalid closing element")),
            Event::Eof => return Err(invalid("mention is unterminated")),
            _ => {},
        }
    }
}

fn parse_mention_empty(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Mention> {
    Ok(Mention {
        mention_person_id: required_guid(
            element,
            b"mentionpersonId",
            decoder,
            "mention person ID",
        )?,
        mention_id: required_guid(element, b"mentionId", decoder, "mention ID")?,
        start_index: required_u32(element, b"startIndex", decoder, "mention start index")?,
        length: required_u32(element, b"length", decoder, "mention length")?,
        attributes: raw_attributes(
            element,
            &[b"mentionpersonId", b"mentionId", b"startIndex", b"length"],
            decoder,
        )?,
        extensions: Vec::new(),
    })
}

fn write_person(xml: &mut String, person: &Person) -> Result<()> {
    xml.push_str("<person displayName=\"");
    xml.push_str(&escape_xml(&person.display_name));
    xml.push_str("\" id=\"");
    xml.push_str(&escape_xml(&person.id));
    xml.push('"');
    if let Some(value) = person.user_id.as_deref() {
        xml.push_str(" userId=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
    if let Some(value) = person.provider_id.as_deref() {
        xml.push_str(" providerId=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
    write_attributes(
        xml,
        &person.attributes,
        &[b"displayName", b"id", b"userId", b"providerId"],
    )?;
    if person.extensions.is_empty() {
        xml.push_str("/>");
    } else {
        xml.push('>');
        write_extensions(xml, &person.extensions)?;
        xml.push_str("</person>");
    }
    Ok(())
}

fn write_comment(xml: &mut String, comment: &Comment) -> Result<()> {
    xml.push_str("<threadedComment");
    if let Some(value) = comment.cell_ref.as_deref() {
        xml.push_str(" ref=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
    xml.push_str(" id=\"");
    xml.push_str(&escape_xml(&comment.id));
    xml.push_str("\" personId=\"");
    xml.push_str(&escape_xml(&comment.person_id));
    xml.push('"');
    if let Some(value) = comment.parent_id.as_deref() {
        xml.push_str(" parentId=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
    if let Some(value) = comment.date_time.as_deref() {
        xml.push_str(" dT=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
    if let Some(value) = comment.done {
        xml.push_str(" done=\"");
        xml.push_str(if value { "1" } else { "0" });
        xml.push('"');
    }
    write_attributes(
        xml,
        &comment.attributes,
        &[b"ref", b"id", b"personId", b"parentId", b"dT", b"done"],
    )?;
    if comment.text.is_none() && comment.mentions.is_empty() && comment.extensions.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    if let Some(value) = comment.text.as_deref() {
        xml.push_str("<text>");
        xml.push_str(&escape_xml(value));
        xml.push_str("</text>");
    }
    if !comment.mentions.is_empty() {
        xml.push_str("<mentions>");
        for mention in &comment.mentions {
            write_mention(xml, mention)?;
        }
        xml.push_str("</mentions>");
    }
    write_extensions(xml, &comment.extensions)?;
    xml.push_str("</threadedComment>");
    Ok(())
}

fn write_mention(xml: &mut String, mention: &Mention) -> Result<()> {
    xml.push_str("<mention mentionpersonId=\"");
    xml.push_str(&escape_xml(&mention.mention_person_id));
    xml.push_str("\" mentionId=\"");
    xml.push_str(&escape_xml(&mention.mention_id));
    let _ = write!(
        xml,
        "\" startIndex=\"{}\" length=\"{}\"",
        mention.start_index, mention.length
    );
    write_attributes(
        xml,
        &mention.attributes,
        &[b"mentionpersonId", b"mentionId", b"startIndex", b"length"],
    )?;
    if mention.extensions.is_empty() {
        xml.push_str("/>");
    } else {
        xml.push('>');
        write_extensions(xml, &mention.extensions)?;
        xml.push_str("</mention>");
    }
    Ok(())
}

fn write_attributes(xml: &mut String, attributes: &[RawAttribute], known: &[&[u8]]) -> Result<()> {
    for attribute in attributes {
        if known.iter().any(|name| *name == attribute.name.as_bytes()) {
            return Err(invalid(format!(
                "preserved XML attribute '{}' conflicts with a typed attribute",
                attribute.name
            )));
        }
        xml.push(' ');
        xml.push_str(&attribute.name);
        xml.push_str("=\"");
        xml.push_str(&escape_xml(&attribute.value));
        xml.push('"');
    }
    Ok(())
}

fn write_extensions(xml: &mut String, extensions: &[RawXml]) -> Result<()> {
    let total = extensions.iter().try_fold(0usize, |total, extension| {
        total
            .checked_add(extension.bytes.len())
            .ok_or_else(|| invalid("preserved XML extension size overflows"))
    })?;
    if extensions.len() > MAX_EXTENSIONS || total > MAX_EXTENSION_BYTES {
        return Err(invalid(
            "preserved XML extensions exceed the resource bound",
        ));
    }
    for extension in extensions {
        let value = std::str::from_utf8(&extension.bytes)
            .map_err(|error| Error::Encoding(error.to_string()))?;
        xml.push_str(value);
    }
    Ok(())
}

fn next_event(reader: &mut NsReader<&[u8]>) -> Result<Event<'static>> {
    reader
        .read_event()
        .map(Event::into_owned)
        .map_err(xml_error)
}

fn capture_unknown(reader: &mut NsReader<&[u8]>, first: Event<'static>) -> Result<RawXml> {
    let mut writer = quick_xml::Writer::new(Vec::new());
    let mut depth = 0usize;
    if matches!(first, Event::Start(_)) {
        depth = 1;
    }
    writer.write_event(first).map_err(xml_error)?;
    while depth != 0 {
        let event = next_event(reader)?;
        match &event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("unknown XML nesting overflows"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(invalid("unknown XML nesting exceeds the resource bound"));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unknown XML closing element is unexpected"))?;
            },
            Event::Eof => return Err(invalid("unknown XML element is unterminated")),
            _ => {},
        }
        writer.write_event(event).map_err(xml_error)?;
    }
    let bytes = writer.into_inner();
    if bytes.len() > MAX_EXTENSION_BYTES {
        return Err(invalid("unknown XML payload exceeds the resource bound"));
    }
    Ok(RawXml::new(bytes))
}

fn raw_attributes(
    element: &BytesStart<'_>,
    known: &[&[u8]],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<RawAttribute>> {
    let mut attributes = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Encoding(error.to_string()))?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || known.iter().any(|item| *item == name) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Encoding(error.to_string()))?
            .into_owned();
        let name = std::str::from_utf8(name)
            .map_err(|error| Error::Encoding(error.to_string()))?
            .to_owned();
        attributes.push(RawAttribute { name, value });
    }
    Ok(attributes)
}

fn require_root(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local: &[u8],
    closed: bool,
) -> Result<()> {
    if closed || !is_name(namespace, name, local) {
        return Err(invalid(format!(
            "part must have one '{}' root in the threaded-comments namespace",
            String::from_utf8_lossy(local)
        )));
    }
    Ok(())
}

fn is_name(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == NAMESPACE)
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<String> {
    let value = unqualified_attribute_value(element, name, decoder)
        .map_err(|error| Error::Encoding(error.to_string()))?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))?;
    if value.is_empty() {
        return Err(invalid(format!("{description} cannot be empty")));
    }
    Ok(value)
}

fn required_guid(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<String> {
    let value = required_string(element, name, decoder, description)?;
    if !valid_guid(&value) {
        return Err(invalid(format!("invalid {description} GUID '{value}'")));
    }
    Ok(value)
}

fn optional_guid(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<Option<String>> {
    let Some(value) = unqualified_attribute_value(element, name, decoder)
        .map_err(|error| Error::Encoding(error.to_string()))?
    else {
        return Ok(None);
    };
    if !valid_guid(&value) {
        return Err(invalid(format!("invalid {description} GUID '{value}'")));
    }
    Ok(Some(value))
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<u32> {
    let value = required_string(element, name, decoder, description)?;
    value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid {description} '{value}'")))
}

fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    description: &str,
) -> Result<Option<bool>> {
    let Some(value) = unqualified_attribute_value(element, name, decoder)
        .map_err(|error| Error::Encoding(error.to_string()))?
    else {
        return Ok(None);
    };
    match value.as_str() {
        "1" | "true" => Ok(Some(true)),
        "0" | "false" => Ok(Some(false)),
        _ => Err(invalid(format!("invalid {description} '{value}'"))),
    }
}

fn check_size(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_PART_BYTES {
        Err(invalid("threaded-comments part exceeds the resource bound"))
    } else {
        Ok(())
    }
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Encoding(error.to_string())
}

fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}
