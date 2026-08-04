//! Bounded SpreadsheetML threaded-comments model and XML codec.

use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

use chrono::{DateTime, NaiveDateTime};
use litchi_core::sheet::Result as SheetResult;
use litchi_core::xml::escape_xml;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Comment, Comments, Graph, Mention, People, Person};
use super::{
    MAX_COMMENTS, MAX_IDENTITY_BYTES, MAX_MENTIONS, MAX_PART_BYTES, MAX_PERSONS, MAX_TEXT_UTF16,
};
use litchi_ooxml_common::custom_xml::valid_guid;
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};

const THREADED_COMMENTS_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// Validate an optional threaded-comment timestamp.
pub fn validate_timestamp(value: Option<&str>) -> SheetResult<()> {
    let Some(value) = value else {
        return Ok(());
    };
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(format!("invalid threaded-comment timestamp '{value}'").into())
    }
}

fn validate_cell_ref(reference: &str) -> SheetResult<()> {
    let bytes = reference.as_bytes();
    let column_end = bytes
        .iter()
        .position(u8::is_ascii_digit)
        .ok_or_else(|| format!("Invalid cell reference: {reference}"))?;
    if column_end == 0 || column_end == bytes.len() {
        return Err(format!("Invalid cell reference: {reference}").into());
    }

    let mut column = 0u32;
    for byte in &bytes[..column_end] {
        if !byte.is_ascii_alphabetic() {
            return Err(format!("Invalid column in cell reference: {reference}").into());
        }
        let digit = u32::from(byte.to_ascii_uppercase() - b'A' + 1);
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(digit))
            .ok_or_else(|| format!("Column overflows in cell reference: {reference}"))?;
    }
    if column > 16_384 {
        return Err(format!("Column exceeds Excel limits in cell reference: {reference}").into());
    }

    let row = bytes[column_end..]
        .iter()
        .all(u8::is_ascii_digit)
        .then(|| std::str::from_utf8(&bytes[column_end..]))
        .transpose()?
        .ok_or_else(|| format!("Invalid row in cell reference: {reference}"))?
        .parse::<u32>()
        .map_err(|_| format!("Invalid row number in cell reference: {reference}"))?;
    if row == 0 || row > 1_048_576 {
        return Err(format!("Row exceeds Excel limits in cell reference: {reference}").into());
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PersonContext {
    PersonList,
    Person,
    Other,
}

pub fn parse_persons(xml: &str) -> SheetResult<People> {
    if xml.len() > MAX_PART_BYTES {
        return Err("persons part exceeds the configured resource bound".into());
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut stack = Vec::new();
    let mut persons = Vec::new();
    let mut ids = HashSet::new();
    let mut closed_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if stack.is_empty() => {
                validate_root(
                    closed_root,
                    &namespace,
                    element.name(),
                    b"personList",
                    "people",
                )?;
                stack.push(PersonContext::PersonList);
            },
            Event::Empty(element) if stack.is_empty() => {
                validate_root(
                    closed_root,
                    &namespace,
                    element.name(),
                    b"personList",
                    "people",
                )?;
                closed_root = true;
            },
            Event::Start(element) => {
                let parent = *stack.last().ok_or("people part is missing its root")?;
                stack.push(start_person_element(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut persons,
                    &mut ids,
                )?);
            },
            Event::Empty(element) => {
                let parent = *stack.last().ok_or("people part is missing its root")?;
                start_person_element(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut persons,
                    &mut ids,
                )?;
            },
            Event::End(element) => {
                let context = stack
                    .pop()
                    .ok_or("people part has a closing element outside its root")?;
                if context == PersonContext::PersonList {
                    if !is_threaded_name(&namespace, element.name(), b"personList") {
                        return Err("people part has an invalid root closing element".into());
                    }
                    closed_root = true;
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => {
                return Err("people part has a missing or unterminated root".into());
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if persons.len() > MAX_PERSONS {
        return Err("persons part contains too many people".into());
    }
    Ok(People { persons })
}

fn start_person_element(
    parent: PersonContext,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    persons: &mut Vec<Person>,
    ids: &mut HashSet<String>,
) -> SheetResult<PersonContext> {
    if parent != PersonContext::PersonList
        || !is_threaded_name(namespace, element.name(), b"person")
    {
        return Ok(PersonContext::Other);
    }
    let person = parse_person(element, decoder)?;
    if !ids.insert(person.id.clone()) {
        return Err(format!("duplicate person ID '{}'", person.id).into());
    }
    persons.push(person);
    Ok(PersonContext::Person)
}

fn parse_person(element: &BytesStart<'_>, decoder: Decoder) -> SheetResult<Person> {
    Ok(Person {
        display_name: required_string(element, b"displayName", decoder, "person display name")?,
        id: required_guid(element, b"id", decoder, "person ID")?,
        user_id: unqualified_attribute_value(element, b"userId", decoder)?,
        provider_id: unqualified_attribute_value(element, b"providerId", decoder)?,
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum CommentContext {
    ThreadedComments,
    Comment,
    Text,
    Mentions,
    Mention,
    Other,
}

struct PendingComment {
    comment: Comment,
    saw_text: bool,
    saw_mentions: bool,
}

pub fn parse_comments(xml: &str) -> SheetResult<Comments> {
    if xml.len() > MAX_PART_BYTES {
        return Err("threaded-comments part exceeds the configured resource bound".into());
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut stack = Vec::new();
    let mut comments = Vec::new();
    let mut comment_ids = HashSet::new();
    let mut mention_ids = HashSet::new();
    let mut pending: Option<PendingComment> = None;
    let mut closed_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if stack.is_empty() => {
                validate_root(
                    closed_root,
                    &namespace,
                    element.name(),
                    b"ThreadedComments",
                    "threaded comments",
                )?;
                stack.push(CommentContext::ThreadedComments);
            },
            Event::Empty(element) if stack.is_empty() => {
                validate_root(
                    closed_root,
                    &namespace,
                    element.name(),
                    b"ThreadedComments",
                    "threaded comments",
                )?;
                closed_root = true;
            },
            Event::Start(element) => {
                let parent = *stack
                    .last()
                    .ok_or("threaded-comments part is missing its root")?;
                let context = start_comment_element(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut pending,
                    &mut comment_ids,
                    &mut mention_ids,
                )?;
                stack.push(context);
            },
            Event::Empty(element) => {
                let parent = *stack
                    .last()
                    .ok_or("threaded-comments part is missing its root")?;
                let context = start_comment_element(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut pending,
                    &mut comment_ids,
                    &mut mention_ids,
                )?;
                finish_comment_element(context, &mut pending, &mut comments)?;
            },
            Event::Text(text) if stack.last() == Some(&CommentContext::Text) => {
                pending
                    .as_mut()
                    .ok_or("threaded-comment text outside a comment")?
                    .comment
                    .text
                    .get_or_insert_with(String::new)
                    .push_str(&text.decode()?);
            },
            Event::CData(text) if stack.last() == Some(&CommentContext::Text) => {
                pending
                    .as_mut()
                    .ok_or("threaded-comment text outside a comment")?
                    .comment
                    .text
                    .get_or_insert_with(String::new)
                    .push_str(&text.decode()?);
            },
            Event::GeneralRef(reference) if stack.last() == Some(&CommentContext::Text) => {
                pending
                    .as_mut()
                    .ok_or("threaded-comment text outside a comment")?
                    .comment
                    .text
                    .get_or_insert_with(String::new)
                    .push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                let context = stack
                    .pop()
                    .ok_or("threaded-comments part has a closing element outside its root")?;
                finish_comment_element(context, &mut pending, &mut comments)?;
                if context == CommentContext::ThreadedComments {
                    if !is_threaded_name(&namespace, element.name(), b"ThreadedComments") {
                        return Err(
                            "threaded-comments part has an invalid root closing element".into()
                        );
                    }
                    closed_root = true;
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => {
                return Err("threaded-comments part has a missing or unterminated root".into());
            },
            Event::Eof => break,
            _ => {},
        }
    }

    for comment in &comments {
        if let Some(parent_id) = comment.parent_id.as_deref()
            && (!comment_ids.contains(parent_id) || parent_id == comment.id)
        {
            return Err(format!(
                "threaded comment '{}' has invalid parent ID '{parent_id}'",
                comment.id
            )
            .into());
        }
    }
    if comments.len() > MAX_COMMENTS {
        return Err("threaded-comment part contains too many comments".into());
    }
    let mentions = comments
        .iter()
        .map(|comment| comment.mentions.len())
        .sum::<usize>();
    if mentions > MAX_MENTIONS {
        return Err("threaded-comments part contains too many mentions".into());
    }
    if comments.iter().any(|comment| {
        comment
            .text
            .as_deref()
            .is_some_and(|text| text.encode_utf16().count() > MAX_TEXT_UTF16)
    }) {
        return Err("threaded-comment text exceeds the configured resource bound".into());
    }
    Ok(Comments { comments })
}

fn start_comment_element(
    parent: CommentContext,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    pending: &mut Option<PendingComment>,
    comment_ids: &mut HashSet<String>,
    mention_ids: &mut HashSet<String>,
) -> SheetResult<CommentContext> {
    if parent == CommentContext::ThreadedComments
        && is_threaded_name(namespace, element.name(), b"threadedComment")
    {
        if pending.is_some() {
            return Err("nested threaded comment".into());
        }
        let comment = parse_threaded_comment(element, decoder)?;
        if !comment_ids.insert(comment.id.clone()) {
            return Err(format!("duplicate threaded-comment ID '{}'", comment.id).into());
        }
        *pending = Some(PendingComment {
            comment,
            saw_text: false,
            saw_mentions: false,
        });
        return Ok(CommentContext::Comment);
    }
    if parent == CommentContext::Comment && is_threaded_name(namespace, element.name(), b"text") {
        let pending = pending
            .as_mut()
            .ok_or("threaded-comment text outside a comment")?;
        if pending.saw_text || pending.saw_mentions {
            return Err("duplicate or out-of-order threaded-comment text".into());
        }
        pending.saw_text = true;
        pending.comment.text = Some(String::new());
        return Ok(CommentContext::Text);
    }
    if parent == CommentContext::Comment && is_threaded_name(namespace, element.name(), b"mentions")
    {
        let pending = pending
            .as_mut()
            .ok_or("threaded-comment mentions outside a comment")?;
        if pending.saw_mentions {
            return Err("duplicate threaded-comment mentions element".into());
        }
        pending.saw_mentions = true;
        return Ok(CommentContext::Mentions);
    }
    if parent == CommentContext::Mentions && is_threaded_name(namespace, element.name(), b"mention")
    {
        let mention = parse_mention(element, decoder)?;
        if !mention_ids.insert(mention.mention_id.clone()) {
            return Err(format!("duplicate mention ID '{}'", mention.mention_id).into());
        }
        pending
            .as_mut()
            .ok_or("mention outside a threaded comment")?
            .comment
            .mentions
            .push(mention);
        return Ok(CommentContext::Mention);
    }
    Ok(CommentContext::Other)
}

fn finish_comment_element(
    context: CommentContext,
    pending: &mut Option<PendingComment>,
    comments: &mut Vec<Comment>,
) -> SheetResult<()> {
    if context != CommentContext::Comment {
        return Ok(());
    }
    let pending = pending.take().ok_or("missing pending threaded comment")?;
    if let Some(text) = pending.comment.text.as_deref() {
        let text_len = u32::try_from(text.encode_utf16().count())
            .map_err(|_| "threaded-comment text is too long")?;
        for mention in &pending.comment.mentions {
            let end = mention
                .start_index
                .checked_add(mention.length)
                .ok_or("mention range overflows")?;
            if end > text_len {
                return Err(format!(
                    "mention '{}' range exceeds threaded-comment text",
                    mention.mention_id
                )
                .into());
            }
        }
    }
    comments.push(pending.comment);
    Ok(())
}

fn parse_threaded_comment(element: &BytesStart<'_>, decoder: Decoder) -> SheetResult<Comment> {
    let cell_ref = unqualified_attribute_value(element, b"ref", decoder)?;
    if let Some(cell_ref) = cell_ref.as_deref() {
        validate_cell_ref(cell_ref)?;
    }
    let date_time = unqualified_attribute_value(element, b"dT", decoder)?;
    validate_timestamp(date_time.as_deref())?;
    Ok(Comment {
        cell_ref,
        id: required_guid(element, b"id", decoder, "threaded-comment ID")?,
        parent_id: optional_guid(element, b"parentId", decoder, "threaded-comment parent ID")?,
        person_id: required_guid(element, b"personId", decoder, "threaded-comment person ID")?,
        text: None,
        date_time,
        done: optional_bool(element, b"done", decoder, "threaded-comment done flag")?,
        mentions: Vec::new(),
    })
}

fn parse_mention(element: &BytesStart<'_>, decoder: Decoder) -> SheetResult<Mention> {
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
    })
}

fn validate_root(
    closed_root: bool,
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    description: &str,
) -> SheetResult<()> {
    if closed_root || !is_threaded_name(namespace, name, local_name) {
        return Err(format!(
            "{description} part must have one namespaced '{}' root",
            String::from_utf8_lossy(local_name)
        )
        .into());
    }
    Ok(())
}

fn is_threaded_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if *value == THREADED_COMMENTS_NAMESPACE
        )
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<String> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| format!("missing {description} attribute"))?;
    if value.is_empty() {
        return Err(format!("{description} cannot be empty").into());
    }
    Ok(value)
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<u32> {
    let value = required_string(element, name, decoder, description)?;
    value
        .parse::<u32>()
        .map_err(|_| format!("invalid {description} '{value}'").into())
}

fn required_guid(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<String> {
    let value = required_string(element, name, decoder, description)?;
    validate_guid(&value, description)?;
    Ok(value)
}

fn optional_guid(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<String>> {
    let Some(value) = unqualified_attribute_value(element, name, decoder)? else {
        return Ok(None);
    };
    validate_guid(&value, description)?;
    Ok(Some(value))
}

pub fn validate_guid(value: &str, description: &str) -> SheetResult<()> {
    if !valid_guid(value) {
        return Err(format!("invalid {description} GUID '{value}'").into());
    }
    Ok(())
}

fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<bool>> {
    let Some(value) = unqualified_attribute_value(element, name, decoder)? else {
        return Ok(None);
    };
    match value.as_str() {
        "1" | "true" => Ok(Some(true)),
        "0" | "false" => Ok(Some(false)),
        _ => Err(format!("invalid {description} '{value}'").into()),
    }
}

const XML_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const THREADED_COMMENTS_NS: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// Write person list to XML.
///
/// Generates the `xl/persons/person.xml` part containing all persons
/// who can author threaded comments in the workbook.
pub fn write_persons(person_list: &People) -> SheetResult<String> {
    validate_people(person_list)?;
    let mut xml = String::with_capacity(1024);

    xml.push_str(XML_HEADER);
    xml.push('\n');
    write!(
        &mut xml,
        r#"<personList xmlns="{}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        THREADED_COMMENTS_NS
    )?;

    for person in &person_list.persons {
        write_person(&mut xml, person)?;
    }

    xml.push_str("</personList>");
    Ok(xml)
}

/// Write a single person to XML.
fn write_person(xml: &mut String, person: &Person) -> SheetResult<()> {
    write!(
        xml,
        r#"<person displayName="{}" id="{}""#,
        escape_xml(&person.display_name),
        escape_xml(&person.id)
    )?;

    if let Some(user_id) = &person.user_id {
        write!(xml, r#" userId="{}""#, escape_xml(user_id))?;
    }
    if let Some(provider_id) = &person.provider_id {
        write!(xml, r#" providerId="{}""#, escape_xml(provider_id))?;
    }

    xml.push_str("/>");
    Ok(())
}

/// Write threaded comments to XML.
///
/// Generates the `xl/threadedComments/threadedCommentN.xml` part containing
/// all threaded comments for a specific worksheet.
pub fn write_comments(comments: &Comments) -> SheetResult<String> {
    validate_comments(comments)?;
    let mut xml = String::with_capacity(4096);

    xml.push_str(XML_HEADER);
    xml.push('\n');
    write!(
        &mut xml,
        r#"<ThreadedComments xmlns="{}">"#,
        THREADED_COMMENTS_NS
    )?;

    for comment in &comments.comments {
        write_threaded_comment(&mut xml, comment)?;
    }

    xml.push_str("</ThreadedComments>");
    Ok(xml)
}

/// Write a single threaded comment to XML.
fn write_threaded_comment(xml: &mut String, comment: &Comment) -> SheetResult<()> {
    xml.push_str("<threadedComment");

    if let Some(cell_ref) = &comment.cell_ref {
        write!(xml, r#" ref="{}""#, escape_xml(cell_ref))?;
    }

    write!(
        xml,
        r#" id="{}" personId="{}""#,
        escape_xml(&comment.id),
        escape_xml(&comment.person_id)
    )?;

    if let Some(parent_id) = &comment.parent_id {
        write!(xml, r#" parentId="{}""#, escape_xml(parent_id))?;
    }

    if let Some(date_time) = &comment.date_time {
        write!(xml, r#" dT="{}""#, escape_xml(date_time))?;
    }

    if let Some(done) = comment.done {
        write!(xml, r#" done="{}""#, if done { "1" } else { "0" })?;
    }

    if comment.text.is_none() && comment.mentions.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }

    xml.push('>');

    if let Some(text) = &comment.text {
        write!(xml, "<text>{}</text>", escape_xml(text))?;
    }

    if !comment.mentions.is_empty() {
        write_mentions(xml, &comment.mentions)?;
    }

    xml.push_str("</threadedComment>");
    Ok(())
}

/// Write mentions to XML.
fn write_mentions(xml: &mut String, mentions: &[Mention]) -> SheetResult<()> {
    xml.push_str("<mentions>");

    for mention in mentions {
        write!(
            xml,
            r#"<mention mentionpersonId="{}" mentionId="{}" startIndex="{}" length="{}"/>"#,
            escape_xml(&mention.mention_person_id),
            escape_xml(&mention.mention_id),
            mention.start_index,
            mention.length
        )?;
    }

    xml.push_str("</mentions>");
    Ok(())
}

pub fn validate_people(person_list: &People) -> SheetResult<()> {
    if person_list.persons.len() > MAX_PERSONS {
        return Err("persons list contains too many people".into());
    }
    let mut ids = HashSet::with_capacity(person_list.persons.len());
    for person in &person_list.persons {
        validate_guid(&person.id, "person ID")?;
        if person.display_name.len() > MAX_IDENTITY_BYTES
            || person
                .user_id
                .as_ref()
                .is_some_and(|value| value.len() > MAX_IDENTITY_BYTES)
            || person
                .provider_id
                .as_ref()
                .is_some_and(|value| value.len() > MAX_IDENTITY_BYTES)
        {
            return Err(format!("person '{}' has oversized identity metadata", person.id).into());
        }
        if !ids.insert(person.id.as_str()) {
            return Err(format!("duplicate person ID '{}'", person.id).into());
        }
    }
    Ok(())
}

pub fn validate_comments(comments: &Comments) -> SheetResult<()> {
    if comments.comments.len() > MAX_COMMENTS {
        return Err("threaded-comments list contains too many comments".into());
    }
    let mut comment_ids = HashSet::with_capacity(comments.comments.len());
    let mention_count = comments
        .comments
        .iter()
        .map(|comment| comment.mentions.len())
        .sum();
    if mention_count > MAX_MENTIONS {
        return Err("threaded-comments list contains too many mentions".into());
    }
    let mut mention_ids = HashSet::with_capacity(mention_count);

    for comment in &comments.comments {
        validate_guid(&comment.id, "threaded-comment ID")?;
        validate_guid(&comment.person_id, "threaded-comment person ID")?;
        if let Some(parent_id) = comment.parent_id.as_deref() {
            validate_guid(parent_id, "threaded-comment parent ID")?;
        }
        if let Some(cell_ref) = comment.cell_ref.as_deref() {
            validate_cell_ref(cell_ref)?;
        }
        validate_timestamp(comment.date_time.as_deref())?;
        if !comment_ids.insert(comment.id.as_str()) {
            return Err(format!("duplicate threaded-comment ID '{}'", comment.id).into());
        }

        let text_len = comment
            .text
            .as_deref()
            .map(|text| {
                u32::try_from(text.encode_utf16().count())
                    .map_err(|_| "threaded-comment text is too long")
            })
            .transpose()?;
        if text_len.is_some_and(|length| length as usize > MAX_TEXT_UTF16) {
            return Err(format!("threaded-comment text '{}' is too long", comment.id).into());
        }
        for mention in &comment.mentions {
            validate_guid(&mention.mention_person_id, "mention person ID")?;
            validate_guid(&mention.mention_id, "mention ID")?;
            if !mention_ids.insert(mention.mention_id.as_str()) {
                return Err(format!("duplicate mention ID '{}'", mention.mention_id).into());
            }
            if let Some(text_len) = text_len {
                let end = mention
                    .start_index
                    .checked_add(mention.length)
                    .ok_or("mention range overflows")?;
                if end > text_len {
                    return Err(format!(
                        "mention '{}' range exceeds threaded-comment text",
                        mention.mention_id
                    )
                    .into());
                }
            }
        }
    }

    for comment in &comments.comments {
        if let Some(parent_id) = comment.parent_id.as_deref()
            && (!comment_ids.contains(parent_id) || parent_id == comment.id)
        {
            return Err(format!(
                "threaded comment '{}' has invalid parent ID '{parent_id}'",
                comment.id
            )
            .into());
        }
    }
    Ok(())
}

pub fn validate_graph(graph: &Graph) -> SheetResult<()> {
    let empty_persons = People::default();
    let persons = graph
        .persons
        .as_ref()
        .map(|part| &part.persons)
        .unwrap_or(&empty_persons);
    write_persons(persons)?;
    let person_ids: HashSet<&str> = persons
        .persons
        .iter()
        .map(|person| person.id.as_str())
        .collect();
    let mut comment_ids = HashSet::new();
    let mut mention_ids = HashSet::new();
    for sheet in &graph.worksheets {
        write_comments(&sheet.comments)?;
        let mut root_cells = HashSet::new();
        for comment in &sheet.comments.comments {
            if !comment_ids.insert(comment.id.as_str()) {
                return Err(
                    format!("duplicate workbook threaded-comment ID '{}'", comment.id).into(),
                );
            }
            if !person_ids.contains(comment.person_id.as_str()) {
                return Err(format!(
                    "threaded comment '{}' references missing person '{}'",
                    comment.id, comment.person_id
                )
                .into());
            }
            validate_timestamp(comment.date_time.as_deref())?;
            if comment.parent_id.is_none() {
                let cell = comment.cell_ref.as_deref().ok_or_else(|| {
                    format!(
                        "threaded-comment root '{}' is missing its cell reference",
                        comment.id
                    )
                })?;
                if !root_cells.insert(cell) {
                    return Err(
                        format!("worksheet has multiple threaded-comment roots at {cell}").into(),
                    );
                }
            } else if comment.cell_ref.is_some() {
                return Err(format!(
                    "threaded-comment reply '{}' must not carry a cell reference",
                    comment.id
                )
                .into());
            }
            for mention in &comment.mentions {
                if !person_ids.contains(mention.mention_person_id.as_str()) {
                    return Err(format!(
                        "mention '{}' references missing person '{}'",
                        mention.mention_id, mention.mention_person_id
                    )
                    .into());
                }
                if !mention_ids.insert(mention.mention_id.as_str()) {
                    return Err(
                        format!("duplicate workbook mention ID '{}'", mention.mention_id).into(),
                    );
                }
            }
        }
        let roots: HashSet<&str> = sheet
            .comments
            .comments
            .iter()
            .filter(|comment| comment.parent_id.is_none())
            .map(|comment| comment.id.as_str())
            .collect();
        for reply in sheet
            .comments
            .comments
            .iter()
            .filter(|comment| comment.parent_id.is_some())
        {
            if !roots.contains(reply.parent_id.as_deref().expect("filtered")) {
                return Err(format!(
                    "threaded-comment reply '{}' must reference a root in the same worksheet",
                    reply.id
                )
                .into());
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod reader_tests {
    use super::*;

    const NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
    const PERSON_ID: &str = "{11111111-1111-1111-1111-111111111111}";
    const COMMENT_ID: &str = "{22222222-2222-2222-2222-222222222222}";
    const BOB_ID: &str = "{33333333-3333-3333-3333-333333333333}";
    const MENTION_ID: &str = "{44444444-4444-4444-4444-444444444444}";

    #[test]
    fn parses_prefixed_people_and_threaded_comments() {
        let people = parse_persons(&format!(
            r#"<tc:personList xmlns:tc="{NS}" xmlns:f="urn:foreign">
                <f:person displayName="Ignored" id="ignored"/>
                <tc:person displayName="Alice &amp; Bob" id="{PERSON_ID}" userId="alice" providerId="aad"/>
            </tc:personList>"#
        ))
        .unwrap();
        assert_eq!(people.persons.len(), 1);
        assert_eq!(people.persons[0].display_name, "Alice & Bob");

        let comments = parse_comments(&format!(
            r#"<tc:ThreadedComments xmlns:tc="{NS}" xmlns:f="urn:foreign">
                <f:threadedComment id="ignored" personId="ignored"/>
                <tc:threadedComment ref="B2" id="{COMMENT_ID}" personId="{PERSON_ID}" dT="2026-07-14T10:00:00Z" done="true">
                    <tc:text>Hello &amp; @Bob</tc:text>
                    <tc:mentions><tc:mention mentionpersonId="{BOB_ID}" mentionId="{MENTION_ID}" startIndex="8" length="4"/></tc:mentions>
                </tc:threadedComment>
            </tc:ThreadedComments>"#
        ))
        .unwrap();
        let comment = &comments.comments[0];
        assert_eq!(comment.cell_ref.as_deref(), Some("B2"));
        assert_eq!(comment.text.as_deref(), Some("Hello & @Bob"));
        assert_eq!(comment.done, Some(true));
        assert_eq!(comment.mentions.len(), 1);
    }

    #[test]
    fn accepts_empty_present_parts() {
        let people = parse_persons(&format!(r#"<personList xmlns="{NS}"/>"#)).unwrap();
        let comments = parse_comments(&format!(r#"<ThreadedComments xmlns="{NS}"/>"#)).unwrap();

        assert!(people.persons.is_empty());
        assert!(comments.comments.is_empty());
    }

    #[test]
    fn rejects_malformed_threaded_parts() {
        let invalid_people = [
            r#"<personList xmlns="urn:foreign"/>"#.to_string(),
            format!(r#"<personList xmlns="{NS}"><person id="x"/></personList>"#),
            format!(
                r#"<personList xmlns="{NS}"><person displayName="A" id="{PERSON_ID}"/><person displayName="B" id="{PERSON_ID}"/></personList>"#
            ),
        ];
        for xml in invalid_people {
            assert!(parse_persons(&xml).is_err(), "accepted {xml}");
        }

        let invalid_comments = [
            r#"<ThreadedComments xmlns="urn:foreign"/>"#.to_string(),
            format!(
                r#"<ThreadedComments xmlns="{NS}"><threadedComment personId="p"/></ThreadedComments>"#
            ),
            format!(
                r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}" done="yes"/></ThreadedComments>"#
            ),
            format!(
                r#"<ThreadedComments xmlns="{NS}"><threadedComment ref="A0" id="{COMMENT_ID}" personId="{PERSON_ID}"/></ThreadedComments>"#
            ),
            format!(
                r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}"><mentions/><text>x</text></threadedComment></ThreadedComments>"#
            ),
            format!(
                r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}"><text>x</text><text>y</text></threadedComment></ThreadedComments>"#
            ),
            format!(
                r#"<ThreadedComments xmlns="{NS}"><threadedComment id="{COMMENT_ID}" personId="{PERSON_ID}"><text>x</text><mentions><mention mentionpersonId="{PERSON_ID}" mentionId="{MENTION_ID}" startIndex="1" length="1"/></mentions></threadedComment></ThreadedComments>"#
            ),
        ];
        for xml in invalid_comments {
            assert!(parse_comments(&xml).is_err(), "accepted {xml}");
        }
    }
}

#[cfg(test)]
mod writer_tests {
    use super::*;

    const PERSON_ID: &str = "{11111111-1111-1111-1111-111111111111}";
    const COMMENT_ID: &str = "{22222222-2222-2222-2222-222222222222}";
    const REPLY_ID: &str = "{33333333-3333-3333-3333-333333333333}";
    const MENTION_ID: &str = "{44444444-4444-4444-4444-444444444444}";

    #[test]
    fn writes_schema_valid_people_and_comments() {
        let people = People {
            persons: vec![Person {
                display_name: "Alice & Bob".into(),
                id: PERSON_ID.into(),
                user_id: Some("alice@example.com".into()),
                provider_id: None,
            }],
        };
        let people_xml = write_persons(&people).unwrap();
        assert!(people_xml.contains("Alice &amp; Bob"));

        let comments = Comments {
            comments: vec![
                Comment {
                    cell_ref: Some("A1".into()),
                    id: COMMENT_ID.into(),
                    person_id: PERSON_ID.into(),
                    text: Some("Hi @Bob".into()),
                    mentions: vec![Mention {
                        mention_person_id: PERSON_ID.into(),
                        mention_id: MENTION_ID.into(),
                        start_index: 3,
                        length: 4,
                    }],
                    ..Default::default()
                },
                Comment {
                    id: REPLY_ID.into(),
                    person_id: PERSON_ID.into(),
                    parent_id: Some(COMMENT_ID.into()),
                    ..Default::default()
                },
            ],
        };
        let comments_xml = write_comments(&comments).unwrap();
        assert!(comments_xml.contains("<text>Hi @Bob</text>"));
        assert!(comments_xml.contains(&format!(r#" parentId="{COMMENT_ID}""#)));
    }

    #[test]
    fn rejects_invalid_people_and_comments() {
        let duplicate_people = People {
            persons: vec![
                Person {
                    id: PERSON_ID.into(),
                    ..Default::default()
                },
                Person {
                    id: PERSON_ID.into(),
                    ..Default::default()
                },
            ],
        };
        assert!(write_persons(&duplicate_people).is_err());

        let invalid = [
            Comment {
                id: "not-a-guid".into(),
                person_id: PERSON_ID.into(),
                ..Default::default()
            },
            Comment {
                cell_ref: Some("A0".into()),
                id: COMMENT_ID.into(),
                person_id: PERSON_ID.into(),
                ..Default::default()
            },
            Comment {
                id: COMMENT_ID.into(),
                person_id: PERSON_ID.into(),
                parent_id: Some(REPLY_ID.into()),
                ..Default::default()
            },
            Comment {
                id: COMMENT_ID.into(),
                person_id: PERSON_ID.into(),
                text: Some("x".into()),
                mentions: vec![Mention {
                    mention_person_id: PERSON_ID.into(),
                    mention_id: MENTION_ID.into(),
                    start_index: 1,
                    length: 1,
                }],
                ..Default::default()
            },
        ];
        for comment in invalid {
            assert!(
                write_comments(&Comments {
                    comments: vec![comment]
                })
                .is_err()
            );
        }
    }
}
