//! Namespace-aware readers for threaded comments and people parts.

use std::collections::HashSet;

use litchi_core::sheet::Result as SheetResult;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Relationships};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::person::{Mention, Person, PersonList};
use super::{
    MAX_THREADED_COMMENTS, MAX_THREADED_MENTIONS, MAX_THREADED_PART_BYTES,
    MAX_THREADED_PERSONS, MAX_THREADED_TEXT_UTF16, ThreadedComment, ThreadedComments,
    validate_threaded_timestamp,
};
use crate::common::xml::{decode_xml_reference, unqualified_attribute_value};
use crate::xlsx::Cell;

const THREADED_COMMENTS_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// Read the person list related to the workbook's actual main-document part.
pub fn read_persons(package: &OpcPackage) -> SheetResult<Option<PersonList>> {
    let workbook_part = package.main_document_part()?;
    let Some(persons_uri) = related_part_uri(workbook_part.rels(), rt::PERSONS, "people")? else {
        return Ok(None);
    };
    let persons_part = package.get_part(&persons_uri)?;
    require_content_type(&persons_uri, persons_part.content_type(), ct::SML_PERSONS)?;
    if persons_part.blob().len() > MAX_THREADED_PART_BYTES {
        return Err("persons part exceeds the configured resource bound".into());
    }
    let bytes = crate::common::mce::process_part(persons_part)?;
    let xml = std::str::from_utf8(bytes.as_ref())?;
    let persons = parse_person_list(xml)?;
    if persons.persons.len() > MAX_THREADED_PERSONS {
        return Err("persons part contains too many people".into());
    }
    Ok(Some(persons))
}

/// Read the threaded-comments part related to a worksheet.
pub fn read_threaded_comments(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
) -> SheetResult<Option<ThreadedComments>> {
    let worksheet_part = package.get_part(worksheet_uri)?;
    let Some(comments_uri) = related_part_uri(
        worksheet_part.rels(),
        rt::THREADED_COMMENTS,
        "threaded comments",
    )?
    else {
        return Ok(None);
    };
    let comments_part = package.get_part(&comments_uri)?;
    require_content_type(
        &comments_uri,
        comments_part.content_type(),
        ct::SML_THREADED_COMMENTS,
    )?;
    if comments_part.blob().len() > MAX_THREADED_PART_BYTES {
        return Err("threaded-comments part exceeds the configured resource bound".into());
    }
    let bytes = crate::common::mce::process_part(comments_part)?;
    let xml = std::str::from_utf8(bytes.as_ref())?;
    let comments = parse_threaded_comments(xml)?;
    if comments.comments.len() > MAX_THREADED_COMMENTS {
        return Err("threaded-comments part contains too many comments".into());
    }
    let mentions = comments.comments.iter().map(|comment| comment.mentions.len()).sum::<usize>();
    if mentions > MAX_THREADED_MENTIONS {
        return Err("threaded-comments part contains too many mentions".into());
    }
    if comments.comments.iter().any(|comment| {
        comment.text.as_deref().is_some_and(|text| text.encode_utf16().count() > MAX_THREADED_TEXT_UTF16)
    }) {
        return Err("threaded-comment text exceeds the configured resource bound".into());
    }
    Ok(Some(comments))
}

fn related_part_uri(
    relationships: &Relationships,
    relationship_type: &str,
    description: &str,
) -> SheetResult<Option<PackURI>> {
    let mut matching = relationships
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type);
    let Some(relationship) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() {
        return Err(format!("part has multiple {description} relationships").into());
    }
    if relationship.is_external() {
        return Err(format!("{description} relationship cannot be external").into());
    }
    Ok(Some(relationship.target_partname()?))
}

fn require_content_type(uri: &PackURI, actual: &str, expected: &str) -> SheetResult<()> {
    if actual != expected {
        return Err(
            format!("part '{uri}' has content type '{actual}', expected '{expected}'").into(),
        );
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PersonContext {
    PersonList,
    Person,
    Other,
}

fn parse_person_list(xml: &str) -> SheetResult<PersonList> {
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

    Ok(PersonList { persons })
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
    comment: ThreadedComment,
    saw_text: bool,
    saw_mentions: bool,
}

fn parse_threaded_comments(xml: &str) -> SheetResult<ThreadedComments> {
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
    Ok(ThreadedComments { comments })
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
    comments: &mut Vec<ThreadedComment>,
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

fn parse_threaded_comment(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> SheetResult<ThreadedComment> {
    let cell_ref = unqualified_attribute_value(element, b"ref", decoder)?;
    if let Some(cell_ref) = cell_ref.as_deref() {
        Cell::reference_to_coords(cell_ref)?;
    }
    let date_time = unqualified_attribute_value(element, b"dT", decoder)?;
    validate_threaded_timestamp(date_time.as_deref())?;
    Ok(ThreadedComment {
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

pub(super) fn validate_guid(value: &str, description: &str) -> SheetResult<()> {
    let bytes = value.as_bytes();
    let valid = bytes.len() == 38
        && bytes[0] == b'{'
        && bytes[37] == b'}'
        && [9, 14, 19, 24].iter().all(|&index| bytes[index] == b'-')
        && bytes[1..37]
            .iter()
            .enumerate()
            .all(|(index, byte)| matches!(index + 1, 9 | 14 | 19 | 24) || byte.is_ascii_hexdigit());
    if !valid {
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

#[cfg(test)]
mod tests {
    use litchi_opc::Part;
    use litchi_opc::part::BlobPart;

    use super::*;

    const NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
    const PERSON_ID: &str = "{11111111-1111-1111-1111-111111111111}";
    const COMMENT_ID: &str = "{22222222-2222-2222-2222-222222222222}";
    const BOB_ID: &str = "{33333333-3333-3333-3333-333333333333}";
    const MENTION_ID: &str = "{44444444-4444-4444-4444-444444444444}";

    #[test]
    fn parses_prefixed_people_and_threaded_comments() {
        let people = parse_person_list(&format!(
            r#"<tc:personList xmlns:tc="{NS}" xmlns:f="urn:foreign">
                <f:person displayName="Ignored" id="ignored"/>
                <tc:person displayName="Alice &amp; Bob" id="{PERSON_ID}" userId="alice" providerId="aad"/>
            </tc:personList>"#
        ))
        .unwrap();
        assert_eq!(people.persons.len(), 1);
        assert_eq!(people.persons[0].display_name, "Alice & Bob");

        let comments = parse_threaded_comments(&format!(
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
        let people = parse_person_list(&format!(r#"<personList xmlns="{NS}"/>"#)).unwrap();
        let comments =
            parse_threaded_comments(&format!(r#"<ThreadedComments xmlns="{NS}"/>"#)).unwrap();

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
            assert!(parse_person_list(&xml).is_err(), "accepted {xml}");
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
            assert!(parse_threaded_comments(&xml).is_err(), "accepted {xml}");
        }
    }

    fn package_with_threaded_parts() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
        let worksheet_uri = PackURI::new("/custom/sheets/sheet.xml").unwrap();
        let mut workbook_part =
            BlobPart::new(workbook_uri, ct::SML_SHEET_MAIN.to_string(), Vec::new());
        workbook_part.relate_to("people.xml", rt::PERSONS);
        let mut worksheet_part = BlobPart::new(
            worksheet_uri.clone(),
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        );
        worksheet_part.relate_to("../threads.xml", rt::THREADED_COMMENTS);
        package.relate_to("custom/book.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(worksheet_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/people.xml").unwrap(),
            ct::SML_PERSONS.to_string(),
            format!(r#"<personList xmlns="{NS}"/>"#).into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/threads.xml").unwrap(),
            ct::SML_THREADED_COMMENTS.to_string(),
            format!(r#"<ThreadedComments xmlns="{NS}"/>"#).into_bytes(),
        )));
        (package, worksheet_uri)
    }

    #[test]
    fn resolves_custom_part_locations() {
        let (package, worksheet_uri) = package_with_threaded_parts();

        assert!(read_persons(&package).unwrap().is_some());
        assert!(
            read_threaded_comments(&package, &worksheet_uri)
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn rejects_external_duplicate_and_wrong_content_type_relationships() {
        let (mut package, worksheet_uri) = package_with_threaded_parts();
        let relationships = package.get_part_mut(&worksheet_uri).unwrap().rels_mut();
        relationships.remove("rId1").unwrap();
        relationships.add_relationship(
            rt::THREADED_COMMENTS.to_string(),
            "https://example.com/thread.xml".to_string(),
            "rId1".to_string(),
            true,
        );
        assert!(read_threaded_comments(&package, &worksheet_uri).is_err());

        let (mut package, worksheet_uri) = package_with_threaded_parts();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::THREADED_COMMENTS.to_string(),
                "https://example.com/thread.xml".to_string(),
                "rId2".to_string(),
                true,
            );
        assert!(read_threaded_comments(&package, &worksheet_uri).is_err());

        let (mut package, worksheet_uri) = package_with_threaded_parts();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/threads.xml").unwrap(),
            ct::SML_PERSONS.to_string(),
            Vec::new(),
        )));
        assert!(read_threaded_comments(&package, &worksheet_uri).is_err());
    }
}
