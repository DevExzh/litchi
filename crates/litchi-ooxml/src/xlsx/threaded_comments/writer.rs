//! Writer module for threaded comments XML generation.

use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

use litchi_core::sheet::Result as SheetResult;
use litchi_core::xml::escape_xml;

use super::person::{Mention, Person, PersonList};
use super::reader::validate_guid;
use super::{
    MAX_THREADED_COMMENTS, MAX_THREADED_IDENTITY_BYTES, MAX_THREADED_MENTIONS,
    MAX_THREADED_PERSONS, MAX_THREADED_TEXT_UTF16, ThreadedComment, ThreadedComments,
    validate_threaded_timestamp,
};
use crate::xlsx::Cell;

const XML_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const THREADED_COMMENTS_NS: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// Write person list to XML.
///
/// Generates the `xl/persons/person.xml` part containing all persons
/// who can author threaded comments in the workbook.
pub fn write_persons(person_list: &PersonList) -> SheetResult<String> {
    validate_person_list(person_list)?;
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
pub fn write_threaded_comments(comments: &ThreadedComments) -> SheetResult<String> {
    validate_threaded_comments(comments)?;
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
fn write_threaded_comment(xml: &mut String, comment: &ThreadedComment) -> SheetResult<()> {
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

fn validate_person_list(person_list: &PersonList) -> SheetResult<()> {
    if person_list.persons.len() > MAX_THREADED_PERSONS {
        return Err("persons list contains too many people".into());
    }
    let mut ids = HashSet::with_capacity(person_list.persons.len());
    for person in &person_list.persons {
        validate_guid(&person.id, "person ID")?;
        if person.display_name.len() > MAX_THREADED_IDENTITY_BYTES
            || person.user_id.as_ref().is_some_and(|value| value.len() > MAX_THREADED_IDENTITY_BYTES)
            || person.provider_id.as_ref().is_some_and(|value| value.len() > MAX_THREADED_IDENTITY_BYTES)
        {
            return Err(format!("person '{}' has oversized identity metadata", person.id).into());
        }
        if !ids.insert(person.id.as_str()) {
            return Err(format!("duplicate person ID '{}'", person.id).into());
        }
    }
    Ok(())
}

fn validate_threaded_comments(comments: &ThreadedComments) -> SheetResult<()> {
    if comments.comments.len() > MAX_THREADED_COMMENTS {
        return Err("threaded-comments list contains too many comments".into());
    }
    let mut comment_ids = HashSet::with_capacity(comments.comments.len());
    let mention_count = comments
        .comments
        .iter()
        .map(|comment| comment.mentions.len())
        .sum();
    if mention_count > MAX_THREADED_MENTIONS {
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
            Cell::reference_to_coords(cell_ref)?;
        }
        validate_threaded_timestamp(comment.date_time.as_deref())?;
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
        if text_len.is_some_and(|length| length as usize > MAX_THREADED_TEXT_UTF16) {
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

#[cfg(test)]
mod tests {
    use super::*;

    const PERSON_ID: &str = "{11111111-1111-1111-1111-111111111111}";
    const COMMENT_ID: &str = "{22222222-2222-2222-2222-222222222222}";
    const REPLY_ID: &str = "{33333333-3333-3333-3333-333333333333}";
    const MENTION_ID: &str = "{44444444-4444-4444-4444-444444444444}";

    #[test]
    fn writes_schema_valid_people_and_comments() {
        let people = PersonList {
            persons: vec![Person {
                display_name: "Alice & Bob".into(),
                id: PERSON_ID.into(),
                user_id: Some("alice@example.com".into()),
                provider_id: None,
            }],
        };
        let people_xml = write_persons(&people).unwrap();
        assert!(people_xml.contains("Alice &amp; Bob"));

        let comments = ThreadedComments {
            comments: vec![
                ThreadedComment {
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
                ThreadedComment {
                    id: REPLY_ID.into(),
                    person_id: PERSON_ID.into(),
                    parent_id: Some(COMMENT_ID.into()),
                    ..Default::default()
                },
            ],
        };
        let comments_xml = write_threaded_comments(&comments).unwrap();
        assert!(comments_xml.contains("<text>Hi @Bob</text>"));
        assert!(comments_xml.contains(&format!(r#" parentId="{COMMENT_ID}""#)));
    }

    #[test]
    fn rejects_invalid_people_and_comments() {
        let duplicate_people = PersonList {
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
            ThreadedComment {
                id: "not-a-guid".into(),
                person_id: PERSON_ID.into(),
                ..Default::default()
            },
            ThreadedComment {
                cell_ref: Some("A0".into()),
                id: COMMENT_ID.into(),
                person_id: PERSON_ID.into(),
                ..Default::default()
            },
            ThreadedComment {
                id: COMMENT_ID.into(),
                person_id: PERSON_ID.into(),
                parent_id: Some(REPLY_ID.into()),
                ..Default::default()
            },
            ThreadedComment {
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
                write_threaded_comments(&ThreadedComments {
                    comments: vec![comment]
                })
                .is_err()
            );
        }
    }
}
