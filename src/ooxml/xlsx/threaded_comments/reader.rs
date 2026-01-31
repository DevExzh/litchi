//! Reader module for threaded comments XML parsing.

use crate::ooxml::opc::constants::relationship_type as rt;
use crate::ooxml::opc::{OpcPackage, PackURI};
use crate::sheet::Result as SheetResult;

use super::person::{Mention, Person, PersonList};
use super::{ThreadedComment, ThreadedComments};

/// Read the person list from a workbook package.
///
/// Persons are stored in `xl/persons/person.xml` and referenced by their
/// unique IDs in threaded comments.
pub fn read_persons(package: &OpcPackage) -> SheetResult<Option<PersonList>> {
    let workbook_uri = PackURI::new("/xl/workbook.xml")?;
    let workbook_part = package.get_part(&workbook_uri)?;
    let workbook_rels = workbook_part.rels();

    for rel in workbook_rels.iter() {
        if rel.reltype() != rt::PERSONS {
            continue;
        }

        let persons_uri = rel.target_partname()?;
        let persons_part = package.get_part(&persons_uri)?;
        let xml = std::str::from_utf8(persons_part.blob())?;

        return parse_person_list(xml);
    }

    Ok(None)
}

/// Read threaded comments from a worksheet.
///
/// Threaded comments are stored in parts related to the worksheet
/// (e.g., `xl/threadedComments/threadedComment1.xml`).
pub fn read_threaded_comments(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
) -> SheetResult<Option<ThreadedComments>> {
    let worksheet_part = package.get_part(worksheet_uri)?;
    let worksheet_rels = worksheet_part.rels();

    for rel in worksheet_rels.iter() {
        if rel.reltype() != rt::THREADED_COMMENTS {
            continue;
        }

        let comments_uri = rel.target_partname()?;
        let comments_part = package.get_part(&comments_uri)?;
        let xml = std::str::from_utf8(comments_part.blob())?;

        return parse_threaded_comments(xml);
    }

    Ok(None)
}

/// Parse person list XML.
fn parse_person_list(xml: &str) -> SheetResult<Option<PersonList>> {
    let mut persons = Vec::new();

    let mut pos = 0;
    while let Some(person_start) = xml[pos..].find("<person ") {
        let abs_start = pos + person_start;
        let after_tag = &xml[abs_start..];

        let tag_end = match after_tag.find('>') {
            Some(e) => e,
            None => break,
        };

        let tag = &after_tag[..tag_end + 1];

        if let Some(person) = parse_person(tag) {
            persons.push(person);
        }

        pos = abs_start + tag_end + 1;
    }

    if persons.is_empty() {
        Ok(None)
    } else {
        Ok(Some(PersonList { persons }))
    }
}

/// Parse a single person from XML tag.
fn parse_person(tag: &str) -> Option<Person> {
    let display_name = extract_attr(tag, "displayName")?;
    let id = extract_attr(tag, "id")?;
    let user_id = extract_attr(tag, "userId");
    let provider_id = extract_attr(tag, "providerId");

    Some(Person {
        display_name,
        id,
        user_id,
        provider_id,
    })
}

/// Parse threaded comments XML.
fn parse_threaded_comments(xml: &str) -> SheetResult<Option<ThreadedComments>> {
    let mut comments = Vec::new();

    let mut pos = 0;
    while let Some(comment_start) = xml[pos..].find("<threadedComment") {
        let abs_start = pos + comment_start;
        let after_comment = &xml[abs_start..];

        let comment_end = match after_comment.find("</threadedComment>") {
            Some(e) => e + "</threadedComment>".len(),
            None => {
                // Self-closing tag
                if let Some(e) = after_comment.find("/>") {
                    e + 2
                } else {
                    break;
                }
            },
        };

        let comment_xml = &after_comment[..comment_end];
        if let Some(comment) = parse_threaded_comment(comment_xml) {
            comments.push(comment);
        }

        pos = abs_start + comment_end;
    }

    if comments.is_empty() {
        Ok(None)
    } else {
        Ok(Some(ThreadedComments { comments }))
    }
}

/// Parse a single threaded comment from XML.
fn parse_threaded_comment(xml: &str) -> Option<ThreadedComment> {
    let tag_end = xml.find('>')?;
    let tag = &xml[..tag_end + 1];

    let id = extract_attr(tag, "id")?;
    let person_id = extract_attr(tag, "personId")?;
    let cell_ref = extract_attr(tag, "ref");
    let parent_id = extract_attr(tag, "parentId");
    let date_time = extract_attr(tag, "dT");
    let done = extract_attr(tag, "done").map(|v| v == "1" || v.to_lowercase() == "true");

    let text = extract_text_element(xml);
    let mentions = parse_mentions(xml);

    Some(ThreadedComment {
        cell_ref,
        id,
        parent_id,
        person_id,
        text,
        date_time,
        done,
        mentions,
    })
}

/// Extract text from <text> element.
fn extract_text_element(xml: &str) -> Option<String> {
    let text_start = xml.find("<text>")?;
    let after_start = &xml[text_start + "<text>".len()..];
    let text_end = after_start.find("</text>")?;
    Some(after_start[..text_end].to_string())
}

/// Parse mentions from XML.
fn parse_mentions(xml: &str) -> Vec<Mention> {
    let mut mentions = Vec::new();

    let mentions_start = match xml.find("<mentions>") {
        Some(s) => s,
        None => return mentions,
    };

    let after_mentions = &xml[mentions_start..];
    let mentions_end = match after_mentions.find("</mentions>") {
        Some(e) => e,
        None => return mentions,
    };

    let mentions_section = &after_mentions[..mentions_end];

    let mut pos = 0;
    while let Some(mention_start) = mentions_section[pos..].find("<mention ") {
        let abs_start = pos + mention_start;
        let after_tag = &mentions_section[abs_start..];

        let tag_end = match after_tag.find("/>") {
            Some(e) => e + 2,
            None => break,
        };

        let tag = &after_tag[..tag_end];
        if let Some(mention) = parse_mention(tag) {
            mentions.push(mention);
        }

        pos = abs_start + tag_end;
    }

    mentions
}

/// Parse a single mention from XML tag.
fn parse_mention(tag: &str) -> Option<Mention> {
    let mention_person_id = extract_attr(tag, "mentionpersonId")?;
    let mention_id = extract_attr(tag, "mentionId")?;
    let start_index = extract_attr(tag, "startIndex")?.parse().ok()?;
    let length = extract_attr(tag, "length")?.parse().ok()?;

    Some(Mention {
        mention_person_id,
        mention_id,
        start_index,
        length,
    })
}

/// Extract an attribute value from an XML tag.
fn extract_attr(tag: &str, attr_name: &str) -> Option<String> {
    let pattern = format!("{}=\"", attr_name);
    let start = tag.find(&pattern)? + pattern.len();
    let after_start = &tag[start..];
    let end = after_start.find('"')?;
    Some(after_start[..end].to_string())
}
