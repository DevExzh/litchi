//! Writer module for threaded comments XML generation.

use crate::common::xml::escape_xml;
use crate::sheet::Result as SheetResult;
use std::fmt::Write as FmtWrite;

use super::person::{Mention, Person, PersonList};
use super::{ThreadedComment, ThreadedComments};

const XML_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const THREADED_COMMENTS_NS: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// Write person list to XML.
///
/// Generates the `xl/persons/person.xml` part containing all persons
/// who can author threaded comments in the workbook.
pub fn write_persons(person_list: &PersonList) -> SheetResult<String> {
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
