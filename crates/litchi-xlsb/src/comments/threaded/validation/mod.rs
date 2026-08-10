#![allow(
    clippy::expect_used,
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error, an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Semantic and graph validation for XLSB threaded comments.

use std::collections::TryReserveError;
use std::collections::{HashMap, HashSet};

use litchi_ooxml_common::custom_xml::valid_guid;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use super::semantic::{Comment, Comments, Graph, People, RawAttribute, RawXml, Thread};
use super::{
    MAX_ATTRIBUTES, MAX_COMMENTS, MAX_EXTENSION_BYTES, MAX_EXTENSIONS, MAX_IDENTITY_BYTES,
    MAX_MENTIONS, MAX_PERSONS, MAX_TEXT_UTF16,
};

/// Result type for the threaded-comment owner.
pub type Result<T> = std::result::Result<T, Error>;

/// A malformed or unsafe threaded-comment semantic graph.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// A semantic or relationship invariant was violated.
    #[error("invalid threaded-comment graph: {0}")]
    Invalid(String),
    /// A timestamp, GUID, reference, or XML value could not be encoded.
    #[error("threaded-comment encoding error: {0}")]
    Encoding(String),
    /// A bounded collection could not reserve its requested memory.
    #[error("allocation failed for {resource}: {source}")]
    Allocation {
        /// Collection being grown.
        resource: &'static str,
        /// Allocator failure.
        source: TryReserveError,
    },
}

/// Validate one persons part.
pub fn validate_people(people: &People) -> Result<()> {
    if people.persons.len() > MAX_PERSONS {
        return Err(invalid("persons list contains too many people"));
    }
    validate_attributes(&people.attributes)?;
    validate_extensions(&people.extensions)?;

    let mut ids = HashSet::new();
    ids.try_reserve(people.persons.len())
        .map_err(|source| allocation("person IDs", source))?;
    for person in &people.persons {
        validate_guid_value(&person.id, "person ID")?;
        validate_identity(&person.display_name, "person display name")?;
        if let Some(value) = person.user_id.as_deref() {
            validate_identity(value, "person user ID")?;
        }
        if let Some(value) = person.provider_id.as_deref() {
            validate_identity(value, "person provider ID")?;
        }
        validate_attributes(&person.attributes)?;
        validate_extensions(&person.extensions)?;
        if !ids.insert(person.id.as_str()) {
            return Err(invalid(format!("duplicate person ID '{}'", person.id)));
        }
    }
    Ok(())
}

/// Validate one worksheet threaded-comments part without resolving people.
pub fn validate_comments(comments: &Comments) -> Result<()> {
    if comments.comments.len() > MAX_COMMENTS {
        return Err(invalid("threaded-comments list contains too many comments"));
    }
    validate_attributes(&comments.attributes)?;
    validate_extensions(&comments.extensions)?;

    let mention_count = comments
        .comments
        .iter()
        .try_fold(0usize, |count, comment| {
            count
                .checked_add(comment.mentions.len())
                .ok_or_else(|| invalid("mention count overflows"))
        })?;
    if mention_count > MAX_MENTIONS {
        return Err(invalid("threaded-comments list contains too many mentions"));
    }

    let mut comment_ids = HashSet::new();
    comment_ids
        .try_reserve(comments.comments.len())
        .map_err(|source| allocation("threaded-comment IDs", source))?;
    let mut mention_ids = HashSet::new();
    mention_ids
        .try_reserve(mention_count)
        .map_err(|source| allocation("mention IDs", source))?;

    for comment in &comments.comments {
        validate_comment(comment, &mut comment_ids, &mut mention_ids)?;
    }

    for comment in &comments.comments {
        let Some(parent_id) = comment.parent_id.as_deref() else {
            continue;
        };
        if parent_id == comment.id {
            return Err(invalid(format!(
                "threaded comment '{}' cannot parent itself",
                comment.id
            )));
        }
        if !comment_ids.contains(parent_id) {
            return Err(invalid(format!(
                "threaded comment '{}' references missing parent '{}'",
                comment.id, parent_id
            )));
        }
    }

    let mut roots = HashSet::new();
    for comment in &comments.comments {
        if comment.is_root() {
            let cell = comment.cell_ref.as_deref().ok_or_else(|| {
                invalid(format!(
                    "threaded-comment root '{}' is missing its cell reference",
                    comment.id
                ))
            })?;
            if !roots.insert(cell) {
                return Err(invalid(format!(
                    "worksheet has multiple threaded-comment roots at {cell}"
                )));
            }
        } else if comment.cell_ref.is_some() {
            return Err(invalid(format!(
                "threaded-comment reply '{}' must not carry a cell reference",
                comment.id
            )));
        }
    }

    for reply in comments
        .comments
        .iter()
        .filter(|comment| !comment.is_root())
    {
        let parent = reply.parent_id.as_deref().expect("filtered reply");
        if !comments
            .comments
            .iter()
            .any(|comment| comment.is_root() && comment.id == parent)
        {
            return Err(invalid(format!(
                "threaded-comment reply '{}' must reference a root in the same worksheet",
                reply.id
            )));
        }
    }
    Ok(())
}

/// Validate all cross-part person, comment, mention, and thread references.
pub fn validate_graph(graph: &Graph) -> Result<()> {
    if let Some(persons) = graph.persons.as_ref() {
        validate_people(&persons.persons)?;
    }
    let empty = People::default();
    let people = graph
        .persons
        .as_ref()
        .map(|part| &part.persons)
        .unwrap_or(&empty);
    let mut person_ids = HashSet::new();
    person_ids
        .try_reserve(people.persons.len())
        .map_err(|source| allocation("workbook person IDs", source))?;
    for person in &people.persons {
        person_ids.insert(person.id.as_str());
    }

    let mut comment_ids = HashSet::new();
    let mut mention_ids = HashSet::new();
    for worksheet in &graph.worksheets {
        validate_comments(&worksheet.comments)?;
        for comment in &worksheet.comments.comments {
            if !comment_ids.insert(comment.id.as_str()) {
                return Err(invalid(format!(
                    "duplicate workbook threaded-comment ID '{}'",
                    comment.id
                )));
            }
            if !person_ids.contains(comment.person_id.as_str()) {
                return Err(invalid(format!(
                    "threaded comment '{}' references missing person '{}'",
                    comment.id, comment.person_id
                )));
            }
            for mention in &comment.mentions {
                if !person_ids.contains(mention.mention_person_id.as_str()) {
                    return Err(invalid(format!(
                        "mention '{}' references missing person '{}'",
                        mention.mention_id, mention.mention_person_id
                    )));
                }
                if !mention_ids.insert(mention.mention_id.as_str()) {
                    return Err(invalid(format!(
                        "duplicate workbook mention ID '{}'",
                        mention.mention_id
                    )));
                }
            }
        }
    }
    Ok(())
}

/// Group one validated worksheet collection into root-and-reply threads.
pub fn group_threads(comments: &Comments) -> Result<Vec<Thread>> {
    validate_comments(comments)?;
    let mut by_root: HashMap<&str, usize> = HashMap::new();
    by_root
        .try_reserve(comments.comments.len())
        .map_err(|source| allocation("thread roots", source))?;
    let mut threads = Vec::new();
    for comment in &comments.comments {
        if comment.is_root() {
            by_root.insert(comment.id.as_str(), threads.len());
            threads.push(Thread {
                root: comment.clone(),
                replies: Vec::new(),
            });
        }
    }
    for comment in &comments.comments {
        let Some(parent) = comment.parent_id.as_deref() else {
            continue;
        };
        let index = by_root.get(parent).copied().ok_or_else(|| {
            invalid(format!(
                "threaded comment '{}' does not reference a root",
                comment.id
            ))
        })?;
        threads[index].replies.push(comment.clone());
    }
    Ok(threads)
}

fn validate_comment<'a>(
    comment: &'a Comment,
    comment_ids: &mut HashSet<&'a str>,
    mention_ids: &mut HashSet<&'a str>,
) -> Result<()> {
    validate_guid_value(&comment.id, "threaded-comment ID")?;
    validate_guid_value(&comment.person_id, "threaded-comment person ID")?;
    if let Some(value) = comment.parent_id.as_deref() {
        validate_guid_value(value, "threaded-comment parent ID")?;
    }
    if let Some(value) = comment.cell_ref.as_deref() {
        validate_cell_ref(value)?;
    }
    validate_timestamp(comment.date_time.as_deref())?;
    if let Some(text) = comment.text.as_deref() {
        if text.encode_utf16().count() > MAX_TEXT_UTF16 {
            return Err(invalid(format!(
                "threaded-comment '{}' text is too long",
                comment.id
            )));
        }
        validate_xml_chars(text, "threaded-comment text")?;
    } else if !comment.mentions.is_empty() {
        return Err(invalid(format!(
            "threaded-comment '{}' has mentions but no text",
            comment.id
        )));
    }
    validate_attributes(&comment.attributes)?;
    validate_extensions(&comment.extensions)?;
    if !comment_ids.insert(comment.id.as_str()) {
        return Err(invalid(format!(
            "duplicate threaded-comment ID '{}'",
            comment.id
        )));
    }
    for mention in &comment.mentions {
        validate_guid_value(&mention.mention_person_id, "mention person ID")?;
        validate_guid_value(&mention.mention_id, "mention ID")?;
        let end = mention
            .start_index
            .checked_add(mention.length)
            .ok_or_else(|| invalid("mention range overflows"))?;
        if let Some(text) = comment.text.as_deref() {
            let text_len = u32::try_from(text.encode_utf16().count())
                .map_err(|_| invalid("threaded-comment text length overflows"))?;
            if end > text_len {
                return Err(invalid(format!(
                    "mention '{}' range exceeds threaded-comment text",
                    mention.mention_id
                )));
            }
        }
        validate_attributes(&mention.attributes)?;
        validate_extensions(&mention.extensions)?;
        if !mention_ids.insert(mention.mention_id.as_str()) {
            return Err(invalid(format!(
                "duplicate mention ID '{}'",
                mention.mention_id
            )));
        }
    }
    Ok(())
}

fn validate_guid_value(value: &str, name: &str) -> Result<()> {
    if valid_guid(value) {
        Ok(())
    } else {
        Err(invalid(format!("invalid {name} GUID '{value}'")))
    }
}

fn validate_identity(value: &str, name: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_IDENTITY_BYTES {
        return Err(invalid(format!(
            "{name} must contain between one and {MAX_IDENTITY_BYTES} bytes"
        )));
    }
    validate_xml_chars(value, name)?;
    Ok(())
}

fn validate_timestamp(value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    validate_xml_chars(value, "threaded-comment timestamp")?;
    if chrono::DateTime::parse_from_rfc3339(value).is_ok()
        || chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "invalid threaded-comment timestamp '{value}'"
        )))
    }
}

fn validate_cell_ref(reference: &str) -> Result<()> {
    let bytes = reference.as_bytes();
    let Some(column_end) = bytes.iter().position(u8::is_ascii_digit) else {
        return Err(invalid(format!("invalid cell reference '{reference}'")));
    };
    if column_end == 0 || column_end == bytes.len() {
        return Err(invalid(format!("invalid cell reference '{reference}'")));
    }
    let mut column = 0u32;
    for byte in &bytes[..column_end] {
        if !byte.is_ascii_alphabetic() {
            return Err(invalid(format!("invalid cell reference '{reference}'")));
        }
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid(format!("cell reference '{reference}' overflows")))?;
    }
    if column == 0 || column > 16_384 {
        return Err(invalid(format!(
            "cell reference '{reference}' exceeds XLSB columns"
        )));
    }
    let row = std::str::from_utf8(&bytes[column_end..])
        .map_err(|_| invalid(format!("invalid cell reference '{reference}'")))?
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid cell reference '{reference}'")))?;
    if row == 0 || row > 1_048_576 {
        return Err(invalid(format!(
            "cell reference '{reference}' exceeds XLSB rows"
        )));
    }
    Ok(())
}

fn validate_attributes(attributes: &[RawAttribute]) -> Result<()> {
    if attributes.len() > MAX_ATTRIBUTES {
        return Err(invalid("too many preserved XML attributes"));
    }
    let mut names = HashSet::new();
    for attribute in attributes {
        if attribute.name.len() > MAX_IDENTITY_BYTES
            || attribute.value.len() > MAX_IDENTITY_BYTES
            || !valid_attribute_name(&attribute.name)
        {
            return Err(invalid("invalid preserved XML attribute"));
        }
        validate_xml_chars(&attribute.value, "preserved XML attribute value")?;
        if !names.insert(attribute.name.as_str()) {
            return Err(invalid(format!(
                "duplicate preserved XML attribute '{}'",
                attribute.name
            )));
        }
    }
    Ok(())
}

fn validate_extensions(extensions: &[RawXml]) -> Result<()> {
    if extensions.len() > MAX_EXTENSIONS {
        return Err(invalid("too many preserved XML extensions"));
    }
    let total = extensions.iter().try_fold(0usize, |total, extension| {
        if extension.bytes.is_empty() {
            return Err(invalid("preserved XML extension is empty"));
        }
        validate_fragment(&extension.bytes)?;
        total
            .checked_add(extension.bytes.len())
            .ok_or_else(|| invalid("preserved XML extension size overflows"))
    })?;
    if total > MAX_EXTENSION_BYTES {
        return Err(invalid(
            "preserved XML extensions exceed the resource bound",
        ));
    }
    Ok(())
}

fn valid_attribute_name(name: &str) -> bool {
    if name == "xmlns" {
        return false;
    }
    let mut parts = name.split(':');
    let Some(prefix) = parts.next() else {
        return false;
    };
    let Some(local) = parts.next() else {
        return valid_name_part(prefix);
    };
    parts.next().is_none() && valid_name_part(prefix) && valid_name_part(local)
}

fn valid_name_part(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
}

fn validate_xml_chars(value: &str, description: &str) -> Result<()> {
    if value
        .chars()
        .any(|character| !matches!(character, '\t' | '\n' | '\r') && character < '\u{20}')
    {
        return Err(invalid(format!(
            "{description} contains an XML control character"
        )));
    }
    Ok(())
}

fn validate_fragment(bytes: &[u8]) -> Result<()> {
    std::str::from_utf8(bytes).map_err(|error| Error::Encoding(error.to_string()))?;
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Encoding(error.to_string()))?;
        match event {
            Event::Start(_) => {
                if root_closed {
                    return Err(invalid("preserved XML has multiple root elements"));
                }
                root_seen = true;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("preserved XML nesting overflows"))?;
                if depth > super::MAX_XML_DEPTH {
                    return Err(invalid("preserved XML nesting exceeds the resource bound"));
                }
            },
            Event::Empty(_) => {
                if root_closed {
                    return Err(invalid("preserved XML has multiple root elements"));
                }
                if !root_seen {
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("preserved XML has an unexpected closing element"));
                }
                depth -= 1;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                let value = text
                    .decode()
                    .map_err(|error| Error::Encoding(error.to_string()))?;
                validate_xml_chars(&value, "preserved XML text")?;
                if depth == 0 && !value.trim().is_empty() {
                    return Err(invalid("preserved XML has character data outside its root"));
                }
            },
            Event::CData(text) => {
                let value = text
                    .decode()
                    .map_err(|error| Error::Encoding(error.to_string()))?;
                validate_xml_chars(&value, "preserved XML CDATA")?;
                if depth == 0 && !value.trim().is_empty() {
                    return Err(invalid("preserved XML has character data outside its root"));
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(invalid("preserved XML has a non-element outside its root"));
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "preserved XML fragments cannot contain declarations",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 {
                    return Err(invalid("preserved XML fragment is incomplete"));
                }
                return Ok(());
            },
            _ => {},
        }
    }
}

fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}
