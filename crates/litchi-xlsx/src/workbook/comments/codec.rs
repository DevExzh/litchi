//! Bounded XML codec for classic `SpreadsheetML` comments.

use std::collections::{BTreeMap, HashSet};
use std::fmt::Write as FmtWrite;

use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Result, invalid};

use super::model::{Comment, Comments};
use super::{MAX_AUTHORS, MAX_COMMENTS, MAX_PART_BYTES, MAX_TEXT_BYTES};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Comments,
    Authors,
    Author,
    CommentList,
    Comment,
    CommentText,
    RichRun,
    Text(TextTarget),
    Other,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    Author,
    Comment,
}

struct PendingComment {
    cell_ref: String,
    author_id: u32,
    text: String,
    guid: Option<String>,
    shape_id: Option<u32>,
    saw_text: bool,
    text_mode: Option<CommentTextMode>,
}

#[derive(Clone, Copy)]
enum CommentTextMode {
    Simple,
    Rich,
}

struct Parser {
    authors: Vec<String>,
    comments: BTreeMap<String, Comment>,
    author: Option<String>,
    comment: Option<PendingComment>,
    seen_authors: bool,
    seen_comment_list: bool,
    comment_list_start: usize,
    run_saw_text: bool,
}

impl Parser {
    fn new() -> Self {
        Self {
            authors: Vec::new(),
            comments: BTreeMap::new(),
            author: None,
            comment: None,
            seen_authors: false,
            seen_comment_list: false,
            comment_list_start: 0,
            run_saw_text: false,
        }
    }

    fn parse(content: &str) -> Result<Comments> {
        if content.len() > MAX_PART_BYTES {
            return Err(invalid(
                "classic comments part exceeds the configured resource bound",
            ));
        }
        let processed = litchi_ooxml_common::mce::process_str(content)?;
        let mut reader = NsReader::from_reader(processed.as_bytes());
        let mut parser = Self::new();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| invalid(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"comments")
                    {
                        return Err(invalid(
                            "comments part must have one SpreadsheetML comments root",
                        ));
                    }
                    stack.push(Context::Comments);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"comments")
                    {
                        return Err(invalid(
                            "comments part must have one SpreadsheetML comments root",
                        ));
                    }
                    return Err(invalid("comments part is missing authors and commentList"));
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("comments parser is missing its root context"))?;
                    stack.push(parser.start(parent, &namespace, &element, decoder)?);
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("comments parser is missing its root context"))?;
                    let context = parser.start(parent, &namespace, &element, decoder)?;
                    parser.finish(context)?;
                },
                Event::Text(text) => {
                    if let Some(target) = text_target(&stack) {
                        let value = text.decode().map_err(|error| invalid(error.to_string()))?;
                        parser.push_text(target, &value)?;
                    }
                },
                Event::CData(text) => {
                    if let Some(target) = text_target(&stack) {
                        let value = text.decode().map_err(|error| invalid(error.to_string()))?;
                        parser.push_text(target, &value)?;
                    }
                },
                Event::GeneralRef(reference) => {
                    if let Some(target) = text_target(&stack) {
                        let value = decode_xml_reference(&reference)?;
                        parser.push_text(target, &value)?;
                    }
                },
                Event::End(element) => {
                    let context = stack.pop().ok_or_else(|| {
                        invalid("comments part has a closing element outside its root")
                    })?;
                    parser.finish(context)?;
                    if context == Context::Comments {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"comments") {
                            return Err(invalid(
                                "comments part has an invalid root closing element",
                            ));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(invalid("comments part has a missing or unterminated root"));
                },
                Event::Eof => break,
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
            }
        }

        if !parser.seen_authors || parser.authors.is_empty() {
            return Err(invalid("comments part contains no authors"));
        }
        if !parser.seen_comment_list || parser.comments.is_empty() {
            return Err(invalid("comments part contains no comments"));
        }
        let value = Comments {
            authors: parser.authors,
            comments: parser.comments,
        };
        validate_comments(&value)?;
        Ok(value)
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Context> {
        if parent == Context::Comments
            && is_spreadsheetml_name(namespace, element.name(), b"authors")
        {
            if self.seen_authors {
                return Err(invalid("duplicate comments authors element"));
            }
            self.seen_authors = true;
            return Ok(Context::Authors);
        }
        if parent == Context::Authors && is_spreadsheetml_name(namespace, element.name(), b"author")
        {
            if self.author.is_some() {
                return Err(invalid("nested comment author"));
            }
            self.author = Some(String::new());
            return Ok(Context::Author);
        }
        if parent == Context::Comments
            && is_spreadsheetml_name(namespace, element.name(), b"commentList")
        {
            if self.seen_comment_list {
                return Err(invalid("duplicate comments commentList element"));
            }
            self.seen_comment_list = true;
            self.comment_list_start = self.comments.len();
            return Ok(Context::CommentList);
        }
        if parent == Context::CommentList
            && is_spreadsheetml_name(namespace, element.name(), b"comment")
        {
            self.start_comment(element, decoder)?;
            return Ok(Context::Comment);
        }
        if parent == Context::Comment && is_spreadsheetml_name(namespace, element.name(), b"text") {
            let comment = self
                .comment
                .as_mut()
                .ok_or_else(|| invalid("comment text outside a comment"))?;
            if comment.saw_text {
                return Err(invalid("duplicate comment text element"));
            }
            comment.saw_text = true;
            return Ok(Context::CommentText);
        }
        if parent == Context::CommentText && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            let comment = self
                .comment
                .as_mut()
                .ok_or_else(|| invalid("simple comment text outside a comment"))?;
            if comment.text_mode.is_some() {
                return Err(invalid("comment text has duplicate or mixed text content"));
            }
            comment.text_mode = Some(CommentTextMode::Simple);
            return Ok(Context::Text(TextTarget::Comment));
        }
        if parent == Context::CommentText && is_spreadsheetml_name(namespace, element.name(), b"r")
        {
            let comment = self
                .comment
                .as_mut()
                .ok_or_else(|| invalid("rich comment run outside a comment"))?;
            if matches!(comment.text_mode, Some(CommentTextMode::Simple)) {
                return Err(invalid("comment text mixes simple text and rich runs"));
            }
            comment.text_mode = Some(CommentTextMode::Rich);
            self.run_saw_text = false;
            return Ok(Context::RichRun);
        }
        if parent == Context::RichRun && is_spreadsheetml_name(namespace, element.name(), b"t") {
            if self.run_saw_text {
                return Err(invalid("comment rich-text run has duplicate text"));
            }
            self.run_saw_text = true;
            return Ok(Context::Text(TextTarget::Comment));
        }
        Ok(Context::Other)
    }

    fn start_comment(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.comment.is_some() {
            return Err(invalid("nested worksheet comment"));
        }
        let cell_ref = required_string(element, b"ref", decoder, "comment cell reference")?;
        litchi_sheet::Cell::from_a1(&cell_ref)?;
        if self.comments.contains_key(&cell_ref) {
            return Err(invalid(format!("duplicate comment reference '{cell_ref}'")));
        }
        self.comment = Some(PendingComment {
            cell_ref,
            author_id: required_u32(element, b"authorId", decoder, "comment author ID")?,
            text: String::new(),
            guid: unqualified_attribute_value(element, b"guid", decoder)?,
            shape_id: optional_u32(element, b"shapeId", decoder, "comment shape ID")?,
            saw_text: false,
            text_mode: None,
        });
        Ok(())
    }

    fn push_text(&mut self, target: TextTarget, value: &str) -> Result<()> {
        match target {
            TextTarget::Author => self
                .author
                .as_mut()
                .ok_or_else(|| invalid("author text outside an author"))?
                .push_str(value),
            TextTarget::Comment => self
                .comment
                .as_mut()
                .ok_or_else(|| invalid("comment text outside a comment"))?
                .text
                .push_str(value),
        }
        let size = match target {
            TextTarget::Author => self.author.as_ref().map_or(0, String::len),
            TextTarget::Comment => self
                .comment
                .as_ref()
                .map_or(0, |comment| comment.text.len()),
        };
        if size > MAX_TEXT_BYTES {
            return Err(invalid(
                "classic comments text exceeds the configured resource bound",
            ));
        }
        Ok(())
    }

    fn finish(&mut self, context: Context) -> Result<()> {
        match context {
            Context::Author => {
                let author = self
                    .author
                    .take()
                    .ok_or_else(|| invalid("missing pending comment author"))?;
                if author.len() > MAX_TEXT_BYTES {
                    return Err(invalid(
                        "classic comments author exceeds the configured resource bound",
                    ));
                }
                self.authors.push(author);
                if self.authors.len() > MAX_AUTHORS {
                    return Err(invalid(
                        "classic comments author count exceeds the configured resource bound",
                    ));
                }
            },
            Context::RichRun if !self.run_saw_text => {
                return Err(invalid("comment rich-text run is missing its text"));
            },
            Context::Comment => {
                let pending = self
                    .comment
                    .take()
                    .ok_or_else(|| invalid("missing pending comment"))?;
                if !pending.saw_text {
                    return Err(invalid(format!(
                        "comment '{}' is missing its text element",
                        pending.cell_ref
                    )));
                }
                let author = self
                    .authors
                    .get(usize::try_from(pending.author_id).unwrap_or(usize::MAX))
                    .cloned()
                    .ok_or_else(|| {
                        invalid(format!(
                            "comment '{}' references missing author {}",
                            pending.cell_ref, pending.author_id
                        ))
                    })?;
                let cell_ref = pending.cell_ref;
                self.comments.insert(
                    cell_ref.clone(),
                    Comment {
                        cell_ref,
                        author,
                        author_id: pending.author_id,
                        text: pending.text,
                        guid: pending.guid,
                        shape_id: pending.shape_id,
                    },
                );
                if self.comments.len() > MAX_COMMENTS {
                    return Err(invalid(
                        "classic comments count exceeds the configured resource bound",
                    ));
                }
            },
            Context::Authors if self.authors.is_empty() => {
                return Err(invalid("comments authors element contains no author"));
            },
            Context::CommentList if self.comments.len() == self.comment_list_start => {
                return Err(invalid("comments commentList contains no comment"));
            },
            Context::Comments
            | Context::Authors
            | Context::CommentList
            | Context::CommentText
            | Context::RichRun
            | Context::Text(_)
            | Context::Other => {},
        }
        Ok(())
    }
}

/// Parse one complete classic `SpreadsheetML` comments part.
pub fn parse_comments(content: &str) -> Result<Comments> {
    Parser::parse(content)
}

fn text_target(stack: &[Context]) -> Option<TextTarget> {
    match stack.last() {
        Some(Context::Author) => Some(TextTarget::Author),
        Some(Context::Text(target)) => Some(*target),
        _ => None,
    }
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_source| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn is_spreadsheetml_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == SML.as_bytes()
            || value.as_ref() == STRICT_SML.as_bytes())
}

/// Validate the semantic graph before serialization or package insertion.
pub fn validate_comments(value: &Comments) -> Result<()> {
    if value.authors.is_empty() {
        return Err(invalid("classic comments require at least one author"));
    }
    if value.authors.len() > MAX_AUTHORS {
        return Err(invalid(
            "classic comments author count exceeds the configured resource bound",
        ));
    }
    if value.comments.is_empty() {
        return Err(invalid("classic comments require at least one comment"));
    }
    if value.comments.len() > MAX_COMMENTS {
        return Err(invalid(
            "classic comments count exceeds the configured resource bound",
        ));
    }
    let mut cells = HashSet::with_capacity(value.comments.len());
    for (cell_ref, comment) in &value.comments {
        if cell_ref != &comment.cell_ref {
            return Err(invalid(
                "classic comment map key does not match comment cell_ref",
            ));
        }
        let cell = litchi_sheet::Cell::from_a1(cell_ref)?;
        if !cells.insert(cell) {
            return Err(invalid(format!(
                "classic comments contain duplicate semantic cell '{cell_ref}'"
            )));
        }
        if comment.author_id as usize >= value.authors.len() {
            return Err(invalid(format!(
                "comment '{cell_ref}' references missing author {}",
                comment.author_id
            )));
        }
        if comment.author != value.authors[comment.author_id as usize] {
            return Err(invalid(format!(
                "comment '{cell_ref}' author does not match authorId {}",
                comment.author_id
            )));
        }
        super::validation::text(&comment.text, "text")?;
        super::validation::text(&comment.author, "author")?;
        if let Some(shape_id) = comment.shape_id {
            let _ = shape_id;
        }
    }
    Ok(())
}

/// Serialize a complete classic comments part with deterministic ordering.
pub fn write_comments(value: &Comments) -> Result<Vec<u8>> {
    validate_comments(value)?;
    let mut xml = String::with_capacity(256);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>"#,
    );
    for author in &value.authors {
        write!(xml, "<author>{}</author>", escape_text(author))
            .map_err(|_source| invalid("classic comments XML formatting failed"))?;
    }
    xml.push_str("</authors><commentList>");
    for comment in value.comments.values() {
        write!(
            xml,
            "<comment ref=\"{}\" authorId=\"{}\"",
            escape_attribute(&comment.cell_ref),
            comment.author_id
        )
        .map_err(|_source| invalid("classic comments XML formatting failed"))?;
        if let Some(guid) = &comment.guid {
            write!(xml, " guid=\"{}\"", escape_attribute(guid))
                .map_err(|_source| invalid("classic comments XML formatting failed"))?;
        }
        if let Some(shape_id) = comment.shape_id {
            write!(xml, " shapeId=\"{shape_id}\"")
                .map_err(|_source| invalid("classic comments XML formatting failed"))?;
        }
        xml.push('>');
        write!(
            xml,
            "<text><t>{}</t></text></comment>",
            escape_text(&comment.text)
        )
        .map_err(|_source| invalid("classic comments XML formatting failed"))?;
    }
    xml.push_str("</commentList></comments>");
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid(
            "classic comments part exceeds the configured resource bound",
        ));
    }
    Ok(xml.into_bytes())
}

fn escape_attribute(value: &str) -> String {
    value
        .chars()
        .flat_map(|character| match character {
            '&' => "&amp;".chars().collect::<Vec<_>>(),
            '<' => "&lt;".chars().collect(),
            '"' => "&quot;".chars().collect(),
            '\t' => "&#x9;".chars().collect(),
            '\n' => "&#xA;".chars().collect(),
            '\r' => "&#xD;".chars().collect(),
            _ => vec![character],
        })
        .collect()
}

fn escape_text(value: &str) -> String {
    value
        .chars()
        .flat_map(|character| match character {
            '&' => "&amp;".chars().collect::<Vec<_>>(),
            '<' => "&lt;".chars().collect(),
            '>' => "&gt;".chars().collect(),
            _ => vec![character],
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    const STRICT_S: &str = STRICT_SML;

    #[test]
    fn parses_strict_comments_and_rich_text() {
        let xml = format!(
            r#"<s:comments xmlns:s="{STRICT_S}" xmlns:f="urn:foreign">
                <f:authors><s:author>Ignored</s:author></f:authors>
                <s:authors><s:author>Alice &amp; Bob</s:author><s:author>陈</s:author></s:authors>
                <s:commentList>
                    <s:comment ref="B2" authorId="1" guid="{{guid}}" shapeId="7">
                        <s:text><s:r><s:rPr><s:b/></s:rPr><s:t>Hello </s:t></s:r><s:r><s:t>世界</s:t></s:r></s:text>
                    </s:comment>
                </s:commentList>
            </s:comments>"#
        );
        let comments = parse_comments(&xml).unwrap();
        let comment = &comments.comments["B2"];
        assert_eq!(comment.author, "陈");
        assert_eq!(comment.author_id, 1);
        assert_eq!(comment.text, "Hello 世界");
        assert_eq!(comment.guid.as_deref(), Some("{guid}"));
        assert_eq!(comment.shape_id, Some(7));
    }

    #[test]
    fn rejects_malformed_comments() {
        let invalid_xml = [
            format!(r#"<comments xmlns="{STRICT_S}"/>"#),
            format!(r#"<comments xmlns="{STRICT_S}"><authors/><commentList/></comments>"#),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment authorId="0"><text/></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A0" authorId="0"><text/></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A1"><text/></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="1"><text/></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"/></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text/></comment><comment ref="A1" authorId="0"><text/></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><authors><author>B</author></authors><commentList><comment ref="A1" authorId="0"><text/></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t>x</t><r><t>y</t></r></text></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t>x</t><t>y</t></text></comment></commentList></comments>"#
            ),
            format!(
                r#"<comments xmlns="{STRICT_S}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><r/></text></comment></commentList></comments>"#
            ),
        ];
        for xml in invalid_xml {
            assert!(parse_comments(&xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn writer_round_trips_semantics() {
        let mut comments = Comments {
            authors: vec!["A&B".into()],
            comments: BTreeMap::new(),
        };
        comments.comments.insert(
            "A1".into(),
            Comment {
                cell_ref: "A1".into(),
                author: "A&B".into(),
                author_id: 0,
                text: "x < y".into(),
                guid: None,
                shape_id: None,
            },
        );
        let xml = write_comments(&comments).unwrap();
        assert_eq!(
            parse_comments(std::str::from_utf8(&xml).unwrap()).unwrap(),
            comments
        );
    }
}
