//! Namespace-aware parser for classic SpreadsheetML comments parts.

use std::collections::HashMap;

use litchi_core::sheet::Result as SheetResult;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::cell::Cell;
use super::namespace::is_spreadsheetml_name;
use super::worksheet::Comment;
use crate::common::xml::{decode_xml_reference, unqualified_attribute_value};

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

struct CommentParser {
    authors: Vec<String>,
    comments: HashMap<String, Comment>,
    author: Option<String>,
    comment: Option<PendingComment>,
    seen_authors: bool,
    seen_comment_list: bool,
    comment_list_start: usize,
    run_saw_text: bool,
}

impl CommentParser {
    fn new() -> Self {
        Self {
            authors: Vec::new(),
            comments: HashMap::new(),
            author: None,
            comment: None,
            seen_authors: false,
            seen_comment_list: false,
            comment_list_start: 0,
            run_saw_text: false,
        }
    }

    fn parse(content: &str) -> SheetResult<HashMap<String, Comment>> {
        let mut reader = NsReader::from_reader(content.as_bytes());
        let mut parser = Self::new();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader.read_event()?.into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"comments")
                    {
                        return Err(
                            "comments part must have one SpreadsheetML comments root".into()
                        );
                    }
                    stack.push(Context::Comments);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"comments")
                    {
                        return Err(
                            "comments part must have one SpreadsheetML comments root".into()
                        );
                    }
                    return Err("comments part is missing authors and commentList".into());
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or("comments parser is missing its root context")?;
                    stack.push(parser.start(parent, &namespace, &element, decoder)?);
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or("comments parser is missing its root context")?;
                    let context = parser.start(parent, &namespace, &element, decoder)?;
                    parser.finish(context)?;
                },
                Event::Text(text) => {
                    if let Some(target) = text_target(&stack) {
                        parser.push_text(target, &text.decode()?)?;
                    }
                },
                Event::CData(text) => {
                    if let Some(target) = text_target(&stack) {
                        parser.push_text(target, &text.decode()?)?;
                    }
                },
                Event::GeneralRef(reference) => {
                    if let Some(target) = text_target(&stack) {
                        parser.push_text(target, &decode_xml_reference(&reference)?)?;
                    }
                },
                Event::End(element) => {
                    let context = stack
                        .pop()
                        .ok_or("comments part has a closing element outside its root")?;
                    parser.finish(context)?;
                    if context == Context::Comments {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"comments") {
                            return Err("comments part has an invalid root closing element".into());
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err("comments part has a missing or unterminated root".into());
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !parser.seen_authors || parser.authors.is_empty() {
            return Err("comments part contains no authors".into());
        }
        if !parser.seen_comment_list || parser.comments.is_empty() {
            return Err("comments part contains no comments".into());
        }
        Ok(parser.comments)
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> SheetResult<Context> {
        if parent == Context::Comments
            && is_spreadsheetml_name(namespace, element.name(), b"authors")
        {
            if self.seen_authors {
                return Err("duplicate comments authors element".into());
            }
            self.seen_authors = true;
            return Ok(Context::Authors);
        }
        if parent == Context::Authors && is_spreadsheetml_name(namespace, element.name(), b"author")
        {
            if self.author.is_some() {
                return Err("nested comment author".into());
            }
            self.author = Some(String::new());
            return Ok(Context::Author);
        }
        if parent == Context::Comments
            && is_spreadsheetml_name(namespace, element.name(), b"commentList")
        {
            if self.seen_comment_list {
                return Err("duplicate comments commentList element".into());
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
                .ok_or("comment text outside a comment")?;
            if comment.saw_text {
                return Err("duplicate comment text element".into());
            }
            comment.saw_text = true;
            return Ok(Context::CommentText);
        }
        if parent == Context::CommentText && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            let comment = self
                .comment
                .as_mut()
                .ok_or("simple comment text outside a comment")?;
            if comment.text_mode.is_some() {
                return Err("comment text has duplicate or mixed text content".into());
            }
            comment.text_mode = Some(CommentTextMode::Simple);
            return Ok(Context::Text(TextTarget::Comment));
        }
        if parent == Context::CommentText && is_spreadsheetml_name(namespace, element.name(), b"r")
        {
            let comment = self
                .comment
                .as_mut()
                .ok_or("rich comment run outside a comment")?;
            if matches!(comment.text_mode, Some(CommentTextMode::Simple)) {
                return Err("comment text mixes simple text and rich runs".into());
            }
            comment.text_mode = Some(CommentTextMode::Rich);
            self.run_saw_text = false;
            return Ok(Context::RichRun);
        }
        if parent == Context::RichRun && is_spreadsheetml_name(namespace, element.name(), b"t") {
            if self.run_saw_text {
                return Err("comment rich-text run has duplicate text".into());
            }
            self.run_saw_text = true;
            return Ok(Context::Text(TextTarget::Comment));
        }
        Ok(Context::Other)
    }

    fn start_comment(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> SheetResult<()> {
        if self.comment.is_some() {
            return Err("nested worksheet comment".into());
        }
        let cell_ref = required_string(element, b"ref", decoder, "comment cell reference")?;
        Cell::reference_to_coords(&cell_ref)?;
        if self.comments.contains_key(&cell_ref) {
            return Err(format!("duplicate comment reference '{cell_ref}'").into());
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

    fn push_text(&mut self, target: TextTarget, value: &str) -> SheetResult<()> {
        match target {
            TextTarget::Author => self
                .author
                .as_mut()
                .ok_or("author text outside an author")?
                .push_str(value),
            TextTarget::Comment => self
                .comment
                .as_mut()
                .ok_or("comment text outside a comment")?
                .text
                .push_str(value),
        }
        Ok(())
    }

    fn finish(&mut self, context: Context) -> SheetResult<()> {
        match context {
            Context::Author => {
                self.authors
                    .push(self.author.take().ok_or("missing pending comment author")?);
            },
            Context::RichRun if !self.run_saw_text => {
                return Err("comment rich-text run is missing its text".into());
            },
            Context::Comment => {
                let pending = self.comment.take().ok_or("missing pending comment")?;
                if !pending.saw_text {
                    return Err(format!(
                        "comment '{}' is missing its text element",
                        pending.cell_ref
                    )
                    .into());
                }
                let author_index = usize::try_from(pending.author_id).unwrap_or(usize::MAX);
                let author = self.authors.get(author_index).cloned().ok_or_else(|| {
                    format!(
                        "comment '{}' references missing author {}",
                        pending.cell_ref, pending.author_id
                    )
                })?;
                self.comments.insert(
                    pending.cell_ref.clone(),
                    Comment {
                        cell_ref: pending.cell_ref,
                        author: Some(author),
                        author_id: pending.author_id,
                        text: pending.text,
                        guid: pending.guid,
                        shape_id: pending.shape_id,
                    },
                );
            },
            Context::Authors if self.authors.is_empty() => {
                return Err("comments authors element contains no author".into());
            },
            Context::CommentList if self.comments.len() == self.comment_list_start => {
                return Err("comments commentList contains no comment".into());
            },
            _ => {},
        }
        Ok(())
    }
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
) -> SheetResult<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| format!("invalid {description} '{value}'").into())
        })
        .transpose()
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

pub(crate) fn parse_comments_xml(content: &str) -> SheetResult<HashMap<String, Comment>> {
    CommentParser::parse(content)
}

#[cfg(test)]
mod tests {
    use super::*;

    const STRICT_S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

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
        let comments = parse_comments_xml(&xml).unwrap();
        let comment = &comments["B2"];

        assert_eq!(comment.author.as_deref(), Some("陈"));
        assert_eq!(comment.author_id, 1);
        assert_eq!(comment.text, "Hello 世界");
        assert_eq!(comment.guid.as_deref(), Some("{guid}"));
        assert_eq!(comment.shape_id, Some(7));
    }

    #[test]
    fn rejects_malformed_comments() {
        let invalid = [
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

        for xml in invalid {
            assert!(parse_comments_xml(&xml).is_err(), "accepted {xml}");
        }
    }
}
