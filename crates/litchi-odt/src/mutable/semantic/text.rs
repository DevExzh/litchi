//! Small, lossless edits for ODF inline text controls.

use super::super::model::MutableDocument;
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 4_096;
const MAX_PARAGRAPHS: usize = 1_000_000;

impl MutableDocument {
    /// Append an ODF `text:line-break` to a paragraph selected in document
    /// order, including paragraphs nested in lists, tables, and note bodies.
    ///
    /// The operation edits the authoritative XML snapshot instead of rebuilding
    /// the structural projection, so spans, hyperlinks, fields, and producer
    /// extensions already inside the paragraph remain byte-for-byte intact.
    pub fn append_line_break(&mut self, paragraph_index: usize) -> Result<()> {
        self.edit_content_xml(|xml| append_line_break_xml(xml, paragraph_index))
    }
}

fn append_line_break_xml(xml: &str, paragraph_index: usize) -> Result<String> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("content.xml is too large for an inline text edit");
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut paragraphs = 0usize;
    let mut target_depth = None;
    let mut target_name = None::<String>;
    let mut edit = None::<(usize, usize, String)>;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid content.xml: {error}")))?;
        let text_namespace = crate::elements::xml::is_bound(&namespace, TEXT_NAMESPACE);
        drop(namespace);
        let event_end = reader.buffer_position() as usize;

        match event {
            Event::Start(ref element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| bad("content.xml nesting depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid("content.xml nesting exceeds 4096 levels");
                }
                if text_namespace && element.local_name().as_ref() == b"p" {
                    if paragraphs >= MAX_PARAGRAPHS {
                        return invalid("content.xml contains too many paragraphs");
                    }
                    if paragraphs == paragraph_index {
                        target_depth = Some(depth);
                        target_name = Some(qualified_name(element.name().as_ref())?);
                    }
                    paragraphs += 1;
                }
            },
            Event::Empty(ref element)
                if text_namespace && element.local_name().as_ref() == b"p" =>
            {
                if paragraphs >= MAX_PARAGRAPHS {
                    return invalid("content.xml contains too many paragraphs");
                }
                if paragraphs == paragraph_index {
                    let name = qualified_name(element.name().as_ref())?;
                    let raw = &xml[event_start..event_end];
                    let slash = raw
                        .rfind("/>")
                        .ok_or_else(|| bad("empty paragraph is not self-closing"))?;
                    let mut replacement = String::with_capacity(raw.len() + 64);
                    replacement.push_str(&raw[..slash]);
                    replacement.push('>');
                    replacement.push_str(&line_break_fragment(&name));
                    replacement.push_str("</");
                    replacement.push_str(&name);
                    replacement.push('>');
                    edit = Some((event_start, event_end, replacement));
                }
                paragraphs += 1;
            },
            Event::End(_) => {
                if target_depth == Some(depth) {
                    let name = target_name
                        .take()
                        .ok_or_else(|| bad("paragraph edit lost its qualified name"))?;
                    edit = Some((event_start, event_start, line_break_fragment(&name)));
                    target_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| bad("content.xml nesting depth underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are prohibited");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    let (start, end, replacement) =
        edit.ok_or_else(|| bad(format!("paragraph index {paragraph_index} does not exist")))?;
    let mut output = String::with_capacity(
        xml.len()
            .checked_sub(end - start)
            .and_then(|size| size.checked_add(replacement.len()))
            .ok_or_else(|| bad("content.xml size overflow"))?,
    );
    output.push_str(&xml[..start]);
    output.push_str(&replacement);
    output.push_str(&xml[end..]);
    Ok(output)
}

fn qualified_name(name: &[u8]) -> Result<String> {
    std::str::from_utf8(name)
        .map(str::to_string)
        .map_err(|_| bad("ODF paragraph name is not valid UTF-8"))
}

fn line_break_fragment(paragraph_name: &str) -> String {
    let prefix = paragraph_name
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:"))
        .unwrap_or_default();
    format!("<{prefix}line-break/>")
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(bad(message))
}

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
