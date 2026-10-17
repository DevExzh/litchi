//! Visible text of indexed nodes, read on demand from the source bytes.

use std::borrow::Cow;

use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::xml::{Ns, Prefixes, unescape};
use super::{Kind, Node};

/// The largest text one node read returns.
const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

pub(super) fn text<'a>(node: Node<'a>) -> Result<Cow<'a, str>> {
    let kind = node.kind();
    let range = node.content_range();
    let content = &node.body.content()[range];
    // Character data without markup, references or collapsible white space
    // is the text itself.
    if is_plain(content) {
        return Ok(Cow::Borrowed(content));
    }
    let mut out = String::new();
    walk(node.body.prefixes(), kind, content, &mut out)?;
    Ok(Cow::Owned(out))
}

/// The visible text of `content`, markup of a part bound by `prefixes`, read
/// as the content of a `root` node (record 0328: header and footer regions,
/// shapes and tables of contents).
pub(super) fn text_of(prefixes: &Prefixes, root: Kind, content: &str) -> Result<String> {
    if is_plain(content) {
        return Ok(content.to_owned());
    }
    let mut out = String::new();
    walk(prefixes, root, content, &mut out)?;
    Ok(out)
}

fn is_plain(content: &str) -> bool {
    let bytes = content.as_bytes();
    if bytes.first() == Some(&b' ') {
        return false;
    }
    !bytes
        .windows(2)
        .any(|pair| pair[0] == b' ' && pair[1] == b' ')
        && !bytes
            .iter()
            .any(|byte| matches!(byte, b'<' | b'&' | b'\t' | b'\n' | b'\r'))
}

/// Whether the element `ns:local` is left out of the text of a `root` node.
pub(super) fn skipped(root: Kind, ns: Ns, local: &[u8]) -> bool {
    match (ns, local) {
        (Ns::Text, b"tracked-changes" | b"ruby-text" | b"note-citation") => true,
        (Ns::Text, b"note") => !matches!(root, Kind::Footnote | Kind::Endnote),
        (Ns::Office, b"annotation") => root != Kind::Annotation,
        (Ns::Office, b"forms" | b"binary-data") => true,
        (Ns::Draw, _) => !matches!(root, Kind::Frame | Kind::TextBox | Kind::Image),
        (Ns::Svg, b"title" | b"desc") | (Ns::Dc, _) | (Ns::Meta, b"date-string") => true,
        _ => false,
    }
}

fn walk(prefixes: &Prefixes, root: Kind, content: &str, out: &mut String) -> Result<()> {
    let mut reader = Reader::from_str(content);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut state = Walk {
        skip_depth: 0,
        blocks: 0,
        pending_break: false,
        last_space: true,
    };
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("ODT text markup: {error}")))?;
        let (tag, empty) = match event {
            Event::Start(tag) => (tag, false),
            Event::Empty(tag) => (tag, true),
            Event::End(_) => {
                state.skip_depth = state.skip_depth.saturating_sub(1);
                continue;
            },
            Event::Text(text) => {
                if state.skip_depth == 0 {
                    let raw = std::str::from_utf8(text.as_ref()).map_err(|_error| {
                        Error::InvalidFormat("ODT text is not UTF-8".to_string())
                    })?;
                    let collapsed = collapse(raw, state.last_space);
                    if !collapsed.is_empty() {
                        state.push(out, &collapsed);
                    }
                }
                check_len(out)?;
                continue;
            },
            Event::GeneralRef(reference) => {
                if state.skip_depth == 0 {
                    let raw = std::str::from_utf8(reference.as_ref()).unwrap_or_default();
                    let entity = format!("&{raw};");
                    state.push(out, &unescape(&entity));
                }
                continue;
            },
            Event::CData(data) => {
                if state.skip_depth == 0 {
                    state.push(out, &String::from_utf8_lossy(data.as_ref()));
                }
                check_len(out)?;
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        if state.skip_depth > 0 {
            if !empty {
                state.skip_depth += 1;
            }
            continue;
        }
        let name = tag.name();
        match prefixes.resolve(name.as_ref()) {
            Some((ns, local)) if skipped(root, ns, local) => {
                if !empty {
                    state.skip_depth = 1;
                }
            },
            Some((Ns::Text, b"s")) => {
                let count = super::xml::attributes(tag_text(&reader, content))
                    .find_map(|(name, value)| {
                        prefixes
                            .is(name.as_bytes(), Ns::Text, "c")
                            .then(|| value.parse::<usize>().ok())
                            .flatten()
                    })
                    .unwrap_or(1)
                    .min(MAX_TEXT_BYTES);
                state.push(out, &" ".repeat(count));
                check_len(out)?;
            },
            Some((Ns::Text, b"tab")) => state.push(out, "\t"),
            Some((Ns::Text, b"line-break")) => {
                state.push(out, "\n");
                state.last_space = true;
            },
            Some((Ns::Text, b"p" | b"h")) => {
                if state.blocks > 0 {
                    state.pending_break = true;
                }
                state.blocks += 1;
                state.last_space = true;
            },
            _ => {},
        }
    }
    Ok(())
}

struct Walk {
    skip_depth: usize,
    blocks: usize,
    pending_break: bool,
    last_space: bool,
}

impl Walk {
    fn push(&mut self, out: &mut String, text: &str) {
        if self.pending_break {
            out.push('\n');
            self.pending_break = false;
        }
        out.push_str(text);
        if let Some(last) = text.chars().last() {
            self.last_space = last == ' ';
        }
    }
}

fn check_len(out: &str) -> Result<()> {
    if out.len() > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODT node text exceeds {MAX_TEXT_BYTES} bytes"
        )));
    }
    Ok(())
}

/// The markup of the event the reader just returned.
fn tag_text<'c>(reader: &Reader<&[u8]>, content: &'c str) -> &'c str {
    let end = usize::try_from(reader.buffer_position()).unwrap_or(content.len());
    let start = content[..end].rfind('<').unwrap_or(0);
    &content[start..end]
}

/// ODF white-space processing: every run of white space becomes one space,
/// and a space after a space is dropped.
pub(super) fn collapse(raw: &str, mut last_space: bool) -> String {
    let mut out = String::with_capacity(raw.len());
    for character in raw.chars() {
        if matches!(character, ' ' | '\t' | '\n' | '\r') {
            if !last_space {
                out.push(' ');
                last_space = true;
            }
        } else {
            out.push(character);
            last_space = false;
        }
    }
    out
}

pub(super) fn child_element_text<'a>(node: Node<'a>, ns: Ns, local: &str) -> Option<Cow<'a, str>> {
    let prefixes = node.body.prefixes();
    let content = &node.body.content()[node.content_range()];
    let mut reader = Reader::from_str(content);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut depth = 0usize;
    let mut capture: Option<String> = None;
    loop {
        let event = reader.read_event().ok()?;
        match event {
            Event::Start(tag) => {
                if depth == 0 && capture.is_none() && prefixes.is(tag.name().as_ref(), ns, local) {
                    capture = Some(String::new());
                }
                depth += 1;
            },
            Event::Empty(tag) => {
                if depth == 0 && capture.is_none() && prefixes.is(tag.name().as_ref(), ns, local) {
                    return Some(Cow::Borrowed(""));
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    if let Some(text) = capture.take() {
                        return Some(Cow::Owned(text));
                    }
                }
            },
            Event::Text(text) => {
                if let Some(capture) = capture.as_mut() {
                    capture.push_str(std::str::from_utf8(text.as_ref()).ok()?);
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(capture) = capture.as_mut() {
                    let entity = format!("&{};", std::str::from_utf8(reference.as_ref()).ok()?);
                    capture.push_str(&unescape(&entity));
                }
            },
            Event::Eof => return None,
            _ => {},
        }
    }
}
