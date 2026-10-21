//! What a text set keeps (record 0304).
//!
//! A text set replaces the *text* of a node: its character data, `text:s`,
//! `text:tab` and `text:line-break`, and the `text:span` wrappers that hold
//! that text, formatting included. Everything else inside
//! the node is content the set does not address, and is either kept or
//! refused, never dropped:
//!
//! - An element that adds nothing to [`super::Node::text`] — a frame or
//!   drawing shape, a note, an annotation and its end marker, a bookmark, a
//!   reference, index or change mark, a soft page break, a field that shows
//!   no text — is **kept**. It moves out of any span and goes
//!   to the character offset it had in the old text, clamped to the length
//!   of the new text. Elements at one offset keep their order. Setting the
//!   old text back therefore returns every kept element to its offset.
//! - An element that adds text — a text field such as `text:page-number`,
//!   `text:date` or `text:sequence`, a hyperlink (`text:a`) inside the
//!   node, `text:meta`, `text:ruby`, a frame with text inside a text box —
//!   is **refused**: keeping it would add its text
//!   to the new text, and replacing it would lose the element.
//! - In a block container (a cell, list item, text box, note or
//!   annotation) the new text becomes one paragraph. A block child other
//!   than a paragraph or heading (a table, a list, a section) is refused;
//!   bookmarks, change marks and soft page breaks at block level are kept
//!   like inline elements.
//!
//! Offsets count the characters of the node's text as [`super::Node::text`]
//! reports it, with one line feed between block paragraphs.

use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::text::{collapse, skipped};
use super::xml::{Ns, unescape};
use super::{Kind, Node};

/// One element a text set keeps: its source bytes and the character offset
/// of the old text it sat at.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Kept {
    pub(super) offset: usize,
    pub(super) range: Range<usize>,
}

/// The elements a text set of `node` keeps, in document order, or the
/// refusal naming the first element it can neither keep nor replace.
pub(super) fn plan(node: Node<'_>) -> Result<Vec<Kept>> {
    let root = node.kind();
    let block = match root {
        Kind::Cell | Kind::ListItem | Kind::TextBox | Kind::Annotation => Some(0),
        // A note's paragraphs sit in its `text:note-body`.
        Kind::Footnote | Kind::Endnote => Some(1),
        _ => None,
    };
    let range = node.content_range();
    let content = &node.body.content()[range.clone()];
    // Character data, spaces, tabs and breaks only: nothing to keep.
    if !content.as_bytes().contains(&b'<') {
        return Ok(Vec::new());
    }
    let prefixes = node.body.prefixes();
    let mut reader = Reader::from_str(content);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut walk = Walk {
        chars: 0,
        last_space: true,
        pending_break: false,
        blocks: 0,
    };
    let mut kept = Vec::new();
    let mut depth = 0usize;
    // An element skipped whole: its start, its depth, and whether it is kept.
    let mut skip: Option<(usize, usize, bool)> = None;
    // An element kept only if it adds no text: start, depth, offset, name.
    let mut object: Option<(usize, usize, usize, String)> = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("ODT text markup: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).unwrap_or(content.len());
        let (tag, empty) = match event {
            Event::Start(tag) => (tag, false),
            Event::Empty(tag) => (tag, true),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if let Some((begin, at, keep)) = skip {
                    if at == depth {
                        skip = None;
                        if keep {
                            kept.push(Kept {
                                offset: walk.chars,
                                range: range.start + begin..range.start + end,
                            });
                        }
                    }
                } else if let Some((begin, at, offset, name)) = object.take() {
                    if at == depth {
                        let element = range.start + begin..range.start + end;
                        close_object(&mut kept, &walk, offset, &name, root, element)?;
                    } else {
                        object = Some((begin, at, offset, name));
                    }
                }
                continue;
            },
            Event::Text(text) => {
                if skip.is_none() {
                    let raw = std::str::from_utf8(text.as_ref())
                        .map_err(|_error| Error::InvalidFormat("ODT text is not UTF-8".into()))?;
                    let collapsed = collapse(raw, walk.last_space);
                    walk.push(&collapsed);
                }
                continue;
            },
            Event::GeneralRef(reference) => {
                if skip.is_none() {
                    let raw = std::str::from_utf8(reference.as_ref()).unwrap_or_default();
                    walk.push(&unescape(&format!("&{raw};")));
                }
                continue;
            },
            Event::CData(data) => {
                if skip.is_none() {
                    walk.push(&String::from_utf8_lossy(data.as_ref()));
                }
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        let opens = !empty;
        if skip.is_some() {
            depth += usize::from(opens);
            continue;
        }
        let name = tag.name();
        let resolved = prefixes.resolve(name.as_ref());
        let element = || String::from_utf8_lossy(name.as_ref()).into_owned();
        let at_block = object.is_none() && block.is_some_and(|level| level == depth);
        let in_note = matches!(root, Kind::Footnote | Kind::Endnote) && depth == 0;
        let class = if in_note {
            match resolved {
                // The renderer copies the citation itself.
                Some((Ns::Text, b"note-citation")) => Class::Skip,
                Some((Ns::Text, b"note-body")) => Class::Wrapper,
                _ => return Err(refusal(root, &element(), "a note child")),
            }
        } else if at_block {
            match resolved {
                Some((Ns::Text, b"p" | b"h")) => {
                    if walk.blocks > 0 {
                        walk.pending_break = true;
                    }
                    walk.blocks += 1;
                    walk.last_space = true;
                    Class::Wrapper
                },
                // The renderer copies an annotation's author and date.
                Some((Ns::Dc, _) | (Ns::Meta, b"date-string")) if root == Kind::Annotation => {
                    Class::Skip
                },
                Some((ns, local)) if is_marker(ns, local) => Class::Keep,
                _ => return Err(refusal(root, &element(), "a block element")),
            }
        } else {
            match resolved {
                Some((ns, local)) if skipped(root, ns, local) => {
                    if object.is_some() {
                        Class::Skip
                    } else {
                        Class::Keep
                    }
                },
                Some((Ns::Text, b"s")) => {
                    let count = super::xml::attributes(&content[start..end])
                        .find_map(|(name, value)| {
                            prefixes
                                .is(name.as_bytes(), Ns::Text, "c")
                                .then(|| value.parse::<usize>().ok())
                                .flatten()
                        })
                        .unwrap_or(1)
                        .min(super::MAX_TEXT_BYTES);
                    walk.push_spaces(count);
                    Class::Wrapper
                },
                Some((Ns::Text, b"tab")) => {
                    walk.push("\t");
                    Class::Wrapper
                },
                Some((Ns::Text, b"line-break")) => {
                    walk.push("\n");
                    walk.last_space = true;
                    Class::Wrapper
                },
                Some((Ns::Text, b"span")) => Class::Wrapper,
                Some((Ns::Text, b"p" | b"h")) if object.is_some() => {
                    if walk.blocks > 0 {
                        walk.pending_break = true;
                    }
                    walk.blocks += 1;
                    walk.last_space = true;
                    Class::Wrapper
                },
                _ if object.is_some() => Class::Wrapper,
                _ => Class::Object,
            }
        };
        match class {
            Class::Wrapper => depth += usize::from(opens),
            Class::Skip | Class::Keep => {
                let keep = class == Class::Keep;
                if opens {
                    skip = Some((start, depth, keep));
                    depth += 1;
                } else if keep {
                    kept.push(Kept {
                        offset: walk.chars,
                        range: range.start + start..range.start + end,
                    });
                }
            },
            Class::Object => {
                if opens {
                    object = Some((start, depth, walk.chars, element()));
                    depth += 1;
                } else {
                    // An empty element adds no text.
                    kept.push(Kept {
                        offset: walk.chars,
                        range: range.start + start..range.start + end,
                    });
                }
            },
        }
    }
    Ok(kept)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Class {
    /// Text or a wrapper of text: replaced.
    Wrapper,
    /// Left out of the plan (the renderer handles it).
    Skip,
    /// Kept whole.
    Keep,
    /// Kept when it adds no text, refused otherwise.
    Object,
}

/// Close an [`Class::Object`] element: keep it when it added no text.
fn close_object(
    kept: &mut Vec<Kept>,
    walk: &Walk,
    offset: usize,
    name: &str,
    root: Kind,
    range: Range<usize>,
) -> Result<()> {
    if walk.chars != offset {
        return Err(refusal(root, name, "an element with text of its own"));
    }
    kept.push(Kept { offset, range });
    Ok(())
}

/// Zero-text markers that may sit between block paragraphs.
fn is_marker(ns: Ns, local: &[u8]) -> bool {
    match ns {
        Ns::Text => matches!(
            local,
            b"bookmark"
                | b"bookmark-start"
                | b"bookmark-end"
                | b"change"
                | b"change-start"
                | b"change-end"
                | b"soft-page-break"
        ),
        _ => false,
    }
}

/// A text-set refusal naming the replaced element when it is an ODF
/// element the set can meet (a schema name, never authored content).
fn refusal(_root: Kind, element: &str, _what: &str) -> Error {
    macro_rules! reason {
        ($name:literal) => {
            concat!(
                "an ODT text set would replace ",
                $name,
                ": set the text of the runs around it, or remove it first (record 0304)"
            )
        };
    }
    let reason = match element {
        "text:date" => reason!("text:date"),
        "text:time" => reason!("text:time"),
        "text:page-number" => reason!("text:page-number"),
        "text:page-continuation" => reason!("text:page-continuation"),
        "text:sender-firstname" => reason!("text:sender-firstname"),
        "text:sender-lastname" => reason!("text:sender-lastname"),
        "text:sender-initials" => reason!("text:sender-initials"),
        "text:sender-title" => reason!("text:sender-title"),
        "text:sender-position" => reason!("text:sender-position"),
        "text:sender-email" => reason!("text:sender-email"),
        "text:sender-phone-private" => reason!("text:sender-phone-private"),
        "text:sender-fax" => reason!("text:sender-fax"),
        "text:sender-company" => reason!("text:sender-company"),
        "text:sender-phone-work" => reason!("text:sender-phone-work"),
        "text:sender-street" => reason!("text:sender-street"),
        "text:sender-city" => reason!("text:sender-city"),
        "text:sender-postal-code" => reason!("text:sender-postal-code"),
        "text:sender-country" => reason!("text:sender-country"),
        "text:sender-state-or-province" => reason!("text:sender-state-or-province"),
        "text:author-name" => reason!("text:author-name"),
        "text:author-initials" => reason!("text:author-initials"),
        "text:chapter" => reason!("text:chapter"),
        "text:file-name" => reason!("text:file-name"),
        "text:template-name" => reason!("text:template-name"),
        "text:sheet-name" => reason!("text:sheet-name"),
        "text:variable-set" => reason!("text:variable-set"),
        "text:variable-get" => reason!("text:variable-get"),
        "text:variable-input" => reason!("text:variable-input"),
        "text:user-field-get" => reason!("text:user-field-get"),
        "text:user-field-input" => reason!("text:user-field-input"),
        "text:sequence" => reason!("text:sequence"),
        "text:expression" => reason!("text:expression"),
        "text:text-input" => reason!("text:text-input"),
        "text:initial-creator" => reason!("text:initial-creator"),
        "text:creation-date" => reason!("text:creation-date"),
        "text:creation-time" => reason!("text:creation-time"),
        "text:description" => reason!("text:description"),
        "text:user-defined" => reason!("text:user-defined"),
        "text:print-time" => reason!("text:print-time"),
        "text:print-date" => reason!("text:print-date"),
        "text:printed-by" => reason!("text:printed-by"),
        "text:title" => reason!("text:title"),
        "text:subject" => reason!("text:subject"),
        "text:keywords" => reason!("text:keywords"),
        "text:editing-cycles" => reason!("text:editing-cycles"),
        "text:editing-duration" => reason!("text:editing-duration"),
        "text:modification-time" => reason!("text:modification-time"),
        "text:modification-date" => reason!("text:modification-date"),
        "text:creator" => reason!("text:creator"),
        "text:page-count" => reason!("text:page-count"),
        "text:paragraph-count" => reason!("text:paragraph-count"),
        "text:word-count" => reason!("text:word-count"),
        "text:character-count" => reason!("text:character-count"),
        "text:table-count" => reason!("text:table-count"),
        "text:image-count" => reason!("text:image-count"),
        "text:object-count" => reason!("text:object-count"),
        "text:database-display" => reason!("text:database-display"),
        "text:database-next" => reason!("text:database-next"),
        "text:database-row-select" => reason!("text:database-row-select"),
        "text:database-row-number" => reason!("text:database-row-number"),
        "text:database-name" => reason!("text:database-name"),
        "text:page-variable-set" => reason!("text:page-variable-set"),
        "text:page-variable-get" => reason!("text:page-variable-get"),
        "text:placeholder" => reason!("text:placeholder"),
        "text:conditional-text" => reason!("text:conditional-text"),
        "text:hidden-text" => reason!("text:hidden-text"),
        "text:hidden-paragraph" => reason!("text:hidden-paragraph"),
        "text:reference-ref" => reason!("text:reference-ref"),
        "text:bookmark-ref" => reason!("text:bookmark-ref"),
        "text:note-ref" => reason!("text:note-ref"),
        "text:sequence-ref" => reason!("text:sequence-ref"),
        "text:script" => reason!("text:script"),
        "text:execute-macro" => reason!("text:execute-macro"),
        "text:dde-connection" => reason!("text:dde-connection"),
        "text:measure" => reason!("text:measure"),
        "text:table-formula" => reason!("text:table-formula"),
        "text:meta-field" => reason!("text:meta-field"),
        "text:meta" => reason!("text:meta"),
        "text:ruby" => reason!("text:ruby"),
        "text:note" => reason!("text:note"),
        "text:note-body" => reason!("text:note-body"),
        "text:note-citation" => reason!("text:note-citation"),
        "text:a" => reason!("text:a"),
        "text:span" => reason!("text:span"),
        "text:h" => reason!("text:h"),
        "text:p" => reason!("text:p"),
        "text:list" => reason!("text:list"),
        "text:list-item" => reason!("text:list-item"),
        "text:section" => reason!("text:section"),
        "text:soft-page-break" => reason!("text:soft-page-break"),
        "text:bibliography-mark" => reason!("text:bibliography-mark"),
        "text:toc-mark" => reason!("text:toc-mark"),
        "text:alphabetical-index-mark" => reason!("text:alphabetical-index-mark"),
        "text:user-index-mark" => reason!("text:user-index-mark"),
        "text:drop-down" => reason!("text:drop-down"),
        "text:annotation" => reason!("text:annotation"),
        "office:annotation" => reason!("office:annotation"),
        "office:annotation-end" => reason!("office:annotation-end"),
        "draw:frame" => reason!("draw:frame"),
        "draw:a" => reason!("draw:a"),
        "draw:custom-shape" => reason!("draw:custom-shape"),
        "draw:rect" => reason!("draw:rect"),
        "draw:line" => reason!("draw:line"),
        "draw:ellipse" => reason!("draw:ellipse"),
        "draw:control" => reason!("draw:control"),
        "draw:g" => reason!("draw:g"),
        "table:table" => reason!("table:table"),
        "table:table-row" => reason!("table:table-row"),
        "table:table-cell" => reason!("table:table-cell"),
        "table:covered-table-cell" => reason!("table:covered-table-cell"),
        _ => {
            "an ODT text set would replace a non-text element: set the text of the runs around it, \
              or remove it first (record 0304)"
        },
    };
    Error::Refused(litchi_core::edit::Refusal::new(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "odt.body.set_text",
        reason,
    ))
}

/// The character count of the text [`super::text`] reports, without the
/// text itself.
struct Walk {
    chars: usize,
    last_space: bool,
    pending_break: bool,
    blocks: usize,
}

impl Walk {
    fn push(&mut self, text: &str) {
        let Some(last) = text.chars().last() else {
            return;
        };
        if self.pending_break {
            self.chars += 1;
            self.pending_break = false;
        }
        self.chars += text.chars().count();
        self.last_space = last == ' ';
    }

    fn push_spaces(&mut self, count: usize) {
        if count == 0 {
            return;
        }
        if self.pending_break {
            self.chars += 1;
            self.pending_break = false;
        }
        self.chars += count;
        self.last_space = true;
    }
}

/// Encode `text` with the kept elements spliced in at their offsets.
pub(super) fn encode(
    out: &mut String,
    prefixes: &super::xml::Prefixes,
    content: &str,
    text: &str,
    kept: &[Kept],
) {
    if kept.is_empty() {
        super::xml::encode_text(out, prefixes, text);
        return;
    }
    let mut cursor = 0usize; // byte index into `text`
    let mut count = 0usize; // characters before `cursor`
    let mut written = 0usize;
    for element in kept {
        while count < element.offset && cursor < text.len() {
            cursor += text[cursor..].chars().next().map_or(1, char::len_utf8);
            count += 1;
        }
        if cursor > written {
            super::xml::encode_text(out, prefixes, &text[written..cursor]);
            written = cursor;
        }
        out.push_str(&content[element.range.clone()]);
    }
    if written < text.len() {
        super::xml::encode_text(out, prefixes, &text[written..]);
    }
}

#[cfg(test)]
mod tests {
    use super::super::{Kind, Tree};
    use super::*;

    const HEAD: &str = concat!(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:dc="http://purl.org/dc/elements/1.1/" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink">"#,
        "<office:body><office:text>"
    );
    const TAIL: &str = "</office:text></office:body></office:document-content>";
    const FRAME: &str =
        r#"<draw:frame draw:name="f1"><draw:image xlink:href="Pictures/a.png"/></draw:frame>"#;

    /// Set the text of the first `kind` node of `body` and return the new
    /// node's text and markup, or the refusal.
    fn set(body: &str, kind: Kind, text: &str) -> Result<(String, String)> {
        let content = format!("{HEAD}{body}{TAIL}");
        let tree = Tree::build(&content, None)?;
        let view = tree.view(&content, None);
        let node = view.nodes().find(|node| node.kind() == kind).unwrap();
        let kept = plan(node)?;
        let range = node.content_range();
        let mut inner = String::new();
        if matches!(kind, Kind::Cell | Kind::ListItem | Kind::TextBox) {
            inner.push_str("<text:p>");
            encode(&mut inner, view.prefixes(), &content, text, &kept);
            inner.push_str("</text:p>");
        } else {
            encode(&mut inner, view.prefixes(), &content, text, &kept);
        }
        let after = format!(
            "{}{inner}{}",
            &content[..range.start],
            &content[range.end..]
        );
        let tree = Tree::build(&after, None)?;
        let view = tree.view(&after, None);
        let node = view.nodes().find(|node| node.kind() == kind).unwrap();
        Ok((node.text()?.into_owned(), inner))
    }

    #[test]
    fn frames_keep_their_offsets() {
        let body = format!("<text:p>{FRAME}Old <text:span>text</text:span>{FRAME}</text:p>");
        let (text, markup) = set(&body, Kind::Paragraph, "Brand new text").unwrap();
        assert_eq!(text, "Brand new text");
        // The first frame stays at 0, the second at offset 8 ("Brand ne").
        assert_eq!(markup, format!("{FRAME}Brand ne{FRAME}w text"));
        // The old text puts both back where they were.
        let (text, markup) = set(&body, Kind::Paragraph, "Old text").unwrap();
        assert_eq!(text, "Old text");
        assert_eq!(markup, format!("{FRAME}Old text{FRAME}"));
    }

    #[test]
    fn offsets_clamp_to_shorter_text() {
        let body = format!("<text:p>Long text{FRAME}</text:p>");
        let (text, markup) = set(&body, Kind::Paragraph, "ab").unwrap();
        assert_eq!(text, "ab");
        assert_eq!(markup, format!("ab{FRAME}"));
        let (text, markup) = set(&body, Kind::Paragraph, "").unwrap();
        assert_eq!(text, "");
        assert_eq!(markup, FRAME);
    }

    #[test]
    fn markers_and_notes_are_kept() {
        let body = concat!(
            r#"<text:p><text:bookmark-start text:name="b"/>Hi"#,
            r#"<text:note text:id="n1"><text:note-citation>1</text:note-citation>"#,
            r#"<text:note-body><text:p>Note</text:p></text:note-body></text:note>"#,
            r#"<text:bookmark-end text:name="b"/> there<text:soft-page-break/></text:p>"#,
        );
        let (text, markup) = set(body, Kind::Paragraph, "Hello world").unwrap();
        assert_eq!(text, "Hello world");
        assert!(markup.starts_with(r#"<text:bookmark-start text:name="b"/>He<text:note "#));
        assert!(
            markup.contains(r#"</text:note><text:bookmark-end text:name="b"/>llo wo"#),
            "{markup}"
        );
        // "Hi there" put the page break at offset 8.
        assert!(markup.ends_with("llo wo<text:soft-page-break/>rld"));
    }

    #[test]
    fn fields_with_text_are_refused() {
        let body = "<text:p>Page <text:page-number>3</text:page-number></text:p>";
        let error = set(body, Kind::Paragraph, "x").unwrap_err();
        assert!(
            matches!(&error, Error::Refused(refusal) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
        );
        // A field that shows nothing is kept.
        let body = "<text:p>Page <text:page-number></text:page-number></text:p>";
        let (text, markup) = set(body, Kind::Paragraph, "Page x").unwrap();
        assert_eq!(text, "Page x");
        assert_eq!(
            markup,
            "Page<text:s/><text:page-number></text:page-number>x"
        );
    }

    #[test]
    fn cells_merge_paragraphs_and_keep_frames() {
        let body = format!(
            "<table:table><table:table-row><table:table-cell>\
             <text:p>One</text:p><text:p>{FRAME}Two</text:p>\
             </table:table-cell></table:table-row></table:table>"
        );
        let (text, markup) = set(&body, Kind::Cell, "New").unwrap();
        assert_eq!(text, "New");
        // The frame sat at offset 4 ("One\n"), clamped to 3.
        assert_eq!(markup, format!("<text:p>New{FRAME}</text:p>"));
        let nested = "<table:table><table:table-row><table:table-cell>\
             <table:table/><text:p>x</text:p></table:table-cell></table:table-row></table:table>";
        let error = set(nested, Kind::Cell, "y").unwrap_err();
        assert!(
            matches!(&error, Error::Refused(refusal) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
        );
    }

    #[test]
    fn hyperlinks_inside_are_refused() {
        let body =
            r#"<text:p>See <text:a xlink:href="https://example.com/">this</text:a></text:p>"#;
        let error = set(body, Kind::Paragraph, "x").unwrap_err();
        assert!(
            matches!(&error, Error::Refused(refusal) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
        );
        // The hyperlink's own text is set.
        let (text, _) = set(body, Kind::Hyperlink, "that").unwrap();
        assert_eq!(text, "that");
    }

    #[test]
    fn nested_lists_are_refused() {
        let body = "<text:list><text:list-item><text:p>a</text:p>\
             <text:list><text:list-item><text:p>b</text:p></text:list-item></text:list>\
             </text:list-item></text:list>";
        let error = set(body, Kind::ListItem, "y").unwrap_err();
        assert!(
            matches!(&error, Error::Refused(refusal) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
        );
    }
}
