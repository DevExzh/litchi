//! Text fields of the body (record 0295): the `text:*` field elements a
//! paragraph holds inline (`text:date`, `text:page-number`, `text:title`,
//! ...), read with their displayed text, and the staged replacement of
//! that text ([`crate::transaction::Edit::set_field_text`]). Fields are not
//! indexed body nodes, so they are listed by one scan of the body.

use std::borrow::Cow;
use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::xml::{self, Ns};
use super::{Body, Kind};

/// The `text:*` field elements [`text_fields`] lists.
const FIELDS: [&str; 26] = [
    "date",
    "time",
    "page-number",
    "page-count",
    "page-continuation",
    "title",
    "subject",
    "description",
    "keywords",
    "initial-creator",
    "creator",
    "author-name",
    "author-initials",
    "creation-date",
    "creation-time",
    "modification-date",
    "modification-time",
    "print-date",
    "print-time",
    "word-count",
    "character-count",
    "paragraph-count",
    "table-count",
    "image-count",
    "object-count",
    "file-name",
];

/// Every `text:*` field element [`fields`] lists (record 0328): the
/// [`text_fields`] set, then the variable, reference, sequence, chapter,
/// sender, template, database and conditional fields.
const ALL_FIELDS: [&str; 59] = [
    "date",
    "time",
    "page-number",
    "page-count",
    "page-continuation",
    "title",
    "subject",
    "description",
    "keywords",
    "initial-creator",
    "creator",
    "author-name",
    "author-initials",
    "creation-date",
    "creation-time",
    "modification-date",
    "modification-time",
    "print-date",
    "print-time",
    "word-count",
    "character-count",
    "paragraph-count",
    "table-count",
    "image-count",
    "object-count",
    "file-name",
    "editing-cycles",
    "editing-duration",
    "printed-by",
    "user-defined",
    "variable-set",
    "variable-get",
    "variable-input",
    "user-field-get",
    "user-field-input",
    "sequence",
    "expression",
    "text-input",
    "page-variable-set",
    "page-variable-get",
    "placeholder",
    "chapter",
    "template-name",
    "sheet-name",
    "sender-firstname",
    "sender-lastname",
    "sender-initials",
    "sender-company",
    "sender-email",
    "sender-title",
    "bookmark-ref",
    "reference-ref",
    "sequence-ref",
    "note-ref",
    "conditional-text",
    "hidden-text",
    "database-display",
    "database-name",
    "execute-macro",
];

/// One text field of the body.
#[derive(Debug, Clone)]
pub struct TextField<'a> {
    name: &'a str,
    range: Range<usize>,
    start_tag: &'a str,
    content: &'a str,
    body: Body<'a>,
}

impl<'a> TextField<'a> {
    /// The field element of `body` at `range` (a [`Self::range`] of a field
    /// [`fields`] listed), or `None` when no field element spans exactly
    /// those bytes.
    #[must_use]
    pub fn at(body: Body<'a>, range: Range<usize>) -> Option<Self> {
        let content = body.content();
        let element = content.get(range.clone())?;
        let tag_end = xml::tag_end(element.as_bytes(), 0)?;
        let start_tag = &element[..tag_end];
        let qualified = xml::tag_name(start_tag);
        let (ns, local) = body.prefixes().resolve(qualified.as_bytes())?;
        if ns != Ns::Text {
            return None;
        }
        let name = ALL_FIELDS
            .iter()
            .copied()
            .find(|name| name.as_bytes() == local)?;
        let inner = if start_tag.ends_with("/>") {
            ""
        } else {
            let close = element.rfind("</")?;
            element.get(tag_end..close)?
        };
        Some(Self {
            name,
            range,
            start_tag,
            content: inner,
            body,
        })
    }

    /// The element's local name: `date`, `page-number`, `title`, ...
    #[must_use]
    pub const fn name(&self) -> &'a str {
        self.name
    }

    /// The element's bytes in `content.xml`.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// The decoded value of the attribute `local` in the text or style
    /// namespace (`text:fixed`, `style:data-style-name`, `style:num-format`,
    /// `text:select-page`).
    #[must_use]
    pub fn attribute(&self, local: &str) -> Option<Cow<'a, str>> {
        let prefixes = self.body.prefixes();
        xml::attributes(self.start_tag).find_map(|(name, value)| {
            (prefixes.is(name.as_bytes(), Ns::Text, local)
                || prefixes.is(name.as_bytes(), Ns::Style, local))
            .then(|| xml::unescape(value))
        })
    }

    /// The displayed text: the element's character data, with `text:s`,
    /// `text:tab` and `text:line-break` as spaces, a tab and a line feed.
    #[must_use]
    pub fn text(&self) -> String {
        let prefixes = self.body.prefixes();
        let mut out = String::new();
        let mut rest = self.content;
        while let Some(start) = rest.find('<') {
            out.push_str(&xml::unescape(&rest[..start]));
            let tag = &rest[start..];
            let end = tag.find('>').map_or(tag.len(), |end| end + 1);
            let name = xml::tag_name(&tag[..end]);
            match prefixes.resolve(name.as_bytes()) {
                Some((Ns::Text, b"s")) => out.push(' '),
                Some((Ns::Text, b"tab")) => out.push('\t'),
                Some((Ns::Text, b"line-break")) => out.push('\n'),
                _ => {},
            }
            rest = &tag[end..];
        }
        out.push_str(&xml::unescape(rest));
        out
    }
}

/// Every text field of `body` in document order, tracked-change records
/// left out.
///
/// # Errors
///
/// Returns an error for malformed body markup.
pub fn text_fields(body: Body<'_>) -> Result<Vec<TextField<'_>>> {
    scan(body, &FIELDS)
}

/// Every text field of `body` in document order — the [`text_fields`] set
/// and the variable, reference, sequence, chapter, sender, template,
/// database and conditional fields — tracked-change records left out. A
/// field inside another field is not listed.
///
/// # Errors
///
/// Returns an error for malformed body markup.
pub fn fields(body: Body<'_>) -> Result<Vec<TextField<'_>>> {
    scan(body, &ALL_FIELDS)
}

fn scan<'a>(body: Body<'a>, names: &'static [&'static str]) -> Result<Vec<TextField<'a>>> {
    let root = body.root();
    debug_assert_eq!(root.kind(), Kind::Body);
    let content = body.content();
    let base = root.range().start;
    let region = &content[root.range()];
    let prefixes = body.prefixes();
    let mut reader = Reader::from_str(region);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    let mut fields = Vec::new();
    // An open field: its name, start and start-tag end.
    let mut open: Option<(&str, usize, usize, usize)> = None;
    let mut depth = 0usize;
    let mut skip = 0usize;
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("ODT body markup: {error}")))?;
        let after = usize::try_from(reader.buffer_position()).unwrap_or(region.len());
        match event {
            Event::Start(tag) => {
                depth += 1;
                if skip > 0 {
                    skip += 1;
                    continue;
                }
                let qualified = tag.name();
                let resolved = prefixes.resolve(qualified.as_ref());
                if resolved == Some((Ns::Text, b"tracked-changes".as_slice())) {
                    skip = 1;
                    continue;
                }
                if open.is_none()
                    && let Some((Ns::Text, local)) = resolved
                    && let Some(name) = names.iter().find(|name| name.as_bytes() == local)
                {
                    open = Some((name, before, after, depth));
                }
            },
            Event::Empty(tag) => {
                if skip > 0 || open.is_some() {
                    continue;
                }
                let qualified = tag.name();
                if let Some((Ns::Text, local)) = prefixes.resolve(qualified.as_ref())
                    && let Some(name) = names.iter().find(|name| name.as_bytes() == local)
                {
                    fields.push(TextField {
                        name,
                        range: base + before..base + after,
                        start_tag: &region[before..after],
                        content: "",
                        body,
                    });
                }
            },
            Event::End(_) => {
                if skip > 0 {
                    skip -= 1;
                } else if let Some((name, start, tag_end, level)) = open
                    && level == depth
                {
                    fields.push(TextField {
                        name,
                        range: base + start..base + after,
                        start_tag: &region[start..tag_end],
                        content: &region[tag_end..before],
                        body,
                    });
                    open = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(fields)
}

/// The markup of the field element at `range` with its content replaced by
/// `text`.
pub(crate) fn render(body: Body<'_>, range: Range<usize>, text: &str) -> String {
    let element = &body.content()[range];
    let tag = element.find('>').map_or(element, |end| &element[..=end]);
    let mut out = match tag.strip_suffix("/>") {
        Some(open) => format!("{}>", open.trim_end()),
        None => tag.to_owned(),
    };
    xml::encode_text(&mut out, body.prefixes(), text);
    out.push_str("</");
    out.push_str(xml::tag_name(tag));
    out.push('>');
    out
}
