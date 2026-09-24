//! Strict recognizer for the common benign `<sheetData>` body.
//!
//! The materializing worksheet parser, the edit scanner and the publication
//! compactor each walk a complete worksheet through `quick_xml`'s
//! namespace-resolving reader. On a dense worksheet almost every event lies
//! inside `<sheetData>`, and for the benign body that ordinary producers emit
//! that generality buys nothing: every element is an unprefixed `row`, `c` or
//! `v` resolved against the one default namespace in scope at `<sheetData>`,
//! no element declares or changes a namespace binding, no attribute value
//! needs entity decoding or whitespace normalization, and no character data
//! needs unescaping.
//!
//! This module recognizes exactly that subset and reports, in document
//! order, the event stream the reader would have produced for it:
//!
//! ```text
//! body  := (ws | row)* "</" sheet-data-name ">"
//! row   := "<row" attrs "/>" | "<row" attrs ">" (ws | cell)* "</row>"
//! cell  := "<c" attrs "/>"   | "<c" attrs ">" ws? value? ws? "</c>"
//! value := "<v/>" | "<v>" text? "</v>"
//! attrs := (" " name "=\"" value-char* "\"")*
//! name  := ascii-name (":" ascii-name)?   -- never `xmlns`, `xmlns:*` or `xml:*`
//! text  := any byte except `<` and `&`
//! ws    := (" " | "\t" | "\r" | "\n")+
//! ```
//!
//! An attribute value may not contain `"`, `<`, `>`, `&`, tab, carriage
//! return or line feed, so the value `quick_xml` decodes and normalizes is the
//! raw byte slice itself. Duplicate attribute names are declined rather than
//! reproduced, so the ordinary reader keeps ownership of that refusal.
//!
//! Anything else — a comment, processing instruction, CDATA section, entity
//! or character reference, prefixed or namespace-declaring element, an
//! `xml:space` attribute, a formula, an inline string, a second value, any
//! other child, single-quoted or irregularly spaced markup, or a truncated
//! document — makes the recognizer *decline*. A decline is never an error:
//! the caller keeps using the ordinary reader for the whole body, so every
//! refusal and diagnostic outside this subset stays exactly where it was.
//!
//! The recognizer is a pure function of the bytes. Callers run it once with a
//! no-op visitor to learn whether the whole body is admitted, and only then
//! replay the events into their own state, so a decline never leaves partial
//! state behind.

use litchi_core::xml::ReaderOrigin;
use memchr::memchr2;
use quick_xml::events::Event as ReaderEvent;
use quick_xml::name::{NamespaceResolver, QName};
use quick_xml::reader::NsReader;

use crate::error::{allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;

/// Upper bound on attributes in one lane tag. Longer tags are declined so
/// the pairwise duplicate check stays a small bounded loop.
const MAX_TAG_ATTRIBUTES: usize = 32;

const ROW_CLOSE: &[u8] = b"</row>";
const CELL_CLOSE: &[u8] = b"</c>";
const VALUE_OPEN: &[u8] = b"<v>";
const VALUE_EMPTY: &[u8] = b"<v/>";
const VALUE_CLOSE: &[u8] = b"</v>";

/// The three element names the lane admits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Name {
    Row,
    Cell,
    Value,
}

/// One recognized start or empty tag.
///
/// `content` is exactly the byte slice `quick_xml` stores in the
/// corresponding `BytesStart`: the element name followed by its raw
/// attributes, excluding the `<`, the closing `>` and an empty element's
/// trailing `/`. `start` and `end` are absolute offsets of `<` and one past
/// `>` in the document the recognizer was given.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Tag<'a> {
    pub(crate) content: &'a [u8],
    pub(crate) start: usize,
    pub(crate) end: usize,
}

impl<'a> Tag<'a> {
    /// Iterate this tag's attributes as raw `(name, value)` pairs.
    ///
    /// For an admitted tag the raw value equals the value `quick_xml`
    /// decodes and normalizes, and names are unique.
    pub(crate) fn attributes(&self) -> Attributes<'a> {
        let name_len = self
            .content
            .iter()
            .position(|byte| *byte == b' ')
            .unwrap_or(self.content.len());
        Attributes {
            rest: &self.content[name_len..],
        }
    }
}

/// Raw attribute iterator over an admitted tag.
#[derive(Debug, Clone)]
pub(crate) struct Attributes<'a> {
    rest: &'a [u8],
}

impl<'a> Iterator for Attributes<'a> {
    type Item = (&'a [u8], &'a [u8]);

    fn next(&mut self) -> Option<Self::Item> {
        // Admitted attributes are exactly ` name="value"`.
        let rest = self.rest.strip_prefix(b" ")?;
        let equals = rest.iter().position(|byte| *byte == b'=')?;
        let name = &rest[..equals];
        let value_and_rest = rest.get(equals + 2..)?;
        let quote = value_and_rest.iter().position(|byte| *byte == b'"')?;
        let value = &value_and_rest[..quote];
        self.rest = &value_and_rest[quote + 1..];
        Some((name, value))
    }
}

/// One event of the admitted body, in the order `quick_xml` reports it.
#[derive(Debug, Clone, Copy)]
pub(crate) enum Event<'a> {
    /// A start tag of an element with content.
    Start(Name, Tag<'a>),
    /// A self-closing tag.
    Empty(Name, Tag<'a>),
    /// A close tag spanning `start..end`.
    End(Name, usize, usize),
    /// Character data spanning `start..end`. `value` is true exactly when the
    /// text is the content of a `<v>` element; otherwise the run is
    /// whitespace between elements.
    Text {
        start: usize,
        end: usize,
        value: bool,
    },
}

/// What an admitted body contains.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct Summary {
    /// Absolute offset of the `<sheetData>` element's close tag.
    pub(crate) end: usize,
    /// Row elements, including empty ones.
    pub(crate) rows: usize,
    /// Cell elements, including empty ones.
    pub(crate) cells: usize,
    /// Events the ordinary reader would report for the body, excluding the
    /// `</sheetData>` close tag itself.
    pub(crate) events: usize,
    /// Whitespace runs between elements, which every pass other than value
    /// text treats as formatting.
    pub(crate) whitespace: usize,
    /// Whether any admitted tag's attributes contain an apostrophe.
    pub(crate) apostrophe: bool,
}

/// Where a reader stopped so a `<sheetData>` body may take the lane: one
/// past the start tag, and the span of the element's qualified name.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Entry {
    pub(crate) position: usize,
    name_start: usize,
    name_end: usize,
}

impl Entry {
    /// Locate the `<sheetData>` start tag a slice reader over `content` just
    /// delivered.
    ///
    /// `event_start` and `position` are the reader's positions before and
    /// after the event and `name` is the element's qualified name. The lane
    /// works in offsets into `content`, so the positions are converted
    /// through `content`'s [`ReaderOrigin`] — a leading UTF-8 byte-order mark
    /// precedes reader position zero — and used only after the tag they claim
    /// to delimit is found at exactly those offsets: `<`, the same name bytes,
    /// and a closing `>`. `None` keeps the ordinary reader.
    pub(crate) fn locate(
        content: &[u8],
        event_start: u64,
        name: &[u8],
        position: u64,
    ) -> Option<Self> {
        let origin = ReaderOrigin::of(content);
        let tag_start = origin.offset(event_start)?;
        let position = origin.offset(position)?;
        if content.get(tag_start) != Some(&b'<') {
            return None;
        }
        let name_start = tag_start.checked_add(1)?;
        let name_end = name_start.checked_add(name.len())?;
        if content.get(name_start..name_end) != Some(name)
            || !matches!(
                content.get(name_end),
                Some(b'>' | b' ' | b'\t' | b'\r' | b'\n')
            )
            || name_end >= position
            || content.get(position.checked_sub(1)?) != Some(&b'>')
        {
            return None;
        }
        Some(Self {
            position,
            name_start,
            name_end,
        })
    }

    /// The element's qualified name, which its close tag repeats.
    pub(crate) fn name<'a>(&self, content: &'a [u8]) -> crate::Result<&'a [u8]> {
        content
            .get(self.name_start..self.name_end)
            .ok_or_else(|| invalid("worksheet lane entry lies outside its document"))
    }
}

/// Whether unprefixed children of the element a reader just started resolve
/// to a `SpreadsheetML` namespace, so the lane's `row`, `c` and `v` are
/// exactly the names the worksheet passes recognize.
pub(crate) fn children_are_spreadsheetml(resolver: &NamespaceResolver) -> bool {
    let name = QName(b"row");
    let (namespace, _) = resolver.resolve_element(name);
    is_spreadsheetml_name(&namespace, name, b"row")
}

/// Rebuild a document without an admitted `<sheetData>` body.
///
/// Rereading `content[..position]` restores exactly the reader state that
/// preceded the body, because the body is balanced and declares no
/// namespace; `content[resume..]` then begins with the close tag.
pub(crate) fn splice_without_body(
    content: &[u8],
    position: usize,
    resume: usize,
) -> crate::Result<Vec<u8>> {
    let prefix = content
        .get(..position)
        .ok_or_else(|| invalid("worksheet lane entry lies outside its document"))?;
    let suffix = content
        .get(resume..)
        .ok_or_else(|| invalid("worksheet lane resume lies outside its document"))?;
    let length = prefix
        .len()
        .checked_add(suffix.len())
        .ok_or_else(|| invalid("worksheet lane resume size overflows usize"))?;
    let mut spliced = Vec::new();
    spliced
        .try_reserve_exact(length)
        .map_err(|source| allocation("worksheet lane resume", source))?;
    spliced.extend_from_slice(prefix);
    spliced.extend_from_slice(suffix);
    Ok(spliced)
}

/// Advance a reader over the spliced prefix whose events were already
/// delivered, stopping exactly at the lane entry.
///
/// `spliced` is the reader's input and `position` the entry's offset in it;
/// the reader's positions are converted through the input's
/// [`ReaderOrigin`], so a byte-order-marked prefix stops at the same entry.
pub(crate) fn skip_to(
    reader: &mut NsReader<&[u8]>,
    spliced: &[u8],
    position: usize,
) -> crate::Result<()> {
    let origin = ReaderOrigin::of(spliced);
    loop {
        let at = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("worksheet XML position does not fit usize"))?;
        if at == position {
            return Ok(());
        }
        if at > position {
            return Err(invalid("worksheet lane resume passed its entry point"));
        }
        if matches!(
            reader
                .read_event()
                .map_err(|error| invalid(error.to_string()))?,
            ReaderEvent::Eof
        ) {
            return Err(invalid(
                "worksheet lane resume ended before its entry point",
            ));
        }
    }
}

/// Recognize the body that starts at `start` (just after the `<sheetData>`
/// start tag) without reporting events. `name` is the element's qualified
/// name exactly as its start tag spells it, which its close tag repeats.
pub(crate) fn recognize(content: &[u8], start: usize, name: &[u8]) -> Option<Summary> {
    let Ok(summary) = walk(content, start, name, &mut |_event| {
        Ok::<(), std::convert::Infallible>(())
    });
    summary
}

/// Walk the body that starts at `start`, reporting each event to `visit`.
///
/// Returns `Ok(None)` when the body leaves the admitted subset. Events that
/// were reported before a decline are not retracted, so a caller that
/// mutates state must first learn from [`recognize`] that the body is
/// admitted.
pub(crate) fn walk<'a, E>(
    content: &'a [u8],
    start: usize,
    name: &[u8],
    visit: &mut impl FnMut(Event<'a>) -> Result<(), E>,
) -> Result<Option<Summary>, E> {
    let mut summary = Summary::default();
    let mut at = start;
    loop {
        at = formatting(content, at, visit, &mut summary)?;
        if close_tag(content, at, name) {
            summary.end = at;
            return Ok(Some(summary));
        }
        let Some((tag, empty)) = start_tag(content, at, b"row", &mut summary) else {
            return Ok(None);
        };
        summary.rows += 1;
        summary.events += 1;
        at = tag.end;
        if empty {
            visit(Event::Empty(Name::Row, tag))?;
            continue;
        }
        visit(Event::Start(Name::Row, tag))?;
        let Some(next) = row_body(content, at, visit, &mut summary)? else {
            return Ok(None);
        };
        at = next;
    }
}

/// Walk one row's content after its start tag, through its close tag.
fn row_body<'a, E>(
    content: &'a [u8],
    mut at: usize,
    visit: &mut impl FnMut(Event<'a>) -> Result<(), E>,
    summary: &mut Summary,
) -> Result<Option<usize>, E> {
    loop {
        at = formatting(content, at, visit, summary)?;
        if content[at.min(content.len())..].starts_with(ROW_CLOSE) {
            let end = at + ROW_CLOSE.len();
            visit(Event::End(Name::Row, at, end))?;
            summary.events += 1;
            return Ok(Some(end));
        }
        let Some((tag, empty)) = start_tag(content, at, b"c", summary) else {
            return Ok(None);
        };
        summary.cells += 1;
        summary.events += 1;
        at = tag.end;
        if empty {
            visit(Event::Empty(Name::Cell, tag))?;
            continue;
        }
        visit(Event::Start(Name::Cell, tag))?;
        let Some(next) = cell_body(content, at, visit, summary)? else {
            return Ok(None);
        };
        at = next;
    }
}

/// Walk one cell's content after its start tag, through its close tag.
fn cell_body<'a, E>(
    content: &'a [u8],
    mut at: usize,
    visit: &mut impl FnMut(Event<'a>) -> Result<(), E>,
    summary: &mut Summary,
) -> Result<Option<usize>, E> {
    at = formatting(content, at, visit, summary)?;
    let rest = &content[at.min(content.len())..];
    if rest.starts_with(VALUE_EMPTY) {
        let end = at + VALUE_EMPTY.len();
        visit(Event::Empty(
            Name::Value,
            Tag {
                content: &content[at + 1..at + 2],
                start: at,
                end,
            },
        ))?;
        summary.events += 1;
        at = end;
    } else if rest.starts_with(VALUE_OPEN) {
        let open_end = at + VALUE_OPEN.len();
        visit(Event::Start(
            Name::Value,
            Tag {
                content: &content[at + 1..at + 2],
                start: at,
                end: open_end,
            },
        ))?;
        summary.events += 1;
        at = open_end;
        // Character data runs to the next markup; an entity or character
        // reference would split it into reader events, so it declines.
        let Some(relative) = memchr2(b'<', b'&', &content[at..]) else {
            return Ok(None);
        };
        let text_end = at + relative;
        if content[text_end] != b'<' {
            return Ok(None);
        }
        if text_end > at {
            visit(Event::Text {
                start: at,
                end: text_end,
                value: true,
            })?;
            summary.events += 1;
        }
        at = text_end;
        if !content[at..].starts_with(VALUE_CLOSE) {
            return Ok(None);
        }
        let end = at + VALUE_CLOSE.len();
        visit(Event::End(Name::Value, at, end))?;
        summary.events += 1;
        at = end;
    }
    at = formatting(content, at, visit, summary)?;
    if !content[at.min(content.len())..].starts_with(CELL_CLOSE) {
        return Ok(None);
    }
    let end = at + CELL_CLOSE.len();
    visit(Event::End(Name::Cell, at, end))?;
    summary.events += 1;
    Ok(Some(end))
}

/// Whether `</name>` starts at `at`.
#[inline]
fn close_tag(content: &[u8], at: usize, name: &[u8]) -> bool {
    content
        .get(at..)
        .and_then(|rest| rest.strip_prefix(b"</"))
        .and_then(|rest| rest.strip_prefix(name))
        .is_some_and(|rest| rest.first() == Some(&b'>'))
}

/// Report and count the whitespace run starting at `at`, if any, and
/// return where the next markup begins.
fn formatting<'a, E>(
    content: &'a [u8],
    at: usize,
    visit: &mut impl FnMut(Event<'a>) -> Result<(), E>,
    summary: &mut Summary,
) -> Result<usize, E> {
    let Some(end) = whitespace(content, at) else {
        return Ok(at);
    };
    visit(Event::Text {
        start: at,
        end,
        value: false,
    })?;
    summary.events += 1;
    summary.whitespace += 1;
    Ok(end)
}

/// Return the end of a non-empty whitespace run starting at `at`.
#[inline]
fn whitespace(content: &[u8], at: usize) -> Option<usize> {
    let rest = content.get(at..)?;
    let length = rest
        .iter()
        .position(|byte| !matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
        .unwrap_or(rest.len());
    (length > 0).then_some(at + length)
}

/// Recognize `<name attrs>` or `<name attrs/>` at `at`.
fn start_tag<'a>(
    content: &'a [u8],
    at: usize,
    name: &[u8],
    summary: &mut Summary,
) -> Option<(Tag<'a>, bool)> {
    let rest = content.get(at..)?;
    if rest.first() != Some(&b'<') || !rest.get(1..)?.starts_with(name) {
        return None;
    }
    let content_start = at + 1;
    let attributes_start = content_start + name.len();
    let mut cursor = attributes_start;
    let mut count = 0usize;
    loop {
        match *content.get(cursor)? {
            b'>' => {
                return Some((
                    Tag {
                        content: &content[content_start..cursor],
                        start: at,
                        end: cursor + 1,
                    },
                    false,
                ));
            },
            b'/' => {
                if content.get(cursor + 1) != Some(&b'>') {
                    return None;
                }
                return Some((
                    Tag {
                        content: &content[content_start..cursor],
                        start: at,
                        end: cursor + 2,
                    },
                    true,
                ));
            },
            b' ' => {
                let name_start = cursor + 1;
                let name_end = attribute_name(content, name_start)?;
                let attribute = &content[name_start..name_end];
                if reserved_attribute(attribute) {
                    return None;
                }
                if content.get(name_end) != Some(&b'=') || content.get(name_end + 1) != Some(&b'"')
                {
                    return None;
                }
                if count == MAX_TAG_ATTRIBUTES {
                    return None;
                }
                // The attributes before this one are already admitted, so
                // their names can be re-read instead of being stored.
                let mut earlier = Attributes {
                    rest: &content[attributes_start..cursor],
                };
                if earlier.any(|(earlier, _)| earlier == attribute) {
                    return None;
                }
                count += 1;
                let value_start = name_end + 2;
                let value_end = attribute_value(content, value_start)?;
                if content[value_start..value_end].contains(&b'\'') {
                    summary.apostrophe = true;
                }
                cursor = value_end + 1;
            },
            _ => return None,
        }
    }
}

/// Recognize `ascii-name (":" ascii-name)?` starting at `at`.
fn attribute_name(content: &[u8], at: usize) -> Option<usize> {
    let mut cursor = ascii_name(content, at)?;
    if content.get(cursor) == Some(&b':') {
        cursor = ascii_name(content, cursor + 1)?;
    }
    Some(cursor)
}

/// Recognize one ASCII name: a letter or underscore, then letters, digits,
/// underscores, hyphens and periods.
fn ascii_name(content: &[u8], at: usize) -> Option<usize> {
    let first = *content.get(at)?;
    if !(first.is_ascii_alphabetic() || first == b'_') {
        return None;
    }
    let rest = &content[at + 1..];
    let length = rest
        .iter()
        .position(|byte| !(byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')))
        .unwrap_or(rest.len());
    Some(at + 1 + length)
}

/// Namespace declarations change resolution and `xml:*` attributes change
/// whitespace handling; both stay with the ordinary reader.
fn reserved_attribute(name: &[u8]) -> bool {
    name.starts_with(b"xmlns") || name.starts_with(b"xml:")
}

/// Return the offset of the closing quote of a value starting at `at`,
/// declining any byte that decoding or normalization would change and any
/// markup delimiter.
fn attribute_value(content: &[u8], at: usize) -> Option<usize> {
    let rest = content.get(at..)?;
    let length = rest
        .iter()
        .position(|byte| matches!(byte, b'"' | b'<' | b'>' | b'&' | b'\t' | b'\r' | b'\n'))?;
    (rest[length] == b'"').then_some(at + length)
}

/// Deterministic worksheet corpus shared by the differential tests of every
/// pass that takes the lane.
#[cfg(test)]
pub(crate) mod corpus {
    use crate::cell::Text;

    pub(crate) const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    pub(crate) const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    pub(crate) const X14AC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

    pub(crate) fn worksheet(body: &str) -> String {
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
             <worksheet xmlns=\"{MAIN}\" xmlns:x14ac=\"{X14AC}\">\
             <dimension ref=\"A1:C3\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>\
             <sheetFormatPr defaultRowHeight=\"15\"/><cols><col min=\"1\" max=\"2\" width=\"9\"/></cols>\
             <sheetData>{body}</sheetData>\
             <mergeCells count=\"1\"><mergeCell ref=\"D1:E1\"/></mergeCells>\
             <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\
             </worksheet>"
        )
    }

    pub(crate) fn shared_strings() -> Vec<Text> {
        vec![Text::from("zero"), Text::from("one"), Text::from("two")]
    }

    /// Small deterministic generator so failures reproduce exactly.
    pub(crate) struct Lcg(pub(crate) u64);

    impl Lcg {
        pub(crate) fn next(&mut self) -> u64 {
            self.0 = self
                .0
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            self.0 >> 33
        }

        pub(crate) fn pick<'a>(&mut self, items: &[&'a str]) -> &'a str {
            let index = usize::try_from(self.next()).expect("u64 fits usize") % items.len();
            items[index]
        }

        pub(crate) fn chance(&mut self, percent: u64) -> bool {
            self.next() % 100 < percent
        }
    }

    pub(crate) fn column_name(mut column: u32) -> String {
        let mut name = Vec::new();
        loop {
            name.push(b'A' + u8::try_from(column % 26).expect("letter"));
            if column < 26 {
                break;
            }
            column = column / 26 - 1;
        }
        name.reverse();
        String::from_utf8(name).expect("ASCII column")
    }

    /// Generate one body. `hostile` mixes in values and structures the shared
    /// parser methods must refuse.
    pub(crate) fn generated_body(random: &mut Lcg, hostile: bool) -> String {
        let separators = ["", "", "", " ", "\n  ", "\r\n"];
        let values = [
            "1",
            "-2.5",
            "1e3",
            " 7 ",
            "",
            "0",
            "42",
            "3.14159",
            "1E-2",
            "\u{e9}\u{6f22}",
            "a>b",
            "x\r\ny",
        ];
        let hostile_values = ["abc", "1e999", "NaN", "2", "not-a-date"];
        let types = ["", "", "", " t=\"n\"", " t=\"s\"", " t=\"str\"", " t=\"b\""];
        let hostile_types = [" t=\"e\"", " t=\"d\"", " t=\"inlineStr\"", " t=\"x\""];
        let extras = [
            "",
            "",
            " s=\"3\"",
            " s=\"+2\"",
            " cm=\"1\"",
            " vm=\"2\"",
            " ph=\"1\"",
            " x14ac:foo=\"1\"",
            " xr:uid=\"{00000000-0001}\"",
            " note=\"it's\"",
        ];
        let hostile_extras = [
            " s=\"65491\"",
            " s=\"-1\"",
            " s=\"\"",
            " cm=\"0\"",
            " vm=\"2147483648\"",
        ];
        let row_extras = [
            "",
            "",
            " spans=\"1:3\"",
            " ht=\"15\" customHeight=\"1\"",
            " hidden=\"true\"",
            " s=\"2\" customFormat=\"1\"",
            " outlineLevel=\"3\"",
            " x14ac:dyDescent=\"0.25\"",
        ];
        let hostile_row_extras = [" ht=\"abc\"", " hidden=\"maybe\"", " outlineLevel=\"9\""];

        let mut body = String::new();
        let rows = 1 + random.next() % 6;
        let mut row = 0u64;
        for _ in 0..rows {
            row += if hostile && random.chance(10) {
                0
            } else {
                1 + random.next() % 3
            };
            body.push_str(random.pick(&separators));
            let mut row_attributes = String::new();
            if !random.chance(10) {
                row_attributes.push_str(&format!(" r=\"{row}\""));
            }
            row_attributes.push_str(random.pick(&row_extras));
            if hostile && random.chance(10) {
                row_attributes.push_str(random.pick(&hostile_row_extras));
            }
            if random.chance(10) {
                body.push_str(&format!("<row{row_attributes}/>"));
                continue;
            }
            body.push_str(&format!("<row{row_attributes}>"));
            let cells = random.next() % 5;
            let mut column = 0u32;
            for _ in 0..cells {
                column += if hostile && column > 0 && random.chance(10) {
                    0
                } else {
                    1 + u32::try_from(random.next() % 3).expect("small")
                };
                body.push_str(random.pick(&separators));
                let mut attributes = String::new();
                let reference_row = if hostile && random.chance(10) {
                    row + 1
                } else {
                    row
                };
                if !random.chance(15) {
                    attributes.push_str(&format!(
                        " r=\"{}{reference_row}\"",
                        column_name(column - 1)
                    ));
                }
                attributes.push_str(random.pick(&types));
                if hostile && random.chance(15) {
                    attributes.push_str(random.pick(&hostile_types));
                }
                attributes.push_str(random.pick(&extras));
                if hostile && random.chance(10) {
                    attributes.push_str(random.pick(&hostile_extras));
                }
                if attributes.matches(" t=").count() > 1 {
                    // Keep duplicate attributes for the explicit decline tests.
                    attributes = attributes.replacen(" t=", " u=", 1);
                }
                match random.next() % 6 {
                    0 => body.push_str(&format!("<c{attributes}/>")),
                    1 => body.push_str(&format!("<c{attributes}><v/></c>")),
                    2 => body.push_str(&format!("<c{attributes}></c>")),
                    _ => {
                        let value = if hostile && random.chance(20) {
                            random.pick(&hostile_values)
                        } else if attributes.contains("t=\"s\"") {
                            random.pick(&["0", "1", "2"])
                        } else if attributes.contains("t=\"b\"") {
                            random.pick(&["0", "1", "true", "false"])
                        } else {
                            random.pick(&values)
                        };
                        let inner = random.pick(&separators);
                        body.push_str(&format!("<c{attributes}>{inner}<v>{value}</v>{inner}</c>"));
                    },
                }
            }
            body.push_str(random.pick(&separators));
            body.push_str("</row>");
        }
        body.push_str(random.pick(&separators));
        body
    }
}

/// Test-only control and accounting of which route each pass took.
///
/// Differential tests run every pass twice, once with the lane disabled, and
/// assert that the lane was actually taken where they expect it so a
/// byte-for-byte comparison is not vacuous.
#[cfg(test)]
pub(crate) mod route {
    use std::cell::Cell;

    /// The worksheet passes that can take the lane.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub(crate) enum Pass {
        Parse,
        Scan,
        Compact,
        /// A commit verified a changed worksheet through the reduced readback.
        Readback,
    }

    impl Pass {
        const fn index(self) -> usize {
            match self {
                Self::Parse => 0,
                Self::Scan => 1,
                Self::Compact => 2,
                Self::Readback => 3,
            }
        }
    }

    thread_local! {
        static DISABLED: Cell<bool> = const { Cell::new(false) };
        static ADMITTED: Cell<[usize; 4]> = const { Cell::new([0; 4]) };
    }

    /// Whether passes on this thread may take the lane.
    pub(crate) fn enabled() -> bool {
        !DISABLED.with(Cell::get)
    }

    /// Run `work` with the lane disabled on this thread.
    pub(crate) fn without_lane<T>(work: impl FnOnce() -> T) -> T {
        struct Restore(bool);
        impl Drop for Restore {
            fn drop(&mut self) {
                DISABLED.with(|disabled| disabled.set(self.0));
            }
        }
        let _restore = Restore(DISABLED.with(|disabled| disabled.replace(true)));
        work()
    }

    pub(crate) fn note_admitted(pass: Pass) {
        ADMITTED.with(|admitted| {
            let mut counts = admitted.get();
            counts[pass.index()] += 1;
            admitted.set(counts);
        });
    }

    /// Lane admissions of `pass` on this thread since the last reset.
    pub(crate) fn admitted(pass: Pass) -> usize {
        ADMITTED.with(|admitted| admitted.get()[pass.index()])
    }

    pub(crate) fn reset() {
        ADMITTED.with(|admitted| admitted.set([0; 4]));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn events(body: &str) -> Option<(Summary, Vec<String>)> {
        let content = format!("<sheetData>{body}</sheetData>");
        let bytes = content.as_bytes();
        let start = "<sheetData>".len();
        let mut seen = Vec::new();
        let summary = walk(bytes, start, b"sheetData", &mut |event| {
            seen.push(match event {
                Event::Start(name, tag) => {
                    format!("start {name:?} {}", String::from_utf8_lossy(tag.content))
                },
                Event::Empty(name, tag) => {
                    format!("empty {name:?} {}", String::from_utf8_lossy(tag.content))
                },
                Event::End(name, start, end) => {
                    format!(
                        "end {name:?} {}",
                        String::from_utf8_lossy(&bytes[start..end])
                    )
                },
                Event::Text { start, end, value } => format!(
                    "text {value} {:?}",
                    String::from_utf8_lossy(&bytes[start..end])
                ),
            });
            Ok::<(), ()>(())
        })
        .expect("no-op visitor")?;
        assert_eq!(recognize(bytes, start, b"sheetData"), Some(summary));
        Some((summary, seen))
    }

    #[test]
    fn reports_the_reader_event_stream_for_admitted_bodies() {
        let (summary, seen) = events(
            "<row r=\"1\" spans=\"1:2\"><c r=\"A1\" s=\"3\"><v>12</v></c><c r=\"B1\"/></row>\
             <row r=\"2\"/>",
        )
        .expect("admitted");
        assert_eq!(summary.rows, 2);
        assert_eq!(summary.cells, 2);
        assert_eq!(summary.events, 9);
        assert!(!summary.apostrophe);
        assert_eq!(
            seen,
            [
                "start Row row r=\"1\" spans=\"1:2\"",
                "start Cell c r=\"A1\" s=\"3\"",
                "start Value v",
                "text true \"12\"",
                "end Value </v>",
                "end Cell </c>",
                "empty Cell c r=\"B1\"",
                "end Row </row>",
                "empty Row row r=\"2\"",
            ]
        );
    }

    #[test]
    fn reports_whitespace_runs_and_empty_values() {
        let (summary, seen) =
            events("\n <row r=\"1\">\n  <c r=\"A1\"> <v/> </c>\n </row>\n").expect("admitted");
        assert_eq!(summary.events, 11);
        assert_eq!(
            seen,
            [
                "text false \"\\n \"",
                "start Row row r=\"1\"",
                "text false \"\\n  \"",
                "start Cell c r=\"A1\"",
                "text false \" \"",
                "empty Value v",
                "text false \" \"",
                "end Cell </c>",
                "text false \"\\n \"",
                "end Row </row>",
                "text false \"\\n\"",
            ]
        );
    }

    #[test]
    fn keeps_value_text_verbatim_and_omits_empty_text() {
        let (_summary, seen) =
            events("<row r=\"1\"><c r=\"A1\"><v> 1\r\n</v></c><c r=\"B1\"><v></v></c></row>")
                .expect("admitted");
        assert!(seen.contains(&"text true \" 1\\r\\n\"".to_owned()));
        assert!(!seen.iter().any(|event| event == "text true \"\""));
    }

    #[test]
    fn records_apostrophes_in_attribute_values() {
        let (summary, _seen) =
            events("<row r=\"1\"><c r=\"A1\" x=\"it's\"/></row>").expect("admitted");
        assert!(summary.apostrophe);
    }

    #[test]
    fn iterates_raw_attributes() {
        let content = b"c r=\"A1\" x14ac:dyDescent=\"0.25\" t=\"s\"";
        let tag = Tag {
            content,
            start: 0,
            end: 0,
        };
        let attributes = tag.attributes().collect::<Vec<_>>();
        assert_eq!(
            attributes,
            [
                (b"r".as_slice(), b"A1".as_slice()),
                (b"x14ac:dyDescent".as_slice(), b"0.25".as_slice()),
                (b"t".as_slice(), b"s".as_slice()),
            ]
        );
        let bare = Tag {
            content: b"c",
            start: 0,
            end: 0,
        };
        assert_eq!(bare.attributes().count(), 0);
    }

    #[test]
    fn declines_everything_outside_the_subset() {
        for body in [
            // Markup the reader reports as other events.
            "<!-- note --><row r=\"1\"/>",
            "<row r=\"1\"><?pi x?></row>",
            "<row r=\"1\"><c r=\"A1\"><v><![CDATA[1]]></v></c></row>",
            // References split character data or need decoding.
            "<row r=\"1\"><c r=\"A1\"><v>1&amp;2</v></c></row>",
            "<row r=\"1\"><c r=\"A&#49;\"/></row>",
            "<row r=\"1\"><c r=\"A1\" t=\"a&amp;b\"/></row>",
            // Normalized whitespace, markup delimiters and other quotes.
            "<row r=\"1\"><c r=\"A1\" x=\"a\tb\"/></row>",
            "<row r=\"1\"><c r=\"A1\" x=\"a\nb\"/></row>",
            "<row r=\"1\"><c r=\"A1\" x=\"a>b\"/></row>",
            "<row r='1'/>",
            // Irregular spacing inside tags.
            "<row  r=\"1\"/>",
            "<row r=\"1\" />",
            "<row r = \"1\"/>",
            "<row\tr=\"1\"/>",
            "<row r=\"1\"></row >",
            // Namespace and xml attributes.
            "<row r=\"1\" xmlns=\"urn:x\"/>",
            "<row r=\"1\" xmlns:x=\"urn:x\"/>",
            "<row r=\"1\" xml:space=\"preserve\"/>",
            // Duplicate names.
            "<row r=\"1\" r=\"1\"/>",
            // Other elements and children.
            "<x:row r=\"1\"/>",
            "<rows r=\"1\"/>",
            "<extLst/>",
            "<row r=\"1\"><c r=\"A1\"><f>1+1</f><v>2</v></c></row>",
            "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>x</t></is></c></row>",
            "<row r=\"1\"><c r=\"A1\"><v>1</v><v>2</v></c></row>",
            "<row r=\"1\"><c r=\"A1\"><extLst/></c></row>",
            "<row r=\"1\"><cx r=\"A1\"/></row>",
            // Non-ASCII names and names starting with a digit.
            "<row r=\"1\"><c r=\"A1\" \u{e9}=\"1\"/></row>",
            "<row r=\"1\" 1a=\"1\"/>",
            // Stray character data between elements.
            "x<row r=\"1\"/>",
            "<row r=\"1\">x</row>",
        ] {
            assert!(events(body).is_none(), "{body} must decline");
        }
    }

    #[test]
    fn declines_truncated_documents() {
        let content = b"<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>";
        assert_eq!(recognize(content, "<sheetData>".len(), b"sheetData"), None);
        let content = b"<sheetData><row r=\"1\"><c r=\"A1\"><v>1";
        assert_eq!(recognize(content, "<sheetData>".len(), b"sheetData"), None);
        let content = b"<sheetData><row r=\"1";
        assert_eq!(recognize(content, "<sheetData>".len(), b"sheetData"), None);
        // The close tag must spell the start tag's qualified name.
        let content = b"<x:sheetData><row r=\"1\"/></sheetData>";
        assert_eq!(
            recognize(content, "<x:sheetData>".len(), b"x:sheetData"),
            None
        );
        let content = b"<x:sheetData><row r=\"1\"/></x:sheetData>";
        assert!(recognize(content, "<x:sheetData>".len(), b"x:sheetData").is_some());
        let content = b"<sheetData><row r=\"1\"/></sheetData >";
        assert_eq!(recognize(content, "<sheetData>".len(), b"sheetData"), None);
    }

    #[test]
    fn declines_tags_with_too_many_attributes() {
        let mut attributes = String::new();
        for index in 0..=MAX_TAG_ATTRIBUTES {
            attributes.push_str(&format!(" a{index}=\"1\""));
        }
        assert!(events(&format!("<row r=\"1\"><c{attributes}/></row>")).is_none());
        attributes.clear();
        for index in 0..MAX_TAG_ATTRIBUTES {
            attributes.push_str(&format!(" a{index}=\"1\""));
        }
        assert!(events(&format!("<row r=\"1\"><c{attributes}/></row>")).is_some());
    }

    #[test]
    fn entries_require_the_named_tag_at_the_reader_positions() {
        let content = b"<w><sheetData><row/></sheetData></w>";
        let (start, end) = (3, 14);
        let entry = Entry::locate(content, start, b"sheetData", end).expect("aligned tag");
        assert_eq!(entry.position, 14);
        assert_eq!(entry.name(content).expect("name"), b"sheetData");
        // A different or truncated name, or shifted positions, decline.
        for (start, name, end) in [
            (start, b"sheetDat".as_slice(), end),
            (start, b"sheetDatb".as_slice(), end),
            (start, b"x:sheetData".as_slice(), end),
            (start + 1, b"sheetData".as_slice(), end),
            (start, b"sheetData".as_slice(), end - 1),
            (start, b"sheetData".as_slice(), end + 1),
        ] {
            assert!(
                Entry::locate(content, start, name, end).is_none(),
                "{start} {end}"
            );
        }
        let spaced = b"<w><sheetData\n a=\"1\"><row/></sheetData></w>";
        assert!(Entry::locate(spaced, start, b"sheetData", 21).is_some());
    }

    #[test]
    fn byte_order_marked_parts_locate_at_document_offsets() {
        let mut content = b"\xEF\xBB\xBF".to_vec();
        content.extend_from_slice(b"<w><sheetData><row/></sheetData></w>");
        // quick-xml reports positions without the mark; the entry is at the
        // document offsets three bytes later.
        let entry = Entry::locate(&content, 3, b"sheetData", 14).expect("marked entry");
        assert_eq!(entry.position, 17);
        assert_eq!(entry.name(&content).expect("name"), b"sheetData");
        // Document offsets passed as reader positions no longer align.
        assert!(Entry::locate(&content, 6, b"sheetData", 17).is_none());
        // The reader over the spliced document stops at the same entry.
        let mut reader = NsReader::from_reader(content.as_slice());
        skip_to(&mut reader, &content, entry.position).expect("skip to the entry");
        assert_eq!(reader.buffer_position(), 14);
    }

    #[test]
    fn an_empty_body_is_admitted() {
        let (summary, seen) = events("").expect("admitted");
        assert_eq!(
            summary,
            Summary {
                end: "<sheetData>".len(),
                ..Summary::default()
            }
        );
        assert!(seen.is_empty());
    }
}
