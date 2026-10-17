//! Inline and frame details the scene index does not hold (record 0328):
//! a table's declared columns, a paragraph's line breaks, a plugin frame's
//! media, a shape's click action, a link's title and a formula object's
//! inline MathML.
//!
//! Each reader scans the bytes of one indexed node of the text the scene
//! was built from, on demand; the scene itself is unchanged, so every
//! existing read and edit is unaffected. Element and attribute names are
//! matched by their local names inside the node, as the drawing and text
//! vocabularies of ODF use them without ambiguity there.

use std::borrow::Cow;
use std::ops::Range;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::{AutomaticStyles, Object, Paragraph, Run, Table};

/// The local part of a qualified name.
fn local(name: &[u8]) -> &[u8] {
    match name.iter().position(|byte| *byte == b':') {
        Some(colon) => &name[colon + 1..],
        None => name,
    }
}

/// The `(qualified name, raw value)` pairs of a start tag's content (the
/// bytes after `<`), in source order: one linear pass, no map.
fn raw_attributes(tag: &[u8]) -> impl Iterator<Item = (&[u8], &[u8])> {
    let name_end = tag
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')
        .unwrap_or(tag.len());
    let mut rest = &tag[name_end..];
    std::iter::from_fn(move || {
        let start = rest.iter().position(|byte| !byte.is_ascii_whitespace())?;
        rest = &rest[start..];
        let equals = rest.iter().position(|byte| *byte == b'=')?;
        let name = rest[..equals].trim_ascii();
        let after = rest[equals + 1..].trim_ascii_start();
        let quote = *after.first()?;
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        let close = after[1..].iter().position(|byte| *byte == quote)?;
        let value = &after[1..=close];
        rest = &after[close + 2..];
        Some((name, value))
    })
}

/// Decode the predefined entities and numeric references of a raw value.
fn unescape(raw: &[u8]) -> String {
    let raw = String::from_utf8_lossy(raw);
    if !raw.contains('&') {
        return raw.into_owned();
    }
    let mut out = String::with_capacity(raw.len());
    let mut rest = raw.as_ref();
    while let Some(index) = rest.find('&') {
        out.push_str(&rest[..index]);
        rest = &rest[index..];
        let Some(end) = rest.find(';') else {
            break;
        };
        let entity = &rest[1..end];
        let decoded = match entity {
            "amp" => Some('&'),
            "lt" => Some('<'),
            "gt" => Some('>'),
            "quot" => Some('"'),
            "apos" => Some('\''),
            _ => entity.strip_prefix('#').and_then(|number| {
                match number
                    .strip_prefix('x')
                    .or_else(|| number.strip_prefix('X'))
                {
                    Some(hex) => u32::from_str_radix(hex, 16).ok(),
                    None => number.parse().ok(),
                }
                .and_then(char::from_u32)
            }),
        };
        match decoded {
            Some(character) => out.push(character),
            None => out.push_str(&rest[..=end]),
        }
        rest = &rest[end + 1..];
    }
    out.push_str(rest);
    out
}

/// The decoded value of the attribute whose local name is `wanted`.
fn attribute(tag: &BytesStart<'_>, wanted: &[u8]) -> Option<String> {
    raw_attributes(tag.as_ref())
        .find(|(name, _)| local(name) == wanted)
        .map(|(_, value)| unescape(value))
}

/// The decoded value of the attribute whose qualified name is `wanted`.
fn qualified_attribute(tag: &BytesStart<'_>, wanted: &[u8]) -> Option<String> {
    raw_attributes(tag.as_ref())
        .find(|(name, _)| *name == wanted)
        .map(|(_, value)| unescape(value))
}

fn reader(xml: &str) -> Reader<&[u8]> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = false;
    reader
}

/// One `table:table-column` declaration of a table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableColumn {
    range: Range<usize>,
    style_name: Option<String>,
    repeat: u32,
}

impl TableColumn {
    /// The bytes of the declaration.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// `table:style-name`.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// `table:number-columns-repeated`, 1 by default.
    #[must_use]
    pub fn repeat(&self) -> u32 {
        self.repeat
    }

    /// `style:column-width` of the column's automatic style, as written.
    #[must_use]
    pub fn width<'a>(&self, xml: &'a str, styles: &AutomaticStyles) -> Option<Cow<'a, str>> {
        let style = styles.get("table-column", self.style_name.as_deref()?)?;
        let declaration = xml.get(style.range())?;
        let mut reader = reader(declaration);
        loop {
            match reader.read_event().ok()? {
                Event::Start(tag) | Event::Empty(tag)
                    if local(tag.name().as_ref()) == b"table-column-properties" =>
                {
                    return attribute(&tag, b"column-width").map(Cow::Owned);
                },
                Event::Eof => return None,
                _ => {},
            }
        }
    }
}

/// The columns `table` declares, one entry per declaration (a repeated
/// declaration stands for [`TableColumn::repeat`] columns), nested tables'
/// columns left out.
#[must_use]
pub fn table_columns(xml: &str, table: &Table) -> Vec<TableColumn> {
    let Some(source) = xml.get(table.range()) else {
        return Vec::new();
    };
    let base = table.range().start;
    let mut reader = reader(source);
    let mut out = Vec::new();
    // Depth of the table's own element is 1; its column declarations sit
    // at depth 2, or 3 inside a column group.
    let mut depth = 0usize;
    let mut groups = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let Ok(event) = reader.read_event() else {
            break;
        };
        let end = usize::try_from(reader.buffer_position()).unwrap_or(source.len());
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let empty = matches!(source.as_bytes().get(end.saturating_sub(2)), Some(b'/'));
                let name = tag.name();
                let name = local(name.as_ref());
                let level = depth + 1;
                let in_table = level == 2 || (level == 3 && groups > 0);
                if in_table && name == b"table-column" {
                    out.push(TableColumn {
                        range: base + start..base + end,
                        style_name: attribute(&tag, b"style-name"),
                        repeat: attribute(&tag, b"number-columns-repeated")
                            .and_then(|value| value.parse::<u32>().ok())
                            .unwrap_or(1)
                            .clamp(1, 16_384),
                    });
                }
                if !empty {
                    depth += 1;
                    if level == 2
                        && matches!(
                            name,
                            b"table-columns" | b"table-header-columns" | b"table-column-group"
                        )
                    {
                        groups += 1;
                    }
                }
            },
            Event::End(tag) => {
                if depth == 2
                    && matches!(
                        local(tag.name().as_ref()),
                        b"table-columns" | b"table-header-columns" | b"table-column-group"
                    )
                {
                    groups = groups.saturating_sub(1);
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out
}

/// Inline subtrees whose line breaks are not the paragraph's: frames,
/// drawing shapes, notes and annotations.
const SKIPPED: [&[u8]; 12] = [
    b"frame",
    b"note",
    b"annotation",
    b"g",
    b"custom-shape",
    b"rect",
    b"ellipse",
    b"circle",
    b"line",
    b"polygon",
    b"path",
    b"connector",
];

/// The bytes of every `text:line-break` of `paragraph`, in document order;
/// breaks inside frames, notes and annotations the paragraph holds are left
/// out.
#[must_use]
pub fn line_breaks(xml: &str, paragraph: &Paragraph) -> Vec<Range<usize>> {
    let content = paragraph.content.clone();
    let Some(source) = xml.get(content.clone()) else {
        return Vec::new();
    };
    let mut reader = reader(source);
    let mut out = Vec::new();
    let mut skip = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let Ok(event) = reader.read_event() else {
            break;
        };
        let end = usize::try_from(reader.buffer_position()).unwrap_or(source.len());
        match event {
            Event::Start(tag) => {
                let name = tag.name();
                let name = local(name.as_ref());
                if skip > 0 {
                    skip += 1;
                } else if SKIPPED.contains(&name) {
                    skip = 1;
                } else if name == b"line-break" {
                    out.push(content.start + start..content.start + end);
                }
            },
            Event::Empty(tag) => {
                if skip == 0 && local(tag.name().as_ref()) == b"line-break" {
                    out.push(content.start + start..content.start + end);
                }
            },
            Event::End(_) => skip = skip.saturating_sub(1),
            Event::Eof => break,
            _ => {},
        }
    }
    out
}

/// Where a new line break goes in `paragraph`: the end of its content (an
/// empty paragraph first grows an end tag, which [`super::Change`]
/// planning handles).
pub(crate) fn line_break_insertion(paragraph: &Paragraph) -> Option<usize> {
    (!paragraph.empty).then_some(paragraph.content.end)
}

/// The media a `draw:plugin` frame plays.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Plugin {
    mime_type: Option<String>,
    href: Option<String>,
    parameters: Vec<(String, String)>,
}

impl Plugin {
    /// `draw:mime-type`.
    #[must_use]
    pub fn mime_type(&self) -> Option<&str> {
        self.mime_type.as_deref()
    }

    /// `xlink:href`: the media member or link.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// The `draw:param` name and value pairs, in document order.
    #[must_use]
    pub fn parameters(&self) -> &[(String, String)] {
        &self.parameters
    }

    /// The value of parameter `name` (ASCII case-insensitive).
    #[must_use]
    pub fn parameter(&self, name: &str) -> Option<&str> {
        self.parameters
            .iter()
            .find(|(held, _)| held.eq_ignore_ascii_case(name))
            .map(|(_, value)| value.as_str())
    }

    /// Whether the media is video, audio or neither, from the MIME type,
    /// else the member's file extension.
    #[must_use]
    pub fn media_kind(&self) -> Option<MediaKind> {
        if let Some(mime) = self.mime_type.as_deref() {
            let mime = mime.to_ascii_lowercase();
            if mime.starts_with("video/") {
                return Some(MediaKind::Video);
            }
            if mime.starts_with("audio/") {
                return Some(MediaKind::Audio);
            }
            if mime == "application/vnd.sun.star.media" {
                // LibreOffice's generic media type: decide by extension.
            } else {
                return None;
            }
        }
        let href = self.href.as_deref()?;
        let extension = href.rsplit('.').next()?.to_ascii_lowercase();
        match extension.as_str() {
            "mp4" | "m4v" | "mov" | "avi" | "wmv" | "mpg" | "mpeg" | "ogv" | "webm" | "mkv"
            | "3gp" | "flv" => Some(MediaKind::Video),
            "mp3" | "wav" | "wma" | "ogg" | "oga" | "m4a" | "aac" | "flac" | "aif" | "aiff"
            | "mid" | "midi" => Some(MediaKind::Audio),
            _ => None,
        }
    }
}

/// What a plugin frame plays.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum MediaKind {
    /// Video.
    Video,
    /// Audio.
    Audio,
}

/// The `draw:plugin` of a plugin frame.
#[must_use]
pub fn plugin(xml: &str, object: &Object) -> Option<Plugin> {
    let source = xml.get(object.range())?;
    let mut reader = reader(source);
    let mut found: Option<Plugin> = None;
    let mut depth = 0usize;
    let mut plugin_depth = None;
    loop {
        match reader.read_event().ok()? {
            Event::Start(tag) => {
                depth += 1;
                visit_plugin(&tag, depth, &mut found, &mut plugin_depth, false);
            },
            Event::Empty(tag) => {
                visit_plugin(&tag, depth + 1, &mut found, &mut plugin_depth, true);
            },
            Event::End(_) => {
                if plugin_depth == Some(depth) {
                    return found;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => return found,
            _ => {},
        }
    }
}

fn visit_plugin(
    tag: &BytesStart<'_>,
    depth: usize,
    found: &mut Option<Plugin>,
    plugin_depth: &mut Option<usize>,
    empty: bool,
) {
    let name = tag.name();
    match local(name.as_ref()) {
        b"plugin" if found.is_none() && depth == 2 => {
            *found = Some(Plugin {
                mime_type: attribute(tag, b"mime-type"),
                href: attribute(tag, b"href"),
                parameters: Vec::new(),
            });
            if !empty {
                *plugin_depth = Some(depth);
            }
        },
        b"param" if plugin_depth.is_some_and(|open| depth == open + 1) => {
            if let (Some(plugin), Some(name)) = (found.as_mut(), attribute(tag, b"name")) {
                plugin
                    .parameters
                    .push((name, attribute(tag, b"value").unwrap_or_default()));
            }
        },
        _ => {},
    }
}

/// What clicking a shape does: an `office:event-listeners` click listener
/// of the shape, else the `draw:a` the shape sits in.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Click {
    action: Option<String>,
    href: Option<String>,
    /// The listener element's bytes (`None` for a `draw:a` link).
    listener: Option<Range<usize>>,
    /// The `office:event-listeners` container and how many listeners it
    /// holds.
    container: Option<(Range<usize>, usize)>,
}

impl Click {
    /// `presentation:action` of a listener (`show`, `next-page`,
    /// `first-page`, ...); `None` for a `draw:a` link.
    #[must_use]
    pub fn action(&self) -> Option<&str> {
        self.action.as_deref()
    }

    /// `xlink:href`: a URL, or `#PageName` for a slide.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// Whether the click comes from a `draw:a` around the shape.
    #[must_use]
    pub fn is_wrapper(&self) -> bool {
        self.listener.is_none()
    }

    /// The listener element's bytes.
    pub(crate) fn listener(&self) -> Option<Range<usize>> {
        self.listener.clone()
    }

    /// The listener container and its listener count.
    pub(crate) fn container(&self) -> Option<(Range<usize>, usize)> {
        self.container.clone()
    }
}

/// The click action of `object`: its own click listener, else a `draw:a`
/// that holds only it.
#[must_use]
pub fn shape_click(xml: &str, object: &Object) -> Option<Click> {
    if let Some(click) = listener_click(xml, object) {
        return Some(click);
    }
    wrapper_click(xml, object)
}

fn listener_click(xml: &str, object: &Object) -> Option<Click> {
    let range = object.range();
    let source = xml.get(range.clone())?;
    let mut reader = reader(source);
    let mut depth = 0usize;
    let mut container: Option<(usize, usize)> = None;
    let mut listeners = 0usize;
    let mut click: Option<Click> = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let event = reader.read_event().ok()?;
        let end = usize::try_from(reader.buffer_position()).unwrap_or(source.len());
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let empty = matches!(source.as_bytes().get(end.saturating_sub(2)), Some(b'/'));
                let level = depth + 1;
                let name = tag.name();
                let name = local(name.as_ref());
                if level == 2 && name == b"event-listeners" {
                    container = Some((start, if empty { end } else { usize::MAX }));
                }
                if level == 3 && container.is_some() && name == b"event-listener" {
                    listeners += 1;
                    let event_name = attribute(&tag, b"event-name");
                    if click.is_none()
                        && event_name
                            .as_deref()
                            .is_some_and(|name| name.ends_with("click"))
                    {
                        click = Some(Click {
                            action: attribute(&tag, b"action"),
                            href: attribute(&tag, b"href"),
                            listener: Some(range.start + start..range.start + end),
                            container: None,
                        });
                        if !empty {
                            // The listener's end is fixed when it closes.
                            if let Some(click) = click.as_mut() {
                                click.listener = Some(range.start + start..usize::MAX);
                            }
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(tag) => {
                let name = tag.name();
                let name = local(name.as_ref());
                if depth == 3
                    && name == b"event-listener"
                    && let Some(click) = click.as_mut()
                    && let Some(listener) = click.listener.as_mut()
                    && listener.end == usize::MAX
                {
                    listener.end = range.start + end;
                }
                if depth == 2
                    && name == b"event-listeners"
                    && let Some(open) = container.as_mut()
                {
                    open.1 = end;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let mut click = click?;
    if let Some((start, end)) = container {
        click.container = Some((range.start + start..range.start + end, listeners));
    }
    Some(click)
}

/// The start tag and the whole element of a `draw:a` whose only content is
/// `object`.
pub(crate) fn wrapper_ranges(xml: &str, object: &Object) -> Option<(Range<usize>, Range<usize>)> {
    let start = object.range().start;
    let before = xml.get(..start)?.trim_end();
    if !before.ends_with('>') {
        return None;
    }
    let open = before.rfind('<')?;
    let tag = &before[open..];
    if tag.starts_with("</") || tag.ends_with("/>") {
        return None;
    }
    let qualified = tag[1..]
        .split(|character: char| character.is_ascii_whitespace() || character == '>')
        .next()?;
    if local(qualified.as_bytes()) != b"a" {
        return None;
    }
    // The link must close right after the object.
    let end = object.range().end;
    let after = xml.get(end..)?;
    let trimmed = after.trim_start();
    let close = format!("</{qualified}>");
    if !trimmed.starts_with(&close) {
        return None;
    }
    let whole_end = end + (after.len() - trimmed.len()) + close.len();
    Some((open..before.len(), open..whole_end))
}

/// A `draw:a` whose whole content is `object`.
fn wrapper_click(xml: &str, object: &Object) -> Option<Click> {
    let (tag, _) = wrapper_ranges(xml, object)?;
    let mut reader = reader(xml.get(tag)?);
    let Ok(Event::Start(start)) = reader.read_event() else {
        return None;
    };
    Some(Click {
        action: None,
        href: attribute(&start, b"href"),
        listener: None,
        container: None,
    })
}

/// The direct child elements of `object`: their local names and bytes.
pub(crate) fn direct_children(xml: &str, object: &Object) -> Vec<(String, Range<usize>)> {
    let Some(source) = xml.get(object.range()) else {
        return Vec::new();
    };
    let base = object.range().start;
    let mut reader = reader(source);
    let mut out: Vec<(String, Range<usize>)> = Vec::new();
    let mut depth = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let Ok(event) = reader.read_event() else {
            break;
        };
        let end = usize::try_from(reader.buffer_position()).unwrap_or(source.len());
        match event {
            Event::Start(tag) => {
                depth += 1;
                if depth == 2 {
                    let name = String::from_utf8_lossy(local(tag.name().as_ref())).into_owned();
                    out.push((name, base + start..usize::MAX));
                }
            },
            Event::Empty(tag) => {
                if depth == 1 {
                    let name = String::from_utf8_lossy(local(tag.name().as_ref())).into_owned();
                    out.push((name, base + start..base + end));
                }
            },
            Event::End(_) => {
                if depth == 2
                    && let Some((_, range)) = out.last_mut()
                    && range.end == usize::MAX
                {
                    range.end = base + end;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out
}

/// `office:title` of a link run (`text:a`).
#[must_use]
pub fn link_title(xml: &str, run: &Run) -> Option<String> {
    let tag = xml.get(run.start_tag.clone())?;
    let mut reader = reader(tag);
    match reader.read_event().ok()? {
        Event::Start(tag) | Event::Empty(tag) => {
            qualified_attribute(&tag, b"office:title").or_else(|| attribute(&tag, b"title"))
        },
        _ => None,
    }
}

/// The inline MathML (`math:math`) payload of a formula object frame.
#[must_use]
pub fn inline_math<'a>(xml: &'a str, object: &Object) -> Option<&'a str> {
    let source = xml.get(object.range())?;
    let mut reader = reader(source);
    let mut depth = 0usize;
    let mut math: Option<(usize, usize)> = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(0);
        let event = reader.read_event().ok()?;
        let end = usize::try_from(reader.buffer_position()).unwrap_or(source.len());
        match event {
            Event::Start(tag) => {
                depth += 1;
                if math.is_none() && depth == 3 && local(tag.name().as_ref()) == b"math" {
                    math = Some((start, depth));
                }
            },
            Event::End(_) => {
                if let Some((open, level)) = math
                    && level == depth
                {
                    return source.get(open..end);
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => return None,
            _ => {},
        }
    }
}

/// The formula text of a MathML payload: its StarMath annotation when it
/// carries one, else the text of its token elements; `None` when the
/// payload holds no `math` element.
#[must_use]
pub fn formula_text(mathml: &str) -> Option<String> {
    let mut reader = reader(mathml);
    let mut annotation: Option<String> = None;
    let mut in_annotation = false;
    let mut skip = 0usize;
    let mut tokens = String::new();
    let mut saw_math = false;
    loop {
        match reader.read_event().ok()? {
            Event::Start(tag) => {
                let name = tag.name();
                let name = local(name.as_ref());
                if name == b"math" {
                    saw_math = true;
                }
                if skip > 0 {
                    skip += 1;
                } else if name == b"annotation" {
                    let star = attribute(&tag, b"encoding")
                        .is_some_and(|encoding| encoding.contains("StarMath"));
                    if star && annotation.is_none() {
                        in_annotation = true;
                        annotation = Some(String::new());
                    } else {
                        skip = 1;
                    }
                } else if name == b"annotation-xml" {
                    skip = 1;
                }
            },
            Event::End(_) => {
                if skip > 0 {
                    skip -= 1;
                } else {
                    in_annotation = false;
                }
            },
            Event::Text(text) => {
                let raw = std::str::from_utf8(text.as_ref()).ok()?;
                if in_annotation {
                    if let Some(out) = annotation.as_mut() {
                        out.push_str(raw);
                    }
                } else if skip == 0 {
                    tokens.push_str(raw.trim());
                }
            },
            Event::GeneralRef(reference) => {
                let name = std::str::from_utf8(reference.as_ref()).ok()?;
                let decoded = match name {
                    "amp" => "&".to_owned(),
                    "lt" => "<".to_owned(),
                    "gt" => ">".to_owned(),
                    "quot" => "\"".to_owned(),
                    "apos" => "'".to_owned(),
                    other => other
                        .strip_prefix('#')
                        .and_then(|number| match number.strip_prefix('x') {
                            Some(hex) => u32::from_str_radix(hex, 16).ok(),
                            None => number.parse().ok(),
                        })
                        .and_then(char::from_u32)
                        .map(String::from)
                        .unwrap_or_default(),
                };
                if in_annotation {
                    if let Some(out) = annotation.as_mut() {
                        out.push_str(&decoded);
                    }
                } else if skip == 0 {
                    tokens.push_str(&decoded);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !saw_math {
        return None;
    }
    match annotation {
        Some(text) if !text.trim().is_empty() => Some(text.trim().to_owned()),
        _ => Some(tokens),
    }
}
