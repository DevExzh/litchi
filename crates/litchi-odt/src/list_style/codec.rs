//! XML codecs for ODF `text:list-style` declarations.

use super::{
    MAX_BINARY, MAX_DEPTH, MAX_LEVEL, MAX_STYLES, MAX_TOTAL, MAX_VALUE, MAX_XML, bad,
    model::{
        BulletRelativeSize, BulletStyle, ImageSource, Kind, LevelStyle, NumberStyle, Style, Styles,
    },
    name_ok, parse_bool,
};
use crate::outline_style::{NumberFormat, PositiveInteger};
use litchi_core::{Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const STYLE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK_STR: &str = "http://www.w3.org/1999/xlink";

impl LevelStyle {
    fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = match &self.kind {
            Kind::Number(_) => "list-level-style-number",
            Kind::Bullet(_) => "list-level-style-bullet",
            Kind::Image(_) => "list-level-style-image",
        };
        let mut xml = format!(r#"<text:{tag} text:level="{}""#, self.level);
        if let Some(value) = &self.style_name {
            xml.push_str(&format!(r#" text:style-name="{}""#, escape_xml(value)));
        }
        match &self.kind {
            Kind::Number(number) => {
                if let Some(value) = &number.format {
                    xml.push_str(&format!(
                        r#" style:num-format="{}""#,
                        escape_xml(value.as_str())
                    ));
                }
                if let Some(value) = &number.prefix {
                    xml.push_str(&format!(r#" style:num-prefix="{}""#, escape_xml(value)));
                }
                if let Some(value) = &number.suffix {
                    xml.push_str(&format!(r#" style:num-suffix="{}""#, escape_xml(value)));
                }
                if let Some(value) = number.letter_sync {
                    xml.push_str(&format!(r#" style:num-letter-sync="{value}""#));
                }
                if let Some(value) = &number.display_levels {
                    xml.push_str(&format!(
                        r#" text:display-levels="{}""#,
                        escape_xml(value.as_str())
                    ));
                }
                if let Some(value) = &number.start_value {
                    xml.push_str(&format!(
                        r#" text:start-value="{}""#,
                        escape_xml(value.as_str())
                    ));
                }
                xml.push_str("/>");
            },
            Kind::Bullet(bullet) => {
                xml.push_str(&format!(
                    r#" text:bullet-char="{}""#,
                    escape_xml(&bullet.bullet_char.to_string())
                ));
                if let Some(value) = &bullet.relative_size {
                    xml.push_str(&format!(
                        r#" text:bullet-relative-size="{}""#,
                        escape_xml(value.as_str())
                    ));
                }
                if let Some(value) = &bullet.prefix {
                    xml.push_str(&format!(r#" style:num-prefix="{}""#, escape_xml(value)));
                }
                if let Some(value) = &bullet.suffix {
                    xml.push_str(&format!(r#" style:num-suffix="{}""#, escape_xml(value)));
                }
                xml.push_str("/>");
            },
            Kind::Image(ImageSource::Linked(href)) => {
                xml.push_str(&format!(
                    r#" xlink:type="simple" xlink:href="{}" xlink:show="embed" xlink:actuate="onLoad""#,
                    escape_xml(href)
                ));
                xml.push_str("/>");
            },
            Kind::Image(ImageSource::Embedded(data)) => {
                xml.push('>');
                xml.push_str(&format!(
                    "<office:binary-data>{data}</office:binary-data></text:{tag}>"
                ));
            },
        }
        Ok(xml)
    }
}

impl Style {
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<text:list-style xmlns:text="{TEXT_STR}" xmlns:style="{STYLE_STR}" xmlns:office="{OFFICE_STR}" xmlns:xlink="{XLINK_STR}" style:name="{}""#,
            escape_xml(&self.name)
        );
        if let Some(value) = &self.display_name {
            xml.push_str(&format!(r#" style:display-name="{}""#, escape_xml(value)));
        }
        if let Some(value) = self.consecutive_numbering {
            xml.push_str(&format!(r#" text:consecutive-numbering="{value}""#));
        }
        if self.levels.is_empty() {
            xml.push_str("/>");
            return Ok(xml);
        }
        xml.push('>');
        for level in &self.levels {
            xml.push_str(&level.to_xml_fragment()?);
        }
        xml.push_str("</text:list-style>");
        Ok(xml)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Ns {
    Office,
    Style,
    Text,
    Xlink,
    Other,
}
fn known(resolve: ResolveResult<'_>) -> Ns {
    match resolve {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == TEXT => Ns::Text,
        ResolveResult::Bound(value) if value.as_ref() == XLINK => Ns::Xlink,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (known(namespace), local.as_ref().to_vec())
}
fn value(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    attr: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attr.decoded_and_normalized_value(version, reader.decoder())
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| bad(format!("invalid attribute value: {error}")))
}

/// Namespace-resolved attribute bag with one-shot accessors.
struct Attrs(Vec<(Ns, Vec<u8>, String)>);
fn attrs(reader: &NsReader<&[u8]>, version: XmlVersion, start: &BytesStart<'_>) -> Result<Attrs> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|error| bad(format!("invalid attribute: {error}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        let value = value(reader, version, &attr)?;
        if value.len() > MAX_VALUE {
            return Err(bad("attribute value is too large"));
        }
        if !seen.insert(local.as_ref().to_vec()) {
            return Err(bad("duplicate attribute"));
        }
        out.push((known(namespace), local.as_ref().to_vec(), value));
    }
    Ok(Attrs(out))
}
impl Attrs {
    fn take(&mut self, namespace: Ns, local: &[u8]) -> Option<String> {
        let at = self
            .0
            .iter()
            .position(|(n, l, _)| *n == namespace && l.as_slice() == local)?;
        Some(self.0.swap_remove(at).2)
    }
    fn reject_unknown(&self, element: &str) -> Result<()> {
        for (namespace, local, _) in &self.0 {
            if *namespace != Ns::Other {
                return Err(bad(format!(
                    "unknown attribute '{}' on {element}",
                    String::from_utf8_lossy(local)
                )));
            }
        }
        Ok(())
    }
}

fn parse_level_common(attrs: &mut Attrs) -> Result<(u16, Option<String>)> {
    let level = attrs
        .take(Ns::Text, b"level")
        .ok_or_else(|| bad("list level missing text:level"))?
        .parse::<u16>()
        .map_err(|_| bad("invalid text:level"))?;
    if !(1..=MAX_LEVEL).contains(&level) {
        return Err(bad("text:level is outside the supported range"));
    }
    let style_name = attrs.take(Ns::Text, b"style-name");
    if let Some(value) = &style_name {
        name_ok(value, "text:style-name")?;
    }
    Ok((level, style_name))
}

fn parse_number_level(mut attrs: Attrs) -> Result<LevelStyle> {
    let (level, style_name) = parse_level_common(&mut attrs)?;
    let number = NumberStyle {
        format: attrs
            .take(Ns::Style, b"num-format")
            .map(NumberFormat::new)
            .transpose()?,
        prefix: attrs.take(Ns::Style, b"num-prefix"),
        suffix: attrs.take(Ns::Style, b"num-suffix"),
        letter_sync: attrs
            .take(Ns::Style, b"num-letter-sync")
            .map(|value| parse_bool(&value, "style:num-letter-sync"))
            .transpose()?,
        display_levels: attrs
            .take(Ns::Text, b"display-levels")
            .map(PositiveInteger::new)
            .transpose()?,
        start_value: attrs
            .take(Ns::Text, b"start-value")
            .map(PositiveInteger::new)
            .transpose()?,
    };
    attrs.reject_unknown("text:list-level-style-number")?;
    let result = LevelStyle {
        level,
        style_name,
        kind: Kind::Number(number),
    };
    result.validate()?;
    Ok(result)
}

fn parse_bullet_level(mut attrs: Attrs) -> Result<LevelStyle> {
    let (level, style_name) = parse_level_common(&mut attrs)?;
    let bullet_char = attrs
        .take(Ns::Text, b"bullet-char")
        .ok_or_else(|| bad("list level bullet missing text:bullet-char"))?;
    let mut chars = bullet_char.chars();
    let bullet_char = chars
        .next()
        .filter(|_| chars.next().is_none())
        .ok_or_else(|| bad("text:bullet-char must be a single character"))?;
    let bullet = BulletStyle {
        bullet_char,
        relative_size: attrs
            .take(Ns::Text, b"bullet-relative-size")
            .map(BulletRelativeSize::new)
            .transpose()?,
        prefix: attrs.take(Ns::Style, b"num-prefix"),
        suffix: attrs.take(Ns::Style, b"num-suffix"),
    };
    attrs.reject_unknown("text:list-level-style-bullet")?;
    let result = LevelStyle {
        level,
        style_name,
        kind: Kind::Bullet(bullet),
    };
    result.validate()?;
    Ok(result)
}

/// Parse the attributes of an image level. Returns the common level data and the
/// optional `xlink:href`; an absent href means `office:binary-data` is required.
fn parse_image_attrs(mut attrs: Attrs) -> Result<(u16, Option<String>, Option<String>)> {
    let (level, style_name) = parse_level_common(&mut attrs)?;
    let href = attrs.take(Ns::Xlink, b"href");
    for (local, expected) in [
        (b"type".as_slice(), "simple"),
        (b"show".as_slice(), "embed"),
        (b"actuate".as_slice(), "onLoad"),
    ] {
        if let Some(value) = attrs.take(Ns::Xlink, local)
            && value != expected
        {
            return Err(bad(format!(
                "xlink:{} must be '{expected}' on a list level image",
                String::from_utf8_lossy(local)
            )));
        }
    }
    attrs.reject_unknown("text:list-level-style-image")?;
    if let Some(value) = &href {
        name_ok(value, "xlink:href")?;
    }
    Ok((level, style_name, href))
}

/// State of an open `text:list-level-style-image` element.
struct ImageState {
    depth: usize,
    level: u16,
    style_name: Option<String>,
    href: Option<String>,
    binary_seen: bool,
    binary: String,
}

struct Active {
    depth: usize,
    style: Style,
    levels: HashSet<u16>,
    /// Depth of an open non-image level element (number or bullet).
    open_level: Option<usize>,
    image: Option<ImageState>,
    /// Depth of a skipped `style:list-level-properties`/`style:text-properties` subtree.
    skip: Option<usize>,
}
impl Active {
    fn push_level(&mut self, level: LevelStyle) -> Result<()> {
        if self.style.levels.len() >= usize::from(MAX_LEVEL) {
            return Err(bad("too many list levels"));
        }
        if !self.levels.insert(level.level) {
            return Err(bad("duplicate list level"));
        }
        self.style.levels.push(level);
        Ok(())
    }
}

fn push_style(styles: &mut Vec<Style>, style: Style, total: &mut usize) -> Result<()> {
    if styles.len() >= MAX_STYLES {
        return Err(bad("too many list styles"));
    }
    if styles.iter().any(|old| old.name == style.name) {
        return Err(bad("duplicate list style name"));
    }
    *total += style.name.len()
        + style
            .levels
            .iter()
            .map(|level| level.style_name.as_deref().map_or(0, str::len))
            .sum::<usize>();
    if *total > MAX_TOTAL {
        return Err(bad("list style data is too large"));
    }
    styles.push(style);
    Ok(())
}

fn style_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Style> {
    let mut a = attrs(reader, version, start)?;
    let name = a
        .take(Ns::Style, b"name")
        .ok_or_else(|| bad("list style missing style:name"))?;
    let style = Style {
        name,
        display_name: a.take(Ns::Style, b"display-name"),
        consecutive_numbering: a
            .take(Ns::Text, b"consecutive-numbering")
            .map(|value| parse_bool(&value, "text:consecutive-numbering"))
            .transpose()?,
        levels: Vec::new(),
    };
    a.reject_unknown("text:list-style")?;
    style.validate()?;
    Ok(style)
}

fn is_style(current: &(Ns, Vec<u8>), parent: Option<&(Ns, Vec<u8>)>) -> bool {
    parent.is_some_and(|(n, l)| {
        *n == Ns::Office && matches!(l.as_slice(), b"styles" | b"automatic-styles")
    }) && current.0 == Ns::Text
        && current.1 == b"list-style"
}

/// Parse every `text:list-style` declared in a styles or flat document part.
pub fn parse(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    if !xml.contains("list-style") {
        return Ok(Styles::default());
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut styles = Vec::new();
    let mut total = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML nesting is too deep"));
                }
                let current = element(&reader, start.name());
                let direct = is_style(&current, stack.last());
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if active.is_some() {
                        return Err(bad("nested text:list-style"));
                    }
                    active = Some(Active {
                        depth,
                        style: style_attrs(&reader, version, &start)?,
                        levels: HashSet::new(),
                        open_level: None,
                        image: None,
                        skip: None,
                    });
                    continue;
                }
                let Some(state) = active.as_mut() else {
                    continue;
                };
                if state.skip.is_some() {
                    // Inside a skipped properties subtree.
                    continue;
                }
                if let Some(image) = state.image.as_mut() {
                    if depth == image.depth + 1
                        && image.href.is_none()
                        && !image.binary_seen
                        && current.0 == Ns::Office
                        && current.1 == b"binary-data"
                    {
                        image.binary_seen = true;
                        continue;
                    }
                    return Err(bad(
                        "text:list-level-style-image allows only one office:binary-data child",
                    ));
                }
                if let Some(open) = state.open_level {
                    if depth == open + 1
                        && current.0 == Ns::Style
                        && matches!(
                            current.1.as_slice(),
                            b"list-level-properties" | b"text-properties"
                        )
                    {
                        // Properties of the open level: skip their subtree.
                        state.skip = Some(depth);
                        continue;
                    }
                    return Err(bad("unsupported list level child"));
                }
                match (depth == state.depth + 1, current.0, current.1.as_slice()) {
                    (true, Ns::Text, b"list-level-style-number") => {
                        let level = parse_number_level(attrs(&reader, version, &start)?)?;
                        state.push_level(level)?;
                        state.open_level = Some(depth);
                    },
                    (true, Ns::Text, b"list-level-style-bullet") => {
                        let level = parse_bullet_level(attrs(&reader, version, &start)?)?;
                        state.push_level(level)?;
                        state.open_level = Some(depth);
                    },
                    (true, Ns::Text, b"list-level-style-image") => {
                        let (level, style_name, href) =
                            parse_image_attrs(attrs(&reader, version, &start)?)?;
                        if state.style.levels.len() >= usize::from(MAX_LEVEL) {
                            return Err(bad("too many list levels"));
                        }
                        if state.levels.contains(&level) {
                            return Err(bad("duplicate list level"));
                        }
                        state.image = Some(ImageState {
                            depth,
                            level,
                            style_name,
                            href,
                            binary_seen: false,
                            binary: String::new(),
                        });
                    },
                    _ => {
                        return Err(bad("unsupported text:list-style child"));
                    },
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                if is_style(&current, stack.last()) {
                    if active.is_some() {
                        return Err(bad("nested text:list-style"));
                    }
                    push_style(
                        &mut styles,
                        style_attrs(&reader, version, &start)?,
                        &mut total,
                    )?;
                    continue;
                }
                let Some(state) = active.as_mut() else {
                    continue;
                };
                let depth = stack.len() + 1;
                if state.skip.is_some() {
                    continue;
                }
                if let Some(image) = state.image.as_mut() {
                    if depth == image.depth + 1
                        && image.href.is_none()
                        && !image.binary_seen
                        && current.0 == Ns::Office
                        && current.1 == b"binary-data"
                    {
                        image.binary_seen = true;
                        continue;
                    }
                    return Err(bad(
                        "text:list-level-style-image allows only one office:binary-data child",
                    ));
                }
                if let Some(open) = state.open_level {
                    if depth == open + 1
                        && current.0 == Ns::Style
                        && matches!(
                            current.1.as_slice(),
                            b"list-level-properties" | b"text-properties"
                        )
                    {
                        continue;
                    }
                    return Err(bad("unsupported list level child"));
                }
                match (depth == state.depth + 1, current.0, current.1.as_slice()) {
                    (true, Ns::Text, b"list-level-style-number") => {
                        let level = parse_number_level(attrs(&reader, version, &start)?)?;
                        state.push_level(level)?;
                    },
                    (true, Ns::Text, b"list-level-style-bullet") => {
                        let level = parse_bullet_level(attrs(&reader, version, &start)?)?;
                        state.push_level(level)?;
                    },
                    (true, Ns::Text, b"list-level-style-image") => {
                        let (level, style_name, href) =
                            parse_image_attrs(attrs(&reader, version, &start)?)?;
                        let href = href.ok_or_else(|| {
                            bad("text:list-level-style-image requires xlink:href or office:binary-data")
                        })?;
                        state.push_level(LevelStyle {
                            level,
                            style_name,
                            kind: Kind::Image(ImageSource::Linked(href)),
                        })?;
                    },
                    _ => {
                        return Err(bad("unsupported text:list-style child"));
                    },
                }
            },
            Ok(Event::Text(text)) => {
                if let Some(state) = active.as_mut()
                    && let Some(image) = state.image.as_mut()
                    && image.binary_seen
                {
                    let bytes: &[u8] = text.as_ref();
                    if !bytes.is_empty() {
                        let text = text
                            .decode()
                            .map_err(|error| bad(format!("invalid binary data: {error}")))?;
                        if image.binary.len() + text.len() > MAX_BINARY {
                            return Err(bad("office:binary-data is too large"));
                        }
                        image.binary.push_str(&text);
                    }
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(state) = active.as_mut() {
                    if state.skip == Some(depth) {
                        state.skip = None;
                    } else if state.open_level == Some(depth) {
                        state.open_level = None;
                    } else if state
                        .image
                        .as_ref()
                        .is_some_and(|image| image.depth == depth)
                    {
                        let image = state.image.take().unwrap();
                        let source = match (image.href, image.binary_seen) {
                            (Some(href), false) => ImageSource::Linked(href),
                            (None, true) => ImageSource::Embedded(image.binary),
                            _ => {
                                return Err(bad(
                                    "text:list-level-style-image requires exactly one of xlink:href or office:binary-data",
                                ));
                            },
                        };
                        let level = LevelStyle {
                            level: image.level,
                            style_name: image.style_name,
                            kind: Kind::Image(source),
                        };
                        level.validate()?;
                        state.push_level(level)?;
                    } else if state.depth == depth {
                        let state = active.take().unwrap();
                        push_style(&mut styles, state.style, &mut total)?;
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(Styles { styles })
}
