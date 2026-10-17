//! Byte-level XML helpers for the body index: namespace prefixes, start-tag
//! attribute scanning, entity decoding and escaping.
//!
//! The scanner is linear in the tag length and never builds a map, so a
//! hostile tag with many attributes costs one pass (record 0770).

use std::borrow::Cow;

use litchi_core::{Error, Result};

/// The OpenDocument namespaces the body index distinguishes.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub(crate) enum Ns {
    Office,
    Text,
    Table,
    Draw,
    Style,
    Fo,
    Xlink,
    Svg,
    Dc,
    Meta,
    Xml,
    Loext,
}

impl Ns {
    pub(crate) const ALL: [Self; 12] = [
        Self::Office,
        Self::Text,
        Self::Table,
        Self::Draw,
        Self::Style,
        Self::Fo,
        Self::Xlink,
        Self::Svg,
        Self::Dc,
        Self::Meta,
        Self::Xml,
        Self::Loext,
    ];

    pub(crate) const fn uri(self) -> &'static str {
        use litchi_odf_common::namespace as ns;
        match self {
            Self::Office => ns::OFFICENS,
            Self::Text => ns::TEXTNS,
            Self::Table => ns::TABLENS,
            Self::Draw => ns::DRAWNS,
            Self::Style => ns::STYLENS,
            Self::Fo => ns::FONS,
            Self::Xlink => ns::XLINKNS,
            Self::Svg => ns::SVGNS,
            Self::Dc => ns::DCNS,
            Self::Meta => ns::METANS,
            Self::Xml => ns::XMLNS,
            Self::Loext => ns::LOEXTNS,
        }
    }

    /// The conventional prefix, used when the part does not declare one.
    pub(crate) const fn conventional(self) -> &'static str {
        match self {
            Self::Office => "office",
            Self::Text => "text",
            Self::Table => "table",
            Self::Draw => "draw",
            Self::Style => "style",
            Self::Fo => "fo",
            Self::Xlink => "xlink",
            Self::Svg => "svg",
            Self::Dc => "dc",
            Self::Meta => "meta",
            Self::Xml => "xml",
            Self::Loext => "loext",
        }
    }

    const fn slot(self) -> usize {
        self as usize
    }
}

/// The prefixes one XML part binds to the namespaces of [`Ns`], read from
/// its root element (where every ODF producer declares them).
#[derive(Clone, Debug, Default)]
pub(crate) struct Prefixes {
    bound: [Option<Box<str>>; 12],
}

impl Prefixes {
    /// Read the `xmlns:*` declarations of the first start tag of `xml`.
    pub(crate) fn of_root(xml: &str) -> Self {
        let mut prefixes = Self::default();
        let bytes = xml.as_bytes();
        let mut cursor = 0;
        // Skip the declaration, comments and processing instructions.
        while let Some(offset) = memchr::memchr(b'<', &bytes[cursor..]) {
            let start = cursor + offset;
            match bytes.get(start + 1) {
                Some(b'?' | b'!') => {
                    cursor = start + 1;
                },
                Some(_) => {
                    if let Some(end) = tag_end(bytes, start) {
                        for (name, value) in attributes(&xml[start..end]) {
                            if let Some(prefix) = name.strip_prefix("xmlns:") {
                                let value = unescape(value);
                                if let Some(ns) = Ns::ALL.iter().find(|ns| ns.uri() == value) {
                                    prefixes.bound[ns.slot()] = Some(prefix.into());
                                }
                            }
                        }
                    }
                    break;
                },
                None => break,
            }
        }
        prefixes
    }

    /// The prefix bound to `ns`, or its conventional spelling.
    pub(crate) fn prefix(&self, ns: Ns) -> &str {
        if ns == Ns::Xml {
            return "xml";
        }
        self.bound[ns.slot()]
            .as_deref()
            .unwrap_or_else(|| ns.conventional())
    }

    /// Whether the part declares a prefix for `ns`.
    pub(crate) fn is_bound(&self, ns: Ns) -> bool {
        ns == Ns::Xml || self.bound[ns.slot()].is_some()
    }

    /// The namespace a prefix is bound to, if it is one of [`Ns`].
    pub(crate) fn namespace(&self, prefix: &[u8]) -> Option<Ns> {
        if prefix == b"xml" {
            return Some(Ns::Xml);
        }
        Ns::ALL.iter().copied().find(|ns| {
            self.bound[ns.slot()].as_deref().map_or_else(
                || ns.conventional().as_bytes() == prefix,
                |bound| bound.as_bytes() == prefix,
            )
        })
    }

    /// Whether the qualified `name` is `ns:local`.
    pub(crate) fn is(&self, name: &[u8], ns: Ns, local: &str) -> bool {
        let prefix = self.prefix(ns).as_bytes();
        name.len() == prefix.len() + 1 + local.len()
            && name.starts_with(prefix)
            && name[prefix.len()] == b':'
            && &name[prefix.len() + 1..] == local.as_bytes()
    }

    /// Split a qualified name into its namespace and local name.
    pub(crate) fn resolve<'n>(&self, name: &'n [u8]) -> Option<(Ns, &'n [u8])> {
        let colon = memchr::memchr(b':', name)?;
        Some((self.namespace(&name[..colon])?, &name[colon + 1..]))
    }

    /// Write `ns:local` into `out`.
    pub(crate) fn push_name(&self, out: &mut String, ns: Ns, local: &str) {
        out.push_str(self.prefix(ns));
        out.push(':');
        out.push_str(local);
    }
}

/// The offset just past the `>` closing the tag that starts at `start`.
pub(crate) fn tag_end(bytes: &[u8], start: usize) -> Option<usize> {
    let mut quote = None;
    let mut index = start + 1;
    while index < bytes.len() {
        let byte = bytes[index];
        match quote {
            Some(open) if byte == open => quote = None,
            Some(_) => {},
            None => match byte {
                b'"' | b'\'' => quote = Some(byte),
                b'>' => return Some(index + 1),
                _ => {},
            },
        }
        index += 1;
    }
    None
}

/// The qualified name of the tag `tag` (starting with `<`).
pub(crate) fn tag_name(tag: &str) -> &str {
    let body = tag.trim_start_matches('<').trim_start_matches('/');
    let end = body
        .find(|c: char| c.is_ascii_whitespace() || c == '/' || c == '>')
        .unwrap_or(body.len());
    &body[..end]
}

/// The local part of a qualified element name (`math` of `math:math`).
pub(crate) fn tag_name_local(name: &[u8]) -> &[u8] {
    match memchr::memchr(b':', name) {
        Some(colon) => &name[colon + 1..],
        None => name,
    }
}

/// The raw (still escaped) attributes of one start tag, in source order.
pub(crate) fn attributes(tag: &str) -> Attributes<'_> {
    let name = tag_name(tag);
    let offset = tag.find(name).map_or(tag.len(), |index| index + name.len());
    Attributes {
        rest: &tag[offset..],
    }
}

/// An iterator over `(name, raw value)` pairs of a start tag.
pub(crate) struct Attributes<'a> {
    rest: &'a str,
}

impl<'a> Iterator for Attributes<'a> {
    type Item = (&'a str, &'a str);

    fn next(&mut self) -> Option<Self::Item> {
        let rest = self.rest.trim_start();
        // A byte scan: every delimiter is ASCII, so the position is a char
        // boundary (and the scan stays inlined into each caller).
        let name_end = rest.bytes().position(|byte| {
            byte == b'=' || byte.is_ascii_whitespace() || byte == b'>' || byte == b'/'
        })?;
        let name = &rest[..name_end];
        if name.is_empty() {
            self.rest = "";
            return None;
        }
        let after = rest[name_end..].trim_start();
        let after = after.strip_prefix('=')?.trim_start();
        let quote = after.chars().next()?;
        if quote != '"' && quote != '\'' {
            self.rest = "";
            return None;
        }
        let value_rest = &after[1..];
        let quote = if quote == '"' { b'"' } else { b'\'' };
        let close = value_rest.bytes().position(|byte| byte == quote)?;
        let value = &value_rest[..close];
        self.rest = &value_rest[close + 1..];
        Some((name, value))
    }
}

/// The byte range (relative to `tag`) of the whole `name="value"` pair named
/// `qualified`, and of its value.
pub(crate) fn attribute_span(
    tag: &str,
    qualified: &str,
) -> Option<(std::ops::Range<usize>, std::ops::Range<usize>)> {
    let base = tag.as_ptr() as usize;
    for (name, value) in attributes(tag) {
        if name == qualified {
            let name_start = name.as_ptr() as usize - base;
            let value_start = value.as_ptr() as usize - base;
            let value_end = value_start + value.len();
            // Include the closing quote.
            return Some((name_start..value_end + 1, value_start..value_end));
        }
    }
    None
}

/// Decode the five predefined entities and numeric character references.
/// Borrows when `raw` holds no reference.
pub(crate) fn unescape(raw: &str) -> Cow<'_, str> {
    if !raw.contains('&') {
        return Cow::Borrowed(raw);
    }
    let mut out = String::with_capacity(raw.len());
    let mut rest = raw;
    while let Some(index) = rest.find('&') {
        out.push_str(&rest[..index]);
        rest = &rest[index..];
        let Some(end) = rest.find(';') else {
            out.push_str(rest);
            return Cow::Owned(out);
        };
        let entity = &rest[1..end];
        let decoded = match entity {
            "amp" => Some('&'),
            "lt" => Some('<'),
            "gt" => Some('>'),
            "quot" => Some('"'),
            "apos" => Some('\''),
            _ => entity.strip_prefix('#').and_then(|number| {
                let value = match number
                    .strip_prefix('x')
                    .or_else(|| number.strip_prefix('X'))
                {
                    Some(hex) => u32::from_str_radix(hex, 16).ok(),
                    None => number.parse().ok(),
                };
                value.and_then(char::from_u32)
            }),
        };
        match decoded {
            Some(character) => out.push(character),
            None => out.push_str(&rest[..=end]),
        }
        rest = &rest[end + 1..];
    }
    out.push_str(rest);
    Cow::Owned(out)
}

/// Escape character data or an attribute value.
pub(crate) fn escape_into(out: &mut String, text: &str) {
    for character in text.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            other => out.push(other),
        }
    }
}

/// Refuse text no XML 1.0 document can carry.
pub(crate) fn check_text(text: &str, field: &str) -> Result<()> {
    if text.len() > super::MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODT {field} exceeds {} bytes",
            super::MAX_TEXT_BYTES
        )));
    }
    if let Some(character) = text
        .chars()
        .find(|c| matches!(*c as u32, 0..=0x8 | 0xb | 0xc | 0xe..=0x1f | 0xfffe | 0xffff))
    {
        return Err(Error::InvalidFormat(format!(
            "ODT {field} contains U+{:04X}, which XML cannot carry",
            character as u32
        )));
    }
    Ok(())
}

/// Encode plain text as ODF inline content: runs of spaces become
/// `text:s`, tabs `text:tab` and line feeds `text:line-break`, so the
/// published markup keeps every character (ODF collapses white space in
/// character data).
pub(crate) fn encode_text(out: &mut String, prefixes: &Prefixes, text: &str) {
    let text_prefix = prefixes.prefix(Ns::Text);
    let mut chars = text.chars().peekable();
    let mut at_start = true;
    while let Some(character) = chars.next() {
        match character {
            ' ' => {
                // One space survives collapsing unless it starts the content
                // or follows another space.
                let mut count = 1usize;
                while chars.peek() == Some(&' ') {
                    chars.next();
                    count += 1;
                }
                let literal = usize::from(!at_start && chars.peek().is_some());
                if literal == 1 {
                    out.push(' ');
                }
                let encoded = count - literal;
                if encoded == 1 {
                    out.push('<');
                    out.push_str(text_prefix);
                    out.push_str(":s/>");
                } else if encoded > 1 {
                    out.push('<');
                    out.push_str(text_prefix);
                    out.push_str(":s ");
                    out.push_str(text_prefix);
                    out.push_str(":c=\"");
                    out.push_str(&encoded.to_string());
                    out.push_str("\"/>");
                }
            },
            '\t' => {
                out.push('<');
                out.push_str(text_prefix);
                out.push_str(":tab/>");
            },
            '\n' => {
                out.push('<');
                out.push_str(text_prefix);
                out.push_str(":line-break/>");
            },
            '\r' => {},
            other => {
                let mut buffer = [0; 4];
                escape_into(out, other.encode_utf8(&mut buffer));
            },
        }
        at_start = false;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attributes_scan_quotes_and_spacing() {
        let tag = r#"<text:p text:style-name="P1"  xml:id = 'a&amp;b' >"#;
        let pairs: Vec<_> = attributes(tag).collect();
        assert_eq!(
            pairs,
            vec![("text:style-name", "P1"), ("xml:id", "a&amp;b")]
        );
        assert_eq!(unescape("a&amp;b&#x41;&#66;"), "a&bAB");
        let (whole, value) = attribute_span(tag, "text:style-name").unwrap();
        assert_eq!(&tag[whole], r#"text:style-name="P1""#);
        assert_eq!(&tag[value], "P1");
    }

    #[test]
    fn encode_text_keeps_every_space() {
        let prefixes = Prefixes::default();
        let mut out = String::new();
        encode_text(&mut out, &prefixes, " a  b\tc\n<");
        assert_eq!(
            out,
            "<text:s/>a <text:s/>b<text:tab/>c<text:line-break/>&lt;"
        );
        let mut trailing = String::new();
        encode_text(&mut trailing, &prefixes, "a ");
        assert_eq!(trailing, "a<text:s/>");
    }
}
