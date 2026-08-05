//! ODF 1.2/1.3 list-level label alignment.

use crate::{FlatDocument, Package};
use litchi_core::{Error, Result, xml::escape_xml};
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
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const STYLE_S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_S: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const FO_S: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_ENTRIES: usize = 65_536;
const MAX_VALUE: usize = 4096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_LEVEL: u16 = 1024;
fn bad(s: impl Into<String>) -> Error {
    Error::InvalidFormat(s.into())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FollowedBy {
    ListTab,
    Space,
    Nothing,
}
impl FollowedBy {
    fn parse(s: &str) -> Result<Self> {
        match s {
            "listtab" => Ok(Self::ListTab),
            "space" => Ok(Self::Space),
            "nothing" => Ok(Self::Nothing),
            _ => Err(bad("invalid text:label-followed-by")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::ListTab => "listtab",
            Self::Space => "space",
            Self::Nothing => "nothing",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Length(String);
impl Length {
    pub fn new(s: impl Into<String>) -> Result<Self> {
        let s = s.into();
        if s.len() > MAX_VALUE || !length(&s) {
            return Err(bad("invalid ODF list label length"));
        }
        Ok(Self(s))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
fn length(s: &str) -> bool {
    let Some(n) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|u| s.strip_suffix(u))
    else {
        return false;
    };
    let n = n.strip_prefix('-').unwrap_or(n);
    let mut p = n.split('.');
    let a = p.next().unwrap_or("");
    let b = p.next();
    if p.next().is_some() {
        return false;
    }
    let d = |x: &str| x.bytes().all(|c| c.is_ascii_digit());
    match b {
        None => !a.is_empty() && d(a),
        Some(b) => d(a) && d(b) && (!a.is_empty() || !b.is_empty()),
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Alignment {
    pub label_followed_by: FollowedBy,
    pub list_tab_stop_position: Option<Length>,
    pub text_indent: Option<Length>,
    pub margin_left: Option<Length>,
}
impl Alignment {
    pub fn new(label_followed_by: FollowedBy) -> Self {
        Self {
            label_followed_by,
            list_tab_stop_position: None,
            text_indent: None,
            margin_left: None,
        }
    }
    pub fn validate(&self) -> Result<()> {
        for v in [
            &self.list_tab_stop_position,
            &self.text_indent,
            &self.margin_left,
        ]
        .into_iter()
        .flatten()
        {
            if !length(v.as_str()) {
                return Err(bad("invalid ODF list label length"));
            }
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut x = format!(
            r#"<style:list-level-label-alignment xmlns:style="{STYLE_S}" xmlns:text="{TEXT_S}" xmlns:fo="{FO_S}" text:label-followed-by="{}""#,
            self.label_followed_by.xml()
        );
        if let Some(v) = &self.list_tab_stop_position {
            x.push_str(&format!(
                r#" text:list-tab-stop-position="{}""#,
                escape_xml(v.as_str())
            ))
        }
        if let Some(v) = &self.text_indent {
            x.push_str(&format!(r#" fo:text-indent="{}""#, escape_xml(v.as_str())))
        }
        if let Some(v) = &self.margin_left {
            x.push_str(&format!(r#" fo:margin-left="{}""#, escape_xml(v.as_str())))
        }
        x.push_str("/>");
        Ok(x)
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    List,
    Outline,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub list_style_kind: Kind,
    pub list_style_name: String,
    pub level: u16,
    pub alignment: Alignment,
}
impl Style {
    pub fn new(name: impl Into<String>, level: u16, alignment: Alignment) -> Result<Self> {
        Self::new_in(Kind::List, name, level, alignment)
    }
    pub fn new_in(
        list_style_kind: Kind,
        name: impl Into<String>,
        level: u16,
        alignment: Alignment,
    ) -> Result<Self> {
        let x = Self {
            list_style_kind,
            list_style_name: name.into(),
            level,
            alignment,
        };
        x.validate()?;
        Ok(x)
    }
    pub fn validate(&self) -> Result<()> {
        if self.list_style_name.is_empty()
            || self.list_style_name.len() > MAX_VALUE
            || self.list_style_name.chars().any(char::is_control)
        {
            return Err(bad("invalid list style name"));
        }
        if !(1..=MAX_LEVEL).contains(&self.level) {
            return Err(bad("list level outside supported range"));
        }
        self.alignment.validate()
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub levels: Vec<Style>,
}
impl Styles {
    pub fn get(&self, name: &str, level: u16) -> Option<&Style> {
        self.get_in(Kind::List, name, level)
    }
    pub fn get_in(&self, kind: Kind, name: &str, level: u16) -> Option<&Style> {
        self.levels
            .iter()
            .find(|x| x.list_style_kind == kind && x.list_style_name == name && x.level == level)
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Text,
    Fo,
    Other,
}
fn ns(r: ResolveResult<'_>) -> Ns {
    match r {
        ResolveResult::Bound(x) if x.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(x) if x.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(x) if x.as_ref() == TEXT => Ns::Text,
        ResolveResult::Bound(x) if x.as_ref() == FO => Ns::Fo,
        _ => Ns::Other,
    }
}
fn elem(r: &NsReader<&[u8]>, q: QName<'_>) -> (Ns, Vec<u8>) {
    let (n, l) = r.resolver().resolve_element(q);
    (ns(n), l.as_ref().to_vec())
}
fn attrs(
    r: &NsReader<&[u8]>,
    v: XmlVersion,
    e: &BytesStart<'_>,
) -> Result<Vec<(Ns, Vec<u8>, String)>> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(|x| bad(format!("invalid list alignment attribute: {x}")))?;
        if a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (n, l) = r.resolver().resolve_attribute(a.key);
        let k = (ns(n), l.as_ref().to_vec());
        if !seen.insert(k.clone()) {
            return Err(bad("duplicate list alignment attribute"));
        }
        let value = a
            .decoded_and_normalized_value(v, r.decoder())
            .map_err(|x| bad(format!("invalid list alignment value: {x}")))?
            .into_owned();
        if value.len() > MAX_VALUE {
            return Err(bad("list alignment value too large"));
        }
        out.push((k.0, k.1, value));
    }
    Ok(out)
}
fn one(a: &mut Vec<(Ns, Vec<u8>, String)>, n: Ns, l: &[u8]) -> Option<String> {
    a.iter()
        .position(|x| x.0 == n && x.1 == l)
        .map(|i| a.remove(i).2)
}
fn parse_alignment(r: &NsReader<&[u8]>, v: XmlVersion, e: &BytesStart<'_>) -> Result<Alignment> {
    let mut a = attrs(r, v, e)?;
    let followed = one(&mut a, Ns::Text, b"label-followed-by")
        .ok_or_else(|| bad("missing text:label-followed-by"))?;
    let x = Alignment {
        label_followed_by: FollowedBy::parse(&followed)?,
        list_tab_stop_position: one(&mut a, Ns::Text, b"list-tab-stop-position")
            .map(Length::new)
            .transpose()?,
        text_indent: one(&mut a, Ns::Fo, b"text-indent")
            .map(Length::new)
            .transpose()?,
        margin_left: one(&mut a, Ns::Fo, b"margin-left")
            .map(Length::new)
            .transpose()?,
    };
    if !a.is_empty() {
        return Err(bad("unknown list-level-label-alignment attribute"));
    }
    x.validate()?;
    Ok(x)
}

/// Parse every modern list-level label alignment in styles or flat-document XML.
pub fn parse(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("XML too large"));
    }
    if !xml.contains("list-level-label-alignment") {
        return Ok(Default::default());
    }
    let mut r = NsReader::from_reader(xml.as_bytes());
    r.config_mut().trim_text(false);
    let mut ver = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut list: Option<(usize, String, Kind, HashSet<u16>)> = None;
    let mut level: Option<(usize, u16, bool, bool)> = None;
    let mut entries = Vec::new();
    let mut total = 0usize;
    let mut open_align = None;
    loop {
        match r.read_event() {
            Ok(Event::Decl(d)) => {
                ver = d
                    .xml_version()
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::Start(e)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("XML nesting too deep"));
                }
                if open_align.is_some() {
                    return Err(bad("list-level-label-alignment must be empty"));
                }
                let c = elem(&r, e.name());
                let parent = stack.last();
                let direct_list = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && c.0 == Ns::Text
                    && matches!(c.1.as_slice(), b"list-style" | b"outline-style");
                stack.push(c.clone());
                let depth = stack.len();
                if direct_list {
                    let mut a = attrs(&r, ver, &e)?;
                    let name = one(&mut a, Ns::Style, b"name")
                        .ok_or_else(|| bad("list style missing style:name"))?;
                    if name.is_empty() {
                        return Err(bad("empty list style name"));
                    }
                    let kind = if c.1 == b"outline-style" {
                        Kind::Outline
                    } else {
                        Kind::List
                    };
                    list = Some((depth, name, kind, HashSet::new()));
                    continue;
                }
                if let Some((ld, _, _, seen)) = list.as_mut()
                    && depth == *ld + 1
                    && c.0 == Ns::Text
                    && matches!(
                        c.1.as_slice(),
                        b"list-level-style-number"
                            | b"list-level-style-bullet"
                            | b"list-level-style-image"
                            | b"outline-level-style"
                    )
                {
                    let mut a = attrs(&r, ver, &e)?;
                    let n = one(&mut a, Ns::Text, b"level")
                        .ok_or_else(|| bad("list level missing text:level"))?
                        .parse::<u16>()
                        .map_err(|_| bad("invalid text:level"))?;
                    if !(1..=MAX_LEVEL).contains(&n) || !seen.insert(n) {
                        return Err(bad("invalid or duplicate list level"));
                    }
                    level = Some((depth, n, false, false));
                    continue;
                }
                if let Some((d, _, props, align)) = level.as_mut() {
                    if depth == *d + 1 && c.0 == Ns::Style && c.1 == b"list-level-properties" {
                        if *props {
                            return Err(bad("duplicate list-level-properties"));
                        }
                        *props = true;
                        let mut a = attrs(&r, ver, &e)?;
                        if one(&mut a, Ns::Text, b"list-level-position-and-space-mode").as_deref()
                            != Some("label-alignment")
                        {
                            return Err(bad(
                                "label alignment requires label-alignment position mode",
                            ));
                        }
                    } else if depth == *d + 2
                        && c.0 == Ns::Style
                        && c.1 == b"list-level-label-alignment"
                    {
                        if !*props || *align {
                            return Err(bad("invalid or duplicate list-level-label-alignment"));
                        }
                        *align = true;
                        let alignment = parse_alignment(&r, ver, &e)?;
                        let name = &list.as_ref().unwrap().1;
                        let item = Style::new_in(
                            list.as_ref().unwrap().2,
                            name.clone(),
                            level.as_ref().unwrap().1,
                            alignment,
                        )?;
                        total +=
                            item.list_style_name.len() + item.alignment.to_xml_fragment()?.len();
                        if entries.len() >= MAX_ENTRIES || total > MAX_TOTAL {
                            return Err(bad("too many list alignments"));
                        }
                        entries.push(item);
                        open_align = Some(depth)
                    } else if c.1 == b"list-level-label-alignment" {
                        return Err(bad(
                            "list-level-label-alignment has invalid parent or namespace",
                        ));
                    }
                } else if c.1 == b"list-level-label-alignment" {
                    return Err(bad("list-level-label-alignment has invalid parent"));
                }
            },
            Ok(Event::Empty(e)) => {
                let c = elem(&r, e.name());
                let depth = stack.len() + 1;
                if let Some((d, _, props, align)) = level.as_mut() {
                    if depth == *d + 1 && c.0 == Ns::Style && c.1 == b"list-level-properties" {
                        if *props {
                            return Err(bad("duplicate list-level-properties"));
                        }
                        *props = true;
                        return Err(bad(
                            "empty list-level-properties cannot contain label alignment",
                        ));
                    }
                    if depth == *d + 2 && c.0 == Ns::Style && c.1 == b"list-level-label-alignment" {
                        if !*props || *align {
                            return Err(bad("invalid or duplicate list-level-label-alignment"));
                        }
                        *align = true;
                        let alignment = parse_alignment(&r, ver, &e)?;
                        let item = Style::new_in(
                            list.as_ref().unwrap().2,
                            list.as_ref().unwrap().1.clone(),
                            level.as_ref().unwrap().1,
                            alignment,
                        )?;
                        total +=
                            item.list_style_name.len() + item.alignment.to_xml_fragment()?.len();
                        if entries.len() >= MAX_ENTRIES || total > MAX_TOTAL {
                            return Err(bad("too many list alignments"));
                        }
                        entries.push(item)
                    } else if c.1 == b"list-level-label-alignment" {
                        return Err(bad(
                            "list-level-label-alignment has invalid parent or namespace",
                        ));
                    }
                } else if c.1 == b"list-level-label-alignment" {
                    return Err(bad("list-level-label-alignment has invalid parent"));
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if open_align == Some(depth) {
                    open_align = None
                }
                if level.as_ref().is_some_and(|x| x.0 == depth) {
                    level = None
                }
                if list.as_ref().is_some_and(|x| x.0 == depth) {
                    list = None
                }
                stack.pop();
            },
            Ok(Event::Text(t)) if open_align.is_some() => {
                let b: &[u8] = t.as_ref();
                if !b.is_empty() {
                    return Err(bad("list-level-label-alignment must be empty"));
                }
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid XML: {e}"))),
        }
    }
    Ok(Styles { levels: entries })
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML boundary"))
}
/// Replace one existing alignment element, preserving every unrelated byte.
pub(crate) fn set_xml(xml: &str, item: &Style) -> Result<String> {
    item.validate()?;
    let mut r = NsReader::from_reader(xml.as_bytes());
    let mut ver = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut list: Option<(usize, bool)> = None;
    let mut level: Option<(usize, bool)> = None;
    let mut found = None;
    loop {
        match r.read_event() {
            Ok(Event::Decl(d)) => {
                ver = d
                    .xml_version()
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::Start(e)) => {
                let end = r.buffer_position() as usize;
                let c = elem(&r, e.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::Office && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && c.0 == Ns::Text
                    && matches!(c.1.as_slice(), b"list-style" | b"outline-style");
                stack.push(c.clone());
                let d = stack.len();
                if direct {
                    let mut a = attrs(&r, ver, &e)?;
                    list = Some((
                        d,
                        one(&mut a, Ns::Style, b"name").as_deref() == Some(&item.list_style_name)
                            && (if c.1 == b"outline-style" {
                                Kind::Outline
                            } else {
                                Kind::List
                            }) == item.list_style_kind,
                    ));
                } else if list == Some((d - 1, true))
                    && c.0 == Ns::Text
                    && matches!(
                        c.1.as_slice(),
                        b"list-level-style-number"
                            | b"list-level-style-bullet"
                            | b"list-level-style-image"
                            | b"outline-level-style"
                    )
                {
                    let mut a = attrs(&r, ver, &e)?;
                    level = Some((
                        d,
                        one(&mut a, Ns::Text, b"level").and_then(|x| x.parse().ok())
                            == Some(item.level),
                    ));
                } else if level.is_some_and(|(ld, on)| on && d == ld + 2)
                    && c.0 == Ns::Style
                    && c.1 == b"list-level-label-alignment"
                {
                    let start = event_start(xml, end)?;
                    found = Some((start, 0usize, d));
                }
            },
            Ok(Event::Empty(e)) => {
                let end = r.buffer_position() as usize;
                let c = elem(&r, e.name());
                let d = stack.len() + 1;
                if level.is_some_and(|(ld, on)| on && d == ld + 2)
                    && c.0 == Ns::Style
                    && c.1 == b"list-level-label-alignment"
                {
                    if found.is_some() {
                        return Err(bad("duplicate target alignment"));
                    }
                    found = Some((event_start(xml, end)?, end, 0));
                }
            },
            Ok(Event::End(_)) => {
                let end = r.buffer_position() as usize;
                let d = stack.len();
                if let Some((s, 0, depth)) = found
                    && d == depth
                {
                    found = Some((s, end, 0));
                }
                if level.as_ref().is_some_and(|x| x.0 == d) {
                    level = None
                }
                if list.as_ref().is_some_and(|x| x.0 == d) {
                    list = None
                }
                stack.pop();
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid XML: {e}"))),
        }
    }
    let (s, e, _) = found
        .filter(|x| x.1 > 0)
        .ok_or_else(|| bad("target list-level-label-alignment does not exist"))?;
    Ok(format!(
        "{}{}{}",
        &xml[..s],
        item.alignment.to_xml_fragment()?,
        &xml[e..]
    ))
}

impl Package {
    pub fn alignments(&self) -> Result<Styles> {
        self.styles_xml()?
            .map_or_else(|| Ok(Default::default()), |x| parse(&x))
    }
}
impl FlatDocument {
    pub fn alignments(&self) -> Result<Styles> {
        parse(self.xml())
    }
}
