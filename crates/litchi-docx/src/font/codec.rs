//! Bounded Strict/Transitional XML codec for the WordprocessingML font table.

use crate::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::{XmlVersion, events::Event, reader::Reader};
use std::collections::{HashMap, HashSet};
use std::fmt;

use super::model::{
    Charset, Conformance, Embed, Family, Font, FontKey, MAX_DEPTH, MAX_FONTS, MAX_NODES, MAX_TEXT,
    MAX_XML, Pitch, RS, RT, Signature, Style, Table, WS, WT, XMLNS, bounded, invalid, raw, rel_ns,
    validate_attr_name, validate_table_value, word_ns,
};

impl Table {
    pub fn xml(&self, conformance: Conformance) -> Result<Vec<u8>> {
        write(self, conformance)
    }
}

#[derive(Clone)]
struct XmlAttr {
    q: String,
    local: String,
    ns: String,
    value: String,
}

#[derive(Clone)]
struct Node {
    q: String,
    local: String,
    ns: String,
    attrs: Vec<XmlAttr>,
    children: Vec<Node>,
    text: String,
}

/// Parse one bounded `fontTable.xml` payload without resolving OPC resources.
pub fn parse(xml: &[u8]) -> Result<Table> {
    if xml.len() > MAX_XML {
        return Err(invalid("font-table part is too large"));
    }
    let xml = process_ooxml(xml)?;
    if xml.len() > MAX_XML {
        return Err(invalid("MCE-expanded font table is too large"));
    }
    parse_table_node(&parse_tree(xml.as_ref())?)
}

fn parse_tree(xml: &[u8]) -> Result<Node> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut stack = Vec::<Node>::new();
    let mut scopes = vec![HashMap::<String, String>::new()];
    let mut root = None;
    let mut count = 0usize;
    loop {
        let decoder = reader.decoder();
        match reader.read_event_into(&mut buf).map_err(xml_error)? {
            Event::Start(e) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("font-table XML resource limit exceeded"));
                }
                let parent = scopes
                    .last()
                    .ok_or_else(|| invalid("font-table namespace scope is missing"))?;
                let (n, s) = make_node(&e, decoder, parent)?;
                stack.push(n);
                scopes.push(s)
            },
            Event::Empty(e) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("font-table XML resource limit exceeded"));
                }
                let parent = scopes
                    .last()
                    .ok_or_else(|| invalid("font-table namespace scope is missing"))?;
                let (n, _) = make_node(&e, decoder, parent)?;
                attach(n, &mut stack, &mut root)?
            },
            Event::End(_) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end tag"))?;
                if scopes.len() <= 1 {
                    return Err(invalid("font-table namespace scope underflow"));
                }
                scopes.pop();
                attach(n, &mut stack, &mut root)?
            },
            Event::Text(e) => {
                let d = e.decode().map_err(xml_error)?;
                let d = quick_xml::escape::unescape(&d).map_err(xml_error)?;
                append_text(&mut stack, &d)?
            },
            Event::CData(e) => append_text(&mut stack, &e.decode().map_err(xml_error)?)?,
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) | Event::GeneralRef(_) => {},
            Event::Eof => break,
        }
        buf.clear()
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated font-table XML"));
    }
    root.ok_or_else(|| invalid("font-table part has no root"))
}

fn make_node(
    e: &quick_xml::events::BytesStart<'_>,
    d: quick_xml::encoding::Decoder,
    parent: &HashMap<String, String>,
) -> Result<(Node, HashMap<String, String>)> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut scope = parent.clone();
    let mut raw = Vec::new();
    let mut names = HashSet::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        let n = std::str::from_utf8(a.key.as_ref())
            .map_err(xml_error)?
            .to_string();
        if !names.insert(n.clone()) {
            return Err(invalid("duplicate XML attribute"));
        }
        let v = a
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
            .map_err(xml_error)?
            .into_owned();
        bounded(&v)?;
        if n == "xmlns" {
            scope.insert(String::new(), v.clone());
        } else if let Some(p) = n.strip_prefix("xmlns:") {
            scope.insert(p.into(), v.clone());
        }
        raw.push((n, v))
    }
    let ns = resolve(&q, &scope, true)?;
    let mut attrs = Vec::new();
    for (n, v) in raw {
        let ans = if n == "xmlns" || n.starts_with("xmlns:") {
            XMLNS.into()
        } else {
            resolve(&n, &scope, false)?
        };
        attrs.push(XmlAttr {
            local: local(&n).into(),
            q: n,
            ns: ans,
            value: v,
        })
    }
    Ok((
        Node {
            local: local(&q).into(),
            q,
            ns,
            attrs,
            children: Vec::new(),
            text: String::new(),
        },
        scope,
    ))
}
fn resolve(q: &str, scope: &HashMap<String, String>, default: bool) -> Result<String> {
    if let Some((p, _)) = q.split_once(':') {
        scope
            .get(p)
            .cloned()
            .ok_or_else(|| invalid(format!("unbound XML prefix '{p}'")))
    } else if default {
        Ok(scope.get("").cloned().unwrap_or_default())
    } else {
        Ok(String::new())
    }
}
fn local(q: &str) -> &str {
    q.rsplit_once(':').map_or(q, |(_, v)| v)
}
fn attach(n: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.children.push(n)
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple font-table roots"));
    }
    Ok(())
}
fn append_text(stack: &mut [Node], v: &str) -> Result<()> {
    if let Some(n) = stack.last_mut() {
        if n.text.len().saturating_add(v.len()) > MAX_TEXT {
            return Err(invalid("font-table text limit exceeded"));
        }
        n.text.push_str(v)
    } else if !v.trim().is_empty() {
        return Err(invalid("text outside font-table root"));
    }
    Ok(())
}

struct Attributes {
    word: Vec<(String, String)>,
    rels: Vec<(String, String)>,
    extensions: Vec<raw::Attr>,
}
impl Attributes {
    fn new(n: &Node, w: &[&str], r: &[&str]) -> Result<Self> {
        let mut word = Vec::new();
        let mut rels = Vec::new();
        let mut extensions = Vec::new();
        for a in &n.attrs {
            if a.ns == XMLNS {
                continue;
            }
            if word_ns(&a.ns) && w.contains(&a.local.as_str()) {
                if word.iter().any(|(x, _)| x == &a.local) {
                    return Err(invalid("duplicate semantic Word attribute"));
                }
                word.push((a.local.clone(), a.value.clone()))
            } else if rel_ns(&a.ns) && r.contains(&a.local.as_str()) {
                if rels.iter().any(|(x, _)| x == &a.local) {
                    return Err(invalid("duplicate semantic relationship attribute"));
                }
                rels.push((a.local.clone(), a.value.clone()))
            } else if !a.ns.is_empty() && !word_ns(&a.ns) && !rel_ns(&a.ns) {
                extensions.push(raw::Attr {
                    qualified_name: a.q.clone(),
                    value: a.value.clone(),
                })
            } else {
                return Err(invalid(format!(
                    "unexpected attribute '{}' on '{}'",
                    a.q, n.q
                )));
            }
        }
        Ok(Self {
            word,
            rels,
            extensions,
        })
    }
    fn opt(&self, n: &str) -> Result<Option<String>> {
        let v = self
            .word
            .iter()
            .find(|(k, _)| k == n)
            .map(|(_, v)| v.clone());
        if let Some(v) = &v {
            bounded(v)?
        }
        Ok(v)
    }
    fn req(&self, n: &str) -> Result<String> {
        self.opt(n)?
            .ok_or_else(|| invalid(format!("missing w:{n}")))
    }
    fn rel(&self, n: &str) -> Result<String> {
        self.rels
            .iter()
            .find(|(k, _)| k == n)
            .map(|(_, v)| v.clone())
            .ok_or_else(|| invalid(format!("missing r:{n}")))
    }
}

fn parse_table_node(root: &Node) -> Result<Table> {
    require(root, "fonts")?;
    whitespace(root)?;
    let a = Attributes::new(root, &[], &[])?;
    if root.children.len() > MAX_FONTS {
        return Err(invalid("too many fonts"));
    }
    let mut fonts = Vec::with_capacity(root.children.len());
    for n in &root.children {
        require(n, "font")?;
        fonts.push(parse_font(n)?)
    }
    let table = Table {
        fonts,
        namespaces: extension_namespaces(root)?,
        extension_attributes: a.extensions,
    };
    validate_table_value(&table, false)?;
    Ok(table)
}
fn parse_font(n: &Node) -> Result<Font> {
    whitespace(n)?;
    let a = Attributes::new(n, &["name"], &[])?;
    let name = a.req("name")?;
    let (mut alt, mut panose, mut charset, mut family, mut not_tt, mut pitch, mut sig) =
        (None, None, None, None, None, None, None);
    let mut embedded = Vec::new();
    let mut phase = 0u8;
    for c in &n.children {
        require(c, &c.local)?;
        whitespace(c)?;
        let p = match c.local.as_str() {
            "altName" => 1,
            "panose1" => 2,
            "charset" => 3,
            "family" => 4,
            "notTrueType" => 5,
            "pitch" => 6,
            "sig" => 7,
            "embedRegular" => 8,
            "embedBold" => 9,
            "embedItalic" => 10,
            "embedBoldItalic" => 11,
            _ => return Err(invalid(format!("unexpected font child '{}'", c.q))),
        };
        if p <= phase {
            return Err(invalid(format!(
                "duplicate or out-of-order font child '{}'",
                c.local
            )));
        }
        phase = p;
        match c.local.as_str() {
            "altName" => {
                leaf(c)?;
                alt = Some(Attributes::new(c, &["val"], &[])?.req("val")?)
            },
            "panose1" => {
                leaf(c)?;
                panose = Some(fixed_hex::<10>(
                    &Attributes::new(c, &["val"], &[])?.req("val")?,
                    "PANOSE",
                )?)
            },
            "charset" => {
                leaf(c)?;
                charset = parse_charset(c)?
            },
            "family" => {
                leaf(c)?;
                family = Some(Family::parse(
                    &Attributes::new(c, &["val"], &[])?.req("val")?,
                )?)
            },
            "notTrueType" => {
                leaf(c)?;
                not_tt = Some(on_off(
                    &Attributes::new(c, &["val"], &[])?
                        .opt("val")?
                        .unwrap_or_else(|| "true".into()),
                )?)
            },
            "pitch" => {
                leaf(c)?;
                pitch = Some(Pitch::parse(
                    &Attributes::new(c, &["val"], &[])?.req("val")?,
                )?)
            },
            "sig" => {
                leaf(c)?;
                sig = Some(parse_sig(c)?)
            },
            "embedRegular" => embedded.push(parse_embed(c, Style::Regular)?),
            "embedBold" => embedded.push(parse_embed(c, Style::Bold)?),
            "embedItalic" => embedded.push(parse_embed(c, Style::Italic)?),
            "embedBoldItalic" => embedded.push(parse_embed(c, Style::BoldItalic)?),
            unexpected => return Err(invalid(format!("unexpected font child '{unexpected}'"))),
        }
    }
    Ok(Font {
        name,
        alternate_name: alt,
        panose,
        character_set: charset,
        family,
        not_true_type: not_tt,
        pitch,
        signature: sig,
        embedded_fonts: embedded,
        extension_attributes: a.extensions,
    })
}
fn parse_charset(n: &Node) -> Result<Option<Charset>> {
    let a = Attributes::new(n, &["val", "characterSet"], &[])?;
    let old = a
        .opt("val")?
        .map(|v| {
            if !(1..=2).contains(&v.len()) || !v.bytes().all(|b| b.is_ascii_hexdigit()) {
                return Err(invalid(format!("invalid charset '{v}'")));
            }
            u8::from_str_radix(&v, 16)
                .map(Charset::from_legacy)
                .map_err(xml_error)
        })
        .transpose()?;
    let strict = a
        .opt("characterSet")?
        .map(|v| Charset::strict(&v))
        .transpose()?;
    if old.is_some() && strict.is_some() && old != strict {
        return Err(invalid("conflicting font character sets"));
    }
    Ok(strict.or(old))
}
fn parse_sig(n: &Node) -> Result<Signature> {
    let a = Attributes::new(n, &["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"], &[])?;
    let p = |name: &str| -> Result<u32> {
        let v = a.req(name)?;
        if v.len() != 8 || !v.bytes().all(|b| b.is_ascii_hexdigit()) {
            return Err(invalid(format!("invalid font signature '{name}'")));
        }
        u32::from_str_radix(&v, 16).map_err(xml_error)
    };
    Ok(Signature {
        unicode_subsets: [p("usb0")?, p("usb1")?, p("usb2")?, p("usb3")?],
        code_pages: [p("csb0")?, p("csb1")?],
    })
}
fn parse_embed(n: &Node, style: Style) -> Result<Embed> {
    leaf(n)?;
    let a = Attributes::new(n, &["fontKey", "subsetted"], &["id"])?;
    let key = a
        .opt("fontKey")?
        .map(|value| value.parse::<FontKey>())
        .transpose()?;
    Ok(Embed {
        style,
        relationship_id: a.rel("id")?,
        font_key: key,
        subsetted: a.opt("subsetted")?.map(|v| on_off(&v)).transpose()?,
        resource: None,
        extension_attributes: a.extensions,
    })
}

/// Serialize one font table using the requested OOXML conformance family.
pub fn write(t: &Table, c: Conformance) -> Result<Vec<u8>> {
    validate_table_value(t, false)?;
    let mut o = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
    o.extend_from_slice(b"<w:fonts xmlns:w=\"");
    esc(&mut o, c.word());
    o.extend_from_slice(b"\" xmlns:r=\"");
    esc(&mut o, c.rel());
    o.push(b'\"');
    for a in &t.namespaces {
        preserved(&mut o, a)?
    }
    extensions(&mut o, &t.extension_attributes)?;
    if t.fonts.is_empty() {
        o.extend_from_slice(b"/>");
        return Ok(o);
    }
    o.push(b'>');
    for f in &t.fonts {
        write_font(&mut o, f, c)?
    }
    o.extend_from_slice(b"</w:fonts>");
    Ok(o)
}
fn write_font(o: &mut Vec<u8>, f: &Font, c: Conformance) -> Result<()> {
    o.extend_from_slice(b"<w:font");
    extensions(o, &f.extension_attributes)?;
    wa(o, "name", &f.name);
    let empty = f.alternate_name.is_none()
        && f.panose.is_none()
        && f.character_set.is_none()
        && f.family.is_none()
        && f.not_true_type.is_none()
        && f.pitch.is_none()
        && f.signature.is_none()
        && f.embedded_fonts.is_empty();
    if empty {
        o.extend_from_slice(b"/>");
        return Ok(());
    }
    o.push(b'>');
    if let Some(v) = &f.alternate_name {
        value_leaf(o, "altName", v)
    }
    if let Some(v) = f.panose {
        value_leaf(o, "panose1", &hex(&v))
    }
    if let Some(v) = f.character_set {
        o.extend_from_slice(b"<w:charset");
        match c {
            Conformance::Transitional => wa(o, "val", &format!("{:02X}", v.legacy_code())),
            Conformance::Strict => wa(
                o,
                "characterSet",
                v.strict_name()
                    .ok_or_else(|| invalid("legacy charset has no Strict representation"))?,
            ),
        }
        o.extend_from_slice(b"/>")
    }
    if let Some(v) = f.family {
        value_leaf(o, "family", v.text())
    }
    if let Some(v) = f.not_true_type {
        o.extend_from_slice(b"<w:notTrueType");
        if !v {
            wa(o, "val", "0")
        }
        o.extend_from_slice(b"/>")
    }
    if let Some(v) = f.pitch {
        value_leaf(o, "pitch", v.text())
    }
    if let Some(v) = &f.signature {
        o.extend_from_slice(b"<w:sig");
        for (i, x) in v.unicode_subsets.iter().enumerate() {
            wa(o, &format!("usb{i}"), &format!("{x:08X}"))
        }
        for (i, x) in v.code_pages.iter().enumerate() {
            wa(o, &format!("csb{i}"), &format!("{x:08X}"))
        }
        o.extend_from_slice(b"/>")
    }
    for e in &f.embedded_fonts {
        o.extend_from_slice(b"<w:");
        o.extend_from_slice(e.style.element().as_bytes());
        extensions(o, &e.extension_attributes)?;
        ra(o, "id", &e.relationship_id);
        if let Some(v) = e.font_key {
            wa(o, "fontKey", &v.to_string())
        }
        if let Some(v) = e.subsetted {
            wa(o, "subsetted", if v { "1" } else { "0" })
        }
        o.extend_from_slice(b"/>")
    }
    o.extend_from_slice(b"</w:font>");
    Ok(())
}
fn value_leaf(o: &mut Vec<u8>, n: &str, v: &str) {
    o.extend_from_slice(b"<w:");
    o.extend_from_slice(n.as_bytes());
    wa(o, "val", v);
    o.extend_from_slice(b"/>")
}
fn wa(o: &mut Vec<u8>, n: &str, v: &str) {
    attr(o, &format!("w:{n}"), v)
}
fn ra(o: &mut Vec<u8>, n: &str, v: &str) {
    attr(o, &format!("r:{n}"), v)
}
fn extensions(o: &mut Vec<u8>, v: &[raw::Attr]) -> Result<()> {
    for a in v {
        preserved(o, a)?
    }
    Ok(())
}
fn preserved(o: &mut Vec<u8>, a: &raw::Attr) -> Result<()> {
    validate_attr_name(&a.qualified_name)?;
    attr(o, &a.qualified_name, &a.value);
    Ok(())
}
fn attr(o: &mut Vec<u8>, n: &str, v: &str) {
    o.push(b' ');
    o.extend_from_slice(n.as_bytes());
    o.extend_from_slice(b"=\"");
    esc(o, v);
    o.push(b'\"')
}
fn esc(o: &mut Vec<u8>, v: &str) {
    for c in v.chars() {
        match c {
            '&' => o.extend_from_slice(b"&amp;"),
            '<' => o.extend_from_slice(b"&lt;"),
            '"' => o.extend_from_slice(b"&quot;"),
            '\t' => o.extend_from_slice(b"&#x9;"),
            '\n' => o.extend_from_slice(b"&#xA;"),
            '\r' => o.extend_from_slice(b"&#xD;"),
            _ => {
                let mut b = [0; 4];
                o.extend_from_slice(c.encode_utf8(&mut b).as_bytes())
            },
        }
    }
}

fn extension_namespaces(root: &Node) -> Result<Vec<raw::Attr>> {
    fn walk(n: &Node, map: &mut HashMap<String, String>, out: &mut Vec<raw::Attr>) -> Result<()> {
        for a in &n.attrs {
            if a.ns != XMLNS
                || matches!(
                    a.value.as_str(),
                    WT | WS
                        | RT
                        | RS
                        | "http://schemas.openxmlformats.org/markup-compatibility/2006"
                )
            {
                continue;
            }
            let p = a.q.strip_prefix("xmlns:").unwrap_or("").to_string();
            if let Some(v) = map.get(&p) {
                if v != &a.value {
                    return Err(invalid(format!("conflicting namespace prefix '{p}'")));
                }
            } else {
                map.insert(p, a.value.clone());
                out.push(raw::Attr {
                    qualified_name: a.q.clone(),
                    value: a.value.clone(),
                })
            }
        }
        for c in &n.children {
            walk(c, map, out)?
        }
        Ok(())
    }
    let mut out = Vec::new();
    walk(root, &mut HashMap::new(), &mut out)?;
    Ok(out)
}
fn fixed_hex<const N: usize>(v: &str, name: &str) -> Result<[u8; N]> {
    if v.len() != N * 2 || !v.bytes().all(|b| b.is_ascii_hexdigit()) {
        return Err(invalid(format!("invalid {name}")));
    }
    let mut out = [0; N];
    for (x, pair) in out.iter_mut().zip(v.as_bytes().chunks_exact(2)) {
        let pair = std::str::from_utf8(pair).map_err(xml_error)?;
        *x = u8::from_str_radix(pair, 16).map_err(xml_error)?
    }
    Ok(out)
}
fn hex<const N: usize>(v: &[u8; N]) -> String {
    let mut s = String::with_capacity(N * 2);
    for b in v {
        s.push_str(&format!("{b:02X}"))
    }
    s
}
fn on_off(v: &str) -> Result<bool> {
    match v {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(invalid(format!("invalid on/off value '{v}'"))),
    }
}
fn require(n: &Node, name: &str) -> Result<()> {
    if word_ns(&n.ns) && n.local == name {
        Ok(())
    } else {
        Err(invalid(format!("expected w:{name}, found '{}'", n.q)))
    }
}
fn whitespace(n: &Node) -> Result<()> {
    if n.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in '{}'", n.q)))
    }
}
fn leaf(n: &Node) -> Result<()> {
    whitespace(n)?;
    if n.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("'{}' must be empty", n.q)))
    }
}
pub(in crate::font) fn xml_error(e: impl fmt::Display) -> Error {
    Error::Xml(e.to_string())
}
