//! Typed ODF paragraph flow and pagination properties.
use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;
const O: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const SS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65536;
const MAX_VALUE: usize = 4096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_COUNT: u32 = 1_000_000;
fn bad(x: impl Into<String>) -> Error {
    Error::InvalidFormat(x.into())
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Keep {
    Auto,
    Always,
}
impl Keep {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "auto" => Ok(Self::Auto),
            "always" => Ok(Self::Always),
            _ => Err(bad("invalid ODF keep value")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Always => "always",
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyphenationKeep {
    Auto,
    Page,
}
impl HyphenationKeep {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "auto" => Ok(Self::Auto),
            "page" => Ok(Self::Page),
            _ => Err(bad("invalid fo:hyphenation-keep")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Page => "page",
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyphenationLadder {
    NoLimit,
    Lines(u32),
}
impl HyphenationLadder {
    fn parse(x: &str) -> Result<Self> {
        if x == "no-limit" {
            return Ok(Self::NoLimit);
        }
        let n = x
            .parse()
            .map_err(|_| bad("invalid hyphenation ladder count"))?;
        if !(1..=MAX_COUNT).contains(&n) {
            return Err(bad("hyphenation ladder count out of range"));
        }
        Ok(Self::Lines(n))
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineBreak {
    Normal,
    Strict,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PunctuationWrap {
    Simple,
    Hanging,
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphFlowProperties {
    pub keep_together: Option<Keep>,
    pub keep_with_next: Option<Keep>,
    pub widows: Option<u32>,
    pub orphans: Option<u32>,
    pub hyphenation_keep: Option<HyphenationKeep>,
    pub hyphenation_ladder_count: Option<HyphenationLadder>,
    pub line_break: Option<LineBreak>,
    pub punctuation_wrap: Option<PunctuationWrap>,
}
impl ParagraphFlowProperties {
    pub fn validate(&self) -> Result<()> {
        for (n, k) in [(self.widows, "fo:widows"), (self.orphans, "fo:orphans")] {
            if let Some(n) = n {
                if n > MAX_COUNT {
                    return Err(bad(format!("{k} out of range")));
                }
            }
        }
        if let Some(HyphenationLadder::Lines(n)) = self.hyphenation_ladder_count {
            if !(1..=MAX_COUNT).contains(&n) {
                return Err(bad("hyphenation ladder count out of range"));
            }
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut x = format!(r#"<style:paragraph-properties xmlns:style="{SS}" xmlns:fo="{FS}""#);
        if let Some(v) = self.keep_together {
            x.push_str(&format!(r#" fo:keep-together="{}""#, v.xml()))
        }
        if let Some(v) = self.keep_with_next {
            x.push_str(&format!(r#" fo:keep-with-next="{}""#, v.xml()))
        }
        if let Some(v) = self.widows {
            x.push_str(&format!(r#" fo:widows="{v}""#))
        }
        if let Some(v) = self.orphans {
            x.push_str(&format!(r#" fo:orphans="{v}""#))
        }
        if let Some(v) = self.hyphenation_keep {
            x.push_str(&format!(r#" fo:hyphenation-keep="{}""#, v.xml()))
        }
        if let Some(v) = self.hyphenation_ladder_count {
            x.push_str(&format!(
                r#" fo:hyphenation-ladder-count="{}""#,
                match v {
                    HyphenationLadder::NoLimit => "no-limit".into(),
                    HyphenationLadder::Lines(n) => n.to_string(),
                }
            ))
        }
        if let Some(v) = self.line_break {
            x.push_str(&format!(
                r#" style:line-break="{}""#,
                match v {
                    LineBreak::Normal => "normal",
                    LineBreak::Strict => "strict",
                }
            ))
        }
        if let Some(v) = self.punctuation_wrap {
            x.push_str(&format!(
                r#" style:punctuation-wrap="{}""#,
                match v {
                    PunctuationWrap::Simple => "simple",
                    PunctuationWrap::Hanging => "hanging",
                }
            ))
        }
        x.push_str("/>");
        Ok(x)
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphStyleFlow {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<ParagraphFlowProperties>,
}
impl ParagraphStyleFlow {
    pub fn named(
        name: impl Into<String>,
        properties: Option<ParagraphFlowProperties>,
    ) -> Result<Self> {
        let x = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        x.validate()?;
        Ok(x)
    }
    pub fn default_style(properties: Option<ParagraphFlowProperties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(n), false)
                if !n.is_empty() && n.len() <= MAX_VALUE && !n.chars().any(char::is_control) => {},
            (None, true) => {},
            _ => return Err(bad("invalid paragraph flow style identity")),
        }
        if let Some(n) = &self.parent_style_name {
            if self.is_default_style || n.is_empty() || n.len() > MAX_VALUE {
                return Err(bad("invalid parent style name"));
            }
        }
        if let Some(p) = &self.properties {
            p.validate()?
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut x = format!(r#"<style:{tag} xmlns:style="{SS}" style:family="paragraph""#);
        if let Some(n) = &self.name {
            x.push_str(&format!(r#" style:name="{}""#, escape_xml(n)))
        }
        if let Some(n) = &self.parent_style_name {
            x.push_str(&format!(r#" style:parent-style-name="{}""#, escape_xml(n)))
        }
        if let Some(p) = &self.properties {
            x.push('>');
            x.push_str(&p.to_xml_fragment()?);
            x.push_str(&format!("</style:{tag}>"))
        } else {
            x.push_str("/>")
        }
        Ok(x)
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphStyleFlowSet {
    pub styles: Vec<ParagraphStyleFlow>,
}
impl ParagraphStyleFlowSet {
    pub fn get(&self, n: &str) -> Option<&ParagraphStyleFlow> {
        self.styles.iter().find(|x| x.name.as_deref() == Some(n))
    }
    pub fn default_style(&self) -> Option<&ParagraphStyleFlow> {
        self.styles.iter().find(|x| x.is_default_style)
    }
}
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    O,
    S,
    F,
    X,
}
fn ns(x: ResolveResult<'_>) -> Ns {
    match x {
        ResolveResult::Bound(n) if n.as_ref() == O => Ns::O,
        ResolveResult::Bound(n) if n.as_ref() == S => Ns::S,
        ResolveResult::Bound(n) if n.as_ref() == F => Ns::F,
        _ => Ns::X,
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
    let mut o = Vec::new();
    let mut seen = HashSet::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(|e| bad(format!("invalid flow attribute: {e}")))?;
        if a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (n, l) = r.resolver().resolve_attribute(a.key);
        let k = (ns(n), l.as_ref().to_vec());
        if !seen.insert(k.clone()) {
            return Err(bad("duplicate flow attribute"));
        }
        let x = a
            .decoded_and_normalized_value(v, r.decoder())
            .map_err(|e| bad(format!("invalid flow value: {e}")))?
            .into_owned();
        if x.len() > MAX_VALUE {
            return Err(bad("flow value too large"));
        }
        o.push((k.0, k.1, x))
    }
    Ok(o)
}
fn take(a: &mut Vec<(Ns, Vec<u8>, String)>, n: Ns, l: &[u8]) -> Option<String> {
    a.iter()
        .position(|x| x.0 == n && x.1 == l)
        .map(|i| a.remove(i).2)
}
fn properties(
    r: &NsReader<&[u8]>,
    v: XmlVersion,
    e: &BytesStart<'_>,
) -> Result<ParagraphFlowProperties> {
    let mut a = attrs(r, v, e)?;
    let num = |x: Option<String>, n: &str| -> Result<Option<u32>> {
        x.map(|x| {
            let v = x.parse().map_err(|_| bad(format!("invalid {n}")))?;
            if v > MAX_COUNT {
                return Err(bad(format!("{n} out of range")));
            }
            Ok(v)
        })
        .transpose()
    };
    let p = ParagraphFlowProperties {
        keep_together: take(&mut a, Ns::F, b"keep-together")
            .map(|x| Keep::parse(&x))
            .transpose()?,
        keep_with_next: take(&mut a, Ns::F, b"keep-with-next")
            .map(|x| Keep::parse(&x))
            .transpose()?,
        widows: num(take(&mut a, Ns::F, b"widows"), "fo:widows")?,
        orphans: num(take(&mut a, Ns::F, b"orphans"), "fo:orphans")?,
        hyphenation_keep: take(&mut a, Ns::F, b"hyphenation-keep")
            .map(|x| HyphenationKeep::parse(&x))
            .transpose()?,
        hyphenation_ladder_count: take(&mut a, Ns::F, b"hyphenation-ladder-count")
            .map(|x| HyphenationLadder::parse(&x))
            .transpose()?,
        line_break: take(&mut a, Ns::S, b"line-break")
            .map(|x| match x.as_str() {
                "normal" => Ok(LineBreak::Normal),
                "strict" => Ok(LineBreak::Strict),
                _ => Err(bad("invalid style:line-break")),
            })
            .transpose()?,
        punctuation_wrap: take(&mut a, Ns::S, b"punctuation-wrap")
            .map(|x| match x.as_str() {
                "simple" => Ok(PunctuationWrap::Simple),
                "hanging" => Ok(PunctuationWrap::Hanging),
                _ => Err(bad("invalid style:punctuation-wrap")),
            })
            .transpose()?,
    };
    p.validate()?;
    Ok(p)
}
pub fn parse_paragraph_style_flows(xml: &str) -> Result<ParagraphStyleFlowSet> {
    if xml.len() > MAX_XML {
        return Err(bad("XML too large"));
    }
    if ![
        "keep-together",
        "keep-with-next",
        "widows",
        "orphans",
        "hyphenation-keep",
        "hyphenation-ladder-count",
        "punctuation-wrap",
        "line-break",
    ]
    .iter()
    .any(|x| xml.contains(x))
    {
        return Ok(Default::default());
    }
    let mut r = NsReader::from_reader(xml.as_bytes());
    let mut v = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<(usize, ParagraphStyleFlow, bool)> = None;
    let mut out = Vec::new();
    let mut total = 0;
    loop {
        match r.read_event() {
            Ok(Event::Decl(d)) => {
                v = d
                    .xml_version()
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::Start(e)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("XML too deep"));
                }
                let c = elem(&r, e.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::O && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && c.0 == Ns::S
                    && matches!(c.1.as_slice(), b"style" | b"default-style");
                stack.push(c.clone());
                let d = stack.len();
                if direct {
                    let mut a = attrs(&r, v, &e)?;
                    if take(&mut a, Ns::S, b"family").as_deref() == Some("paragraph") {
                        let default = c.1 == b"default-style";
                        let style = ParagraphStyleFlow {
                            name: take(&mut a, Ns::S, b"name"),
                            parent_style_name: take(&mut a, Ns::S, b"parent-style-name"),
                            is_default_style: default,
                            properties: None,
                        };
                        style.validate()?;
                        active = Some((d, style, false))
                    }
                } else if let Some((sd, style, seen)) = active.as_mut() {
                    if d == *sd + 1 && c.0 == Ns::S && c.1 == b"paragraph-properties" {
                        if *seen {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                        *seen = true;
                        style.properties = Some(properties(&r, v, &e)?)
                    } else if c.1 == b"paragraph-properties" && c.0 != Ns::S {
                        return Err(bad("paragraph-properties wrong namespace"));
                    }
                }
            },
            Ok(Event::Empty(e)) => {
                let c = elem(&r, e.name());
                let d = stack.len() + 1;
                if let Some((sd, style, seen)) = active.as_mut() {
                    if d == *sd + 1 && c.0 == Ns::S && c.1 == b"paragraph-properties" {
                        if *seen {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                        *seen = true;
                        style.properties = Some(properties(&r, v, &e)?)
                    } else if c.1 == b"paragraph-properties" && c.0 != Ns::S {
                        return Err(bad("paragraph-properties wrong namespace"));
                    }
                }
            },
            Ok(Event::End(_)) => {
                let d = stack.len();
                if active.as_ref().is_some_and(|x| x.0 == d) {
                    let s = active.take().unwrap().1;
                    total += s.to_xml_fragment()?.len();
                    if out.len() >= MAX_STYLES || total > MAX_TOTAL {
                        return Err(bad("too many paragraph flow styles"));
                    }
                    if out.iter().any(|x: &ParagraphStyleFlow| {
                        x.name == s.name && x.is_default_style == s.is_default_style
                    }) {
                        return Err(bad("duplicate paragraph style identity"));
                    }
                    out.push(s)
                }
                stack.pop();
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid XML: {e}"))),
        }
    }
    Ok(ParagraphStyleFlowSet { styles: out })
}
impl OpenDocumentPackage {
    pub fn paragraph_style_flows(&self) -> Result<ParagraphStyleFlowSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |x| parse_paragraph_style_flows(&x),
        )
    }
}
impl FlatOpenDocument {
    pub fn paragraph_style_flows(&self) -> Result<ParagraphStyleFlowSet> {
        parse_paragraph_style_flows(self.xml())
    }
}
