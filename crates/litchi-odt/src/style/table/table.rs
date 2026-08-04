//! Complete typed ODF `style:table-properties` support.

use super::row::{
    HorizontalBackgroundPosition, TableRowBackgroundColor, TableRowBackgroundImage,
    TableRowBackgroundPosition, TableRowBackgroundRepeat, TableRowBackgroundSource, TableRowBreak,
    TableRowKeepTogether, TableRowOpacity, VerticalBackgroundPosition,
};
use crate::{FlatOpenDocument, OpenDocumentPackage};
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
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE: usize = 4096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_BINARY: usize = 8 * 1024 * 1024;
const MAX_ATTRIBUTES: usize = 48;
const MAX_PAGE: u64 = 1_000_000_000;
fn bad(x: impl Into<String>) -> Error {
    Error::InvalidFormat(x.into())
}
fn safe(x: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && x.is_empty()) || x.len() > MAX_VALUE || x.chars().any(char::is_control) {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn decimal(x: &str, signed: bool) -> bool {
    let x = if signed {
        x.strip_prefix('-').unwrap_or(x)
    } else if x.starts_with('-') {
        return false;
    } else {
        x
    };
    let mut p = x.split('.');
    let a = p.next().unwrap_or_default();
    let b = p.next();
    p.next().is_none()
        && a.bytes().all(|x| x.is_ascii_digit())
        && b.is_none_or(|x| x.bytes().all(|x| x.is_ascii_digit()))
        && (!a.is_empty() || b.is_some_and(|x| !x.is_empty()))
}
fn physical(x: &str, signed: bool, positive: bool) -> bool {
    let Some(n) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|u| x.strip_suffix(u))
    else {
        return false;
    };
    decimal(n, signed) && (!positive || n.bytes().any(|x| x.is_ascii_digit() && x != b'0'))
}

/// ODF percentage lexical value, including signed percentages where the RNG permits them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStylePercent(String);
impl TableStylePercent {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        let Some(n) = x.strip_suffix('%') else {
            return Err(bad("table percentage requires %"));
        };
        if x.len() > MAX_VALUE || !decimal(n, true) {
            return Err(bad("invalid table percentage"));
        }
        Ok(Self(x))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
/// An ODF physical length or percentage used by table margins.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableStyleMeasure {
    Length(String),
    Percent(TableStylePercent),
}
impl TableStyleMeasure {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        if x.ends_with('%') {
            Ok(Self::Percent(TableStylePercent::new(x)?))
        } else if physical(&x, true, false) && x.len() <= MAX_VALUE {
            Ok(Self::Length(x))
        } else {
            Err(bad("invalid table margin measure"))
        }
    }
    pub fn as_str(&self) -> &str {
        match self {
            Self::Length(x) => x,
            Self::Percent(x) => x.as_str(),
        }
    }
    fn nonnegative_length(&self) -> bool {
        !matches!(self,Self::Length(x)if x.starts_with('-'))
    }
}
/// Positive physical table width.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyleWidth(String);
impl TableStyleWidth {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        if x.len() > MAX_VALUE || !physical(&x, false, true) {
            return Err(bad("style:width must be a positive physical length"));
        }
        Ok(Self(x))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableAlignment {
    Left,
    Center,
    Right,
    Margins,
}
impl TableAlignment {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            "margins" => Ok(Self::Margins),
            _ => Err(bad("invalid table:align")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Margins => "margins",
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TablePageNumber {
    Auto,
    Number(u64),
}
impl TablePageNumber {
    fn parse(x: &str) -> Result<Self> {
        if x == "auto" {
            return Ok(Self::Auto);
        }
        let n = x.parse().map_err(|_| bad("invalid style:page-number"))?;
        if n == 0 || n > MAX_PAGE {
            return Err(bad("style:page-number out of range"));
        }
        Ok(Self::Number(n))
    }
    fn xml(self) -> String {
        match self {
            Self::Auto => "auto".into(),
            Self::Number(x) => x.to_string(),
        }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableShadow(String);
impl TableShadow {
    pub fn new(x: impl Into<String>) -> Result<Self> {
        let x = x.into();
        safe(&x, "style:shadow", true)?;
        Ok(Self(x))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableBorderModel {
    Collapsing,
    Separating,
}
impl TableBorderModel {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "collapsing" => Ok(Self::Collapsing),
            "separating" => Ok(Self::Separating),
            _ => Err(bad("invalid table:border-model")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::Collapsing => "collapsing",
            Self::Separating => "separating",
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableWritingMode {
    LrTb,
    RlTb,
    TbRl,
    TbLr,
    Lr,
    Rl,
    Tb,
    Page,
}
impl TableWritingMode {
    fn parse(x: &str) -> Result<Self> {
        match x {
            "lr-tb" => Ok(Self::LrTb),
            "rl-tb" => Ok(Self::RlTb),
            "tb-rl" => Ok(Self::TbRl),
            "tb-lr" => Ok(Self::TbLr),
            "lr" => Ok(Self::Lr),
            "rl" => Ok(Self::Rl),
            "tb" => Ok(Self::Tb),
            "page" => Ok(Self::Page),
            _ => Err(bad("invalid style:writing-mode")),
        }
    }
    fn xml(self) -> &'static str {
        match self {
            Self::LrTb => "lr-tb",
            Self::RlTb => "rl-tb",
            Self::TbRl => "tb-rl",
            Self::TbLr => "tb-lr",
            Self::Lr => "lr",
            Self::Rl => "rl",
            Self::Tb => "tb",
            Self::Page => "page",
        }
    }
}

/// Complete `style:table-properties` value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableProperties {
    pub width: Option<TableStyleWidth>,
    pub relative_width: Option<TableStylePercent>,
    pub align: Option<TableAlignment>,
    pub margin_left: Option<TableStyleMeasure>,
    pub margin_right: Option<TableStyleMeasure>,
    pub margin_top: Option<TableStyleMeasure>,
    pub margin_bottom: Option<TableStyleMeasure>,
    pub margin: Option<TableStyleMeasure>,
    pub page_number: Option<TablePageNumber>,
    pub break_before: Option<TableRowBreak>,
    pub break_after: Option<TableRowBreak>,
    pub background_color: Option<TableRowBackgroundColor>,
    pub background_image: Option<TableRowBackgroundImage>,
    pub shadow: Option<TableShadow>,
    pub keep_with_next: Option<TableRowKeepTogether>,
    pub may_break_between_rows: Option<bool>,
    pub border_model: Option<TableBorderModel>,
    pub writing_mode: Option<TableWritingMode>,
    pub display: Option<bool>,
}
impl TableProperties {
    pub fn validate(&self) -> Result<()> {
        for (x, n) in [
            (&self.margin_top, "fo:margin-top"),
            (&self.margin_bottom, "fo:margin-bottom"),
            (&self.margin, "fo:margin"),
        ] {
            if x.as_ref().is_some_and(|x| !x.nonnegative_length()) {
                return Err(bad(format!("{n} length cannot be negative")));
            }
        }
        if let Some(TablePageNumber::Number(n)) = self.page_number
            && (n == 0 || n > MAX_PAGE)
        {
            return Err(bad("style:page-number out of range"));
        }
        if let Some(x) = &self.background_image {
            x.validate()?
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut x = format!(
            r#"<style:table-properties xmlns:style="{STYLE_NS}" xmlns:office="{OFFICE_NS}" xmlns:fo="{FO_NS}" xmlns:table="{TABLE_NS}" xmlns:draw="{DRAW_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        if let Some(v) = &self.width {
            x.push_str(&format!(r#" style:width="{}""#, v.as_str()))
        }
        if let Some(v) = &self.relative_width {
            x.push_str(&format!(r#" style:rel-width="{}""#, v.as_str()))
        }
        if let Some(v) = self.align {
            x.push_str(&format!(r#" table:align="{}""#, v.xml()))
        }
        for (n, v) in [
            ("margin-left", &self.margin_left),
            ("margin-right", &self.margin_right),
            ("margin-top", &self.margin_top),
            ("margin-bottom", &self.margin_bottom),
            ("margin", &self.margin),
        ] {
            if let Some(v) = v {
                x.push_str(&format!(r#" fo:{n}="{}""#, v.as_str()))
            }
        }
        if let Some(v) = self.page_number {
            x.push_str(&format!(r#" style:page-number="{}""#, v.xml()))
        }
        if let Some(v) = self.break_before {
            x.push_str(&format!(r#" fo:break-before="{}""#, break_xml(v)))
        }
        if let Some(v) = self.break_after {
            x.push_str(&format!(r#" fo:break-after="{}""#, break_xml(v)))
        }
        if let Some(v) = &self.background_color {
            x.push_str(&format!(r#" fo:background-color="{}""#, v.as_str()))
        }
        if let Some(v) = &self.shadow {
            x.push_str(&format!(r#" style:shadow="{}""#, escape_xml(v.as_str())))
        }
        if let Some(v) = self.keep_with_next {
            x.push_str(&format!(r#" fo:keep-with-next="{}""#, keep_xml(v)))
        }
        if let Some(v) = self.may_break_between_rows {
            x.push_str(&format!(r#" style:may-break-between-rows="{v}""#))
        }
        if let Some(v) = self.border_model {
            x.push_str(&format!(r#" table:border-model="{}""#, v.xml()))
        }
        if let Some(v) = self.writing_mode {
            x.push_str(&format!(r#" style:writing-mode="{}""#, v.xml()))
        }
        if let Some(v) = self.display {
            x.push_str(&format!(r#" table:display="{v}""#))
        }
        if let Some(v) = &self.background_image {
            x.push('>');
            x.push_str(&v.to_xml_fragment()?);
            x.push_str("</style:table-properties>")
        } else {
            x.push_str("/>")
        }
        Ok(x)
    }
}
fn break_xml(x: TableRowBreak) -> &'static str {
    match x {
        TableRowBreak::Auto => "auto",
        TableRowBreak::Column => "column",
        TableRowBreak::Page => "page",
    }
}
fn keep_xml(x: TableRowKeepTogether) -> &'static str {
    match x {
        TableRowKeepTogether::Auto => "auto",
        TableRowKeepTogether::Always => "always",
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyleProperties {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<TableProperties>,
}
impl TableStyleProperties {
    pub fn named(name: impl Into<String>, properties: Option<TableProperties>) -> Result<Self> {
        let x = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        x.validate()?;
        Ok(x)
    }
    pub fn default_style(properties: Option<TableProperties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(x), false) => safe(x, "table style name", false)?,
            (None, true) => {},
            _ => return Err(bad("invalid table style identity")),
        }
        if let Some(x) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default table style cannot have a parent"));
            }
            safe(x, "parent table style name", false)?
        }
        if let Some(x) = &self.properties {
            x.validate()?
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
        let mut x = format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="table""#);
        if let Some(v) = &self.name {
            x.push_str(&format!(r#" style:name="{}""#, escape_xml(v)))
        }
        if let Some(v) = &self.parent_style_name {
            x.push_str(&format!(r#" style:parent-style-name="{}""#, escape_xml(v)))
        }
        if let Some(v) = &self.properties {
            x.push('>');
            x.push_str(&v.to_xml_fragment()?);
            x.push_str(&format!("</style:{tag}>"))
        } else {
            x.push_str("/>")
        }
        Ok(x)
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableStylePropertiesSet {
    pub styles: Vec<TableStyleProperties>,
}
impl TableStylePropertiesSet {
    pub fn get(&self, n: &str) -> Option<&TableStyleProperties> {
        self.styles.iter().find(|x| x.name.as_deref() == Some(n))
    }
    pub fn default_style(&self) -> Option<&TableStyleProperties> {
        self.styles.iter().find(|x| x.is_default_style)
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    O,
    S,
    F,
    T,
    D,
    X,
    Z,
}
fn ns(x: ResolveResult<'_>) -> Ns {
    match x {
        ResolveResult::Bound(v) if v.as_ref() == OFFICE => Ns::O,
        ResolveResult::Bound(v) if v.as_ref() == STYLE => Ns::S,
        ResolveResult::Bound(v) if v.as_ref() == FO => Ns::F,
        ResolveResult::Bound(v) if v.as_ref() == TABLE => Ns::T,
        ResolveResult::Bound(v) if v.as_ref() == DRAW => Ns::D,
        ResolveResult::Bound(v) if v.as_ref() == XLINK => Ns::X,
        _ => Ns::Z,
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
        let a = a.map_err(|e| bad(format!("invalid table property attribute: {e}")))?;
        if a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if o.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many table property attributes"));
        }
        let (n, l) = r.resolver().resolve_attribute(a.key);
        let k = (ns(n), l.as_ref().to_vec());
        if !seen.insert(k.clone()) {
            return Err(bad("duplicate table property attribute"));
        }
        let x = a
            .decoded_and_normalized_value(v, r.decoder())
            .map_err(|e| bad(format!("invalid table property value: {e}")))?
            .into_owned();
        if x.len() > MAX_VALUE {
            return Err(bad("table property value too large"));
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
fn boolean(x: &str) -> Result<bool> {
    match x {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(bad("ODF boolean must be true or false")),
    }
}
fn parse_break(x: &str) -> Result<TableRowBreak> {
    match x {
        "auto" => Ok(TableRowBreak::Auto),
        "column" => Ok(TableRowBreak::Column),
        "page" => Ok(TableRowBreak::Page),
        _ => Err(bad("invalid table break")),
    }
}
fn parse_keep(x: &str) -> Result<TableRowKeepTogether> {
    match x {
        "auto" => Ok(TableRowKeepTogether::Auto),
        "always" => Ok(TableRowKeepTogether::Always),
        _ => Err(bad("invalid fo:keep-with-next")),
    }
}
fn header(
    r: &NsReader<&[u8]>,
    v: XmlVersion,
    e: &BytesStart<'_>,
    default: bool,
) -> Result<Option<TableStyleProperties>> {
    let mut a = attrs(r, v, e)?;
    if a.iter().any(|(n, l, _)| {
        *n != Ns::S && matches!(l.as_slice(), b"family" | b"name" | b"parent-style-name")
    }) {
        return Err(bad("table style identity attribute uses wrong namespace"));
    }
    if take(&mut a, Ns::S, b"family").as_deref() != Some("table") {
        return Ok(None);
    }
    let x = TableStyleProperties {
        name: take(&mut a, Ns::S, b"name"),
        parent_style_name: take(&mut a, Ns::S, b"parent-style-name"),
        is_default_style: default,
        properties: None,
    };
    x.validate()?;
    Ok(Some(x))
}
fn properties(r: &NsReader<&[u8]>, v: XmlVersion, e: &BytesStart<'_>) -> Result<TableProperties> {
    let mut a = attrs(r, v, e)?;
    let p = TableProperties {
        width: take(&mut a, Ns::S, b"width")
            .map(TableStyleWidth::new)
            .transpose()?,
        relative_width: take(&mut a, Ns::S, b"rel-width")
            .map(TableStylePercent::new)
            .transpose()?,
        align: take(&mut a, Ns::T, b"align")
            .map(|x| TableAlignment::parse(&x))
            .transpose()?,
        margin_left: take(&mut a, Ns::F, b"margin-left")
            .map(TableStyleMeasure::new)
            .transpose()?,
        margin_right: take(&mut a, Ns::F, b"margin-right")
            .map(TableStyleMeasure::new)
            .transpose()?,
        margin_top: take(&mut a, Ns::F, b"margin-top")
            .map(TableStyleMeasure::new)
            .transpose()?,
        margin_bottom: take(&mut a, Ns::F, b"margin-bottom")
            .map(TableStyleMeasure::new)
            .transpose()?,
        margin: take(&mut a, Ns::F, b"margin")
            .map(TableStyleMeasure::new)
            .transpose()?,
        page_number: take(&mut a, Ns::S, b"page-number")
            .map(|x| TablePageNumber::parse(&x))
            .transpose()?,
        break_before: take(&mut a, Ns::F, b"break-before")
            .map(|x| parse_break(&x))
            .transpose()?,
        break_after: take(&mut a, Ns::F, b"break-after")
            .map(|x| parse_break(&x))
            .transpose()?,
        background_color: take(&mut a, Ns::F, b"background-color")
            .map(TableRowBackgroundColor::new)
            .transpose()?,
        background_image: None,
        shadow: take(&mut a, Ns::S, b"shadow")
            .map(TableShadow::new)
            .transpose()?,
        keep_with_next: take(&mut a, Ns::F, b"keep-with-next")
            .map(|x| parse_keep(&x))
            .transpose()?,
        may_break_between_rows: take(&mut a, Ns::S, b"may-break-between-rows")
            .map(|x| boolean(&x))
            .transpose()?,
        border_model: take(&mut a, Ns::T, b"border-model")
            .map(|x| TableBorderModel::parse(&x))
            .transpose()?,
        writing_mode: take(&mut a, Ns::S, b"writing-mode")
            .map(|x| TableWritingMode::parse(&x))
            .transpose()?,
        display: take(&mut a, Ns::T, b"display")
            .map(|x| boolean(&x))
            .transpose()?,
    };
    if !a.is_empty() {
        return Err(bad("unknown style:table-properties attribute"));
    }
    p.validate()?;
    Ok(p)
}
fn position(x: &str) -> Result<TableRowBackgroundPosition> {
    let w: Vec<_> = x.split_ascii_whitespace().collect();
    let h = |x| match x {
        "left" => Some(HorizontalBackgroundPosition::Left),
        "center" => Some(HorizontalBackgroundPosition::Center),
        "right" => Some(HorizontalBackgroundPosition::Right),
        _ => None,
    };
    let v = |x| match x {
        "top" => Some(VerticalBackgroundPosition::Top),
        "center" => Some(VerticalBackgroundPosition::Center),
        "bottom" => Some(VerticalBackgroundPosition::Bottom),
        _ => None,
    };
    match w.as_slice() {
        ["left"] => Ok(TableRowBackgroundPosition::Left),
        ["center"] => Ok(TableRowBackgroundPosition::Center),
        ["right"] => Ok(TableRowBackgroundPosition::Right),
        ["top"] => Ok(TableRowBackgroundPosition::Top),
        ["bottom"] => Ok(TableRowBackgroundPosition::Bottom),
        [a, b] => h(a)
            .zip(v(b))
            .or_else(|| h(b).zip(v(a)))
            .map(|(h, v)| TableRowBackgroundPosition::Pair(h, v))
            .ok_or_else(|| bad("invalid background position")),
        _ => Err(bad("invalid background position")),
    }
}
struct Image {
    value: TableRowBackgroundImage,
    linked: bool,
}
fn image(r: &NsReader<&[u8]>, v: XmlVersion, e: &BytesStart<'_>) -> Result<Image> {
    let mut a = attrs(r, v, e)?;
    let repeat = take(&mut a, Ns::S, b"repeat")
        .map(|x| match x.as_str() {
            "no-repeat" => Ok(TableRowBackgroundRepeat::NoRepeat),
            "repeat" => Ok(TableRowBackgroundRepeat::Repeat),
            "stretch" => Ok(TableRowBackgroundRepeat::Stretch),
            _ => Err(bad("invalid background repeat")),
        })
        .transpose()?;
    let pos = take(&mut a, Ns::S, b"position")
        .map(|x| position(&x))
        .transpose()?;
    let filter = take(&mut a, Ns::S, b"filter-name");
    if let Some(x) = &filter {
        safe(x, "style:filter-name", true)?
    }
    let opacity = take(&mut a, Ns::D, b"opacity")
        .map(TableRowOpacity::new)
        .transpose()?;
    let kind = take(&mut a, Ns::X, b"type");
    let href = take(&mut a, Ns::X, b"href");
    let show = take(&mut a, Ns::X, b"show");
    let actuate = take(&mut a, Ns::X, b"actuate");
    if !a.is_empty() {
        return Err(bad("unknown background-image attribute"));
    }
    let linked = kind.is_some() || href.is_some() || show.is_some() || actuate.is_some();
    let source = if linked {
        if kind.as_deref() != Some("simple")
            || href.is_none()
            || show.as_deref().is_some_and(|x| x != "embed")
            || actuate.as_deref().is_some_and(|x| x != "onLoad")
        {
            return Err(bad("invalid background-image link group"));
        }
        TableRowBackgroundSource::Link {
            href: href.unwrap(),
            show_embed: show.is_some(),
            actuate_on_load: actuate.is_some(),
        }
    } else {
        TableRowBackgroundSource::Empty
    };
    let value = TableRowBackgroundImage {
        repeat,
        position: pos,
        filter_name: filter,
        opacity,
        source,
    };
    value.validate()?;
    Ok(Image { value, linked })
}
struct Active {
    depth: usize,
    style: TableStyleProperties,
    seen: bool,
    pd: Option<usize>,
    id: Option<usize>,
    bd: Option<usize>,
    binary: String,
    linked: bool,
}
fn push(
    out: &mut Vec<TableStyleProperties>,
    x: TableStyleProperties,
    total: &mut usize,
) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out
            .iter()
            .any(|o| o.name == x.name && o.is_default_style == x.is_default_style)
    {
        return Err(bad("duplicate or excessive table style"));
    }
    *total += x.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("table style data too large"));
    }
    out.push(x);
    Ok(())
}
/// Parse direct table styles from regular and automatic style containers.
pub fn parse_table_style_properties(xml: &str) -> Result<TableStylePropertiesSet> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML too large"));
    }
    let mut r = NsReader::from_reader(xml.as_bytes());
    r.config_mut().trim_text(false);
    let mut ver = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut out = Vec::new();
    let mut total = 0;
    loop {
        match r.read_event() {
            Ok(Event::Start(e)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML too deep"));
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
                    if let Some(style) = header(&r, ver, &e, c.1 == b"default-style")? {
                        active = Some(Active {
                            depth: d,
                            style,
                            seen: false,
                            pd: None,
                            id: None,
                            bd: None,
                            binary: String::new(),
                            linked: false,
                        })
                    }
                    continue;
                }
                if let Some(s) = active.as_mut() {
                    if d == s.depth + 1 && c.0 == Ns::S && c.1 == b"table-properties" {
                        if s.seen {
                            return Err(bad("duplicate style:table-properties"));
                        }
                        s.seen = true;
                        s.style.properties = Some(properties(&r, ver, &e)?);
                        s.pd = Some(d)
                    } else if c.1 == b"table-properties" {
                        return Err(bad(
                            "style:table-properties has invalid namespace or parent",
                        ));
                    } else if s.pd.is_some_and(|p| d == p + 1)
                        && c.0 == Ns::S
                        && c.1 == b"background-image"
                    {
                        if s.style
                            .properties
                            .as_ref()
                            .unwrap()
                            .background_image
                            .is_some()
                        {
                            return Err(bad("duplicate style:background-image"));
                        }
                        let i = image(&r, ver, &e)?;
                        s.style.properties.as_mut().unwrap().background_image = Some(i.value);
                        s.linked = i.linked;
                        s.id = Some(d)
                    } else if c.1 == b"background-image" {
                        return Err(bad("background-image has invalid namespace or parent"));
                    } else if s.id.is_some_and(|i| d == i + 1)
                        && c.0 == Ns::O
                        && c.1 == b"binary-data"
                    {
                        if s.linked || s.bd.is_some() {
                            return Err(bad("invalid office:binary-data"));
                        }
                        s.bd = Some(d);
                        s.binary.clear()
                    } else if s.pd.is_some_and(|p| d > p) {
                        return Err(bad("unexpected table-properties child"));
                    }
                }
            },
            Ok(Event::Empty(e)) => {
                let c = elem(&r, e.name());
                let parent = stack.last();
                let d = stack.len() + 1;
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::O && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && c.0 == Ns::S
                    && matches!(c.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(x) = header(&r, ver, &e, c.1 == b"default-style")? {
                        push(&mut out, x, &mut total)?
                    }
                    continue;
                }
                if let Some(s) = active.as_mut() {
                    if d == s.depth + 1 && c.0 == Ns::S && c.1 == b"table-properties" {
                        if s.seen {
                            return Err(bad("duplicate style:table-properties"));
                        }
                        s.seen = true;
                        s.style.properties = Some(properties(&r, ver, &e)?)
                    } else if c.1 == b"table-properties" {
                        return Err(bad(
                            "style:table-properties has invalid namespace or parent",
                        ));
                    } else if s.pd.is_some_and(|p| d == p + 1)
                        && c.0 == Ns::S
                        && c.1 == b"background-image"
                    {
                        if s.style
                            .properties
                            .as_ref()
                            .unwrap()
                            .background_image
                            .is_some()
                        {
                            return Err(bad("duplicate style:background-image"));
                        }
                        s.style.properties.as_mut().unwrap().background_image =
                            Some(image(&r, ver, &e)?.value)
                    } else if c.1 == b"background-image" {
                        return Err(bad("background-image has invalid namespace or parent"));
                    } else if s.id.is_some_and(|i| d == i + 1)
                        && c.0 == Ns::O
                        && c.1 == b"binary-data"
                    {
                        if s.linked {
                            return Err(bad("linked image cannot contain binary data"));
                        }
                        s.style
                            .properties
                            .as_mut()
                            .unwrap()
                            .background_image
                            .as_mut()
                            .unwrap()
                            .source = TableRowBackgroundSource::Embedded(Vec::new())
                    } else if s.pd.is_some_and(|p| d > p) {
                        return Err(bad("unexpected table-properties child"));
                    }
                }
            },
            Ok(Event::Text(t)) => {
                let bytes: &[u8] = t.as_ref();
                if let Some(s) = active.as_mut() {
                    if s.bd.is_some() {
                        if s.binary.len() + bytes.len() > MAX_BINARY * 2 {
                            return Err(bad("encoded binary data too large"));
                        }
                        s.binary.push_str(&String::from_utf8_lossy(bytes))
                    } else if s.pd.is_some() && !bytes.iter().all(u8::is_ascii_whitespace) {
                        return Err(bad("unexpected text in table-properties"));
                    }
                }
            },
            Ok(Event::CData(t)) => {
                let bytes: &[u8] = t.as_ref();
                if let Some(s) = active.as_mut() {
                    if s.bd.is_some() {
                        if s.binary.len() + bytes.len() > MAX_BINARY * 2 {
                            return Err(bad("encoded binary data too large"));
                        }
                        s.binary.push_str(&String::from_utf8_lossy(bytes))
                    } else if s.pd.is_some() && !bytes.iter().all(u8::is_ascii_whitespace) {
                        return Err(bad("unexpected text in table-properties"));
                    }
                }
            },
            Ok(Event::End(_)) => {
                let d = stack.len();
                if let Some(s) = active.as_mut() {
                    if s.bd == Some(d) {
                        let data = b64_decode(&s.binary)?;
                        s.style
                            .properties
                            .as_mut()
                            .unwrap()
                            .background_image
                            .as_mut()
                            .unwrap()
                            .source = TableRowBackgroundSource::Embedded(data);
                        s.bd = None
                    }
                    if s.id == Some(d) {
                        s.id = None
                    }
                    if s.pd == Some(d) {
                        s.pd = None
                    }
                }
                if active.as_ref().is_some_and(|x| x.depth == d) {
                    push(&mut out, active.take().unwrap().style, &mut total)?
                }
                stack.pop();
            },
            Ok(Event::Decl(d)) => {
                ver = d
                    .xml_version()
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid styles XML: {e}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(TableStylePropertiesSet { styles: out })
}
fn b64_decode(x: &str) -> Result<Vec<u8>> {
    let x: Vec<_> = x.bytes().filter(|x| !x.is_ascii_whitespace()).collect();
    if x.len() % 4 != 0 {
        return Err(bad("invalid binary-data base64"));
    }
    let val = |x| match x {
        b'A'..=b'Z' => Some(x - b'A'),
        b'a'..=b'z' => Some(x - b'a' + 26),
        b'0'..=b'9' => Some(x - b'0' + 52),
        b'+' => Some(62),
        b'/' => Some(63),
        _ => None,
    };
    let mut out = Vec::with_capacity(x.len() / 4 * 3);
    let groups = x.len() / 4;
    for (i, c) in x.chunks(4).enumerate() {
        let pad = (c[2] == b'=') as usize + (c[3] == b'=') as usize;
        if (i + 1 < groups && pad != 0) || (c[2] == b'=' && c[3] != b'=') || pad > 2 {
            return Err(bad("invalid binary-data padding"));
        }
        let a = val(c[0]).ok_or_else(|| bad("invalid base64"))? as u32;
        let b = val(c[1]).ok_or_else(|| bad("invalid base64"))? as u32;
        let z = if c[2] == b'=' {
            0
        } else {
            val(c[2]).ok_or_else(|| bad("invalid base64"))? as u32
        };
        let q = if c[3] == b'=' {
            0
        } else {
            val(c[3]).ok_or_else(|| bad("invalid base64"))? as u32
        };
        let n = a << 18 | b << 12 | z << 6 | q;
        out.push((n >> 16) as u8);
        if pad < 2 {
            out.push((n >> 8) as u8)
        }
        if pad < 1 {
            out.push(n as u8)
        }
        if out.len() > MAX_BINARY {
            return Err(bad("binary data too large"));
        }
    }
    Ok(out)
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct Spans {
    style: Span,
    properties: Option<Span>,
}
fn boundary(x: &str, end: usize) -> Result<usize> {
    x[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}
fn replace(x: &str, s: &Span, v: &str) -> String {
    format!("{}{}{}", &x[..s.start], v, &x[s.end..])
}
fn expand(x: &str, s: &Span, v: &str) -> Result<String> {
    let raw = &x[s.start..s.end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| bad("invalid empty element"))?;
    Ok(replace(
        x,
        s,
        &format!("{}>{v}</{}>", &raw[..slash], s.qname),
    ))
}
/// Losslessly replace, insert, or remove one existing table style's property element.
pub fn set_table_style_properties_xml(xml: &str, want: &TableStyleProperties) -> Result<String> {
    want.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML too large"));
    }
    let mut r = NsReader::from_reader(xml.as_bytes());
    let mut ver = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut td = None;
    let mut active: Option<Spans> = None;
    let mut found = None;
    loop {
        match r.read_event() {
            Ok(Event::Start(e)) => {
                let end = r.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let c = elem(&r, e.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::O && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && c.0 == Ns::S
                    && matches!(c.1.as_slice(), b"style" | b"default-style");
                stack.push(c.clone());
                let d = stack.len();
                if direct {
                    if let Some(x) = header(&r, ver, &e, c.1 == b"default-style")?
                        && x.name == want.name
                        && x.is_default_style == want.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target table style"));
                        }
                        td = Some(d);
                        active = Some(Spans {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        })
                    }
                } else if td.is_some_and(|x| d == x + 1)
                    && c.0 == Ns::S
                    && c.1 == b"table-properties"
                {
                    let s = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(s).is_some() {
                        return Err(bad("duplicate style:table-properties"));
                    }
                }
            },
            Ok(Event::Empty(e)) => {
                let end = r.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let c = elem(&r, e.name());
                let parent = stack.last();
                let d = stack.len() + 1;
                let direct = parent.is_some_and(|x| {
                    x.0 == Ns::O && matches!(x.1.as_slice(), b"styles" | b"automatic-styles")
                }) && c.0 == Ns::S
                    && matches!(c.1.as_slice(), b"style" | b"default-style");
                let s = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(x) = header(&r, ver, &e, c.1 == b"default-style")?
                        && x.name == want.name
                        && x.is_default_style == want.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target table style"));
                        }
                        found = Some(Spans {
                            style: s,
                            ..Default::default()
                        })
                    }
                } else if td.is_some_and(|x| d == x + 1)
                    && c.0 == Ns::S
                    && c.1 == b"table-properties"
                    && active.as_mut().unwrap().properties.replace(s).is_some()
                {
                    return Err(bad("duplicate style:table-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = r.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let d = stack.len();
                if let Some(s) = active.as_mut() {
                    if s.properties.as_ref().is_some_and(|x| x.end == 0)
                        && td.is_some_and(|x| d == x + 1)
                    {
                        let x = s.properties.as_mut().unwrap();
                        x.end_start = begin;
                        x.end = end
                    }
                    if td == Some(d) {
                        s.style.end_start = begin;
                        s.style.end = end;
                        found = active.take();
                        td = None
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(d)) => {
                ver = d
                    .xml_version()
                    .map_err(|e| bad(format!("unsupported XML version: {e}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(e) => return Err(bad(format!("invalid styles XML: {e}"))),
        }
    }
    let s = found.ok_or_else(|| bad("target table style does not exist"))?;
    let value = want
        .properties
        .as_ref()
        .map(TableProperties::to_xml_fragment)
        .transpose()?;
    if let Some(p) = &s.properties {
        return Ok(replace(xml, p, value.as_deref().unwrap_or("")));
    }
    let Some(value) = value else {
        return Ok(xml.to_owned());
    };
    if s.style.empty {
        return expand(xml, &s.style, &value);
    }
    let mut out = xml.to_owned();
    out.insert_str(s.style.end_start, &value);
    Ok(out)
}
impl OpenDocumentPackage {
    pub fn table_style_properties(&self) -> Result<TableStylePropertiesSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |x| parse_table_style_properties(&x),
        )
    }
}
impl FlatOpenDocument {
    pub fn table_style_properties(&self) -> Result<TableStylePropertiesSet> {
        parse_table_style_properties(self.xml())
    }
}
