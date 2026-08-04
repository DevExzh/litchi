//! Semantic WordprocessingML font-table model and inert embedding values.

use crate::{Error, Result};
use caseless::Caseless;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;
use std::sync::Arc;
use unicode_normalization::UnicodeNormalization;

pub(in crate::font) const WT: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(in crate::font) const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
pub(in crate::font) const RT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(in crate::font) const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(in crate::font) const FT_RT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
pub(in crate::font) const FT_RS: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable";
pub(in crate::font) const FONT_RT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
pub(in crate::font) const FONT_RS: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/font";
pub(in crate::font) const FT_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
pub(in crate::font) const FONT_CT: &str =
    "application/vnd.openxmlformats-officedocument.obfuscatedFont";
pub(in crate::font) const XMLNS: &str = "http://www.w3.org/2000/xmlns/";
pub(in crate::font) const MAX_XML: usize = 8 * 1024 * 1024;
pub(in crate::font) const MAX_FONT: usize = 32 * 1024 * 1024;
pub(in crate::font) const MAX_ALL_FONTS: usize = 128 * 1024 * 1024;
pub(in crate::font) const MAX_FONTS: usize = 4096;
pub(in crate::font) const MAX_NODES: usize = 64_000;
pub(in crate::font) const MAX_DEPTH: usize = 32;
pub(in crate::font) const MAX_TEXT: usize = 32_768;

pub(in crate::font) fn is_font_table_relationship(value: &str) -> bool {
    matches!(value, FT_RT | FT_RS)
}

pub(in crate::font) fn is_font_relationship(value: &str) -> bool {
    matches!(value, FONT_RT | FONT_RS)
}

pub(in crate::font) fn word_ns(value: &str) -> bool {
    matches!(value, WT | WS)
}

pub(in crate::font) fn rel_ns(value: &str) -> bool {
    matches!(value, RT | RS)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}
impl Conformance {
    pub(in crate::font) fn word(self) -> &'static str {
        match self {
            Self::Transitional => WT,
            Self::Strict => WS,
        }
    }
    pub(in crate::font) fn rel(self) -> &'static str {
        match self {
            Self::Transitional => RT,
            Self::Strict => RS,
        }
    }
    pub(in crate::font) fn font_table_rel(self) -> &'static str {
        match self {
            Self::Transitional => FT_RT,
            Self::Strict => FT_RS,
        }
    }
    pub(in crate::font) fn font_rel(self) -> &'static str {
        match self {
            Self::Transitional => FONT_RT,
            Self::Strict => FONT_RS,
        }
    }
}

/// Validated OpenType `OS/2.fsType` embedding metadata supplied by a caller.
///
/// This module never searches, parses, loads, renders, or executes a font
/// program. Licensing bits therefore remain inert metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(transparent)]
pub struct License(u16);

impl License {
    /// Validate and retain the compact OpenType `fsType` bit field.
    pub fn new(fs_type: u16) -> Result<Self> {
        const DEFINED: u16 = 0x0002 | 0x0004 | 0x0008 | 0x0100 | 0x0200;
        if fs_type & !DEFINED != 0 {
            return Err(invalid(format!(
                "font fsType contains reserved bits 0x{:04X}",
                fs_type & !DEFINED
            )));
        }
        if [0x0002, 0x0004, 0x0008]
            .into_iter()
            .filter(|bit| fs_type & *bit != 0)
            .count()
            > 1
        {
            return Err(invalid(
                "font fsType has contradictory restricted, preview/print, and editable modes",
            ));
        }
        Ok(Self(fs_type))
    }

    /// Return the validated OpenType bit field.
    pub const fn bits(self) -> u16 {
        self.0
    }

    pub const fn restricted(self) -> bool {
        self.0 & 0x0002 != 0
    }

    pub const fn preview_print(self) -> bool {
        self.0 & 0x0004 != 0
    }

    pub const fn editable(self) -> bool {
        self.0 & 0x0008 != 0
    }

    pub const fn no_subsetting(self) -> bool {
        self.0 & 0x0100 != 0
    }

    pub const fn bitmap_only(self) -> bool {
        self.0 & 0x0200 != 0
    }

    pub const fn installable(self) -> bool {
        self.0 & 0x000E == 0
    }
}

/// Compact OOXML font-obfuscation GUID.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct FontKey([u8; 16]);

impl FontKey {
    /// Construct a key from its binary GUID representation.
    pub const fn new(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Borrow the binary GUID representation.
    pub const fn bytes(&self) -> &[u8; 16] {
        &self.0
    }

    /// Move out the binary GUID representation.
    pub const fn into_bytes(self) -> [u8; 16] {
        self.0
    }
}

impl From<[u8; 16]> for FontKey {
    fn from(bytes: [u8; 16]) -> Self {
        Self::new(bytes)
    }
}

impl FromStr for FontKey {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let bytes = value.as_bytes();
        let valid = bytes.len() == 38
            && bytes.first() == Some(&b'{')
            && bytes.last() == Some(&b'}')
            && [9, 14, 19, 24]
                .iter()
                .all(|index| bytes.get(*index) == Some(&b'-'))
            && bytes.get(1..37).is_some_and(|body| {
                body.iter().enumerate().all(|(index, byte)| {
                    [8, 13, 18, 23].contains(&index)
                        || byte.is_ascii_digit()
                        || (b'A'..=b'F').contains(byte)
                })
            });
        if !valid {
            return Err(invalid(format!("invalid font key '{value}'")));
        }

        let mut key = [0u8; 16];
        let mut digits = bytes
            .get(1..37)
            .ok_or_else(|| invalid("font key body is missing"))?
            .iter()
            .copied()
            .filter(|byte| *byte != b'-');
        for output in &mut key {
            let high = digits
                .next()
                .and_then(hex_digit)
                .ok_or_else(|| invalid("font key has an invalid high nibble"))?;
            let low = digits
                .next()
                .and_then(hex_digit)
                .ok_or_else(|| invalid("font key has an invalid low nibble"))?;
            *output = (high << 4) | low;
        }
        Ok(Self(key))
    }
}

impl fmt::Display for FontKey {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("{")?;
        for (index, byte) in self.0.iter().enumerate() {
            if matches!(index, 4 | 6 | 8 | 10) {
                formatter.write_str("-")?;
            }
            write!(formatter, "{byte:02X}")?;
        }
        formatter.write_str("}")
    }
}

const fn hex_digit(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

/// Apply the reversible OOXML GUID XOR transformation to the first 32 bytes.
pub fn obfuscate(data: &mut [u8], font_key: FontKey) -> Result<()> {
    if data.len() < 32 {
        return Err(invalid("OOXML font obfuscation requires at least 32 bytes"));
    }
    for (byte, key_byte) in data
        .iter_mut()
        .take(32)
        .zip(font_key.bytes().iter().rev().cycle())
    {
        *byte ^= *key_byte;
    }
    Ok(())
}

/// Reverse [`obfuscate`]. XOR makes both operations identical.
pub fn deobfuscate(data: &mut [u8], font_key: FontKey) -> Result<()> {
    obfuscate(data, font_key)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Family {
    Decorative,
    Modern,
    Roman,
    Script,
    Swiss,
    Auto,
}
impl Family {
    pub(in crate::font) fn parse(v: &str) -> Result<Self> {
        match v {
            "decorative" => Ok(Self::Decorative),
            "modern" => Ok(Self::Modern),
            "roman" => Ok(Self::Roman),
            "script" => Ok(Self::Script),
            "swiss" => Ok(Self::Swiss),
            "auto" => Ok(Self::Auto),
            _ => Err(invalid(format!("invalid font family '{v}'"))),
        }
    }
    pub(in crate::font) fn text(self) -> &'static str {
        match self {
            Self::Decorative => "decorative",
            Self::Modern => "modern",
            Self::Roman => "roman",
            Self::Script => "script",
            Self::Swiss => "swiss",
            Self::Auto => "auto",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Pitch {
    Fixed,
    Variable,
    Default,
}
impl Pitch {
    pub(in crate::font) fn parse(v: &str) -> Result<Self> {
        match v {
            "fixed" => Ok(Self::Fixed),
            "variable" => Ok(Self::Variable),
            "default" => Ok(Self::Default),
            _ => Err(invalid(format!("invalid font pitch '{v}'"))),
        }
    }
    pub(in crate::font) fn text(self) -> &'static str {
        match self {
            Self::Fixed => "fixed",
            Self::Variable => "variable",
            Self::Default => "default",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Charset {
    Ansi,
    Default,
    Symbol,
    Macintosh,
    ShiftJis,
    Hangeul,
    Johab,
    Gb2312,
    ChineseBig5,
    Greek,
    Turkish,
    Vietnamese,
    Hebrew,
    Arabic,
    Baltic,
    Russian,
    Thai,
    EastEurope,
    Oem,
    Legacy(u8),
}
impl Charset {
    /// Convert a legacy one-byte Word charset code into a typed value.
    pub fn from_legacy(v: u8) -> Self {
        match v {
            0x00 => Self::Ansi,
            0x01 => Self::Default,
            0x02 => Self::Symbol,
            0x4D => Self::Macintosh,
            0x80 => Self::ShiftJis,
            0x81 => Self::Hangeul,
            0x82 => Self::Johab,
            0x86 => Self::Gb2312,
            0x88 => Self::ChineseBig5,
            0xA1 => Self::Greek,
            0xA2 => Self::Turkish,
            0xA3 => Self::Vietnamese,
            0xB1 => Self::Hebrew,
            0xB2 => Self::Arabic,
            0xBA => Self::Baltic,
            0xCC => Self::Russian,
            0xDE => Self::Thai,
            0xEE => Self::EastEurope,
            0xFF => Self::Oem,
            x => Self::Legacy(x),
        }
    }
    pub(in crate::font) fn strict(v: &str) -> Result<Self> {
        match v {
            "iso-8859-1" => Ok(Self::Ansi),
            "macintosh" => Ok(Self::Macintosh),
            "shift_jis" => Ok(Self::ShiftJis),
            "ks_c-5601-1987" => Ok(Self::Hangeul),
            "KS_C-5601-1992" => Ok(Self::Johab),
            "GBK" => Ok(Self::Gb2312),
            "Big5" => Ok(Self::ChineseBig5),
            "windows-1253" => Ok(Self::Greek),
            "iso-8859-9" => Ok(Self::Turkish),
            "windows-1258" => Ok(Self::Vietnamese),
            "windows-1255" => Ok(Self::Hebrew),
            "windows-1256" => Ok(Self::Arabic),
            "windows-1257" => Ok(Self::Baltic),
            "windows-1251" => Ok(Self::Russian),
            "windows-874" => Ok(Self::Thai),
            "windows-1250" => Ok(Self::EastEurope),
            _ => Err(invalid(format!("invalid strict character set '{v}'"))),
        }
    }
    pub fn legacy_code(self) -> u8 {
        match self {
            Self::Ansi => 0x00,
            Self::Default => 0x01,
            Self::Symbol => 0x02,
            Self::Macintosh => 0x4D,
            Self::ShiftJis => 0x80,
            Self::Hangeul => 0x81,
            Self::Johab => 0x82,
            Self::Gb2312 => 0x86,
            Self::ChineseBig5 => 0x88,
            Self::Greek => 0xA1,
            Self::Turkish => 0xA2,
            Self::Vietnamese => 0xA3,
            Self::Hebrew => 0xB1,
            Self::Arabic => 0xB2,
            Self::Baltic => 0xBA,
            Self::Russian => 0xCC,
            Self::Thai => 0xDE,
            Self::EastEurope => 0xEE,
            Self::Oem => 0xFF,
            Self::Legacy(v) => v,
        }
    }
    pub fn strict_name(self) -> Option<&'static str> {
        Some(match self {
            Self::Ansi => "iso-8859-1",
            Self::Default | Self::Symbol | Self::Oem => return None,
            Self::Macintosh => "macintosh",
            Self::ShiftJis => "shift_jis",
            Self::Hangeul => "ks_c-5601-1987",
            Self::Johab => "KS_C-5601-1992",
            Self::Gb2312 => "GBK",
            Self::ChineseBig5 => "Big5",
            Self::Greek => "windows-1253",
            Self::Turkish => "iso-8859-9",
            Self::Vietnamese => "windows-1258",
            Self::Hebrew => "windows-1255",
            Self::Arabic => "windows-1256",
            Self::Baltic => "windows-1257",
            Self::Russian => "windows-1251",
            Self::Thai => "windows-874",
            Self::EastEurope => "windows-1250",
            Self::Legacy(_) => return None,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Style {
    Regular,
    Bold,
    Italic,
    BoldItalic,
}
impl Style {
    pub(in crate::font) fn element(self) -> &'static str {
        match self {
            Self::Regular => "embedRegular",
            Self::Bold => "embedBold",
            Self::Italic => "embedItalic",
            Self::BoldItalic => "embedBoldItalic",
        }
    }
    pub(in crate::font) fn rank(self) -> u8 {
        match self {
            Self::Regular => 0,
            Self::Bold => 1,
            Self::Italic => 2,
            Self::BoldItalic => 3,
        }
    }
}

// Private implementation spellings keep the package codec compact without
// leaking the retired monolith vocabulary into rustdoc or downstream code.

/// Low-level, lossless XML details that normal font authoring can ignore.
pub mod raw {
    use super::{bounded, validate_attr_name};
    use crate::Result;

    /// An extension attribute preserved verbatim by the XML codec.
    #[derive(Debug, Clone, PartialEq, Eq)]
    pub struct Attr {
        pub(in crate::font) qualified_name: String,
        pub(in crate::font) value: String,
    }

    impl Attr {
        pub fn new(qualified_name: impl Into<String>, value: impl Into<String>) -> Result<Self> {
            let qualified_name = qualified_name.into();
            let value = value.into();
            validate_attr_name(&qualified_name)?;
            bounded(&value)?;
            Ok(Self {
                qualified_name,
                value,
            })
        }

        pub fn qualified_name(&self) -> &str {
            &self.qualified_name
        }

        pub fn value(&self) -> &str {
            &self.value
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Signature {
    pub(in crate::font) unicode_subsets: [u32; 4],
    pub(in crate::font) code_pages: [u32; 2],
}
impl Signature {
    pub fn new(unicode_subsets: [u32; 4], code_pages: [u32; 2]) -> Self {
        Self {
            unicode_subsets,
            code_pages,
        }
    }
    pub fn unicode_subsets(&self) -> &[u32; 4] {
        &self.unicode_subsets
    }
    pub fn code_pages(&self) -> &[u32; 2] {
        &self.code_pages
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Resource {
    // OPC identity is deliberately private: safe authoring allocates it.
    pub(in crate::font) part_name: String,
    pub(in crate::font) content_type: String,
    pub(in crate::font) data: Arc<Vec<u8>>,
}
impl Resource {
    /// Own an inert, already-obfuscated embedded-font payload.
    pub fn new(data: Vec<u8>) -> Result<Self> {
        Self::from_shared(Arc::new(data))
    }

    /// Share an inert payload without copying its allocation.
    pub(in crate::font) fn from_shared(data: Arc<Vec<u8>>) -> Result<Self> {
        validate_resource_len(data.len())?;
        Ok(Self {
            part_name: String::new(),
            content_type: FONT_CT.into(),
            data,
        })
    }

    /// Borrow the payload bytes.
    pub fn bytes(&self) -> &[u8] {
        self.data.as_slice()
    }

    /// Share the payload allocation with another owner.
    pub(in crate::font) fn share(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.data)
    }

    /// Whether two resources retain the same package allocation.
    pub fn shares_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.data, &other.data)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Embed {
    pub(in crate::font) style: Style,
    pub(in crate::font) relationship_id: String,
    pub(in crate::font) font_key: Option<FontKey>,
    pub(in crate::font) subsetted: Option<bool>,
    pub(in crate::font) resource: Option<Resource>,
    pub(in crate::font) extension_attributes: Vec<raw::Attr>,
}
impl Embed {
    pub fn new(style: Style, font_key: FontKey, resource: Resource) -> Self {
        Self {
            style,
            relationship_id: String::new(),
            font_key: Some(font_key),
            subsetted: None,
            resource: Some(resource),
            extension_attributes: Vec::new(),
        }
    }

    /// Replace the OOXML GUID used to obfuscate this payload.
    pub fn rekey(&mut self, font_key: FontKey) -> Option<FontKey> {
        self.font_key.replace(font_key)
    }

    pub fn with_subset(mut self, subsetted: bool) -> Self {
        self.subsetted = Some(subsetted);
        self
    }

    pub fn style(&self) -> Style {
        self.style
    }

    pub fn key(&self) -> Option<FontKey> {
        self.font_key
    }

    pub fn subsetted(&self) -> Option<bool> {
        self.subsetted
    }

    pub fn resource(&self) -> Option<&Resource> {
        self.resource.as_ref()
    }

    pub fn attrs(&self) -> &[raw::Attr] {
        &self.extension_attributes
    }

    pub fn with_attr(mut self, attr: raw::Attr) -> Self {
        self.extension_attributes.push(attr);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    pub(in crate::font) name: String,
    pub(in crate::font) alternate_name: Option<String>,
    pub(in crate::font) panose: Option<[u8; 10]>,
    pub(in crate::font) character_set: Option<Charset>,
    pub(in crate::font) family: Option<Family>,
    pub(in crate::font) not_true_type: Option<bool>,
    pub(in crate::font) pitch: Option<Pitch>,
    pub(in crate::font) signature: Option<Signature>,
    pub(in crate::font) embedded_fonts: Vec<Embed>,
    pub(in crate::font) extension_attributes: Vec<raw::Attr>,
}
impl Font {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_font_name(&name, "font name")?;
        Ok(Self {
            name,
            alternate_name: None,
            panose: None,
            character_set: None,
            family: None,
            not_true_type: None,
            pitch: None,
            signature: None,
            embedded_fonts: Vec::new(),
            extension_attributes: Vec::new(),
        })
    }

    pub fn with_alt(mut self, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_font_name(&value, "alternate font name")?;
        self.alternate_name = Some(value);
        Ok(self)
    }

    pub fn with_panose(mut self, value: [u8; 10]) -> Self {
        self.panose = Some(value);
        self
    }
    pub fn with_charset(mut self, value: Charset) -> Self {
        self.character_set = Some(value);
        self
    }
    /// Replace or clear the character-set hint.
    pub fn set_charset(&mut self, value: Option<Charset>) -> Option<Charset> {
        std::mem::replace(&mut self.character_set, value)
    }
    pub fn with_family(mut self, value: Family) -> Self {
        self.family = Some(value);
        self
    }
    pub fn with_not_true_type(mut self, value: bool) -> Self {
        self.not_true_type = Some(value);
        self
    }
    pub fn with_pitch(mut self, value: Pitch) -> Self {
        self.pitch = Some(value);
        self
    }
    pub fn with_signature(mut self, value: Signature) -> Self {
        self.signature = Some(value);
        self
    }
    pub fn with_embed(mut self, value: Embed) -> Result<Self> {
        self.add_embed(value)?;
        Ok(self)
    }

    pub fn rename(&mut self, value: impl Into<String>) -> Result<()> {
        let value = value.into();
        validate_font_name(&value, "font name")?;
        self.name = value;
        Ok(())
    }

    /// Add one embedded face, preserving schema order and rejecting duplicates.
    pub fn add_embed(&mut self, value: Embed) -> Result<()> {
        if self
            .embedded_fonts
            .iter()
            .any(|existing| existing.style == value.style)
        {
            return Err(invalid("embedded-font style already exists"));
        }
        self.embedded_fonts.push(value);
        self.embedded_fonts.sort_by_key(|embed| embed.style.rank());
        Ok(())
    }

    /// Add or replace one embedded face by its typed style.
    pub fn put(&mut self, value: Embed) -> Result<Option<Embed>> {
        if let Some(index) = self
            .embedded_fonts
            .iter()
            .position(|embedded| embedded.style == value.style)
        {
            let len = self.embedded_fonts.len();
            let slot = self
                .embedded_fonts
                .get_mut(index)
                .ok_or_else(|| invalid(format!("embedded-font index {index} exceeds {len}")))?;
            return Ok(Some(std::mem::replace(slot, value)));
        }
        self.add_embed(value)?;
        Ok(None)
    }

    /// Remove and return one embedded face, if present.
    pub fn remove(&mut self, style: Style) -> Option<Embed> {
        self.embedded_fonts
            .iter()
            .position(|embedded| embedded.style == style)
            .map(|index| self.embedded_fonts.remove(index))
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn alt(&self) -> Option<&str> {
        self.alternate_name.as_deref()
    }
    pub fn panose(&self) -> Option<&[u8; 10]> {
        self.panose.as_ref()
    }
    pub fn charset(&self) -> Option<Charset> {
        self.character_set
    }
    pub fn family(&self) -> Option<Family> {
        self.family
    }
    pub fn not_true_type(&self) -> Option<bool> {
        self.not_true_type
    }
    pub fn pitch(&self) -> Option<Pitch> {
        self.pitch
    }
    pub fn signature(&self) -> Option<&Signature> {
        self.signature.as_ref()
    }
    pub fn embeds(&self) -> &[Embed] {
        &self.embedded_fonts
    }

    pub fn attrs(&self) -> &[raw::Attr] {
        &self.extension_attributes
    }

    pub fn with_attr(mut self, attr: raw::Attr) -> Self {
        self.extension_attributes.push(attr);
        self
    }
}

/// A safe table selector. Names are the primary stable selector; numeric
/// positions remain available for import and inspection workflows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Key<'a> {
    Name(&'a str),
    Index(usize),
}

impl<'a> From<&'a str> for Key<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for Key<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    pub(in crate::font) fonts: Vec<Font>,
    pub(in crate::font) namespaces: Vec<raw::Attr>,
    pub(in crate::font) extension_attributes: Vec<raw::Attr>,
}

impl Table {
    pub fn new() -> Self {
        Self {
            fonts: Vec::new(),
            namespaces: Vec::new(),
            extension_attributes: Vec::new(),
        }
    }
    pub fn fonts(&self) -> &[Font] {
        &self.fonts
    }

    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Font> {
        self.fonts.iter()
    }

    pub fn len(&self) -> usize {
        self.fonts.len()
    }

    pub fn is_empty(&self) -> bool {
        self.fonts.is_empty()
    }

    pub fn get<'a, 's>(&'a self, key: impl Into<Key<'s>>) -> Result<Option<&'a Font>> {
        match key.into() {
            Key::Name(name) => Ok(
                unique_font_offset(&self.fonts, name)?.and_then(|offset| self.fonts.get(offset))
            ),
            Key::Index(offset) => Ok(self.fonts.get(offset)),
        }
    }

    pub fn add(&mut self, font: Font) -> Result<()> {
        validate_font_entry(&font, false)?;
        if unique_font_offset(&self.fonts, font.name())?.is_some() {
            return Err(invalid(format!("font '{}' already exists", font.name())));
        }
        self.fonts.push(font);
        Ok(())
    }

    pub fn replace<'s>(
        &mut self,
        key: impl Into<Key<'s>>,
        replacement: Font,
    ) -> Result<Option<Font>> {
        validate_font_entry(&replacement, false)?;
        let offset = match key.into() {
            Key::Name(name) => unique_font_offset(&self.fonts, name)?,
            Key::Index(offset) => (offset < self.fonts.len()).then_some(offset),
        };
        let Some(offset) = offset else {
            return Ok(None);
        };
        let replacement_key = name_key(replacement.name());
        if self
            .fonts
            .iter()
            .enumerate()
            .any(|(index, font)| index != offset && name_key(&font.name) == replacement_key)
        {
            return Err(invalid(format!(
                "font '{}' already exists",
                replacement.name()
            )));
        }
        let slot = self
            .fonts
            .get_mut(offset)
            .ok_or_else(|| invalid("font selector changed during replacement"))?;
        Ok(Some(std::mem::replace(slot, replacement)))
    }

    pub fn remove<'s>(&mut self, key: impl Into<Key<'s>>) -> Result<Option<Font>> {
        let offset = match key.into() {
            Key::Name(name) => unique_font_offset(&self.fonts, name)?,
            Key::Index(offset) => (offset < self.fonts.len()).then_some(offset),
        };
        Ok(offset.map(|offset| self.fonts.remove(offset)))
    }

    pub fn reorder<S: AsRef<str>>(&mut self, ordered_names: &[S]) -> Result<()> {
        let mut rank = HashMap::with_capacity(ordered_names.len());
        for (offset, name) in ordered_names.iter().enumerate() {
            let name = name_key(name.as_ref());
            if rank.insert(name, offset).is_some() {
                return Err(invalid("font-table reorder contains duplicate names"));
            }
        }
        let expected = self
            .fonts
            .iter()
            .map(|font| name_key(&font.name))
            .collect::<HashSet<_>>();
        if rank.len() != self.fonts.len()
            || rank.keys().any(|name| !expected.contains(name))
            || expected.len() != self.fonts.len()
        {
            return Err(invalid("font-table reorder is not a font-name permutation"));
        }
        self.fonts.sort_by_cached_key(|font| {
            rank.get(&name_key(&font.name))
                .copied()
                .unwrap_or(usize::MAX)
        });
        Ok(())
    }

    pub fn attrs(&self) -> &[raw::Attr] {
        &self.extension_attributes
    }

    pub fn namespaces(&self) -> &[raw::Attr] {
        &self.namespaces
    }

    pub fn with_attr(mut self, attr: raw::Attr) -> Self {
        self.extension_attributes.push(attr);
        self
    }

    pub fn with_namespace(mut self, attr: raw::Attr) -> Result<Self> {
        if attr.qualified_name != "xmlns" && !attr.qualified_name.starts_with("xmlns:") {
            return Err(invalid(
                "namespace attribute must be named xmlns or xmlns:prefix",
            ));
        }
        self.namespaces.push(attr);
        Ok(self)
    }
}
impl Default for Table {
    fn default() -> Self {
        Self::new()
    }
}
pub(in crate::font) fn name_key(value: &str) -> String {
    value.chars().nfd().default_case_fold().nfd().collect()
}

pub(in crate::font) fn unique_font_offset(fonts: &[Font], name: &str) -> Result<Option<usize>> {
    let key = name_key(name);
    let mut matching = fonts
        .iter()
        .enumerate()
        .filter(|(_, font)| name_key(&font.name) == key)
        .map(|(offset, _)| offset);
    let first = matching.next();
    if first.is_some() && matching.next().is_some() {
        Err(invalid(format!(
            "font name '{name}' is ambiguous in the font table"
        )))
    } else {
        Ok(first)
    }
}

pub(in crate::font) fn validate_table_value(value: &Table, require_resources: bool) -> Result<()> {
    if value.fonts.len() > MAX_FONTS {
        return Err(invalid("too many fonts"));
    }
    let mut total = 0usize;
    let mut resource_names = HashSet::new();
    for font in &value.fonts {
        validate_font_entry(font, require_resources)?;
        for embedded in &font.embedded_fonts {
            if require_resources {
                let resource = embedded.resource.as_ref().ok_or_else(|| {
                    invalid("embedded-font resource is required for package storage")
                })?;
                if resource.content_type != FONT_CT {
                    return Err(invalid(format!(
                        "embedded font has invalid content type '{}'",
                        resource.content_type
                    )));
                }
                validate_resource_len(resource.data.len())?;
                if resource.part_name.is_empty() {
                    return Err(invalid("embedded-font part name is empty"));
                }
                if resource_names.insert(resource.part_name.clone()) {
                    total = total
                        .checked_add(resource.data.len())
                        .ok_or_else(|| invalid("embedded-font size overflow"))?;
                    if total > MAX_ALL_FONTS {
                        return Err(invalid("embedded fonts exceed total size limit"));
                    }
                }
            }
        }
    }
    Ok(())
}

pub(in crate::font) fn validate_font_entry(font: &Font, require_resources: bool) -> Result<()> {
    validate_font_name(&font.name, "font name")?;
    if let Some(name) = &font.alternate_name {
        validate_font_name(name, "alternate font name")?;
    }
    for pair in font.embedded_fonts.windows(2) {
        let [left, right] = pair else {
            return Err(invalid("invalid embedded-font ordering window"));
        };
        if left.style.rank() >= right.style.rank() {
            return Err(invalid(
                "embedded-font styles are duplicated or out of schema order",
            ));
        }
    }
    for embedded in &font.embedded_fonts {
        if require_resources
            && (embedded.relationship_id.is_empty() || embedded.relationship_id.len() > MAX_TEXT)
        {
            return Err(invalid(
                "embedded-font relationship ID is empty or too long",
            ));
        }
        if embedded.font_key.is_none() && require_resources {
            return Err(invalid("fontKey is required for package storage"));
        }
        if require_resources && embedded.resource.is_none() {
            return Err(invalid(
                "embedded-font resource is required for package storage",
            ));
        }
    }
    Ok(())
}

pub(in crate::font) fn validate_resource_len(len: usize) -> Result<()> {
    if (32..=MAX_FONT).contains(&len) {
        Ok(())
    } else {
        Err(invalid("embedded font size is outside the allowed bounds"))
    }
}

pub(in crate::font) fn validate_font_name(value: &str, kind: &str) -> Result<()> {
    let count = value.chars().count();
    if !(1..=31).contains(&count) {
        Err(invalid(format!(
            "{kind} must contain 1 through 31 characters"
        )))
    } else {
        Ok(())
    }
}

pub(in crate::font) fn bounded(v: &str) -> Result<()> {
    if v.len() <= MAX_TEXT {
        Ok(())
    } else {
        Err(invalid("font-table string limit exceeded"))
    }
}
pub(in crate::font) fn validate_attr_name(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_TEXT
        || value.bytes().any(|byte| {
            byte.is_ascii_whitespace() || matches!(byte, b'<' | b'>' | b'=' | b'\'' | b'\"')
        })
    {
        Err(invalid("invalid extension attribute name"))
    } else {
        Ok(())
    }
}
pub(in crate::font) fn invalid(e: impl Into<String>) -> Error {
    Error::Invalid(e.into())
}
