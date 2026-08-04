//! Typed WordprocessingML font tables and inert embedded-font resources.

use crate::{Error, Result};
use caseless::Caseless;
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, XmlPart};
use quick_xml::{XmlVersion, events::Event, reader::Reader};
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str::FromStr;
use std::sync::Arc;
use unicode_normalization::UnicodeNormalization;

const WT: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const RT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const FT_RT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
const FT_RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable";
const FONT_RT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const FONT_RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/font";
const FT_CT: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
const FONT_CT: &str = "application/vnd.openxmlformats-officedocument.obfuscatedFont";
const XMLNS: &str = "http://www.w3.org/2000/xmlns/";
const MAX_XML: usize = 8 * 1024 * 1024;
const MAX_FONT: usize = 32 * 1024 * 1024;
const MAX_ALL_FONTS: usize = 128 * 1024 * 1024;
const MAX_FONTS: usize = 4096;
const MAX_NODES: usize = 64_000;
const MAX_DEPTH: usize = 32;
const MAX_TEXT: usize = 32_768;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}
impl Conformance {
    fn word(self) -> &'static str {
        match self {
            Self::Transitional => WT,
            Self::Strict => WS,
        }
    }
    fn rel(self) -> &'static str {
        match self {
            Self::Transitional => RT,
            Self::Strict => RS,
        }
    }
    fn font_table_rel(self) -> &'static str {
        match self {
            Self::Transitional => FT_RT,
            Self::Strict => FT_RS,
        }
    }
    fn font_rel(self) -> &'static str {
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
    fn parse(v: &str) -> Result<Self> {
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
    fn text(self) -> &'static str {
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
    fn parse(v: &str) -> Result<Self> {
        match v {
            "fixed" => Ok(Self::Fixed),
            "variable" => Ok(Self::Variable),
            "default" => Ok(Self::Default),
            _ => Err(invalid(format!("invalid font pitch '{v}'"))),
        }
    }
    fn text(self) -> &'static str {
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
    fn strict(v: &str) -> Result<Self> {
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
    fn element(self) -> &'static str {
        match self {
            Self::Regular => "embedRegular",
            Self::Bold => "embedBold",
            Self::Italic => "embedItalic",
            Self::BoldItalic => "embedBoldItalic",
        }
    }
    fn rank(self) -> u8 {
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
        pub(super) qualified_name: String,
        pub(super) value: String,
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
    unicode_subsets: [u32; 4],
    code_pages: [u32; 2],
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
    part_name: String,
    content_type: String,
    data: Arc<Vec<u8>>,
}
impl Resource {
    /// Own an inert, already-obfuscated embedded-font payload.
    pub fn new(data: Vec<u8>) -> Result<Self> {
        Self::from_shared(Arc::new(data))
    }

    /// Share an inert payload without copying its allocation.
    fn from_shared(data: Arc<Vec<u8>>) -> Result<Self> {
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
    fn share(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.data)
    }

    /// Whether two resources retain the same package allocation.
    pub fn shares_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.data, &other.data)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Embed {
    style: Style,
    relationship_id: String,
    font_key: Option<FontKey>,
    subsetted: Option<bool>,
    resource: Option<Resource>,
    extension_attributes: Vec<raw::Attr>,
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
    name: String,
    alternate_name: Option<String>,
    panose: Option<[u8; 10]>,
    character_set: Option<Charset>,
    family: Option<Family>,
    not_true_type: Option<bool>,
    pitch: Option<Pitch>,
    signature: Option<Signature>,
    embedded_fonts: Vec<Embed>,
    extension_attributes: Vec<raw::Attr>,
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
    fonts: Vec<Font>,
    namespaces: Vec<raw::Attr>,
    extension_attributes: Vec<raw::Attr>,
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

    pub fn xml(&self, conformance: Conformance) -> Result<Vec<u8>> {
        write(self, conformance)
    }

    pub(crate) fn extract_from_part(part: &dyn Part, pkg: &OpcPackage) -> Result<Self> {
        if part.content_type() != FT_CT {
            return Err(Error::ContentType {
                expected: FT_CT.into(),
                actual: part.content_type().into(),
            });
        }
        let mut v = parse(part.blob())?;
        v.resolve(part, pkg)?;
        Ok(v)
    }
    fn resolve(&mut self, source: &dyn Part, pkg: &OpcPackage) -> Result<()> {
        validate_font_relationship_sources(pkg, source.partname())?;
        let mut used = HashSet::new();
        let mut cached = HashMap::<String, Resource>::new();
        let mut targets = HashSet::new();
        let mut total = 0usize;
        for font in &mut self.fonts {
            for embed in &mut font.embedded_fonts {
                used.insert(embed.relationship_id.clone());
                let rel = source.rels().get(&embed.relationship_id).ok_or_else(|| {
                    invalid(format!(
                        "missing embedded-font relationship '{}'",
                        embed.relationship_id
                    ))
                })?;
                if !is_font_relationship(rel.reltype()) {
                    return Err(invalid(format!(
                        "invalid embedded-font relationship type '{}'",
                        rel.reltype()
                    )));
                }
                if rel.is_external() {
                    return Err(invalid("embedded-font relationship cannot be external"));
                }
                let uri = rel.target_partname()?;
                let target_name = uri.to_string();
                targets.insert(target_name.clone());
                if let Some(v) = cached.get(&target_name) {
                    embed.resource = Some(v.clone());
                    continue;
                }
                let part = pkg.get_part(&uri)?;
                if part.content_type() != FONT_CT {
                    return Err(Error::ContentType {
                        expected: FONT_CT.into(),
                        actual: part.content_type().into(),
                    });
                }
                if part.blob().len() > MAX_FONT {
                    return Err(invalid(format!("embedded font '{uri}' is too large")));
                }
                if part.blob().len() < 32 {
                    return Err(invalid(format!("embedded font '{uri}' is too short")));
                }
                total = total
                    .checked_add(part.blob().len())
                    .ok_or_else(|| invalid("embedded-font size overflow"))?;
                if total > MAX_ALL_FONTS {
                    return Err(invalid("embedded fonts exceed total size limit"));
                }
                if part.rels().iter().next().is_some() {
                    return Err(invalid(format!(
                        "embedded font '{uri}' has nested relationships"
                    )));
                }
                let resource = Resource {
                    part_name: uri.to_string(),
                    content_type: part.content_type().into(),
                    data: part.blob_arc(),
                };
                cached.insert(target_name, resource.clone());
                embed.resource = Some(resource)
            }
        }
        for rel in source.rels().iter() {
            if is_font_relationship(rel.reltype()) && !used.contains(rel.r_id()) {
                return Err(invalid(format!(
                    "unreferenced font-table relationship '{}'",
                    rel.r_id()
                )));
            }
        }
        reject_orphan_font_parts(pkg, &targets)?;
        Ok(())
    }
}
impl Default for Table {
    fn default() -> Self {
        Self::new()
    }
}

/// Read the document font table and its bounded, inert font resources.
///
/// Embedded payload allocations are shared with the OPC package rather than
/// copied. The returned table can therefore be queried repeatedly without
/// reparsing or rediscovering the package graph.
pub fn read(package: &OpcPackage) -> Result<Option<Table>> {
    let (main_name, table_name, _) = locate_font_table(package)?;
    validate_font_table_relationship_sources(package, &main_name)?;
    let Some(table_name) = table_name else {
        reject_orphan_font_parts(package, &HashSet::new())?;
        return Ok(None);
    };
    let part = package.get_part(&table_name)?;
    Ok(Some(Table::extract_from_part(part, package)?))
}

/// Move a complete font table into the package after validating the staged XML
/// and OPC graph.
///
/// Font bytes are stored exactly as supplied. Callers that have unobfuscated
/// bytes must explicitly call [`obfuscate`] first. The API
/// operates on an already decrypted in-memory `OpcPackage` and invalidates any
/// package signatures immediately before the mutation phase. Moving a default,
/// empty [`Table`] removes the optional font-table graph and any font resources
/// that become unreferenced.
pub fn put(package: &mut OpcPackage, mut value: Table, conformance: Conformance) -> Result<bool> {
    validate_package_conformance(package, conformance)?;
    let old = read(package)?.unwrap_or_default();
    let (main_name, old_table_name, old_table_relationship_id) = locate_font_table(package)?;
    if value.fonts.is_empty()
        && value.namespaces.is_empty()
        && value.extension_attributes.is_empty()
    {
        return remove_graph(
            package,
            &old,
            &main_name,
            old_table_name.as_ref(),
            old_table_relationship_id.as_deref(),
        );
    }
    if old == value {
        return Ok(false);
    }
    allocate_font_identifiers(package, &mut value)?;
    validate_table_value(&value, true)?;
    let table_name = match old_table_name.clone() {
        Some(name) => name,
        None => next_font_table_part_name(package)?,
    };
    let table_relationship_id = match old_table_relationship_id.clone() {
        Some(id) => id,
        None => next_named_relationship_id(package.get_part(&main_name)?, "rIdTable")?,
    };
    if let Some(existing) = &old_table_name {
        let replaced = old_table_relationship_id
            .iter()
            .cloned()
            .collect::<HashSet<_>>();
        if has_inbound_outside_relationships(package, existing, &main_name, &replaced)? {
            return Err(invalid(format!(
                "shared font-table part '{existing}' cannot be overwritten"
            )));
        }
    }

    let xml = write(&value, conformance)?;
    let staged = parse(&xml)?;
    if !same_metadata(&staged, &value) {
        return Err(invalid("staged font-table XML did not round-trip"));
    }

    let old_relationship_ids = if let Some(name) = &old_table_name {
        package
            .get_part(name)?
            .rels()
            .iter()
            .filter(|relationship| is_font_relationship(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect::<HashSet<_>>()
    } else {
        HashSet::new()
    };
    let old_part_names = old
        .fonts
        .iter()
        .flat_map(|font| font.embedded_fonts.iter())
        .filter_map(|font| {
            font.resource
                .as_ref()
                .map(|resource| resource.part_name.clone())
        })
        .collect::<HashSet<_>>();
    let old_part_uris = old_part_names
        .iter()
        .map(|name| PackURI::new(name).map_err(Error::Uri))
        .collect::<Result<Vec<_>>>()?;

    let table_part = old_table_name
        .as_ref()
        .map(|name| package.get_part(name))
        .transpose()?;
    let mut relationships = HashMap::<String, PackURI>::new();
    let mut resources = HashMap::<String, (String, Arc<Vec<u8>>)>::new();
    for font in &value.fonts {
        for embedded in &font.embedded_fonts {
            if let Some(part) = table_part
                && part.rels().get(&embedded.relationship_id).is_some()
                && !old_relationship_ids.contains(&embedded.relationship_id)
            {
                return Err(invalid(format!(
                    "relationship ID '{}' already exists",
                    embedded.relationship_id
                )));
            }
            let resource = embedded
                .resource
                .as_ref()
                .ok_or_else(|| invalid("embedded-font resource is required for package storage"))?;
            let uri = PackURI::new(&resource.part_name).map_err(Error::Uri)?;
            if let Some(previous) = relationships.get(&embedded.relationship_id) {
                if previous != &uri {
                    return Err(invalid(format!(
                        "relationship ID '{}' has conflicting font targets",
                        embedded.relationship_id
                    )));
                }
            } else {
                relationships.insert(embedded.relationship_id.clone(), uri.clone());
            }
            if let Some((content_type, data)) = resources.get(uri.as_str()) {
                if content_type != &resource.content_type || data.as_slice() != resource.bytes() {
                    return Err(invalid(format!(
                        "shared font part '{uri}' has conflicting resources"
                    )));
                }
            } else {
                resources.insert(
                    uri.to_string(),
                    (resource.content_type.clone(), resource.share()),
                );
            }
        }
    }

    for (part_name, (content_type, data)) in &resources {
        let uri = PackURI::new(part_name).map_err(Error::Uri)?;
        if let Ok(part) = package.get_part(&uri) {
            if part.content_type() != content_type {
                return Err(invalid(format!("font part '{uri}' content type collision")));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{uri}' has outbound relationships"
                )));
            }
            if part.blob() != data.as_slice() && !old_part_names.contains(part_name) {
                return Err(invalid(format!("font part '{uri}' data collision")));
            }
            if part.blob() != data.as_slice()
                && old_table_name.as_ref().is_some_and(|table| {
                    has_inbound_outside_relationships(package, &uri, table, &old_relationship_ids)
                        .unwrap_or(true)
                })
            {
                return Err(invalid(format!(
                    "shared font part '{uri}' cannot be overwritten"
                )));
            }
        }
    }
    validate_all_internal_relationship_targets(package)?;

    let resource_parts = resources
        .into_iter()
        .map(|(name, (content_type, data))| {
            PackURI::new(&name)
                .map(|uri| (uri, content_type, data))
                .map_err(Error::Uri)
        })
        .collect::<Result<Vec<_>>>()?;
    package.unsign();

    for (uri, content_type, data) in resource_parts {
        if let Ok(part) = package.get_part_mut(&uri) {
            part.set_blob_shared(data);
        } else {
            package.add_part(Box::new(BlobPart::new_shared(uri, content_type, data)));
        }
    }
    if let Some(existing) = &old_table_name {
        let part = package.get_part_mut(existing)?;
        let font_relationships = part
            .rels()
            .iter()
            .filter(|relationship| is_font_relationship(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect::<Vec<_>>();
        for relationship_id in font_relationships {
            part.rels_mut().remove(&relationship_id);
        }
        for (relationship_id, target) in &relationships {
            part.rels_mut().add_relationship(
                conformance.font_rel().into(),
                target.relative_ref(table_name.base_uri()),
                relationship_id.clone(),
                false,
            );
        }
        part.set_blob(xml);
    } else {
        let mut part = XmlPart::new(table_name.clone(), FT_CT.into(), xml);
        for (relationship_id, target) in &relationships {
            part.rels_mut().add_relationship(
                conformance.font_rel().into(),
                target.relative_ref(table_name.base_uri()),
                relationship_id.clone(),
                false,
            );
        }
        package.add_part(Box::new(part));
        package
            .get_part_mut(&main_name)?
            .rels_mut()
            .add_relationship(
                conformance.font_table_rel().into(),
                table_name.relative_ref(main_name.base_uri()),
                table_relationship_id,
                false,
            );
    }

    let retained = relationships
        .values()
        .map(ToString::to_string)
        .collect::<HashSet<_>>();
    for uri in old_part_uris {
        if !retained.contains(uri.as_str()) && !part_is_referenced(package, &uri)? {
            package.remove_part(&uri);
        }
    }
    Ok(true)
}

/// Remove the optional font-table graph and every font resource that becomes
/// unreferenced.
///
/// The complete relationship graph is validated before signatures, parts, or
/// relationships are mutated. Resources shared by another source are retained.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    let old = read(package)?.unwrap_or_default();
    let (main_name, table_name, relationship_id) = locate_font_table(package)?;
    remove_graph(
        package,
        &old,
        &main_name,
        table_name.as_ref(),
        relationship_id.as_deref(),
    )
}

fn remove_graph(
    package: &mut OpcPackage,
    old: &Table,
    main_name: &PackURI,
    table_name: Option<&PackURI>,
    table_relationship_id: Option<&str>,
) -> Result<bool> {
    let Some(table_name) = table_name else {
        return Ok(false);
    };
    let table_relationship_id =
        table_relationship_id.ok_or_else(|| invalid("font-table relationship ID is missing"))?;
    let table_part = package.get_part(table_name)?;
    if table_part
        .rels()
        .iter()
        .any(|relationship| !is_font_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "font table with unknown outbound relationships cannot be removed safely",
        ));
    }
    let font_relationship_ids = table_part
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    let replaced_table_relationship = HashSet::from([table_relationship_id.to_owned()]);
    if has_inbound_outside_relationships(
        package,
        table_name,
        main_name,
        &replaced_table_relationship,
    )? {
        return Err(invalid(format!(
            "shared font-table part '{table_name}' cannot be removed"
        )));
    }

    let resource_names = old
        .fonts
        .iter()
        .flat_map(|font| font.embedded_fonts.iter())
        .filter_map(|embed| embed.resource.as_ref())
        .map(|resource| resource.part_name.as_str())
        .collect::<HashSet<_>>();
    let mut resources_to_remove = Vec::with_capacity(resource_names.len());
    for name in resource_names {
        let uri = PackURI::new(name).map_err(Error::Uri)?;
        if !has_inbound_outside_relationships(package, &uri, table_name, &font_relationship_ids)? {
            resources_to_remove.push(uri);
        }
    }
    validate_all_internal_relationship_targets(package)?;

    package.unsign();
    package
        .get_part_mut(main_name)?
        .rels_mut()
        .remove(table_relationship_id);
    package.remove_part(table_name);
    for uri in resources_to_remove {
        package.remove_part(&uri);
    }
    Ok(true)
}

/// Reject embedded typefaces that are not directly named by any `w:rFonts`.
/// Theme-based font resolution is intentionally not attempted.
pub fn validate_usage(package: &OpcPackage, table: &Table) -> Result<()> {
    let used = directly_used_font_names(package)?;
    let unused = table
        .fonts
        .iter()
        .filter(|font| !font.embedded_fonts.is_empty())
        .filter(|font| !used.contains(&name_key(&font.name)))
        .map(|font| font.name.clone())
        .collect::<Vec<_>>();
    if unused.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!(
            "embedded fonts are not directly used by the document: {}",
            unused.join(", ")
        )))
    }
}

fn locate_font_table(package: &OpcPackage) -> Result<(PackURI, Option<PackURI>, Option<String>)> {
    let main = package.main_document_part()?;
    let main_name = main.partname().clone();
    let mut matching = main
        .rels()
        .iter()
        .filter(|relationship| is_font_table_relationship(relationship.reltype()));
    let Some(relationship) = matching.next() else {
        return Ok((main_name, None, None));
    };
    if matching.next().is_some() {
        return Err(invalid("document has multiple font-table relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("font-table relationship cannot be external"));
    }
    Ok((
        main_name,
        Some(relationship.target_partname()?),
        Some(relationship.r_id().to_owned()),
    ))
}

fn allocate_font_identifiers(package: &OpcPackage, table: &mut Table) -> Result<()> {
    let (_, table_name, _) = locate_font_table(package)?;
    let mut relationship_ids = table_name
        .as_ref()
        .map(|name| {
            package.get_part(name).map(|part| {
                part.rels()
                    .iter()
                    .map(|relationship| relationship.r_id().to_owned())
                    .collect::<HashSet<_>>()
            })
        })
        .transpose()?
        .unwrap_or_default();
    relationship_ids.extend(table.fonts.iter().flat_map(|font| {
        font.embedded_fonts
            .iter()
            .filter(|embedded| !embedded.relationship_id.is_empty())
            .map(|embedded| embedded.relationship_id.clone())
    }));
    let mut part_names = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    part_names.extend(table.fonts.iter().flat_map(|font| {
        font.embedded_fonts.iter().filter_map(|embedded| {
            embedded
                .resource
                .as_ref()
                .filter(|resource| !resource.part_name.is_empty())
                .map(|resource| resource.part_name.clone())
        })
    }));
    let mut shared_names = HashMap::<usize, String>::new();
    for font in &table.fonts {
        for embedded in &font.embedded_fonts {
            if let Some(resource) = &embedded.resource
                && !resource.part_name.is_empty()
            {
                shared_names.insert(
                    Arc::as_ptr(&resource.data) as usize,
                    resource.part_name.clone(),
                );
            }
        }
    }
    for font in &mut table.fonts {
        for embedded in &mut font.embedded_fonts {
            if embedded.relationship_id.is_empty() {
                embedded.relationship_id = next_font_relationship_id(&relationship_ids)?;
            }
            relationship_ids.insert(embedded.relationship_id.clone());
            let resource = embedded
                .resource
                .as_mut()
                .ok_or_else(|| invalid("embedded-font resource is required"))?;
            if resource.part_name.is_empty() {
                let identity = Arc::as_ptr(&resource.data) as usize;
                resource.part_name = match shared_names.get(&identity) {
                    Some(name) => name.clone(),
                    None => {
                        let name = next_font_part_name(&part_names)?;
                        shared_names.insert(identity, name.clone());
                        name
                    },
                };
            }
            part_names.insert(resource.part_name.clone());
            if resource.content_type.is_empty() {
                resource.content_type = FONT_CT.into();
            }
        }
    }
    Ok(())
}

fn next_font_relationship_id(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("rIdFont{index}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid("too many font relationship IDs"))
}
fn next_font_part_name(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("/word/fonts/font{index}.odttf");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid("too many font part names"))
}
fn next_font_table_part_name(package: &OpcPackage) -> Result<PackURI> {
    let used = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    if !used.contains("/word/fontTable.xml") {
        return PackURI::new("/word/fontTable.xml").map_err(Error::Uri);
    }
    for index in 1..=u32::MAX {
        let candidate = format!("/word/fontTable{index}.xml");
        if !used.contains(&candidate) {
            return PackURI::new(&candidate).map_err(Error::Uri);
        }
    }
    Err(invalid("too many font-table part names"))
}
fn next_named_relationship_id(source: &dyn Part, prefix: &str) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("{prefix}{index}");
        if source.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("too many relationship IDs"))
}

fn same_metadata(left: &Table, right: &Table) -> bool {
    left.namespaces == right.namespaces
        && left.extension_attributes == right.extension_attributes
        && left.fonts.len() == right.fonts.len()
        && left.fonts.iter().zip(&right.fonts).all(|(left, right)| {
            left.name == right.name
                && left.alternate_name == right.alternate_name
                && left.panose == right.panose
                && left.character_set == right.character_set
                && left.family == right.family
                && left.not_true_type == right.not_true_type
                && left.pitch == right.pitch
                && left.signature == right.signature
                && left.extension_attributes == right.extension_attributes
                && left.embedded_fonts.len() == right.embedded_fonts.len()
                && left
                    .embedded_fonts
                    .iter()
                    .zip(&right.embedded_fonts)
                    .all(|(left, right)| {
                        left.style == right.style
                            && left.relationship_id == right.relationship_id
                            && left.font_key == right.font_key
                            && left.subsetted == right.subsetted
                            && left.extension_attributes == right.extension_attributes
                    })
        })
}

fn name_key(value: &str) -> String {
    value.chars().nfd().default_case_fold().nfd().collect()
}

fn unique_font_offset(fonts: &[Font], name: &str) -> Result<Option<usize>> {
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

fn validate_font_table_relationship_sources(package: &OpcPackage, main: &PackURI) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_font_table_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot source a font-table relationship",
        ));
    }
    for part in package.iter_parts() {
        if part.partname() != main
            && part
                .rels()
                .iter()
                .any(|relationship| is_font_table_relationship(relationship.reltype()))
            && part.content_type()
                != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"
        {
            return Err(invalid(format!(
                "font-table relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_font_relationship_sources(package: &OpcPackage, table: &PackURI) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_font_relationship(relationship.reltype()))
    {
        return Err(invalid("package root cannot source a font relationship"));
    }
    for part in package.iter_parts() {
        if part.partname() != table
            && part
                .rels()
                .iter()
                .any(|relationship| is_font_relationship(relationship.reltype()))
            && part.content_type() != FT_CT
        {
            return Err(invalid(format!(
                "font relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_package_conformance(package: &OpcPackage, requested: Conformance) -> Result<()> {
    const STRICT_OFFICE_DOCUMENT: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    | STRICT_OFFICE_DOCUMENT
            )
        })
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    let actual = if relationship.reltype() == STRICT_OFFICE_DOCUMENT {
        Conformance::Strict
    } else {
        Conformance::Transitional
    };
    if actual == requested {
        Ok(())
    } else {
        Err(invalid(
            "requested font-table conformance does not match the package relationship namespace",
        ))
    }
}

fn reject_orphan_font_parts(package: &OpcPackage, targets: &HashSet<String>) -> Result<()> {
    for part in package.iter_parts() {
        if (part.content_type() == FONT_CT || part.partname().as_str().starts_with("/word/fonts/"))
            && !targets.contains(part.partname().as_str())
            && !part_is_referenced(package, part.partname())?
        {
            return Err(invalid(format!("orphan font part '{}'", part.partname())));
        }
    }
    Ok(())
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn has_inbound_outside_relationships(
    package: &OpcPackage,
    target: &PackURI,
    table: &PackURI,
    replaced_relationships: &HashSet<String>,
) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == *target
                && (part.partname() != table
                    || !replaced_relationships.contains(relationship.r_id()))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn validate_all_internal_relationship_targets(package: &OpcPackage) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        relationship.target_partname()?;
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            relationship.target_partname()?;
        }
    }
    Ok(())
}

fn directly_used_font_names(package: &OpcPackage) -> Result<HashSet<String>> {
    let mut output = HashSet::new();
    for part in package.iter_parts().filter(|part| {
        part.content_type().contains("wordprocessingml")
            && part.content_type().ends_with("+xml")
            && part.content_type() != FT_CT
    }) {
        if part.blob().len() > MAX_XML {
            return Err(invalid(format!(
                "WordprocessingML part '{}' is too large for font-usage validation",
                part.partname()
            )));
        }
        let mut reader = Reader::from_reader(part.blob());
        let mut nodes = 0usize;
        loop {
            match reader.read_event().map_err(xml_error)? {
                Event::Start(element) | Event::Empty(element)
                    if element.local_name().as_ref() == b"rFonts" =>
                {
                    nodes += 1;
                    if nodes > MAX_NODES {
                        return Err(invalid("font-usage XML node limit exceeded"));
                    }
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(xml_error)?;
                        if matches!(
                            attribute.key.local_name().as_ref(),
                            b"ascii" | b"hAnsi" | b"eastAsia" | b"cs"
                        ) {
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(xml_error)?;
                            bounded(&value)?;
                            output.insert(name_key(&value));
                        }
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("DTDs and processing instructions are rejected"));
                },
                Event::Eof => break,
                _ => {},
            }
        }
    }
    Ok(output)
}

fn validate_table_value(value: &Table, require_resources: bool) -> Result<()> {
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

fn validate_font_entry(font: &Font, require_resources: bool) -> Result<()> {
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

fn validate_resource_len(len: usize) -> Result<()> {
    if (32..=MAX_FONT).contains(&len) {
        Ok(())
    } else {
        Err(invalid("embedded font size is outside the allowed bounds"))
    }
}

fn validate_font_name(value: &str, kind: &str) -> Result<()> {
    let count = value.chars().count();
    if !(1..=31).contains(&count) {
        Err(invalid(format!(
            "{kind} must contain 1 through 31 characters"
        )))
    } else {
        Ok(())
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

fn is_font_table_relationship(v: &str) -> bool {
    matches!(v, FT_RT | FT_RS)
}
fn is_font_relationship(v: &str) -> bool {
    matches!(v, FONT_RT | FONT_RS)
}
fn word_ns(v: &str) -> bool {
    matches!(v, WT | WS)
}
fn rel_ns(v: &str) -> bool {
    matches!(v, RT | RS)
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
fn bounded(v: &str) -> Result<()> {
    if v.len() <= MAX_TEXT {
        Ok(())
    } else {
        Err(invalid("font-table string limit exceeded"))
    }
}
fn validate_attr_name(value: &str) -> Result<()> {
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
fn xml_error(e: impl fmt::Display) -> Error {
    Error::Xml(e.to_string())
}
fn invalid(e: impl Into<String>) -> Error {
    Error::Invalid(e.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    const KEY: FontKey = FontKey::new([
        0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE,
        0xFF,
    ]);

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        let document = PackURI::new("/word/document.xml").expect("test URI");
        package.add_part(Box::new(XmlPart::new(
            document,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .into(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
        )));
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "word/document.xml".into(),
            "rId1".into(),
            false,
        );
        package
    }

    #[test]
    fn strict_round_trip_and_safe_selectors() {
        let xml = br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="A&amp;B"><w:altName w:val="Alias"/><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:notTrueType w:val="0"/><w:pitch w:val="variable"/><w:sig w:usb0="E10002FF" w:usb1="4000ACFF" w:usb2="00000009" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/><w:embedRegular r:id="rId1" w:fontKey="{01014A78-CABC-4EF0-12AC-5CD89AEFDE01}" w:subsetted="1"/></w:font></w:fonts>"#;
        let table = parse(xml).expect("parse");
        let by_name = table.get("a&b").expect("lookup").expect("font");
        let by_index = table.get(0usize).expect("lookup").expect("font");
        assert_eq!(by_name, by_index);
        assert!(table.get(9usize).expect("lookup").is_none());
        assert_eq!(
            by_name.signature().expect("signature").code_pages()[0],
            0x19F
        );

        let strict = write(&table, Conformance::Strict).expect("write");
        assert!(std::str::from_utf8(&strict).expect("UTF-8").contains(WS));
        assert_eq!(parse(&strict).expect("reparse"), table);
    }

    #[test]
    fn mce_and_real_strict_fixture() {
        let xml = br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:font/></mc:Choice><mc:Fallback><w:font w:name="Fallback"><w:family w:val="roman"/></w:font></mc:Fallback></mc:AlternateContent></w:fonts>"#;
        assert_eq!(
            parse(xml)
                .expect("MCE parse")
                .get("Fallback")
                .expect("lookup")
                .expect("font")
                .family(),
            Some(Family::Roman)
        );

        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
            root.join("test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx"),
        )
        .expect("open fixture");
        let uri = PackURI::new("/word/fontTable.xml").expect("test URI");
        let table = parse(&physical.blob_for(&uri).expect("font table")).expect("parse");
        assert!(table.get("Calibri").expect("lookup").is_some());
        assert_eq!(
            table.get(0usize).expect("lookup").expect("font").charset(),
            Some(Charset::Ansi)
        );
    }

    #[test]
    fn malformed_order_and_bounds_are_rejected() {
        for xml in [
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font/></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:family w:val="fantasy"/></w:font></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:panose1 w:val="1234"/></w:font></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:pitch w:val="fixed"/><w:family w:val="roman"/></w:font></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="x"><w:embedRegular r:id="rId1" w:fontKey="bad"/></w:font></w:fonts>"#,
            r#"<!DOCTYPE x><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        ] {
            assert!(parse(xml.as_bytes()).is_err(), "{xml}");
        }
        assert!(parse(&vec![b' '; MAX_XML + 1]).is_err());
    }

    #[test]
    fn poi_resources_share_package_allocations() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let package =
            OpcPackage::open(root.join("test-data/poi/test-data/document/saut_page.docx"))
                .expect("open fixture");
        let table = read(&package).expect("read").expect("font table");
        assert_eq!(table.len(), 7);
        let embeds = table.iter().flat_map(Font::embeds).collect::<Vec<_>>();
        assert_eq!(embeds.len(), 20);
        let first = embeds
            .iter()
            .find_map(|embed| embed.resource())
            .expect("embedded resource");
        let uri = PackURI::new(&first.part_name).expect("fixture part URI");
        assert!(Arc::ptr_eq(
            &first.data,
            &package.get_part(&uri).expect("font part").blob_arc()
        ));
        assert!(embeds.iter().all(|embed| embed.resource().is_some()));
    }

    #[test]
    fn obfuscation_and_compact_license_are_checked() {
        let original = (0u8..64).collect::<Vec<_>>();
        let mut data = original.clone();
        obfuscate(&mut data, KEY).expect("obfuscate");
        assert_ne!(data, original);
        assert_eq!(&data[32..], &original[32..]);
        deobfuscate(&mut data, KEY).expect("deobfuscate");
        assert_eq!(data, original);
        assert!(obfuscate(&mut [0; 31], KEY).is_err());
        assert!("bad".parse::<FontKey>().is_err());
        assert_eq!(KEY.to_string(), "{00112233-4455-6677-8899-AABBCCDDEEFF}");

        assert!(License::new(0).expect("license").installable());
        let editable = License::new(0x0108).expect("license");
        assert!(editable.editable() && editable.no_subsetting());
        assert_eq!(editable.bits(), 0x0108);
        assert!(License::new(0x0006).is_err());
        assert!(License::new(0x8000).is_err());
        assert_eq!(size_of::<License>(), size_of::<u16>());
    }

    #[test]
    fn move_first_crud_preserves_shared_resources_and_extensions() {
        let mut package = package();
        let shared = Resource::new((0u8..64).collect()).expect("resource");
        let first = Font::new("Alpha")
            .expect("font")
            .with_alt("Alpha Alt")
            .expect("alternate")
            .with_panose([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
            .with_charset(Charset::Ansi)
            .with_family(Family::Swiss)
            .with_pitch(Pitch::Variable)
            .with_signature(Signature::new([1, 2, 3, 4], [5, 6]))
            .with_embed(Embed::new(Style::Regular, KEY, shared.clone()).with_subset(true))
            .expect("face")
            .with_attr(raw::Attr::new("x:flag", "kept").expect("attribute"));
        let mut table = Table::new()
            .with_namespace(raw::Attr::new("xmlns:x", "urn:test-fonts").expect("namespace"))
            .expect("namespace");
        table.add(first).expect("add");
        put(&mut package, table, Conformance::Transitional).expect("put");

        let mut table = read(&package).expect("read").expect("table");
        table
            .add(
                Font::new("Beta")
                    .expect("font")
                    .with_embed(Embed::new(Style::Regular, KEY, shared))
                    .expect("face"),
            )
            .expect("add");
        table.reorder(&["Beta", "Alpha"]).expect("reorder");
        put(&mut package, table, Conformance::Transitional).expect("put");

        let mut table = read(&package).expect("read").expect("table");
        assert_eq!(
            table.get(0usize).expect("lookup").expect("font").name(),
            "Beta"
        );
        assert_eq!(
            table
                .get("alpha")
                .expect("lookup")
                .expect("font")
                .attrs()
                .first()
                .expect("attribute")
                .value(),
            "kept"
        );
        let beta = table.get("Beta").expect("lookup").expect("font");
        let alpha = table.get("Alpha").expect("lookup").expect("font");
        assert!(
            beta.embeds()[0]
                .resource()
                .expect("resource")
                .shares_with(alpha.embeds()[0].resource().expect("resource"))
        );
        let shared_part = beta.embeds()[0]
            .resource()
            .expect("resource")
            .part_name
            .clone();

        assert!(table.remove("Alpha").expect("remove").is_some());
        put(&mut package, table, Conformance::Transitional).expect("put");
        assert!(
            package
                .get_part(&PackURI::new(&shared_part).expect("part URI"))
                .is_ok()
        );
        assert!(
            read(&package)
                .expect("read")
                .expect("table")
                .get("Beta")
                .expect("lookup")
                .is_some()
        );

        put(&mut package, Table::new(), Conformance::Transitional).expect("remove graph");
        assert!(read(&package).expect("read").is_none());
        assert!(
            package
                .get_part(&PackURI::new(&shared_part).expect("part URI"))
                .is_err()
        );
    }

    #[test]
    fn graph_delete_keeps_resources_referenced_outside_the_table() {
        let mut package = package();
        let font = Font::new("Shared")
            .expect("font")
            .with_embed(Embed::new(
                Style::Regular,
                KEY,
                Resource::new(vec![0; 32]).expect("resource"),
            ))
            .expect("face");
        let mut table = Table::new();
        table.add(font).expect("add");
        put(&mut package, table, Conformance::Transitional).expect("put");

        let table = read(&package).expect("read").expect("table");
        let resource_name = table.fonts[0].embedded_fonts[0]
            .resource
            .as_ref()
            .expect("resource")
            .part_name
            .clone();
        let resource = PackURI::new(&resource_name).expect("resource URI");
        let main_name = package
            .main_document_part()
            .expect("main")
            .partname()
            .clone();
        package
            .get_part_mut(&main_name)
            .expect("main")
            .rels_mut()
            .add_relationship(
                "urn:litchi:test:keep-font".into(),
                resource.relative_ref(main_name.base_uri()),
                "rIdKeepFont".into(),
                false,
            );

        assert!(remove(&mut package).expect("remove graph"));
        assert!(read(&package).expect("read").is_none());
        assert!(package.get_part(&resource).is_ok());
        assert!(!remove(&mut package).expect("already absent"));
    }

    #[test]
    fn constructors_prevent_invalid_authoring_state() {
        assert!(Font::new("").is_err());
        assert!(Font::new("12345678901234567890123456789012").is_err());
        assert!(Resource::new(vec![0; 31]).is_err());
        assert!("bad".parse::<FontKey>().is_err());
        assert!(raw::Attr::new("", "value").is_err());

        let mut table = Table::new();
        table.add(Font::new("Alpha").expect("font")).expect("add");
        assert!(table.add(Font::new("alpha").expect("font")).is_err());
        assert!(
            table
                .replace(9usize, Font::new("Beta").expect("font"))
                .expect("replace")
                .is_none()
        );
        assert!(table.remove(9usize).expect("remove").is_none());
    }

    #[test]
    fn unicode_caseless_identity_is_consistent_and_ambiguous_sources_fail() {
        let mut table = Table::new();
        table.add(Font::new("Straße").expect("font")).expect("add");
        assert!(table.add(Font::new("STRASSE").expect("font")).is_err());
        table.add(Font::new("École").expect("font")).expect("add");

        assert!(table.get("strasse").expect("lookup").is_some());
        assert!(table.get("e\u{301}COLE").expect("lookup").is_some());
        table
            .reorder(&["e\u{301}cole", "STRASSE"])
            .expect("reorder");
        assert_eq!(
            table.get(0usize).expect("lookup").expect("font").name(),
            "École"
        );

        let malformed_xml = r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Straße"/><w:font w:name="STRASSE"/></w:fonts>"#;
        let malformed = parse(malformed_xml.as_bytes()).expect("parse malformed producer table");
        assert!(malformed.get("strasse").is_err());
    }
}
