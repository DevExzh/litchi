//! Typed WordprocessingML font tables and inert embedded-font resources.

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, XmlPart};
use quick_xml::{XmlVersion, events::Event, reader::Reader};
use std::collections::{HashMap, HashSet};

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
pub enum FontTableConformance {
    Transitional,
    Strict,
}
impl FontTableConformance {
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
pub struct EmbeddedFontLicensing {
    pub fs_type: u16,
    pub restricted_license: bool,
    pub preview_and_print: bool,
    pub editable: bool,
    pub no_subsetting: bool,
    pub bitmap_only: bool,
}
impl EmbeddedFontLicensing {
    pub fn from_fs_type(fs_type: u16) -> Result<Self> {
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
        Ok(Self {
            fs_type,
            restricted_license: fs_type & 0x0002 != 0,
            preview_and_print: fs_type & 0x0004 != 0,
            editable: fs_type & 0x0008 != 0,
            no_subsetting: fs_type & 0x0100 != 0,
            bitmap_only: fs_type & 0x0200 != 0,
        })
    }
    pub fn installable(self) -> bool {
        self.fs_type & 0x000E == 0
    }
}

/// Apply the reversible OOXML GUID XOR transformation to the first 32 bytes.
pub fn obfuscate_embedded_font_data(data: &mut [u8], font_key: &str) -> Result<()> {
    if data.len() < 32 {
        return Err(invalid("OOXML font obfuscation requires at least 32 bytes"));
    }
    let key = parse_font_key_bytes(font_key)?;
    for index in 0..32 {
        data[index] ^= key[15 - (index % 16)];
    }
    Ok(())
}

/// Reverse [`obfuscate_embedded_font_data`]. XOR makes both operations identical.
pub fn deobfuscate_embedded_font_data(data: &mut [u8], font_key: &str) -> Result<()> {
    obfuscate_embedded_font_data(data, font_key)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FontFamily {
    Decorative,
    Modern,
    Roman,
    Script,
    Swiss,
    Auto,
}
impl FontFamily {
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
pub enum FontPitch {
    Fixed,
    Variable,
    Default,
}
impl FontPitch {
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
pub enum FontCharacterSet {
    Ansi,
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
    Legacy(u8),
}
impl FontCharacterSet {
    fn legacy(v: u8) -> Self {
        match v {
            0x00 => Self::Ansi,
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
            Self::Legacy(v) => v,
        }
    }
    pub fn strict_name(self) -> Option<&'static str> {
        Some(match self {
            Self::Ansi => "iso-8859-1",
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
pub enum EmbeddedFontStyle {
    Regular,
    Bold,
    Italic,
    BoldItalic,
}
impl EmbeddedFontStyle {
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontTableExtensionAttribute {
    qualified_name: String,
    value: String,
}
impl FontTableExtensionAttribute {
    pub fn new(qualified_name: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            qualified_name: qualified_name.into(),
            value: value.into(),
        }
    }
    pub fn qualified_name(&self) -> &str {
        &self.qualified_name
    }
    pub fn value(&self) -> &str {
        &self.value
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontSignature {
    unicode_subsets: [u32; 4],
    code_pages: [u32; 2],
}
impl FontSignature {
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
pub struct EmbeddedFontResource {
    part_name: String,
    content_type: String,
    data: Vec<u8>,
}
impl EmbeddedFontResource {
    pub fn new(part_name: impl Into<String>, data: Vec<u8>) -> Self {
        Self {
            part_name: part_name.into(),
            content_type: FONT_CT.into(),
            data,
        }
    }
    pub fn with_content_type(mut self, content_type: impl Into<String>) -> Self {
        self.content_type = content_type.into();
        self
    }
    pub fn set_part_name(&mut self, part_name: impl Into<String>) {
        self.part_name = part_name.into();
    }
    pub fn set_data(&mut self, data: Vec<u8>) {
        self.data = data;
    }
    pub fn part_name(&self) -> &str {
        &self.part_name
    }
    pub fn content_type(&self) -> &str {
        &self.content_type
    }
    pub fn data(&self) -> &[u8] {
        &self.data
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedFont {
    style: EmbeddedFontStyle,
    relationship_id: String,
    font_key: Option<String>,
    subsetted: Option<bool>,
    resource: Option<EmbeddedFontResource>,
    extension_attributes: Vec<FontTableExtensionAttribute>,
}
impl EmbeddedFont {
    pub fn new(style: EmbeddedFontStyle, resource: EmbeddedFontResource) -> Self {
        Self {
            style,
            relationship_id: String::new(),
            font_key: None,
            subsetted: None,
            resource: Some(resource),
            extension_attributes: Vec::new(),
        }
    }
    pub fn with_relationship_id(mut self, relationship_id: impl Into<String>) -> Self {
        self.relationship_id = relationship_id.into();
        self
    }
    pub fn with_font_key(mut self, font_key: impl Into<String>) -> Self {
        self.font_key = Some(font_key.into());
        self
    }
    pub fn with_subsetted(mut self, subsetted: bool) -> Self {
        self.subsetted = Some(subsetted);
        self
    }
    pub fn set_relationship_id(&mut self, relationship_id: impl Into<String>) {
        self.relationship_id = relationship_id.into();
    }
    pub fn set_font_key(&mut self, font_key: Option<String>) {
        self.font_key = font_key;
    }
    pub fn set_subsetted(&mut self, subsetted: Option<bool>) {
        self.subsetted = subsetted;
    }
    pub fn resource_mut(&mut self) -> Option<&mut EmbeddedFontResource> {
        self.resource.as_mut()
    }
    pub fn style(&self) -> EmbeddedFontStyle {
        self.style
    }
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
    pub fn font_key(&self) -> Option<&str> {
        self.font_key.as_deref()
    }
    pub fn subsetted(&self) -> Option<bool> {
        self.subsetted
    }
    pub fn resource(&self) -> Option<&EmbeddedFontResource> {
        self.resource.as_ref()
    }
    pub fn extension_attributes(&self) -> &[FontTableExtensionAttribute] {
        &self.extension_attributes
    }
    pub fn extension_attributes_mut(&mut self) -> &mut Vec<FontTableExtensionAttribute> {
        &mut self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    name: String,
    alternate_name: Option<String>,
    panose: Option<[u8; 10]>,
    character_set: Option<FontCharacterSet>,
    family: Option<FontFamily>,
    not_true_type: Option<bool>,
    pitch: Option<FontPitch>,
    signature: Option<FontSignature>,
    embedded_fonts: Vec<EmbeddedFont>,
    extension_attributes: Vec<FontTableExtensionAttribute>,
}
impl Font {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            alternate_name: None,
            panose: None,
            character_set: None,
            family: None,
            not_true_type: None,
            pitch: None,
            signature: None,
            embedded_fonts: Vec::new(),
            extension_attributes: Vec::new(),
        }
    }
    pub fn with_alternate_name(mut self, value: impl Into<String>) -> Self {
        self.alternate_name = Some(value.into());
        self
    }
    pub fn with_panose(mut self, value: [u8; 10]) -> Self {
        self.panose = Some(value);
        self
    }
    pub fn with_character_set(mut self, value: FontCharacterSet) -> Self {
        self.character_set = Some(value);
        self
    }
    pub fn with_family(mut self, value: FontFamily) -> Self {
        self.family = Some(value);
        self
    }
    pub fn with_not_true_type(mut self, value: bool) -> Self {
        self.not_true_type = Some(value);
        self
    }
    pub fn with_pitch(mut self, value: FontPitch) -> Self {
        self.pitch = Some(value);
        self
    }
    pub fn with_signature(mut self, value: FontSignature) -> Self {
        self.signature = Some(value);
        self
    }
    pub fn with_embedded_font(mut self, value: EmbeddedFont) -> Self {
        self.embedded_fonts.push(value);
        self
    }
    pub fn set_name(&mut self, value: impl Into<String>) {
        self.name = value.into();
    }
    pub fn embedded_fonts_mut(&mut self) -> &mut Vec<EmbeddedFont> {
        &mut self.embedded_fonts
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn alternate_name(&self) -> Option<&str> {
        self.alternate_name.as_deref()
    }
    pub fn panose(&self) -> Option<&[u8; 10]> {
        self.panose.as_ref()
    }
    pub fn character_set(&self) -> Option<FontCharacterSet> {
        self.character_set
    }
    pub fn family(&self) -> Option<FontFamily> {
        self.family
    }
    pub fn not_true_type(&self) -> Option<bool> {
        self.not_true_type
    }
    pub fn pitch(&self) -> Option<FontPitch> {
        self.pitch
    }
    pub fn signature(&self) -> Option<&FontSignature> {
        self.signature.as_ref()
    }
    pub fn embedded_fonts(&self) -> &[EmbeddedFont] {
        &self.embedded_fonts
    }
    pub fn extension_attributes(&self) -> &[FontTableExtensionAttribute] {
        &self.extension_attributes
    }
    pub fn extension_attributes_mut(&mut self) -> &mut Vec<FontTableExtensionAttribute> {
        &mut self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontTable {
    fonts: Vec<Font>,
    namespaces: Vec<FontTableExtensionAttribute>,
    extension_attributes: Vec<FontTableExtensionAttribute>,
}
impl FontTable {
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
    pub fn extension_attributes(&self) -> &[FontTableExtensionAttribute] {
        &self.extension_attributes
    }
    pub fn namespaces(&self) -> &[FontTableExtensionAttribute] {
        &self.namespaces
    }
    pub fn fonts_mut(&mut self) -> &mut Vec<Font> {
        &mut self.fonts
    }
    pub fn namespaces_mut(&mut self) -> &mut Vec<FontTableExtensionAttribute> {
        &mut self.namespaces
    }
    pub fn extension_attributes_mut(&mut self) -> &mut Vec<FontTableExtensionAttribute> {
        &mut self.extension_attributes
    }
    pub fn to_xml(&self, c: FontTableConformance) -> Result<Vec<u8>> {
        write_font_table(self, c)
    }
    pub(crate) fn extract_from_part(part: &dyn Part, pkg: &OpcPackage) -> Result<Self> {
        if part.content_type() != FT_CT {
            return Err(OoxmlError::InvalidContentType {
                expected: FT_CT.into(),
                got: part.content_type().into(),
            });
        }
        let mut v = parse_font_table(part.blob())?;
        v.resolve(part, pkg)?;
        Ok(v)
    }
    fn resolve(&mut self, source: &dyn Part, pkg: &OpcPackage) -> Result<()> {
        validate_font_relationship_sources(pkg, source.partname())?;
        let mut used = HashSet::new();
        let mut cached = HashMap::<String, EmbeddedFontResource>::new();
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
                    return Err(OoxmlError::InvalidContentType {
                        expected: FONT_CT.into(),
                        got: part.content_type().into(),
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
                let resource = EmbeddedFontResource {
                    part_name: uri.to_string(),
                    content_type: part.content_type().into(),
                    data: part.blob().to_vec(),
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
impl Default for FontTable {
    fn default() -> Self {
        Self::new()
    }
}

/// Load the document font table and its bounded, inert font resources.
pub fn load_font_table(package: &OpcPackage) -> Result<Option<FontTable>> {
    let (main_name, table_name, _) = locate_font_table(package)?;
    validate_font_table_relationship_sources(package, &main_name)?;
    let Some(table_name) = table_name else {
        reject_orphan_font_parts(package, &HashSet::new())?;
        return Ok(None);
    };
    let part = package.get_part(&table_name)?;
    Ok(Some(FontTable::extract_from_part(part, package)?))
}

/// Store a complete font table after validating the staged XML and OPC graph.
///
/// Font bytes are stored exactly as supplied. Callers that have unobfuscated
/// bytes must explicitly call [`obfuscate_embedded_font_data`] first. The API
/// operates on an already decrypted in-memory `OpcPackage` and invalidates any
/// package signatures immediately before the mutation phase.
pub fn store_font_table(
    package: &mut OpcPackage,
    value: &FontTable,
    conformance: FontTableConformance,
) -> Result<()> {
    validate_package_conformance(package, conformance)?;
    validate_table_value(value, true)?;
    let old = load_font_table(package)?.unwrap_or_default();
    let (main_name, old_table_name, old_table_relationship_id) = locate_font_table(package)?;
    if old_table_name.is_none() && value.fonts.is_empty() {
        return Ok(());
    }
    let table_name = old_table_name.clone().unwrap_or_else(|| {
        next_font_table_part_name(package)
            .expect("the bounded part-name allocation was preflighted")
    });
    let table_relationship_id = old_table_relationship_id.clone().unwrap_or_else(|| {
        next_named_relationship_id(package.get_part(&main_name).unwrap(), "rIdFontTable")
    });
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

    let xml = write_font_table(value, conformance)?;
    let staged = parse_font_table(&xml)?;
    if staged != metadata_only(value) {
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
        .map(|name| PackURI::new(name).map_err(OoxmlError::InvalidUri))
        .collect::<Result<Vec<_>>>()?;

    let table_part = old_table_name
        .as_ref()
        .map(|name| package.get_part(name))
        .transpose()?;
    let mut relationships = HashMap::<String, PackURI>::new();
    let mut resources = HashMap::<String, (String, Vec<u8>)>::new();
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
            let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
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
                if content_type != &resource.content_type || data != &resource.data {
                    return Err(invalid(format!(
                        "shared font part '{uri}' has conflicting resources"
                    )));
                }
            } else {
                resources.insert(
                    uri.to_string(),
                    (resource.content_type.clone(), resource.data.clone()),
                );
            }
        }
    }

    for (part_name, (content_type, data)) in &resources {
        let uri = PackURI::new(part_name).map_err(OoxmlError::InvalidUri)?;
        if let Ok(part) = package.get_part(&uri) {
            if part.content_type() != content_type {
                return Err(invalid(format!("font part '{uri}' content type collision")));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{uri}' has outbound relationships"
                )));
            }
            if part.blob() != data && !old_part_names.contains(part_name) {
                return Err(invalid(format!("font part '{uri}' data collision")));
            }
            if part.blob() != data
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
                .map_err(OoxmlError::InvalidUri)
        })
        .collect::<Result<Vec<_>>>()?;
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::Other(format!("cannot invalidate package signatures: {error}"))
    })?;

    for (uri, content_type, data) in resource_parts {
        if let Ok(part) = package.get_part_mut(&uri) {
            part.set_blob(data);
        } else {
            package.add_part(Box::new(BlobPart::new(uri, content_type, data)));
        }
    }
    if let Some(existing) = &old_table_name {
        let part = package
            .get_part_mut(existing)
            .expect("font-table part was validated before mutation");
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
            .get_part_mut(&main_name)
            .expect("main document was validated before mutation")
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
        if !retained.contains(uri.as_str())
            && !part_is_referenced(package, &uri)
                .expect("all internal relationship targets were preflighted")
        {
            package.remove_part(&uri);
        }
    }
    Ok(())
}

pub fn find_font(package: &OpcPackage, name: &str) -> Result<Option<Font>> {
    let Some(table) = load_font_table(package)? else {
        return Ok(None);
    };
    Ok(unique_font_offset(&table.fonts, name)?.map(|offset| table.fonts[offset].clone()))
}

pub fn add_font(
    package: &mut OpcPackage,
    mut font: Font,
    conformance: FontTableConformance,
) -> Result<()> {
    let mut table = load_font_table(package)?.unwrap_or_default();
    if table
        .fonts
        .iter()
        .any(|item| item.name.eq_ignore_ascii_case(&font.name))
    {
        return Err(invalid(format!("font '{}' already exists", font.name)));
    }
    allocate_font_identifiers(package, &mut font, &table)?;
    table.fonts.push(font);
    store_font_table(package, &table, conformance)
}

pub fn update_font(
    package: &mut OpcPackage,
    name: &str,
    mut replacement: Font,
    conformance: FontTableConformance,
) -> Result<()> {
    let mut table =
        load_font_table(package)?.ok_or_else(|| invalid("document has no font table"))?;
    let offset = unique_font_offset(&table.fonts, name)?
        .ok_or_else(|| invalid(format!("font '{name}' was not found")))?;
    table.fonts.remove(offset);
    allocate_font_identifiers(package, &mut replacement, &table)?;
    table.fonts.insert(offset, replacement);
    store_font_table(package, &table, conformance)
}

pub fn replace_font(
    package: &mut OpcPackage,
    name: &str,
    replacement: Font,
    conformance: FontTableConformance,
) -> Result<()> {
    update_font(package, name, replacement, conformance)
}

pub fn remove_font(
    package: &mut OpcPackage,
    name: &str,
    conformance: FontTableConformance,
) -> Result<bool> {
    let Some(mut table) = load_font_table(package)? else {
        return Ok(false);
    };
    let Some(offset) = unique_font_offset(&table.fonts, name)? else {
        return Ok(false);
    };
    table.fonts.remove(offset);
    store_font_table(package, &table, conformance)?;
    Ok(true)
}

pub fn reorder_fonts(
    package: &mut OpcPackage,
    ordered_names: &[String],
    conformance: FontTableConformance,
) -> Result<()> {
    let mut table =
        load_font_table(package)?.ok_or_else(|| invalid("document has no font table"))?;
    let expected = table
        .fonts
        .iter()
        .map(|font| font.name.to_lowercase())
        .collect::<HashSet<_>>();
    if expected.len() != table.fonts.len() {
        return Err(invalid(
            "cannot reorder a font table with ambiguous case-insensitive names",
        ));
    }
    let actual = ordered_names
        .iter()
        .map(|name| name.to_lowercase())
        .collect::<HashSet<_>>();
    if expected != actual || ordered_names.len() != table.fonts.len() {
        return Err(invalid("font-table reorder is not a font-name permutation"));
    }
    table.fonts = ordered_names
        .iter()
        .map(|name| {
            table
                .fonts
                .iter()
                .find(|font| font.name.eq_ignore_ascii_case(name))
                .expect("the font-name permutation was validated")
                .clone()
        })
        .collect();
    store_font_table(package, &table, conformance)
}

/// Reject embedded typefaces that are not directly named by any `w:rFonts`.
/// Theme-based font resolution is intentionally not attempted.
pub fn validate_embedded_font_usage(package: &OpcPackage, table: &FontTable) -> Result<()> {
    let used = directly_used_font_names(package)?;
    let unused = table
        .fonts
        .iter()
        .filter(|font| !font.embedded_fonts.is_empty())
        .filter(|font| !used.contains(&font.name.to_lowercase()))
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

fn allocate_font_identifiers(
    package: &OpcPackage,
    font: &mut Font,
    existing: &FontTable,
) -> Result<()> {
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
    relationship_ids.extend(existing.fonts.iter().flat_map(|font| {
        font.embedded_fonts
            .iter()
            .map(|embedded| embedded.relationship_id.clone())
    }));
    let mut part_names = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    part_names.extend(existing.fonts.iter().flat_map(|font| {
        font.embedded_fonts.iter().filter_map(|embedded| {
            embedded
                .resource
                .as_ref()
                .map(|resource| resource.part_name.clone())
        })
    }));
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
            resource.part_name = next_font_part_name(&part_names)?;
        }
        part_names.insert(resource.part_name.clone());
        if resource.content_type.is_empty() {
            resource.content_type = FONT_CT.into();
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
        return PackURI::new("/word/fontTable.xml").map_err(OoxmlError::InvalidUri);
    }
    for index in 1..=u32::MAX {
        let candidate = format!("/word/fontTable{index}.xml");
        if !used.contains(&candidate) {
            return PackURI::new(&candidate).map_err(OoxmlError::InvalidUri);
        }
    }
    Err(invalid("too many font-table part names"))
}
fn next_named_relationship_id(source: &dyn Part, prefix: &str) -> String {
    for index in 1..=u32::MAX {
        let candidate = format!("{prefix}{index}");
        if source.rels().get(&candidate).is_none() {
            return candidate;
        }
    }
    unreachable!("u32 relationship ID space exhausted")
}

fn metadata_only(value: &FontTable) -> FontTable {
    let mut value = value.clone();
    for font in &mut value.fonts {
        for embedded in &mut font.embedded_fonts {
            embedded.resource = None;
        }
    }
    value
}

fn unique_font_offset(fonts: &[Font], name: &str) -> Result<Option<usize>> {
    let mut matching = fonts
        .iter()
        .enumerate()
        .filter(|(_, font)| font.name.eq_ignore_ascii_case(name))
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

fn validate_package_conformance(
    package: &OpcPackage,
    requested: FontTableConformance,
) -> Result<()> {
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
        FontTableConformance::Strict
    } else {
        FontTableConformance::Transitional
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
                            output.insert(value.to_lowercase());
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

fn validate_table_value(value: &FontTable, require_resources: bool) -> Result<()> {
    if value.fonts.len() > MAX_FONTS {
        return Err(invalid("too many fonts"));
    }
    let mut total = 0usize;
    let mut resource_names = HashSet::new();
    for font in &value.fonts {
        validate_font_name(&font.name, "font name")?;
        if let Some(name) = &font.alternate_name {
            validate_font_name(name, "alternate font name")?;
        }
        for pair in font.embedded_fonts.windows(2) {
            if pair[0].style.rank() >= pair[1].style.rank() {
                return Err(invalid(
                    "embedded-font styles are duplicated or out of schema order",
                ));
            }
        }
        for embedded in &font.embedded_fonts {
            if embedded.relationship_id.is_empty() || embedded.relationship_id.len() > MAX_TEXT {
                return Err(invalid(
                    "embedded-font relationship ID is empty or too long",
                ));
            }
            if let Some(key) = &embedded.font_key {
                font_key(key)?;
            } else if require_resources {
                return Err(invalid("fontKey is required for package storage"));
            }
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
                if resource.data.len() < 32 || resource.data.len() > MAX_FONT {
                    return Err(invalid("embedded font size is outside the allowed bounds"));
                }
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

fn parse_font_key_bytes(value: &str) -> Result<[u8; 16]> {
    font_key(value)?;
    let value = value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
        .unwrap_or(value);
    let digits = value
        .bytes()
        .filter(|byte| *byte != b'-')
        .collect::<Vec<_>>();
    let mut key = [0u8; 16];
    for (index, output) in key.iter_mut().enumerate() {
        let pair = std::str::from_utf8(&digits[index * 2..index * 2 + 2]).map_err(xml_error)?;
        *output = u8::from_str_radix(pair, 16).map_err(xml_error)?;
    }
    Ok(key)
}

#[derive(Clone)]
struct Attr {
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
    attrs: Vec<Attr>,
    children: Vec<Node>,
    text: String,
}

pub fn is_font_table_relationship(v: &str) -> bool {
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

pub fn parse_font_table(xml: &[u8]) -> Result<FontTable> {
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
                let (n, s) = make_node(&e, decoder, scopes.last().unwrap())?;
                stack.push(n);
                scopes.push(s)
            },
            Event::Empty(e) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("font-table XML resource limit exceeded"));
                }
                let (n, _) = make_node(&e, decoder, scopes.last().unwrap())?;
                attach(n, &mut stack, &mut root)?
            },
            Event::End(_) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end tag"))?;
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
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, d)
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
        attrs.push(Attr {
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
    extensions: Vec<FontTableExtensionAttribute>,
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
                extensions.push(FontTableExtensionAttribute {
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

fn parse_table_node(root: &Node) -> Result<FontTable> {
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
    let table = FontTable {
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
                family = Some(FontFamily::parse(
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
                pitch = Some(FontPitch::parse(
                    &Attributes::new(c, &["val"], &[])?.req("val")?,
                )?)
            },
            "sig" => {
                leaf(c)?;
                sig = Some(parse_sig(c)?)
            },
            "embedRegular" => embedded.push(parse_embed(c, EmbeddedFontStyle::Regular)?),
            "embedBold" => embedded.push(parse_embed(c, EmbeddedFontStyle::Bold)?),
            "embedItalic" => embedded.push(parse_embed(c, EmbeddedFontStyle::Italic)?),
            "embedBoldItalic" => embedded.push(parse_embed(c, EmbeddedFontStyle::BoldItalic)?),
            _ => unreachable!(),
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
fn parse_charset(n: &Node) -> Result<Option<FontCharacterSet>> {
    let a = Attributes::new(n, &["val", "characterSet"], &[])?;
    let old = a
        .opt("val")?
        .map(|v| {
            if !(1..=2).contains(&v.len()) || !v.bytes().all(|b| b.is_ascii_hexdigit()) {
                return Err(invalid(format!("invalid charset '{v}'")));
            }
            u8::from_str_radix(&v, 16)
                .map(FontCharacterSet::legacy)
                .map_err(xml_error)
        })
        .transpose()?;
    let strict = a
        .opt("characterSet")?
        .map(|v| FontCharacterSet::strict(&v))
        .transpose()?;
    if old.is_some() && strict.is_some() && old != strict {
        return Err(invalid("conflicting font character sets"));
    }
    Ok(strict.or(old))
}
fn parse_sig(n: &Node) -> Result<FontSignature> {
    let a = Attributes::new(n, &["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"], &[])?;
    let p = |name: &str| -> Result<u32> {
        let v = a.req(name)?;
        if v.len() != 8 || !v.bytes().all(|b| b.is_ascii_hexdigit()) {
            return Err(invalid(format!("invalid font signature '{name}'")));
        }
        u32::from_str_radix(&v, 16).map_err(xml_error)
    };
    Ok(FontSignature {
        unicode_subsets: [p("usb0")?, p("usb1")?, p("usb2")?, p("usb3")?],
        code_pages: [p("csb0")?, p("csb1")?],
    })
}
fn parse_embed(n: &Node, style: EmbeddedFontStyle) -> Result<EmbeddedFont> {
    leaf(n)?;
    let a = Attributes::new(n, &["fontKey", "subsetted"], &["id"])?;
    let key = a.opt("fontKey")?;
    if let Some(v) = &key {
        font_key(v)?
    }
    Ok(EmbeddedFont {
        style,
        relationship_id: a.rel("id")?,
        font_key: key,
        subsetted: a.opt("subsetted")?.map(|v| on_off(&v)).transpose()?,
        resource: None,
        extension_attributes: a.extensions,
    })
}

pub fn write_font_table(t: &FontTable, c: FontTableConformance) -> Result<Vec<u8>> {
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
fn write_font(o: &mut Vec<u8>, f: &Font, c: FontTableConformance) -> Result<()> {
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
            FontTableConformance::Transitional => wa(o, "val", &format!("{:02X}", v.legacy_code())),
            FontTableConformance::Strict => wa(
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
        if let Some(v) = &e.font_key {
            font_key(v)?;
            wa(o, "fontKey", v)
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
fn extensions(o: &mut Vec<u8>, v: &[FontTableExtensionAttribute]) -> Result<()> {
    for a in v {
        preserved(o, a)?
    }
    Ok(())
}
fn preserved(o: &mut Vec<u8>, a: &FontTableExtensionAttribute) -> Result<()> {
    if a.qualified_name.is_empty()
        || a.qualified_name
            .bytes()
            .any(|b| b.is_ascii_whitespace() || matches!(b, b'<' | b'>' | b'=' | b'\'' | b'\"'))
    {
        return Err(invalid("invalid preserved attribute name"));
    }
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

fn extension_namespaces(root: &Node) -> Result<Vec<FontTableExtensionAttribute>> {
    fn walk(
        n: &Node,
        map: &mut HashMap<String, String>,
        out: &mut Vec<FontTableExtensionAttribute>,
    ) -> Result<()> {
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
                out.push(FontTableExtensionAttribute {
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
    for (i, x) in out.iter_mut().enumerate() {
        *x = u8::from_str_radix(&v[i * 2..i * 2 + 2], 16).map_err(xml_error)?
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
fn font_key(v: &str) -> Result<()> {
    let b = v.as_bytes();
    let ok = b.len() == 38
        && b[0] == b'{'
        && b[37] == b'}'
        && [9, 14, 19, 24].iter().all(|i| b[*i] == b'-')
        && b[1..37].iter().enumerate().all(|(i, x)| {
            [8, 13, 18, 23].contains(&i) || x.is_ascii_digit() || (b'A'..=b'F').contains(x)
        });
    if ok {
        Ok(())
    } else {
        Err(invalid(format!("invalid font key '{v}'")))
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
fn xml_error(e: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(e.to_string())
}
fn invalid(e: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(e.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn complete_strict_round_trip() {
        let xml=br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="A&amp;B"><w:altName w:val="Alias"/><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:notTrueType w:val="0"/><w:pitch w:val="variable"/><w:sig w:usb0="E10002FF" w:usb1="4000ACFF" w:usb2="00000009" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/><w:embedRegular r:id="rId1" w:fontKey="{01014A78-CABC-4EF0-12AC-5CD89AEFDE01}" w:subsetted="1"/></w:font></w:fonts>"#;
        let t = parse_font_table(xml).unwrap();
        assert_eq!(t.fonts()[0].name(), "A&B");
        assert_eq!(t.fonts()[0].signature().unwrap().code_pages()[0], 0x19F);
        let strict = t.to_xml(FontTableConformance::Strict).unwrap();
        let s = std::str::from_utf8(&strict).unwrap();
        assert!(s.contains(WS));
        assert!(s.contains("w:characterSet=\"iso-8859-1\""));
        assert_eq!(parse_font_table(&strict).unwrap(), t)
    }
    #[test]
    fn mce_and_real_strict_fixture() {
        let xml=br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:font/></mc:Choice><mc:Fallback><w:font w:name="Fallback"><w:family w:val="roman"/></w:font></mc:Fallback></mc:AlternateContent></w:fonts>"#;
        assert_eq!(parse_font_table(xml).unwrap().fonts()[0].name(), "Fallback");
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let p = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
            root.join("test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx"),
        )
        .unwrap();
        let u = litchi_opc::PackURI::new("/word/fontTable.xml").unwrap();
        let t = parse_font_table(&p.blob_for(&u).unwrap()).unwrap();
        assert!(t.fonts().iter().any(|f| f.name() == "Calibri"));
        assert_eq!(t.fonts()[0].character_set(), Some(FontCharacterSet::Ansi))
    }
    #[test]
    fn malformed_order_and_bounds() {
        for xml in [
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font/></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:family w:val="fantasy"/></w:font></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:panose1 w:val="1234"/></w:font></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:pitch w:val="fixed"/><w:family w:val="roman"/></w:font></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="x"><w:sig w:usb0="0"/></w:font></w:fonts>"#,
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="x"><w:embedRegular r:id="rId1" w:fontKey="bad"/></w:font></w:fonts>"#,
            r#"<!DOCTYPE x><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
            r#"<?bad x?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        ] {
            assert!(parse_font_table(xml.as_bytes()).is_err(), "{xml}")
        }
        assert!(parse_font_table(&vec![b' '; MAX_XML + 1]).is_err())
    }
    #[test]
    fn real_poi_embedded_resources_are_inert() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let p = crate::docx::Package::open(
            root.join("test-data/poi/test-data/document/saut_page.docx"),
        )
        .unwrap();
        let t = p.font_table().unwrap().unwrap();
        assert_eq!(t.fonts().len(), 7);
        let e: Vec<_> = t.fonts().iter().flat_map(Font::embedded_fonts).collect();
        assert_eq!(e.len(), 20);
        assert_eq!(
            e.iter()
                .map(|v| v.relationship_id())
                .collect::<HashSet<_>>()
                .len(),
            16
        );
        assert!(e.iter().all(|v| v.resource().is_some()));
        assert!(
            e.iter()
                .all(|v| v.resource().unwrap().content_type() == FONT_CT)
        )
    }

    #[test]
    fn guid_obfuscation_is_reversible_and_fs_type_is_validated() {
        let key = "{00112233-4455-6677-8899-AABBCCDDEEFF}";
        let original = (0u8..64).collect::<Vec<_>>();
        let mut data = original.clone();
        obfuscate_embedded_font_data(&mut data, key).unwrap();
        assert_ne!(data, original);
        assert_eq!(&data[32..], &original[32..]);
        deobfuscate_embedded_font_data(&mut data, key).unwrap();
        assert_eq!(data, original);
        assert!(obfuscate_embedded_font_data(&mut [0; 31], key).is_err());
        assert!(obfuscate_embedded_font_data(&mut [0; 32], "bad").is_err());

        assert!(
            EmbeddedFontLicensing::from_fs_type(0)
                .unwrap()
                .installable()
        );
        let editable = EmbeddedFontLicensing::from_fs_type(0x0108).unwrap();
        assert!(editable.editable && editable.no_subsetting);
        assert!(EmbeddedFontLicensing::from_fs_type(0x0006).is_err());
        assert!(EmbeddedFontLicensing::from_fs_type(0x8000).is_err());
    }

    #[test]
    fn generated_shared_font_crud_and_extensions_remain_inert() {
        let mut package = OpcPackage::new();
        let document_name = PackURI::new("/word/document.xml").unwrap();
        package.add_part(Box::new(XmlPart::new(
            document_name,
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

        let key = "{00112233-4455-6677-8899-AABBCCDDEEFF}";
        let resource =
            EmbeddedFontResource::new("/word/fonts/shared.odttf", (0u8..64).collect::<Vec<_>>());
        let mut first = Font::new("Alpha")
            .with_alternate_name("Alpha Alt")
            .with_panose([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
            .with_character_set(FontCharacterSet::Ansi)
            .with_family(FontFamily::Swiss)
            .with_pitch(FontPitch::Variable)
            .with_signature(FontSignature::new([1, 2, 3, 4], [5, 6]))
            .with_embedded_font(
                EmbeddedFont::new(EmbeddedFontStyle::Regular, resource.clone())
                    .with_relationship_id("rIdFont1")
                    .with_font_key(key)
                    .with_subsetted(true),
            );
        first
            .extension_attributes_mut()
            .push(FontTableExtensionAttribute::new("x:flag", "kept"));
        let mut table = FontTable::new();
        table
            .namespaces_mut()
            .push(FontTableExtensionAttribute::new(
                "xmlns:x",
                "urn:test-fonts",
            ));
        table.fonts_mut().push(first);
        store_font_table(&mut package, &table, FontTableConformance::Transitional).unwrap();

        add_font(
            &mut package,
            Font::new("Beta").with_embedded_font(
                EmbeddedFont::new(EmbeddedFontStyle::Regular, resource).with_font_key(key),
            ),
            FontTableConformance::Transitional,
        )
        .unwrap();
        let loaded = load_font_table(&package).unwrap().unwrap();
        assert_eq!(loaded.fonts().len(), 2);
        assert_eq!(loaded.fonts()[0].extension_attributes()[0].value(), "kept");
        let relationships = loaded
            .fonts()
            .iter()
            .map(|font| font.embedded_fonts()[0].relationship_id())
            .collect::<HashSet<_>>();
        assert_eq!(relationships.len(), 2);
        assert_eq!(
            loaded.fonts()[0].embedded_fonts()[0]
                .resource()
                .unwrap()
                .part_name(),
            loaded.fonts()[1].embedded_fonts()[0]
                .resource()
                .unwrap()
                .part_name()
        );

        reorder_fonts(
            &mut package,
            &["Beta".into(), "Alpha".into()],
            FontTableConformance::Transitional,
        )
        .unwrap();
        assert!(remove_font(&mut package, "Alpha", FontTableConformance::Transitional).unwrap());
        let shared = PackURI::new("/word/fonts/shared.odttf").unwrap();
        assert!(package.get_part(&shared).is_ok());
        assert!(find_font(&package, "beta").unwrap().is_some());
    }

    #[test]
    fn word_font_name_limits_are_rejected_but_fixture_duplicates_round_trip() {
        let mut table = FontTable::new();
        table.fonts_mut().push(Font::new("A"));
        table.fonts_mut().push(Font::new("a"));
        assert!(write_font_table(&table, FontTableConformance::Transitional).is_ok());
        assert!(parse_font_table(
            br#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="12345678901234567890123456789012"/></w:fonts>"#
        )
        .is_err());
    }
}
