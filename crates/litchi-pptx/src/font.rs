//! Typed PresentationML embedded fonts and inert OPC resources.

use crate::error::{Error, Result};
use caseless::Caseless;
use litchi_ooxml_common::mce::{
    ActiveOffsetLimits, MceCapabilities, MceLimits, active_offsets, process_ooxml,
};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;
use unicode_normalization::UnicodeNormalization;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FONT_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const STRICT_FONT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/font";
#[cfg(test)]
const PRESENTATION_CT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const FONT_DATA_CT: &str = "application/x-fontdata";
const FONT_TTF_CT: &str = "application/x-font-ttf";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 256;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_FONTS: usize = 4096;
const MAX_FONT_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_FONT_BYTES: usize = 256 * 1024 * 1024;
const MAX_MCE_MARKED_BYTES: usize = MAX_XML_BYTES + MAX_NODES * 64;

/// Validated OpenType `OS/2.fsType` embedding metadata.
///
/// This value is supplied by callers or a separate font-metadata reader. This
/// module never searches, parses, loads, or executes a font program.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Permission {
    Installable,
    Restricted,
    PreviewPrint,
    Editable,
}

bitflags::bitflags! {
    /// Independent OpenType embedding restrictions.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Restrictions: u16 {
        const NO_SUBSETTING = 0x0100;
        const BITMAP_ONLY = 0x0200;
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct License {
    fs_type: u16,
    permission: Permission,
    restrictions: Restrictions,
}

impl License {
    /// Validate the defined embedding bits and reject contradictory modes.
    pub fn from_fs_type(fs_type: u16) -> Result<Self> {
        const DEFINED: u16 = 0x0002 | 0x0004 | 0x0008 | 0x0100 | 0x0200;
        if fs_type & !DEFINED != 0 {
            return Err(invalid(format!(
                "font fsType contains reserved bits 0x{:04X}",
                fs_type & !DEFINED
            )));
        }
        let modes = [0x0002, 0x0004, 0x0008]
            .into_iter()
            .filter(|bit| fs_type & *bit != 0)
            .count();
        if modes > 1 {
            return Err(invalid(
                "font fsType has contradictory restricted, preview/print, and editable modes",
            ));
        }
        let permission = if fs_type & 0x0002 != 0 {
            Permission::Restricted
        } else if fs_type & 0x0004 != 0 {
            Permission::PreviewPrint
        } else if fs_type & 0x0008 != 0 {
            Permission::Editable
        } else {
            Permission::Installable
        };
        Ok(Self {
            fs_type,
            permission,
            restrictions: Restrictions::from_bits_retain(fs_type & 0x0300),
        })
    }

    /// Return the original validated `fsType` bits.
    pub fn fs_type(self) -> u16 {
        self.fs_type
    }

    /// Return the mutually exclusive embedding permission.
    pub fn permission(self) -> Permission {
        self.permission
    }

    /// Return compact independent restrictions.
    pub fn restrictions(self) -> Restrictions {
        self.restrictions
    }

    /// Report installable embedding permission.
    pub fn installable(self) -> bool {
        self.permission == Permission::Installable
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    fn pml(self) -> &'static str {
        match self {
            Self::Transitional => PML,
            Self::Strict => STRICT_PML,
        }
    }
    fn rel_ns(self) -> &'static str {
        match self {
            Self::Transitional => REL_NS,
            Self::Strict => STRICT_REL_NS,
        }
    }
    fn font_rel(self) -> &'static str {
        match self {
            Self::Transitional => FONT_REL,
            Self::Strict => STRICT_FONT_REL,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Style {
    Regular,
    Bold,
    Italic,
    BoldItalic,
}

impl Style {
    fn element(self) -> &'static str {
        match self {
            Self::Regular => "regular",
            Self::Bold => "bold",
            Self::Italic => "italic",
            Self::BoldItalic => "boldItalic",
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
    fn parse_raw(value: &str) -> Option<Self> {
        match value {
            "regular" => Some(Self::Regular),
            "bold" => Some(Self::Bold),
            "italic" => Some(Self::Italic),
            "boldItalic" => Some(Self::BoldItalic),
            _ => None,
        }
    }
}

/// Font-pitch component of DrawingML `ST_PitchFamily`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Pitch {
    Default,
    Fixed,
    Variable,
}

impl Pitch {
    const fn wire(self) -> u8 {
        match self {
            Self::Default => 0,
            Self::Fixed => 1,
            Self::Variable => 2,
        }
    }
}

/// Font-family component of DrawingML `ST_PitchFamily`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Family {
    None,
    Roman,
    Swiss,
    Modern,
    Script,
    Decorative,
}

impl Family {
    const fn wire(self) -> u8 {
        match self {
            Self::None => 0,
            Self::Roman => 1,
            Self::Swiss => 2,
            Self::Modern => 3,
            Self::Script => 4,
            Self::Decorative => 5,
        }
    }
}

/// The closed 18-value DrawingML pitch/family domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PitchFamily {
    pitch: Pitch,
    family: Family,
}

impl PitchFamily {
    /// Combine typed pitch and family values; every combination is valid.
    pub const fn new(pitch: Pitch, family: Family) -> Self {
        Self { pitch, family }
    }

    /// Return the typed pitch component.
    pub const fn pitch(self) -> Pitch {
        self.pitch
    }

    /// Return the typed family component.
    pub const fn family(self) -> Family {
        self.family
    }

    const fn wire(self) -> u8 {
        self.family.wire() * 16 + self.pitch.wire()
    }

    fn from_wire(value: u8) -> Result<Self> {
        let pitch = match value & 0x0F {
            0 => Pitch::Default,
            1 => Pitch::Fixed,
            2 => Pitch::Variable,
            _ => return Err(invalid(format!("invalid pitchFamily value '{value}'"))),
        };
        let family = match value >> 4 {
            0 => Family::None,
            1 => Family::Roman,
            2 => Family::Swiss,
            3 => Family::Modern,
            4 => Family::Script,
            5 => Family::Decorative,
            _ => return Err(invalid(format!("invalid pitchFamily value '{value}'"))),
        };
        Ok(Self { pitch, family })
    }
}

/// The fixed ten-byte PANOSE classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Panose([u8; 10]);

impl Panose {
    /// Construct a PANOSE value from its fixed-size classification bytes.
    pub const fn new(bytes: [u8; 10]) -> Self {
        Self(bytes)
    }

    /// Borrow the ten classification bytes.
    pub const fn bytes(&self) -> &[u8; 10] {
        &self.0
    }

    /// Move out the ten classification bytes.
    pub const fn into_bytes(self) -> [u8; 10] {
        self.0
    }
}

impl From<[u8; 10]> for Panose {
    fn from(value: [u8; 10]) -> Self {
        Self::new(value)
    }
}

/// A Windows font charset code with private PresentationML wire conversion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Charset(u8);

impl Charset {
    pub const ANSI: Self = Self(0);
    pub const DEFAULT: Self = Self(1);
    pub const SYMBOL: Self = Self(2);
    pub const MACINTOSH: Self = Self(77);
    pub const SHIFT_JIS: Self = Self(128);
    pub const HANGEUL: Self = Self(129);
    pub const JOHAB: Self = Self(130);
    pub const GB2312: Self = Self(134);
    pub const CHINESE_BIG5: Self = Self(136);
    pub const GREEK: Self = Self(161);
    pub const TURKISH: Self = Self(162);
    pub const VIETNAMESE: Self = Self(163);
    pub const HEBREW: Self = Self(177);
    pub const ARABIC: Self = Self(178);
    pub const BALTIC: Self = Self(186);
    pub const RUSSIAN: Self = Self(204);
    pub const THAI: Self = Self(222);
    pub const EAST_EUROPE: Self = Self(238);
    pub const OEM: Self = Self(255);

    /// Preserve any Windows charset byte, including producer-defined values.
    pub const fn new(code: u8) -> Self {
        Self(code)
    }

    /// Return the Windows charset code without a string conversion.
    pub const fn code(self) -> u8 {
        self.0
    }

    const fn from_wire(value: i8) -> Self {
        Self(value as u8)
    }

    const fn wire(self) -> i8 {
        self.0 as i8
    }
}

/// Physical font representation permitted by PresentationML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Format {
    /// PowerPoint-compatible storage (`application/x-fontdata`).
    PowerPoint,
    /// Standards-only `application/x-font-ttf` preservation.
    Standard,
}

impl Format {
    fn content_type(self) -> &'static str {
        match self {
            Self::PowerPoint => FONT_DATA_CT,
            Self::Standard => FONT_TTF_CT,
        }
    }

    fn parse(content_type: &str) -> Result<Self> {
        match content_type {
            FONT_DATA_CT => Ok(Self::PowerPoint),
            FONT_TTF_CT => Ok(Self::Standard),
            _ => Err(invalid(format!(
                "unsupported embedded-font content type '{content_type}'"
            ))),
        }
    }
}

/// One immutable, cheaply shared inert font program.
///
/// Moving a `Vec<u8>` into [`Data::new`] adopts its allocation. Cloning this
/// value only increments an `Arc`; it never copies the font program.
#[derive(Debug, Clone)]
pub struct Data {
    format: Format,
    bytes: Arc<Vec<u8>>,
}

impl Data {
    /// Adopt and validate an owned font container.
    pub fn new(bytes: Vec<u8>, format: Format) -> Result<Self> {
        Self::shared(Arc::new(bytes), format)
    }

    /// Validate and share an existing immutable allocation without copying it.
    pub fn shared(bytes: Arc<Vec<u8>>, format: Format) -> Result<Self> {
        validate_font_bytes(&bytes)?;
        match format {
            Format::PowerPoint => validate_eot(&bytes)?,
            Format::Standard => validate_sfnt(&bytes)?,
        }
        Ok(Self { format, bytes })
    }

    /// Adopt a PowerPoint-compatible EOT/MTX container.
    pub fn powerpoint(bytes: Vec<u8>) -> Result<Self> {
        Self::new(bytes, Format::PowerPoint)
    }

    /// Preserve standards-only `application/x-font-ttf` storage explicitly.
    pub fn standard(bytes: Vec<u8>) -> Result<Self> {
        Self::new(bytes, Format::Standard)
    }

    /// Preserve a bounded producer payload already present in a loaded package.
    fn preserve(bytes: Arc<Vec<u8>>, format: Format) -> Result<Self> {
        validate_font_bytes(&bytes)?;
        Ok(Self { format, bytes })
    }

    /// Return the physical representation.
    pub fn format(&self) -> Format {
        self.format
    }

    /// Borrow the inert bytes.
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Move out the shared allocation without copying it.
    pub fn into_shared(self) -> Arc<Vec<u8>> {
        self.bytes
    }
}

impl PartialEq for Data {
    fn eq(&self, other: &Self) -> bool {
        self.format == other.format
            && (Arc::ptr_eq(&self.bytes, &other.bytes) || self.bytes == other.bytes)
    }
}

impl Eq for Data {}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Source {
    relationship_id: String,
    part_name: String,
}

/// One typed style face and its inert program.
#[derive(Debug, Clone)]
pub struct Face {
    style: Style,
    data: Data,
    source: Option<Source>,
}

impl Face {
    /// Pair a typed face style with an owned or shared font program.
    pub fn new(style: Style, data: Data) -> Self {
        Self {
            style,
            data,
            source: None,
        }
    }

    /// Return the schema-level style.
    pub fn style(&self) -> Style {
        self.style
    }

    /// Borrow the inert font program.
    pub fn data(&self) -> &Data {
        &self.data
    }

    /// Replace the program by move, returning the previous allocation.
    pub fn set(&mut self, data: Data) -> Data {
        self.source = None;
        std::mem::replace(&mut self.data, data)
    }
}

impl PartialEq for Face {
    fn eq(&self, other: &Self) -> bool {
        self.style == other.style && self.data == other.data
    }
}

impl Eq for Face {}

/// One invariant-bearing embedded typeface.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    typeface: String,
    key: String,
    panose: Option<Panose>,
    pitch_family: Option<PitchFamily>,
    charset: Option<Charset>,
    faces: Vec<Face>,
}

impl Font {
    /// Construct a typeface. The schema permits a descriptor with no faces;
    /// add one concisely with [`Font::with`] or [`Font::put`].
    pub fn new(typeface: impl Into<String>) -> Result<Self> {
        let typeface = typeface.into();
        validate_typeface(&typeface)?;
        Ok(Self {
            key: name_key(&typeface),
            typeface,
            panose: None,
            pitch_family: None,
            charset: None,
            faces: Vec::new(),
        })
    }

    /// Construct a typeface with one face in a single expression.
    pub fn from_face(typeface: impl Into<String>, face: Face) -> Result<Self> {
        let mut font = Self::new(typeface)?;
        font.put(face)?;
        Ok(font)
    }

    /// Add a face with builder-style chaining.
    pub fn with(mut self, face: Face) -> Result<Self> {
        self.put(face)?;
        Ok(self)
    }

    /// Return the producer spelling of the typeface.
    pub fn name(&self) -> &str {
        &self.typeface
    }

    /// Return the optional ten-byte PANOSE classification.
    pub fn panose(&self) -> Option<Panose> {
        self.panose
    }

    /// Return the optional combined pitch/family byte.
    pub fn pitch_family(&self) -> Option<PitchFamily> {
        self.pitch_family
    }

    /// Return the optional Windows character-set byte.
    pub fn charset(&self) -> Option<Charset> {
        self.charset
    }

    /// Return faces in the schema-defined style order.
    pub fn faces(&self) -> &[Face] {
        &self.faces
    }

    /// Set the PANOSE classification with struct-update-like builder syntax.
    pub fn with_panose(mut self, value: impl Into<Panose>) -> Self {
        self.panose = Some(value.into());
        self
    }

    /// Replace or clear the PANOSE classification.
    pub fn set_panose(&mut self, value: Option<Panose>) -> Option<Panose> {
        std::mem::replace(&mut self.panose, value)
    }

    /// Set the combined pitch/family byte.
    pub fn with_pitch_family(mut self, value: PitchFamily) -> Self {
        self.pitch_family = Some(value);
        self
    }

    /// Replace or clear the compact pitch/family classification.
    pub fn set_pitch_family(&mut self, value: Option<PitchFamily>) -> Option<PitchFamily> {
        std::mem::replace(&mut self.pitch_family, value)
    }

    /// Set the Windows character-set byte.
    pub fn with_charset(mut self, value: Charset) -> Self {
        self.charset = Some(value);
        self
    }

    /// Replace or clear the Windows charset code.
    pub fn set_charset(&mut self, value: Option<Charset>) -> Option<Charset> {
        std::mem::replace(&mut self.charset, value)
    }

    /// Add or replace one typed style face, returning the previous face.
    pub fn put(&mut self, mut face: Face) -> Result<Option<Face>> {
        face.source = None;
        if let Some(index) = self.faces.iter().position(|item| item.style == face.style) {
            let len = self.faces.len();
            let previous = std::mem::replace(
                self.faces
                    .get_mut(index)
                    .ok_or(Error::FontIndexOutOfBounds { index, len })?,
                face,
            );
            return Ok(Some(previous));
        }
        if self.faces.len() == 4 {
            return Err(invalid("embedded font already has all four style faces"));
        }
        self.faces
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "embedded-font faces",
                source,
            })?;
        self.faces.push(face);
        self.faces.sort_by_key(|item| item.style.rank());
        Ok(None)
    }

    /// Select one face by its typed style.
    pub fn get(&self, style: Style) -> Option<&Face> {
        self.faces.iter().find(|face| face.style == style)
    }

    /// Remove and return one face. Face-less descriptors remain schema-valid.
    pub fn remove(&mut self, style: Style) -> Result<Face> {
        let index = self
            .faces
            .iter()
            .position(|face| face.style == style)
            .ok_or_else(|| invalid(format!("embedded font has no {style:?} face")))?;
        Ok(self.faces.remove(index))
    }

    /// Rename a detached font. A containing [`Fonts`] rechecks uniqueness when
    /// the value is inserted or replaced.
    pub fn rename(&mut self, typeface: impl Into<String>) -> Result<String> {
        let typeface = typeface.into();
        validate_typeface(&typeface)?;
        let key = name_key(&typeface);
        let previous = std::mem::replace(&mut self.typeface, typeface);
        self.key = key;
        Ok(previous)
    }
}

/// Semantic font name or checked source-order position.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Key<'a> {
    /// Unicode-caseless typeface spelling.
    Name(&'a str),
    /// Zero-based source-order position for repair and deterministic ordering.
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

/// A detached, source-ordered embedded-font collection.
#[derive(Debug, Clone, Copy)]
struct NameIndex {
    first: usize,
    matches: usize,
}

#[derive(Debug, Clone)]
pub struct Fonts {
    fonts: Vec<Font>,
    by_name: HashMap<String, NameIndex>,
}

impl Fonts {
    /// Construct an empty collection.
    pub fn new() -> Self {
        Self {
            fonts: Vec::new(),
            by_name: HashMap::new(),
        }
    }

    /// Return the number of typefaces.
    pub fn len(&self) -> usize {
        self.fonts.len()
    }

    /// Report whether no fonts are embedded.
    pub fn is_empty(&self) -> bool {
        self.fonts.is_empty()
    }

    /// Iterate in source order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Font> {
        self.fonts.iter()
    }

    /// Select by semantic typeface or checked numeric position.
    pub fn get<'a, 'k>(&'a self, key: impl Into<Key<'k>>) -> Result<&'a Font> {
        let index = self.offset(key.into())?;
        self.fonts.get(index).ok_or(Error::FontIndexOutOfBounds {
            index,
            len: self.fonts.len(),
        })
    }

    /// Append an owned typeface without copying its programs.
    pub fn add(&mut self, font: Font) -> Result<()> {
        self.ensure_unique(&font, None)?;
        if self.fonts.len() == MAX_FONTS {
            return Err(limit("embedded fonts"));
        }
        let next_index = build_name_index(
            self.fonts.len() + 1,
            self.fonts
                .iter()
                .enumerate()
                .chain(std::iter::once((self.fonts.len(), &font))),
        )?;
        self.fonts
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "embedded-font collection",
                source,
            })?;
        self.fonts.push(font);
        self.by_name = next_index;
        Ok(())
    }

    /// Replace a selected typeface and return the previous owned value.
    pub fn replace<'k>(&mut self, key: impl Into<Key<'k>>, font: Font) -> Result<Font> {
        let index = self.offset(key.into())?;
        self.ensure_unique(&font, Some(index))?;
        let next_index = build_name_index(
            self.fonts.len(),
            self.fonts.iter().enumerate().map(|(position, current)| {
                (position, if position == index { &font } else { current })
            }),
        )?;
        let len = self.fonts.len();
        let previous = std::mem::replace(
            self.fonts
                .get_mut(index)
                .ok_or(Error::FontIndexOutOfBounds { index, len })?,
            font,
        );
        self.by_name = next_index;
        Ok(previous)
    }

    /// Remove and return a selected typeface.
    pub fn remove<'k>(&mut self, key: impl Into<Key<'k>>) -> Result<Font> {
        let index = self.offset(key.into())?;
        let next_index = build_name_index(
            self.fonts.len().saturating_sub(1),
            self.fonts
                .iter()
                .enumerate()
                .filter(|(position, _)| *position != index)
                .map(|(_, font)| font)
                .enumerate(),
        )?;
        let removed = self.fonts.remove(index);
        self.by_name = next_index;
        Ok(removed)
    }

    /// Apply a complete checked permutation without cloning font programs.
    pub fn reorder<'k, K>(&mut self, order: &'k [K]) -> Result<()>
    where
        K: Copy + Into<Key<'k>>,
    {
        if order.len() != self.fonts.len() {
            return Err(Error::FontOrderLength {
                expected: self.fonts.len(),
                actual: order.len(),
            });
        }
        let mut ranks = Vec::new();
        ranks
            .try_reserve_exact(self.fonts.len())
            .map_err(|source| Error::Allocation {
                resource: "embedded-font reorder",
                source,
            })?;
        ranks.resize(self.fonts.len(), usize::MAX);
        for (rank, key) in order.iter().copied().enumerate() {
            let index = self.offset(key.into())?;
            let len = ranks.len();
            let slot = ranks
                .get_mut(index)
                .ok_or(Error::FontIndexOutOfBounds { index, len })?;
            if *slot != usize::MAX {
                return Err(Error::DuplicateFontSelection { index });
            }
            *slot = rank;
        }
        let next_index = build_name_index(
            self.fonts.len(),
            self.fonts
                .iter()
                .zip(&ranks)
                .map(|(font, rank)| (*rank, font)),
        )?;
        for position in 0..ranks.len() {
            while ranks[position] != position {
                let target = ranks[position];
                self.fonts.swap(position, target);
                ranks.swap(position, target);
            }
        }
        self.by_name = next_index;
        Ok(())
    }

    /// Move all fonts out in source order.
    pub fn into_fonts(self) -> Vec<Font> {
        self.fonts
    }

    fn offset(&self, key: Key<'_>) -> Result<usize> {
        match key {
            Key::Name(name) => match self.by_name.get(&name_key(name)).copied() {
                Some(NameIndex { first, matches: 1 }) => Ok(first),
                Some(NameIndex { matches, .. }) => Err(Error::AmbiguousFontName {
                    name: name.into(),
                    matches,
                }),
                None => Err(Error::FontNotFound(name.into())),
            },
            Key::Index(index) if index < self.fonts.len() => Ok(index),
            Key::Index(index) => Err(Error::FontIndexOutOfBounds {
                index,
                len: self.fonts.len(),
            }),
        }
    }

    fn ensure_unique(&self, font: &Font, replaced: Option<usize>) -> Result<()> {
        let matches = self
            .fonts
            .iter()
            .enumerate()
            .filter(|(index, item)| Some(*index) != replaced && item.key == font.key)
            .count();
        if matches == 0 {
            Ok(())
        } else {
            Err(Error::DuplicateFontName {
                name: font.typeface.clone(),
                matches,
            })
        }
    }

    fn reindex(&mut self) -> Result<()> {
        let next_index = build_name_index(self.fonts.len(), self.fonts.iter().enumerate())?;
        self.by_name = next_index;
        Ok(())
    }
}

fn build_name_index<'a>(
    len: usize,
    fonts: impl IntoIterator<Item = (usize, &'a Font)>,
) -> Result<HashMap<String, NameIndex>> {
    let mut index = HashMap::new();
    index.try_reserve(len).map_err(|source| Error::Allocation {
        resource: "embedded-font name index",
        source,
    })?;
    for (position, font) in fonts {
        index
            .entry(font.key.clone())
            .and_modify(|entry: &mut NameIndex| {
                entry.first = entry.first.min(position);
                entry.matches = entry.matches.saturating_add(1);
            })
            .or_insert(NameIndex {
                first: position,
                matches: 1,
            });
    }
    Ok(index)
}

impl Default for Fonts {
    fn default() -> Self {
        Self::new()
    }
}

impl PartialEq for Fonts {
    fn eq(&self, other: &Self) -> bool {
        self.fonts == other.fonts
    }
}

impl Eq for Fonts {}

impl IntoIterator for Fonts {
    type Item = Font;
    type IntoIter = std::vec::IntoIter<Font>;

    fn into_iter(self) -> Self::IntoIter {
        self.fonts.into_iter()
    }
}

impl<'a> IntoIterator for &'a Fonts {
    type Item = &'a Font;
    type IntoIter = std::slice::Iter<'a, Font>;

    fn into_iter(self) -> Self::IntoIter {
        self.fonts.iter()
    }
}

#[derive(Debug, Clone)]
struct RawResource {
    part_name: String,
    content_type: String,
    /// The font program is deliberately retained as inert bytes.
    data: Arc<Vec<u8>>,
}

impl PartialEq for RawResource {
    fn eq(&self, other: &Self) -> bool {
        self.part_name == other.part_name
            && self.content_type == other.content_type
            && (Arc::ptr_eq(&self.data, &other.data) || self.data == other.data)
    }
}

impl Eq for RawResource {}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawFace {
    style: Style,
    relationship_id: String,
    /// Present after package loading and required for package storage.
    resource: Option<RawResource>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawFont {
    has_descriptor: bool,
    typeface: String,
    panose: Option<Panose>,
    pitch_family: Option<PitchFamily>,
    charset: Option<Charset>,
    faces: Vec<RawFace>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct RawFonts {
    fonts: Vec<RawFont>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Presentation,
    List,
    Font(usize),
    Leaf,
    Other,
}

struct ParsedPresentation {
    conformance: Conformance,
    value: Option<RawFonts>,
}

/// Parse the optional embedded-font markup from a complete presentation part.
#[cfg(test)]
fn parse_raw(xml: &[u8]) -> Result<Option<RawFonts>> {
    Ok(parse_presentation(xml)?.value)
}

fn parse_presentation(xml: &[u8]) -> Result<ParsedPresentation> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("presentation XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_XML_BYTES {
        return Err(limit("MCE-processed presentation XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Context> = Vec::new();
    let mut fonts = Vec::new();
    let mut saw_root = false;
    let mut saw_list = false;
    let mut conformance = None;
    let mut root_rank = None;
    let mut nodes = 0usize;
    let mut string_bytes = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let empty_event = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = empty_event;
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                let namespace =
                    resolved_namespace(reader.resolver().resolve_element(element.name()).0)?;
                let local = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let parent = stack.last().copied();
                if parent == Some(Context::Presentation)
                    && namespace == conformance.map(Conformance::pml).unwrap_or_default()
                    && let Some(rank) = presentation_child_rank(&local)
                {
                    if root_rank.is_some_and(|previous| previous > rank) {
                        return Err(invalid(format!(
                            "presentation child '{local}' is out of schema order"
                        )));
                    }
                    root_rank = Some(rank);
                }
                let context = if parent.is_none() {
                    if saw_root || local != "presentation" {
                        return Err(invalid("expected one PresentationML presentation root"));
                    }
                    let c = match namespace.as_str() {
                        PML => Conformance::Transitional,
                        STRICT_PML => Conformance::Strict,
                        _ => return Err(invalid("presentation root has an unsupported namespace")),
                    };
                    saw_root = true;
                    conformance = Some(c);
                    Context::Presentation
                } else if parent == Some(Context::Presentation)
                    && namespace == conformance.map(Conformance::pml).unwrap_or_default()
                    && local == "embeddedFontLst"
                {
                    if saw_list {
                        return Err(invalid(
                            "presentation has multiple embeddedFontLst elements",
                        ));
                    }
                    reject_unqualified_attributes(
                        &reader,
                        element,
                        reader.decoder(),
                        &[],
                        &mut string_bytes,
                    )?;
                    saw_list = true;
                    Context::List
                } else if parent == Some(Context::List) {
                    if namespace != conformance.map(Conformance::pml).unwrap_or_default()
                        || local != "embeddedFont"
                    {
                        return Err(invalid("embeddedFontLst contains a non-embeddedFont child"));
                    }
                    reject_unqualified_attributes(
                        &reader,
                        element,
                        reader.decoder(),
                        &[],
                        &mut string_bytes,
                    )?;
                    if fonts.len() >= MAX_FONTS {
                        return Err(limit("embedded fonts"));
                    }
                    fonts.push(RawFont {
                        has_descriptor: false,
                        typeface: String::new(),
                        panose: None,
                        pitch_family: None,
                        charset: None,
                        faces: Vec::new(),
                    });
                    Context::Font(fonts.len() - 1)
                } else if let Some(Context::Font(index)) = parent {
                    if namespace != conformance.map(Conformance::pml).unwrap_or_default() {
                        return Err(invalid("embeddedFont contains a foreign child"));
                    }
                    if local == "font" {
                        if fonts[index].has_descriptor {
                            return Err(invalid("embeddedFont has multiple font descriptors"));
                        }
                        if !fonts[index].faces.is_empty() {
                            return Err(invalid(
                                "embeddedFont descriptor must precede every style face",
                            ));
                        }
                        parse_descriptor(
                            &reader,
                            element,
                            reader.decoder(),
                            &mut fonts[index],
                            &mut string_bytes,
                        )?;
                        fonts[index].has_descriptor = true;
                        Context::Leaf
                    } else if let Some(style) = Style::parse_raw(&local) {
                        if !fonts[index].has_descriptor {
                            return Err(invalid(
                                "embeddedFont descriptor must precede every style face",
                            ));
                        }
                        let relationship_id = parse_face(
                            &reader,
                            element,
                            reader.decoder(),
                            conformance.ok_or_else(|| invalid("missing presentation profile"))?,
                            &mut string_bytes,
                        )?;
                        if fonts[index].faces.iter().any(|face| face.style == style) {
                            return Err(invalid(format!(
                                "duplicate embedded-font style '{local}'"
                            )));
                        }
                        if fonts[index]
                            .faces
                            .last()
                            .is_some_and(|face| face.style.rank() >= style.rank())
                        {
                            return Err(invalid("embedded-font styles are out of schema order"));
                        }
                        fonts[index].faces.push(RawFace {
                            style,
                            relationship_id,
                            resource: None,
                        });
                        Context::Leaf
                    } else {
                        return Err(invalid(format!("unexpected embeddedFont child '{local}'")));
                    }
                } else if matches!(parent, Some(Context::List | Context::Leaf)) {
                    return Err(invalid("embedded-font leaf element contains child content"));
                } else {
                    Context::Other
                };
                stack.push(context);
                if empty {
                    let ended = stack
                        .pop()
                        .ok_or_else(|| invalid("missing empty-element context"))?;
                    if let Context::Font(index) = ended {
                        finish_font(&fonts[index])?;
                    }
                }
            },
            Event::End(_) => {
                let ended = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                if let Context::Font(index) = ended {
                    finish_font(&fonts[index])?;
                }
            },
            Event::Text(text) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    let value = text.decode().map_err(xml_error)?;
                    if !value.trim().is_empty() {
                        return Err(invalid("embedded-font markup contains text"));
                    }
                } else if stack.is_empty() && !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("text occurs outside the presentation root"));
                }
            },
            Event::CData(_) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    return Err(invalid("CDATA is rejected in embedded-font markup"));
                }
            },
            Event::GeneralRef(_) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    return Err(invalid(
                        "entity references are rejected in embedded-font markup",
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation XML"));
    }
    let conformance = conformance.ok_or_else(|| invalid("missing presentation root"))?;
    let value = saw_list.then_some(RawFonts { fonts });
    if let Some(value) = &value {
        validate_value(value, false)?;
    }
    Ok(ParsedPresentation { conformance, value })
}

fn parse_descriptor(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    font: &mut RawFont,
    strings: &mut usize,
) -> Result<()> {
    let attrs = collect_unqualified_attributes(
        reader,
        element,
        decoder,
        &["typeface", "panose", "pitchFamily", "charset"],
        strings,
    )?;
    font.typeface = attrs
        .get("typeface")
        .cloned()
        .ok_or_else(|| invalid("font descriptor is missing typeface"))?;
    font.panose = attrs
        .get("panose")
        .map(|value| parse_panose(value))
        .transpose()?;
    font.pitch_family = attrs
        .get("pitchFamily")
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_| invalid(format!("invalid pitchFamily value '{value}'")))
                .and_then(PitchFamily::from_wire)
        })
        .transpose()?;
    font.charset = attrs
        .get("charset")
        .map(|value| {
            value
                .parse::<i8>()
                .map_err(|_| invalid(format!("invalid charset byte value '{value}'")))
                .map(Charset::from_wire)
        })
        .transpose()?;
    Ok(())
}

fn parse_face(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    conformance: Conformance,
    strings: &mut usize,
) -> Result<String> {
    let mut id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified = attribute.key.as_ref();
        if qualified == b"xmlns" || qualified.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != b"id"
            || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.rel_ns().as_bytes())
        {
            return Err(invalid("embedded-font face has an unexpected attribute"));
        }
        if id.is_some() {
            return Err(invalid("embedded-font face has duplicate relationship IDs"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_string_bytes(strings, value.len())?;
        id = Some(value);
    }
    let id = id
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("embedded-font face is missing r:id"))?;
    validate_relationship_id(&id)?;
    Ok(id)
}

fn collect_unqualified_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&str],
    strings: &mut usize,
) -> Result<HashMap<String, String>> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified = attribute.key.as_ref();
        if qualified == b"xmlns" || qualified.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        if namespace != ResolveResult::Unbound || !allowed.contains(&local) {
            return Err(invalid(format!("unexpected attribute '{local}'")));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_string_bytes(strings, local.len() + value.len())?;
        if result.insert(local.to_owned(), value).is_some() {
            return Err(invalid(format!("duplicate attribute '{local}'")));
        }
    }
    Ok(result)
}

fn reject_unqualified_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&str],
    strings: &mut usize,
) -> Result<()> {
    collect_unqualified_attributes(reader, element, decoder, allowed, strings).map(|_| ())
}

/// Deterministically serializes a self-contained `p:embeddedFontLst` fragment.
fn write_raw(value: &RawFonts, conformance: Conformance) -> Result<Vec<u8>> {
    validate_value(value, false)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<p:embeddedFontLst xmlns:p=\"");
    escape(&mut output, conformance.pml());
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut output, conformance.rel_ns());
    if value.fonts.is_empty() {
        output.extend_from_slice(b"\"/>");
        return Ok(output);
    }
    output.extend_from_slice(b"\">");
    for font in &value.fonts {
        output.extend_from_slice(b"<p:embeddedFont><p:font typeface=\"");
        escape(&mut output, &font.typeface);
        output.push(b'\"');
        if let Some(panose) = font.panose {
            attribute(&mut output, "panose", &hex_panose(panose)?);
        }
        if let Some(value) = font.pitch_family {
            attribute(&mut output, "pitchFamily", &value.wire().to_string());
        }
        if let Some(value) = font.charset {
            attribute(&mut output, "charset", &value.wire().to_string());
        }
        output.extend_from_slice(b"/>");
        for face in &font.faces {
            output.extend_from_slice(b"<p:");
            output.extend_from_slice(face.style.element().as_bytes());
            output.extend_from_slice(b" r:id=\"");
            escape(&mut output, &face.relationship_id);
            output.extend_from_slice(b"\"/>");
        }
        output.extend_from_slice(b"</p:embeddedFont>");
    }
    output.extend_from_slice(b"</p:embeddedFontLst>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized embedded-font XML bytes"));
    }
    Ok(output)
}

/// Loads embedded-font metadata and validates every referenced inert font part.
fn load_raw(package: &OpcPackage) -> Result<Option<RawFonts>> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation)?;
    let presentation_name = presentation.partname().to_string();
    let parsed = parse_presentation(presentation.blob())?;
    let conformance = parsed.conformance;
    validate_font_relationship_sources(package, &presentation_name)?;
    let Some(mut value) = parsed.value else {
        reject_orphan_font_parts(package, &HashSet::new())?;
        return Ok(None);
    };
    let mut targets = HashSet::new();
    let mut references = HashSet::new();
    let mut resources = HashMap::<String, RawResource>::new();
    let mut total_bytes = 0usize;
    for font in &mut value.fonts {
        for face in &mut font.faces {
            references.insert(face.relationship_id.clone());
            let relationship = presentation
                .rels()
                .get(&face.relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "missing embedded-font relationship '{}'",
                        face.relationship_id
                    ))
                })?;
            if relationship.reltype() != conformance.font_rel() {
                return Err(invalid(format!(
                    "relationship '{}' does not match the presentation conformance",
                    face.relationship_id,
                )));
            }
            if relationship.is_external() {
                return Err(invalid("embedded-font relationship must be internal"));
            }
            let target = relationship.target_partname()?;
            let target_name = target.to_string();
            targets.insert(target_name.clone());
            if let Some(resource) = resources.get(&target_name) {
                face.resource = Some(resource.clone());
                continue;
            }
            let part = package.get_part(&target)?;
            if !is_font_content_type(part.content_type()) {
                return Err(invalid(format!(
                    "font part '{target}' has invalid content type '{}'",
                    part.content_type()
                )));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{target}' has outbound relationships"
                )));
            }
            if part.blob().len() > MAX_FONT_BYTES {
                return Err(limit("individual font bytes"));
            }
            total_bytes = total_bytes
                .checked_add(part.blob().len())
                .ok_or_else(|| limit("total font bytes"))?;
            if total_bytes > MAX_TOTAL_FONT_BYTES {
                return Err(limit("total font bytes"));
            }
            let resource = RawResource {
                part_name: target_name.clone(),
                content_type: part.content_type().to_owned(),
                data: part.blob_arc(),
            };
            resources.insert(target_name, resource.clone());
            face.resource = Some(resource);
        }
    }
    validate_inbound_font_graph(
        package,
        &presentation_name,
        presentation,
        &references,
        &targets,
    )?;
    reject_orphan_font_parts(package, &targets)?;
    Ok(Some(value))
}

/// Atomically stores the complete embedded-font graph.
///
/// Existing font relationships are replaced. RawFont parts still referenced by
/// another relationship are retained, and unrelated presentation XML is copied
/// byte-for-byte.
fn put_raw(package: &mut OpcPackage, value: &RawFonts, conformance: Conformance) -> Result<bool> {
    validate_value(value, true)?;
    let old = load_raw(package)?;
    let presentation = package.main_document_part()?;
    let presentation_name = presentation.partname().clone();
    let parsed = parse_presentation(presentation.blob())?;
    if parsed.conformance != conformance {
        return Err(invalid(
            "requested conformance does not match the presentation namespace",
        ));
    }
    let enabled = !value.fonts.is_empty();
    if enabled && old.as_ref() == Some(value) {
        return Ok(false);
    }
    if !enabled && old.is_none() && !embedding_enabled(presentation.blob())? {
        return Ok(false);
    }
    let fragment = if value.fonts.is_empty() {
        Vec::new()
    } else {
        write_raw(value, conformance)?
    };
    let updated_xml = patch_embedding_flag(
        &patch_font_list(presentation.blob(), &fragment, conformance)?,
        enabled,
    )?;
    let staged = parse_presentation(&updated_xml)?;
    let expected = enabled.then(|| metadata_only(value));
    if staged.conformance != conformance || staged.value != expected {
        return Err(invalid("staged embedded-font XML did not round-trip"));
    }
    let old_relationship_ids = old
        .iter()
        .flat_map(|value| &value.fonts)
        .flat_map(|font| font.faces.iter().map(|face| face.relationship_id.clone()))
        .collect::<HashSet<_>>();
    let old_part_names = old
        .iter()
        .flat_map(|value| &value.fonts)
        .flat_map(|font| font.faces.iter())
        .filter_map(|face| {
            face.resource
                .as_ref()
                .map(|resource| resource.part_name.clone())
        })
        .collect::<HashSet<_>>();
    let mut relationship_ids = HashMap::<String, PackURI>::new();
    let mut resources = HashMap::<String, (String, Arc<Vec<u8>>)>::new();
    let mut relationships = Vec::new();
    for font in &value.fonts {
        for face in &font.faces {
            let resource = face
                .resource
                .as_ref()
                .ok_or_else(|| invalid("embedded-font resource is required for package storage"))?;
            let uri = PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
            if let Some(existing) = relationship_ids.get(&face.relationship_id) {
                if existing != &uri {
                    return Err(invalid(format!(
                        "relationship ID '{}' resolves to conflicting font parts",
                        face.relationship_id
                    )));
                }
            } else {
                if presentation.rels().get(&face.relationship_id).is_some()
                    && !old_relationship_ids.contains(&face.relationship_id)
                {
                    return Err(invalid(format!(
                        "relationship ID '{}' already exists",
                        face.relationship_id
                    )));
                }
                relationship_ids.insert(face.relationship_id.clone(), uri.clone());
                relationships.push((uri.clone(), face.relationship_id.clone()));
            }
            if let Some((content_type, data)) = resources.get(uri.as_str()) {
                if content_type != &resource.content_type
                    || (!Arc::ptr_eq(data, &resource.data)
                        && data.as_slice() != resource.data.as_slice())
                {
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
        let uri = PackURI::new(part_name).map_err(Error::Invalid)?;
        if let Ok(part) = package.get_part(&uri) {
            if part.content_type() != content_type {
                return Err(invalid(format!("font part '{uri}' content type collision")));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{uri}' has outbound relationships"
                )));
            }
            let stored = part.blob_arc();
            let same_data = Arc::ptr_eq(&stored, data) || stored.as_slice() == data.as_slice();
            if !same_data && !old_part_names.contains(part_name) {
                return Err(invalid(format!("font part '{uri}' data collision")));
            }
            if !same_data
                && has_inbound_outside_relationships(
                    package,
                    &uri,
                    &presentation_name,
                    &old_relationship_ids,
                )?
            {
                return Err(invalid(format!(
                    "shared font part '{uri}' cannot be overwritten"
                )));
            }
        }
    }

    let mut candidate = package.clone();
    candidate.unsign();
    let existing_font_relationships = candidate
        .get_part(&presentation_name)?
        .rels()
        .iter()
        .filter(|relationship| is_font_relationship(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<Vec<_>>();
    for relationship_id in existing_font_relationships {
        candidate
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .remove(&relationship_id);
    }
    for (uri, relationship_id) in &relationships {
        candidate
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .add_relationship(
                conformance.font_rel().into(),
                uri.relative_ref(presentation_name.base_uri()),
                relationship_id.clone(),
                false,
            );
    }
    for (part_name, (content_type, data)) in resources {
        let uri = PackURI::new(&part_name).map_err(Error::Invalid)?;
        if let Ok(part) = candidate.get_part_mut(&uri) {
            part.set_blob_shared(data);
        } else {
            candidate.add_part(Box::new(BlobPart::new_shared(uri, content_type, data)));
        }
    }
    candidate
        .get_part_mut(&presentation_name)?
        .set_blob(updated_xml);
    let retained = relationships
        .iter()
        .map(|(uri, _)| uri.to_string())
        .collect::<HashSet<_>>();
    for old_part in old_part_names {
        if !retained.contains(&old_part) {
            let uri = PackURI::new(&old_part).map_err(Error::Invalid)?;
            if !part_is_referenced(&candidate, &uri)? {
                candidate.remove_part(&uri);
            }
        }
    }
    *package = candidate;
    Ok(true)
}

/// Load the complete semantic embedded-font collection.
///
/// Relationship IDs and part names remain private provenance. Font programs
/// that share one physical part share the same allocation in memory.
pub fn load(package: &OpcPackage) -> Result<Option<Fonts>> {
    load_raw(package)?.map(fonts_from_raw).transpose()
}

/// Atomically publish a complete collection, consuming its owned values.
///
/// Returns `false` for an exact semantic and physical no-op, preserving any
/// valid package signatures. A real mutation invalidates signatures only after
/// every bounded validation and staging operation succeeds.
pub fn put(package: &mut OpcPackage, fonts: Fonts) -> Result<bool> {
    validate_fonts(&fonts, true)?;
    let current = load_raw(package)?;
    let presentation = package.main_document_part()?;
    let conformance = parse_presentation(presentation.blob())?.conformance;
    let empty = RawFonts::default();
    let raw = fonts_into_raw(package, fonts, current.as_ref().unwrap_or(&empty))?;
    if current.as_ref() == Some(&raw) {
        return Ok(false);
    }
    put_raw(package, &raw, conformance)
}

/// Remove the complete embedded-font graph and return its previous semantic
/// value. Absence is an exact no-op.
pub fn remove(package: &mut OpcPackage) -> Result<Option<Fonts>> {
    let Some(current) = load_raw(package)? else {
        return Ok(None);
    };
    let value = fonts_from_raw(current.clone())?;
    let presentation = package.main_document_part()?;
    let conformance = parse_presentation(presentation.blob())?.conformance;
    let changed = put_raw(package, &RawFonts::default(), conformance)?;
    if changed { Ok(Some(value)) } else { Ok(None) }
}

/// Detect the presentation namespace profile without exposing XML details.
pub fn conformance(package: &OpcPackage) -> Result<Conformance> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation)?;
    Ok(parse_presentation(presentation.blob())?.conformance)
}

fn fonts_from_raw(raw: RawFonts) -> Result<Fonts> {
    let mut fonts = Fonts::new();
    fonts
        .fonts
        .try_reserve(raw.fonts.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-font collection",
            source,
        })?;
    for raw_font in raw.fonts {
        let mut faces = Vec::new();
        faces
            .try_reserve(raw_font.faces.len())
            .map_err(|source| Error::Allocation {
                resource: "embedded-font faces",
                source,
            })?;
        for raw_face in raw_font.faces {
            let resource = raw_face
                .resource
                .ok_or_else(|| invalid("loaded embedded-font face has no resource"))?;
            let format = Format::parse(&resource.content_type)?;
            let data = Data::preserve(resource.data, format)?;
            faces.push(Face {
                style: raw_face.style,
                data,
                source: Some(Source {
                    relationship_id: raw_face.relationship_id,
                    part_name: resource.part_name,
                }),
            });
        }
        validate_typeface(&raw_font.typeface)?;
        fonts.fonts.push(Font {
            key: name_key(&raw_font.typeface),
            typeface: raw_font.typeface,
            panose: raw_font.panose,
            pitch_family: raw_font.pitch_family,
            charset: raw_font.charset,
            faces,
        });
    }
    fonts.reindex()?;
    validate_fonts(&fonts, false)?;
    Ok(fonts)
}

fn fonts_into_raw(package: &OpcPackage, fonts: Fonts, current: &RawFonts) -> Result<RawFonts> {
    let presentation = package.main_document_part()?;
    let mut relationship_ids = presentation
        .rels()
        .iter()
        .filter(|relationship| !is_font_relationship(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    let current_sources = current
        .fonts
        .iter()
        .flat_map(|font| &font.faces)
        .filter_map(|face| {
            face.resource.as_ref().map(|resource| {
                (
                    (face.relationship_id.clone(), resource.part_name.clone()),
                    (resource.content_type.clone(), Arc::clone(&resource.data)),
                )
            })
        })
        .collect::<HashMap<_, _>>();
    let mut part_names = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    let mut shared_parts = HashMap::<(usize, Format), String>::new();
    let mut claimed_parts = HashMap::<String, (usize, Format)>::new();
    let mut claimed_relationships = HashMap::<String, String>::new();
    let mut raw_fonts = Vec::new();
    raw_fonts
        .try_reserve(fonts.fonts.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-font staging",
            source,
        })?;
    for font in fonts.fonts {
        let mut raw_faces = Vec::new();
        raw_faces
            .try_reserve(font.faces.len())
            .map_err(|source| Error::Allocation {
                resource: "embedded-font face staging",
                source,
            })?;
        for face in font.faces {
            let data_id = (Arc::as_ptr(&face.data.bytes) as usize, face.data.format);
            let valid_source = face.source.as_ref().filter(|source| {
                current_sources
                    .get(&(source.relationship_id.clone(), source.part_name.clone()))
                    .is_some_and(|(content_type, data)| {
                        content_type == face.data.format.content_type()
                            && Arc::ptr_eq(data, &face.data.bytes)
                    })
            });
            let part_name = if let Some(part_name) = shared_parts.get(&data_id) {
                part_name.clone()
            } else {
                let reusable =
                    valid_source
                        .map(|source| source.part_name.clone())
                        .filter(|part_name| {
                            claimed_parts
                                .get(part_name)
                                .is_none_or(|claimed| *claimed == data_id)
                        });
                let part_name = match reusable {
                    Some(part_name) => part_name,
                    None => next_font_part_name(&part_names, face.data.format)?,
                };
                part_names.insert(part_name.clone());
                claimed_parts.insert(part_name.clone(), data_id);
                shared_parts.insert(data_id, part_name.clone());
                part_name
            };
            let relationship_id = if let Some(source) = valid_source {
                match claimed_relationships.get(&source.relationship_id) {
                    Some(existing) if existing == &part_name => source.relationship_id.clone(),
                    None if !relationship_ids.contains(&source.relationship_id) => {
                        source.relationship_id.clone()
                    },
                    _ => next_font_relationship_id(&relationship_ids)?,
                }
            } else {
                next_font_relationship_id(&relationship_ids)?
            };
            relationship_ids.insert(relationship_id.clone());
            claimed_relationships
                .entry(relationship_id.clone())
                .or_insert_with(|| part_name.clone());
            raw_faces.push(RawFace {
                style: face.style,
                relationship_id,
                resource: Some(RawResource {
                    part_name,
                    content_type: face.data.format.content_type().into(),
                    data: face.data.bytes,
                }),
            });
        }
        raw_fonts.push(RawFont {
            has_descriptor: true,
            typeface: font.typeface,
            panose: font.panose,
            pitch_family: font.pitch_family,
            charset: font.charset,
            faces: raw_faces,
        });
    }
    let raw = RawFonts { fonts: raw_fonts };
    validate_value(&raw, true)?;
    Ok(raw)
}

fn next_font_relationship_id(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("rIdFont{index}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(limit("relationship IDs"))
}

fn next_font_part_name(used: &HashSet<String>, format: Format) -> Result<String> {
    let extension = match format {
        Format::PowerPoint => "fntdata",
        Format::Standard => "ttf",
    };
    for index in 1..=u32::MAX {
        let candidate = format!("/ppt/fonts/font{index}.{extension}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(limit("part names"))
}

fn embedding_enabled(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"presentation"
                    || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PML.as_bytes() || value == STRICT_PML.as_bytes())
                {
                    return Err(invalid("expected a PresentationML presentation root"));
                }
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    if attribute.key.as_ref() == b"embedTrueTypeFonts" {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(xml_error)?;
                        return match value.as_ref() {
                            "1" | "true" => Ok(true),
                            "0" | "false" => Ok(false),
                            _ => Err(invalid(format!(
                                "invalid embedTrueTypeFonts boolean '{value}'"
                            ))),
                        };
                    }
                }
                return Ok(false);
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => return Err(invalid("missing presentation root")),
            _ => return Err(invalid("unexpected content before presentation root")),
        }
    }
}

fn patch_embedding_flag(xml: &[u8], enabled: bool) -> Result<Vec<u8>> {
    let (start, end) = presentation_start_tag(xml)?;
    let tag = xml
        .get(start..end)
        .ok_or_else(|| invalid("presentation start-tag range is invalid"))?;
    let value_range = find_unqualified_attribute_value(tag, b"embedTrueTypeFonts")?;
    let replacement = if enabled {
        b"1".as_slice()
    } else {
        b"0".as_slice()
    };
    if let Some((value_start, value_end)) = value_range {
        if tag.get(value_start..value_end) == Some(replacement) {
            return Ok(xml.to_vec());
        }
        let absolute_start = start
            .checked_add(value_start)
            .ok_or_else(|| limit("updated presentation XML bytes"))?;
        let absolute_end = start
            .checked_add(value_end)
            .ok_or_else(|| limit("updated presentation XML bytes"))?;
        return replace_bytes(xml, absolute_start, absolute_end, replacement);
    }
    if !enabled {
        return Ok(xml.to_vec());
    }
    let mut insertion = end
        .checked_sub(1)
        .ok_or_else(|| invalid("presentation start tag is empty"))?;
    if xml.get(insertion.wrapping_sub(1)) == Some(&b'/') {
        insertion = insertion
            .checked_sub(1)
            .ok_or_else(|| invalid("presentation start tag is empty"))?;
    }
    replace_bytes(xml, insertion, insertion, b" embedTrueTypeFonts=\"1\"")
}

fn presentation_start_tag(xml: &[u8]) -> Result<(usize, usize)> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("presentation XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"presentation"
                    || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PML.as_bytes() || value == STRICT_PML.as_bytes())
                {
                    return Err(invalid("expected a PresentationML presentation root"));
                }
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_| invalid("presentation XML offset overflow"))?;
                return Ok((start, end));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => return Err(invalid("missing presentation root")),
            _ => return Err(invalid("unexpected content before presentation root")),
        }
    }
}

fn find_unqualified_attribute_value(tag: &[u8], wanted: &[u8]) -> Result<Option<(usize, usize)>> {
    let mut offset = tag
        .iter()
        .position(|byte| *byte == b'<')
        .ok_or_else(|| invalid("presentation start tag has no opening delimiter"))?
        + 1;
    while tag
        .get(offset)
        .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'>' | b'/'))
    {
        offset += 1;
    }
    loop {
        while tag.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        if tag
            .get(offset)
            .is_none_or(|byte| matches!(byte, b'>' | b'/'))
        {
            return Ok(None);
        }
        let name_start = offset;
        while tag
            .get(offset)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'=' | b'>' | b'/'))
        {
            offset += 1;
        }
        let name_end = offset;
        while tag.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        if tag.get(offset) != Some(&b'=') {
            return Err(invalid("presentation root contains a malformed attribute"));
        }
        offset += 1;
        while tag.get(offset).is_some_and(u8::is_ascii_whitespace) {
            offset += 1;
        }
        let quote = *tag
            .get(offset)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| invalid("presentation root attribute is not quoted"))?;
        offset += 1;
        let value_start = offset;
        while tag.get(offset).is_some_and(|byte| *byte != quote) {
            offset += 1;
        }
        let value_end = offset;
        if tag.get(offset) != Some(&quote) {
            return Err(invalid("presentation root attribute is unterminated"));
        }
        offset += 1;
        if tag.get(name_start..name_end) == Some(wanted) {
            return Ok(Some((value_start, value_end)));
        }
    }
}

fn replace_bytes(xml: &[u8], start: usize, end: usize, replacement: &[u8]) -> Result<Vec<u8>> {
    if start > end || end > xml.len() {
        return Err(invalid("presentation XML replacement range is invalid"));
    }
    let length = xml
        .len()
        .checked_sub(end - start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| limit("updated presentation XML bytes"))?;
    if length > MAX_XML_BYTES {
        return Err(limit("updated presentation XML bytes"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "embedded-font presentation patch",
            source,
        })?;
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[end..]);
    Ok(output)
}

fn patch_font_list(xml: &[u8], fragment: &[u8], conformance: Conformance) -> Result<Vec<u8>> {
    let scan = active_direct_elements(xml, conformance)?;
    let mut lists = scan
        .elements
        .iter()
        .filter(|element| element.rank == presentation_child_rank("embeddedFontLst"));
    let list = lists.next();
    if lists.next().is_some() {
        return Err(invalid(
            "presentation has multiple active direct embeddedFontLst elements",
        ));
    }
    if let Some(list) = list {
        return replace_bytes(xml, list.start, list.end, fragment);
    }
    if fragment.is_empty() {
        Ok(xml.to_vec())
    } else {
        insert_font_list(xml, fragment, &scan)
    }
}

#[derive(Clone, Copy)]
struct DirectElement {
    start: usize,
    end: usize,
    rank: Option<usize>,
}

struct DirectScan {
    elements: Vec<DirectElement>,
    root_close: usize,
}

#[derive(Clone, Copy)]
struct DirectFrame {
    mce_wrapper: bool,
    element: Option<usize>,
}

fn active_direct_elements(xml: &[u8], conformance: Conformance) -> Result<DirectScan> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("presentation XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut frames = Vec::<DirectFrame>::new();
    let mut elements = Vec::<DirectElement>::new();
    let mut offsets = Vec::<u32>::new();
    let mut root_seen = false;
    let mut root_close = None;
    let mut nodes = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("presentation XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                if frames.len() >= MAX_DEPTH {
                    return Err(limit("presentation XML depth"));
                }
                let is_pml = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value))
                        if value == conformance.pml().as_bytes()
                );
                let is_mce = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value)) if value == MCE_NS.as_bytes()
                );
                let local = element.local_name();
                if frames.is_empty() {
                    if root_seen || !is_pml || local.as_ref() != b"presentation" {
                        return Err(invalid(
                            "presentation root does not match requested conformance",
                        ));
                    }
                    root_seen = true;
                }
                let effective_direct =
                    !frames.is_empty() && frames.iter().skip(1).all(|frame| frame.mce_wrapper);
                let rank = effective_direct
                    .then(|| std::str::from_utf8(local.as_ref()).ok())
                    .flatten()
                    .and_then(presentation_child_rank);
                let direct = if is_pml && rank.is_some() {
                    let index = elements.len();
                    elements.push(DirectElement {
                        start,
                        end: 0,
                        rank,
                    });
                    offsets.push(
                        u32::try_from(start)
                            .map_err(|_| invalid("presentation XML offset overflow"))?,
                    );
                    Some(index)
                } else {
                    None
                };
                let mce_wrapper = is_mce
                    && matches!(
                        local.as_ref(),
                        b"AlternateContent" | b"Choice" | b"Fallback"
                    );
                frames.push(DirectFrame {
                    mce_wrapper,
                    element: direct,
                });
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                if frames.is_empty() {
                    return Err(invalid("presentation root cannot be empty"));
                }
                let is_pml = matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value))
                        if value == conformance.pml().as_bytes()
                );
                let effective_direct = frames.iter().skip(1).all(|frame| frame.mce_wrapper);
                let local = element.local_name();
                let rank = effective_direct
                    .then(|| std::str::from_utf8(local.as_ref()).ok())
                    .flatten()
                    .and_then(presentation_child_rank);
                if is_pml && rank.is_some() {
                    elements.push(DirectElement {
                        start,
                        end: usize::try_from(reader.buffer_position())
                            .map_err(|_| invalid("presentation XML offset overflow"))?,
                        rank,
                    });
                    offsets.push(
                        u32::try_from(start)
                            .map_err(|_| invalid("presentation XML offset overflow"))?,
                    );
                }
            },
            Event::End(_) => {
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("unexpected presentation closing element"))?;
                if let Some(index) = frame.element {
                    let end = usize::try_from(reader.buffer_position())
                        .map_err(|_| invalid("presentation XML offset overflow"))?;
                    let len = elements.len();
                    elements
                        .get_mut(index)
                        .ok_or(Error::FontIndexOutOfBounds { index, len })?
                        .end = end;
                }
                if frames.is_empty() {
                    root_close = Some(start);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !frames.is_empty() {
        return Err(invalid("invalid presentation XML"));
    }
    let defaults = ActiveOffsetLimits::default();
    let limits = ActiveOffsetLimits {
        max_source_bytes: MAX_XML_BYTES,
        max_offsets: MAX_NODES,
        max_marked_bytes: MAX_MCE_MARKED_BYTES,
        mce: MceLimits {
            max_input_bytes: MAX_MCE_MARKED_BYTES,
            max_output_bytes: MAX_MCE_MARKED_BYTES,
            max_depth: MAX_DEPTH,
            ..defaults.mce
        },
    };
    let active = active_offsets(xml, &offsets, &MceCapabilities::default(), &limits)?;
    let mut active = active.into_iter().peekable();
    elements.retain(|element| {
        let Ok(start) = u32::try_from(element.start) else {
            return false;
        };
        while active.peek().is_some_and(|offset| *offset < start) {
            active.next();
        }
        if active.peek() == Some(&start) {
            active.next();
            true
        } else {
            false
        }
    });
    Ok(DirectScan {
        elements,
        root_close: root_close.ok_or_else(|| invalid("missing presentation closing element"))?,
    })
}

fn insert_font_list(xml: &[u8], fragment: &[u8], scan: &DirectScan) -> Result<Vec<u8>> {
    let font_rank = presentation_child_rank("embeddedFontLst")
        .ok_or_else(|| invalid("missing embeddedFontLst schema rank"))?;
    let position = scan
        .elements
        .iter()
        .filter(|element| element.rank.is_some_and(|rank| rank > font_rank))
        .map(|element| element.start)
        .min()
        .unwrap_or(scan.root_close);
    let length = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated presentation XML bytes"))?;
    if length > MAX_XML_BYTES {
        return Err(limit("updated presentation XML bytes"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "embedded-font presentation patch",
            source,
        })?;
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn presentation_child_rank(local: &str) -> Option<usize> {
    [
        "sldMasterIdLst",
        "notesMasterIdLst",
        "handoutMasterIdLst",
        "sldIdLst",
        "sldSz",
        "notesSz",
        "smartTags",
        "embeddedFontLst",
        "custShowLst",
        "photoAlbum",
        "custDataLst",
        "kinsoku",
        "defaultTextStyle",
        "modifyVerifier",
        "extLst",
    ]
    .iter()
    .position(|name| *name == local)
}

fn validate_fonts(value: &Fonts, require_unique_names: bool) -> Result<()> {
    if value.fonts.len() > MAX_FONTS {
        return Err(limit("embedded fonts"));
    }
    let mut names = HashSet::new();
    let mut resources = HashSet::new();
    let mut total = 0usize;
    for font in &value.fonts {
        validate_typeface(&font.typeface)?;
        if require_unique_names && !names.insert(font.key.as_str()) {
            return Err(Error::DuplicateFontName {
                name: font.typeface.clone(),
                matches: 2,
            });
        }
        if font.faces.len() > 4 {
            return Err(invalid("embeddedFont has more than four styles"));
        }
        for pair in font.faces.windows(2) {
            if pair[0].style.rank() >= pair[1].style.rank() {
                return Err(invalid(
                    "embedded-font styles are duplicated or out of schema order",
                ));
            }
        }
        for face in &font.faces {
            validate_font_bytes(&face.data.bytes)?;
            let identity = (Arc::as_ptr(&face.data.bytes) as usize, face.data.format);
            if resources.insert(identity) {
                total = total
                    .checked_add(face.data.bytes.len())
                    .ok_or_else(|| limit("total font bytes"))?;
                if total > MAX_TOTAL_FONT_BYTES {
                    return Err(limit("total font bytes"));
                }
            }
        }
    }
    Ok(())
}

fn validate_typeface(value: &str) -> Result<()> {
    bounded_string(value)?;
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid(
            "embedded-font typeface contains an XML 1.0-forbidden character",
        ))
    }
}

fn validate_font_bytes(value: &[u8]) -> Result<()> {
    if value.len() > MAX_FONT_BYTES {
        Err(limit("individual font bytes"))
    } else {
        Ok(())
    }
}

fn validate_eot(value: &[u8]) -> Result<()> {
    const VERSION_1: u32 = 0x0001_0000;
    const VERSION_2_1: u32 = 0x0002_0001;
    const VERSION_2_2: u32 = 0x0002_0002;
    const FLAGS: u32 = 0x1000_00F5;
    const EUDC: u32 = 0x0000_0020;

    let eot_size = usize::try_from(le_u32(value, 0)?)
        .map_err(|_| invalid("EOT size does not fit this platform"))?;
    if eot_size != value.len() {
        return Err(invalid("EOT size does not match the container length"));
    }
    let font_size = usize::try_from(le_u32(value, 4)?)
        .map_err(|_| invalid("EOT font-data size does not fit this platform"))?;
    if font_size == 0 {
        return Err(invalid("EOT font-data payload is empty"));
    }
    let font_start = value
        .len()
        .checked_sub(font_size)
        .ok_or_else(|| invalid("EOT font-data size exceeds the container"))?;
    let version = le_u32(value, 8)?;
    if !matches!(version, VERSION_1 | VERSION_2_1 | VERSION_2_2) {
        return Err(invalid(format!("unsupported EOT version 0x{version:08X}")));
    }
    let flags = le_u32(value, 12)?;
    if flags & !FLAGS != 0 {
        return Err(invalid(format!(
            "EOT processing flags contain reserved bits 0x{:08X}",
            flags & !FLAGS
        )));
    }
    if version == VERSION_1 && flags & EUDC != 0 {
        return Err(invalid("EOT version 1 cannot contain an EUDC font"));
    }
    if value.get(27).copied().is_none_or(|italic| italic > 1) {
        return Err(invalid("EOT italic byte must be zero or one"));
    }
    License::from_fs_type(le_u16(value, 32)?)?;
    if le_u16(value, 34)? != 0x504C {
        return Err(invalid("EOT magic number is not 0x504C"));
    }
    if value
        .get(64..80)
        .is_none_or(|reserved| reserved.iter().any(|byte| *byte != 0))
    {
        return Err(invalid("EOT reserved header words must be zero"));
    }
    if le_u16(value, 80)? != 0 {
        return Err(invalid("EOT header padding must be zero"));
    }

    let mut cursor = 82usize;
    for name in ["family", "style", "version", "full"] {
        eot_utf16(value, &mut cursor, font_start, name)?;
        if name != "full" && eot_u16(value, &mut cursor, font_start, "name padding")? != 0 {
            return Err(invalid("EOT name padding must be zero"));
        }
    }

    if version != VERSION_1 {
        if eot_u16(value, &mut cursor, font_start, "root padding")? != 0 {
            return Err(invalid("EOT root-string padding must be zero"));
        }
        let root = eot_sized(value, &mut cursor, font_start, "root string")?;
        if root.len() % 2 != 0 {
            return Err(invalid("EOT root string is not UTF-16 byte-aligned"));
        }
        if version == VERSION_2_2 {
            let checksum = eot_u32(value, &mut cursor, font_start, "root checksum")?;
            let expected = root
                .iter()
                .fold(0u32, |sum, byte| sum.wrapping_add(u32::from(*byte)))
                ^ 0x5047_5342;
            if checksum != expected {
                return Err(invalid("EOT root-string checksum is invalid"));
            }
            let _code_page = eot_u32(value, &mut cursor, font_start, "EUDC code page")?;
            if eot_u16(value, &mut cursor, font_start, "signature padding")? != 0 {
                return Err(invalid("EOT signature padding must be zero"));
            }
            let signature = eot_sized(value, &mut cursor, font_start, "signature")?;
            if !signature.is_empty() {
                return Err(invalid("EOT reserved signature must be empty"));
            }
            let eudc_flags = eot_u32(value, &mut cursor, font_start, "EUDC flags")?;
            if eudc_flags & !FLAGS != 0 {
                return Err(invalid("EOT EUDC flags contain reserved bits"));
            }
            let eudc_size = usize::try_from(eot_u32(
                value,
                &mut cursor,
                font_start,
                "EUDC font-data size",
            )?)
            .map_err(|_| invalid("EOT EUDC font-data size does not fit this platform"))?;
            eot_take(value, &mut cursor, eudc_size, font_start, "EUDC font data")?;
            if (flags & EUDC != 0) != (eudc_size != 0) {
                return Err(invalid("EOT EUDC flag and payload disagree"));
            }
        }
    }
    if cursor != font_start {
        return Err(invalid(
            "EOT variable header overlaps or precedes font data",
        ));
    }
    if flags & (0x0000_0004 | 0x1000_0000) == 0 {
        validate_sfnt(
            value
                .get(font_start..)
                .ok_or_else(|| invalid("EOT font-data range is invalid"))?,
        )?;
    }
    Ok(())
}

fn validate_sfnt(value: &[u8]) -> Result<()> {
    match value.get(..4) {
        Some(b"ttcf") => {
            let version = be_u32(value, 4)?;
            if !matches!(version, 0x0001_0000 | 0x0002_0000) {
                return Err(invalid(format!(
                    "unsupported TrueType Collection version 0x{version:08X}"
                )));
            }
            let fonts = usize::try_from(be_u32(value, 8)?)
                .map_err(|_| invalid("TrueType Collection font count does not fit"))?;
            if fonts == 0 {
                return Err(invalid("TrueType Collection contains no fonts"));
            }
            let offsets_end = 12usize
                .checked_add(
                    fonts
                        .checked_mul(4)
                        .ok_or_else(|| invalid("TrueType Collection offset table overflows"))?,
                )
                .ok_or_else(|| invalid("TrueType Collection offset table overflows"))?;
            if offsets_end > value.len() {
                return Err(invalid("TrueType Collection offset table is truncated"));
            }
            for index in 0..fonts {
                let field = 12usize
                    .checked_add(index * 4)
                    .ok_or_else(|| invalid("TrueType Collection font offset overflows"))?;
                let offset = usize::try_from(be_u32(value, field)?)
                    .map_err(|_| invalid("TrueType Collection font offset does not fit"))?;
                if offset % 4 != 0 {
                    return Err(invalid("TrueType Collection font offset is not aligned"));
                }
                validate_sfnt_at(value, offset)?;
            }
            Ok(())
        },
        Some(_) => validate_sfnt_at(value, 0),
        None => Err(invalid("sfnt container is missing its signature")),
    }
}

fn validate_sfnt_at(value: &[u8], offset: usize) -> Result<()> {
    let signature_end = offset
        .checked_add(4)
        .ok_or_else(|| invalid("sfnt offset overflows"))?;
    if !matches!(
        value.get(offset..signature_end),
        Some(b"\0\x01\0\0" | b"OTTO" | b"true" | b"typ1")
    ) {
        return Err(invalid("font data has an unsupported sfnt signature"));
    }
    let tables = usize::from(be_u16(value, offset + 4)?);
    let directory_end = offset
        .checked_add(12)
        .and_then(|base| {
            tables
                .checked_mul(16)
                .and_then(|size| base.checked_add(size))
        })
        .ok_or_else(|| invalid("sfnt table directory overflows"))?;
    if directory_end > value.len() {
        return Err(invalid("sfnt table directory is truncated"));
    }
    Ok(())
}

fn eot_utf16(value: &[u8], cursor: &mut usize, limit: usize, name: &str) -> Result<()> {
    let data = eot_sized(value, cursor, limit, name)?;
    if data.len() % 2 != 0 {
        return Err(invalid(format!(
            "EOT {name} name is not UTF-16 byte-aligned"
        )));
    }
    let words = data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]));
    if char::decode_utf16(words).any(|character| character.is_err()) {
        return Err(invalid(format!("EOT {name} name contains invalid UTF-16")));
    }
    Ok(())
}

fn eot_sized<'a>(
    value: &'a [u8],
    cursor: &mut usize,
    limit: usize,
    name: &str,
) -> Result<&'a [u8]> {
    let size = usize::from(eot_u16(value, cursor, limit, name)?);
    eot_take(value, cursor, size, limit, name)
}

fn eot_take<'a>(
    value: &'a [u8],
    cursor: &mut usize,
    size: usize,
    limit: usize,
    name: &str,
) -> Result<&'a [u8]> {
    let end = cursor
        .checked_add(size)
        .ok_or_else(|| invalid(format!("EOT {name} size overflows")))?;
    if end > limit {
        return Err(invalid(format!("EOT {name} extends into font data")));
    }
    let data = value
        .get(*cursor..end)
        .ok_or_else(|| invalid(format!("EOT {name} is truncated")))?;
    *cursor = end;
    Ok(data)
}

fn eot_u16(value: &[u8], cursor: &mut usize, limit: usize, name: &str) -> Result<u16> {
    let bytes = eot_take(value, cursor, 2, limit, name)?;
    let bytes = <[u8; 2]>::try_from(bytes).map_err(xml_error)?;
    Ok(u16::from_le_bytes(bytes))
}

fn eot_u32(value: &[u8], cursor: &mut usize, limit: usize, name: &str) -> Result<u32> {
    let bytes = eot_take(value, cursor, 4, limit, name)?;
    let bytes = <[u8; 4]>::try_from(bytes).map_err(xml_error)?;
    Ok(u32::from_le_bytes(bytes))
}

fn le_u16(value: &[u8], offset: usize) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u16::from_le_bytes(
        <[u8; 2]>::try_from(bytes).map_err(xml_error)?,
    ))
}

fn le_u32(value: &[u8], offset: usize) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u32::from_le_bytes(
        <[u8; 4]>::try_from(bytes).map_err(xml_error)?,
    ))
}

fn be_u16(value: &[u8], offset: usize) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u16::from_be_bytes(
        <[u8; 2]>::try_from(bytes).map_err(xml_error)?,
    ))
}

fn be_u32(value: &[u8], offset: usize) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u32::from_be_bytes(
        <[u8; 4]>::try_from(bytes).map_err(xml_error)?,
    ))
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&value)
        || ('\u{E000}'..='\u{FFFD}').contains(&value)
        || ('\u{10000}'..='\u{10FFFF}').contains(&value)
}

fn name_key(value: &str) -> String {
    value.chars().nfd().default_case_fold().nfd().collect()
}

fn validate_value(value: &RawFonts, require_resources: bool) -> Result<()> {
    if value.fonts.len() > MAX_FONTS {
        return Err(limit("embedded fonts"));
    }
    let mut resources = HashMap::<&str, (&str, &Arc<Vec<u8>>)>::new();
    let mut total = 0usize;
    for font in &value.fonts {
        bounded_string(&font.typeface)?;
        finish_font(font)?;
        for face in &font.faces {
            validate_relationship_id(&face.relationship_id)?;
            if require_resources && face.resource.is_none() {
                return Err(invalid(
                    "embedded-font resource is required for package storage",
                ));
            }
            if let Some(resource) = &face.resource {
                PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
                if !is_font_content_type(&resource.content_type) {
                    return Err(invalid(format!(
                        "invalid embedded-font content type '{}'",
                        resource.content_type
                    )));
                }
                validate_font_bytes(&resource.data)?;
                if let Some((content_type, data)) = resources.get(resource.part_name.as_str()) {
                    if *content_type != resource.content_type
                        || (!Arc::ptr_eq(data, &resource.data)
                            && data.as_slice() != resource.data.as_slice())
                    {
                        return Err(invalid(format!(
                            "shared font part '{}' has conflicting resources",
                            resource.part_name
                        )));
                    }
                } else {
                    resources.insert(
                        resource.part_name.as_str(),
                        (resource.content_type.as_str(), &resource.data),
                    );
                    total = total
                        .checked_add(resource.data.len())
                        .ok_or_else(|| limit("total font bytes"))?;
                    if total > MAX_TOTAL_FONT_BYTES {
                        return Err(limit("total font bytes"));
                    }
                }
            }
        }
    }
    Ok(())
}

fn finish_font(font: &RawFont) -> Result<()> {
    if !font.has_descriptor {
        return Err(invalid("embeddedFont is missing its font descriptor"));
    }
    if font.faces.len() > 4 {
        return Err(invalid("embeddedFont has more than four styles"));
    }
    for pair in font.faces.windows(2) {
        if pair[0].style.rank() >= pair[1].style.rank() {
            return Err(invalid(
                "embedded-font styles are duplicated or out of schema order",
            ));
        }
    }
    Ok(())
}

fn validate_font_relationship_sources(package: &OpcPackage, presentation: &str) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_font_relationship(relationship.reltype()))
    {
        return Err(invalid("package root cannot source font relationships"));
    }
    for part in package.iter_parts() {
        if part.partname().as_str() != presentation
            && part
                .rels()
                .iter()
                .any(|relationship| is_font_relationship(relationship.reltype()))
        {
            return Err(invalid(format!(
                "font relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_inbound_font_graph(
    package: &OpcPackage,
    presentation_name: &str,
    presentation: &dyn Part,
    references: &HashSet<String>,
    targets: &HashSet<String>,
) -> Result<()> {
    for relationship in presentation
        .rels()
        .iter()
        .filter(|relationship| is_font_relationship(relationship.reltype()))
    {
        if !references.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced font relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        let target = relationship.target_partname()?;
        if targets.contains(target.as_str()) {
            return Err(invalid(format!(
                "font part '{target}' has an invalid package-root relationship"
            )));
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = relationship.target_partname()?;
            if targets.contains(target.as_str())
                && (part.partname().as_str() != presentation_name
                    || !is_font_relationship(relationship.reltype())
                    || !references.contains(relationship.r_id()))
            {
                return Err(invalid(format!(
                    "font part '{target}' has an invalid inbound relationship"
                )));
            }
        }
    }
    Ok(())
}

fn reject_orphan_font_parts(package: &OpcPackage, targets: &HashSet<String>) -> Result<()> {
    for part in package.iter_parts() {
        if is_font_content_type(part.content_type())
            && !targets.contains(part.partname().as_str())
            && !part_is_referenced(package, part.partname())?
        {
            return Err(invalid(format!("orphan font part '{}'", part.partname())));
        }
    }
    Ok(())
}

fn require_presentation(part: &dyn Part) -> Result<()> {
    if matches!(
        part.content_type(),
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main part has unsupported presentation content type '{}'",
            part.content_type()
        )))
    }
}

fn metadata_only(value: &RawFonts) -> RawFonts {
    let mut value = value.clone();
    for font in &mut value.fonts {
        for face in &mut font.faces {
            face.resource = None;
        }
    }
    value
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
    presentation: &PackURI,
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
                && (part.partname() != presentation
                    || !replaced_relationships.contains(relationship.r_id()))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn parse_panose(value: &str) -> Result<Panose> {
    if value.len() != 20 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid("panose must contain exactly 20 hexadecimal digits"));
    }
    let mut output = [0u8; 10];
    for (index, byte) in output.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16).map_err(xml_error)?;
    }
    Ok(Panose::new(output))
}

fn hex_panose(value: Panose) -> Result<String> {
    let mut output = String::with_capacity(20);
    for byte in value.bytes() {
        use std::fmt::Write;
        write!(&mut output, "{byte:02X}").map_err(|_| Error::Write)?;
    }
    Ok(output)
}

fn is_font_relationship(value: &str) -> bool {
    matches!(value, FONT_REL | STRICT_FONT_REL)
}
fn is_font_content_type(value: &str) -> bool {
    matches!(value, FONT_DATA_CT | FONT_TTF_CT)
}
fn bounded_string(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("embedded-font string bytes"))
    }
}
fn add_string_bytes(total: &mut usize, count: usize) -> Result<()> {
    *total = total
        .checked_add(count)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn validate_relationship_id(value: &str) -> Result<()> {
    if !litchi_ooxml_common::xml::is_ncname(value) {
        return Err(invalid(format!("invalid relationship ID '{value}'")));
    }
    Ok(())
}
fn resolved_namespace(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'\"');
}
fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(name: &'static str) -> Error {
    Error::Limit {
        resource: name,
        limit: match name {
            "presentation XML bytes"
            | "MCE-processed presentation XML bytes"
            | "serialized embedded-font XML bytes"
            | "updated presentation XML bytes" => MAX_XML_BYTES,
            "XML nodes" => MAX_NODES,
            "XML depth" | "presentation XML depth" => MAX_DEPTH,
            "embedded fonts" => MAX_FONTS,
            "individual font bytes" => MAX_FONT_BYTES,
            "total font bytes" => MAX_TOTAL_FONT_BYTES,
            "embedded-font string bytes" | "XML string bytes" => MAX_STRING_BYTES,
            _ => usize::MAX,
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn root() -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
    }
    fn package(conformance: Conformance) -> OpcPackage {
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let xml = format!(
            "<p:presentation xmlns:p=\"{}\"><p:sldMasterIdLst/><p:defaultTextStyle/></p:presentation>",
            conformance.pml()
        );
        package.add_part(Box::new(BlobPart::new(
            uri,
            PRESENTATION_CT.into(),
            xml.into_bytes(),
        )));
        let office_rel = match conformance {
            Conformance::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            },
            Conformance::Strict => {
                "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
            },
        };
        package.rels_mut().add_relationship(
            office_rel.into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        package
    }
    fn eot(marker: u8) -> Vec<u8> {
        let mut value = vec![0; 96];
        value[0..4].copy_from_slice(&108u32.to_le_bytes());
        value[4..8].copy_from_slice(&12u32.to_le_bytes());
        value[8..12].copy_from_slice(&0x0001_0000u32.to_le_bytes());
        value[16] = marker;
        value[34..36].copy_from_slice(&0x504Cu16.to_le_bytes());
        value.extend_from_slice(&[0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        value
    }
    fn value() -> RawFonts {
        RawFonts {
            fonts: vec![RawFont {
                has_descriptor: true,
                typeface: "A&B".into(),
                panose: Some(Panose::new([2, 11, 6, 4, 2, 2, 2, 2, 2, 4])),
                pitch_family: Some(PitchFamily::new(Pitch::Variable, Family::Swiss)),
                charset: Some(Charset::ANSI),
                faces: vec![
                    RawFace {
                        style: Style::Regular,
                        relationship_id: "rIdFont1".into(),
                        resource: Some(RawResource {
                            part_name: "/ppt/fonts/font1.fntdata".into(),
                            content_type: FONT_DATA_CT.into(),
                            data: Arc::new(vec![0, 1, 2, 3]),
                        }),
                    },
                    RawFace {
                        style: Style::BoldItalic,
                        relationship_id: "rIdFont2".into(),
                        resource: Some(RawResource {
                            part_name: "/ppt/fonts/font2.fntdata".into(),
                            content_type: FONT_DATA_CT.into(),
                            data: Arc::new(vec![4, 5, 6]),
                        }),
                    },
                ],
            }],
        }
    }

    #[test]
    fn strict_xml_round_trip_and_mce_fallback() {
        let expected = value();
        let fragment = write_raw(&expected, Conformance::Strict).unwrap();
        let xml = [
            format!("<p:presentation xmlns:p=\"{STRICT_PML}\">").as_bytes(),
            fragment.as_slice(),
            b"</p:presentation>",
        ]
        .concat();
        let parsed = parse_raw(&xml).unwrap().unwrap();
        assert_eq!(parsed.fonts[0].typeface, "A&B");
        assert!(
            parsed.fonts[0]
                .faces
                .iter()
                .all(|face| face.resource.is_none())
        );
        let mce = format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback><p:embeddedFontLst><p:embeddedFont><p:font typeface="Fallback"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></mc:Fallback></mc:AlternateContent></p:presentation>"#
        );
        assert_eq!(
            parse_raw(mce.as_bytes()).unwrap().unwrap().fonts[0].typeface,
            "Fallback"
        );
    }

    #[test]
    fn loads_libreoffice_and_poi_reference_packages() {
        let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
            root().join("test-data/libreoffice-core/sd/qa/unit/data/BoldonseFontEmbedded.pptx"),
        )
        .unwrap();
        let mut libreoffice = package(Conformance::Transitional);
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        libreoffice
            .get_part_mut(&presentation_uri)
            .unwrap()
            .set_blob(physical.blob_for(&presentation_uri).unwrap());
        let font_uri = PackURI::new("/ppt/fonts/font1.fntdata").unwrap();
        let font_data = physical.blob_for(&font_uri).unwrap();
        assert!(Data::powerpoint(font_data.clone()).is_ok());
        libreoffice.add_part(Box::new(BlobPart::new(
            font_uri.clone(),
            FONT_DATA_CT.into(),
            font_data,
        )));
        libreoffice
            .get_part_mut(&presentation_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                FONT_REL.into(),
                "fonts/font1.fntdata".into(),
                "rId3".into(),
                false,
            );
        let fonts = load_raw(&libreoffice).unwrap().unwrap();
        assert_eq!(fonts.fonts[0].typeface, "Boldonse");
        assert_eq!(
            fonts.fonts[0].faces[0]
                .resource
                .as_ref()
                .unwrap()
                .data
                .len(),
            36_187
        );
        let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
            root().join("test-data/poi/test-data/slideshow/placeholder-layout-color.pptx"),
        )
        .unwrap();
        let mut poi = package(Conformance::Transitional);
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        poi.get_part_mut(&presentation_uri)
            .unwrap()
            .set_blob(physical.blob_for(&presentation_uri).unwrap());
        for (index, relationship_id) in
            (1..=6).zip(["rId4", "rId5", "rId6", "rId7", "rId8", "rId9"])
        {
            let uri = PackURI::new(format!("/ppt/fonts/font{index}.fntdata")).unwrap();
            let data = physical.blob_for(&uri).unwrap();
            poi.add_part(Box::new(BlobPart::new(uri, FONT_DATA_CT.into(), data)));
            poi.get_part_mut(&presentation_uri)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    FONT_REL.into(),
                    format!("fonts/font{index}.fntdata"),
                    relationship_id.into(),
                    false,
                );
        }
        let fonts = load_raw(&poi).unwrap().unwrap();
        assert_eq!(fonts.fonts.len(), 3);
        let roboto = fonts
            .fonts
            .iter()
            .find(|font| font.typeface == "Roboto")
            .unwrap();
        assert_eq!(roboto.faces.len(), 4);
        assert!(roboto.faces.iter().all(|face| {
            face.resource
                .as_ref()
                .is_some_and(|resource| !resource.data.is_empty())
        }));
    }

    #[test]
    fn package_writer_round_trips_strict_graph_and_schema_position() {
        let mut package = package(Conformance::Strict);
        let expected = value();
        put_raw(&mut package, &expected, Conformance::Strict).unwrap();
        assert_eq!(load_raw(&package).unwrap().unwrap(), expected);
        let xml = package.main_document_part().unwrap().blob();
        let list = memchr::memmem::find(xml, b"<p:embeddedFontLst").unwrap();
        let defaults = memchr::memmem::find(xml, b"<p:defaultTextStyle").unwrap();
        assert!(list < defaults);
    }

    #[test]
    fn rejects_malformed_xml_duplicates_and_caps() {
        for xml in [
            format!(r#"<p:presentation xmlns:p="{PML}"/>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" panose="12"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:bold xmlns:r="{REL_NS}" r:id="rId1"/><p:regular r:id="rId2"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<!DOCTYPE x><p:presentation xmlns:p="{PML}"/>"#),
        ].into_iter().skip(1) { assert!(parse_raw(xml.as_bytes()).is_err(), "{xml}"); }
        assert!(parse_raw(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let face = Face::new(Style::Regular, Data::powerpoint(eot(1)).unwrap());
        let font = Font::from_face("Duplicate", face).unwrap();
        let mut duplicate = Fonts::new();
        duplicate.add(font.clone()).unwrap();
        assert!(duplicate.add(font).is_err());
    }

    #[test]
    fn rejects_external_orphan_and_outbound_graphs() {
        let mut external = package(Conformance::Transitional);
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL_NS}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:regular r:id="rIdFont1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .set_blob(xml.into_bytes());
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                FONT_REL.into(),
                "https://invalid.example/font".into(),
                "rIdFont1".into(),
                true,
            );
        assert!(load_raw(&external).is_err());

        let mut orphan = package(Conformance::Transitional);
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/fonts/orphan.fntdata").unwrap(),
            FONT_DATA_CT.into(),
            vec![1],
        )));
        assert!(load_raw(&orphan).is_err());

        let mut outbound = package(Conformance::Transitional);
        put_raw(&mut outbound, &value(), Conformance::Transitional).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/fonts/font1.fntdata").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.bin".into(),
                "rId1".into(),
                false,
            );
        assert!(load_raw(&outbound).is_err());

        let mut root_owned = package(Conformance::Transitional);
        put_raw(&mut root_owned, &value(), Conformance::Transitional).unwrap();
        root_owned.rels_mut().add_relationship(
            "urn:not-a-font-owner".into(),
            "ppt/fonts/font1.fntdata".into(),
            "rIdOther".into(),
            false,
        );
        assert!(load_raw(&root_owned).is_err());
    }

    #[test]
    fn fs_type_is_compact_and_validated() {
        assert!(License::from_fs_type(0).unwrap().installable());
        assert!(
            License::from_fs_type(0x0008)
                .unwrap()
                .restrictions()
                .is_empty()
        );
        let editable = License::from_fs_type(0x0108).unwrap();
        assert_eq!(editable.permission(), Permission::Editable);
        assert!(
            editable
                .restrictions()
                .contains(Restrictions::NO_SUBSETTING)
        );
        assert!(!editable.installable());
        assert!(License::from_fs_type(0x0006).is_err());
        assert!(License::from_fs_type(0x8000).is_err());
    }

    #[test]
    fn typed_metadata_can_be_cleared_without_rebuilding_the_font() {
        let panose = Panose::new([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
        let pitch = PitchFamily::new(Pitch::Fixed, Family::Roman);
        let mut font = Font::new("Metadata")
            .unwrap()
            .with_panose(panose)
            .with_pitch_family(pitch)
            .with_charset(Charset::SHIFT_JIS);
        assert_eq!(font.set_panose(None), Some(panose));
        assert_eq!(font.set_pitch_family(None), Some(pitch));
        assert_eq!(font.set_charset(None), Some(Charset::SHIFT_JIS));
        assert_eq!(font.panose(), None);
        assert_eq!(font.pitch_family(), None);
        assert_eq!(font.charset(), None);
    }

    #[test]
    fn fresh_font_containers_are_structurally_checked() {
        assert!(Data::powerpoint(eot(3)).is_ok());
        let mut wrong_size = eot(3);
        wrong_size[0..4].copy_from_slice(&1u32.to_le_bytes());
        assert!(Data::powerpoint(wrong_size).is_err());
        let mut reserved = eot(3);
        reserved[64] = 1;
        assert!(Data::powerpoint(reserved).is_err());
        let sfnt = vec![0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        assert!(Data::standard(sfnt).is_ok());
        assert!(Data::standard(b"not-a-font".to_vec()).is_err());
    }

    #[test]
    fn present_empty_and_loaded_false_flag_round_trip_exactly() {
        let uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let mut empty = package(Conformance::Transitional);
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}"><p:sldMasterIdLst/><p:embeddedFontLst/><p:defaultTextStyle/></p:presentation>"#
        );
        empty.get_part_mut(&uri).unwrap().set_blob(xml.into_bytes());
        let loaded = load(&empty).unwrap().unwrap();
        assert!(loaded.is_empty());
        let before = empty.get_part(&uri).unwrap().blob().to_vec();
        empty.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(!put(&mut empty, loaded).unwrap());
        assert!(empty.is_signed());
        assert_eq!(empty.get_part(&uri).unwrap().blob(), before);
        assert!(remove(&mut empty).unwrap().is_some());
        assert!(load(&empty).unwrap().is_none());
        assert!(
            memchr::memmem::find(empty.get_part(&uri).unwrap().blob(), b"embeddedFontLst")
                .is_none()
        );

        let mut disabled = package(Conformance::Transitional);
        put_raw(&mut disabled, &value(), Conformance::Transitional).unwrap();
        let xml = patch_embedding_flag(disabled.get_part(&uri).unwrap().blob(), false).unwrap();
        disabled.get_part_mut(&uri).unwrap().set_blob(xml);
        let loaded = load(&disabled).unwrap().unwrap();
        let before = disabled.get_part(&uri).unwrap().blob().to_vec();
        disabled.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(!put(&mut disabled, loaded).unwrap());
        assert!(disabled.is_signed());
        assert_eq!(disabled.get_part(&uri).unwrap().blob(), before);
    }

    #[test]
    fn font_list_crud_edits_only_the_active_direct_mce_branch() {
        let mut package = package(Conformance::Transitional);
        let uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:mc="{MCE_NS}" xmlns:x="urn:future"><p:sldMasterIdLst/><mc:AlternateContent><mc:Choice Requires="x"><p:embeddedFontLst><p:embeddedFont><p:font typeface="Inactive"/></p:embeddedFont></p:embeddedFontLst></mc:Choice><mc:Fallback><p:defaultTextStyle/></mc:Fallback></mc:AlternateContent></p:presentation>"#
        );
        package
            .get_part_mut(&uri)
            .unwrap()
            .set_blob(xml.into_bytes());
        let mut fonts = Fonts::new();
        fonts
            .add(
                Font::from_face(
                    "Active",
                    Face::new(Style::Regular, Data::powerpoint(eot(9)).unwrap()),
                )
                .unwrap(),
            )
            .unwrap();
        assert!(put(&mut package, fonts).unwrap());
        let xml = package.get_part(&uri).unwrap().blob();
        assert!(memchr::memmem::find(xml, b"typeface=\"Inactive\"").is_some());
        assert!(memchr::memmem::find(xml, b"typeface=\"Active\"").is_some());
        let active = memchr::memmem::find(xml, b"typeface=\"Active\"").unwrap();
        let defaults = memchr::memmem::find(xml, b"<p:defaultTextStyle").unwrap();
        assert!(active < defaults);
        assert_eq!(
            load(&package)
                .unwrap()
                .unwrap()
                .get("Active")
                .unwrap()
                .name(),
            "Active"
        );
        assert!(remove(&mut package).unwrap().is_some());
        let xml = package.get_part(&uri).unwrap().blob();
        assert!(memchr::memmem::find(xml, b"typeface=\"Inactive\"").is_some());
        assert!(memchr::memmem::find(xml, b"typeface=\"Active\"").is_none());
        assert!(load(&package).unwrap().is_none());
    }

    #[test]
    fn descriptor_root_order_and_unicode_relationship_ids_are_checked() {
        let face_first = format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL_NS}"><p:embeddedFontLst><p:embeddedFont><p:regular r:id="rId1"/><p:font typeface="A"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        assert!(parse_raw(face_first.as_bytes()).is_err());
        let root_out_of_order = format!(
            r#"<p:presentation xmlns:p="{PML}"><p:defaultTextStyle/><p:embeddedFontLst/></p:presentation>"#
        );
        assert!(parse_raw(root_out_of_order.as_bytes()).is_err());
        let unicode = format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL_NS}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:regular r:id="字体"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        assert_eq!(
            parse_raw(unicode.as_bytes()).unwrap().unwrap().fonts[0].faces[0].relationship_id,
            "字体"
        );
    }

    #[test]
    fn rejected_collection_edits_leave_indexes_and_order_unchanged() {
        let mut fonts = Fonts::new();
        fonts.add(Font::new("First").unwrap()).unwrap();
        fonts.add(Font::new("Second").unwrap()).unwrap();
        let before = fonts.clone();
        assert!(fonts.reorder(&["First", "First"]).is_err());
        assert_eq!(fonts, before);
        assert!(
            fonts
                .replace("First", Font::new("SECOND").unwrap())
                .is_err()
        );
        assert_eq!(fonts, before);
        assert!(fonts.remove(9_usize).is_err());
        assert_eq!(fonts, before);
    }

    #[test]
    fn generated_crud_allocates_collisions_and_preserves_unknown_xml_atomically() {
        let mut package = package(Conformance::Transitional);
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let original = package.get_part(&presentation_uri).unwrap().blob();
        let marker = memchr::memmem::find(original, b"<p:defaultTextStyle").unwrap();
        let mut xml = original.to_vec();
        xml.splice(marker..marker, b"<!--font-preserve-->".iter().copied());
        package
            .get_part_mut(&presentation_uri)
            .unwrap()
            .set_blob(xml);

        let generated = Font::from_face(
            "Generated",
            Face::new(Style::Regular, Data::powerpoint(eot(7)).unwrap()),
        )
        .unwrap()
        .with_panose([2, 11, 6, 4, 2, 2, 2, 2, 2, 4])
        .with_pitch_family(PitchFamily::new(Pitch::Variable, Family::Swiss))
        .with_charset(Charset::ANSI);
        let mut fonts = Fonts::new();
        fonts.add(generated).unwrap();
        assert!(put(&mut package, fonts).unwrap());
        let loaded = load(&package).unwrap().unwrap();
        let found = loaded.get("generated").unwrap();
        assert_eq!(found.faces()[0].data().bytes(), eot(7));
        assert!(package.contains_part(&PackURI::new("/ppt/fonts/font1.fntdata").unwrap()));
        assert!(
            package
                .get_part(&presentation_uri)
                .unwrap()
                .blob()
                .windows(b"<!--font-preserve-->".len())
                .any(|window| window == b"<!--font-preserve-->")
        );
        assert!(embedding_enabled(package.get_part(&presentation_uri).unwrap().blob()).unwrap());

        let before = package.get_part(&presentation_uri).unwrap().blob().to_vec();
        let parts = package.part_count();
        let mut duplicate = loaded.clone();
        assert!(duplicate.add(found.clone()).is_err());
        assert_eq!(package.get_part(&presentation_uri).unwrap().blob(), before);
        assert_eq!(package.part_count(), parts);
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
        assert!(!put(&mut package, loaded.clone()).unwrap());
        assert!(package.is_signed());
        assert_eq!(package.get_part(&presentation_uri).unwrap().blob(), before);

        let mut changed = loaded;
        let mut replacement = changed
            .remove("Generated")
            .unwrap()
            .with_charset(Charset::DEFAULT);
        replacement.rename("Renamed").unwrap();
        changed.add(replacement).unwrap();
        assert!(put(&mut package, changed).unwrap());
        assert!(!package.is_signed());
        assert_eq!(
            load(&package)
                .unwrap()
                .unwrap()
                .get("renamed")
                .unwrap()
                .charset(),
            Some(Charset::DEFAULT)
        );
        assert!(remove(&mut package).unwrap().is_some());
        assert!(load(&package).unwrap().is_none());
        assert!(!embedding_enabled(package.get_part(&presentation_uri).unwrap().blob()).unwrap());
    }

    #[test]
    fn pitch_charset_zero_face_and_mixed_dialects_are_checked() {
        for value in [
            0_u8, 1, 2, 16, 17, 18, 32, 33, 34, 48, 49, 50, 64, 65, 66, 80, 81, 82,
        ] {
            let xml = format!(
                r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" pitchFamily="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
            );
            assert!(parse_raw(xml.as_bytes()).is_ok(), "pitchFamily={value}");
        }
        for value in [3_u8, 15, 19, 31, 35, 255] {
            let xml = format!(
                r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" pitchFamily="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
            );
            assert!(parse_raw(xml.as_bytes()).is_err(), "pitchFamily={value}");
        }
        for value in ["-128", "127"] {
            let xml = format!(
                r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" charset="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
            );
            assert!(parse_raw(xml.as_bytes()).is_ok(), "charset={value}");
        }
        for value in ["-129", "128"] {
            let xml = format!(
                r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" charset="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
            );
            assert!(parse_raw(xml.as_bytes()).is_err(), "charset={value}");
        }
        let mixed = format!(
            r#"<p:presentation xmlns:p="{STRICT_PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        assert!(parse_raw(mixed.as_bytes()).is_err());
    }

    #[test]
    fn malformed_unicode_duplicates_remain_numerically_repairable() {
        let mut package = package(Conformance::Transitional);
        let uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="Straße"/></p:embeddedFont><p:embeddedFont><p:font typeface="STRASSE"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        package
            .get_part_mut(&uri)
            .unwrap()
            .set_blob(xml.into_bytes());
        let mut fonts = load(&package).unwrap().unwrap();
        assert!(matches!(
            fonts.get("strasse"),
            Err(Error::AmbiguousFontName { matches: 2, .. })
        ));
        assert_eq!(fonts.get(0_usize).unwrap().name(), "Straße");
        fonts.remove(1_usize).unwrap();
        assert_eq!(fonts.get("STRASSE").unwrap().name(), "Straße");

        let mut authored = Fonts::new();
        authored.add(Font::new("é").unwrap()).unwrap();
        assert!(authored.add(Font::new("e\u{301}").unwrap()).is_err());
    }

    #[test]
    fn noncanonical_targets_and_every_main_profile_round_trip() {
        let mut noncanonical = package(Conformance::Transitional);
        let mut value = value();
        for face in &mut value.fonts[0].faces {
            if let Some(resource) = &mut face.resource {
                resource.part_name = format!("/custom/{}.bin", face.style.element());
            }
        }
        put_raw(&mut noncanonical, &value, Conformance::Transitional).unwrap();
        assert_eq!(load_raw(&noncanonical).unwrap().unwrap(), value);

        for content_type in [
            ct::PML_PRESENTATION_MAIN,
            ct::PML_SLIDESHOW_MAIN,
            ct::PML_TEMPLATE_MAIN,
            ct::PML_PRES_MACRO_MAIN,
            ct::PML_SLIDESHOW_MACRO_MAIN,
            ct::PML_TEMPLATE_MACRO_MAIN,
        ] {
            let mut package = package(Conformance::Transitional);
            let uri = PackURI::new("/ppt/presentation.xml").unwrap();
            package
                .get_part_mut(&uri)
                .unwrap()
                .set_content_type(content_type.into())
                .unwrap();
            let mut fonts = Fonts::new();
            fonts
                .add(
                    Font::from_face(
                        "Profile",
                        Face::new(Style::Regular, Data::powerpoint(eot(3)).unwrap()),
                    )
                    .unwrap(),
                )
                .unwrap();
            assert!(put(&mut package, fonts).unwrap(), "{content_type}");
            assert_eq!(load(&package).unwrap().unwrap().len(), 1, "{content_type}");
        }
    }

    #[test]
    fn shared_font_parts_survive_face_removal_and_reject_other_owners() {
        let mut package = package(Conformance::Transitional);
        let shared = RawResource {
            part_name: "/ppt/fonts/shared.fntdata".into(),
            content_type: FONT_DATA_CT.into(),
            data: Arc::new(vec![3; 64]),
        };
        let graph = RawFonts {
            fonts: vec![
                RawFont {
                    has_descriptor: true,
                    typeface: "First".into(),
                    panose: None,
                    pitch_family: None,
                    charset: None,
                    faces: vec![RawFace {
                        style: Style::Regular,
                        relationship_id: "rIdFontA".into(),
                        resource: Some(shared.clone()),
                    }],
                },
                RawFont {
                    has_descriptor: true,
                    typeface: "Second".into(),
                    panose: None,
                    pitch_family: None,
                    charset: None,
                    faces: vec![RawFace {
                        style: Style::Regular,
                        relationship_id: "rIdFontA".into(),
                        resource: Some(shared),
                    }],
                },
            ],
        };
        put_raw(&mut package, &graph, Conformance::Transitional).unwrap();
        assert_eq!(load_raw(&package).unwrap().unwrap().fonts.len(), 2);
        let mut fonts = load(&package).unwrap().unwrap();
        fonts.remove("First").unwrap();
        put(&mut package, fonts).unwrap();
        let font_uri = PackURI::new("/ppt/fonts/shared.fntdata").unwrap();
        assert!(package.contains_part(&font_uri));

        let owner_uri = PackURI::new("/ppt/unknown-owner.bin").unwrap();
        let mut owner = BlobPart::new(
            owner_uri.clone(),
            "application/octet-stream".into(),
            vec![1],
        );
        owner.rels_mut().add_relationship(
            "urn:shared-resource".into(),
            "fonts/shared.fntdata".into(),
            "rIdShared".into(),
            false,
        );
        package.add_part(Box::new(owner));
        assert!(matches!(load(&package), Err(Error::Invalid(_))));
        assert!(package.contains_part(&font_uri));
    }
}
