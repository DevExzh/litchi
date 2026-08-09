//! Public embedded-font values and detached collection semantics.

use super::{
    FONT_DATA_CT, FONT_REL, FONT_TTF_CT, MAX_FONTS, MAX_TOTAL_FONT_BYTES, PML, REL_NS,
    STRICT_FONT_REL, STRICT_PML, STRICT_REL_NS, codec, invalid, limit,
};
use crate::error::{Error, Result};
use caseless::Caseless;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;
use unicode_normalization::UnicodeNormalization;

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
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
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
    #[must_use]
    pub fn fs_type(self) -> u16 {
        self.fs_type
    }

    /// Return the mutually exclusive embedding permission.
    #[must_use]
    pub fn permission(self) -> Permission {
        self.permission
    }

    /// Return compact independent restrictions.
    #[must_use]
    pub fn restrictions(self) -> Restrictions {
        self.restrictions
    }

    /// Report installable embedding permission.
    #[must_use]
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
    pub(super) fn pml(self) -> &'static str {
        match self {
            Self::Transitional => PML,
            Self::Strict => STRICT_PML,
        }
    }
    pub(super) fn rel_ns(self) -> &'static str {
        match self {
            Self::Transitional => REL_NS,
            Self::Strict => STRICT_REL_NS,
        }
    }
    pub(super) fn font_rel(self) -> &'static str {
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
    pub(super) fn element(self) -> &'static str {
        match self {
            Self::Regular => "regular",
            Self::Bold => "bold",
            Self::Italic => "italic",
            Self::BoldItalic => "boldItalic",
        }
    }
    pub(super) fn rank(self) -> u8 {
        match self {
            Self::Regular => 0,
            Self::Bold => 1,
            Self::Italic => 2,
            Self::BoldItalic => 3,
        }
    }
    pub(super) fn parse_raw(value: &str) -> Option<Self> {
        match value {
            "regular" => Some(Self::Regular),
            "bold" => Some(Self::Bold),
            "italic" => Some(Self::Italic),
            "boldItalic" => Some(Self::BoldItalic),
            _ => None,
        }
    }
}

/// Font-pitch component of `DrawingML` `ST_PitchFamily`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Pitch {
    Default,
    Fixed,
    Variable,
}

impl Pitch {
    pub(super) const fn wire(self) -> u8 {
        match self {
            Self::Default => 0,
            Self::Fixed => 1,
            Self::Variable => 2,
        }
    }
}

/// Font-family component of `DrawingML` `ST_PitchFamily`.
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
    pub(super) const fn wire(self) -> u8 {
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

/// The closed 18-value `DrawingML` pitch/family domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PitchFamily {
    pitch: Pitch,
    family: Family,
}

impl PitchFamily {
    /// Combine typed pitch and family values; every combination is valid.
    #[must_use]
    pub const fn new(pitch: Pitch, family: Family) -> Self {
        Self { pitch, family }
    }

    /// Return the typed pitch component.
    #[must_use]
    pub const fn pitch(self) -> Pitch {
        self.pitch
    }

    /// Return the typed family component.
    #[must_use]
    pub const fn family(self) -> Family {
        self.family
    }

    pub(super) const fn wire(self) -> u8 {
        self.family.wire() * 16 + self.pitch.wire()
    }

    pub(super) fn from_wire(value: u8) -> Result<Self> {
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
    #[must_use]
    pub const fn new(bytes: [u8; 10]) -> Self {
        Self(bytes)
    }

    /// Borrow the ten classification bytes.
    #[must_use]
    pub const fn bytes(&self) -> &[u8; 10] {
        &self.0
    }

    /// Move out the ten classification bytes.
    #[must_use]
    pub const fn into_bytes(self) -> [u8; 10] {
        self.0
    }
}

impl From<[u8; 10]> for Panose {
    fn from(value: [u8; 10]) -> Self {
        Self::new(value)
    }
}

/// A Windows font charset code with private `PresentationML` wire conversion.
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
    #[must_use]
    pub const fn new(code: u8) -> Self {
        Self(code)
    }

    /// Return the Windows charset code without a string conversion.
    #[must_use]
    pub const fn code(self) -> u8 {
        self.0
    }

    pub(super) const fn from_wire(value: i8) -> Self {
        Self(value.cast_unsigned())
    }

    pub(super) const fn wire(self) -> i8 {
        self.0.cast_signed()
    }
}

/// Physical font representation permitted by `PresentationML`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Format {
    /// PowerPoint-compatible storage (`application/x-fontdata`).
    PowerPoint,
    /// Standards-only `application/x-font-ttf` preservation.
    Standard,
}

impl Format {
    pub(super) fn content_type(self) -> &'static str {
        match self {
            Self::PowerPoint => FONT_DATA_CT,
            Self::Standard => FONT_TTF_CT,
        }
    }

    pub(super) fn parse(content_type: &str) -> Result<Self> {
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
    pub(super) format: Format,
    pub(super) bytes: Arc<Vec<u8>>,
}

impl Data {
    /// Adopt and validate an owned font container.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(bytes: Vec<u8>, format: Format) -> Result<Self> {
        Self::shared(Arc::new(bytes), format)
    }

    /// Validate and share an existing immutable allocation without copying it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shared(bytes: Arc<Vec<u8>>, format: Format) -> Result<Self> {
        codec::validate_font_bytes(&bytes)?;
        match format {
            Format::PowerPoint => codec::validate_eot(&bytes)?,
            Format::Standard => codec::validate_sfnt(&bytes)?,
        }
        Ok(Self { format, bytes })
    }

    /// Adopt a PowerPoint-compatible EOT/MTX container.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn powerpoint(bytes: Vec<u8>) -> Result<Self> {
        Self::new(bytes, Format::PowerPoint)
    }

    /// Preserve standards-only `application/x-font-ttf` storage explicitly.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn standard(bytes: Vec<u8>) -> Result<Self> {
        Self::new(bytes, Format::Standard)
    }

    /// Preserve a bounded producer payload already present in a loaded package.
    pub(super) fn preserve(bytes: Arc<Vec<u8>>, format: Format) -> Result<Self> {
        codec::validate_font_bytes(&bytes)?;
        Ok(Self { format, bytes })
    }

    /// Return the physical representation.
    #[must_use]
    pub fn format(&self) -> Format {
        self.format
    }

    /// Borrow the inert bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    /// Move out the shared allocation without copying it.
    #[must_use]
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
pub(super) struct Source {
    pub(super) relationship_id: String,
    pub(super) part_name: String,
}

/// One typed style face and its inert program.
#[derive(Debug, Clone)]
pub struct Face {
    pub(super) style: Style,
    pub(super) data: Data,
    pub(super) source: Option<Source>,
}

impl Face {
    /// Pair a typed face style with an owned or shared font program.
    #[must_use]
    pub fn new(style: Style, data: Data) -> Self {
        Self {
            style,
            data,
            source: None,
        }
    }

    /// Return the schema-level style.
    #[must_use]
    pub fn style(&self) -> Style {
        self.style
    }

    /// Borrow the inert font program.
    #[must_use]
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
    pub(super) typeface: String,
    pub(super) key: String,
    pub(super) panose: Option<Panose>,
    pub(super) pitch_family: Option<PitchFamily>,
    pub(super) charset: Option<Charset>,
    pub(super) faces: Vec<Face>,
}

impl Font {
    /// Construct a typeface. The schema permits a descriptor with no faces;
    /// add one concisely with [`Font::with`] or [`Font::put`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(typeface: impl Into<String>) -> Result<Self> {
        let typeface = typeface.into();
        codec::validate_typeface(&typeface)?;
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
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_face(typeface: impl Into<String>, face: Face) -> Result<Self> {
        let mut font = Self::new(typeface)?;
        font.put(face)?;
        Ok(font)
    }

    /// Add a face with builder-style chaining.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn with(mut self, face: Face) -> Result<Self> {
        self.put(face)?;
        Ok(self)
    }

    /// Return the producer spelling of the typeface.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.typeface
    }

    /// Return the optional ten-byte PANOSE classification.
    #[must_use]
    pub fn panose(&self) -> Option<Panose> {
        self.panose
    }

    /// Return the optional combined pitch/family byte.
    #[must_use]
    pub fn pitch_family(&self) -> Option<PitchFamily> {
        self.pitch_family
    }

    /// Return the optional Windows character-set byte.
    #[must_use]
    pub fn charset(&self) -> Option<Charset> {
        self.charset
    }

    /// Return faces in the schema-defined style order.
    #[must_use]
    pub fn faces(&self) -> &[Face] {
        &self.faces
    }

    /// Set the PANOSE classification with struct-update-like builder syntax.
    #[must_use]
    pub fn with_panose(mut self, value: impl Into<Panose>) -> Self {
        self.panose = Some(value.into());
        self
    }

    /// Replace or clear the PANOSE classification.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_panose(&mut self, value: Option<Panose>) -> Option<Panose> {
        std::mem::replace(&mut self.panose, value)
    }

    /// Set the combined pitch/family byte.
    #[must_use]
    pub fn with_pitch_family(mut self, value: PitchFamily) -> Self {
        self.pitch_family = Some(value);
        self
    }

    /// Replace or clear the compact pitch/family classification.
    pub fn set_pitch_family(&mut self, value: Option<PitchFamily>) -> Option<PitchFamily> {
        std::mem::replace(&mut self.pitch_family, value)
    }

    /// Set the Windows character-set byte.
    #[must_use]
    pub fn with_charset(mut self, value: Charset) -> Self {
        self.charset = Some(value);
        self
    }

    /// Replace or clear the Windows charset code.
    pub fn set_charset(&mut self, value: Option<Charset>) -> Option<Charset> {
        std::mem::replace(&mut self.charset, value)
    }

    /// Add or replace one typed style face, returning the previous face.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[must_use]
    pub fn get(&self, style: Style) -> Option<&Face> {
        self.faces.iter().find(|face| face.style == style)
    }

    /// Remove and return one face. Face-less descriptors remain schema-valid.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn rename(&mut self, typeface: impl Into<String>) -> Result<String> {
        let typeface = typeface.into();
        codec::validate_typeface(&typeface)?;
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
    pub(super) first: usize,
    pub(super) matches: usize,
}

#[derive(Debug, Clone)]
pub struct Fonts {
    pub(super) fonts: Vec<Font>,
    by_name: HashMap<String, NameIndex>,
}

impl Fonts {
    /// Construct an empty collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[must_use]
    pub fn new() -> Self {
        Self {
            fonts: Vec::new(),
            by_name: HashMap::new(),
        }
    }

    /// Return the number of typefaces.
    #[must_use]
    pub fn len(&self) -> usize {
        self.fonts.len()
    }

    /// Report whether no fonts are embedded.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.fonts.is_empty()
    }

    /// Iterate in source order.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Font> {
        self.fonts.iter()
    }

    /// Select by semantic typeface or checked numeric position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn get<'a, 'k>(&'a self, key: impl Into<Key<'k>>) -> Result<&'a Font> {
        let index = self.offset(key.into())?;
        self.fonts.get(index).ok_or(Error::FontIndexOutOfBounds {
            index,
            len: self.fonts.len(),
        })
    }

    /// Append an owned typeface without copying its programs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    #[must_use]
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

    pub(super) fn reindex(&mut self) -> Result<()> {
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

pub(super) fn name_key(value: &str) -> String {
    value.chars().nfd().default_case_fold().nfd().collect()
}

/// Validate the detached semantic collection before package publication.
pub(super) fn validate_fonts(value: &Fonts, require_unique_names: bool) -> Result<()> {
    if value.fonts.len() > MAX_FONTS {
        return Err(limit("embedded fonts"));
    }
    let mut names = HashSet::new();
    let mut resources = HashSet::new();
    let mut total = 0usize;
    for font in &value.fonts {
        codec::validate_typeface(&font.typeface)?;
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
            codec::validate_font_bytes(&face.data.bytes)?;
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
