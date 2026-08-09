//! Typed, inert `PowerPoint` font models.

use crate::package::RecordLimits;
use std::sync::Arc;
use std::{fmt, ops::Deref};

/// The font-reference namespace owned by a collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Scope {
    /// `FontCollectionContainer` in the live document `Environment`.
    Base,
    /// `FontCollection10Container` in the live document `___PPT10` tag.
    International,
}

/// One of the four embedded-font facets defined by MS-PPT 2.9.9.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[repr(u8)]
pub enum Facet {
    Plain = 0,
    Bold = 1,
    Italic = 2,
    BoldItalic = 3,
}

impl TryFrom<u8> for Facet {
    type Error = crate::package::Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Plain),
            1 => Ok(Self::Bold),
            2 => Ok(Self::Italic),
            3 => Ok(Self::BoldItalic),
            _ => Err(crate::package::Error::Corrupted(
                "embedded font facet is outside 0..=3".into(),
            )),
        }
    }
}

/// A bounded, non-executing view of the fixed EOT header prefix.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EotMetadata {
    pub declared_size: u32,
    pub font_data_size: u32,
    pub version: u32,
    pub flags: u32,
    pub charset: u8,
    pub italic: bool,
    pub weight: u32,
    pub embedding_permissions: u16,
    pub magic: u16,
}

/// One embedded OpenType font facet. Its payload is always inert.
#[derive(Clone)]
pub struct SharedFontData(SharedFontDataInner);

#[derive(Clone)]
enum SharedFontDataInner {
    Vec(Arc<Vec<u8>>),
    Slice(Arc<[u8]>),
}

impl SharedFontData {
    #[must_use]
    pub fn as_slice(&self) -> &[u8] {
        match &self.0 {
            SharedFontDataInner::Vec(bytes) => bytes.as_slice(),
            SharedFontDataInner::Slice(bytes) => bytes.as_ref(),
        }
    }

    /// Whether two owners point at the same backing allocation and range.
    #[must_use]
    pub fn ptr_eq(&self, other: &Self) -> bool {
        match (&self.0, &other.0) {
            (SharedFontDataInner::Vec(left), SharedFontDataInner::Vec(right)) => {
                Arc::ptr_eq(left, right)
            },
            (SharedFontDataInner::Slice(left), SharedFontDataInner::Slice(right)) => {
                Arc::ptr_eq(left, right)
            },
            _ => false,
        }
    }

    /// Extract the uniquely owned byte vector behind this payload.
    ///
    /// # Errors
    ///
    /// Returns `Err(Self)` with the payload unchanged when it is not a
    /// uniquely owned `Vec` (shared owner or slice-backed).
    pub fn try_unwrap_vec(self) -> Result<Vec<u8>, Self> {
        match self.0 {
            SharedFontDataInner::Vec(bytes) => match Arc::try_unwrap(bytes) {
                Ok(owned) => Ok(owned),
                Err(shared) => Err(Self(SharedFontDataInner::Vec(shared))),
            },
            SharedFontDataInner::Slice(bytes) => Err(Self(SharedFontDataInner::Slice(bytes))),
        }
    }
}

impl From<Vec<u8>> for SharedFontData {
    fn from(bytes: Vec<u8>) -> Self {
        Self(SharedFontDataInner::Vec(Arc::new(bytes)))
    }
}

impl From<Arc<Vec<u8>>> for SharedFontData {
    fn from(bytes: Arc<Vec<u8>>) -> Self {
        Self(SharedFontDataInner::Vec(bytes))
    }
}

impl From<Arc<[u8]>> for SharedFontData {
    fn from(bytes: Arc<[u8]>) -> Self {
        Self(SharedFontDataInner::Slice(bytes))
    }
}

impl From<&[u8]> for SharedFontData {
    fn from(bytes: &[u8]) -> Self {
        Self(SharedFontDataInner::Slice(Arc::from(bytes)))
    }
}

impl AsRef<[u8]> for SharedFontData {
    fn as_ref(&self) -> &[u8] {
        self.as_slice()
    }
}

impl Deref for SharedFontData {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        self.as_slice()
    }
}

impl fmt::Debug for SharedFontData {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SharedFontData")
            .field("len", &self.len())
            .finish_non_exhaustive()
    }
}

impl PartialEq for SharedFontData {
    fn eq(&self, other: &Self) -> bool {
        self.as_slice() == other.as_slice()
    }
}

impl Eq for SharedFontData {}

/// One embedded OpenType font facet. Its payload is always inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedFont {
    /// Facet index retained for source compatibility (`0..=3`).
    pub style: u8,
    /// Exact EOT bytes. Litchi never loads, renders, registers, or executes them.
    pub data: SharedFontData,
}

impl EmbeddedFont {
    /// Create one embedded font facet from EOT bytes validated against the
    /// default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if `data` is not a structurally valid EOT 1.0 payload
    /// within the default limits.
    pub fn new(facet: Facet, data: impl Into<SharedFontData>) -> crate::package::Result<Self> {
        let payload = data.into();
        validate_eot_facet(payload.as_ref(), Limits::default())?;
        Ok(Self {
            style: facet as u8,
            data: payload,
        })
    }

    pub(crate) fn from_preserved(facet: Facet, data: impl Into<SharedFontData>) -> Self {
        Self {
            style: facet as u8,
            data: data.into(),
        }
    }

    /// Exact inert EOT bytes. Cloning the owner does not clone this payload.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.data
    }

    /// The typed facet for this embedded font.
    ///
    /// # Errors
    ///
    /// Returns an error if the retained facet index is outside `0..=3`.
    pub fn facet(&self) -> crate::package::Result<Facet> {
        Facet::try_from(self.style)
    }

    /// Inspect the fixed header of a structurally plausible uncompressed EOT
    /// 1.0 payload. This remains a cheap inert probe; the optional font crate
    /// performs authoritative SFNT and licensing validation.
    #[must_use]
    pub fn eot_metadata(&self) -> Option<EotMetadata> {
        validate_eot_facet(self.data.as_ref(), Limits::default()).ok()
    }
}

/// Font attributes from one `FontEntityAtom`.
#[allow(
    clippy::struct_excessive_bools,
    reason = "each bool mirrors one MS-PPT `FontEntityAtom` flag bit; merging them would obscure the format projection"
)]
#[derive(Debug, Clone)]
pub struct Font {
    /// Stable zero-based ordinal used by `FontIndexRef`/`FontIndexRef10`.
    pub index: u16,
    /// Producer-supplied `recInstance`, retained independently of the ordinal.
    pub raw_instance: u16,
    pub name: String,
    pub charset: u8,
    /// Raw byte; undefined bits 1..=7 are retained.
    pub font_flags: u8,
    pub embedded_subset: bool,
    /// Raw byte; the reserved high nibble is retained and ignored on read.
    pub font_type_flags: u8,
    pub raster: bool,
    pub device: bool,
    pub truetype: bool,
    pub no_substitution: bool,
    pub pitch_and_family: u8,
    pub embedded_fonts: Vec<EmbeddedFont>,
    pub(crate) source_name: Option<[u8; 64]>,
}

impl PartialEq for Font {
    fn eq(&self, other: &Self) -> bool {
        self.index == other.index
            && self.raw_instance == other.raw_instance
            && self.name == other.name
            && self.charset == other.charset
            && self.font_flags == other.font_flags
            && self.embedded_subset == other.embedded_subset
            && self.font_type_flags == other.font_type_flags
            && self.raster == other.raster
            && self.device == other.device
            && self.truetype == other.truetype
            && self.no_substitution == other.no_substitution
            && self.pitch_and_family == other.pitch_and_family
            && self.embedded_fonts == other.embedded_fonts
    }
}

impl Eq for Font {}

impl Font {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            index: 0,
            raw_instance: 0,
            name: name.into(),
            charset: 0,
            font_flags: 0,
            embedded_subset: false,
            font_type_flags: 0x04,
            raster: false,
            device: false,
            truetype: true,
            no_substitution: false,
            pitch_and_family: 0,
            embedded_fonts: Vec::new(),
            source_name: None,
        }
    }

    #[must_use]
    pub fn facet(&self, facet: Facet) -> Option<&EmbeddedFont> {
        self.embedded_fonts
            .iter()
            .find(|value| value.style == facet as u8)
    }
}

/// Parsed base or `PowerPoint` 10 font collection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontCollection {
    /// Retained compatibility projection of [`Self::scope`].
    pub international: bool,
    pub fonts: Vec<Font>,
}

impl FontCollection {
    #[must_use]
    pub fn new(scope: Scope) -> Self {
        Self {
            international: scope == Scope::International,
            fonts: Vec::new(),
        }
    }

    #[must_use]
    pub const fn scope(&self) -> Scope {
        if self.international {
            Scope::International
        } else {
            Scope::Base
        }
    }

    #[must_use]
    pub fn get(&self, index: u16) -> Option<&Font> {
        self.fonts.get(usize::from(index))
    }

    pub fn get_mut(&mut self, index: u16) -> Option<&mut Font> {
        self.fonts.get_mut(usize::from(index))
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.fonts.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.fonts.is_empty()
    }

    #[must_use]
    pub fn has_embedded_fonts(&self) -> bool {
        self.fonts
            .iter()
            .any(|font| !font.embedded_fonts.is_empty())
    }

    /// Append a font, assigning it the next ordinal, and return that ordinal.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection already holds 129 fonts or the
    /// ordinal exceeds `u16`.
    pub fn try_push(&mut self, mut font: Font) -> crate::package::Result<u16> {
        let index = u16::try_from(self.fonts.len())
            .map_err(|_err| crate::package::Error::Corrupted("font ordinal exceeds u16".into()))?;
        if index > 128 {
            return Err(crate::package::Error::Corrupted(
                "font collection exceeds the 129-font format limit".into(),
            ));
        }
        font.index = index;
        font.raw_instance = index;
        font.source_name = None;
        self.fonts.push(font);
        Ok(index)
    }

    /// Replace the font at `index`, retaining its raw instance, and return the
    /// previous font.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is not a known font ordinal.
    pub fn replace(&mut self, index: u16, mut font: Font) -> crate::package::Result<Font> {
        let slot = self.get_mut(index).ok_or_else(|| {
            crate::package::Error::Corrupted(format!("unknown font ordinal {index}"))
        })?;
        if font.name != slot.name {
            font.source_name = None;
        }
        font.index = index;
        font.raw_instance = slot.raw_instance;
        Ok(std::mem::replace(slot, font))
    }

    /// Validate and set one embedded facet on the font at `index`, returning
    /// the replaced facet if one existed.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is unknown or `data` is not a structurally
    /// valid EOT 1.0 payload within the default limits.
    pub fn set_facet(
        &mut self,
        index: u16,
        facet: Facet,
        data: impl Into<SharedFontData>,
    ) -> crate::package::Result<Option<EmbeddedFont>> {
        let payload = data.into();
        validate_eot_facet(payload.as_ref(), Limits::default())?;
        self.set_facet_preserved(index, facet, payload)
    }

    pub(crate) fn set_facet_preserved(
        &mut self,
        index: u16,
        facet: Facet,
        data: impl Into<SharedFontData>,
    ) -> crate::package::Result<Option<EmbeddedFont>> {
        let payload = data.into();
        let font = self.get_mut(index).ok_or_else(|| {
            crate::package::Error::Corrupted(format!("unknown font ordinal {index}"))
        })?;
        if let Some(position) = font
            .embedded_fonts
            .iter()
            .position(|value| value.style == facet as u8)
        {
            return Ok(Some(std::mem::replace(
                &mut font.embedded_fonts[position],
                EmbeddedFont::from_preserved(facet, payload),
            )));
        }
        let position = font
            .embedded_fonts
            .partition_point(|value| value.style < facet as u8);
        font.embedded_fonts
            .insert(position, EmbeddedFont::from_preserved(facet, payload));
        Ok(None)
    }

    /// Check that appending `font` would keep the collection valid under
    /// `limits`, without mutating it.
    ///
    /// # Errors
    ///
    /// Returns an error if the resulting collection would violate the format
    /// or `limits`.
    pub fn validate_append(&self, font: &Font, limits: Limits) -> crate::package::Result<()> {
        let mut candidate = self.clone();
        candidate.try_push(font.clone())?;
        super::validation::validate_authored_collection(&candidate, limits)
    }

    /// Check that replacing the font at `index` with `font` would keep the
    /// collection valid under `limits`, without mutating it.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is unknown or the resulting collection
    /// would violate the format or `limits`.
    pub fn validate_replacement(
        &self,
        index: u16,
        font: &Font,
        limits: Limits,
    ) -> crate::package::Result<()> {
        let mut candidate = self.clone();
        candidate.replace(index, font.clone())?;
        super::validation::validate_authored_collection(&candidate, limits)
    }

    /// Check that setting `facet` with `data` on the font at `index` would
    /// keep the collection valid under `limits`, without mutating it.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is unknown, `data` fails EOT validation, or
    /// the resulting collection would violate the format or `limits`.
    pub fn validate_facet(
        &self,
        index: u16,
        facet: Facet,
        data: &[u8],
        limits: Limits,
    ) -> crate::package::Result<()> {
        let mut candidate = self.clone();
        candidate.set_facet(index, facet, data.to_vec())?;
        super::validation::validate_authored_collection(&candidate, limits)
    }

    /// Remove `facet` from the font at `index`, returning it if present.
    ///
    /// # Errors
    ///
    /// Returns an error if `index` is not a known font ordinal.
    pub fn remove_facet(
        &mut self,
        index: u16,
        facet: Facet,
    ) -> crate::package::Result<Option<EmbeddedFont>> {
        let font = self.get_mut(index).ok_or_else(|| {
            crate::package::Error::Corrupted(format!("unknown font ordinal {index}"))
        })?;
        Ok(font
            .embedded_fonts
            .iter()
            .position(|value| value.style == facet as u8)
            .map(|position| font.embedded_fonts.remove(position)))
    }
}

/// `PowerPoint` 10 document-wide embedding settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FontEmbeddingFlags {
    /// Undefined bits 2..=31 are retained.
    pub raw: u32,
    pub subset: bool,
    pub subset_option_confirmed: bool,
}

impl FontEmbeddingFlags {
    #[must_use]
    pub const fn new(subset: bool, subset_option_confirmed: bool) -> Self {
        Self {
            raw: (subset as u32) | ((subset_option_confirmed as u32) << 1),
            subset,
            subset_option_confirmed,
        }
    }
}

/// Base and PP10 font owners from the exact live `DocumentContainer`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FontCollections {
    pub base: Option<FontCollection>,
    pub international: Option<FontCollection>,
    pub embedding_flags: Option<FontEmbeddingFlags>,
}

impl FontCollections {
    #[must_use]
    pub fn collection(&self, scope: Scope) -> Option<&FontCollection> {
        match scope {
            Scope::Base => self.base.as_ref(),
            Scope::International => self.international.as_ref(),
        }
    }

    pub fn collection_mut(&mut self, scope: Scope) -> Option<&mut FontCollection> {
        match scope {
            Scope::Base => self.base.as_mut(),
            Scope::International => self.international.as_mut(),
        }
    }

    #[must_use]
    pub fn get_base(&self, index: u16) -> Option<&Font> {
        self.base.as_ref()?.get(index)
    }

    #[must_use]
    pub fn get_international(&self, index: u16) -> Option<&Font> {
        self.international.as_ref()?.get(index)
    }

    pub fn has_embedded_fonts(&self) -> bool {
        self.base
            .as_ref()
            .is_some_and(FontCollection::has_embedded_fonts)
            || self
                .international
                .as_ref()
                .is_some_and(FontCollection::has_embedded_fonts)
    }

    pub fn base_font_count(&self) -> usize {
        self.base.as_ref().map_or(0, FontCollection::len)
    }
}

/// Composed record and font resource limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub records: RecordLimits,
    pub max_fonts_per_collection: usize,
    pub max_facets: usize,
    pub max_facet_bytes: usize,
    pub max_embedded_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            records: RecordLimits::default(),
            max_fonts_per_collection: 129,
            max_facets: 516,
            max_facet_bytes: 32 * 1024 * 1024,
            max_embedded_bytes: 128 * 1024 * 1024,
        }
    }
}

/// Validate the inert structural envelope required for newly authored EOT 1.0
/// data. This does not decompress, load, render, register, or execute the font.
///
/// # Errors
///
/// Returns an error if the payload exceeds the facet byte limit or is not a
/// structurally valid EOT 1.0 envelope.
pub fn validate_eot_facet(bytes: &[u8], limits: Limits) -> crate::package::Result<EotMetadata> {
    use crate::package::Error;

    if bytes.len() > limits.max_facet_bytes {
        return Err(Error::ResourceLimit(
            "embedded font facet exceeds its byte limit".into(),
        ));
    }
    if bytes.len() < 96 {
        return Err(Error::Corrupted("EOT 1.0 envelope is truncated".into()));
    }
    // The 96-byte minimum above guarantees the fixed 36-byte header prefix.
    let data = &bytes[..36];
    let u32_at = |offset: usize| {
        u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ])
    };
    let metadata = EotMetadata {
        declared_size: u32_at(0),
        font_data_size: u32_at(4),
        version: u32_at(8),
        flags: u32_at(12),
        charset: data[26],
        italic: data[27] != 0,
        weight: u32_at(28),
        embedding_permissions: u16::from_le_bytes([data[32], data[33]]),
        magic: u16::from_le_bytes([data[34], data[35]]),
    };
    let declared_size = usize::try_from(metadata.declared_size)
        .map_err(|_err| Error::Corrupted("EOT declared size overflows this platform".into()))?;
    let font_data_size = usize::try_from(metadata.font_data_size)
        .map_err(|_err| Error::Corrupted("EOT font size overflows this platform".into()))?;
    if metadata.magic != 0x504c || metadata.version != 0x0001_0000 || declared_size != bytes.len() {
        return Err(Error::Corrupted(
            "EOT header magic, version, or declared size is invalid".into(),
        ));
    }
    let mut cursor = 82usize;
    for index in 0..4 {
        let length_end = cursor
            .checked_add(2)
            .ok_or_else(|| Error::Corrupted("EOT name offset overflow".into()))?;
        let length_bytes = bytes
            .get(cursor..length_end)
            .ok_or_else(|| Error::Corrupted("EOT name length is truncated".into()))?;
        let length = usize::from(u16::from_le_bytes([length_bytes[0], length_bytes[1]]));
        if length % 2 != 0 {
            return Err(Error::Corrupted(
                "EOT UTF-16 name has an odd byte length".into(),
            ));
        }
        cursor = length_end
            .checked_add(length)
            .ok_or_else(|| Error::Corrupted("EOT name range overflow".into()))?;
        let name_start = length_end;
        let name = bytes
            .get(name_start..cursor)
            .ok_or_else(|| Error::Corrupted("EOT name is truncated".into()))?;
        if char::decode_utf16(
            name.chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
        )
        .any(|value| value.is_err())
        {
            return Err(Error::Corrupted(
                "EOT name contains malformed UTF-16".into(),
            ));
        }
        if index != 3 {
            let padding = bytes
                .get(cursor..cursor + 2)
                .ok_or_else(|| Error::Corrupted("EOT name padding is truncated".into()))?;
            if padding != [0, 0] {
                return Err(Error::Corrupted("EOT name padding is nonzero".into()));
            }
            cursor += 2;
        }
    }
    if cursor.checked_add(font_data_size) != Some(declared_size) {
        return Err(Error::Corrupted(
            "EOT font data size does not match its trailing payload".into(),
        ));
    }
    Ok(metadata)
}
