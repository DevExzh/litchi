//! Shared font values and policies.

use std::collections::HashMap;

use roaring::RoaringBitmap;

/// A validated set of Unicode scalar values needed from one font face.
///
/// The private bitmap prevents surrogate code points and values above
/// `U+10FFFF` from entering the embedding pipeline.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Glyphs(RoaringBitmap);

impl Glyphs {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn insert(&mut self, value: char) -> bool {
        self.0.insert(u32::from(value))
    }

    pub fn len(&self) -> u64 {
        self.0.len()
    }

    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    pub fn iter(&self) -> impl Iterator<Item = u32> + '_ {
        self.0.iter()
    }
}

impl Extend<char> for Glyphs {
    fn extend<T: IntoIterator<Item = char>>(&mut self, iter: T) {
        for value in iter {
            self.insert(value);
        }
    }
}

impl FromIterator<char> for Glyphs {
    fn from_iter<T: IntoIterator<Item = char>>(iter: T) -> Self {
        let mut glyphs = Self::new();
        glyphs.extend(iter);
        glyphs
    }
}

impl std::ops::BitOrAssign for Glyphs {
    fn bitor_assign(&mut self, rhs: Self) {
        self.0 |= rhs.0;
    }
}

/// The four Office font-face styles.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum Style {
    #[default]
    Regular,
    Bold,
    Italic,
    BoldItalic,
}

impl Style {
    pub const fn from_flags(bold: bool, italic: bool) -> Self {
        match (bold, italic) {
            (false, false) => Self::Regular,
            (true, false) => Self::Bold,
            (false, true) => Self::Italic,
            (true, true) => Self::BoldItalic,
        }
    }

    pub const fn is_bold(self) -> bool {
        matches!(self, Self::Bold | Self::BoldItalic)
    }

    pub const fn is_italic(self) -> bool {
        matches!(self, Self::Italic | Self::BoldItalic)
    }
}

/// One system-font face request used as a glyph-map key.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Request {
    family: String,
    style: Style,
}

impl Request {
    pub fn new(family: impl Into<String>, style: Style) -> Self {
        Self {
            family: family.into(),
            style,
        }
    }

    pub fn regular(family: impl Into<String>) -> Self {
        Self::new(family, Style::Regular)
    }

    pub fn family(&self) -> &str {
        &self.family
    }

    pub const fn style(&self) -> Style {
        self.style
    }
}

/// Font-family requests keyed by their producer-visible names.
pub type GlyphMap = HashMap<Request, Glyphs>;

/// Trait for document types that can collect all glyphs (characters) used in the document.
///
/// This is used to determine which fonts need to be embedded and which glyphs
/// should be included in font subsets.
///
/// Uses `RoaringBitmap` instead of `HashSet<char>` for better cache locality and memory efficiency.
/// The bitmap stores Unicode code points (u32 values from chars).
pub trait CollectGlyphs {
    /// Returns typed font-face requests and their Unicode scalar values.
    fn collect_glyphs(&self) -> GlyphMap;
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontData {
    pub name: String,
    pub data: Vec<u8>,
    pub index: u32,
    pub properties: Option<FontProperties>,
}

/// A validated OpenType `OS/2.fsType` embedding permission.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct License(u16);

/// The mutually exclusive embedding permission encoded by [`License`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Permission {
    Installable,
    Restricted,
    PreviewPrint,
    Editable,
}

/// Independent embedding restrictions encoded by [`License`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Restrictions(u16);

impl Restrictions {
    pub const NO_SUBSETTING: Self = Self(0x0100);
    pub const BITMAP_ONLY: Self = Self(0x0200);

    const MASK: u16 = Self::NO_SUBSETTING.0 | Self::BITMAP_ONLY.0;

    const fn from_license(bits: u16) -> Self {
        Self(bits & Self::MASK)
    }

    pub const fn bits(self) -> u16 {
        self.0
    }

    pub const fn contains(self, other: Self) -> bool {
        self.0 & other.0 == other.0
    }
}

impl std::ops::BitOr for Restrictions {
    type Output = Self;

    fn bitor(self, rhs: Self) -> Self::Output {
        Self(self.0 | rhs.0)
    }
}

impl std::ops::BitOrAssign for Restrictions {
    fn bitor_assign(&mut self, rhs: Self) {
        self.0 |= rhs.0;
    }
}

impl std::ops::BitAnd for Restrictions {
    type Output = Self;

    fn bitand(self, rhs: Self) -> Self::Output {
        Self(self.0 & rhs.0)
    }
}

impl License {
    const DEFINED: u16 = 0x0002 | 0x0004 | 0x0008 | 0x0100 | 0x0200;
    const PERMISSION_MASK: u16 = 0x0002 | 0x0004 | 0x0008;

    /// Validate an OpenType `fsType` bit field.
    pub fn new(bits: u16) -> Result<Self, LicenseError> {
        let reserved = bits & !Self::DEFINED;
        if reserved != 0 {
            return Err(LicenseError::ReservedBits(reserved));
        }
        if (bits & Self::PERMISSION_MASK).count_ones() > 1 {
            return Err(LicenseError::ContradictoryPermission);
        }
        Ok(Self(bits))
    }

    /// Interpret `fsType` using the rules assigned to an OS/2 table version.
    ///
    /// Versions zero and one define only the permission nibble, so later bits
    /// are ignored. Versions zero through two permit multiple permission bits;
    /// their effective value is the least restrictive permission present.
    pub(crate) fn from_os2(version: u16, bits: u16) -> Result<Self, LicenseError> {
        let assigned = if version <= 1 { bits & 0x000F } else { bits };
        let reserved = assigned & !Self::DEFINED;
        if reserved != 0 {
            return Err(LicenseError::ReservedBits(reserved));
        }

        if version >= 3 {
            return Self::new(assigned);
        }

        let permission = if assigned & 0x0008 != 0 {
            0x0008
        } else if assigned & 0x0004 != 0 {
            0x0004
        } else {
            assigned & 0x0002
        };
        Ok(Self((assigned & Restrictions::MASK) | permission))
    }

    pub const fn bits(self) -> u16 {
        self.0
    }

    pub const fn permission(self) -> Permission {
        if self.0 & 0x0002 != 0 {
            Permission::Restricted
        } else if self.0 & 0x0004 != 0 {
            Permission::PreviewPrint
        } else if self.0 & 0x0008 != 0 {
            Permission::Editable
        } else {
            Permission::Installable
        }
    }

    pub const fn restrictions(self) -> Restrictions {
        Restrictions::from_license(self.0)
    }

    /// Whether the license permits embedding outline data.
    pub const fn may_embed_outlines(self) -> bool {
        !matches!(self.permission(), Permission::Restricted)
            && !self.restrictions().contains(Restrictions::BITMAP_ONLY)
    }

    pub const fn may_subset(self) -> bool {
        !self.restrictions().contains(Restrictions::NO_SUBSETTING)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
pub enum LicenseError {
    #[error("font license contains reserved fsType bits 0x{0:04X}")]
    ReservedBits(u16),
    #[error("font license contains contradictory embedding permissions")]
    ContradictoryPermission,
}

/// The fixed ten-byte PANOSE classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Panose([u8; 10]);

impl Panose {
    pub const fn new(bytes: [u8; 10]) -> Self {
        Self(bytes)
    }

    pub const fn bytes(&self) -> &[u8; 10] {
        &self.0
    }

    pub const fn into_bytes(self) -> [u8; 10] {
        self.0
    }
}

impl From<[u8; 10]> for Panose {
    fn from(value: [u8; 10]) -> Self {
        Self::new(value)
    }
}

/// A Windows font charset code retained without string conversion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Charset(u8);

impl Charset {
    pub const ANSI: Self = Self(0);
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

    pub const fn new(code: u8) -> Self {
        Self(code)
    }

    pub const fn code(self) -> u8 {
        self.0
    }

    /// PresentationML represents this byte with XML Schema's signed `byte`.
    pub const fn signed(self) -> i8 {
        self.0 as i8
    }
}

/// The compact Office font-family classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Family {
    Auto,
    Roman,
    Swiss,
    Modern,
    Script,
    Decorative,
}

/// Whether glyph advances are fixed or variable.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Pitch {
    Default,
    Fixed,
    Variable,
}

/// OpenType Unicode-range and code-page signature words.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Signature {
    unicode: [u32; 4],
    code_pages: [u32; 2],
}

impl Signature {
    pub const fn new(unicode: [u32; 4], code_pages: [u32; 2]) -> Self {
        Self {
            unicode,
            code_pages,
        }
    }

    pub const fn unicode(&self) -> &[u32; 4] {
        &self.unicode
    }

    pub const fn code_pages(&self) -> &[u32; 2] {
        &self.code_pages
    }
}

/// Typed metadata needed by format-specific font publishers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FontProperties {
    license: License,
    panose: Panose,
    charset: Option<Charset>,
    family: Family,
    pitch: Pitch,
    signature: Signature,
}

impl FontProperties {
    pub const fn new(
        license: License,
        panose: Panose,
        charset: Option<Charset>,
        family: Family,
        pitch: Pitch,
        signature: Signature,
    ) -> Self {
        Self {
            license,
            panose,
            charset,
            family,
            pitch,
            signature,
        }
    }

    pub const fn license(self) -> License {
        self.license
    }

    pub const fn panose(self) -> Panose {
        self.panose
    }

    /// Return a character-set hint only when OS/2 evidence is unambiguous.
    pub const fn charset(self) -> Option<Charset> {
        self.charset
    }

    pub const fn family(self) -> Family {
        self.family
    }

    pub const fn pitch(self) -> Pitch {
        self.pitch
    }

    pub const fn signature(self) -> Signature {
        self.signature
    }
}

#[derive(Debug, thiserror::Error)]
pub enum FontError {
    #[error("Font not found: {0}")]
    NotFound(String),
    #[error("Invalid font data")]
    InvalidData,
    #[error("invalid or unreadable font face index {0}")]
    InvalidFaceIndex(u32),
    #[error("unsupported OpenType OS/2 table version {0}")]
    UnsupportedOs2Version(u16),
    #[error(
        "truncated OpenType OS/2 version {version} table: expected at least {expected} bytes, got {actual}"
    )]
    TruncatedOs2 {
        version: u16,
        expected: usize,
        actual: usize,
    },
    #[error(transparent)]
    InvalidLicense(#[from] LicenseError),
    #[error("Subsetting failed: {0}")]
    SubsettingFailed(String),
    #[error("font embedding failed: {0}")]
    EmbeddingFailed(String),
    #[error("font '{name}' forbids outline embedding")]
    EmbeddingForbidden { name: String },
    #[error("font '{name}' permits bitmap embedding only")]
    BitmapOnly { name: String },
    #[error("font '{name}' has no readable OpenType OS/2 metadata")]
    MissingProperties { name: String },
    #[error("font embedding requires one standalone OpenType face")]
    RequiresStandaloneFace,
    #[error("font allocation for {resource} failed: {source}")]
    Allocation {
        resource: &'static str,
        #[source]
        source: std::collections::TryReserveError,
    },
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn license_rejects_reserved_and_contradictory_bits() {
        assert_eq!(
            License::new(0x8000),
            Err(LicenseError::ReservedBits(0x8000))
        );
        assert_eq!(
            License::new(0x0006),
            Err(LicenseError::ContradictoryPermission)
        );
    }

    #[test]
    fn license_exposes_compact_typed_policy() {
        let license = License::new(0x0108).expect("valid editable license");
        assert_eq!(license.permission(), Permission::Editable);
        assert!(license.may_embed_outlines());
        assert!(!license.may_subset());
        assert_eq!(license.bits(), 0x0108);
    }

    #[test]
    fn legacy_os2_license_rules_are_version_aware() {
        let version_one = License::from_os2(1, 0x030C).expect("legacy license");
        assert_eq!(version_one.bits(), 0x0008);
        assert_eq!(version_one.permission(), Permission::Editable);
        assert!(version_one.may_subset());

        let version_two = License::from_os2(2, 0x010C).expect("legacy license");
        assert_eq!(version_two.bits(), 0x0108);
        assert!(!version_two.may_subset());

        assert_eq!(
            License::from_os2(3, 0x000C),
            Err(LicenseError::ContradictoryPermission)
        );
    }

    #[test]
    fn charset_preserves_unsigned_and_presentationml_views() {
        assert_eq!(Charset::GB2312.code(), 134);
        assert_eq!(Charset::GB2312.signed(), -122);
    }

    #[test]
    fn glyph_sets_accept_only_unicode_scalars_and_union_by_value() {
        let mut glyphs = "A😀".chars().collect::<Glyphs>();
        let other = "AB".chars().collect::<Glyphs>();
        glyphs |= other;
        assert_eq!(glyphs.iter().collect::<Vec<_>>(), [65, 66, 0x1F600]);
    }
}
