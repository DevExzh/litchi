#![forbid(unsafe_code)]
#![deny(missing_docs)]
//! Exact, typed legacy code-page selection and text conversion.
//!
//! [`Page`] is a one-byte validated identifier. Unsupported identifiers are
//! rejected during construction; decoding is strict unless the caller names
//! the lossy recovery path explicitly. Record-level terminators and other
//! container rules remain the responsibility of the concrete format crate.
//! [`Mbcs`] excludes wide UTF-16 pages from byte-stream records, while [`Ansi`]
//! further narrows the value to the exact `[MS-OSHARED]` ANSI page set.

use std::borrow::Cow;
use std::fmt;

use encoding_rs::Encoding;

/// Code-page selection or conversion failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// The numeric identifier has no exact supported codec.
    Unsupported(u32),
    /// The input byte sequence is malformed for the selected page.
    Invalid(u32),
    /// The Unicode input cannot be represented by the selected page.
    Unmappable(u32),
}

impl Error {
    /// Numeric code-page identifier associated with the failure.
    pub const fn page(self) -> u32 {
        match self {
            Self::Unsupported(page) | Self::Invalid(page) | Self::Unmappable(page) => page,
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Unsupported(page) => write!(formatter, "unsupported code page {page}"),
            Self::Invalid(page) => write!(formatter, "invalid byte sequence for code page {page}"),
            Self::Unmappable(page) => {
                write!(formatter, "text is not representable in code page {page}")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Validated legacy code-page identifier.
///
/// The value stores a private one-byte discriminant and resolves the static
/// codec with an exhaustive match, keeping the capability compact enough to
/// embed directly in parser state while making invalid states unrepresentable.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Page(Kind);

impl Page {
    /// UTF-16 little endian (Windows code page 1200).
    pub const UTF_16LE: Self = Self(Kind::Utf16Le);
    /// UTF-16 big endian (Windows code page 1201).
    pub const UTF_16BE: Self = Self(Kind::Utf16Be);
    /// Windows-1252.
    pub const WINDOWS_1252: Self = Self(Kind::Windows1252);
    /// Macintosh Roman.
    pub const MACINTOSH: Self = Self(Kind::Macintosh);
    /// UTF-8 (Windows code page 65001).
    pub const UTF_8: Self = Self(Kind::Utf8);

    /// Validates a numeric code-page identifier.
    ///
    /// Identifiers for which the backing codec would only be an approximation
    /// are rejected. In particular, CP437/CP850 are not treated as IBM866 and
    /// UTF-7 is not treated as UTF-8.
    pub fn new(page: u32) -> Option<Self> {
        Kind::from_id(page).map(Self)
    }

    /// Validates a numeric identifier and preserves the unsupported value in
    /// a typed error.
    pub fn require(page: u32) -> Result<Self, Error> {
        Self::new(page).ok_or(Error::Unsupported(page))
    }

    /// Numeric Windows code-page identifier.
    pub const fn id(self) -> u32 {
        self.0.id() as u32
    }

    /// Exact 16-bit storage form of this supported identifier.
    pub const fn id16(self) -> u16 {
        self.0.id()
    }

    /// Canonical codec name.
    pub fn name(self) -> &'static str {
        self.codec().name()
    }

    /// Strictly decodes bytes without BOM handling.
    ///
    /// Malformed input returns [`Error::Invalid`] rather than inserting a
    /// replacement character. The returned text borrows the input when the
    /// codec can do so safely.
    pub fn decode<'a>(self, bytes: &'a [u8]) -> Result<Cow<'a, str>, Error> {
        let (text, malformed) = self.codec().decode_without_bom_handling(bytes);
        if malformed {
            Err(Error::Invalid(self.id()))
        } else {
            Ok(text)
        }
    }

    /// Decodes bytes with replacement characters for malformed sequences.
    ///
    /// Use this only when the owning format explicitly defines recovery or a
    /// diagnostic records that recovery occurred.
    pub fn decode_lossy<'a>(self, bytes: &'a [u8]) -> Cow<'a, str> {
        self.recover(bytes).0
    }

    /// Decodes with replacement and reports whether recovery was required.
    ///
    /// This is intended for formats such as VBA that retain malformed source
    /// bytes and expose the recovery status to callers.
    pub fn recover<'a>(self, bytes: &'a [u8]) -> (Cow<'a, str>, bool) {
        self.codec().decode_without_bom_handling(bytes)
    }

    /// Strictly encodes text into this code page.
    ///
    /// Unrepresentable text returns [`Error::Unmappable`] rather than writing
    /// replacement bytes. The output may borrow an already-compatible input.
    pub fn encode<'a>(self, text: &'a str) -> Result<Cow<'a, [u8]>, Error> {
        if let Some(order) = self.0.utf16_order() {
            return Ok(Cow::Owned(encode_utf16(text, order)));
        }
        let (bytes, _, unmappable) = self.codec().encode(text);
        if unmappable {
            Err(Error::Unmappable(self.id()))
        } else {
            Ok(bytes)
        }
    }

    fn codec(self) -> &'static Encoding {
        self.0.codec()
    }
}

impl TryFrom<u32> for Page {
    type Error = Error;

    fn try_from(page: u32) -> Result<Self, Self::Error> {
        Self::require(page)
    }
}

impl From<Page> for u32 {
    fn from(page: Page) -> Self {
        page.id()
    }
}

impl fmt::Debug for Page {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Page")
            .field("id", &self.id())
            .field("name", &self.name())
            .finish()
    }
}

/// Validated byte-stream code page.
///
/// Unlike [`Page`], this capability excludes UTF-16LE and UTF-16BE. Formats
/// whose records use a one-byte NUL terminator can therefore store an `Mbcs`
/// without admitting a wide-character page into that path.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Mbcs(Page);

impl Mbcs {
    /// Shift JIS (Windows code page 932).
    pub const SHIFT_JIS: Self = Self(Page(Kind::ShiftJis));
    /// Windows-1252.
    pub const WINDOWS_1252: Self = Self(Page::WINDOWS_1252);
    /// Macintosh Roman.
    pub const MACINTOSH: Self = Self(Page::MACINTOSH);
    /// UTF-8 (Windows code page 65001).
    pub const UTF_8: Self = Self(Page::UTF_8);

    /// Validates a byte-stream code-page identifier.
    pub fn new(page: u32) -> Option<Self> {
        let page = Page::new(page)?;
        match page.0 {
            Kind::Utf16Le | Kind::Utf16Be => None,
            _ => Some(Self(page)),
        }
    }

    /// Validates a byte-stream identifier and retains the unsupported value.
    pub fn require(page: u32) -> Result<Self, Error> {
        Self::new(page).ok_or(Error::Unsupported(page))
    }

    /// General code-page capability represented by this value.
    pub const fn page(self) -> Page {
        self.0
    }

    /// Numeric Windows code-page identifier.
    pub const fn id(self) -> u32 {
        self.0.id()
    }

    /// Exact 16-bit storage form of this supported identifier.
    pub const fn id16(self) -> u16 {
        self.0.id16()
    }

    /// Canonical codec name.
    pub fn name(self) -> &'static str {
        self.0.name()
    }

    /// Strictly decodes bytes without BOM or terminator handling.
    pub fn decode<'a>(self, bytes: &'a [u8]) -> Result<Cow<'a, str>, Error> {
        self.0.decode(bytes)
    }

    /// Decodes bytes with replacement for explicitly recoverable formats.
    pub fn decode_lossy<'a>(self, bytes: &'a [u8]) -> Cow<'a, str> {
        self.0.decode_lossy(bytes)
    }

    /// Decodes with replacement and reports whether recovery was required.
    pub fn recover<'a>(self, bytes: &'a [u8]) -> (Cow<'a, str>, bool) {
        self.0.recover(bytes)
    }

    /// Strictly encodes text, rejecting unrepresentable characters.
    pub fn encode<'a>(self, text: &'a str) -> Result<Cow<'a, [u8]>, Error> {
        self.0.encode(text)
    }
}

impl TryFrom<u32> for Mbcs {
    type Error = Error;

    fn try_from(page: u32) -> Result<Self, Self::Error> {
        Self::require(page)
    }
}

impl From<Mbcs> for Page {
    fn from(page: Mbcs) -> Self {
        page.page()
    }
}

impl From<Mbcs> for u32 {
    fn from(page: Mbcs) -> Self {
        page.id()
    }
}

impl fmt::Debug for Mbcs {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Mbcs")
            .field("id", &self.id())
            .field("name", &self.name())
            .finish()
    }
}

/// `[MS-OSHARED]` ANSI character-set capability.
///
/// This is narrower than [`Mbcs`]: only code pages 874, 932, 936, 949, 950,
/// and 1250 through 1258 can inhabit it. A smart-tag ANSI `PBString` can thus
/// carry this type directly instead of repeatedly validating a raw number.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Ansi(Mbcs);

impl Ansi {
    /// Windows-1252.
    pub const WINDOWS_1252: Self = Self(Mbcs::WINDOWS_1252);

    /// Validates an `[MS-OSHARED]` ANSI code-page identifier.
    pub fn new(page: u32) -> Option<Self> {
        match page {
            874 | 932 | 936 | 949 | 950 | 1250..=1258 => Mbcs::new(page).map(Self),
            _ => None,
        }
    }

    /// Validates an ANSI identifier and retains the unsupported value.
    pub fn require(page: u32) -> Result<Self, Error> {
        Self::new(page).ok_or(Error::Unsupported(page))
    }

    /// General byte-stream capability represented by this ANSI page.
    pub const fn mbcs(self) -> Mbcs {
        self.0
    }

    /// Numeric Windows code-page identifier.
    pub const fn id(self) -> u32 {
        self.0.id()
    }

    /// Exact 16-bit storage form of this supported identifier.
    pub const fn id16(self) -> u16 {
        self.0.id16()
    }

    /// Canonical codec name.
    pub fn name(self) -> &'static str {
        self.0.name()
    }

    /// Strictly decodes ANSI bytes without terminator handling.
    pub fn decode<'a>(self, bytes: &'a [u8]) -> Result<Cow<'a, str>, Error> {
        self.0.decode(bytes)
    }

    /// Strictly encodes text, rejecting unrepresentable characters.
    pub fn encode<'a>(self, text: &'a str) -> Result<Cow<'a, [u8]>, Error> {
        self.0.encode(text)
    }
}

impl TryFrom<u32> for Ansi {
    type Error = Error;

    fn try_from(page: u32) -> Result<Self, Self::Error> {
        Self::require(page)
    }
}

impl From<Ansi> for Mbcs {
    fn from(page: Ansi) -> Self {
        page.mbcs()
    }
}

impl From<Ansi> for Page {
    fn from(page: Ansi) -> Self {
        page.mbcs().page()
    }
}

impl From<Ansi> for u32 {
    fn from(page: Ansi) -> Self {
        page.id()
    }
}

impl fmt::Debug for Ansi {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Ansi")
            .field("id", &self.id())
            .field("name", &self.name())
            .finish()
    }
}

#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(u8)]
enum Kind {
    Ibm866,
    Windows874,
    ShiftJis,
    Gbk,
    EucKr,
    Big5,
    Utf16Le,
    Utf16Be,
    Windows1250,
    Windows1251,
    Windows1252,
    Windows1253,
    Windows1254,
    Windows1255,
    Windows1256,
    Windows1257,
    Windows1258,
    Macintosh,
    Koi8R,
    EucJp,
    Koi8U,
    Iso8859_2,
    Iso8859_3,
    Iso8859_4,
    Iso8859_5,
    Iso8859_6,
    Iso8859_7,
    Iso8859_8,
    Iso8859_13,
    Iso8859_15,
    Gb18030,
    Utf8,
}

#[derive(Clone, Copy)]
enum Utf16Order {
    Little,
    Big,
}

fn encode_utf16(text: &str, order: Utf16Order) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(text.encode_utf16().count().saturating_mul(2));
    for unit in text.encode_utf16() {
        let pair = match order {
            Utf16Order::Little => unit.to_le_bytes(),
            Utf16Order::Big => unit.to_be_bytes(),
        };
        bytes.extend_from_slice(&pair);
    }
    bytes
}

impl Kind {
    const fn utf16_order(self) -> Option<Utf16Order> {
        match self {
            Self::Utf16Le => Some(Utf16Order::Little),
            Self::Utf16Be => Some(Utf16Order::Big),
            _ => None,
        }
    }

    const fn from_id(page: u32) -> Option<Self> {
        match page {
            866 => Some(Self::Ibm866),
            874 => Some(Self::Windows874),
            932 => Some(Self::ShiftJis),
            936 => Some(Self::Gbk),
            949 => Some(Self::EucKr),
            950 => Some(Self::Big5),
            1200 => Some(Self::Utf16Le),
            1201 => Some(Self::Utf16Be),
            1250 => Some(Self::Windows1250),
            1251 => Some(Self::Windows1251),
            1252 => Some(Self::Windows1252),
            1253 => Some(Self::Windows1253),
            1254 => Some(Self::Windows1254),
            1255 => Some(Self::Windows1255),
            1256 => Some(Self::Windows1256),
            1257 => Some(Self::Windows1257),
            1258 => Some(Self::Windows1258),
            10000 => Some(Self::Macintosh),
            20866 => Some(Self::Koi8R),
            20932 => Some(Self::EucJp),
            21866 => Some(Self::Koi8U),
            28592 => Some(Self::Iso8859_2),
            28593 => Some(Self::Iso8859_3),
            28594 => Some(Self::Iso8859_4),
            28595 => Some(Self::Iso8859_5),
            28596 => Some(Self::Iso8859_6),
            28597 => Some(Self::Iso8859_7),
            28598 => Some(Self::Iso8859_8),
            28603 => Some(Self::Iso8859_13),
            28605 => Some(Self::Iso8859_15),
            54936 => Some(Self::Gb18030),
            65001 => Some(Self::Utf8),
            _ => None,
        }
    }

    const fn id(self) -> u16 {
        match self {
            Self::Ibm866 => 866,
            Self::Windows874 => 874,
            Self::ShiftJis => 932,
            Self::Gbk => 936,
            Self::EucKr => 949,
            Self::Big5 => 950,
            Self::Utf16Le => 1200,
            Self::Utf16Be => 1201,
            Self::Windows1250 => 1250,
            Self::Windows1251 => 1251,
            Self::Windows1252 => 1252,
            Self::Windows1253 => 1253,
            Self::Windows1254 => 1254,
            Self::Windows1255 => 1255,
            Self::Windows1256 => 1256,
            Self::Windows1257 => 1257,
            Self::Windows1258 => 1258,
            Self::Macintosh => 10000,
            Self::Koi8R => 20866,
            Self::EucJp => 20932,
            Self::Koi8U => 21866,
            Self::Iso8859_2 => 28592,
            Self::Iso8859_3 => 28593,
            Self::Iso8859_4 => 28594,
            Self::Iso8859_5 => 28595,
            Self::Iso8859_6 => 28596,
            Self::Iso8859_7 => 28597,
            Self::Iso8859_8 => 28598,
            Self::Iso8859_13 => 28603,
            Self::Iso8859_15 => 28605,
            Self::Gb18030 => 54936,
            Self::Utf8 => 65001,
        }
    }

    fn codec(self) -> &'static Encoding {
        match self {
            Self::Ibm866 => encoding_rs::IBM866,
            Self::Windows874 => encoding_rs::WINDOWS_874,
            Self::ShiftJis => encoding_rs::SHIFT_JIS,
            Self::Gbk => encoding_rs::GBK,
            Self::EucKr => encoding_rs::EUC_KR,
            Self::Big5 => encoding_rs::BIG5,
            Self::Utf16Le => encoding_rs::UTF_16LE,
            Self::Utf16Be => encoding_rs::UTF_16BE,
            Self::Windows1250 => encoding_rs::WINDOWS_1250,
            Self::Windows1251 => encoding_rs::WINDOWS_1251,
            Self::Windows1252 => encoding_rs::WINDOWS_1252,
            Self::Windows1253 => encoding_rs::WINDOWS_1253,
            Self::Windows1254 => encoding_rs::WINDOWS_1254,
            Self::Windows1255 => encoding_rs::WINDOWS_1255,
            Self::Windows1256 => encoding_rs::WINDOWS_1256,
            Self::Windows1257 => encoding_rs::WINDOWS_1257,
            Self::Windows1258 => encoding_rs::WINDOWS_1258,
            Self::Macintosh => encoding_rs::MACINTOSH,
            Self::Koi8R => encoding_rs::KOI8_R,
            Self::EucJp => encoding_rs::EUC_JP,
            Self::Koi8U => encoding_rs::KOI8_U,
            Self::Iso8859_2 => encoding_rs::ISO_8859_2,
            Self::Iso8859_3 => encoding_rs::ISO_8859_3,
            Self::Iso8859_4 => encoding_rs::ISO_8859_4,
            Self::Iso8859_5 => encoding_rs::ISO_8859_5,
            Self::Iso8859_6 => encoding_rs::ISO_8859_6,
            Self::Iso8859_7 => encoding_rs::ISO_8859_7,
            Self::Iso8859_8 => encoding_rs::ISO_8859_8,
            Self::Iso8859_13 => encoding_rs::ISO_8859_13,
            Self::Iso8859_15 => encoding_rs::ISO_8859_15,
            Self::Gb18030 => encoding_rs::GB18030,
            Self::Utf8 => encoding_rs::UTF_8,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem;

    #[test]
    fn page_is_compact_and_checked() {
        assert_eq!(mem::size_of::<Page>(), 1);
        assert_eq!(mem::size_of::<Mbcs>(), 1);
        assert_eq!(mem::size_of::<Ansi>(), 1);
        assert_eq!(Page::new(1252), Some(Page::WINDOWS_1252));
        assert_eq!(Page::WINDOWS_1252.name(), "windows-1252");
        assert_eq!(Page::try_from(99999), Err(Error::Unsupported(99999)));
    }

    #[test]
    fn approximate_legacy_substitutions_are_rejected() {
        for page in [437, 850, 1028, 1041, 1042, 65000] {
            assert_eq!(Page::new(page), None, "page {page}");
        }
        assert_eq!(Page::new(866).map(Page::name), Some("IBM866"));
        assert_eq!(Mbcs::new(1200), None);
        assert_eq!(Mbcs::new(1201), None);
        assert_eq!(Ansi::new(1252), Some(Ansi::WINDOWS_1252));
        for page in [866, 1200, 65001, 28592] {
            assert_eq!(Ansi::new(page), None, "ANSI page {page}");
        }
    }

    #[test]
    fn strict_and_lossy_decoding_are_distinct() {
        let malformed_shift_jis = [0x82];
        assert_eq!(
            Page::require(932)
                .and_then(|page| page.decode(&malformed_shift_jis).map(Cow::into_owned)),
            Err(Error::Invalid(932))
        );
        assert_eq!(
            Page::require(932)
                .expect("supported page")
                .decode_lossy(&malformed_shift_jis),
            "�"
        );
    }

    #[test]
    fn strict_encoding_rejects_unrepresentable_text() {
        assert_eq!(
            Page::WINDOWS_1252.encode("界"),
            Err(Error::Unmappable(1252))
        );
        assert_eq!(
            Page::WINDOWS_1252
                .encode("café")
                .expect("representable text")
                .as_ref(),
            b"caf\xe9"
        );
    }

    #[test]
    fn decoding_does_not_guess_record_terminators() {
        let text = Page::WINDOWS_1252
            .decode(b"a\0b")
            .expect("valid code-page bytes");
        assert_eq!(text, "a\0b");
    }

    #[test]
    fn utf16_round_trips_strictly_and_rejects_partial_units() {
        let encoded = Page::UTF_16LE.encode("A界").expect("valid Unicode");
        assert_eq!(encoded.as_ref(), b"A\0\x4c\x75");
        assert_eq!(
            Page::UTF_16LE
                .decode(encoded.as_ref())
                .expect("complete UTF-16 units"),
            "A界"
        );
        assert_eq!(Page::UTF_16LE.decode(b"A"), Err(Error::Invalid(1200)));
    }
}
