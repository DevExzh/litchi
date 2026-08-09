//! Borrowed, typed `OfficeArt` image objects.
//!
//! The model layer contains no traversal policy and no wire parsing entry
//! points.  It owns the zero-copy objects exposed by the image facade; the
//! codec and validation layers fill and check these objects.

use core::num::NonZeroU16;

use crate::{Error, Record, Result};

pub(super) const FBSE_FIXED_LEN: usize = 36;
pub(super) const MAX_STORE_ENTRIES: u16 = 0x0FFF;

/// Resource ceilings for `OfficeArt` image parsing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Largest accepted BLIP body.
    pub max_blip_bytes: u32,
    /// Largest accepted BLIP-store entry count.
    pub max_store_entries: u16,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_blip_bytes: 256 * 1024 * 1024,
            max_store_entries: MAX_STORE_ENTRIES,
        }
    }
}

/// Persistence format used by an `OfficeArt` BLIP or FBSE platform field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Error reading an image.
    Error,
    /// Image type is unknown.
    Unknown,
    /// Enhanced Metafile.
    Emf,
    /// Windows Metafile.
    Wmf,
    /// Macintosh PICT.
    Pict,
    /// RGB JPEG.
    Jpeg,
    /// PNG.
    Png,
    /// Device-independent bitmap.
    Dib,
    /// TIFF.
    Tiff,
    /// YCCK or CMYK JPEG.
    CmykJpeg,
    /// Producer extension retaining its exact value.
    Other(u8),
}

impl Kind {
    /// Decodes an `MSOBLIPTYPE` wire value without losing extensions.
    #[must_use]
    pub const fn from_raw(raw: u8) -> Self {
        match raw {
            0x00 => Self::Error,
            0x01 => Self::Unknown,
            0x02 => Self::Emf,
            0x03 => Self::Wmf,
            0x04 => Self::Pict,
            0x05 => Self::Jpeg,
            0x06 => Self::Png,
            0x07 => Self::Dib,
            0x11 => Self::Tiff,
            0x12 => Self::CmykJpeg,
            value => Self::Other(value),
        }
    }

    /// Returns the exact `MSOBLIPTYPE` value.
    #[must_use]
    pub const fn raw(self) -> u8 {
        match self {
            Self::Error => 0x00,
            Self::Unknown => 0x01,
            Self::Emf => 0x02,
            Self::Wmf => 0x03,
            Self::Pict => 0x04,
            Self::Jpeg => 0x05,
            Self::Png => 0x06,
            Self::Dib => 0x07,
            Self::Tiff => 0x11,
            Self::CmykJpeg => 0x12,
            Self::Other(value) => value,
        }
    }

    /// Returns whether this format is an `OfficeArt` metafile.
    #[must_use]
    pub const fn is_meta(self) -> bool {
        matches!(self, Self::Emf | Self::Wmf | Self::Pict)
    }

    /// Returns the conventional file extension.
    #[must_use]
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Emf => "emf",
            Self::Wmf => "wmf",
            Self::Pict => "pict",
            Self::Jpeg | Self::CmykJpeg => "jpg",
            Self::Png => "png",
            Self::Dib => "dib",
            Self::Tiff => "tiff",
            Self::Error | Self::Unknown | Self::Other(_) => "bin",
        }
    }
}

/// Exact JPEG record-type flavor retained for lossless writing.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum JpegFlavor {
    /// Original `0xF01D` record type.
    Original,
    /// Later `0xF02A` record type.
    Alternate,
}

/// A 16-byte `OfficeArt` image identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[repr(transparent)]
pub struct Uid([u8; 16]);

impl Uid {
    /// Creates an identifier from its wire bytes.
    #[must_use]
    pub const fn new(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Returns the wire bytes.
    #[must_use]
    pub const fn bytes(self) -> [u8; 16] {
        self.0
    }

    /// Borrows the wire bytes.
    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }

    /// Returns whether every byte is zero.
    #[must_use]
    pub fn is_zero(self) -> bool {
        self.0 == [0; 16]
    }
}

impl From<[u8; 16]> for Uid {
    fn from(bytes: [u8; 16]) -> Self {
        Self::new(bytes)
    }
}

impl From<Uid> for [u8; 16] {
    fn from(uid: Uid) -> Self {
        uid.bytes()
    }
}

/// The one or two UIDs carried by a BLIP record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Uids {
    first: Uid,
    second: Option<Uid>,
}

impl Uids {
    /// Creates a UID pair.
    #[must_use]
    pub const fn new(first: Uid, second: Option<Uid>) -> Self {
        Self { first, second }
    }

    /// Returns the first UID.
    #[must_use]
    pub const fn first(self) -> Uid {
        self.first
    }

    /// Returns the optional second UID.
    #[must_use]
    pub const fn second(self) -> Option<Uid> {
        self.second
    }

    /// Returns the effective UID defined by MS-ODRAW.
    #[must_use]
    pub fn effective(self) -> Uid {
        self.second
            .filter(|uid| !uid.is_zero())
            .unwrap_or(self.first)
    }
}

/// Metafile payload compression.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Compression {
    /// RFC1950-wrapped DEFLATE (`0x00`).
    Deflate,
    /// Uncompressed bytes (`0xFE`).
    None,
}

/// Signed `OfficeArt` rectangle.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Rect {
    /// Left coordinate.
    pub left: i32,
    /// Top coordinate.
    pub top: i32,
    /// Right coordinate.
    pub right: i32,
    /// Bottom coordinate.
    pub bottom: i32,
}

/// Signed `OfficeArt` point or extent.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Point {
    /// Horizontal component.
    pub x: i32,
    /// Vertical component.
    pub y: i32,
}

/// The 34-byte metadata header carried by a metafile BLIP.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MetaHeader {
    /// Uncompressed metafile size (`cbSize`).
    pub size: u32,
    /// Logical clipping bounds.
    pub bounds: Rect,
    /// Rendering size in EMUs.
    pub extent: Point,
    /// Stored payload size (`cbSave`).
    pub saved: u32,
    /// Payload compression.
    pub compression: Compression,
}

/// A borrowed metafile BLIP view.
#[derive(Debug, Clone)]
pub struct Meta<'data> {
    pub(super) record: Record<'data>,
    pub(super) kind: Kind,
    pub(super) uids: Uids,
    pub(super) header: MetaHeader,
    pub(super) data: &'data [u8],
}

impl<'data> Meta<'data> {
    /// Returns the persistence kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns both UID fields.
    #[must_use]
    pub const fn uids(&self) -> Uids {
        self.uids
    }

    /// Returns the metafile header.
    #[must_use]
    pub const fn header(&self) -> MetaHeader {
        self.header
    }

    /// Returns the stored, possibly compressed metafile bytes.
    #[must_use]
    pub const fn data(&self) -> &'data [u8] {
        self.data
    }

    /// Returns the underlying `OfficeArt` record.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// A borrowed bitmap BLIP view.
#[derive(Debug, Clone)]
pub struct Bitmap<'data> {
    pub(super) record: Record<'data>,
    pub(super) kind: Kind,
    pub(super) uids: Uids,
    pub(super) tag: u8,
    pub(super) data: &'data [u8],
}

impl<'data> Bitmap<'data> {
    /// Returns the persistence kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns the JPEG record flavor, when applicable.
    #[must_use]
    pub fn jpeg_flavor(&self) -> Option<JpegFlavor> {
        match self.record.raw_kind() {
            0xF01D => Some(JpegFlavor::Original),
            0xF02A => Some(JpegFlavor::Alternate),
            _ => None,
        }
    }

    /// Returns both UID fields.
    #[must_use]
    pub const fn uids(&self) -> Uids {
        self.uids
    }

    /// Returns the application-defined resource tag.
    #[must_use]
    pub const fn tag(&self) -> u8 {
        self.tag
    }

    /// Returns the encoded image bytes.
    #[must_use]
    pub const fn data(&self) -> &'data [u8] {
        self.data
    }

    /// Returns the underlying `OfficeArt` record.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// An unknown BLIP-range record retained without interpreting its body.
#[derive(Debug, Clone)]
pub struct Opaque<'data> {
    pub(super) record: Record<'data>,
}

impl<'data> Opaque<'data> {
    /// Returns the exact record type.
    #[must_use]
    pub const fn raw_kind(&self) -> u16 {
        self.record.raw_kind()
    }

    /// Returns the uninterpreted record body.
    #[must_use]
    pub const fn data(&self) -> &'data [u8] {
        self.record.data()
    }

    /// Returns the underlying `OfficeArt` record.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// A borrowed `OfficeArt` BLIP.
#[derive(Debug, Clone)]
pub enum Blip<'data> {
    /// Enhanced Metafile.
    Emf(Meta<'data>),
    /// Windows Metafile.
    Wmf(Meta<'data>),
    /// Macintosh PICT.
    Pict(Meta<'data>),
    /// RGB or CMYK JPEG.
    Jpeg(Bitmap<'data>),
    /// PNG.
    Png(Bitmap<'data>),
    /// Device-independent bitmap.
    Dib(Bitmap<'data>),
    /// TIFF.
    Tiff(Bitmap<'data>),
    /// An unrecognized record in the BLIP file-block range.
    Opaque(Opaque<'data>),
}

impl<'data> Blip<'data> {
    /// Returns the persistence kind when known.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Emf(meta) | Self::Wmf(meta) | Self::Pict(meta) => meta.kind(),
            Self::Jpeg(bitmap) | Self::Png(bitmap) | Self::Dib(bitmap) | Self::Tiff(bitmap) => {
                bitmap.kind()
            },
            Self::Opaque(_) => Kind::Unknown,
        }
    }

    /// Returns the exact `OfficeArt` record type.
    #[must_use]
    pub const fn raw_kind(&self) -> u16 {
        self.record().raw_kind()
    }

    /// Returns the underlying `OfficeArt` record.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        match self {
            Self::Emf(meta) | Self::Wmf(meta) | Self::Pict(meta) => meta.record(),
            Self::Jpeg(bitmap) | Self::Png(bitmap) | Self::Dib(bitmap) | Self::Tiff(bitmap) => {
                bitmap.record()
            },
            Self::Opaque(opaque) => opaque.record(),
        }
    }

    /// Returns the stored file data without decompression or adaptation.
    #[must_use]
    pub const fn data(&self) -> &'data [u8] {
        match self {
            Self::Emf(meta) | Self::Wmf(meta) | Self::Pict(meta) => meta.data(),
            Self::Jpeg(bitmap) | Self::Png(bitmap) | Self::Dib(bitmap) | Self::Tiff(bitmap) => {
                bitmap.data()
            },
            Self::Opaque(opaque) => opaque.data(),
        }
    }

    /// Returns the one or two declared UIDs for a known BLIP.
    #[must_use]
    pub const fn uids(&self) -> Option<Uids> {
        match self {
            Self::Emf(meta) | Self::Wmf(meta) | Self::Pict(meta) => Some(meta.uids()),
            Self::Jpeg(bitmap) | Self::Png(bitmap) | Self::Dib(bitmap) | Self::Tiff(bitmap) => {
                Some(bitmap.uids())
            },
            Self::Opaque(_) => None,
        }
    }

    /// Borrows metafile metadata when this is EMF, WMF, or PICT.
    #[must_use]
    pub const fn meta(&self) -> Option<&Meta<'data>> {
        match self {
            Self::Emf(meta) | Self::Wmf(meta) | Self::Pict(meta) => Some(meta),
            Self::Jpeg(_) | Self::Png(_) | Self::Dib(_) | Self::Tiff(_) | Self::Opaque(_) => None,
        }
    }

    /// Borrows bitmap metadata when this is JPEG, PNG, DIB, or TIFF.
    #[must_use]
    pub const fn bitmap(&self) -> Option<&Bitmap<'data>> {
        match self {
            Self::Jpeg(bitmap) | Self::Png(bitmap) | Self::Dib(bitmap) | Self::Tiff(bitmap) => {
                Some(bitmap)
            },
            Self::Emf(_) | Self::Wmf(_) | Self::Pict(_) | Self::Opaque(_) => None,
        }
    }
}

/// A checked one-based index into a `BStore` container.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Id(NonZeroU16);

impl Id {
    /// Creates an identifier in the representable `OfficeArt` range.
    ///
    /// # Errors
    ///
    /// Returns `Error::InvalidBlipId` if `value` is zero or exceeds the
    /// representable `OfficeArt` BLIP identifier range.
    pub fn new(value: u32) -> Result<Self> {
        let id = u16::try_from(value).map_err(|_err| Error::InvalidBlipId { value })?;
        if id > MAX_STORE_ENTRIES {
            return Err(Error::InvalidBlipId {
                value: u32::from(id),
            });
        }
        NonZeroU16::new(id)
            .map(Self)
            .ok_or(Error::InvalidBlipId { value: 0 })
    }

    /// Returns the one-based numeric value.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0.get()
    }
}

impl TryFrom<u32> for Id {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

impl From<Id> for u32 {
    fn from(value: Id) -> Self {
        u32::from(value.get())
    }
}

/// A checked byte offset in an `OfficeArt` delay store.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Offset(pub(super) u32);

impl Offset {
    /// Returns the byte offset.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

/// A zero-copy UTF-16LE FBSE name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Name<'data> {
    pub(super) bytes: &'data [u8],
}

impl<'data> Name<'data> {
    /// Returns the encoded bytes, including the terminating NUL.
    #[must_use]
    pub const fn as_bytes(self) -> &'data [u8] {
        self.bytes
    }

    /// Decodes the name, omitting the terminating NUL.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedImage` if the encoded bytes are not valid
    /// UTF-16.
    pub fn to_string(self) -> Result<String> {
        char::decode_utf16(
            self.bytes[..self.bytes.len() - 2]
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
        )
        .collect::<core::result::Result<String, _>>()
        .map_err(|_err| Error::MalformedImage {
            reason: "FBSE name is not valid UTF-16",
        })
    }
}

/// Physical storage selected by one FBSE.
#[derive(Debug, Clone)]
pub enum Storage<'data> {
    /// BLIP record embedded in the FBSE.
    Embedded(Blip<'data>),
    /// BLIP record located at an associated delay-store offset.
    Delay(Offset),
    /// Empty store slot (`cRef == 0`).
    Empty,
}

/// A borrowed `OfficeArt` FBSE.
#[derive(Debug, Clone)]
pub struct Entry<'data> {
    pub(super) record: Record<'data>,
    pub(super) win: Kind,
    pub(super) mac: Kind,
    pub(super) uid: Uid,
    pub(super) tag: u16,
    pub(super) size: u32,
    pub(super) refs: u32,
    pub(super) delay: u32,
    pub(super) unused1: u8,
    pub(super) name: Option<Name<'data>>,
    pub(super) unused2: u8,
    pub(super) unused3: u8,
    pub(super) embedded: &'data [u8],
    pub(super) limits: Limits,
}

impl<'data> Entry<'data> {
    /// Returns the Windows persistence kind.
    #[must_use]
    pub const fn win(&self) -> Kind {
        self.win
    }

    /// Returns the persistence kind selected by the FBSE record instance.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedImage` if the record instance is not a valid
    /// `MSOBLIPTYPE` value.
    pub fn kind(&self) -> Result<Kind> {
        let raw = u8::try_from(self.record.instance()).map_err(|_err| Error::MalformedImage {
            reason: "FBSE instance is not an MSOBLIPTYPE value",
        })?;
        Ok(Kind::from_raw(raw))
    }

    /// Returns the Macintosh persistence kind.
    #[must_use]
    pub const fn mac(&self) -> Kind {
        self.mac
    }

    /// Returns the FBSE UID.
    #[must_use]
    pub const fn uid(&self) -> Uid {
        self.uid
    }

    /// Returns the application-defined resource tag.
    #[must_use]
    pub const fn tag(&self) -> u16 {
        self.tag
    }

    /// Returns the declared BLIP size.
    #[must_use]
    pub const fn size(&self) -> u32 {
        self.size
    }

    /// Returns the reference count.
    #[must_use]
    pub const fn refs(&self) -> u32 {
        self.refs
    }

    /// Returns a delay offset when the entry declares one.
    #[must_use]
    pub const fn delay_offset(&self) -> Option<Offset> {
        if self.delay == u32::MAX {
            None
        } else {
            Some(Offset(self.delay))
        }
    }

    /// Returns the optional borrowed name.
    #[must_use]
    pub const fn name(&self) -> Option<Name<'data>> {
        self.name
    }

    /// Returns the undefined bytes without interpreting producer extensions.
    #[must_use]
    pub const fn unused(&self) -> [u8; 3] {
        [self.unused1, self.unused2, self.unused3]
    }

    /// Returns the underlying FBSE record.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// Lazy iterator over BLIPs embedded in an FBSE.
#[derive(Debug, Clone)]
pub struct Embedded<'data> {
    pub(super) records: crate::Children<'data>,
    pub(super) limits: Limits,
}

/// One file block in a `BStore` container or delay store.
#[derive(Debug, Clone)]
pub enum Block<'data> {
    /// File BLIP Store Entry.
    Entry(Entry<'data>),
    /// Direct BLIP record.
    Blip(Blip<'data>),
}

/// A checked, lazy view of an `OfficeArt` `BStore` container.
#[derive(Debug, Clone)]
pub struct Store<'data> {
    pub(super) record: Record<'data>,
    pub(super) count: u16,
    pub(super) limits: Limits,
}

/// Lazy, count-validating `BStore` iterator.
#[derive(Debug, Clone)]
pub struct Blocks<'data> {
    pub(super) records: crate::Children<'data>,
    pub(super) expected: u16,
    pub(super) seen: u16,
    pub(super) done: bool,
    pub(super) limits: Limits,
}

/// Headerless `OfficeArt` `BStoreDelay` sequence.
#[derive(Debug, Clone, Copy)]
pub struct Delay<'data> {
    pub(super) data: &'data [u8],
    pub(super) limits: Limits,
}

/// Lazy iterator over a headerless delay store.
#[derive(Debug, Clone)]
pub struct DelayBlocks<'data> {
    pub(super) records: crate::Children<'data>,
    pub(super) limits: Limits,
    pub(super) seen: u32,
    pub(super) done: bool,
}

/// Host-provided resources needed to resolve delayed images.
#[derive(Debug, Clone, Copy, Default)]
pub struct Context<'data> {
    pub(super) delay: Option<Delay<'data>>,
}

impl<'data> Context<'data> {
    /// Creates a context without a delay store.
    #[must_use]
    pub const fn new() -> Self {
        Self { delay: None }
    }

    /// Adds the host's associated delay store.
    #[must_use]
    pub const fn with_delay(mut self, delay: Delay<'data>) -> Self {
        self.delay = Some(delay);
        self
    }
}
