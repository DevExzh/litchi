//! Checked, borrowed OfficeArt BLIP stores and image records.
//!
//! This module owns the format grammar shared by DOC, PPT, and XLS. It does
//! not decompress or render image payloads; those operations belong in a
//! codec crate and consume the borrowed views defined here.

use core::num::NonZeroU16;

use crate::{Error, ImageLimit, Record, RecordKind, Result};

pub mod write;

const FBSE_FIXED_LEN: usize = 36;
const MAX_STORE_ENTRIES: u16 = 0x0FFF;

/// Resource ceilings for OfficeArt image parsing.
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

/// Persistence format used by an OfficeArt BLIP or FBSE platform field.
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

    /// Returns whether this format is an OfficeArt metafile.
    pub const fn is_meta(self) -> bool {
        matches!(self, Self::Emf | Self::Wmf | Self::Pict)
    }

    /// Returns the conventional file extension.
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

/// A 16-byte OfficeArt image identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[repr(transparent)]
pub struct Uid([u8; 16]);

impl Uid {
    /// Creates an identifier from its wire bytes.
    pub const fn new(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Returns the wire bytes.
    pub const fn bytes(self) -> [u8; 16] {
        self.0
    }

    /// Borrows the wire bytes.
    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }

    /// Returns whether every byte is zero.
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
    pub const fn new(first: Uid, second: Option<Uid>) -> Self {
        Self { first, second }
    }

    /// Returns the first UID.
    pub const fn first(self) -> Uid {
        self.first
    }

    /// Returns the optional second UID.
    pub const fn second(self) -> Option<Uid> {
        self.second
    }

    /// Returns the effective UID defined by MS-ODRAW.
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

/// Signed OfficeArt rectangle.
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

/// Signed OfficeArt point or extent.
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
    record: Record<'data>,
    kind: Kind,
    uids: Uids,
    header: MetaHeader,
    data: &'data [u8],
}

impl<'data> Meta<'data> {
    /// Returns the persistence kind.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns both UID fields.
    pub const fn uids(&self) -> Uids {
        self.uids
    }

    /// Returns the metafile header.
    pub const fn header(&self) -> MetaHeader {
        self.header
    }

    /// Returns the stored, possibly compressed metafile bytes.
    pub const fn data(&self) -> &'data [u8] {
        self.data
    }

    /// Returns the underlying OfficeArt record.
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// A borrowed bitmap BLIP view.
#[derive(Debug, Clone)]
pub struct Bitmap<'data> {
    record: Record<'data>,
    kind: Kind,
    uids: Uids,
    tag: u8,
    data: &'data [u8],
}

impl<'data> Bitmap<'data> {
    /// Returns the persistence kind.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns the JPEG record flavor, when applicable.
    pub fn jpeg_flavor(&self) -> Option<JpegFlavor> {
        match self.record.raw_kind() {
            0xF01D => Some(JpegFlavor::Original),
            0xF02A => Some(JpegFlavor::Alternate),
            _ => None,
        }
    }

    /// Returns both UID fields.
    pub const fn uids(&self) -> Uids {
        self.uids
    }

    /// Returns the application-defined resource tag.
    pub const fn tag(&self) -> u8 {
        self.tag
    }

    /// Returns the encoded image bytes.
    pub const fn data(&self) -> &'data [u8] {
        self.data
    }

    /// Returns the underlying OfficeArt record.
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// An unknown BLIP-range record retained without interpreting its body.
#[derive(Debug, Clone)]
pub struct Opaque<'data> {
    record: Record<'data>,
}

impl<'data> Opaque<'data> {
    /// Returns the exact record type.
    pub const fn raw_kind(&self) -> u16 {
        self.record.raw_kind()
    }

    /// Returns the uninterpreted record body.
    pub const fn data(&self) -> &'data [u8] {
        self.record.data()
    }

    /// Returns the underlying OfficeArt record.
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// A borrowed OfficeArt BLIP.
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
    /// Parses exactly one complete BLIP record.
    pub fn parse(data: &'data [u8]) -> Result<Self> {
        Self::parse_with(data, Limits::default())
    }

    /// Parses exactly one complete BLIP record under explicit limits.
    pub fn parse_with(data: &'data [u8], limits: Limits) -> Result<Self> {
        let (record, consumed) = Record::parse(data, 0)?;
        if consumed != data.len() {
            return Err(Error::TrailingData { offset: consumed });
        }
        Self::from_record_with(record, limits)
    }

    /// Parses a previously checked OfficeArt record.
    pub fn from_record(record: Record<'data>) -> Result<Self> {
        Self::from_record_with(record, Limits::default())
    }

    /// Parses a previously checked OfficeArt record under explicit limits.
    pub fn from_record_with(record: Record<'data>, limits: Limits) -> Result<Self> {
        if record.len() > limits.max_blip_bytes {
            return Err(Error::ImageLimitExceeded {
                limit: ImageLimit::BlipBytes,
                maximum: u64::from(limits.max_blip_bytes),
            });
        }
        if !record.kind().is_blip() {
            return Err(Error::NotImageRecord {
                raw_kind: record.raw_kind(),
            });
        }

        let layout = layout(record.raw_kind(), record.instance());
        let Some((kind, family, two_uids)) = layout else {
            if matches!(record.kind(), RecordKind::Unknown(_)) {
                return Ok(Self::Opaque(Opaque { record }));
            }
            return Err(Error::MalformedImage {
                reason: "BLIP record has an invalid instance",
            });
        };
        if record.version() != 0 {
            return Err(Error::MalformedImage {
                reason: "BLIP record version is not zero",
            });
        }

        let body = record.data();
        let (uids, mut offset) = parse_uids(body, two_uids)?;
        match family {
            Family::Meta => {
                let end = offset.checked_add(34).ok_or(Error::ArithmeticOverflow {
                    context: "metafile header extent",
                })?;
                let header_bytes = body.get(offset..end).ok_or(Error::MalformedImage {
                    reason: "metafile BLIP header is truncated",
                })?;
                let size = le_u32(header_bytes, 0)?;
                let bounds = Rect {
                    left: le_i32(header_bytes, 4)?,
                    top: le_i32(header_bytes, 8)?,
                    right: le_i32(header_bytes, 12)?,
                    bottom: le_i32(header_bytes, 16)?,
                };
                let extent = Point {
                    x: le_i32(header_bytes, 20)?,
                    y: le_i32(header_bytes, 24)?,
                };
                let saved = le_u32(header_bytes, 28)?;
                let compression = match header_bytes[32] {
                    0x00 => Compression::Deflate,
                    0xFE => Compression::None,
                    _ => {
                        return Err(Error::MalformedImage {
                            reason: "metafile BLIP compression is not 0x00 or 0xFE",
                        });
                    },
                };
                if header_bytes[33] != 0xFE {
                    return Err(Error::MalformedImage {
                        reason: "metafile BLIP filter is not 0xFE",
                    });
                }
                offset = end;
                let data = &body[offset..];
                let actual = u64::try_from(data.len()).map_err(|_| Error::ArithmeticOverflow {
                    context: "metafile BLIP data length",
                })?;
                if u64::from(saved) != actual {
                    return Err(Error::ImageSizeMismatch {
                        field: "cbSave",
                        declared: u64::from(saved),
                        actual: data.len(),
                    });
                }
                if compression == Compression::None && size != saved {
                    return Err(Error::ImageSizeMismatch {
                        field: "cbSize",
                        declared: u64::from(size),
                        actual: data.len(),
                    });
                }
                let meta = Meta {
                    record,
                    kind,
                    uids,
                    header: MetaHeader {
                        size,
                        bounds,
                        extent,
                        saved,
                        compression,
                    },
                    data,
                };
                match kind {
                    Kind::Emf => Ok(Self::Emf(meta)),
                    Kind::Wmf => Ok(Self::Wmf(meta)),
                    Kind::Pict => Ok(Self::Pict(meta)),
                    _ => Err(Error::MalformedImage {
                        reason: "metafile layout resolved to a non-metafile kind",
                    }),
                }
            },
            Family::Bitmap => {
                let tag = *body.get(offset).ok_or(Error::MalformedImage {
                    reason: "bitmap BLIP tag is missing",
                })?;
                offset += 1;
                let bitmap = Bitmap {
                    record,
                    kind,
                    uids,
                    tag,
                    data: &body[offset..],
                };
                match kind {
                    Kind::Jpeg | Kind::CmykJpeg => Ok(Self::Jpeg(bitmap)),
                    Kind::Png => Ok(Self::Png(bitmap)),
                    Kind::Dib => Ok(Self::Dib(bitmap)),
                    Kind::Tiff => Ok(Self::Tiff(bitmap)),
                    _ => Err(Error::MalformedImage {
                        reason: "bitmap layout resolved to a non-bitmap kind",
                    }),
                }
            },
        }
    }

    /// Returns the persistence kind when known.
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Emf(meta) | Self::Wmf(meta) | Self::Pict(meta) => meta.kind(),
            Self::Jpeg(bitmap) | Self::Png(bitmap) | Self::Dib(bitmap) | Self::Tiff(bitmap) => {
                bitmap.kind()
            },
            Self::Opaque(_) => Kind::Unknown,
        }
    }

    /// Returns the exact OfficeArt record type.
    pub const fn raw_kind(&self) -> u16 {
        self.record().raw_kind()
    }

    /// Returns the underlying OfficeArt record.
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
    pub const fn meta(&self) -> Option<&Meta<'data>> {
        match self {
            Self::Emf(meta) | Self::Wmf(meta) | Self::Pict(meta) => Some(meta),
            _ => None,
        }
    }

    /// Borrows bitmap metadata when this is JPEG, PNG, DIB, or TIFF.
    pub const fn bitmap(&self) -> Option<&Bitmap<'data>> {
        match self {
            Self::Jpeg(bitmap) | Self::Png(bitmap) | Self::Dib(bitmap) | Self::Tiff(bitmap) => {
                Some(bitmap)
            },
            _ => None,
        }
    }
}

impl<'data> TryFrom<Record<'data>> for Blip<'data> {
    type Error = Error;

    fn try_from(record: Record<'data>) -> Result<Self> {
        Self::from_record(record)
    }
}

#[derive(Debug, Clone, Copy)]
enum Family {
    Meta,
    Bitmap,
}

fn layout(raw_kind: u16, instance: u16) -> Option<(Kind, Family, bool)> {
    let value = match (raw_kind, instance) {
        (0xF01A, 0x3D4) => (Kind::Emf, Family::Meta, false),
        (0xF01A, 0x3D5) => (Kind::Emf, Family::Meta, true),
        (0xF01B, 0x216) => (Kind::Wmf, Family::Meta, false),
        (0xF01B, 0x217) => (Kind::Wmf, Family::Meta, true),
        (0xF01C, 0x542) => (Kind::Pict, Family::Meta, false),
        (0xF01C, 0x543) => (Kind::Pict, Family::Meta, true),
        (0xF01D | 0xF02A, 0x46A) => (Kind::Jpeg, Family::Bitmap, false),
        (0xF01D | 0xF02A, 0x46B) => (Kind::Jpeg, Family::Bitmap, true),
        (0xF01D | 0xF02A, 0x6E2) => (Kind::CmykJpeg, Family::Bitmap, false),
        (0xF01D | 0xF02A, 0x6E3) => (Kind::CmykJpeg, Family::Bitmap, true),
        (0xF01E, 0x6E0) => (Kind::Png, Family::Bitmap, false),
        (0xF01E, 0x6E1) => (Kind::Png, Family::Bitmap, true),
        (0xF01F, 0x7A8) => (Kind::Dib, Family::Bitmap, false),
        (0xF01F, 0x7A9) => (Kind::Dib, Family::Bitmap, true),
        (0xF029, 0x6E4) => (Kind::Tiff, Family::Bitmap, false),
        (0xF029, 0x6E5) => (Kind::Tiff, Family::Bitmap, true),
        _ => return None,
    };
    Some(value)
}

fn parse_uids(data: &[u8], second: bool) -> Result<(Uids, usize)> {
    let first = uid_at(data, 0)?;
    if second {
        Ok((Uids::new(first, Some(uid_at(data, 16)?)), 32))
    } else {
        Ok((Uids::new(first, None), 16))
    }
}

fn uid_at(data: &[u8], offset: usize) -> Result<Uid> {
    let end = offset.checked_add(16).ok_or(Error::ArithmeticOverflow {
        context: "BLIP UID extent",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::MalformedImage {
        reason: "BLIP UID is truncated",
    })?;
    let mut uid = [0; 16];
    uid.copy_from_slice(bytes);
    Ok(Uid::new(uid))
}

/// A checked one-based index into a BStore container.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Id(NonZeroU16);

impl Id {
    /// Creates an identifier in the representable OfficeArt range.
    pub fn new(value: u32) -> Result<Self> {
        let value = u16::try_from(value).map_err(|_| Error::InvalidBlipId { value })?;
        if value > MAX_STORE_ENTRIES {
            return Err(Error::InvalidBlipId {
                value: u32::from(value),
            });
        }
        NonZeroU16::new(value)
            .map(Self)
            .ok_or(Error::InvalidBlipId { value: 0 })
    }

    /// Returns the one-based numeric value.
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

/// A checked byte offset in an OfficeArt delay store.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub struct Offset(u32);

impl Offset {
    /// Returns the byte offset.
    pub const fn get(self) -> u32 {
        self.0
    }
}

/// A zero-copy UTF-16LE FBSE name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Name<'data> {
    bytes: &'data [u8],
}

impl<'data> Name<'data> {
    /// Returns the encoded bytes, including the terminating NUL.
    pub const fn as_bytes(self) -> &'data [u8] {
        self.bytes
    }

    /// Decodes the name, omitting the terminating NUL.
    pub fn to_string(self) -> Result<String> {
        char::decode_utf16(
            self.bytes[..self.bytes.len() - 2]
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
        )
        .collect::<core::result::Result<String, _>>()
        .map_err(|_| Error::MalformedImage {
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

/// A borrowed OfficeArt FBSE.
#[derive(Debug, Clone)]
pub struct Entry<'data> {
    record: Record<'data>,
    win: Kind,
    mac: Kind,
    uid: Uid,
    tag: u16,
    size: u32,
    refs: u32,
    delay: u32,
    unused1: u8,
    name: Option<Name<'data>>,
    unused2: u8,
    unused3: u8,
    embedded: &'data [u8],
    limits: Limits,
}

impl<'data> Entry<'data> {
    /// Parses an FBSE record.
    pub fn parse(record: Record<'data>) -> Result<Self> {
        Self::parse_with(record, Limits::default())
    }

    /// Parses an FBSE record under explicit image limits.
    pub fn parse_with(record: Record<'data>, limits: Limits) -> Result<Self> {
        if record.kind() != RecordKind::Bse {
            return Err(Error::NotImageRecord {
                raw_kind: record.raw_kind(),
            });
        }
        if record.version() != 2 {
            return Err(Error::MalformedImage {
                reason: "FBSE record version is not two",
            });
        }
        let instance = u8::try_from(record.instance()).map_err(|_| Error::MalformedImage {
            reason: "FBSE instance is not an MSOBLIPTYPE value",
        })?;
        let body = record.data();
        if body.len() < FBSE_FIXED_LEN {
            return Err(Error::MalformedImage {
                reason: "FBSE fixed fields are truncated",
            });
        }
        let win = Kind::from_raw(body[0]);
        let mac = Kind::from_raw(body[1]);
        if instance != win.raw() && instance != mac.raw() {
            return Err(Error::MalformedImage {
                reason: "FBSE instance matches neither platform kind",
            });
        }
        let uid = uid_at(body, 2)?;
        let tag = le_u16(body, 18)?;
        let size = le_u32(body, 20)?;
        let refs = le_u32(body, 24)?;
        let delay = le_u32(body, 28)?;
        let unused1 = body[32];
        let name_len = usize::from(body[33]);
        let unused2 = body[34];
        let unused3 = body[35];
        if name_len % 2 != 0 {
            return Err(Error::MalformedImage {
                reason: "FBSE name length is odd",
            });
        }
        let name_end = FBSE_FIXED_LEN
            .checked_add(name_len)
            .ok_or(Error::ArithmeticOverflow {
                context: "FBSE name extent",
            })?;
        let name_bytes = body
            .get(FBSE_FIXED_LEN..name_end)
            .ok_or(Error::MalformedImage {
                reason: "FBSE name extends past the record",
            })?;
        let name = if name_bytes.is_empty() {
            None
        } else {
            if name_bytes[name_bytes.len() - 2..] != [0, 0] {
                return Err(Error::MalformedImage {
                    reason: "FBSE name is not NUL terminated",
                });
            }
            let decoded = name_bytes[..name_bytes.len() - 2]
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]));
            if char::decode_utf16(decoded).any(|value| value.is_err()) {
                return Err(Error::MalformedImage {
                    reason: "FBSE name is not valid UTF-16",
                });
            }
            Some(Name { bytes: name_bytes })
        };
        let embedded = &body[name_end..];
        let embedded_len =
            u64::try_from(embedded.len()).map_err(|_| Error::ArithmeticOverflow {
                context: "embedded BLIP byte length",
            })?;
        if !embedded.is_empty() && u64::from(size) != embedded_len {
            return Err(Error::ImageSizeMismatch {
                field: "FBSE size",
                declared: u64::from(size),
                actual: embedded.len(),
            });
        }
        if !embedded.is_empty() && refs == 0 {
            return Err(Error::MalformedImage {
                reason: "FBSE with an embedded BLIP has an empty-slot cRef",
            });
        }
        if embedded.is_empty() && delay == u32::MAX && refs != 0 {
            return Err(Error::MalformedImage {
                reason: "FBSE without delay data uses the sentinel with nonzero cRef",
            });
        }

        let entry = Self {
            record,
            win,
            mac,
            uid,
            tag,
            size,
            refs,
            delay,
            unused1,
            name,
            unused2,
            unused3,
            embedded,
            limits,
        };
        if !entry.embedded.is_empty() {
            let mut selected = false;
            for blip in entry.embedded() {
                let blip = blip?;
                if matches_instance(&blip, instance) {
                    entry.validate_selected(&blip, false)?;
                    selected = true;
                }
            }
            if !selected {
                return Err(Error::MalformedImage {
                    reason: "FBSE has no embedded BLIP matching its instance",
                });
            }
        }
        Ok(entry)
    }

    /// Returns the Windows persistence kind.
    pub const fn win(&self) -> Kind {
        self.win
    }

    /// Returns the persistence kind selected by the FBSE record instance.
    pub fn kind(&self) -> Result<Kind> {
        let raw = u8::try_from(self.record.instance()).map_err(|_| Error::MalformedImage {
            reason: "FBSE instance is not an MSOBLIPTYPE value",
        })?;
        Ok(Kind::from_raw(raw))
    }

    /// Returns the Macintosh persistence kind.
    pub const fn mac(&self) -> Kind {
        self.mac
    }

    /// Returns the FBSE UID.
    pub const fn uid(&self) -> Uid {
        self.uid
    }

    /// Returns the application-defined resource tag.
    pub const fn tag(&self) -> u16 {
        self.tag
    }

    /// Returns the declared BLIP size.
    pub const fn size(&self) -> u32 {
        self.size
    }

    /// Returns the reference count.
    pub const fn refs(&self) -> u32 {
        self.refs
    }

    /// Returns a delay offset when the entry declares one.
    pub const fn delay_offset(&self) -> Option<Offset> {
        if self.delay == u32::MAX {
            None
        } else {
            Some(Offset(self.delay))
        }
    }

    /// Returns the optional borrowed name.
    pub const fn name(&self) -> Option<Name<'data>> {
        self.name
    }

    /// Returns the undefined bytes without interpreting producer extensions.
    pub const fn unused(&self) -> [u8; 3] {
        [self.unused1, self.unused2, self.unused3]
    }

    /// Iterates all embedded BLIPs without allocating.
    pub fn embedded(&self) -> Embedded<'data> {
        Embedded {
            records: crate::Children::new(self.embedded),
            limits: self.limits,
        }
    }

    /// Returns the physical storage selected by this entry.
    pub fn storage(&self) -> Result<Storage<'data>> {
        if !self.embedded.is_empty() {
            let instance =
                u8::try_from(self.record.instance()).map_err(|_| Error::MalformedImage {
                    reason: "FBSE instance is not an MSOBLIPTYPE value",
                })?;
            for blip in self.embedded() {
                let blip = blip?;
                if matches_instance(&blip, instance) {
                    return Ok(Storage::Embedded(blip));
                }
            }
            return Err(Error::MalformedImage {
                reason: "FBSE has no embedded BLIP matching its instance",
            });
        }
        if self.refs == 0 {
            return Ok(Storage::Empty);
        }
        if self.delay == u32::MAX {
            return Err(Error::MalformedImage {
                reason: "nonempty FBSE has no embedded or delayed BLIP",
            });
        }
        Ok(Storage::Delay(Offset(self.delay)))
    }

    /// Resolves this entry using a host-supplied context.
    pub fn resolve(&self, context: Context<'data>) -> Result<Option<Blip<'data>>> {
        match self.storage()? {
            Storage::Embedded(blip) => Ok(Some(blip)),
            Storage::Empty => Ok(None),
            Storage::Delay(offset) => {
                let delay = context.delay.ok_or(Error::MissingDelayStore)?;
                match delay.at(offset)? {
                    Block::Blip(blip) => {
                        self.validate_selected(&blip, true)?;
                        Ok(Some(blip))
                    },
                    Block::Entry(_) => Err(Error::MalformedImage {
                        reason: "FBSE delay offset does not point to a BLIP",
                    }),
                }
            },
        }
    }

    /// Returns the underlying FBSE record.
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }

    fn validate_selected(&self, blip: &Blip<'_>, delayed: bool) -> Result<()> {
        let instance = u8::try_from(self.record.instance()).map_err(|_| Error::MalformedImage {
            reason: "FBSE instance is not an MSOBLIPTYPE value",
        })?;
        if !matches_instance(blip, instance) {
            return Err(Error::MalformedImage {
                reason: "resolved BLIP kind does not match the FBSE instance",
            });
        }
        if delayed {
            let actual = blip
                .record()
                .len()
                .checked_add(8)
                .ok_or(Error::ArithmeticOverflow {
                    context: "resolved BLIP wire length",
                })?;
            if actual != self.size {
                let actual = usize::try_from(actual).map_err(|_| Error::ArithmeticOverflow {
                    context: "resolved BLIP wire length",
                })?;
                return Err(Error::ImageSizeMismatch {
                    field: "FBSE size",
                    declared: u64::from(self.size),
                    actual,
                });
            }
        }
        if let Some(uids) = blip.uids()
            && uids.effective() != self.uid
        {
            return Err(Error::MalformedImage {
                reason: "resolved BLIP UID does not match the FBSE UID",
            });
        }
        Ok(())
    }
}

fn matches_instance(blip: &Blip<'_>, instance: u8) -> bool {
    blip.kind().raw() == instance
        || matches!(blip, Blip::Opaque(_))
            && matches!(Kind::from_raw(instance), Kind::Unknown | Kind::Other(_))
}

/// Lazy iterator over BLIPs embedded in an FBSE.
#[derive(Debug, Clone)]
pub struct Embedded<'data> {
    records: crate::Children<'data>,
    limits: Limits,
}

impl<'data> Iterator for Embedded<'data> {
    type Item = Result<Blip<'data>>;

    fn next(&mut self) -> Option<Self::Item> {
        self.records
            .next()
            .map(|record| record.and_then(|record| Blip::from_record_with(record, self.limits)))
    }
}

impl std::iter::FusedIterator for Embedded<'_> {}

/// One file block in a BStore container or delay store.
#[derive(Debug, Clone)]
pub enum Block<'data> {
    /// File BLIP Store Entry.
    Entry(Entry<'data>),
    /// Direct BLIP record.
    Blip(Blip<'data>),
}

fn block<'data>(record: Record<'data>, limits: Limits) -> Result<Block<'data>> {
    if record.kind() == RecordKind::Bse {
        Entry::parse_with(record, limits).map(Block::Entry)
    } else if record.kind().is_blip() {
        Blip::from_record_with(record, limits).map(Block::Blip)
    } else {
        Err(Error::MalformedImage {
            reason: "BLIP store contains a non-image file block",
        })
    }
}

/// A checked, lazy view of an OfficeArt BStore container.
#[derive(Debug, Clone)]
pub struct Store<'data> {
    record: Record<'data>,
    count: u16,
    limits: Limits,
}

impl<'data> Store<'data> {
    /// Parses exactly one BStore container.
    pub fn parse(data: &'data [u8]) -> Result<Self> {
        Self::parse_with(data, Limits::default())
    }

    /// Parses exactly one BStore container under explicit limits.
    pub fn parse_with(data: &'data [u8], limits: Limits) -> Result<Self> {
        let (record, consumed) = Record::parse(data, 0)?;
        if consumed != data.len() {
            return Err(Error::TrailingData { offset: consumed });
        }
        Self::from_record_with(record, limits)
    }

    /// Validates a previously parsed BStore record.
    pub fn from_record(record: Record<'data>) -> Result<Self> {
        Self::from_record_with(record, Limits::default())
    }

    /// Validates a previously parsed BStore record under explicit limits.
    pub fn from_record_with(record: Record<'data>, limits: Limits) -> Result<Self> {
        if record.kind() != RecordKind::BStoreContainer
            || record.version() != 0x0F
            || !record.is_container()
        {
            return Err(Error::MalformedImage {
                reason: "record is not an OfficeArt BStore container",
            });
        }
        let count = record.instance();
        if count > limits.max_store_entries {
            return Err(Error::ImageLimitExceeded {
                limit: ImageLimit::StoreEntries,
                maximum: u64::from(limits.max_store_entries),
            });
        }
        let mut actual = 0u16;
        for child in crate::Children::new(record.data()) {
            child?;
            actual = actual.checked_add(1).ok_or(Error::MalformedImage {
                reason: "BStore file-block count exceeds 4095",
            })?;
        }
        if actual != count {
            return Err(Error::MalformedImage {
                reason: "BStore file-block count does not match recInstance",
            });
        }
        Ok(Self {
            record,
            count,
            limits,
        })
    }

    /// Returns the number declared by the container header.
    pub const fn len(&self) -> u16 {
        self.count
    }

    /// Returns whether the store is empty.
    pub const fn is_empty(&self) -> bool {
        self.count == 0
    }

    /// Iterates file blocks lazily and validates the exact declared count.
    pub fn iter(&self) -> Blocks<'data> {
        Blocks {
            records: crate::Children::new(self.record.data()),
            expected: self.count,
            seen: 0,
            done: false,
            limits: self.limits,
        }
    }

    /// Looks up a checked one-based identifier.
    pub fn get(&self, id: Id) -> Result<Option<Block<'data>>> {
        if id.get() > self.count {
            return Ok(None);
        }
        let wanted = id.get() - 1;
        for (index, item) in self.iter().enumerate() {
            let item = item?;
            if index == usize::from(wanted) {
                return Ok(Some(item));
            }
        }
        Ok(None)
    }

    /// Resolves an image identifier with optional delay-store context.
    pub fn resolve(&self, id: Id, context: Context<'data>) -> Result<Option<Blip<'data>>> {
        match self.get(id)? {
            None => Ok(None),
            Some(Block::Blip(blip)) => Ok(Some(blip)),
            Some(Block::Entry(entry)) => entry.resolve(context),
        }
    }

    /// Returns the underlying BStore record.
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// Lazy, count-validating BStore iterator.
#[derive(Debug, Clone)]
pub struct Blocks<'data> {
    records: crate::Children<'data>,
    expected: u16,
    seen: u16,
    done: bool,
    limits: Limits,
}

impl<'data> Iterator for Blocks<'data> {
    type Item = Result<Block<'data>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.records.next() {
            Some(Ok(_)) if self.seen >= self.expected => {
                self.done = true;
                Some(Err(Error::MalformedImage {
                    reason: "BStore contains more file blocks than recInstance",
                }))
            },
            Some(Ok(record)) => {
                self.seen += 1;
                Some(block(record, self.limits))
            },
            Some(Err(error)) => {
                self.done = true;
                Some(Err(error))
            },
            None if self.seen != self.expected => {
                self.done = true;
                Some(Err(Error::MalformedImage {
                    reason: "BStore contains fewer file blocks than recInstance",
                }))
            },
            None => {
                self.done = true;
                None
            },
        }
    }
}

impl std::iter::FusedIterator for Blocks<'_> {}

/// Headerless OfficeArt BStoreDelay sequence.
#[derive(Debug, Clone, Copy)]
pub struct Delay<'data> {
    data: &'data [u8],
    limits: Limits,
}

impl<'data> Delay<'data> {
    /// Borrows a delay-store byte sequence.
    pub const fn new(data: &'data [u8]) -> Self {
        Self::with_limits(
            data,
            Limits {
                max_blip_bytes: 256 * 1024 * 1024,
                max_store_entries: MAX_STORE_ENTRIES,
            },
        )
    }

    /// Borrows a delay store under explicit image limits.
    pub const fn with_limits(data: &'data [u8], limits: Limits) -> Self {
        Self { data, limits }
    }

    /// Iterates every file block in order.
    pub fn iter(self) -> DelayBlocks<'data> {
        DelayBlocks {
            records: crate::Children::new(self.data),
            limits: self.limits,
            seen: 0,
            done: false,
        }
    }

    /// Parses the file block beginning at an exact delay offset.
    pub fn at(self, offset: Offset) -> Result<Block<'data>> {
        let start = usize::try_from(offset.get()).map_err(|_| Error::DelayOffsetOutOfBounds {
            offset: offset.get(),
            available: self.data.len(),
        })?;
        if start >= self.data.len() {
            return Err(Error::DelayOffsetOutOfBounds {
                offset: offset.get(),
                available: self.data.len(),
            });
        }
        let (record, _) = Record::parse(self.data, start)?;
        block(record, self.limits)
    }

    /// Returns the borrowed delay-store bytes.
    pub const fn as_bytes(self) -> &'data [u8] {
        self.data
    }
}

/// Lazy iterator over a headerless delay store.
#[derive(Debug, Clone)]
pub struct DelayBlocks<'data> {
    records: crate::Children<'data>,
    limits: Limits,
    seen: u32,
    done: bool,
}

impl<'data> Iterator for DelayBlocks<'data> {
    type Item = Result<Block<'data>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.records.next() {
            Some(Ok(_)) if self.seen >= u32::from(self.limits.max_store_entries) => {
                self.done = true;
                Some(Err(Error::ImageLimitExceeded {
                    limit: ImageLimit::StoreEntries,
                    maximum: u64::from(self.limits.max_store_entries),
                }))
            },
            Some(Ok(record)) => {
                self.seen += 1;
                Some(block(record, self.limits))
            },
            Some(Err(error)) => {
                self.done = true;
                Some(Err(error))
            },
            None => {
                self.done = true;
                None
            },
        }
    }
}

impl std::iter::FusedIterator for DelayBlocks<'_> {}

/// Host-provided resources needed to resolve delayed images.
#[derive(Debug, Clone, Copy, Default)]
pub struct Context<'data> {
    delay: Option<Delay<'data>>,
}

impl<'data> Context<'data> {
    /// Creates a context without a delay store.
    pub const fn new() -> Self {
        Self { delay: None }
    }

    /// Adds the host's associated delay store.
    pub const fn with_delay(mut self, delay: Delay<'data>) -> Self {
        self.delay = Some(delay);
        self
    }
}

fn le_u16(data: &[u8], offset: usize) -> Result<u16> {
    let end = offset.checked_add(2).ok_or(Error::ArithmeticOverflow {
        context: "image u16 field extent",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::MalformedImage {
        reason: "image integer field is truncated",
    })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn le_u32(data: &[u8], offset: usize) -> Result<u32> {
    let end = offset.checked_add(4).ok_or(Error::ArithmeticOverflow {
        context: "image u32 field extent",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::MalformedImage {
        reason: "image integer field is truncated",
    })?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn le_i32(data: &[u8], offset: usize) -> Result<i32> {
    let end = offset.checked_add(4).ok_or(Error::ArithmeticOverflow {
        context: "image i32 field extent",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::MalformedImage {
        reason: "image integer field is truncated",
    })?;
    Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(version: u8, instance: u16, kind: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + body.len());
        bytes.extend_from_slice(&(u16::from(version) | (instance << 4)).to_le_bytes());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
        bytes.extend_from_slice(body);
        bytes
    }

    fn png_body(two: bool, data: &[u8]) -> Vec<u8> {
        let mut body = vec![1; 16];
        if two {
            body.extend_from_slice(&[2; 16]);
        }
        body.push(0xFF);
        body.extend_from_slice(data);
        body
    }

    #[test]
    fn parses_borrowed_two_uid_bitmap_and_retains_jpeg_flavor() {
        let bytes = record(0, 0x6E1, 0xF01E, &png_body(true, b"png"));
        let blip = Blip::parse(&bytes).expect("valid PNG BLIP");
        let Blip::Png(bitmap) = blip else {
            panic!("expected PNG")
        };
        assert_eq!(bitmap.uids().second(), Some(Uid::new([2; 16])));
        assert_eq!(bitmap.data(), b"png");
        assert_eq!(bitmap.data().as_ptr(), bytes[8 + 33..].as_ptr());

        let jpeg = record(0, 0x46A, 0xF02A, &png_body(false, b"jpeg"));
        let Blip::Jpeg(bitmap) = Blip::parse(&jpeg).expect("valid alternate JPEG") else {
            panic!("expected JPEG")
        };
        assert_eq!(bitmap.jpeg_flavor(), Some(JpegFlavor::Alternate));
        assert_eq!(bitmap.record().raw_kind(), 0xF02A);
    }

    #[test]
    fn validates_metafile_sizes_and_compression() {
        let mut body = vec![0; 16];
        body.extend_from_slice(&3u32.to_le_bytes());
        body.extend_from_slice(&[0; 24]);
        body.extend_from_slice(&3u32.to_le_bytes());
        body.push(0xFE);
        body.push(0xFE);
        body.extend_from_slice(b"wmf");
        let bytes = record(0, 0x216, 0xF01B, &body);
        let Blip::Wmf(meta) = Blip::parse(&bytes).expect("valid WMF") else {
            panic!("expected WMF")
        };
        assert_eq!(meta.header().compression, Compression::None);
        assert_eq!(meta.data(), b"wmf");

        let mut invalid = bytes;
        invalid[8 + 16 + 28] = 4;
        assert!(matches!(
            Blip::parse(&invalid),
            Err(Error::ImageSizeMismatch {
                field: "cbSave",
                ..
            })
        ));
    }

    #[test]
    fn preserves_unknown_file_block_kinds() {
        let bytes = record(7, 0x123, 0xF020, b"future");
        let Blip::Opaque(opaque) = Blip::parse(&bytes).expect("opaque BLIP") else {
            panic!("expected opaque BLIP")
        };
        assert_eq!(opaque.raw_kind(), 0xF020);
        assert_eq!(opaque.data(), b"future");
    }

    #[test]
    fn lazily_validates_direct_and_fbse_store_blocks() {
        let png = record(0, 0x6E0, 0xF01E, &png_body(false, b"x"));
        let mut fbse = Vec::new();
        fbse.extend_from_slice(&[Kind::Png.raw(), Kind::Pict.raw()]);
        fbse.extend_from_slice(&[1; 16]);
        fbse.extend_from_slice(&0u16.to_le_bytes());
        fbse.extend_from_slice(&(png.len() as u32).to_le_bytes());
        fbse.extend_from_slice(&1u32.to_le_bytes());
        fbse.extend_from_slice(&0u32.to_le_bytes());
        fbse.extend_from_slice(&[0; 4]);
        fbse.extend_from_slice(&png);
        let fbse = record(2, u16::from(Kind::Png.raw()), 0xF007, &fbse);

        let mut body = fbse;
        body.extend_from_slice(&png);
        let store = record(0x0F, 2, 0xF001, &body);
        let store = Store::parse(&store).expect("valid store");
        assert_eq!(store.len(), 2);
        let Some(Block::Entry(entry)) = store.get(Id::new(1).unwrap()).unwrap() else {
            panic!("expected FBSE")
        };
        assert_eq!(entry.win(), Kind::Png);
        assert_eq!(entry.mac(), Kind::Pict);
        assert!(matches!(
            store.get(Id::new(2).unwrap()).unwrap(),
            Some(Block::Blip(Blip::Png(_)))
        ));
    }

    #[test]
    fn resolves_delay_offset_zero_and_rejects_missing_context() {
        let png = record(0, 0x6E0, 0xF01E, &png_body(false, b"x"));
        let mut fbse = Vec::new();
        fbse.extend_from_slice(&[Kind::Png.raw(), Kind::Png.raw()]);
        fbse.extend_from_slice(&[1; 16]);
        fbse.extend_from_slice(&0u16.to_le_bytes());
        fbse.extend_from_slice(&(png.len() as u32).to_le_bytes());
        fbse.extend_from_slice(&1u32.to_le_bytes());
        fbse.extend_from_slice(&0u32.to_le_bytes());
        fbse.extend_from_slice(&[0; 4]);
        let fbse = record(2, u16::from(Kind::Png.raw()), 0xF007, &fbse);
        let store = record(0x0F, 1, 0xF001, &fbse);
        let store = Store::parse(&store).unwrap();
        let id = Id::new(1).unwrap();
        assert_eq!(
            store.resolve(id, Context::new()).unwrap_err(),
            Error::MissingDelayStore
        );
        assert!(matches!(
            store
                .resolve(id, Context::new().with_delay(Delay::new(&png)))
                .unwrap(),
            Some(Blip::Png(_))
        ));
        let mut wrong_uid = png;
        wrong_uid[8] ^= 0xFF;
        assert!(matches!(
            store.resolve(id, Context::new().with_delay(Delay::new(&wrong_uid))),
            Err(Error::MalformedImage { .. })
        ));
    }

    #[test]
    fn detects_store_count_mismatch_on_iteration() {
        let bytes = record(0x0F, 1, 0xF001, &[]);
        assert!(matches!(
            Store::parse(&bytes),
            Err(Error::MalformedImage { .. })
        ));
    }

    #[test]
    fn bounds_headerless_delay_block_count_and_fuses_after_error() {
        let png = record(0, 0x6E0, 0xF01E, &png_body(false, b"x"));
        let mut bytes = png.clone();
        bytes.extend_from_slice(&png);
        let mut blocks = Delay::with_limits(
            &bytes,
            Limits {
                max_store_entries: 1,
                ..Limits::default()
            },
        )
        .iter();

        assert!(blocks.next().is_some_and(|block| block.is_ok()));
        assert!(matches!(
            blocks.next(),
            Some(Err(Error::ImageLimitExceeded {
                limit: ImageLimit::StoreEntries,
                maximum: 1,
            }))
        ));
        assert!(blocks.next().is_none());
    }
}
