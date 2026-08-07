//! `OfficeArt` image wire decoding and lazy record traversal.

use super::model::{
    Bitmap, Blip, Block, Compression, Context, Delay, DelayBlocks, Embedded, Entry, Kind, Limits,
    Meta, MetaHeader, Offset, Point, Rect, Store, Uid, Uids,
};
use crate::{Error, ImageLimit, Record, RecordKind, Result};

#[derive(Debug, Clone, Copy)]
enum Family {
    Meta,
    Bitmap,
}

/// Parses exactly one BLIP record.
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

    /// Parses a previously checked `OfficeArt` record.
    pub fn from_record(record: Record<'data>) -> Result<Self> {
        Self::from_record_with(record, Limits::default())
    }

    /// Parses a previously checked `OfficeArt` record under explicit limits.
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
                return Ok(Self::Opaque(super::model::Opaque { record }));
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
                let actual =
                    u64::try_from(data.len()).map_err(|_err| Error::ArithmeticOverflow {
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
}

impl<'data> TryFrom<Record<'data>> for Blip<'data> {
    type Error = Error;

    fn try_from(record: Record<'data>) -> Result<Self> {
        Self::from_record(record)
    }
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

pub(super) fn uid_at(data: &[u8], offset: usize) -> Result<Uid> {
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
        let instance = u8::try_from(record.instance()).map_err(|_err| Error::MalformedImage {
            reason: "FBSE instance is not an MSOBLIPTYPE value",
        })?;
        let body = record.data();
        if body.len() < super::model::FBSE_FIXED_LEN {
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
        let name_end = super::model::FBSE_FIXED_LEN.checked_add(name_len).ok_or(
            Error::ArithmeticOverflow {
                context: "FBSE name extent",
            },
        )?;
        let name_bytes =
            body.get(super::model::FBSE_FIXED_LEN..name_end)
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
            Some(super::model::Name { bytes: name_bytes })
        };
        let embedded = &body[name_end..];
        let embedded_len =
            u64::try_from(embedded.len()).map_err(|_err| Error::ArithmeticOverflow {
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
        super::validation::validate_embedded(&entry, instance)?;
        Ok(entry)
    }
}

impl<'data> Entry<'data> {
    /// Iterates all embedded BLIPs without allocating.
    #[must_use]
    pub fn embedded(&self) -> Embedded<'data> {
        Embedded {
            records: crate::Children::new(self.embedded),
            limits: self.limits,
        }
    }
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

pub(super) fn block(record: Record<'_>, limits: Limits) -> Result<Block<'_>> {
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

impl<'data> Store<'data> {
    /// Parses exactly one `BStore` container.
    pub fn parse(data: &'data [u8]) -> Result<Self> {
        Self::parse_with(data, Limits::default())
    }

    /// Parses exactly one `BStore` container under explicit limits.
    pub fn parse_with(data: &'data [u8], limits: Limits) -> Result<Self> {
        let (record, consumed) = Record::parse(data, 0)?;
        if consumed != data.len() {
            return Err(Error::TrailingData { offset: consumed });
        }
        Self::from_record_with(record, limits)
    }

    /// Returns the number declared by the container header.
    #[must_use]
    pub const fn len(&self) -> u16 {
        self.count
    }

    /// Returns whether the store is empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.count == 0
    }

    /// Iterates file blocks lazily and validates the exact declared count.
    #[must_use]
    pub fn iter(&self) -> super::model::Blocks<'data> {
        super::model::Blocks {
            records: crate::Children::new(self.record.data()),
            expected: self.count,
            seen: 0,
            done: false,
            limits: self.limits,
        }
    }

    /// Looks up a checked one-based identifier.
    pub fn get(&self, id: super::model::Id) -> Result<Option<Block<'data>>> {
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
    pub fn resolve(
        &self,
        id: super::model::Id,
        context: Context<'data>,
    ) -> Result<Option<Blip<'data>>> {
        match self.get(id)? {
            None => Ok(None),
            Some(Block::Blip(blip)) => Ok(Some(blip)),
            Some(Block::Entry(entry)) => entry.resolve(context),
        }
    }

    /// Returns the underlying `BStore` record.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

impl<'data> Iterator for super::model::Blocks<'data> {
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

impl std::iter::FusedIterator for super::model::Blocks<'_> {}

impl<'data> Delay<'data> {
    /// Borrows a delay-store byte sequence.
    #[must_use]
    pub const fn new(data: &'data [u8]) -> Self {
        Self::with_limits(
            data,
            Limits {
                max_blip_bytes: 256 * 1024 * 1024,
                max_store_entries: super::model::MAX_STORE_ENTRIES,
            },
        )
    }

    /// Borrows a delay store under explicit image limits.
    #[must_use]
    pub const fn with_limits(data: &'data [u8], limits: Limits) -> Self {
        Self { data, limits }
    }

    /// Iterates every file block in order.
    #[must_use]
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
        let start =
            usize::try_from(offset.get()).map_err(|_err| Error::DelayOffsetOutOfBounds {
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
    #[must_use]
    pub const fn as_bytes(self) -> &'data [u8] {
        self.data
    }
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
