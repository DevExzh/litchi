//! Streaming `OfficeArt` BLIP and `BStore` writers.

use std::{
    borrow::Cow,
    io::{self, Write},
};

use super::{Id, JpegFlavor, Kind, Point, Rect, Uid};
use crate::write::{self as record_write, Atom, Container};

/// Physical placement of a BLIP selected by an FBSE.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Placement {
    /// Store the BLIP inside its FBSE.
    Embedded,
    /// Store the BLIP in the host's associated `BStoreDelay` sequence.
    Delay,
}

/// Move-friendly builder for one known `OfficeArt` BLIP.
#[derive(Debug)]
pub struct BlipBuilder<'data> {
    kind: Kind,
    flavor: JpegFlavor,
    data: Cow<'data, [u8]>,
    uid: Uid,
    tag: u8,
    bounds: Rect,
    extent: Point,
    two_uids: bool,
}

impl<'data> BlipBuilder<'data> {
    /// Creates a bitmap BLIP from borrowed or moved encoded bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if `kind` is not a bitmap kind (JPEG, CMYK JPEG, PNG,
    /// DIB, or TIFF).
    pub fn bitmap(kind: Kind, data: impl Into<Cow<'data, [u8]>>) -> io::Result<Self> {
        if !matches!(
            kind,
            Kind::Jpeg | Kind::CmykJpeg | Kind::Png | Kind::Dib | Kind::Tiff
        ) {
            return Err(invalid("bitmap builder requires a bitmap kind"));
        }
        let data_cow = data.into();
        let uid = digest(&data_cow);
        Ok(Self {
            kind,
            flavor: JpegFlavor::Original,
            data: data_cow,
            uid,
            tag: 0xFF,
            bounds: Rect::default(),
            extent: Point::default(),
            two_uids: false,
        })
    }

    /// Creates an uncompressed metafile BLIP from borrowed or moved bytes.
    ///
    /// The writer emits the specification's `0xFE` no-compression marker. A
    /// codec may compress beforehand, but compressed writer input is excluded
    /// here so raw DEFLATE cannot accidentally be labeled as RFC1950 data.
    ///
    /// # Errors
    ///
    /// Returns an error if `kind` is not a metafile kind (EMF, WMF, or PICT).
    pub fn meta(
        kind: Kind,
        data: impl Into<Cow<'data, [u8]>>,
        bounds: Rect,
        extent: Point,
    ) -> io::Result<Self> {
        if !kind.is_meta() {
            return Err(invalid("metafile builder requires EMF, WMF, or PICT"));
        }
        let data_cow = data.into();
        let uid = digest(&data_cow);
        Ok(Self {
            kind,
            flavor: JpegFlavor::Original,
            data: data_cow,
            uid,
            tag: 0xFF,
            bounds,
            extent,
            two_uids: false,
        })
    }

    /// Sets the application-defined bitmap resource tag.
    #[must_use]
    pub const fn tag(mut self, tag: u8) -> Self {
        self.tag = tag;
        self
    }

    /// Selects the exact JPEG record-type flavor.
    ///
    /// # Errors
    ///
    /// Returns an error if this BLIP's kind is not JPEG or CMYK JPEG.
    pub fn jpeg_flavor(mut self, flavor: JpegFlavor) -> io::Result<Self> {
        if !matches!(self.kind, Kind::Jpeg | Kind::CmykJpeg) {
            return Err(invalid("JPEG flavor can only be set on a JPEG BLIP"));
        }
        self.flavor = flavor;
        Ok(self)
    }

    /// Emits the two-UID form. Both fields contain the required MD4 digest.
    #[must_use]
    pub const fn two_uids(mut self) -> Self {
        self.two_uids = true;
        self
    }

    /// Returns the image persistence kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Borrows the uncompressed image file data.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Returns the MD4 UID required by MS-ODRAW.
    #[must_use]
    pub const fn uid(&self) -> Uid {
        self.uid
    }

    /// Returns the complete record size including its eight-byte header.
    ///
    /// # Errors
    ///
    /// Returns an error if the BLIP file data length or total body length
    /// exceeds `u32`.
    pub fn wire_len(&self) -> io::Result<u32> {
        self.body_len()?
            .checked_add(8)
            .ok_or_else(|| invalid("BLIP record length exceeds u32"))
    }

    fn body_len(&self) -> io::Result<u32> {
        let uid_bytes = if self.two_uids { 32u32 } else { 16 };
        let framing = if self.kind.is_meta() { 34 } else { 1 };
        let data = u32::try_from(self.data.len())
            .map_err(|_err| invalid("BLIP file data length exceeds u32"))?;
        uid_bytes
            .checked_add(framing)
            .and_then(|value| value.checked_add(data))
            .ok_or_else(|| invalid("BLIP body length exceeds u32"))
    }

    fn atom(&self) -> io::Result<Atom> {
        let atom = match self.kind {
            Kind::Emf => Atom::BlipEmf,
            Kind::Wmf => Atom::BlipWmf,
            Kind::Pict => Atom::BlipPict,
            Kind::Jpeg | Kind::CmykJpeg => match self.flavor {
                JpegFlavor::Original => Atom::BlipJpeg,
                JpegFlavor::Alternate => Atom::BlipJpeg2,
            },
            Kind::Png => Atom::BlipPng,
            Kind::Dib => Atom::BlipDib,
            Kind::Tiff => Atom::BlipTiff,
            Kind::Error | Kind::Unknown | Kind::Other(_) => {
                return Err(invalid("unsupported BLIP kind"));
            },
        };
        Ok(atom)
    }

    fn instance(&self) -> io::Result<u16> {
        let second = u16::from(self.two_uids);
        let instance = match self.kind {
            Kind::Emf => 0x3D4 + second,
            Kind::Wmf => 0x216 + second,
            Kind::Pict => 0x542 + second,
            Kind::Jpeg => 0x46A + second,
            Kind::CmykJpeg => 0x6E2 + second,
            Kind::Png => 0x6E0 + second,
            Kind::Dib => 0x7A8 + second,
            Kind::Tiff => 0x6E4 + second,
            Kind::Error | Kind::Unknown | Kind::Other(_) => {
                return Err(invalid("unsupported BLIP kind"));
            },
        };
        Ok(instance)
    }

    /// Streams this BLIP record without copying its file data.
    ///
    /// # Errors
    ///
    /// Returns an error if the BLIP kind has no writable record type, if a
    /// length exceeds `u32`, or if the writer fails.
    pub fn write<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        record_write::atom_header(writer, self.instance()?, self.atom()?, self.body_len()?)?;
        let uid = self.uid();
        writer.write_all(uid.as_bytes())?;
        if self.two_uids {
            writer.write_all(uid.as_bytes())?;
        }
        if self.kind.is_meta() {
            let size = u32::try_from(self.data.len())
                .map_err(|_err| invalid("metafile data length exceeds u32"))?;
            writer.write_all(&size.to_le_bytes())?;
            writer.write_all(&self.bounds.left.to_le_bytes())?;
            writer.write_all(&self.bounds.top.to_le_bytes())?;
            writer.write_all(&self.bounds.right.to_le_bytes())?;
            writer.write_all(&self.bounds.bottom.to_le_bytes())?;
            writer.write_all(&self.extent.x.to_le_bytes())?;
            writer.write_all(&self.extent.y.to_le_bytes())?;
            writer.write_all(&size.to_le_bytes())?;
            writer.write_all(&[0xFE, 0xFE])?;
        } else {
            writer.write_all(&[self.tag])?;
        }
        writer.write_all(&self.data)
    }
}

/// Builder for one FBSE and its physical BLIP placement.
#[derive(Debug)]
pub struct EntryBuilder<'data> {
    blip: BlipBuilder<'data>,
    placement: Placement,
    win: Kind,
    mac: Kind,
    tag: u16,
    refs: u32,
    name: Option<Cow<'data, str>>,
}

impl<'data> EntryBuilder<'data> {
    /// Creates an embedded FBSE.
    #[must_use]
    pub fn embedded(blip: BlipBuilder<'data>) -> Self {
        Self::new(blip, Placement::Embedded)
    }

    /// Creates a delay-loaded FBSE.
    #[must_use]
    pub fn delayed(blip: BlipBuilder<'data>) -> Self {
        Self::new(blip, Placement::Delay)
    }

    fn new(blip: BlipBuilder<'data>, placement: Placement) -> Self {
        let kind = blip.kind();
        Self {
            blip,
            placement,
            win: kind,
            mac: kind,
            tag: 0x00FF,
            refs: 1,
            name: None,
        }
    }

    /// Sets both platform persistence fields.
    #[must_use]
    pub const fn platforms(mut self, win: Kind, mac: Kind) -> Self {
        self.win = win;
        self.mac = mac;
        self
    }

    /// Sets the internal resource tag.
    #[must_use]
    pub const fn tag(mut self, tag: u16) -> Self {
        self.tag = tag;
        self
    }

    /// Sets the reference count.
    #[must_use]
    pub const fn refs(mut self, refs: u32) -> Self {
        self.refs = refs;
        self
    }

    /// Sets an optional Unicode name from borrowed or moved text.
    #[must_use]
    pub fn name(mut self, name: impl Into<Cow<'data, str>>) -> Self {
        self.name = Some(name.into());
        self
    }

    fn validate(&self) -> io::Result<()> {
        let selected = self.blip.kind().raw();
        if selected != self.win.raw() && selected != self.mac.raw() {
            return Err(invalid("FBSE instance matches neither platform kind"));
        }
        if self.refs == 0 {
            return Err(invalid("nonempty FBSE requires a nonzero reference count"));
        }
        self.name_len()?;
        Ok(())
    }

    fn name_len(&self) -> io::Result<u8> {
        let Some(name) = &self.name else {
            return Ok(0);
        };
        let units = name.encode_utf16().count();
        let bytes = units
            .checked_add(1)
            .and_then(|value| value.checked_mul(2))
            .ok_or_else(|| invalid("FBSE name length overflows"))?;
        if bytes > 0xFE {
            return Err(invalid("FBSE name exceeds 254 encoded bytes"));
        }
        u8::try_from(bytes).map_err(|_err| invalid("FBSE name length exceeds u8"))
    }

    fn body_len(&self) -> io::Result<u32> {
        self.validate()?;
        let embedded = match self.placement {
            Placement::Embedded => self.blip.wire_len()?,
            Placement::Delay => 0,
        };
        36u32
            .checked_add(u32::from(self.name_len()?))
            .and_then(|value| value.checked_add(embedded))
            .ok_or_else(|| invalid("FBSE body length exceeds u32"))
    }

    fn wire_len(&self) -> io::Result<u32> {
        self.body_len()?
            .checked_add(8)
            .ok_or_else(|| invalid("FBSE record length exceeds u32"))
    }

    fn write<W: Write>(&self, writer: &mut W, delay_offset: u32) -> io::Result<()> {
        let body_len = self.body_len()?;
        record_write::atom_header(
            writer,
            u16::from(self.blip.kind().raw()),
            Atom::Bse,
            body_len,
        )?;
        writer.write_all(&[self.win.raw(), self.mac.raw()])?;
        writer.write_all(self.blip.uid().as_bytes())?;
        writer.write_all(&self.tag.to_le_bytes())?;
        writer.write_all(&self.blip.wire_len()?.to_le_bytes())?;
        writer.write_all(&self.refs.to_le_bytes())?;
        let offset = match self.placement {
            Placement::Embedded => 0,
            Placement::Delay => delay_offset,
        };
        writer.write_all(&offset.to_le_bytes())?;
        writer.write_all(&[0, self.name_len()?, 0, 0])?;
        if let Some(name) = &self.name {
            for unit in name.encode_utf16() {
                writer.write_all(&unit.to_le_bytes())?;
            }
            writer.write_all(&0u16.to_le_bytes())?;
        }
        if self.placement == Placement::Embedded {
            self.blip.write(writer)?;
        }
        Ok(())
    }
}

/// Move-first builder for a `BStore` container and its associated delay store.
#[derive(Debug, Default)]
pub struct StoreBuilder<'data> {
    entries: Vec<EntryBuilder<'data>>,
}

impl<'data> StoreBuilder<'data> {
    /// Creates an empty store builder.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Moves one configured FBSE into the store and returns its checked ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the entry fails validation, if its name exceeds the
    /// encodable length, or if the entry count overflows `usize` or `u32`.
    pub fn push(&mut self, entry: EntryBuilder<'data>) -> io::Result<Id> {
        entry.validate()?;
        let next = self
            .entries
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("BStore entry count overflows usize"))?;
        let value =
            u32::try_from(next).map_err(|_err| invalid("BStore entry count exceeds u32"))?;
        let id = Id::new(value).map_err(|error| invalid(error.to_string()))?;
        self.entries.push(entry);
        Ok(id)
    }

    /// Adds an embedded BLIP.
    ///
    /// # Errors
    ///
    /// Returns an error if the entry fails validation or the entry count
    /// overflows; see [`Self::push`].
    pub fn add_embedded(&mut self, blip: BlipBuilder<'data>) -> io::Result<Id> {
        self.push(EntryBuilder::embedded(blip))
    }

    /// Adds a delay-loaded BLIP.
    ///
    /// # Errors
    ///
    /// Returns an error if the entry fails validation or the entry count
    /// overflows; see [`Self::push`].
    pub fn add_delayed(&mut self, blip: BlipBuilder<'data>) -> io::Result<Id> {
        self.push(EntryBuilder::delayed(blip))
    }

    /// Returns the number of FBSEs.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Returns whether no FBSEs have been added.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    fn store_body_len(&self) -> io::Result<u32> {
        self.entries.iter().try_fold(0u32, |total, entry| {
            total
                .checked_add(entry.wire_len()?)
                .ok_or_else(|| invalid("BStore body length exceeds u32"))
        })
    }

    fn delay_body_len(&self) -> io::Result<u32> {
        self.entries.iter().try_fold(0u32, |total, entry| {
            if entry.placement == Placement::Delay {
                total
                    .checked_add(entry.blip.wire_len()?)
                    .ok_or_else(|| invalid("BStoreDelay length exceeds u32"))
            } else {
                Ok(total)
            }
        })
    }

    /// Streams the `BStore` container. Delayed payloads are not copied here.
    ///
    /// # Errors
    ///
    /// Returns an error if the entry count exceeds 4095, if a length or offset
    /// computation exceeds `u32`, or if the writer fails.
    pub fn write_store<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        let count = u16::try_from(self.entries.len())
            .map_err(|_err| invalid("BStore entry count exceeds u16"))?;
        if count > 0x0FFF {
            return Err(invalid("BStore entry count exceeds 4095"));
        }
        let body_len = self.store_body_len()?;
        self.delay_body_len()?;
        record_write::container_header(writer, count, Container::BStore, body_len)?;
        let mut delay_offset = 0u32;
        for entry in &self.entries {
            entry.write(writer, delay_offset)?;
            if entry.placement == Placement::Delay {
                delay_offset = delay_offset
                    .checked_add(entry.blip.wire_len()?)
                    .ok_or_else(|| invalid("BStoreDelay offset exceeds u32"))?;
            }
        }
        Ok(())
    }

    /// Streams the headerless `BStoreDelay` sequence.
    ///
    /// # Errors
    ///
    /// Returns an error if a length computation exceeds `u32` or if the writer
    /// fails.
    pub fn write_delay<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        self.delay_body_len()?;
        for entry in &self.entries {
            if entry.placement == Placement::Delay {
                entry.blip.write(writer)?;
            }
        }
        Ok(())
    }
}

/// Copies a parsed BLIP while retaining its exact record type and instance.
///
/// # Errors
///
/// Returns an error if the record kind is not a known BLIP atom, or if the
/// writer fails.
pub fn copy<W: Write>(writer: &mut W, blip: &super::Blip<'_>) -> io::Result<()> {
    let record = blip.record();
    let atom = match record.raw_kind() {
        0xF01A => Atom::BlipEmf,
        0xF01B => Atom::BlipWmf,
        0xF01C => Atom::BlipPict,
        0xF01D => Atom::BlipJpeg,
        0xF01E => Atom::BlipPng,
        0xF01F => Atom::BlipDib,
        0xF029 => Atom::BlipTiff,
        0xF02A => Atom::BlipJpeg2,
        raw => Atom::unknown(raw, record.version())?,
    };
    record_write::atom_header(writer, record.instance(), atom, record.len())?;
    writer.write_all(record.data())
}

/// Computes the RFC1320 MD4 digest used for `OfficeArt` image UIDs.
#[must_use]
pub fn digest(data: &[u8]) -> Uid {
    let mut state = [0x6745_2301, 0xEFCD_AB89, 0x98BA_DCFE, 0x1032_5476];
    let mut chunks = data.chunks_exact(64);
    for chunk in &mut chunks {
        let mut block = [0; 64];
        block.copy_from_slice(chunk);
        compress(&mut state, &block);
    }

    let remainder = chunks.remainder();
    let mut tail = [0; 128];
    tail[..remainder.len()].copy_from_slice(remainder);
    tail[remainder.len()] = 0x80;
    let padded = if remainder.len() < 56 { 64 } else { 128 };
    let bit_len = (data.len() as u64).wrapping_mul(8);
    tail[padded - 8..padded].copy_from_slice(&bit_len.to_le_bytes());
    for chunk in tail[..padded].chunks_exact(64) {
        let mut block = [0; 64];
        block.copy_from_slice(chunk);
        compress(&mut state, &block);
    }

    let mut digest = [0; 16];
    for (bytes, word) in digest.chunks_exact_mut(4).zip(state) {
        bytes.copy_from_slice(&word.to_le_bytes());
    }
    Uid::new(digest)
}

fn compress(state: &mut [u32; 4], block: &[u8; 64]) {
    let mut words = [0u32; 16];
    for (word, bytes) in words.iter_mut().zip(block.chunks_exact(4)) {
        *word = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    }
    let [mut aa, mut bb, mut cc, mut dd] = *state;
    macro_rules! round {
        ($f:expr, $k:expr, [$($index:expr),*], [$($shift:expr),*]) => {{
            let indices = [$($index),*];
            let shifts = [$($shift),*];
            for (step, (&index, &shift)) in indices.iter().zip(&shifts).enumerate() {
                let value = match step & 3 {
                    0 => aa.wrapping_add($f(bb, cc, dd)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                    1 => dd.wrapping_add($f(aa, bb, cc)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                    2 => cc.wrapping_add($f(dd, aa, bb)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                    _ => bb.wrapping_add($f(cc, dd, aa)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                };
                match step & 3 { 0 => aa = value, 1 => dd = value, 2 => cc = value, _ => bb = value }
            }
        }};
    }
    round!(
        |left: u32, mid: u32, right: u32| (left & mid) | (!left & right),
        0,
        [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
        [3, 7, 11, 19, 3, 7, 11, 19, 3, 7, 11, 19, 3, 7, 11, 19]
    );
    round!(
        |left: u32, mid: u32, right: u32| (left & mid) | (left & right) | (mid & right),
        0x5A82_7999,
        [0, 4, 8, 12, 1, 5, 9, 13, 2, 6, 10, 14, 3, 7, 11, 15],
        [3, 5, 9, 13, 3, 5, 9, 13, 3, 5, 9, 13, 3, 5, 9, 13]
    );
    round!(
        |left: u32, mid: u32, right: u32| left ^ mid ^ right,
        0x6ED9_EBA1,
        [0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15],
        [3, 9, 11, 15, 3, 9, 11, 15, 3, 9, 11, 15, 3, 9, 11, 15]
    );
    state[0] = state[0].wrapping_add(aa);
    state[1] = state[1].wrapping_add(bb);
    state[2] = state[2].wrapping_add(cc);
    state[3] = state[3].wrapping_add(dd);
}

fn invalid(error: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, error.into())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;
    use crate::image::{Blip, Compression, Context, Delay, Store};

    #[test]
    fn md4_matches_rfc_vectors() {
        assert_eq!(
            digest(b"").bytes(),
            [
                0x31, 0xD6, 0xCF, 0xE0, 0xD1, 0x6A, 0xE9, 0x31, 0xB7, 0x3C, 0x59, 0xD7, 0xE0, 0xC0,
                0x89, 0xC0,
            ]
        );
        assert_eq!(
            digest(b"abc").bytes(),
            [
                0xA4, 0x48, 0x01, 0x7A, 0xAF, 0x21, 0xD8, 0x52, 0x5F, 0xC1, 0x0A, 0xE8, 0x7A, 0xA6,
                0x72, 0x9D,
            ]
        );
    }

    #[test]
    fn metafile_writer_marks_raw_data_uncompressed() {
        let blip = BlipBuilder::meta(
            Kind::Wmf,
            b"wmf".as_slice(),
            Rect::default(),
            Point { x: 10, y: 20 },
        )
        .unwrap();
        let mut bytes = Vec::new();
        blip.write(&mut bytes).unwrap();
        let Blip::Wmf(parsed) = Blip::parse(&bytes).unwrap() else {
            panic!("expected WMF")
        };
        assert_eq!(parsed.header().compression, Compression::None);
        assert_eq!(parsed.header().size, 3);
        assert_eq!(parsed.uids().effective(), digest(b"wmf"));
    }

    #[test]
    fn writes_two_uid_alternate_jpeg_losslessly() {
        let blip = BlipBuilder::bitmap(Kind::Jpeg, b"jpeg".as_slice())
            .unwrap()
            .jpeg_flavor(JpegFlavor::Alternate)
            .unwrap()
            .two_uids();
        let mut bytes = Vec::new();
        blip.write(&mut bytes).unwrap();
        let Blip::Jpeg(parsed) = Blip::parse(&bytes).unwrap() else {
            panic!("expected JPEG")
        };
        assert_eq!(parsed.jpeg_flavor(), Some(JpegFlavor::Alternate));
        assert_eq!(parsed.uids().second(), Some(digest(b"jpeg")));
    }

    #[test]
    fn store_writer_resolves_embedded_and_delayed_payloads() {
        let mut builder = StoreBuilder::new();
        let embedded = builder
            .add_embedded(BlipBuilder::bitmap(Kind::Png, b"one".as_slice()).unwrap())
            .unwrap();
        let delayed = builder
            .add_delayed(BlipBuilder::bitmap(Kind::Png, Vec::from(&b"two"[..])).unwrap())
            .unwrap();
        let mut store_bytes = Vec::new();
        let mut delay_bytes = Vec::new();
        builder.write_store(&mut store_bytes).unwrap();
        builder.write_delay(&mut delay_bytes).unwrap();

        let store = Store::parse(&store_bytes).unwrap();
        assert!(matches!(
            store.resolve(embedded, Context::new()).unwrap(),
            Some(Blip::Png(_))
        ));
        assert!(matches!(
            store
                .resolve(delayed, Context::new().with_delay(Delay::new(&delay_bytes)))
                .unwrap(),
            Some(Blip::Png(_))
        ));
    }
}
