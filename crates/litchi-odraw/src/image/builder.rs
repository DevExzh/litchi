//! Authoring builders for OfficeArt image records.
//!
//! The builders stage BLIPs and FBSE entries.  Their streaming output methods
//! use the checked primitives in [`super::write`], while the `write` module
//! itself remains reserved for copying and digesting already admitted data.

use std::{
    borrow::Cow,
    io::{self, Write},
};

use super::write::{self as image_write, digest};
use super::{Id, JpegFlavor, Kind, Limits, Point, Rect, Uid};
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
            return Err(image_write::invalid(
                "bitmap builder requires a bitmap kind",
            ));
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
            return Err(image_write::invalid(
                "metafile builder requires EMF, WMF, or PICT",
            ));
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
            return Err(image_write::invalid(
                "JPEG flavor can only be set on a JPEG BLIP",
            ));
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
            .ok_or_else(|| image_write::invalid("BLIP record length exceeds u32"))
    }

    fn body_len(&self) -> io::Result<u32> {
        let uid_bytes = if self.two_uids { 32u32 } else { 16 };
        let framing = if self.kind.is_meta() { 34 } else { 1 };
        let data = u32::try_from(self.data.len())
            .map_err(|_err| image_write::invalid("BLIP file data length exceeds u32"))?;
        uid_bytes
            .checked_add(framing)
            .and_then(|value| value.checked_add(data))
            .ok_or_else(|| image_write::invalid("BLIP body length exceeds u32"))
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
                return Err(image_write::invalid("unsupported BLIP kind"));
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
                return Err(image_write::invalid("unsupported BLIP kind"));
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
    pub fn write<W: Write>(&self, writer: &mut W, limits: Limits) -> io::Result<()> {
        let body_len = image_write::bounded_blip_len(self.body_len()?, limits)?;
        record_write::atom_header(
            writer,
            self.instance()?,
            self.atom()?,
            body_len,
            limits.records,
        )?;
        let uid = self.uid();
        writer.write_all(uid.as_bytes())?;
        if self.two_uids {
            writer.write_all(uid.as_bytes())?;
        }
        if self.kind.is_meta() {
            let size = u32::try_from(self.data.len())
                .map_err(|_err| image_write::invalid("metafile data length exceeds u32"))?;
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
            return Err(image_write::invalid(
                "FBSE instance matches neither platform kind",
            ));
        }
        if self.refs == 0 {
            return Err(image_write::invalid(
                "nonempty FBSE requires a nonzero reference count",
            ));
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
            .ok_or_else(|| image_write::invalid("FBSE name length overflows"))?;
        if bytes > 0xFE {
            return Err(image_write::invalid("FBSE name exceeds 254 encoded bytes"));
        }
        u8::try_from(bytes).map_err(|_err| image_write::invalid("FBSE name length exceeds u8"))
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
            .ok_or_else(|| image_write::invalid("FBSE body length exceeds u32"))
    }

    fn wire_len(&self) -> io::Result<u32> {
        self.body_len()?
            .checked_add(8)
            .ok_or_else(|| image_write::invalid("FBSE record length exceeds u32"))
    }

    fn write<W: Write>(&self, writer: &mut W, delay_offset: u32, limits: Limits) -> io::Result<()> {
        let body_len = image_write::bounded_record_len(self.body_len()?, limits, "FBSE body")?;
        record_write::atom_header(
            writer,
            u16::from(self.blip.kind().raw()),
            Atom::Bse,
            body_len,
            limits.records,
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
            self.blip.write(writer, limits)?;
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
            .ok_or_else(|| image_write::invalid("BStore entry count overflows usize"))?;
        let value = u32::try_from(next)
            .map_err(|_err| image_write::invalid("BStore entry count exceeds u32"))?;
        let id = Id::new(value)
            .map_err(|_error| image_write::invalid("BStore entry ID exceeds 4095"))?;
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
                .ok_or_else(|| image_write::invalid("BStore body length exceeds u32"))
        })
    }

    fn delay_body_len(&self) -> io::Result<u32> {
        self.entries.iter().try_fold(0u32, |total, entry| {
            if entry.placement == Placement::Delay {
                total
                    .checked_add(entry.blip.wire_len()?)
                    .ok_or_else(|| image_write::invalid("BStoreDelay length exceeds u32"))
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
    pub fn write_store<W: Write>(&self, writer: &mut W, limits: Limits) -> io::Result<()> {
        let count = u16::try_from(self.entries.len())
            .map_err(|_err| image_write::invalid("BStore entry count exceeds u16"))?;
        if count > 0x0FFF {
            return Err(image_write::invalid("BStore entry count exceeds 4095"));
        }
        if count > limits.max_store_entries {
            return Err(image_write::invalid(
                "BStore entry count exceeds the configured limit",
            ));
        }
        let body_len = self.store_body_len()?;
        image_write::bounded_record_len(body_len, limits, "BStore body")?;
        // `BStoreDelay` is a headerless sequence rather than one record. Its
        // aggregate size may exceed the per-record ceiling; each delayed BLIP
        // is checked by `BlipBuilder::write` below.
        self.delay_body_len()?;
        record_write::container_header(writer, count, Container::BStore, body_len, limits.records)?;
        let mut delay_offset = 0u32;
        for entry in &self.entries {
            entry.write(writer, delay_offset, limits)?;
            if entry.placement == Placement::Delay {
                delay_offset = delay_offset
                    .checked_add(entry.blip.wire_len()?)
                    .ok_or_else(|| image_write::invalid("BStoreDelay offset exceeds u32"))?;
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
    pub fn write_delay<W: Write>(&self, writer: &mut W, limits: Limits) -> io::Result<()> {
        let count = u16::try_from(self.entries.len())
            .map_err(|_err| image_write::invalid("BStore entry count exceeds u16"))?;
        if count > 0x0FFF || count > limits.max_store_entries {
            return Err(image_write::invalid(
                "BStore entry count exceeds the configured limit",
            ));
        }
        // The delay stream has no enclosing length field. Check each BLIP's
        // record body while streaming it rather than imposing an aggregate
        // ceiling on an otherwise valid sequence.
        self.delay_body_len()?;
        for entry in &self.entries {
            if entry.placement == Placement::Delay {
                entry.blip.write(writer, limits)?;
            }
        }
        Ok(())
    }
}
