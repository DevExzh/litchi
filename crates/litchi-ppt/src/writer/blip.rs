//! PowerPoint image registration and host framing.
//!
//! The OfficeArt record grammar and MD4 UID implementation live in
//! `litchi-odraw`. PowerPoint contributes only format sniffing and the rule
//! that FBSEs live in the Dgg BStore while BLIPs form the headerless
//! `Pictures` BStoreDelay stream.

use std::io::{self, Write};

use litchi_odraw::image::write::{BlipBuilder, StoreBuilder, digest};
pub use litchi_odraw::image::{Id, Kind, Uid};
use litchi_odraw::image::{Point, Rect};

/// One moved native image registered with the presentation.
#[derive(Debug, Clone)]
pub struct PictureData {
    data: Vec<u8>,
    kind: Kind,
    uid: Uid,
}

impl PictureData {
    /// Detects a supported native format and moves its encoded bytes.
    pub fn new(data: Vec<u8>) -> io::Result<Self> {
        let kind = detect(&data).ok_or_else(|| invalid("unsupported picture format"))?;
        Self::with_kind(data, kind)
    }

    /// Moves encoded bytes with an explicit native OfficeArt kind.
    pub fn with_kind(data: Vec<u8>, kind: Kind) -> io::Result<Self> {
        let blip = make_blip(kind, &data)?;
        blip.wire_len()?;
        Ok(Self {
            uid: digest(&data),
            data,
            kind,
        })
    }

    /// Borrows the unchanged native file bytes.
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Returns the native OfficeArt kind.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns the specification-defined MD4 UID.
    pub const fn uid(&self) -> Uid {
        self.uid
    }
}

/// Move-first collection of PowerPoint pictures.
#[derive(Debug, Default)]
pub struct Pictures {
    pictures: Vec<PictureData>,
}

impl Pictures {
    /// Creates an empty picture collection.
    pub const fn new() -> Self {
        Self {
            pictures: Vec::new(),
        }
    }

    /// Detects and registers one image, returning its checked one-based ID.
    pub fn add(&mut self, data: Vec<u8>) -> io::Result<Id> {
        self.push(PictureData::new(data)?)
    }

    /// Registers one explicitly typed image.
    pub fn add_as(&mut self, data: Vec<u8>, kind: Kind) -> io::Result<Id> {
        self.push(PictureData::with_kind(data, kind)?)
    }

    /// Moves one already validated image into this collection.
    pub fn push(&mut self, picture: PictureData) -> io::Result<Id> {
        let next = self
            .pictures
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("picture count overflows usize"))?;
        let raw = u32::try_from(next).map_err(|_| invalid("picture count exceeds u32"))?;
        let id = Id::new(raw).map_err(|error| invalid(error.to_string()))?;
        self.pictures.push(picture);
        Ok(id)
    }

    /// Finds an existing picture by its MD4 UID.
    pub fn id_by_uid(&self, uid: Uid) -> Option<Id> {
        self.pictures
            .iter()
            .position(|picture| picture.uid == uid)
            .and_then(|index| index.checked_add(1))
            .and_then(|index| u32::try_from(index).ok())
            .and_then(|raw| Id::new(raw).ok())
    }

    /// Returns the number of registered images.
    pub fn len(&self) -> usize {
        self.pictures.len()
    }

    /// Returns whether no images are registered.
    pub fn is_empty(&self) -> bool {
        self.pictures.is_empty()
    }

    /// Streams the Dgg `BStoreContainer` containing delay-loaded FBSEs.
    pub fn write_store<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        self.builder()?.write_store(writer)
    }

    /// Streams the headerless `Pictures` BStoreDelay sequence.
    pub fn write_delay<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        self.builder()?.write_delay(writer)
    }

    /// Builds the Dgg `BStoreContainer` bytes.
    pub fn store(&self) -> io::Result<Vec<u8>> {
        let mut bytes = Vec::new();
        self.write_store(&mut bytes)?;
        Ok(bytes)
    }

    /// Builds the headerless `Pictures` stream bytes.
    pub fn delay(&self) -> io::Result<Vec<u8>> {
        let mut bytes = Vec::new();
        self.write_delay(&mut bytes)?;
        Ok(bytes)
    }

    fn builder(&self) -> io::Result<StoreBuilder<'_>> {
        let mut builder = StoreBuilder::new();
        for picture in &self.pictures {
            builder.add_delayed(make_blip(picture.kind, &picture.data)?)?;
        }
        Ok(builder)
    }
}

fn make_blip<'data>(kind: Kind, data: &'data [u8]) -> io::Result<BlipBuilder<'data>> {
    if kind.is_meta() {
        BlipBuilder::meta(kind, data, Rect::default(), Point::default())
    } else {
        BlipBuilder::bitmap(kind, data).map(|blip| blip.tag(0))
    }
}

fn detect(data: &[u8]) -> Option<Kind> {
    if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
        Some(Kind::Jpeg)
    } else if data.starts_with(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]) {
        Some(Kind::Png)
    } else if data.starts_with(b"BM") {
        Some(Kind::Dib)
    } else if data.starts_with(&[0x49, 0x49, 0x2A, 0x00])
        || data.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
    {
        Some(Kind::Tiff)
    } else if data.len() >= 44 && data.get(40..44) == Some(&[0x20, 0x45, 0x4D, 0x46]) {
        Some(Kind::Emf)
    } else if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A])
        || data.starts_with(&[0x01, 0x00, 0x09, 0x00])
    {
        Some(Kind::Wmf)
    } else {
        None
    }
}

fn invalid(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message.into())
}

/// Escher property IDs used by picture-frame records.
pub mod prop_id {
    /// Picture BLIP reference.
    pub const PIC_BLIP: u16 = 0x4104;
    /// Picture crop from left.
    pub const CROP_LEFT: u16 = 0x0102;
    /// Picture crop from top.
    pub const CROP_TOP: u16 = 0x0103;
    /// Picture crop from right.
    pub const CROP_RIGHT: u16 = 0x0104;
    /// Picture crop from bottom.
    pub const CROP_BOTTOM: u16 = 0x0105;
    /// Picture flags.
    pub const PIC_FLAGS: u16 = 0x017F;
}

/// Checked picture-frame property builder.
pub struct PictureShapeBuilder {
    id: Id,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    crop: [i32; 4],
}

impl PictureShapeBuilder {
    /// Creates a frame referencing a checked image ID.
    pub const fn new(id: Id, x: i32, y: i32, width: i32, height: i32) -> Self {
        Self {
            id,
            x,
            y,
            width,
            height,
            crop: [0; 4],
        }
    }

    /// Creates a frame from a raw host index after checking it.
    pub fn from_index(
        index: u32,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
    ) -> litchi_odraw::Result<Self> {
        Ok(Self::new(Id::new(index)?, x, y, width, height))
    }

    /// Sets left, top, right, and bottom crop values.
    pub const fn crop(mut self, left: i32, top: i32, right: i32, bottom: i32) -> Self {
        self.crop = [left, top, right, bottom];
        self
    }

    /// Builds simple OfficeArt properties.
    pub fn properties(&self) -> Vec<(u16, u32)> {
        let mut props = vec![(prop_id::PIC_BLIP, u32::from(self.id))];
        for (id, value) in [
            (prop_id::CROP_LEFT, self.crop[0]),
            (prop_id::CROP_TOP, self.crop[1]),
            (prop_id::CROP_RIGHT, self.crop[2]),
            (prop_id::CROP_BOTTOM, self.crop[3]),
        ] {
            if value != 0 {
                props.push((id, value as u32));
            }
        }
        props.push((prop_id::PIC_FLAGS, 0x0008_0000));
        props
    }

    /// Returns the frame rectangle.
    pub const fn position(&self) -> (i32, i32, i32, i32) {
        (self.x, self.y, self.width, self.height)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_odraw::image::{Context, Delay, Store};

    #[test]
    fn detects_and_writes_semantic_delay_topology() {
        let mut pictures = Pictures::new();
        let id = pictures
            .add(vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
            .expect("register PNG");
        assert_eq!(id.get(), 1);
        assert_eq!(
            pictures.id_by_uid(digest(pictures.pictures[0].data())),
            Some(id)
        );

        let store = pictures.store().expect("BStore");
        let delay = pictures.delay().expect("Pictures stream");
        let store = Store::parse(&store).expect("parse BStore");
        let blip = store
            .resolve(id, Context::new().with_delay(Delay::new(&delay)))
            .expect("resolve")
            .expect("image");
        assert_eq!(blip.kind(), Kind::Png);
        assert_eq!(blip.data(), pictures.pictures[0].data());
        assert_eq!(
            blip.uids().expect("known UID").effective(),
            pictures.pictures[0].uid()
        );
    }

    #[test]
    fn rejects_unknown_data_and_excess_ids_without_mutation() {
        let mut pictures = Pictures::new();
        assert!(pictures.add(vec![1, 2, 3]).is_err());
        assert!(pictures.is_empty());
        assert!(PictureShapeBuilder::from_index(0, 0, 0, 1, 1).is_err());
    }
}
