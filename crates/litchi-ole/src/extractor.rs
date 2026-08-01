//! Host-neutral image discovery for legacy Office containers.
//!
//! OfficeArt grammar is provided by `litchi-odraw`. This module only maps the
//! DOC/PPT host topology onto borrowed BLIP views. Rasterization and
//! decompression remain behind the optional `imgconv` feature.

use litchi_core::error::{Error, Result};
use litchi_odraw::image::{Blip, Block, Context, Delay, Entry, Id, Kind, Store};
use litchi_odraw::{Children, Container, Parser, Record, RecordKind};

#[cfg(feature = "imgconv")]
use std::borrow::Cow;

const MAX_IMAGES: usize = 0x0FFF;

fn parse_error(error: impl std::fmt::Display) -> Error {
    Error::ParseError(error.to_string())
}

/// Backing storage for an extracted image.
///
/// Borrowed views are zero-copy. Owned records are parsed on demand, avoiding
/// a self-referential allocation.
#[derive(Debug, Clone)]
enum Source<'data> {
    View(Blip<'data>),
    Bytes(Vec<u8>),
}

/// One validated OfficeArt image with host metadata.
#[derive(Debug, Clone)]
pub struct ExtractedImage<'data> {
    source: Source<'data>,
    kind: Kind,
    /// Optional name or filename hint.
    pub name: Option<String>,
    /// Zero-based position in the host image collection.
    pub index: usize,
}

impl<'data> ExtractedImage<'data> {
    /// Wraps a validated borrowed BLIP without copying its file data.
    pub fn new(blip: Blip<'data>, name: Option<String>, index: usize) -> Self {
        let kind = blip.kind();
        Self {
            source: Source::View(blip),
            kind,
            name,
            index,
        }
    }

    /// Parses a fresh borrowed BLIP view.
    pub fn blip(&self) -> Result<Blip<'_>> {
        match &self.source {
            Source::View(blip) => Ok(blip.clone()),
            Source::Bytes(bytes) => Blip::parse(bytes).map_err(parse_error),
        }
    }

    /// Returns the native OfficeArt image kind.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns the conventional native extension.
    pub const fn extension(&self) -> &'static str {
        self.kind.extension()
    }

    /// Returns the recommended extraction extension.
    pub const fn output_extension(&self) -> &'static str {
        match self.kind {
            Kind::Emf | Kind::Wmf => "svg",
            Kind::Pict => "png",
            _ => self.kind.extension(),
        }
    }

    /// Returns a suggested output filename.
    pub fn suggested_filename(&self) -> String {
        match &self.name {
            Some(name) => {
                let stem = name
                    .rsplit_once('.')
                    .map_or(name.as_str(), |(stem, _)| stem);
                format!("{stem}.{}", self.output_extension())
            },
            None => format!("image_{:03}.{}", self.index, self.output_extension()),
        }
    }

    /// Borrows the stored image file data without decoding it.
    pub fn data(&self) -> Result<&[u8]> {
        match &self.source {
            Source::View(blip) => Ok(blip.data()),
            Source::Bytes(bytes) => Blip::parse(bytes)
                .map(|blip| blip.data())
                .map_err(parse_error),
        }
    }

    /// Moves or copies this image into an independently owned record.
    pub fn into_owned(self) -> Result<ExtractedImage<'static>> {
        let bytes = match self.source {
            Source::Bytes(bytes) => bytes,
            Source::View(blip) => {
                let mut bytes = Vec::new();
                litchi_odraw::image::write::copy(&mut bytes, &blip)?;
                bytes
            },
        };
        Ok(ExtractedImage {
            source: Source::Bytes(bytes),
            kind: self.kind,
            name: self.name,
            index: self.index,
        })
    }

    /// Returns bounded, decompressed native file data.
    #[cfg(feature = "imgconv")]
    pub fn decode(&self, limits: litchi_imgconv::Limits) -> Result<Cow<'_, [u8]>> {
        let blip = self.blip()?;
        litchi_imgconv::decode_data(&blip, &limits)
    }

    /// Converts this image to PNG under explicit sizing and resource limits.
    #[cfg(feature = "imgconv")]
    pub fn to_png(&self, options: litchi_imgconv::Options) -> Result<Vec<u8>> {
        litchi_imgconv::to_png(&self.blip()?, options)
    }

    /// Converts this image to JPEG under explicit sizing and resource limits.
    #[cfg(feature = "imgconv")]
    pub fn to_jpeg(&self, options: litchi_imgconv::Options) -> Result<Vec<u8>> {
        litchi_imgconv::to_jpeg(&self.blip()?, options)
    }

    /// Converts an EMF or WMF image to SVG under explicit limits.
    #[cfg(feature = "imgconv")]
    pub fn to_svg(&self, options: litchi_imgconv::Options) -> Result<String> {
        litchi_imgconv::to_svg(&self.blip()?, options)
    }

    /// Extracts the recommended representation under explicit limits.
    #[cfg(feature = "imgconv")]
    pub fn extract(&self, options: litchi_imgconv::Options) -> Result<Vec<u8>> {
        match self.kind {
            Kind::Emf | Kind::Wmf => self.to_svg(options).map(String::into_bytes),
            Kind::Pict => self.to_png(options),
            Kind::Jpeg | Kind::CmykJpeg | Kind::Png | Kind::Dib | Kind::Tiff => {
                self.decode(options.limits).map(|data| data.into_owned())
            },
            Kind::Error | Kind::Unknown | Kind::Other(_) => Err(Error::Unsupported(
                "unknown OfficeArt images cannot be decoded".to_string(),
            )),
        }
    }
}

/// Borrowed image discovery for OfficeArt-bearing OLE streams.
pub struct ImageExtractor;

impl ImageExtractor {
    /// Finds the unique BStore below an OfficeArt drawing root.
    pub fn store(data: &[u8]) -> Result<Option<Store<'_>>> {
        let mut found = None;
        for record in Parser::new(data).records() {
            let record = record.map_err(parse_error)?;
            if record.kind() == RecordKind::BStoreContainer {
                set_store(&mut found, record)?;
                continue;
            }
            if record.is_container() {
                let container = Container::try_new(record).map_err(parse_error)?;
                for store in container
                    .find_recursive(RecordKind::BStoreContainer)
                    .map_err(parse_error)?
                {
                    set_store(&mut found, store)?;
                }
            }
        }
        found
            .map(Store::from_record)
            .transpose()
            .map_err(parse_error)
    }

    /// Resolves one checked semantic BStore ID against an optional delay store.
    pub fn resolve<'data>(
        store: &Store<'data>,
        id: Id,
        delay: Option<&'data [u8]>,
    ) -> Result<Option<ExtractedImage<'data>>> {
        let Some(block) = store.get(id).map_err(parse_error)? else {
            return Ok(None);
        };
        let index = usize::from(id.get() - 1);
        match block {
            Block::Blip(blip) => Ok(Some(ExtractedImage::new(blip, None, index))),
            Block::Entry(entry) => image_from_entry(entry, delay, index),
        }
    }

    /// Resolves every semantic BStore slot in one-based ID order.
    pub fn from_store<'data>(
        store: &Store<'data>,
        delay: Option<&'data [u8]>,
    ) -> Result<Vec<ExtractedImage<'data>>> {
        let mut images = Vec::with_capacity(usize::from(store.len()));
        for raw in 1..=u32::from(store.len()) {
            let id = Id::new(raw).map_err(parse_error)?;
            if let Some(image) = Self::resolve(store, id, delay)? {
                images.push(image);
            }
        }
        Ok(images)
    }

    /// Extracts an image from a direct BLIP or FBSE record.
    pub fn from_record<'data>(record: &Record<'data>) -> Result<ExtractedImage<'data>> {
        Self::from_record_with_delay(record, None)
    }

    /// Extracts an image from a direct BLIP or FBSE with a host delay store.
    pub fn from_record_with_delay<'data>(
        record: &Record<'data>,
        delay: Option<&'data [u8]>,
    ) -> Result<ExtractedImage<'data>> {
        if record.kind().is_blip() {
            let blip = Blip::from_record(record.clone()).map_err(parse_error)?;
            return Ok(ExtractedImage::new(blip, None, 0));
        }
        if record.kind() == RecordKind::Bse {
            let entry = Entry::parse(record.clone()).map_err(parse_error)?;
            return image_from_entry(entry, delay, 0)?.ok_or_else(|| {
                Error::ParseError("OfficeArt FBSE is an empty image slot".to_string())
            });
        }
        Err(Error::ParseError(format!(
            "OfficeArt record 0x{:04X} is not an image file block",
            record.raw_kind()
        )))
    }

    /// Recursively discovers BLIPs and FBSEs in an OfficeArt record sequence.
    pub fn blips(data: &[u8]) -> Result<Vec<ExtractedImage<'_>>> {
        let mut images = Vec::new();
        collect_sequence(data, None, &mut images)?;
        for (index, image) in images.iter_mut().enumerate() {
            image.index = index;
        }
        Ok(images)
    }

    /// Discovers images below one already checked OfficeArt container.
    pub fn from_container<'data>(
        container: &Container<'data>,
        delay: Option<&'data [u8]>,
    ) -> Result<Vec<ExtractedImage<'data>>> {
        let mut images = Vec::new();
        collect_sequence(container.record().data(), delay, &mut images)?;
        for (index, image) in images.iter_mut().enumerate() {
            image.index = index;
        }
        Ok(images)
    }

    /// Parses the headerless PPT `Pictures` BStoreDelay sequence.
    pub fn pictures(data: &[u8]) -> Result<Vec<ExtractedImage<'_>>> {
        let delay = Delay::new(data);
        let mut images = Vec::new();
        for block in delay.iter() {
            let index = images.len();
            match block.map_err(parse_error)? {
                Block::Blip(blip) => images.push(ExtractedImage::new(blip, None, index)),
                Block::Entry(entry) => {
                    if let Some(image) = image_from_entry(entry, Some(data), index)? {
                        images.push(image);
                    }
                },
            }
        }
        Ok(images)
    }

    /// Searches a DOC data stream for validated BLIP record signatures.
    fn search(data: &[u8]) -> Result<Vec<ExtractedImage<'_>>> {
        let mut images = Vec::new();
        let mut offset = 0usize;
        while data.len().saturating_sub(offset) >= 8 {
            match Record::parse(data, offset) {
                Ok((record, consumed)) if record.kind().is_blip() => {
                    if let Ok(blip) = Blip::from_record(record) {
                        check_image_count(images.len(), 1)?;
                        images.push(ExtractedImage::new(blip, None, images.len()));
                        if let Some(next) = offset.checked_add(consumed) {
                            offset = next;
                            continue;
                        }
                        break;
                    }
                },
                _ => {},
            }
            offset += 1;
        }
        Ok(images)
    }
}

fn set_store<'data>(slot: &mut Option<Record<'data>>, record: Record<'data>) -> Result<()> {
    if slot.replace(record).is_some() {
        return Err(Error::ParseError(
            "OfficeArt drawing contains multiple BStore containers".to_string(),
        ));
    }
    Ok(())
}

fn entry_name(entry: &Entry<'_>) -> Result<Option<String>> {
    entry
        .name()
        .map(|name| name.to_string().map_err(parse_error))
        .transpose()
}

fn image_from_entry<'data>(
    entry: Entry<'data>,
    delay: Option<&'data [u8]>,
    index: usize,
) -> Result<Option<ExtractedImage<'data>>> {
    let name = entry_name(&entry)?;
    let context = delay.map_or_else(Context::new, |data| {
        Context::new().with_delay(Delay::new(data))
    });
    entry
        .resolve(context)
        .map_err(parse_error)
        .map(|blip| blip.map(|blip| ExtractedImage::new(blip, name, index)))
}

fn check_image_count(current: usize, additional: usize) -> Result<()> {
    if current
        .checked_add(additional)
        .is_none_or(|count| count > MAX_IMAGES)
    {
        return Err(Error::ParseError(
            "OfficeArt image collection exceeds 4095 file blocks".to_string(),
        ));
    }
    Ok(())
}

fn collect_sequence<'data>(
    data: &'data [u8],
    delay: Option<&'data [u8]>,
    images: &mut Vec<ExtractedImage<'data>>,
) -> Result<()> {
    const MAX_DEPTH: usize = 64;
    const MAX_RECORDS: u32 = 1_000_000;

    let mut stack = vec![Children::new(data)];
    let mut visited = 0u32;
    while let Some(records) = stack.last_mut() {
        let Some(record) = records.next() else {
            stack.pop();
            continue;
        };
        let record = record.map_err(parse_error)?;
        visited = visited
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("OfficeArt record count overflow".to_string()))?;
        if visited > MAX_RECORDS {
            return Err(Error::ParseError(
                "OfficeArt image traversal exceeds one million records".to_string(),
            ));
        }
        if record.kind().is_blip() {
            let blip = Blip::from_record(record).map_err(parse_error)?;
            check_image_count(images.len(), 1)?;
            images.push(ExtractedImage::new(blip, None, images.len()));
        } else if record.kind() == RecordKind::Bse {
            let entry = Entry::parse(record).map_err(parse_error)?;
            if let Some(image) = image_from_entry(entry, delay, images.len())? {
                check_image_count(images.len(), 1)?;
                images.push(image);
            }
        } else if record.kind() == RecordKind::BStoreContainer {
            let store = Store::from_record(record).map_err(parse_error)?;
            check_image_count(images.len(), usize::from(store.len()))?;
            images.extend(ImageExtractor::from_store(&store, delay)?);
        } else if record.is_container() {
            if stack.len() >= MAX_DEPTH {
                return Err(Error::ParseError(
                    "OfficeArt image traversal exceeds 64 containers".to_string(),
                ));
            }
            stack.push(Children::new(record.data()));
        }
    }
    Ok(())
}

/// High-level extraction from an OLE-backed PPT file.
pub mod ppt {
    use super::*;
    use crate::OleFile;
    use std::io::{Read, Seek};

    impl ImageExtractor {
        /// Extracts physical records from the PPT `Pictures` stream.
        pub fn from_ppt<R: Read + Seek>(
            ole: &mut OleFile<R>,
        ) -> Result<Vec<ExtractedImage<'static>>> {
            if !ole.exists(&["Pictures"]) {
                return Ok(Vec::new());
            }
            let data = ole.open_stream(&["Pictures"]).map_err(parse_error)?;
            Self::pictures(&data)
                .and_then(|images| images.into_iter().map(ExtractedImage::into_owned).collect())
        }
    }
}

/// High-level extraction from an OLE-backed DOC file.
pub mod doc {
    use super::*;
    use crate::OleFile;
    use std::io::{Read, Seek};

    impl ImageExtractor {
        /// Searches the DOC `Data` stream for validated native BLIPs.
        pub fn from_doc<R: Read + Seek>(
            ole: &mut OleFile<R>,
        ) -> Result<Vec<ExtractedImage<'static>>> {
            if !ole.exists(&["Data"]) {
                return Ok(Vec::new());
            }
            let data = ole.open_stream(&["Data"]).map_err(parse_error)?;
            Self::search(&data)?
                .into_iter()
                .map(ExtractedImage::into_owned)
                .collect()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn png_blip(data: &[u8]) -> Vec<u8> {
        let mut payload = vec![0; 16];
        payload.push(0xff);
        payload.extend_from_slice(data);
        let mut record = Vec::new();
        record.extend_from_slice(&(0x6e0u16 << 4).to_le_bytes());
        record.extend_from_slice(&0xf01eu16.to_le_bytes());
        record.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        record.extend_from_slice(&payload);
        record
    }

    fn delayed_fbse(blip: &[u8], offset: u32, name: &[u8]) -> Vec<u8> {
        let mut payload = vec![Kind::Png.raw(), Kind::Png.raw()];
        payload.extend_from_slice(&[0; 16]);
        payload.extend_from_slice(&0xffu16.to_le_bytes());
        payload.extend_from_slice(&(blip.len() as u32).to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&offset.to_le_bytes());
        payload.push(0);
        payload.push(name.len() as u8);
        payload.extend_from_slice(&[0, 0]);
        payload.extend_from_slice(name);
        let mut record = Vec::new();
        record.extend_from_slice(&0x62u16.to_le_bytes());
        record.extend_from_slice(&0xf007u16.to_le_bytes());
        record.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        record.extend_from_slice(&payload);
        record
    }

    #[test]
    fn extracted_image_reparses_owned_bytes_without_self_reference() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes).expect("valid PNG");
        let image = ExtractedImage::new(blip, Some("photo.old".to_string()), 4)
            .into_owned()
            .expect("copy image");
        assert_eq!(image.kind(), Kind::Png);
        assert_eq!(image.data().expect("image data"), b"png");
        assert_eq!(image.suggested_filename(), "photo.png");
    }

    #[test]
    fn resolves_offset_zero_from_bstore_metadata() {
        let blip = png_blip(b"png");
        let fbse = delayed_fbse(&blip, 0, &[]);
        let mut store = vec![0x1f, 0, 0x01, 0xf0];
        store.extend_from_slice(&(fbse.len() as u32).to_le_bytes());
        store.extend_from_slice(&fbse);
        let store = ImageExtractor::store(&store)
            .expect("valid store")
            .expect("store exists");
        let image = ImageExtractor::resolve(&store, Id::new(1).expect("valid ID"), Some(&blip))
            .expect("resolve")
            .expect("image");
        assert_eq!(image.data().expect("image data"), b"png");
    }

    #[test]
    fn pictures_is_a_headerless_delay_sequence() {
        let first = png_blip(b"one");
        let second = png_blip(b"two");
        let mut pictures = first;
        pictures.extend_from_slice(&second);
        let images = ImageExtractor::pictures(&pictures).expect("valid Pictures stream");
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].data().expect("first image"), b"one");
        assert_eq!(images[1].data().expect("second image"), b"two");
    }

    #[test]
    fn recursive_discovery_caps_the_semantic_image_collection() {
        let data = png_blip(b"x").repeat(MAX_IMAGES + 1);
        assert!(matches!(
            ImageExtractor::blips(&data),
            Err(Error::ParseError(message)) if message.contains("4095")
        ));
    }
}
