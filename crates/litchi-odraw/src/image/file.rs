//! Move-first OfficeArt image-file discovery.
//!
//! A [`File`] is either a zero-copy view over caller-owned bytes or an owned
//! OfficeArt record produced by consuming a view.  The scanner understands
//! OfficeArt framing only; DOC and PPT remain responsible for locating their
//! host streams, and image codecs remain in `litchi-imgconv`.

use super::{Blip, Block, Context, Delay, Entry, Id, Kind, Store};
use crate::{Children, Container, Error, ImageLimit, Limit, Parser, Record, RecordKind, Result};

const MAX_FILES: usize = 0x0fff;
const MAX_DEPTH: usize = 64;
const MAX_RECORDS: u32 = 1_000_000;

/// One validated OfficeArt image file.
#[derive(Debug)]
pub struct File<'data> {
    source: Source<'data>,
    kind: Kind,
    name: Option<String>,
    index: usize,
}

#[derive(Debug)]
enum Source<'data> {
    View(Blip<'data>),
    Owned { record: Vec<u8>, data_start: usize },
}

impl<'data> File<'data> {
    /// Wraps an already validated BLIP without copying its file data.
    pub fn new(blip: Blip<'data>, name: Option<String>, index: usize) -> Self {
        let kind = blip.kind();
        Self {
            source: Source::View(blip),
            kind,
            name,
            index,
        }
    }

    /// Parses a fresh borrowed BLIP view over this file's storage.
    pub fn blip(&self) -> Result<Blip<'_>> {
        match &self.source {
            Source::View(blip) => Ok(blip.clone()),
            Source::Owned { record, .. } => Blip::parse(record),
        }
    }

    /// Returns the native OfficeArt image kind.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Returns the conventional extension for the native image bytes.
    pub const fn extension(&self) -> &'static str {
        self.kind.extension()
    }

    /// Returns the optional producer-supplied filename hint.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the zero-based position in the discovered image collection.
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Reassigns the collection position when a host builds a semantic view.
    pub const fn with_index(mut self, index: usize) -> Self {
        self.index = index;
        self
    }

    /// Returns a suggested filename without changing the native representation.
    pub fn filename(&self) -> String {
        let stem = self
            .name
            .as_deref()
            .and_then(safe_stem)
            .unwrap_or_else(|| format!("image_{:03}", self.index));
        format!("{stem}.{}", self.extension())
    }

    /// Borrows the stored native image bytes without decoding them.
    pub fn data(&self) -> Result<&[u8]> {
        match &self.source {
            Source::View(blip) => Ok(blip.data()),
            Source::Owned { record, data_start } => {
                record.get(*data_start..).ok_or(Error::MalformedImage {
                    reason: "owned BLIP data range is invalid",
                })
            },
        }
    }

    /// Consumes this view into an independently owned OfficeArt record.
    pub fn into_owned(self) -> Result<File<'static>> {
        let (record, data_start) = match self.source {
            Source::Owned { record, data_start } => (record, data_start),
            Source::View(blip) => {
                let data_len = blip.data().len();
                let mut record = Vec::new();
                super::write::copy(&mut record, &blip).map_err(|_| Error::MalformedImage {
                    reason: "owned BLIP framing could not be encoded",
                })?;
                let data_start =
                    record
                        .len()
                        .checked_sub(data_len)
                        .ok_or(Error::ArithmeticOverflow {
                            context: "owned BLIP data range",
                        })?;
                (record, data_start)
            },
        };
        Ok(File {
            source: Source::Owned { record, data_start },
            kind: self.kind,
            name: self.name,
            index: self.index,
        })
    }
}

/// Finds the unique BStore below an OfficeArt drawing root.
pub fn store(data: &[u8]) -> Result<Option<Store<'_>>> {
    let mut found = None;
    for record in Parser::new(data).records() {
        let record = record?;
        if record.kind() == RecordKind::BStoreContainer {
            set_store(&mut found, record)?;
            continue;
        }
        if record.is_container() {
            let container = Container::try_new(record)?;
            for store in container.find_recursive(RecordKind::BStoreContainer)? {
                set_store(&mut found, store)?;
            }
        }
    }
    found.map(Store::from_record).transpose()
}

/// Resolves one checked semantic BStore ID against an optional delay store.
pub fn get<'data>(
    store: &Store<'data>,
    id: Id,
    delay_store: Option<&'data [u8]>,
) -> Result<Option<File<'data>>> {
    let Some(block) = store.get(id)? else {
        return Ok(None);
    };
    let index = usize::from(id.get() - 1);
    match block {
        Block::Blip(blip) => Ok(Some(File::new(blip, None, index))),
        Block::Entry(entry) => file_from_entry(entry, delay_store, index),
    }
}

/// Resolves every semantic BStore slot in one-based ID order.
pub fn all<'data>(
    store: &Store<'data>,
    delay_store: Option<&'data [u8]>,
) -> Result<Vec<File<'data>>> {
    let mut files = Vec::with_capacity(usize::from(store.len()));
    for raw in 1..=u32::from(store.len()) {
        let id = Id::new(raw)?;
        if let Some(file) = get(store, id, delay_store)? {
            files.push(file);
        }
    }
    Ok(files)
}

/// Extracts an image from a direct BLIP or FBSE record.
pub fn record<'data>(record: &Record<'data>) -> Result<File<'data>> {
    record_with_delay(record, None)
}

/// Extracts an image from a direct BLIP or FBSE with a host delay store.
pub fn record_with_delay<'data>(
    record: &Record<'data>,
    delay_store: Option<&'data [u8]>,
) -> Result<File<'data>> {
    if record.kind().is_blip() {
        let blip = Blip::from_record(record.clone())?;
        return Ok(File::new(blip, None, 0));
    }
    if record.kind() == RecordKind::Bse {
        let entry = Entry::parse(record.clone())?;
        return file_from_entry(entry, delay_store, 0)?.ok_or(Error::MalformedImage {
            reason: "FBSE is an empty image slot",
        });
    }
    Err(Error::NotImageRecord {
        raw_kind: record.raw_kind(),
    })
}

/// Recursively discovers image files in an OfficeArt record sequence.
pub fn scan(data: &[u8]) -> Result<Vec<File<'_>>> {
    let mut files = Vec::new();
    collect(data, None, &mut files)?;
    for (index, file) in files.iter_mut().enumerate() {
        file.index = index;
    }
    Ok(files)
}

/// Discovers files below one already checked OfficeArt container.
pub fn container<'data>(
    container: &Container<'data>,
    delay_store: Option<&'data [u8]>,
) -> Result<Vec<File<'data>>> {
    let mut files = Vec::new();
    collect(container.record().data(), delay_store, &mut files)?;
    for (index, file) in files.iter_mut().enumerate() {
        file.index = index;
    }
    Ok(files)
}

/// Parses a headerless BStoreDelay sequence, such as PPT's `Pictures` stream.
pub fn delay(data: &[u8]) -> Result<Vec<File<'_>>> {
    let delay = Delay::new(data);
    let mut files = Vec::new();
    for block in delay.iter() {
        let index = files.len();
        match block? {
            Block::Blip(blip) => {
                check_count(files.len(), 1)?;
                files.push(File::new(blip, None, index));
            },
            Block::Entry(entry) => {
                if let Some(file) = file_from_entry(entry, Some(data), index)? {
                    check_count(files.len(), 1)?;
                    files.push(file);
                }
            },
        }
    }
    Ok(files)
}

fn safe_stem(name: &str) -> Option<String> {
    const MAX_STEM_BYTES: usize = 120;

    let basename = name.rsplit(['/', '\\']).next()?;
    let stem = basename
        .rsplit_once('.')
        .map_or(basename, |(stem, _)| stem)
        .trim_matches([' ', '.']);
    let mut safe = String::with_capacity(stem.len().min(MAX_STEM_BYTES));
    for character in stem.chars() {
        let character = if character.is_control()
            || matches!(character, '<' | '>' | ':' | '"' | '|' | '?' | '*')
        {
            '_'
        } else {
            character
        };
        if safe.len() + character.len_utf8() > MAX_STEM_BYTES {
            break;
        }
        safe.push(character);
    }
    if safe.is_empty() || safe == "." || safe == ".." {
        return None;
    }
    if is_windows_device_name(&safe) {
        safe.push('_');
    }
    Some(safe)
}

fn is_windows_device_name(stem: &str) -> bool {
    let upper = stem.to_ascii_uppercase();
    matches!(upper.as_str(), "CON" | "PRN" | "AUX" | "NUL")
        || upper
            .strip_prefix("COM")
            .or_else(|| upper.strip_prefix("LPT"))
            .is_some_and(|suffix| {
                matches!(suffix, "1" | "2" | "3" | "4" | "5" | "6" | "7" | "8" | "9")
            })
}

fn set_store<'data>(slot: &mut Option<Record<'data>>, record: Record<'data>) -> Result<()> {
    if slot.replace(record).is_some() {
        return Err(Error::MalformedImage {
            reason: "drawing contains multiple BStore containers",
        });
    }
    Ok(())
}

fn entry_name(entry: &Entry<'_>) -> Result<Option<String>> {
    entry.name().map(|name| name.to_string()).transpose()
}

fn file_from_entry<'data>(
    entry: Entry<'data>,
    delay_store: Option<&'data [u8]>,
    index: usize,
) -> Result<Option<File<'data>>> {
    let name = entry_name(&entry)?;
    let context = delay_store.map_or_else(Context::new, |data| {
        Context::new().with_delay(Delay::new(data))
    });
    entry
        .resolve(context)
        .map(|blip| blip.map(|blip| File::new(blip, name, index)))
}

fn check_count(current: usize, additional: usize) -> Result<()> {
    if current
        .checked_add(additional)
        .is_none_or(|count| count > MAX_FILES)
    {
        return Err(Error::ImageLimitExceeded {
            limit: ImageLimit::StoreEntries,
            maximum: MAX_FILES as u64,
        });
    }
    Ok(())
}

fn collect<'data>(
    data: &'data [u8],
    delay_store: Option<&'data [u8]>,
    files: &mut Vec<File<'data>>,
) -> Result<()> {
    let mut stack = vec![Children::new(data)];
    let mut visited = 0u32;
    while let Some(records) = stack.last_mut() {
        let Some(record) = records.next() else {
            stack.pop();
            continue;
        };
        let record = record?;
        visited = visited.checked_add(1).ok_or(Error::ArithmeticOverflow {
            context: "image record count",
        })?;
        if visited > MAX_RECORDS {
            return Err(Error::LimitExceeded {
                limit: Limit::Records,
                maximum: MAX_RECORDS,
            });
        }
        if record.kind().is_blip() {
            let blip = Blip::from_record(record)?;
            check_count(files.len(), 1)?;
            files.push(File::new(blip, None, files.len()));
        } else if record.kind() == RecordKind::Bse {
            let entry = Entry::parse(record)?;
            if let Some(file) = file_from_entry(entry, delay_store, files.len())? {
                check_count(files.len(), 1)?;
                files.push(file);
            }
        } else if record.kind() == RecordKind::BStoreContainer {
            let store = Store::from_record(record)?;
            check_count(files.len(), usize::from(store.len()))?;
            files.extend(all(&store, delay_store)?);
        } else if record.is_container() {
            if stack.len() >= MAX_DEPTH {
                return Err(Error::LimitExceeded {
                    limit: Limit::Depth,
                    maximum: MAX_DEPTH as u32,
                });
            }
            stack.push(Children::new(record.data()));
        }
    }
    Ok(())
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
    fn owned_file_reparses_without_a_self_reference() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes).expect("valid PNG");
        let file = File::new(blip, Some("photo.old".to_string()), 4)
            .into_owned()
            .expect("copy image");
        assert_eq!(file.kind(), Kind::Png);
        assert_eq!(file.data().expect("image data"), b"png");
        assert_eq!(file.filename(), "photo.png");
    }

    #[test]
    fn producer_names_cannot_escape_the_output_directory() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes).expect("valid PNG");
        let file = File::new(blip, Some("../../unsafe:name.old".to_string()), 4);
        assert_eq!(file.filename(), "unsafe_name.png");
    }

    #[test]
    fn producer_names_are_portable_and_bounded() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes).expect("valid PNG");
        let reserved = File::new(blip.clone(), Some("CON.old".to_string()), 0);
        assert_eq!(reserved.filename(), "CON_.png");

        let long = File::new(blip, Some(format!("{}.old", "🦀".repeat(100))), 0);
        assert!(long.filename().len() <= 124);
        assert!(long.filename().ends_with(".png"));
    }

    #[test]
    fn host_can_assign_a_stable_semantic_position() {
        let bytes = png_blip(b"png");
        let file = File::new(Blip::parse(&bytes).unwrap(), None, 0).with_index(7);
        assert_eq!(file.index(), 7);
        assert_eq!(file.filename(), "image_007.png");
    }

    #[test]
    fn resolves_offset_zero_from_store_metadata() {
        let blip = png_blip(b"png");
        let fbse = delayed_fbse(&blip, 0, &[]);
        let mut bytes = vec![0x1f, 0, 0x01, 0xf0];
        bytes.extend_from_slice(&(fbse.len() as u32).to_le_bytes());
        bytes.extend_from_slice(&fbse);
        let store = store(&bytes).expect("valid store").expect("store exists");
        let file = get(&store, Id::new(1).expect("valid ID"), Some(&blip))
            .expect("resolve")
            .expect("image");
        assert_eq!(file.data().expect("image data"), b"png");
    }

    #[test]
    fn delay_accepts_a_headerless_record_sequence() {
        let first = png_blip(b"one");
        let second = png_blip(b"two");
        let mut pictures = first;
        pictures.extend_from_slice(&second);
        let files = delay(&pictures).expect("valid delay stream");
        assert_eq!(files.len(), 2);
        assert_eq!(files[0].data().expect("first image"), b"one");
        assert_eq!(files[1].data().expect("second image"), b"two");
    }

    #[test]
    fn recursive_scan_caps_the_semantic_collection() {
        let data = png_blip(b"x").repeat(MAX_FILES + 1);
        assert!(matches!(
            scan(&data),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimit::StoreEntries,
                maximum: 4095,
            })
        ));
    }
}
