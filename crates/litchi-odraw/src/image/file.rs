//! Move-first `OfficeArt` image-file discovery.
//!
//! A [`File`] is either a zero-copy view over caller-owned bytes or an owned
//! `OfficeArt` record produced by consuming a view.  The scanner understands
//! `OfficeArt` framing only; DOC and PPT remain responsible for locating their
//! host streams, and image codecs remain in `litchi-imgconv`.

use super::LimitKind as ImageLimitKind;
use super::{Blip, Block, Context, Delay, Entry, Id, Kind, Limits, Store};
use crate::{Container, Error, Parser, Record, RecordKind, Result};

/// One validated `OfficeArt` image file.
#[derive(Debug)]
pub struct File<'data> {
    source: Source<'data>,
    name: Option<String>,
    index: usize,
}

#[derive(Debug)]
enum Source<'data> {
    View {
        blip: Blip<'data>,
        max_blip_bytes: u32,
    },
    Owned {
        record: Vec<u8>,
        data_start: usize,
        raw_limits: crate::Limits,
        max_blip_bytes: u32,
        kind: Kind,
    },
}

impl<'data> File<'data> {
    /// Wraps an already validated BLIP without copying its file data.
    #[must_use]
    pub fn new(blip: Blip<'data>, name: Option<String>, index: usize, limits: Limits) -> Self {
        let raw_limits = blip.record().limits().constrained_by(limits.records);
        let blip = blip.with_record_limits(raw_limits);
        Self {
            source: Source::View {
                blip,
                max_blip_bytes: limits.max_blip_bytes,
            },
            name,
            index,
        }
    }

    /// Parses a fresh borrowed BLIP view over this file's storage.
    ///
    /// # Errors
    ///
    /// Returns an error when an owned record cannot be re-parsed as a valid
    /// BLIP.
    pub fn blip(&self) -> Result<Blip<'_>> {
        match &self.source {
            Source::View { blip, .. } => Ok(blip.clone()),
            Source::Owned {
                record,
                raw_limits,
                max_blip_bytes,
                ..
            } => Blip::parse(
                record,
                Limits {
                    records: *raw_limits,
                    max_blip_bytes: *max_blip_bytes,
                    ..Limits::default()
                },
            ),
        }
    }

    /// Returns the native `OfficeArt` image kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        match &self.source {
            Source::View { blip, .. } => blip.kind(),
            Source::Owned { kind, .. } => *kind,
        }
    }

    /// Returns the conventional extension for the native image bytes.
    #[must_use]
    pub const fn extension(&self) -> &'static str {
        self.kind().extension()
    }

    /// Returns the optional producer-supplied filename hint.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the zero-based position in the discovered image collection.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Reassigns the collection position when a host builds a semantic view.
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedImage` when the owned BLIP data range is
    /// invalid.
    pub fn data(&self) -> Result<&[u8]> {
        match &self.source {
            Source::View { blip, .. } => Ok(blip.data()),
            Source::Owned {
                record, data_start, ..
            } => record.get(*data_start..).ok_or(Error::MalformedImage {
                reason: "owned BLIP data range is invalid",
            }),
        }
    }

    /// Consumes this view into an independently owned `OfficeArt` record.
    ///
    /// # Errors
    ///
    /// Returns `Error::ImageLimitExceeded` when the retained raw or BLIP
    /// budget no longer admits the record, `Error::MalformedImage` when the
    /// BLIP framing cannot be encoded, or `Error::ArithmeticOverflow` when
    /// the owned data range underflows.
    pub fn into_owned(self) -> Result<File<'static>> {
        let kind = self.kind();
        let (record, data_start, raw_limits, max_blip_bytes) = match self.source {
            Source::Owned {
                record,
                data_start,
                raw_limits,
                max_blip_bytes,
                ..
            } => (record, data_start, raw_limits, max_blip_bytes),
            Source::View {
                blip,
                max_blip_bytes,
            } => {
                let raw_limits = blip.record().limits();
                let record_len = blip.record().len();
                if record_len > raw_limits.max_record_bytes {
                    return Err(Error::ImageLimitExceeded {
                        limit: ImageLimitKind::RecordBytes,
                        observed: u64::from(record_len),
                        maximum: u64::from(raw_limits.max_record_bytes),
                    });
                }
                if record_len > max_blip_bytes {
                    return Err(Error::ImageLimitExceeded {
                        limit: ImageLimitKind::BlipBytes,
                        observed: u64::from(record_len),
                        maximum: u64::from(max_blip_bytes),
                    });
                }
                let limits = Limits {
                    records: raw_limits,
                    max_blip_bytes,
                    ..Limits::default()
                };
                let data_len = blip.data().len();
                let mut record = Vec::new();
                super::write::copy(&mut record, &blip, limits).map_err(|_err| {
                    Error::MalformedImage {
                        reason: "owned BLIP framing could not be encoded",
                    }
                })?;
                let data_start =
                    record
                        .len()
                        .checked_sub(data_len)
                        .ok_or(Error::ArithmeticOverflow {
                            context: "owned BLIP data range",
                        })?;
                (record, data_start, raw_limits, max_blip_bytes)
            },
        };
        Ok(File {
            source: Source::Owned {
                record,
                data_start,
                raw_limits,
                max_blip_bytes,
                kind,
            },
            name: self.name,
            index: self.index,
        })
    }
}

/// Finds the unique `BStore` below an `OfficeArt` drawing root.
///
/// # Errors
///
/// Returns an error when a record is malformed or when the drawing contains
/// multiple `BStore` containers, or when the top-level record traversal
/// exceeds `limits.records.max_records`.
pub fn store(data: &[u8], limits: Limits) -> Result<Option<Store<'_>>> {
    let mut found = None;
    let mut visited = 0_u32;
    // Keep this scanner's image-specific record count classification. The
    // generic Children iterator enforces the same raw budget for callers that
    // consume it directly, but this route reports its own typed image limit.
    let parser_limits = crate::Limits {
        max_records: u32::MAX,
        ..limits.record_limits()
    };
    for result in Parser::new(data, parser_limits).records() {
        let record = result.map_err(super::map_traversal_limit)?;
        visited = visited.checked_add(1).ok_or(Error::ImageLimitExceeded {
            limit: ImageLimitKind::Records,
            observed: u64::from(visited) + 1,
            maximum: u64::from(limits.records.max_records),
        })?;
        if visited > limits.records.max_records {
            return Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Records,
                observed: u64::from(visited),
                maximum: u64::from(limits.records.max_records),
            });
        }
        if record.kind() == RecordKind::BStoreContainer {
            set_store(&mut found, record)?;
            continue;
        }
        if record.is_container() {
            let container = Container::try_new(record)?;
            let stores = container
                .find_recursive(RecordKind::BStoreContainer, limits.record_limits())
                .map_err(super::map_traversal_limit)?;
            for store in stores {
                set_store(&mut found, store)?;
            }
        }
    }
    found
        .map(|record| Store::from_record(record, limits))
        .transpose()
}

/// Resolves one checked semantic `BStore` ID against an optional delay store.
///
/// # Errors
///
/// Returns an error when the store metadata is malformed or a delayed BLIP
/// cannot be resolved.
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
        Block::Blip(blip) => Ok(Some(File::new(blip, None, index, store.limits()))),
        Block::Entry(entry) => file_from_entry(&entry, delay_store, index),
    }
}

/// Resolves every semantic `BStore` slot in one-based ID order.
///
/// # Errors
///
/// Returns an error when a store slot ID is invalid or any slot cannot be
/// resolved.
pub fn all<'data>(
    store: &Store<'data>,
    delay_store: Option<&'data [u8]>,
) -> Result<Vec<File<'data>>> {
    let file_cap = usize::try_from(store.limits().max_files).unwrap_or(usize::MAX);
    let mut files = Vec::with_capacity(usize::from(store.len()).min(file_cap));
    for raw in 1..=u32::from(store.len()) {
        let id = Id::new(raw)?;
        if let Some(file) = get(store, id, delay_store)? {
            check_count(files.len(), 1, store.limits())?;
            files.push(file);
        }
    }
    Ok(files)
}

/// Extracts an image from a direct BLIP or FBSE record.
///
/// # Errors
///
/// Returns `Error::NotImageRecord` when the record is neither a BLIP nor an
/// FBSE, or an error when the record payload is malformed.
pub fn record<'data>(record: &Record<'data>, limits: Limits) -> Result<File<'data>> {
    record_with_delay(record, None, limits)
}

/// Extracts an image from a direct BLIP or FBSE with a host delay store.
///
/// # Errors
///
/// Returns `Error::NotImageRecord` when the record is neither a BLIP nor an
/// FBSE, `Error::MalformedImage` when the payload is malformed or the FBSE
/// slot is empty, or an error when a delayed BLIP cannot be resolved.
pub fn record_with_delay<'data>(
    record: &Record<'data>,
    delay_store: Option<&'data [u8]>,
    limits: Limits,
) -> Result<File<'data>> {
    if record.kind().is_blip() {
        let blip = Blip::from_record(record.clone(), limits)?;
        return Ok(File::new(blip, None, 0, limits));
    }
    if record.kind() == RecordKind::Bse {
        let entry = Entry::parse(record.clone(), limits)?;
        return file_from_entry(&entry, delay_store, 0)?.ok_or(Error::MalformedImage {
            reason: "FBSE is an empty image slot",
        });
    }
    Err(Error::NotImageRecord {
        raw_kind: record.raw_kind(),
    })
}

/// Recursively discovers image files in an `OfficeArt` record sequence.
///
/// # Errors
///
/// Returns an error when a record is malformed or the file, record, or depth
/// limits are exceeded.
pub fn scan(data: &[u8], limits: Limits) -> Result<Vec<File<'_>>> {
    let mut files = Vec::new();
    collect(data, None, &mut files, limits)?;
    for (index, file) in files.iter_mut().enumerate() {
        file.index = index;
    }
    Ok(files)
}

/// Discovers files below one already checked `OfficeArt` container.
///
/// # Errors
///
/// Returns an error when a record below the container is malformed or a
/// discovery limit is exceeded.
pub fn container<'data>(
    container: &Container<'data>,
    delay_store: Option<&'data [u8]>,
    limits: Limits,
) -> Result<Vec<File<'data>>> {
    let limits = Limits {
        records: container.record().limits().constrained_by(limits.records),
        ..limits
    };
    let mut files = Vec::new();
    collect(container.record().data(), delay_store, &mut files, limits)?;
    for (index, file) in files.iter_mut().enumerate() {
        file.index = index;
    }
    Ok(files)
}

/// Parses a headerless `BStoreDelay` sequence, such as PPT's `Pictures` stream.
///
/// # Errors
///
/// Returns an error when a delay record is malformed, a delayed BLIP cannot
/// be resolved, or the file count limit is exceeded.
pub fn delay(data: &[u8], limits: Limits) -> Result<Vec<File<'_>>> {
    let delay = Delay::new(data, limits);
    let mut files = Vec::new();
    for block in delay.iter() {
        let index = files.len();
        match block? {
            Block::Blip(blip) => {
                check_count(files.len(), 1, limits)?;
                files.push(File::new(blip, None, index, limits));
            },
            Block::Entry(entry) => {
                if let Some(file) = file_from_entry(&entry, Some(data), index)? {
                    check_count(files.len(), 1, limits)?;
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
        let sanitized = if character.is_control()
            || matches!(character, '<' | '>' | ':' | '"' | '|' | '?' | '*')
        {
            '_'
        } else {
            character
        };
        if safe.len() + sanitized.len_utf8() > MAX_STEM_BYTES {
            break;
        }
        safe.push(sanitized);
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
    entry.name().map(super::model::Name::to_string).transpose()
}

fn file_from_entry<'data>(
    entry: &Entry<'data>,
    delay_store: Option<&'data [u8]>,
    index: usize,
) -> Result<Option<File<'data>>> {
    let name = entry_name(entry)?;
    let context = delay_store.map_or_else(Context::new, |data| {
        Context::new().with_delay(Delay::new(data, entry.limits))
    });
    entry
        .resolve(context)
        .map(|resolved| resolved.map(|blip| File::new(blip, name, index, entry.limits)))
}

fn check_count(current: usize, additional: usize, limits: Limits) -> Result<()> {
    if current
        .checked_add(additional)
        .is_none_or(|count| count > usize::try_from(limits.max_files).unwrap_or(usize::MAX))
    {
        return Err(Error::ImageLimitExceeded {
            limit: ImageLimitKind::Files,
            observed: u64::try_from(current.saturating_add(additional)).unwrap_or(u64::MAX),
            maximum: u64::from(limits.max_files),
        });
    }
    Ok(())
}

fn collect<'data>(
    data: &'data [u8],
    delay_store: Option<&'data [u8]>,
    files: &mut Vec<File<'data>>,
    limits: Limits,
) -> Result<()> {
    let mut stack = vec![crate::container::Frame::new(data)];
    let mut visited = 0u32;
    while let Some(frame) = stack.last_mut() {
        let Some(result) = frame.next(limits.record_limits()) else {
            stack.pop();
            continue;
        };
        let record = result.map_err(super::map_traversal_limit)?;
        visited = visited.checked_add(1).ok_or(Error::ImageLimitExceeded {
            limit: ImageLimitKind::Records,
            observed: u64::from(visited) + 1,
            maximum: u64::from(limits.records.max_records),
        })?;
        if visited > limits.records.max_records {
            return Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Records,
                observed: u64::from(visited),
                maximum: u64::from(limits.records.max_records),
            });
        }
        if record.kind().is_blip() {
            let blip = Blip::from_record(record, limits)?;
            check_count(files.len(), 1, limits)?;
            files.push(File::new(blip, None, files.len(), limits));
        } else if record.kind() == RecordKind::Bse {
            let entry = Entry::parse(record, limits)?;
            if let Some(file) = file_from_entry(&entry, delay_store, files.len())? {
                check_count(files.len(), 1, limits)?;
                files.push(file);
            }
        } else if record.kind() == RecordKind::BStoreContainer {
            let store = Store::from_record(record, limits)?;
            for raw in 1..=u32::from(store.len()) {
                let id = Id::new(raw)?;
                if let Some(file) = get(&store, id, delay_store)? {
                    check_count(files.len(), 1, limits)?;
                    files.push(file);
                }
            }
        } else if record.is_container() {
            let depth = stack.len().saturating_add(1);
            if depth > usize::from(limits.records.max_depth) {
                return Err(Error::ImageLimitExceeded {
                    limit: ImageLimitKind::Depth,
                    observed: u64::try_from(depth).unwrap_or(u64::MAX),
                    maximum: u64::from(limits.records.max_depth),
                });
            }
            stack.push(crate::container::Frame::new(record.data()));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    fn png_blip(data: &[u8]) -> Vec<u8> {
        let mut payload = vec![0; 16];
        payload.push(0xff);
        payload.extend_from_slice(data);
        let mut record = Vec::new();
        record.extend_from_slice(&(0x6e0u16 << 4).to_le_bytes());
        record.extend_from_slice(&0xf01eu16.to_le_bytes());
        record.extend_from_slice(
            &u32::try_from(payload.len())
                .expect("payload length fits in u32")
                .to_le_bytes(),
        );
        record.extend_from_slice(&payload);
        record
    }

    fn delayed_fbse(blip: &[u8], offset: u32, name: &[u8]) -> Vec<u8> {
        let mut payload = vec![Kind::Png.raw(), Kind::Png.raw()];
        payload.extend_from_slice(&[0; 16]);
        payload.extend_from_slice(&0xffu16.to_le_bytes());
        payload.extend_from_slice(
            &u32::try_from(blip.len())
                .expect("BLIP length fits in u32")
                .to_le_bytes(),
        );
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&offset.to_le_bytes());
        payload.push(0);
        payload.push(u8::try_from(name.len()).expect("name length fits in u8"));
        payload.extend_from_slice(&[0, 0]);
        payload.extend_from_slice(name);
        let mut record = Vec::new();
        record.extend_from_slice(&0x62u16.to_le_bytes());
        record.extend_from_slice(&0xf007u16.to_le_bytes());
        record.extend_from_slice(
            &u32::try_from(payload.len())
                .expect("payload length fits in u32")
                .to_le_bytes(),
        );
        record.extend_from_slice(&payload);
        record
    }

    #[test]
    fn owned_file_reparses_without_a_self_reference() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes, Default::default()).expect("valid PNG");
        let file = File::new(blip, Some("photo.old".to_string()), 4, Limits::default())
            .into_owned()
            .expect("copy image");
        assert_eq!(file.kind(), Kind::Png);
        assert_eq!(file.data().expect("image data"), b"png");
        assert_eq!(file.filename(), "photo.png");
    }

    #[test]
    fn owned_file_reparse_uses_the_admission_policy_it_retains() {
        let bytes = png_blip(b"png");
        let limits = Limits {
            records: crate::Limits {
                max_record_bytes: 19,
                ..crate::Limits::default()
            },
            max_blip_bytes: 19,
            ..Limits::default()
        };
        let file = File {
            source: Source::Owned {
                record: bytes,
                data_start: 25,
                raw_limits: limits.records,
                max_blip_bytes: limits.max_blip_bytes,
                kind: Kind::Png,
            },
            name: None,
            index: 0,
        };
        assert!(matches!(
            file.blip(),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::RecordBytes,
                observed: 20,
                maximum: 19,
            })
        ));
    }

    #[test]
    fn file_view_and_detached_owner_retain_the_intersected_raw_policy() {
        let bytes = png_blip(b"png");
        let body_len = u32::try_from(bytes.len() - 8).expect("body length fits");
        let origin = crate::Limits {
            max_record_bytes: body_len,
            max_depth: 2,
            max_records: 3,
        };
        let blip = Blip::parse(
            &bytes,
            Limits {
                records: origin,
                max_blip_bytes: body_len,
                ..Limits::default()
            },
        )
        .expect("origin policy admits BLIP");
        let strict_caller = Limits {
            records: crate::Limits {
                max_record_bytes: body_len,
                max_depth: 1,
                max_records: 1,
            },
            max_blip_bytes: body_len,
            ..Limits::default()
        };
        let file = File::new(blip, None, 0, strict_caller);
        assert_eq!(
            file.blip().expect("view BLIP").record().limits(),
            strict_caller.records
        );
        let detached = file.into_owned().expect("detached BLIP");
        assert_eq!(
            detached.blip().expect("owned BLIP").record().limits(),
            strict_caller.records
        );

        let permissive_caller = Limits {
            records: crate::Limits::default(),
            max_blip_bytes: body_len,
            ..Limits::default()
        };
        let blip = Blip::parse(
            &bytes,
            Limits {
                records: origin,
                max_blip_bytes: body_len,
                ..Limits::default()
            },
        )
        .expect("origin policy admits BLIP");
        let file = File::new(blip, None, 0, permissive_caller);
        assert_eq!(
            file.blip()
                .expect("origin-restricted BLIP")
                .record()
                .limits(),
            origin
        );
    }

    #[test]
    fn detaching_a_view_rechecks_the_retained_blip_budget() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes, Limits::default()).expect("valid PNG");
        let body_len = u32::try_from(bytes.len() - 8).expect("body length fits");
        let raw_limited = File::new(
            blip.clone(),
            None,
            0,
            Limits {
                records: crate::Limits {
                    max_record_bytes: body_len - 1,
                    ..crate::Limits::default()
                },
                ..Limits::default()
            },
        );
        assert!(matches!(
            raw_limited.into_owned(),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::RecordBytes,
                observed: 20,
                maximum: 19,
            })
        ));
        let file = File::new(
            blip,
            None,
            0,
            Limits {
                max_blip_bytes: body_len - 1,
                ..Limits::default()
            },
        );
        assert!(matches!(
            file.into_owned(),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::BlipBytes,
                observed: 20,
                maximum: 19,
            })
        ));
    }

    #[test]
    fn producer_names_cannot_escape_the_output_directory() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes, Default::default()).expect("valid PNG");
        let file = File::new(
            blip,
            Some("../../unsafe:name.old".to_string()),
            4,
            Limits::default(),
        );
        assert_eq!(file.filename(), "unsafe_name.png");
    }

    #[test]
    fn producer_names_are_portable_and_bounded() {
        let bytes = png_blip(b"png");
        let blip = Blip::parse(&bytes, Default::default()).expect("valid PNG");
        let reserved = File::new(
            blip.clone(),
            Some("CON.old".to_string()),
            0,
            Limits::default(),
        );
        assert_eq!(reserved.filename(), "CON_.png");

        let long = File::new(
            blip,
            Some(format!("{}.old", "🦀".repeat(100))),
            0,
            Limits::default(),
        );
        assert!(long.filename().len() <= 124);
        assert!(
            std::path::Path::new(&long.filename())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("png"))
        );
    }

    #[test]
    fn host_can_assign_a_stable_semantic_position() {
        let bytes = png_blip(b"png");
        let file = File::new(
            Blip::parse(&bytes, Default::default()).unwrap(),
            None,
            0,
            Limits::default(),
        )
        .with_index(7);
        assert_eq!(file.index(), 7);
        assert_eq!(file.filename(), "image_007.png");
    }

    #[test]
    fn resolves_offset_zero_from_store_metadata() {
        let blip = png_blip(b"png");
        let fbse = delayed_fbse(&blip, 0, &[]);
        let mut bytes = vec![0x1f, 0, 0x01, 0xf0];
        bytes.extend_from_slice(
            &u32::try_from(fbse.len())
                .expect("FBSE length fits in u32")
                .to_le_bytes(),
        );
        bytes.extend_from_slice(&fbse);
        let store = store(&bytes, Limits::default())
            .expect("valid store")
            .expect("store exists");
        let file = get(&store, Id::new(1).expect("valid ID"), Some(&blip))
            .expect("resolve")
            .expect("image");
        assert_eq!(file.data().expect("image data"), b"png");
    }

    #[test]
    fn all_enforces_the_retained_file_budget() {
        let blip = png_blip(b"png");
        let fbse = delayed_fbse(&blip, 0, &[]);
        let mut bytes = vec![0x1f, 0, 0x01, 0xf0];
        bytes.extend_from_slice(
            &u32::try_from(fbse.len())
                .expect("FBSE length fits in u32")
                .to_le_bytes(),
        );
        bytes.extend_from_slice(&fbse);
        let store = Store::parse(
            &bytes,
            Limits {
                max_files: 0,
                ..Limits::default()
            },
        )
        .expect("valid store");
        assert!(matches!(
            all(&store, Some(&blip)),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Files,
                observed: 1,
                maximum: 0,
            })
        ));
    }

    #[test]
    fn delayed_resolution_retains_the_entry_limit_with_a_broader_context() {
        let blip = png_blip(b"png");
        let entry = delayed_fbse(&blip, 0, &[]);
        let mut bytes = vec![0x1f, 0, 0x01, 0xf0];
        bytes.extend_from_slice(&u32::try_from(entry.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(&entry);
        let strict = Limits {
            max_blip_bytes: 17,
            max_store_entries: 1,
            ..Default::default()
        };
        let store = Store::parse(&bytes, strict).expect("delayed metadata admitted");
        let id = Id::new(1).unwrap();
        assert!(matches!(
            get(&store, id, Some(&blip)),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::BlipBytes,
                maximum: 17,
                ..
            })
        ));
        let broad = Context::new().with_delay(Delay::new(&blip, Default::default()));
        assert!(matches!(
            store.resolve(id, broad),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::BlipBytes,
                maximum: 17,
                ..
            })
        ));
        let exact = Limits {
            max_blip_bytes: 20,
            ..strict
        };
        let store = Store::parse(&bytes, exact).expect("exact limit store");
        let resolved = store
            .resolve(id, broad)
            .expect("exact limit")
            .expect("image");
        assert_eq!(resolved.data(), b"png");
        assert_eq!(resolved.data().as_ptr(), blip[25..].as_ptr());
        let narrow_context = Context::new().with_delay(Delay::new(&blip, strict));
        assert!(matches!(
            store.resolve(id, narrow_context),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::BlipBytes,
                maximum: 17,
                ..
            })
        ));
    }

    #[test]
    fn delay_accepts_a_headerless_record_sequence() {
        let first = png_blip(b"one");
        let second = png_blip(b"two");
        let mut pictures = first;
        pictures.extend_from_slice(&second);
        let files = delay(&pictures, Limits::default()).expect("valid delay stream");
        assert_eq!(files.len(), 2);
        assert_eq!(files[0].data().expect("first image"), b"one");
        assert_eq!(files[1].data().expect("second image"), b"two");
    }

    #[test]
    fn recursive_scan_caps_the_semantic_collection() {
        let limits = Limits {
            max_files: 0x0fff,
            ..Limits::default()
        };
        let data = png_blip(b"x").repeat(usize::try_from(limits.max_files).unwrap() + 1);
        assert!(matches!(
            scan(&data, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Files,
                observed: 4096,
                maximum: 4095,
                ..
            })
        ));
    }

    #[test]
    fn recursive_scan_maps_raw_record_limit_to_image_kind() {
        let data = png_blip(b"one");
        let body_len = u32::try_from(data.len() - 8).expect("body length fits");
        let limits = Limits {
            records: crate::Limits {
                max_record_bytes: body_len - 1,
                ..crate::Limits::default()
            },
            ..Limits::default()
        };
        assert!(matches!(
            scan(&data, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::RecordBytes,
                observed: 20,
                maximum: 19,
            })
        ));
    }

    #[test]
    fn container_scan_retains_the_admitted_root_policy() {
        let first = png_blip(b"one");
        let second = png_blip(b"two");
        let mut body = first;
        body.extend_from_slice(&second);
        let mut bytes = vec![0x0F, 0x00, 0x04, 0xF0];
        bytes.extend_from_slice(&u32::try_from(body.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(&body);
        let origin = crate::Limits {
            max_records: 1,
            ..crate::Limits::default()
        };
        let record = Record::parse(&bytes, 0, origin).unwrap().0;
        let container = Container::try_new(record).unwrap();
        assert!(matches!(
            super::container(&container, None, Limits::default()),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Records,
                maximum: 1,
                ..
            })
        ));
    }

    #[test]
    fn store_caps_top_level_record_traversal() {
        let mut data = png_blip(b"one");
        data.extend_from_slice(&png_blip(b"two"));
        let limits = Limits {
            records: crate::Limits {
                max_records: 1,
                ..crate::Limits::default()
            },
            ..Limits::default()
        };

        assert!(matches!(
            store(&data, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Records,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn container_scan_maps_raw_record_limit_to_image_kind() {
        let child = png_blip(b"one");
        let data = officeart_container(0xf002, &child);
        let root = Record::parse(&data, 0, crate::Limits::default())
            .expect("valid drawing container")
            .0;
        let container = Container::try_new(root).expect("container");
        let body_len = u32::try_from(child.len() - 8).expect("body length fits");
        let limits = Limits {
            records: crate::Limits {
                max_record_bytes: body_len - 1,
                ..crate::Limits::default()
            },
            ..Limits::default()
        };
        assert!(matches!(
            super::container(&container, None, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::RecordBytes,
                observed: 20,
                maximum: 19,
            })
        ));
    }

    #[test]
    fn delay_maps_raw_record_limit_to_image_kind() {
        let data = png_blip(b"one");
        let body_len = u32::try_from(data.len() - 8).expect("body length fits");
        let limits = Limits {
            records: crate::Limits {
                max_record_bytes: body_len - 1,
                ..crate::Limits::default()
            },
            ..Limits::default()
        };
        assert!(matches!(
            delay(&data, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::RecordBytes,
                observed: 20,
                maximum: 19,
            })
        ));
    }

    fn officeart_container(kind: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = vec![0x0f, 0, 0, 0];
        bytes[2..4].copy_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(
            &u32::try_from(body.len())
                .expect("container body length fits")
                .to_le_bytes(),
        );
        bytes.extend_from_slice(body);
        bytes
    }

    #[test]
    fn store_maps_nested_record_limit_to_image_kind() {
        let mut body = png_blip(b"one");
        body.extend_from_slice(&png_blip(b"two"));
        let data = officeart_container(0xf002, &body);
        let limits = Limits {
            records: crate::Limits {
                max_records: 1,
                ..crate::Limits::default()
            },
            ..Limits::default()
        };
        assert!(matches!(
            store(&data, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Records,
                maximum: 1,
                ..
            })
        ));
    }

    #[test]
    fn store_maps_top_level_record_limit_to_image_kind() {
        let data = png_blip(b"one");
        let body_len = u32::try_from(data.len() - 8).expect("body length fits");
        let limits = Limits {
            records: crate::Limits {
                max_record_bytes: body_len - 1,
                ..crate::Limits::default()
            },
            ..Limits::default()
        };
        assert!(matches!(
            store(&data, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::RecordBytes,
                observed: 20,
                maximum: 19,
            })
        ));
    }

    #[test]
    fn store_maps_nested_depth_limit_to_image_kind() {
        let nested = officeart_container(0xf002, &[]);
        let data = officeart_container(0xf002, &nested);
        let limits = Limits {
            records: crate::Limits {
                max_depth: 0,
                ..crate::Limits::default()
            },
            ..Limits::default()
        };
        assert!(matches!(
            store(&data, limits),
            Err(Error::ImageLimitExceeded {
                limit: ImageLimitKind::Depth,
                maximum: 0,
                ..
            })
        ));
    }
}
