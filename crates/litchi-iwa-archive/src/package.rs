//! Bounded raw package-entry ingress.
//!
//! This module owns the physical ZIP envelope used by mutable iWork package
//! snapshots. It returns ordered, owned entries and deliberately does not
//! validate format-specific paths or decode IWA messages; those policies stay
//! with the facade and the neutral component catalog respectively.

use std::collections::HashSet;
use std::io::Write;
use std::ops::Range;
use std::sync::Arc;

use crate::zip::{PhysicalEntry, PhysicalHeader, ZipArchive};
use crate::{Error, Limits, Result};
use soapberry_zip::office::StreamingArchiveWriter;

/// A raw ZIP entry record retained for an exact preserve-mode write.
///
/// The local record includes the local header, compressed bytes, data
/// descriptor, and any bytes before the next local record or central
/// directory. The central record is retained separately because ZIP stores it
/// in a different part of the archive. Neither view is normalized.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawEntryRecord {
    source: Arc<[u8]>,
    local_record: Range<usize>,
    compressed_data: Range<usize>,
    central_directory_record: Range<usize>,
}

impl RawEntryRecord {
    fn new(source: Arc<[u8]>, entry: &PhysicalEntry) -> Self {
        Self {
            source,
            local_record: entry.local_record(),
            compressed_data: entry.compressed_data_range(),
            central_directory_record: entry.central_record(),
        }
    }

    /// Borrow the exact local record bytes.
    #[must_use]
    pub fn local_record(&self) -> &[u8] {
        &self.source[self.local_record.clone()]
    }

    /// Borrow the exact compressed data bytes.
    #[must_use]
    pub fn compressed_data(&self) -> &[u8] {
        &self.source[self.compressed_data.clone()]
    }

    /// Borrow the exact central-directory record bytes.
    #[must_use]
    pub fn central_directory_record(&self) -> &[u8] {
        &self.source[self.central_directory_record.clone()]
    }
}

/// The raw timestamp fields from one ZIP header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DosDateTime {
    time: u16,
    date: u16,
}

impl DosDateTime {
    /// Return the original packed DOS time field.
    #[must_use]
    pub const fn time(self) -> u16 {
        self.time
    }

    /// Return the original packed DOS date field.
    #[must_use]
    pub const fn date(self) -> u16 {
        self.date
    }
}

/// Physical metadata from one local or central ZIP header.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeaderMetadata {
    version_needed: u16,
    flags: u16,
    compression_method: u16,
    last_modified: DosDateTime,
    name: Box<[u8]>,
    extra: Box<[u8]>,
    comment: Box<[u8]>,
}

impl HeaderMetadata {
    /// Return the ZIP version needed to extract this header's member.
    #[must_use]
    pub const fn version_needed(&self) -> u16 {
        self.version_needed
    }

    /// Return the original general-purpose bit flags.
    #[must_use]
    pub const fn flags(&self) -> u16 {
        self.flags
    }

    /// Return the original numeric ZIP compression method.
    #[must_use]
    pub const fn compression_method(&self) -> u16 {
        self.compression_method
    }

    /// Return the original packed DOS modification timestamp.
    #[must_use]
    pub const fn last_modified(&self) -> DosDateTime {
        self.last_modified
    }

    /// Borrow the exact header filename bytes.
    #[must_use]
    pub fn name(&self) -> &[u8] {
        &self.name
    }

    /// Borrow the exact header extra-field bytes.
    #[must_use]
    pub fn extra(&self) -> &[u8] {
        &self.extra
    }

    /// Borrow the exact header comment bytes.
    #[must_use]
    pub fn comment(&self) -> &[u8] {
        &self.comment
    }
}

/// All physical metadata needed to describe one ZIP member.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EntryMetadata {
    local: HeaderMetadata,
    central: HeaderMetadata,
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
}

impl EntryMetadata {
    fn new(entry: &PhysicalEntry) -> Self {
        Self {
            local: header_metadata(entry.local_header()),
            central: header_metadata(entry.central_header()),
            compressed_size: entry.compressed_size(),
            uncompressed_size: entry.uncompressed_size(),
            crc32: entry.crc32(),
        }
    }

    /// Return metadata from the local file header.
    #[must_use]
    pub const fn local(&self) -> &HeaderMetadata {
        &self.local
    }

    /// Return metadata from the central-directory header.
    #[must_use]
    pub const fn central(&self) -> &HeaderMetadata {
        &self.central
    }

    /// Return the central directory's declared compressed size.
    #[must_use]
    pub const fn compressed_size(&self) -> u64 {
        self.compressed_size
    }

    /// Return the central directory's declared uncompressed size.
    #[must_use]
    pub const fn uncompressed_size(&self) -> u64 {
        self.uncompressed_size
    }

    /// Return the central directory's declared CRC-32.
    #[must_use]
    pub const fn crc32(&self) -> u32 {
        self.crc32
    }
}

/// The safe distinction between decoded content and an unsupported physical
/// member.
#[derive(Debug, Clone, Copy)]
pub enum EntryPayload<'a> {
    /// The member was decoded by the bounded ZIP reader.
    Decoded(&'a [u8]),
    /// The member's compression method is unsupported; use its raw record.
    Opaque(&'a RawEntryRecord),
}

/// One ordered package member with its physical ZIP provenance.
#[derive(Debug)]
pub struct Entry {
    name: Box<str>,
    data: Vec<u8>,
    raw_name: Box<[u8]>,
    metadata: EntryMetadata,
    raw_record: RawEntryRecord,
    opaque: bool,
}

impl Entry {
    fn new(
        name: &str,
        data: Vec<u8>,
        raw_name: Box<[u8]>,
        metadata: EntryMetadata,
        raw_record: RawEntryRecord,
        opaque: bool,
    ) -> Self {
        Self {
            name: name.into(),
            data,
            raw_name,
            metadata,
            raw_record,
            opaque,
        }
    }

    /// Borrow the physical member name in source order.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the decoded member payload.
    ///
    /// For an opaque entry this returns the raw compressed byte stream for
    /// compatibility with the original payload accessor. Use the structured
    /// payload or opaque flag before interpreting it.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Borrow the exact central-directory filename bytes.
    #[must_use]
    pub fn raw_name(&self) -> &[u8] {
        &self.raw_name
    }

    /// Borrow the preserved physical metadata.
    #[must_use]
    pub const fn metadata(&self) -> &EntryMetadata {
        &self.metadata
    }

    /// Borrow the preserved raw ZIP records.
    #[must_use]
    pub const fn raw_record(&self) -> &RawEntryRecord {
        &self.raw_record
    }

    /// Return whether the compression method was not decoded by this crate.
    #[must_use]
    pub const fn is_opaque(&self) -> bool {
        self.opaque
    }

    /// Return decoded content or the opaque raw-record provenance.
    #[must_use]
    pub fn payload(&self) -> EntryPayload<'_> {
        if self.opaque {
            EntryPayload::Opaque(&self.raw_record)
        } else {
            EntryPayload::Decoded(&self.data)
        }
    }

    /// Consume the entry without cloning its name or payload.
    ///
    /// For an opaque entry, the returned bytes are the raw compressed stream;
    /// callers must retain the raw record when they need preserve-mode
    /// serialization.
    #[must_use]
    pub fn into_parts(self) -> (String, Vec<u8>) {
        (self.name.into(), self.data)
    }
}

/// Ordered raw entries extracted from one physical iWork ZIP input.
#[derive(Debug)]
pub struct Catalog {
    entries: Vec<Entry>,
    source: Arc<[u8]>,
}

impl Catalog {
    /// Parse a package with the default physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope or any configured physical
    /// limit is invalid.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse an immutable, already-owned package source with the default
    /// physical limits.
    ///
    /// The catalog retains this exact [`Arc`] allocation for preserve-mode
    /// writes and does not copy the source bytes. The shared byte slice is
    /// immutable, so entries and raw ZIP records can safely borrow from it
    /// for the lifetime of the catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope or any configured physical
    /// limit is invalid.
    pub fn from_shared_bytes(source: Arc<[u8]>) -> Result<Self> {
        Self::from_shared_bytes_with_limits(source, Limits::default())
    }

    /// Parse a package under caller-selected physical limits.
    ///
    /// A legacy package containing `.../Index.zip` is flattened into the
    /// modern entry order used by the mutable facade: nested IWA members are
    /// emitted first, followed by outer entries with the legacy prefix
    /// removed. The nested archive must contain only IWA members.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope, nested index, duplicate entry,
    /// or configured physical limit is invalid.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let checked_limits = limits.validate()?;
        let source: Arc<[u8]> = bytes.to_vec().into();
        Self::from_source_with_checked_limits(source, checked_limits)
    }

    /// Parse an immutable, already-owned package source under caller-selected
    /// physical limits.
    ///
    /// The catalog retains the supplied [`Arc`] allocation for preserve-mode
    /// writes and does not copy the source bytes. A legacy package containing
    /// `.../Index.zip` is flattened into the modern entry order used by the
    /// mutable facade: nested IWA members are emitted first, followed by outer
    /// entries with the legacy prefix removed. The nested archive must contain
    /// only IWA members.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP envelope, nested index, duplicate entry,
    /// or configured physical limit is invalid.
    pub fn from_shared_bytes_with_limits(source: Arc<[u8]>, limits: Limits) -> Result<Self> {
        let checked_limits = limits.validate()?;
        Self::from_source_with_checked_limits(source, checked_limits)
    }

    fn from_source_with_checked_limits(source: Arc<[u8]>, checked_limits: Limits) -> Result<Self> {
        let archive = ZipArchive::new_with_limits(source.as_ref(), checked_limits)?;
        if crate::zip::is_encrypted(&archive) {
            return Err(Error::Encrypted);
        }

        let has_direct_iwa = archive.file_names().any(crate::zip::is_iwa_name);
        let nested_name = crate::zip::nested_index_name(&archive)?;
        if has_direct_iwa && nested_name.is_some() {
            return Err(Error::InvalidBundle(
                "iWork package mixes direct IWA members with a legacy Index.zip".to_owned(),
            ));
        }
        let entries = if has_direct_iwa {
            collect_flat(&archive, &source)?
        } else if let Some(index_name) = nested_name {
            collect_legacy(&archive, &index_name, checked_limits, &source)?
        } else {
            collect_flat(&archive, &source)?
        };
        Ok(Catalog { entries, source })
    }

    /// Return the number of extracted entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether no entries were extracted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Borrow entries in their preserved source order.
    pub fn iter(&self) -> impl Iterator<Item = &Entry> {
        self.entries.iter()
    }

    /// Write the original ZIP bytes without rebuilding or normalizing any
    /// member, central record, archive comment, or opaque entry.
    ///
    /// This is the preserve-mode no-op path. Catalog has no mutating
    /// operations in this bounded slice, so the source remains authoritative.
    ///
    /// # Errors
    ///
    /// Returns an error when the caller-provided sink rejects the source bytes.
    pub fn write_to<W: Write>(&self, mut sink: W) -> Result<()> {
        sink.write_all(&self.source)?;
        Ok(())
    }

    /// Return an exact byte-for-byte copy of the source ZIP.
    ///
    /// # Errors
    ///
    /// Returns an error if the source copy cannot be allocated.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.source.len())
            .map_err(|_error| Error::Allocation {
                resource: "catalog source bytes",
                amount: self.source.len(),
            })?;
        bytes.extend_from_slice(&self.source);
        Ok(bytes)
    }
}

impl IntoIterator for Catalog {
    type Item = Entry;
    type IntoIter = std::vec::IntoIter<Entry>;

    fn into_iter(self) -> Self::IntoIter {
        self.entries.into_iter()
    }
}

fn header_metadata(header: &PhysicalHeader) -> HeaderMetadata {
    HeaderMetadata {
        version_needed: header.version_needed,
        flags: header.flags,
        compression_method: header.compression_method,
        last_modified: DosDateTime {
            time: header.last_mod_time,
            date: header.last_mod_date,
        },
        name: header.name.clone(),
        extra: header.extra.clone(),
        comment: header.comment.clone(),
    }
}

/// Write ordered, uncompressed package members to a physical iWork ZIP.
///
/// The input iterator must be cloneable because the complete member budget is
/// checked before the first byte reaches `sink`. This keeps a rejected
/// package transaction from leaving a partially written physical archive.
/// This is a new logical-package writer: it deliberately emits Store entries
/// and has no physical metadata input. Use `Catalog::write_to` for an untouched
/// preserve-mode round trip.
/// ZIP grammar and implementation details remain private to this crate.
///
/// # Errors
///
/// Returns an error when the entry budget is exceeded or the physical ZIP
/// writer rejects the sink or an entry.
pub fn write_to<'a, I, W>(entries: I, sink: W, limits: Limits) -> Result<()>
where
    I: IntoIterator<Item = (&'a str, &'a [u8])> + Clone,
    W: Write,
{
    let checked_limits = limits.validate()?;
    validate_output(entries.clone(), checked_limits)?;

    let mut writer = StreamingArchiveWriter::with_writer(sink);
    for (name, data) in entries {
        writer.write_stored(name, data)?;
    }
    writer.finish()?;
    Ok(())
}

/// Encode ordered, uncompressed package members as a physical iWork ZIP.
///
/// This is a new logical-package writer and intentionally does not preserve
/// metadata from an Entry. Use `Catalog::to_bytes` for an untouched preserve-mode
/// round trip.
///
/// # Errors
///
/// Returns an error when the entry budget is exceeded or the physical ZIP
/// writer rejects an entry.
pub fn to_bytes<'a, I>(entries: I, limits: Limits) -> Result<Vec<u8>>
where
    I: IntoIterator<Item = (&'a str, &'a [u8])> + Clone,
{
    let mut bytes = Vec::new();
    write_to(entries, &mut bytes, limits)?;
    Ok(bytes)
}

fn validate_output<'a, I>(entries: I, limits: Limits) -> Result<()>
where
    I: IntoIterator<Item = (&'a str, &'a [u8])>,
{
    let maximum_entries = u64::try_from(limits.max_entries()).map_err(|error| {
        Error::InvalidBundle(format!(
            "package output entry limit does not fit u64: {error}"
        ))
    })?;
    let mut count = 0usize;
    let mut total = 0u64;
    for (_name, data) in entries {
        count = count.checked_add(1).ok_or_else(|| {
            Error::InvalidBundle("package output entry count overflow".to_owned())
        })?;
        if count > limits.max_entries() {
            let observed = u64::try_from(count).map_err(|error| {
                Error::InvalidBundle(format!(
                    "package output entry count does not fit u64: {error}"
                ))
            })?;
            return Err(Error::Limit {
                kind: crate::LimitKind::Entries,
                observed,
                maximum: maximum_entries,
            });
        }
        let size = u64::try_from(data.len()).map_err(|error| {
            Error::InvalidBundle(format!(
                "package output member length does not fit u64: {error}"
            ))
        })?;
        if size > limits.max_entry_bytes() {
            return Err(Error::Limit {
                kind: crate::LimitKind::EntryBytes,
                observed: size,
                maximum: limits.max_entry_bytes(),
            });
        }
        total = total
            .checked_add(size)
            .ok_or_else(|| Error::InvalidBundle("package output total size overflow".to_owned()))?;
        if total > limits.max_total_bytes() {
            return Err(Error::Limit {
                kind: crate::LimitKind::TotalBytes,
                observed: total,
                maximum: limits.max_total_bytes(),
            });
        }
    }
    Ok(())
}

fn collect_flat(archive: &ZipArchive<'_>, source: &Arc<[u8]>) -> Result<Vec<Entry>> {
    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        push_entry(
            archive,
            source.clone(),
            entry,
            entry.name(),
            &mut entries,
            &mut seen,
        )?;
    }
    Ok(entries)
}

fn collect_legacy(
    archive: &ZipArchive<'_>,
    index_name: &str,
    limits: Limits,
    source: &Arc<[u8]>,
) -> Result<Vec<Entry>> {
    let prefix = index_name.strip_suffix("Index.zip").ok_or_else(|| {
        Error::InvalidBundle(format!("invalid legacy package index name: {index_name}"))
    })?;
    let index_data = archive.read(index_name)?;
    let index_size = u64::try_from(index_data.len()).map_err(|error| {
        Error::InvalidBundle(format!(
            "legacy iWork Index.zip length does not fit u64: {error}"
        ))
    })?;
    limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index_source: Arc<[u8]> = index_data.into();
    let index = ZipArchive::new_with_limits(index_source.as_ref(), limits)
        .map_err(|error| Error::InvalidBundle(format!("legacy package index: {error}")))?;

    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for entry in index
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        if !crate::zip::is_iwa_name(entry.name()) {
            return Err(Error::InvalidBundle(format!(
                "legacy package index contains a non-IWA member: {}",
                entry.name()
            )));
        }
        push_entry(
            &index,
            index_source.clone(),
            entry,
            entry.name(),
            &mut entries,
            &mut seen,
        )?;
    }
    if entries.is_empty() {
        return Err(Error::InvalidBundle(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }

    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        if entry.name() == index_name {
            continue;
        }
        let normalized = entry.name().strip_prefix(prefix).unwrap_or(entry.name());
        push_entry(
            archive,
            source.clone(),
            entry,
            normalized,
            &mut entries,
            &mut seen,
        )?;
    }
    Ok(entries)
}

fn push_entry(
    archive: &ZipArchive<'_>,
    source: Arc<[u8]>,
    physical: &PhysicalEntry,
    normalized_name: &str,
    entries: &mut Vec<Entry>,
    seen: &mut HashSet<String>,
) -> Result<()> {
    seen.try_reserve(1).map_err(|_error| Error::Allocation {
        resource: "package entry names",
        amount: 1,
    })?;
    if !seen.insert(normalized_name.to_owned()) {
        return Err(Error::InvalidBundle(format!(
            "duplicate package entry is ambiguous: {normalized_name}"
        )));
    }
    entries.try_reserve(1).map_err(|_error| Error::Allocation {
        resource: "package entries",
        amount: 1,
    })?;
    let raw_record = RawEntryRecord::new(source, physical);
    let metadata = EntryMetadata::new(physical);
    let data = if physical.is_supported() {
        archive.read_entry(physical)?
    } else {
        physical.compressed_data(archive.source()).to_vec()
    };
    entries.push(Entry::new(
        normalized_name,
        data,
        physical.raw_name().to_vec().into_boxed_slice(),
        metadata,
        raw_record,
        !physical.is_supported(),
    ));
    Ok(())
}

#[cfg(test)]
mod tests {
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    fn zip(entries: &[(&str, &[u8])]) -> Result<Vec<u8>> {
        let mut writer = StreamingArchiveWriter::new();
        for (name, data) in entries {
            writer.write_stored(name, data)?;
        }
        Ok(writer.finish_to_bytes()?)
    }

    fn push_u16(bytes: &mut Vec<u8>, value: u16) {
        bytes.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(bytes: &mut Vec<u8>, value: u32) {
        bytes.extend_from_slice(&value.to_le_bytes());
    }

    fn checked_u16(value: usize) -> u16 {
        match u16::try_from(value) {
            Ok(converted) => converted,
            Err(error) => panic!("test ZIP field does not fit u16: {error}"),
        }
    }

    fn checked_u32(value: usize) -> u32 {
        match u32::try_from(value) {
            Ok(converted) => converted,
            Err(error) => panic!("test ZIP field does not fit u32: {error}"),
        }
    }

    fn physical_zip(method: u16) -> (Vec<u8>, usize, usize) {
        let name = b"Opaque/entry.bin";
        let data = b"opaque payload";
        let local_extra = b"\xaa\xbb\x03\0xyz";
        let central_extra = b"\xcc\xdd\x02\0\xfe\xed";
        let file_comment = b"entry-comment\0\xff";
        let archive_comment = b"archive-comment\0\xfe";
        let crc32 = soapberry_zip::crc32(data);
        let mut bytes = Vec::new();

        push_u32(&mut bytes, 0x0403_4b50);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 0x0800);
        push_u16(&mut bytes, method);
        push_u16(&mut bytes, 0x1234);
        push_u16(&mut bytes, 0x5678);
        push_u32(&mut bytes, crc32);
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u16(&mut bytes, checked_u16(name.len()));
        push_u16(&mut bytes, checked_u16(local_extra.len()));
        bytes.extend_from_slice(name);
        bytes.extend_from_slice(local_extra);
        bytes.extend_from_slice(data);

        let central_offset = bytes.len();
        push_u32(&mut bytes, 0x0201_4b50);
        push_u16(&mut bytes, 0x0314);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 0x0800);
        push_u16(&mut bytes, method);
        push_u16(&mut bytes, 0x9abc);
        push_u16(&mut bytes, 0xdef0);
        push_u32(&mut bytes, crc32);
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u32(&mut bytes, checked_u32(data.len()));
        push_u16(&mut bytes, checked_u16(name.len()));
        push_u16(&mut bytes, checked_u16(central_extra.len()));
        push_u16(&mut bytes, checked_u16(file_comment.len()));
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u32(&mut bytes, 0);
        push_u32(&mut bytes, 0);
        bytes.extend_from_slice(name);
        bytes.extend_from_slice(central_extra);
        bytes.extend_from_slice(file_comment);
        let central_end = bytes.len();

        push_u32(&mut bytes, 0x0605_4b50);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 1);
        push_u16(&mut bytes, 1);
        push_u32(&mut bytes, checked_u32(central_end - central_offset));
        push_u32(&mut bytes, checked_u32(central_offset));
        push_u16(&mut bytes, checked_u16(archive_comment.len()));
        bytes.extend_from_slice(archive_comment);
        (bytes, central_offset, central_end)
    }

    #[test]
    fn preserves_flat_entry_order_and_payloads() -> Result<()> {
        let bytes = zip(&[("Metadata/a", b"a"), ("Index/Document.iwa", b"iwa")])?;
        let catalog = Catalog::from_bytes(&bytes)?;
        assert_eq!(
            catalog.iter().map(Entry::name).collect::<Vec<_>>(),
            ["Metadata/a", "Index/Document.iwa"]
        );
        assert_eq!(catalog.iter().next().map(Entry::data), Some(&b"a"[..]));
        Ok(())
    }

    #[test]
    fn writes_flat_entry_order_and_payloads_without_exposing_zip_types() -> Result<()> {
        let entries: [(&str, &[u8]); 2] = [("Metadata/a", b"a"), ("Index/Document.iwa", b"opaque")];
        let bytes = to_bytes(entries.iter().copied(), Limits::default())?;
        let catalog = Catalog::from_bytes(&bytes)?;
        assert_eq!(
            catalog
                .into_iter()
                .map(Entry::into_parts)
                .collect::<Vec<_>>(),
            [
                ("Metadata/a".to_owned(), b"a".to_vec()),
                ("Index/Document.iwa".to_owned(), b"opaque".to_vec()),
            ]
        );
        Ok(())
    }

    #[test]
    fn preserves_physical_metadata_and_exact_noop_bytes() -> Result<()> {
        let (bytes, central_offset, central_end) = physical_zip(0);
        let catalog = Catalog::from_bytes(&bytes)?;
        let entry = catalog.iter().next().ok_or_else(|| {
            Error::InvalidBundle("physical metadata test produced no entry".to_owned())
        })?;

        assert!(!entry.is_opaque());
        assert_eq!(entry.raw_name(), b"Opaque/entry.bin");
        assert_eq!(entry.data(), b"opaque payload");
        assert_eq!(entry.metadata().local().flags(), 0x0800);
        assert_eq!(entry.metadata().central().flags(), 0x0800);
        assert_eq!(entry.metadata().local().compression_method(), 0);
        assert_eq!(entry.metadata().central().compression_method(), 0);
        assert_eq!(entry.metadata().local().last_modified().time(), 0x1234);
        assert_eq!(entry.metadata().local().last_modified().date(), 0x5678);
        assert_eq!(entry.metadata().central().last_modified().time(), 0x9abc);
        assert_eq!(entry.metadata().central().last_modified().date(), 0xdef0);
        assert_eq!(entry.metadata().local().extra(), b"\xaa\xbb\x03\0xyz");
        assert_eq!(
            entry.metadata().central().extra(),
            b"\xcc\xdd\x02\0\xfe\xed"
        );
        assert_eq!(entry.metadata().central().comment(), b"entry-comment\0\xff");
        assert_eq!(entry.raw_record().local_record(), &bytes[..central_offset]);
        assert_eq!(
            entry.raw_record().central_directory_record(),
            &bytes[central_offset..central_end]
        );
        assert_eq!(entry.raw_record().compressed_data(), b"opaque payload");
        assert_eq!(catalog.to_bytes()?, bytes);
        let mut streamed = Vec::new();
        catalog.write_to(&mut streamed)?;
        assert_eq!(streamed, bytes);
        Ok(())
    }

    #[test]
    fn accepts_shared_source_without_copying_input() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(0);
        let source: Arc<[u8]> = bytes.into();
        let catalog = Catalog::from_shared_bytes(source.clone())?;

        assert!(Arc::ptr_eq(&source, &catalog.source));
        assert_eq!(catalog.len(), 1);
        Ok(())
    }

    #[test]
    fn shared_source_respects_input_limits() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(0);
        let source: Arc<[u8]> = bytes.into();
        let limits = Limits::new(1, 10, 100, 100, 100)?;

        assert!(matches!(
            Catalog::from_shared_bytes_with_limits(source, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::InputBytes,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn shared_source_preserves_opaque_zip_bytes_exactly() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(99);
        let source: Arc<[u8]> = bytes.into();
        let catalog = Catalog::from_shared_bytes(source.clone())?;
        let entry = catalog.iter().next().ok_or_else(|| {
            Error::InvalidBundle("shared source test produced no entry".to_owned())
        })?;

        assert!(entry.is_opaque());
        assert_eq!(entry.raw_record().compressed_data(), b"opaque payload");
        assert_eq!(catalog.to_bytes()?.as_slice(), source.as_ref());
        let mut streamed = Vec::new();
        catalog.write_to(&mut streamed)?;
        assert_eq!(streamed.as_slice(), source.as_ref());
        Ok(())
    }

    #[test]
    fn retains_unsupported_compression_as_opaque_raw_record() -> Result<()> {
        let (bytes, _central_offset, _central_end) = physical_zip(99);
        let catalog = Catalog::from_bytes(&bytes)?;
        let entry = catalog.iter().next().ok_or_else(|| {
            Error::InvalidBundle("opaque metadata test produced no entry".to_owned())
        })?;

        assert!(entry.is_opaque());
        assert_eq!(entry.metadata().central().compression_method(), 99);
        assert_eq!(entry.data(), b"opaque payload");
        assert!(matches!(
            entry.payload(),
            EntryPayload::Opaque(record) if record.compressed_data() == b"opaque payload"
        ));
        assert_eq!(catalog.to_bytes()?, bytes);
        Ok(())
    }

    #[test]
    fn rejects_truncated_local_metadata_before_materializing_payload() {
        let (mut bytes, _central_offset, _central_end) = physical_zip(0);
        bytes[28..30].copy_from_slice(&u16::MAX.to_le_bytes());
        assert!(matches!(
            Catalog::from_bytes(&bytes),
            Err(Error::InvalidBundle(message)) if message.contains("local file header")
        ));
    }

    #[test]
    fn rejects_output_limits_before_writing_any_bytes() -> Result<()> {
        let entries: [(&str, &[u8]); 2] = [("Data/a", b"a"), ("Data/b", b"b")];
        let limits = Limits::new(1024, 1, 1024, 1024, 1024)?;
        let mut sink = Vec::new();
        let result = write_to(entries.iter().copied(), &mut sink, limits);
        let error = result
            .err()
            .ok_or_else(|| Error::InvalidBundle("output unexpectedly succeeded".to_owned()))?;
        assert!(matches!(
            error,
            Error::Limit {
                kind: crate::LimitKind::Entries,
                ..
            }
        ));
        assert!(sink.is_empty());
        Ok(())
    }

    #[test]
    fn flattens_legacy_index_before_outer_entries() -> Result<()> {
        let index = zip(&[("Index/Document.iwa", b"iwa")])?;
        let bytes = zip(&[
            ("legacy.pages/Index.zip", &index),
            ("legacy.pages/Data/a", b"a"),
        ])?;
        let catalog = Catalog::from_bytes(&bytes)?;
        assert_eq!(
            catalog.iter().map(Entry::name).collect::<Vec<_>>(),
            ["Index/Document.iwa", "Data/a"]
        );
        Ok(())
    }

    #[test]
    fn rejects_legacy_non_iwa_members() -> Result<()> {
        let index = zip(&[("Index/Document.iwa", b"iwa"), ("Metadata/a", b"bad")])?;
        let bytes = zip(&[("legacy.pages/Index.zip", &index)])?;
        assert!(matches!(
            Catalog::from_bytes(&bytes),
            Err(Error::InvalidBundle(message)) if message.contains("non-IWA")
        ));
        Ok(())
    }

    #[test]
    fn rejects_mixed_direct_and_legacy_representations() -> Result<()> {
        let index = zip(&[("Index/Document.iwa", b"iwa")])?;
        let bytes = zip(&[
            ("legacy.pages/Index.zip", &index),
            ("Index/CalculationEngine.iwa", b"iwa"),
        ])?;
        assert!(matches!(
            Catalog::from_bytes(&bytes),
            Err(Error::InvalidBundle(message)) if message.contains("mixes direct IWA")
        ));
        Ok(())
    }
}
