//! Bounded raw package-entry ingress.
//!
//! This module owns the physical ZIP envelope used by mutable iWork package
//! snapshots. It returns ordered, owned entries and deliberately does not
//! validate format-specific paths or decode IWA messages; those policies stay
//! with the facade and the neutral component catalog respectively.

use std::collections::HashSet;

use crate::zip::ZipArchive;
use crate::{Error, Limits, Result};

/// One ordered, uncompressed physical package member.
#[derive(Debug)]
pub struct Entry {
    name: Box<str>,
    data: Vec<u8>,
}

impl Entry {
    fn new(name: &str, data: Vec<u8>) -> Self {
        Self {
            name: name.into(),
            data,
        }
    }

    /// Borrow the physical member name in source order.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the uncompressed member payload.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Consume the entry without cloning its name or payload.
    #[must_use]
    pub fn into_parts(self) -> (String, Vec<u8>) {
        (self.name.into(), self.data)
    }
}

/// Ordered raw entries extracted from one physical iWork ZIP input.
#[derive(Debug)]
pub struct Catalog {
    entries: Vec<Entry>,
}

impl Catalog {
    /// Parse a package with the default physical limits.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a package under caller-selected physical limits.
    ///
    /// A legacy package containing `.../Index.zip` is flattened into the
    /// modern entry order used by the mutable facade: nested IWA members are
    /// emitted first, followed by outer entries with the legacy prefix
    /// removed. The nested archive must contain only IWA members.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let archive = ZipArchive::new_with_limits(bytes, limits)?;
        if crate::zip::is_encrypted(&archive) {
            return Err(Error::Encrypted);
        }

        let has_direct_iwa = archive.file_names().any(crate::zip::is_iwa_name);
        let index_name = crate::zip::nested_index_name(&archive)?;
        if has_direct_iwa && index_name.is_some() {
            return Err(Error::InvalidBundle(
                "iWork package mixes direct IWA members with a legacy Index.zip".to_owned(),
            ));
        }
        if has_direct_iwa {
            return collect_flat(&archive);
        }

        let Some(index_name) = index_name else {
            return collect_flat(&archive);
        };
        collect_legacy(&archive, &index_name, limits)
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
}

impl IntoIterator for Catalog {
    type Item = Entry;
    type IntoIter = std::vec::IntoIter<Entry>;

    fn into_iter(self) -> Self::IntoIter {
        self.entries.into_iter()
    }
}

fn collect_flat(archive: &ZipArchive<'_>) -> Result<Catalog> {
    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for name in archive.file_names().filter(|name| !name.ends_with('/')) {
        push_entry(archive, name, name, &mut entries, &mut seen)?;
    }
    Ok(Catalog { entries })
}

fn collect_legacy(archive: &ZipArchive<'_>, index_name: &str, limits: Limits) -> Result<Catalog> {
    let prefix = index_name.strip_suffix("Index.zip").ok_or_else(|| {
        Error::InvalidBundle(format!("invalid legacy package index name: {index_name}"))
    })?;
    let index_data = archive.read(index_name)?;
    let index_size = u64::try_from(index_data.len()).map_err(|_| {
        Error::InvalidBundle("legacy iWork Index.zip length does not fit u64".to_owned())
    })?;
    limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index = ZipArchive::new_with_limits(&index_data, limits)
        .map_err(|error| Error::InvalidBundle(format!("legacy package index: {error}")))?;

    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for name in index.file_names().filter(|name| !name.ends_with('/')) {
        if !crate::zip::is_iwa_name(name) {
            return Err(Error::InvalidBundle(format!(
                "legacy package index contains a non-IWA member: {name}"
            )));
        }
        push_entry(&index, name, name, &mut entries, &mut seen)?;
    }
    if entries.is_empty() {
        return Err(Error::InvalidBundle(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }

    for name in archive
        .file_names()
        .filter(|name| *name != index_name && !name.ends_with('/'))
    {
        let normalized = name.strip_prefix(prefix).unwrap_or(name);
        push_entry(archive, name, normalized, &mut entries, &mut seen)?;
    }
    Ok(Catalog { entries })
}

fn push_entry(
    archive: &ZipArchive<'_>,
    source_name: &str,
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
    let data = archive.read(source_name)?;
    entries.push(Entry::new(normalized_name, data));
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
