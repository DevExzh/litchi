//! Mutable iWork ZIP package with entry-order preservation.

use std::collections::HashSet;
use std::fs::File;
use std::io::Write;
use std::path::{Component, Path};

use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use tempfile::NamedTempFile;

use crate::archive::Archive;
use crate::snappy::SnappyStream;
use crate::zip_utils::{is_encrypted_iwork_archive, nested_index_zip_name};
use crate::{Error, Result};

/// A mutable single-file Pages, Numbers, or Keynote package.
///
/// All ZIP members are retained as raw uncompressed bytes. IWA entries can be
/// parsed, updated transactionally, and written back while media, previews, and
/// metadata remain byte-for-byte unchanged.
#[derive(Debug, Clone, Default)]
pub struct IWorkPackage {
    entries: Vec<(String, Vec<u8>)>,
}

impl IWorkPackage {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::from_bytes(&std::fs::read(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let archive = ArchiveReader::new(bytes)
            .map_err(|error| Error::Bundle(format!("Failed to open iWork ZIP: {error}")))?;
        if is_encrypted_iwork_archive(&archive) {
            return Err(Error::InvalidFormat(
                "password-protected iWork documents are not supported".to_owned(),
            ));
        }
        if !archive.file_names().any(|name| name.ends_with(".iwa"))
            && let Some(index_name) = nested_index_zip_name(&archive)?
        {
            return Self::from_legacy_bundle(&archive, &index_name);
        }
        Self::from_flat_archive(&archive)
    }

    fn from_flat_archive(archive: &ArchiveReader<'_>) -> Result<Self> {
        let mut entries = Vec::new();
        let mut seen = HashSet::new();
        for name in archive.file_names() {
            validate_entry_name(name)?;
            if !seen.insert(name.to_owned()) {
                return Err(Error::Bundle(format!(
                    "Duplicate package entry is ambiguous: {name}"
                )));
            }
            let data = archive.read(name).map_err(|error| {
                Error::Bundle(format!("Failed to read package entry {name}: {error}"))
            })?;
            entries.push((name.to_owned(), data));
        }
        Ok(Self { entries })
    }

    /// Expand the pre-iWork '13 nested bundle representation into the modern,
    /// flat package representation used by the rest of the mutable API. The
    /// IWA members come first and all non-directory assets are retained with
    /// the legacy bundle prefix removed.
    fn from_legacy_bundle(archive: &ArchiveReader<'_>, index_name: &str) -> Result<Self> {
        let prefix = index_name.strip_suffix("Index.zip").ok_or_else(|| {
            Error::InvalidFormat(format!("invalid legacy package index name: {index_name}"))
        })?;
        let index_data = archive.read(index_name).map_err(|error| {
            Error::Bundle(format!(
                "Failed to read legacy package index {index_name}: {error}"
            ))
        })?;
        let index = ArchiveReader::new(&index_data).map_err(|error| {
            Error::Bundle(format!(
                "Failed to open legacy package index {index_name}: {error}"
            ))
        })?;

        let mut entries = Vec::new();
        let mut seen = HashSet::new();
        for name in index.file_names().filter(|name| !name.ends_with('/')) {
            validate_entry_name(name)?;
            if !name.ends_with(".iwa") {
                return Err(Error::InvalidFormat(format!(
                    "legacy package index contains a non-IWA member: {name}"
                )));
            }
            insert_unique_archive_entry(&index, name, &mut entries, &mut seen)?;
        }
        if entries.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "legacy package index {index_name} contains no IWA components"
            )));
        }

        for outer_name in archive
            .file_names()
            .filter(|name| *name != index_name && !name.ends_with('/'))
        {
            let name = outer_name.strip_prefix(prefix).unwrap_or(outer_name);
            validate_entry_name(name)?;
            if !seen.insert(name.to_owned()) {
                return Err(Error::InvalidFormat(format!(
                    "legacy package entries normalize to the same name: {name}"
                )));
            }
            let data = archive.read(outer_name).map_err(|error| {
                Error::Bundle(format!(
                    "Failed to read legacy package entry {outer_name}: {error}"
                ))
            })?;
            entries.push((name.to_owned(), data));
        }
        Ok(Self { entries })
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn entry_names(&self) -> impl Iterator<Item = &str> {
        self.entries.iter().map(|(name, _)| name.as_str())
    }

    /// Enumerate package members that contain IWA object archives.
    ///
    /// Some legacy packages contain an `OperationStorage.iwa` member whose
    /// `bvxn` payload is a separate operation-log format. It is intentionally
    /// retained as a raw entry but excluded from object-archive scans.
    pub fn iwa_entry_names(&self) -> impl Iterator<Item = &str> {
        self.entries
            .iter()
            .filter(|(name, data)| {
                name.ends_with(".iwa") && !is_legacy_operation_storage(name, data)
            })
            .map(|(name, _)| name.as_str())
    }

    /// Locate the package's calculation-engine component without allocating.
    ///
    /// Pages and Numbers may add numeric suffixes when they save a package,
    /// for example `Index/CalculationEngine-174.iwa`. Multiple matching
    /// components are rejected because choosing one would make formula edits
    /// ambiguous.
    pub fn calculation_engine_entry_name(&self) -> Result<Option<&str>> {
        let mut entries = self
            .iwa_entry_names()
            .filter(|name| is_calculation_engine_entry_name(name));
        let Some(entry) = entries.next() else {
            return Ok(None);
        };
        if entries.next().is_some() {
            return Err(Error::InvalidFormat(
                "iWork package contains multiple CalculationEngine components".to_owned(),
            ));
        }
        Ok(Some(entry))
    }

    pub fn contains_entry(&self, name: &str) -> bool {
        self.entry_position(normalize_entry_name(name)).is_some()
    }

    pub fn entry(&self, name: &str) -> Option<&[u8]> {
        let position = self.entry_position(normalize_entry_name(name))?;
        Some(self.entries[position].1.as_slice())
    }

    pub fn entry_mut(&mut self, name: &str) -> Option<&mut Vec<u8>> {
        let position = self.entry_position(normalize_entry_name(name))?;
        Some(&mut self.entries[position].1)
    }

    /// Create or replace a package member.
    pub fn insert_entry(
        &mut self,
        name: impl Into<String>,
        data: Vec<u8>,
    ) -> Result<Option<Vec<u8>>> {
        let supplied_name = name.into();
        let name = normalize_entry_name(&supplied_name).to_string();
        validate_entry_name(&name)?;
        if let Some(position) = self.entry_position(&name) {
            return Ok(Some(std::mem::replace(&mut self.entries[position].1, data)));
        }
        self.insert_new_entry(name, data);
        Ok(None)
    }

    /// Delete a package member.
    pub fn remove_entry(&mut self, name: &str) -> Option<Vec<u8>> {
        let position = self.entry_position(normalize_entry_name(name))?;
        Some(self.entries.remove(position).1)
    }

    /// Parse a compressed `.iwa` package member.
    pub fn archive(&self, name: &str) -> Result<Archive> {
        let normalized = normalize_entry_name(name);
        if !normalized.ends_with(".iwa") {
            return Err(Error::Bundle(format!(
                "Package entry {normalized} is not an IWA component"
            )));
        }
        let compressed = self
            .entry(normalized)
            .ok_or_else(|| Error::Bundle(format!("IWA package entry not found: {normalized}")))?;
        if is_legacy_operation_storage(normalized, compressed) {
            return Err(Error::InvalidFormat(format!(
                "package entry {normalized} is a legacy operation log, not an IWA object archive"
            )));
        }
        let stream = SnappyStream::decompress(&mut std::io::Cursor::new(compressed))?;
        Archive::parse(stream.data())
    }

    /// Serialize and replace a parsed `.iwa` package member.
    pub fn replace_archive(&mut self, name: &str, archive: &Archive) -> Result<Option<Vec<u8>>> {
        let normalized = normalize_entry_name(name).to_string();
        validate_entry_name(&normalized)?;
        if !normalized.ends_with(".iwa") {
            return Err(Error::Bundle(format!(
                "Package entry {normalized} is not an IWA component"
            )));
        }
        let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
        self.insert_entry(normalized, compressed)
    }

    /// Serialize and insert a new IWA component before an existing package member.
    pub(crate) fn insert_archive_before(
        &mut self,
        name: &str,
        archive: &Archive,
        before: &str,
    ) -> Result<()> {
        let normalized = normalize_entry_name(name).to_string();
        let before = normalize_entry_name(before);
        validate_entry_name(&normalized)?;
        if !normalized.ends_with(".iwa") {
            return Err(Error::Bundle(format!(
                "Package entry {normalized} is not an IWA component"
            )));
        }
        if self.contains_entry(&normalized) {
            return Err(Error::InvalidFormat(format!(
                "IWA package entry already exists: {normalized}"
            )));
        }
        let position = self
            .entry_position(before)
            .ok_or_else(|| Error::Bundle(format!("IWA insertion anchor not found: {before}")))?;
        let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
        self.entries.insert(position, (normalized, compressed));
        Ok(())
    }

    /// Parse, mutate, validate, and replace an IWA component as one operation.
    /// If the callback or serialization fails, the original package is unchanged.
    pub fn update_archive<F>(&mut self, name: &str, update: F) -> Result<()>
    where
        F: FnOnce(&mut Archive) -> Result<()>,
    {
        let mut archive = self.archive(name)?;
        update(&mut archive)?;
        archive.validate()?;
        self.replace_archive(name, &archive)?;
        Ok(())
    }

    /// Encode the package as a ZIP using stored members and the original order.
    ///
    /// Pages and Numbers use a leading `Index/Document.iwa` for package type
    /// discovery, so newly-created document indexes are inserted first.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = StreamingArchiveWriter::new();
        for (name, data) in &self.entries {
            writer.write_stored(name, data).map_err(|error| {
                Error::Bundle(format!("Failed to write package entry {name}: {error}"))
            })?;
        }
        writer
            .finish_to_bytes()
            .map_err(|error| Error::Bundle(format!("Failed to finish iWork ZIP: {error}")))
    }

    /// Atomically save the package to a file in the destination directory.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let path = path.as_ref();
        let parent = path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let bytes = self.to_bytes()?;
        let mut temporary = NamedTempFile::new_in(parent)?;
        temporary.write_all(&bytes)?;
        temporary.as_file_mut().sync_all()?;
        temporary
            .persist(path)
            .map_err(|error| Error::Io(error.error))?;

        // Make the rename durable on filesystems that support directory sync.
        if let Ok(directory) = File::open(parent) {
            directory.sync_all()?;
        }
        Ok(())
    }

    fn entry_position(&self, name: &str) -> Option<usize> {
        self.entries
            .iter()
            .position(|(candidate, _)| candidate == name)
    }

    fn insert_new_entry(&mut self, name: String, data: Vec<u8>) {
        if name == "Index/Document.iwa" {
            self.entries.insert(0, (name, data));
        } else {
            self.entries.push((name, data));
        }
    }
}

fn insert_unique_archive_entry(
    archive: &ArchiveReader<'_>,
    name: &str,
    entries: &mut Vec<(String, Vec<u8>)>,
    seen: &mut HashSet<String>,
) -> Result<()> {
    if !seen.insert(name.to_owned()) {
        return Err(Error::Bundle(format!(
            "Duplicate package entry is ambiguous: {name}"
        )));
    }
    let data = archive
        .read(name)
        .map_err(|error| Error::Bundle(format!("Failed to read package entry {name}: {error}")))?;
    entries.push((name.to_owned(), data));
    Ok(())
}

fn normalize_entry_name(name: &str) -> &str {
    name.strip_prefix('/').unwrap_or(name)
}

pub(crate) fn is_calculation_engine_entry_name(name: &str) -> bool {
    const BASE_NAME: &str = "CalculationEngine.iwa";
    const VERSIONED_PREFIX: &str = "CalculationEngine-";

    name.rsplit('/').next().is_some_and(|file_name| {
        file_name == BASE_NAME
            || file_name
                .strip_prefix(VERSIONED_PREFIX)
                .and_then(|suffix| suffix.strip_suffix(".iwa"))
                .is_some_and(|version| {
                    !version.is_empty()
                        && version.split('-').all(|part| {
                            !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit())
                        })
                })
    })
}

fn is_legacy_operation_storage(name: &str, data: &[u8]) -> bool {
    name.rsplit('/').next() == Some("OperationStorage.iwa") && data.starts_with(b"bvxn")
}

fn validate_entry_name(name: &str) -> Result<()> {
    if name.is_empty() || name.contains('\0') || name.contains('\\') {
        return Err(Error::Bundle(format!(
            "Invalid package entry name: {name:?}"
        )));
    }
    let path = Path::new(name);
    if path.is_absolute()
        || path
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(Error::Bundle(format!("Unsafe package entry name: {name}")));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{ArchiveObject, RawMessage};

    fn archive() -> Archive {
        Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 99,
                        data: vec![1, 2, 3],
                    }],
                )
                .unwrap(),
            ],
        }
    }

    fn legacy_package() -> Vec<u8> {
        let compressed = SnappyStream::compress(&archive().to_bytes().unwrap()).unwrap();
        let mut index = StreamingArchiveWriter::new();
        index
            .write_stored("Index/Document.iwa", &compressed)
            .unwrap();
        let index = index.finish_to_bytes().unwrap();

        let mut outer = StreamingArchiveWriter::new();
        outer
            .write_stored("mac.numbers/preview.jpg", b"preview")
            .unwrap();
        outer.write_stored("mac.numbers/Index.zip", &index).unwrap();
        outer
            .write_stored("mac.numbers/Metadata/Properties.plist", b"plist")
            .unwrap();
        outer.finish_to_bytes().unwrap()
    }

    #[test]
    fn package_entry_and_archive_crud_round_trip() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Metadata/Properties.plist", b"plist".to_vec())
            .unwrap();
        package
            .replace_archive("Index/Document.iwa", &archive())
            .unwrap();

        package
            .update_archive("Index/Document.iwa", |archive| {
                archive.object_mut(1).unwrap().replace_message(
                    0,
                    RawMessage {
                        type_: 100,
                        data: vec![4, 5],
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let bytes = package.to_bytes().unwrap();
        let mut reparsed = IWorkPackage::from_bytes(&bytes).unwrap();
        let document = reparsed.archive("Index/Document.iwa").unwrap();
        assert_eq!(document.object(1).unwrap().messages[0].type_, 100);
        assert_eq!(document.object(1).unwrap().messages[0].data, [4, 5]);
        assert_eq!(
            reparsed.remove_entry("Metadata/Properties.plist"),
            Some(b"plist".to_vec())
        );
        assert_eq!(reparsed.entry_names().next(), Some("Index/Document.iwa"));
    }

    #[test]
    fn preserves_member_order() {
        let mut package = IWorkPackage::new();
        package.insert_entry("Data/a", vec![1]).unwrap();
        package.insert_entry("Data/b", vec![2]).unwrap();
        package
            .replace_archive("Index/Document.iwa", &archive())
            .unwrap();

        let bytes = package.to_bytes().unwrap();
        let reparsed = IWorkPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            reparsed.entry_names().collect::<Vec<_>>(),
            ["Index/Document.iwa", "Data/a", "Data/b"]
        );
    }

    #[test]
    fn rejects_unsafe_entry_names() {
        let mut package = IWorkPackage::new();
        assert!(package.insert_entry("../escape", Vec::new()).is_err());
        assert!(package.insert_entry("/absolute", Vec::new()).is_ok());
        assert!(package.insert_entry("bad\\name", Vec::new()).is_err());
    }

    #[test]
    fn expands_legacy_nested_bundle_for_crud_without_losing_assets() {
        let mut package = IWorkPackage::from_bytes(&legacy_package()).unwrap();
        assert_eq!(
            package.entry_names().collect::<Vec<_>>(),
            [
                "Index/Document.iwa",
                "preview.jpg",
                "Metadata/Properties.plist"
            ]
        );
        assert_eq!(package.entry("preview.jpg"), Some(b"preview".as_slice()));
        assert_eq!(
            package.entry("Metadata/Properties.plist"),
            Some(b"plist".as_slice())
        );

        package
            .update_archive("Index/Document.iwa", |archive| {
                archive.object_mut(1).unwrap().messages[0].type_ = 100;
                Ok(())
            })
            .unwrap();
        let reparsed = IWorkPackage::from_bytes(&package.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reparsed.archive("Index/Document.iwa").unwrap().objects[0].messages[0].type_,
            100
        );
        assert_eq!(reparsed.entry("preview.jpg"), Some(b"preview".as_slice()));
        assert!(!reparsed.contains_entry("Index.zip"));
    }

    #[test]
    fn excludes_legacy_operation_log_from_iwa_archive_scans() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Index/OperationStorage.iwa", b"bvxn log".to_vec())
            .unwrap();
        package
            .replace_archive("Index/Document.iwa", &archive())
            .unwrap();

        assert_eq!(
            package.iwa_entry_names().collect::<Vec<_>>(),
            ["Index/Document.iwa"]
        );
        let error = package.archive("Index/OperationStorage.iwa").unwrap_err();
        assert!(error.to_string().contains("legacy operation log"));
    }

    #[test]
    fn discovers_canonical_and_app_versioned_calculation_engines_strictly() {
        for entry in [
            "Index/CalculationEngine.iwa",
            "Index/CalculationEngine-174.iwa",
            "Index/CalculationEngine-10-2.iwa",
        ] {
            let mut package = IWorkPackage::new();
            package.insert_entry(entry, vec![1]).unwrap();
            assert_eq!(
                package.calculation_engine_entry_name().unwrap(),
                Some(entry)
            );
        }

        for entry in [
            "Index/CalculationEngine-.iwa",
            "Index/CalculationEngine-copy.iwa",
            "Index/CalculationEngine-1-.iwa",
            "Index/CalculationEngine-1.txt",
        ] {
            let mut package = IWorkPackage::new();
            package.insert_entry(entry, vec![1]).unwrap();
            assert_eq!(package.calculation_engine_entry_name().unwrap(), None);
        }
    }

    #[test]
    fn rejects_ambiguous_calculation_engines() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Index/CalculationEngine.iwa", vec![1])
            .unwrap();
        package
            .insert_entry("Index/CalculationEngine-1.iwa", vec![2])
            .unwrap();

        assert!(package.calculation_engine_entry_name().is_err());
    }
}
