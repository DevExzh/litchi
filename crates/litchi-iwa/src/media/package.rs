//! Package-backed and directory-backed media access.

use std::collections::HashMap;
use std::fs;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;

use sha1::{Digest, Sha1};
use tempfile::NamedTempFile;

use crate::archive::RawMessage;
use crate::package::{IWorkPackage, PackageLimits};
use crate::{Error, Result};

use super::codec::{
    PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE, append_data_info, data_entry_name,
    embedded_assets, insert_unique_asset, materialized_file_name, patch_package_metadata,
    remove_data_info, validate_new_media, validate_replacement_type, write_package_entry,
};
use super::model::{EmbeddedMediaAsset, MediaAsset, MediaLimits, MediaStats, MediaType};

const DEFAULT_MAX_REPLACEMENT_BYTES: usize = 1024 * 1024 * 1024;

#[derive(Debug, Clone)]
enum MediaSource {
    Directory(PathBuf),
    File(PathBuf),
    Package(IWorkPackage),
}

/// Read-only media access for directory bundles, package files, and bytes.
#[derive(Debug, Clone)]
pub struct MediaManager {
    pub(super) state: Arc<MediaManagerState>,
}

#[derive(Debug)]
pub(super) struct MediaManagerState {
    source: MediaSource,
    assets: MediaCatalog,
    limits: MediaLimits,
    package_limits: PackageLimits,
}

/// Immutable media catalog with deterministic traversal and checked basename
/// lookup.
///
/// The sorted slice is the public read path; the private map is only an index
/// for the basename selector. Keeping the two responsibilities separate avoids
/// exposing hash-map storage and means repeated ordered traversals do not sort
/// or allocate.
#[derive(Debug, Clone)]
struct MediaCatalog {
    ordered: Box<[MediaAsset]>,
    by_filename: HashMap<String, usize>,
}

impl MediaCatalog {
    fn from_assets(assets: HashMap<String, MediaAsset>) -> Self {
        let mut ordered: Vec<_> = assets.into_values().collect();
        ordered.sort_unstable_by(|left, right| {
            left.path
                .as_os_str()
                .cmp(right.path.as_os_str())
                .then_with(|| left.filename.cmp(&right.filename))
        });

        let by_filename = ordered
            .iter()
            .enumerate()
            .map(|(index, asset)| (asset.filename.clone(), index))
            .collect();

        Self {
            ordered: ordered.into_boxed_slice(),
            by_filename,
        }
    }

    fn as_slice(&self) -> &[MediaAsset] {
        &self.ordered
    }

    fn len(&self) -> usize {
        self.ordered.len()
    }

    fn get(&self, filename: &str) -> Option<&MediaAsset> {
        self.by_filename
            .get(filename)
            .and_then(|&index| self.ordered.get(index))
    }
}

impl MediaManager {
    fn from_parts(
        source: MediaSource,
        assets: HashMap<String, MediaAsset>,
        limits: MediaLimits,
        package_limits: PackageLimits,
    ) -> Self {
        Self {
            state: Arc::new(MediaManagerState {
                source,
                assets: MediaCatalog::from_assets(assets),
                limits,
                package_limits,
            }),
        }
    }

    /// Return another handle to the same immutable media snapshot.
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Open a directory bundle or a single-file `.pages`, `.numbers`, or `.key` package.
    pub fn new<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::new_with_limits(path, MediaLimits::default())
    }

    /// Open a media source under caller-selected resource ceilings.
    pub fn new_with_limits<P: AsRef<Path>>(path: P, limits: MediaLimits) -> Result<Self> {
        Self::new_with_limits_and_package_limits(path, limits, PackageLimits::default())
    }

    /// Open a media source with separate media and package resource ceilings.
    ///
    /// File-backed packages retain the package profile for later extraction;
    /// reopening an asset therefore cannot silently fall back to a broader
    /// default ZIP or Snappy budget.
    pub fn new_with_limits_and_package_limits<P: AsRef<Path>>(
        path: P,
        limits: MediaLimits,
        package_limits: PackageLimits,
    ) -> Result<Self> {
        let path = path.as_ref().to_path_buf();
        let metadata = fs::symlink_metadata(&path).map_err(|error| {
            if error.kind() == std::io::ErrorKind::NotFound {
                Error::Bundle(format!("Media source does not exist: {}", path.display()))
            } else {
                error.into()
            }
        })?;
        if metadata.file_type().is_symlink() {
            return Err(Error::Bundle(format!(
                "Media source must not be a symbolic link: {}",
                path.display()
            )));
        }
        if metadata.is_dir() {
            let mut assets = HashMap::new();
            Self::scan_directory_bundle(&path, &mut assets, limits)?;
            Ok(Self::from_parts(
                MediaSource::Directory(path),
                assets,
                limits,
                package_limits,
            ))
        } else if metadata.is_file() {
            let package = IWorkPackage::open_with_limits(&path, package_limits)?;
            let assets = Self::scan_package(&package, limits)?;
            Ok(Self::from_parts(
                MediaSource::File(path),
                assets,
                limits,
                package_limits,
            ))
        } else {
            Err(Error::Bundle(format!(
                "Media source is not a regular file or directory: {}",
                path.display()
            )))
        }
    }

    /// Open read-only media access over an in-memory iWork package.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, MediaLimits::default())
    }

    /// Open an in-memory iWork package under caller-selected resource ceilings.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: MediaLimits) -> Result<Self> {
        Self::from_bytes_with_limits_and_package_limits(bytes, limits, PackageLimits::default())
    }

    /// Open an in-memory package with separate media and package ceilings.
    pub fn from_bytes_with_limits_and_package_limits(
        bytes: &[u8],
        limits: MediaLimits,
        package_limits: PackageLimits,
    ) -> Result<Self> {
        Self::from_package_with_limits(
            IWorkPackage::from_bytes_with_limits(bytes, package_limits)?,
            limits,
        )
    }

    /// Create read-only media access from an already parsed package.
    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        Self::from_package_with_limits(package, MediaLimits::default())
    }

    /// Create read-only media access from a package under explicit limits.
    pub fn from_package_with_limits(package: IWorkPackage, limits: MediaLimits) -> Result<Self> {
        let package_limits = package.limits();
        let assets = Self::scan_package(&package, limits)?;
        Ok(Self::from_parts(
            MediaSource::Package(package),
            assets,
            limits,
            package_limits,
        ))
    }

    fn scan_directory_bundle(
        bundle_path: &Path,
        assets: &mut HashMap<String, MediaAsset>,
        limits: MediaLimits,
    ) -> Result<()> {
        let data_dir = bundle_path.join("Data");
        match fs::symlink_metadata(&data_dir) {
            Ok(metadata) if metadata.file_type().is_symlink() => {
                return Err(Error::Bundle(format!(
                    "Media Data directory is a symbolic link: {}",
                    data_dir.display()
                )));
            },
            Ok(metadata) if metadata.is_dir() => {
                let mut total_size = 0;
                Self::scan_directory_recursive(
                    &data_dir,
                    bundle_path,
                    assets,
                    limits,
                    &mut total_size,
                )?;
            },
            Ok(_) => {},
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {},
            Err(error) => return Err(error.into()),
        }
        Ok(())
    }

    fn scan_directory_recursive(
        directory: &Path,
        bundle_root: &Path,
        assets: &mut HashMap<String, MediaAsset>,
        limits: MediaLimits,
        total_size: &mut u64,
    ) -> Result<()> {
        for entry in fs::read_dir(directory)? {
            let entry = entry?;
            let file_type = entry.file_type()?;
            let path = entry.path();
            if file_type.is_symlink() {
                return Err(Error::Bundle(format!(
                    "Media bundle contains a symbolic link: {}",
                    path.display()
                )));
            }
            if file_type.is_dir() {
                Self::scan_directory_recursive(&path, bundle_root, assets, limits, total_size)?;
            } else if file_type.is_file() {
                let relative = path
                    .strip_prefix(bundle_root)
                    .map_err(|_| {
                        Error::Bundle(format!(
                            "Media path escapes its bundle root: {}",
                            path.display()
                        ))
                    })?
                    .to_path_buf();
                let asset = MediaAsset::new(relative, entry.metadata()?.len());
                insert_unique_asset(assets, asset, limits, total_size)?;
            }
        }
        Ok(())
    }

    fn scan_package(
        package: &IWorkPackage,
        limits: MediaLimits,
    ) -> Result<HashMap<String, MediaAsset>> {
        let mut assets = HashMap::new();
        let mut total_size = 0;
        for name in package
            .entry_names()
            .filter(|name| name.starts_with("Data/"))
        {
            let data = package.entry(name).ok_or_else(|| {
                Error::Bundle(format!("Package entry disappeared while scanning: {name}"))
            })?;
            let size = u64::try_from(data.len())
                .map_err(|_| Error::Bundle(format!("Media asset is too large: {name}")))?;
            insert_unique_asset(
                &mut assets,
                MediaAsset::new(PathBuf::from(name), size),
                limits,
                &mut total_size,
            )?;
        }
        Ok(assets)
    }

    /// Borrow all materialized assets in deterministic relative-path order.
    ///
    /// The returned slice is backed by the immutable media snapshot and does
    /// not expose the private basename lookup index.
    pub fn assets(&self) -> &[MediaAsset] {
        self.state.assets.as_slice()
    }

    /// Iterate over all materialized assets without allocating.
    pub fn iter_assets(&self) -> impl Iterator<Item = &MediaAsset> + '_ {
        self.state.assets.as_slice().iter()
    }

    /// Return all materialized assets in deterministic relative-path order.
    ///
    /// This is a borrowed view; use [`Self::assets`] when slice access is more
    /// convenient.
    pub fn assets_in_order(&self) -> &[MediaAsset] {
        self.state.assets.as_slice()
    }

    /// Return the checked resource profile used by this manager.
    pub fn limits(&self) -> MediaLimits {
        self.state.limits
    }

    /// Return the package profile retained for file-backed reopen operations.
    pub fn package_limits(&self) -> PackageLimits {
        self.state.package_limits
    }

    pub fn get(&self, filename: &str) -> Option<&MediaAsset> {
        self.state.assets.get(filename)
    }

    /// Return matching assets in deterministic relative-path order.
    pub fn assets_by_type(&self, media_type: MediaType) -> Vec<&MediaAsset> {
        self.iter_assets()
            .filter(|asset| asset.media_type == media_type)
            .collect()
    }

    pub fn images(&self) -> Vec<&MediaAsset> {
        self.assets_by_type(MediaType::Image)
    }

    pub fn videos(&self) -> Vec<&MediaAsset> {
        self.assets_by_type(MediaType::Video)
    }

    pub fn audio(&self) -> Vec<&MediaAsset> {
        self.assets_by_type(MediaType::Audio)
    }

    /// Extract an asset by its basename.
    pub fn extract(&self, filename: &str) -> Result<Vec<u8>> {
        let asset = self
            .get(filename)
            .ok_or_else(|| Error::Bundle(format!("Media asset not found: {filename}")))?;
        if asset.size > self.state.limits.max_asset_bytes {
            return Err(Error::Bundle(format!(
                "Media asset {filename} is {} bytes, exceeding the configured {}-byte limit",
                asset.size, self.state.limits.max_asset_bytes
            )));
        }

        let capacity = usize::try_from(asset.size).map_err(|_| {
            Error::Bundle(format!(
                "Media asset {} does not fit in memory on this target",
                asset.path.display()
            ))
        })?;
        let mut data = Vec::new();
        data.try_reserve_exact(capacity).map_err(|error| {
            Error::Bundle(format!(
                "Unable to reserve {} bytes for media asset {}: {error}",
                capacity,
                asset.path.display()
            ))
        })?;
        self.extract_to_writer(filename, &mut data)?;
        Ok(data)
    }

    /// Stream an asset to a caller-owned sequential sink.
    ///
    /// The configured per-asset limit is checked before reading. Directory
    /// sources are additionally bounded while reading so a file that grows
    /// after discovery cannot cause an unbounded allocation or copy.
    pub fn extract_to_writer<W: Write>(&self, filename: &str, mut sink: W) -> Result<()> {
        let asset = self
            .get(filename)
            .ok_or_else(|| Error::Bundle(format!("Media asset not found: {filename}")))?;
        if asset.size > self.state.limits.max_asset_bytes {
            return Err(Error::Bundle(format!(
                "Media asset {filename} is {} bytes, exceeding the configured {}-byte limit",
                asset.size, self.state.limits.max_asset_bytes
            )));
        }

        match &self.state.source {
            MediaSource::Directory(root) => {
                let path = root.join(&asset.path);
                let metadata = fs::symlink_metadata(&path)?;
                if metadata.file_type().is_symlink() {
                    return Err(Error::Bundle(format!(
                        "Media asset is a symbolic link: {}",
                        path.display()
                    )));
                }
                if !metadata.is_file() {
                    return Err(Error::Bundle(format!(
                        "Media asset is not a regular file: {}",
                        path.display()
                    )));
                }
                if metadata.len() > self.state.limits.max_asset_bytes {
                    return Err(Error::Bundle(format!(
                        "Media asset {} grew to {} bytes, exceeding the configured {}-byte limit",
                        path.display(),
                        metadata.len(),
                        self.state.limits.max_asset_bytes
                    )));
                }

                let file = fs::File::open(&path)?;
                let mut bounded = file.take(self.state.limits.max_asset_bytes.saturating_add(1));
                let written = std::io::copy(&mut bounded, &mut sink)?;
                if written > self.state.limits.max_asset_bytes {
                    return Err(Error::Bundle(format!(
                        "Media asset {} exceeded the configured {}-byte limit while reading",
                        path.display(),
                        self.state.limits.max_asset_bytes
                    )));
                }
                Ok(())
            },
            MediaSource::File(path) => {
                let package = IWorkPackage::open_with_limits(path, self.state.package_limits)?;
                write_package_entry(&package, asset, self.state.limits, &mut sink)
            },
            MediaSource::Package(package) => {
                write_package_entry(package, asset, self.state.limits, &mut sink)
            },
        }
    }

    /// Atomically stream an asset to a regular file.
    pub fn extract_to_file(&self, filename: &str, output_path: &Path) -> Result<()> {
        let parent = output_path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let existing = match fs::symlink_metadata(output_path) {
            Ok(metadata) if metadata.file_type().is_symlink() => {
                return Err(Error::Bundle(format!(
                    "Media output destination must not be a symbolic link: {}",
                    output_path.display()
                )));
            },
            Ok(metadata) if metadata.is_file() => Some(metadata),
            Ok(_) => {
                return Err(Error::Bundle(format!(
                    "Media output destination is not a regular file: {}",
                    output_path.display()
                )));
            },
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
            Err(error) => return Err(error.into()),
        };

        let mut temporary = NamedTempFile::new_in(parent)?;
        self.extract_to_writer(filename, temporary.as_file_mut())?;
        if let Some(metadata) = existing {
            fs::set_permissions(temporary.path(), metadata.permissions())?;
        }
        temporary.as_file_mut().sync_all()?;
        temporary
            .persist(output_path)
            .map_err(|error| Error::Io(error.error))?;
        if let Ok(directory) = fs::File::open(parent) {
            directory.sync_all()?;
        }
        Ok(())
    }

    pub fn stats(&self) -> MediaStats {
        let mut stats = MediaStats {
            total_count: self.state.assets.len(),
            ..MediaStats::default()
        };
        for asset in self.iter_assets() {
            stats.total_size = stats.total_size.saturating_add(asset.size);
            match asset.media_type {
                MediaType::Image => stats.image_count += 1,
                MediaType::Video => stats.video_count += 1,
                MediaType::Audio => stats.audio_count += 1,
                MediaType::Pdf => stats.pdf_count += 1,
                MediaType::Unknown => stats.unknown_count += 1,
            }
        }
        stats
    }
}

/// Transactional editor for existing embedded iWork assets.
///
/// Replacement retains the data identifier and all object/component references.
/// The package member, SHA-1 digest, and materialized length are changed together.
#[derive(Debug, Clone)]
pub struct IWorkMediaEditor {
    package: IWorkPackage,
    assets: Vec<EmbeddedMediaAsset>,
    max_replacement_bytes: usize,
}

impl IWorkMediaEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(IWorkPackage::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(IWorkPackage::from_bytes(bytes)?)
    }

    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        let assets = embedded_assets(&package)?;
        Ok(Self {
            package,
            assets,
            max_replacement_bytes: DEFAULT_MAX_REPLACEMENT_BYTES,
        })
    }

    /// Set the in-memory replacement safety limit. The default is 1 GiB.
    pub fn set_max_replacement_bytes(&mut self, limit: usize) -> Result<()> {
        if limit == 0 {
            return Err(Error::Bundle(
                "Media replacement limit must be greater than zero".to_owned(),
            ));
        }
        self.max_replacement_bytes = limit;
        Ok(())
    }

    pub fn assets(&self) -> &[EmbeddedMediaAsset] {
        &self.assets
    }

    pub fn asset(&self, data_identifier: u64) -> Option<&EmbeddedMediaAsset> {
        self.assets
            .iter()
            .find(|asset| asset.data_identifier == data_identifier)
    }

    pub fn extract(&self, data_identifier: u64) -> Result<Vec<u8>> {
        let asset = self.asset(data_identifier).ok_or_else(|| {
            Error::Bundle(format!("Data identifier {data_identifier} does not exist"))
        })?;
        let path = asset.package_path.as_deref().ok_or_else(|| {
            Error::Bundle(format!(
                "Data identifier {data_identifier} is not materialized in this package"
            ))
        })?;
        self.package
            .entry(path)
            .map(<[u8]>::to_vec)
            .ok_or_else(|| Error::Bundle(format!("Materialized data entry not found: {path}")))
    }

    /// Replace an existing materialized asset and return its previous bytes.
    ///
    /// The operation is staged on a package clone, serialized, reparsed, and
    /// verified before it becomes visible through this editor.
    pub fn replace(&mut self, data_identifier: u64, replacement: &[u8]) -> Result<Vec<u8>> {
        if replacement.is_empty() {
            return Err(Error::Bundle(
                "A materialized media asset cannot be replaced with empty data".to_owned(),
            ));
        }
        if replacement.len() > self.max_replacement_bytes {
            return Err(Error::Bundle(format!(
                "Replacement is {} bytes, exceeding the configured {}-byte limit",
                replacement.len(),
                self.max_replacement_bytes
            )));
        }

        let asset = self.asset(data_identifier).cloned().ok_or_else(|| {
            Error::Bundle(format!("Data identifier {data_identifier} does not exist"))
        })?;
        let path = asset.package_path.clone().ok_or_else(|| {
            Error::Bundle(format!(
                "Data identifier {data_identifier} is not materialized and cannot be replaced safely"
            ))
        })?;
        validate_replacement_type(&asset, replacement)?;

        let digest = Sha1::digest(replacement).to_vec();
        let replacement_length = u64::try_from(replacement.len())
            .map_err(|_| Error::Bundle("Replacement length exceeds u64".to_owned()))?;
        let mut staged = self.package.clone();
        let previous = staged
            .insert_entry(path.clone(), replacement.to_vec())?
            .ok_or_else(|| Error::Bundle(format!("Materialized data entry not found: {path}")))?;

        update_package_metadata(&mut staged, |metadata| {
            patch_package_metadata(metadata, data_identifier, &digest, replacement_length)
        })?;

        let serialized = staged.to_bytes()?;
        let verified = Self::from_bytes(&serialized)?;
        let verified_asset = verified.asset(data_identifier).ok_or_else(|| {
            Error::Bundle("Replaced data identifier vanished during verification".to_owned())
        })?;
        if verified.extract(data_identifier)? != replacement
            || verified_asset.digest != digest
            || verified_asset.declared_size != Some(replacement_length)
            || verified_asset.size != Some(replacement_length)
        {
            return Err(Error::Bundle(
                "Media replacement failed post-serialization verification".to_owned(),
            ));
        }
        self.package = staged;
        self.assets = verified.assets;
        Ok(previous)
    }

    /// Insert a materialized data record without attaching it to an app object.
    ///
    /// The returned asset is deliberately unreferenced and can be attached by a
    /// higher-level drawable mutation or removed with [`Self::remove_unreferenced`].
    pub fn insert_unreferenced(
        &mut self,
        preferred_filename: &str,
        data: &[u8],
    ) -> Result<EmbeddedMediaAsset> {
        validate_new_media(preferred_filename, data, self.max_replacement_bytes)?;
        let data_identifier = self
            .assets
            .iter()
            .map(|asset| asset.data_identifier)
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| Error::Bundle("Data identifier space is exhausted".to_owned()))?;
        let file_name = materialized_file_name(preferred_filename, data_identifier)?;
        let package_path = data_entry_name(&file_name)?;
        if self.package.contains_entry(&package_path) {
            return Err(Error::Bundle(format!(
                "Allocated media package entry already exists: {package_path}"
            )));
        }

        let digest = Sha1::digest(data).to_vec();
        let length = u64::try_from(data.len())
            .map_err(|_| Error::Bundle("Media length exceeds u64".to_owned()))?;
        let mut staged = self.package.clone();
        if staged
            .insert_entry(package_path.clone(), data.to_vec())?
            .is_some()
        {
            return Err(Error::Bundle(format!(
                "Allocated media package entry already exists: {package_path}"
            )));
        }
        update_package_metadata(&mut staged, |metadata| {
            append_data_info(
                metadata,
                data_identifier,
                &digest,
                preferred_filename,
                &file_name,
                length,
            )
        })?;

        let serialized = staged.to_bytes()?;
        let verified = Self::from_bytes(&serialized)?;
        let inserted = verified.asset(data_identifier).cloned().ok_or_else(|| {
            Error::Bundle("Inserted data identifier vanished during verification".to_owned())
        })?;
        if inserted.is_referenced()
            || inserted.package_path.as_deref() != Some(package_path.as_str())
            || inserted.digest != digest
            || inserted.declared_size != Some(length)
            || verified.extract(data_identifier)? != data
        {
            return Err(Error::Bundle(
                "Media insertion failed post-serialization verification".to_owned(),
            ));
        }
        self.package = staged;
        self.assets = verified.assets;
        Ok(inserted)
    }

    /// Remove an asset that is absent from every known package reference index.
    ///
    /// Referenced assets are rejected. For an unmaterialized record the return
    /// value is `None`; otherwise it contains the removed `Data/*` bytes.
    pub fn remove_unreferenced(&mut self, data_identifier: u64) -> Result<Option<Vec<u8>>> {
        let asset = self.asset(data_identifier).cloned().ok_or_else(|| {
            Error::Bundle(format!("Data identifier {data_identifier} does not exist"))
        })?;
        if asset.is_referenced() {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is still referenced and cannot be removed"
            )));
        }

        let mut staged = self.package.clone();
        let previous = asset
            .package_path
            .as_deref()
            .and_then(|path| staged.remove_entry(path));
        update_package_metadata(&mut staged, |metadata| {
            remove_data_info(metadata, data_identifier)
        })?;

        let serialized = staged.to_bytes()?;
        let verified = Self::from_bytes(&serialized)?;
        if verified.asset(data_identifier).is_some()
            || asset
                .package_path
                .as_deref()
                .is_some_and(|path| verified.package.contains_entry(path))
        {
            return Err(Error::Bundle(
                "Media removal failed post-serialization verification".to_owned(),
            ));
        }
        self.package = staged;
        self.assets = verified.assets;
        Ok(previous)
    }

    pub fn package(&self) -> &IWorkPackage {
        &self.package
    }

    pub fn into_package(self) -> IWorkPackage {
        self.package
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

fn update_package_metadata(
    package: &mut IWorkPackage,
    update: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
) -> Result<()> {
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let mut location = None;
        for (object_index, object) in archive.objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                    && location.replace((object_index, message_index)).is_some()
                {
                    return Err(Error::Bundle(
                        "Package contains multiple PackageMetadata payloads".to_owned(),
                    ));
                }
            }
        }
        let (object_index, message_index) = location
            .ok_or_else(|| Error::Bundle("PackageMetadata payload was not found".to_owned()))?;
        let object = &mut archive.objects[object_index];
        let old = &object.messages[message_index];
        let patched = update(&old.data)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: old.type_,
                data: patched,
            },
        )?;
        Ok(())
    })
}
