//! Media discovery, extraction, and transactional replacement.
//!
//! iWork packages store materialized assets as `Data/*` ZIP members and
//! describe them with `TSP.DataInfo` records in `Index/Metadata.iwa`.

use std::collections::VecDeque;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::io::Read;
use std::ops::Range;
use std::path::{Component, Path, PathBuf};

use prost::Message;
use sha1::{Digest, Sha1};

use crate::archive::RawMessage;
use crate::package::IWorkPackage;
use crate::protobuf;
use crate::varint::{decode_varint_from_bytes, encode_varint};
use crate::{Error, Result};

const PACKAGE_METADATA_ENTRY: &str = "Index/Metadata.iwa";
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;
const DEFAULT_MAX_REPLACEMENT_BYTES: usize = 1024 * 1024 * 1024;

/// Types of media assets that can be found in iWork documents.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MediaType {
    /// Image file (PNG, JPEG, TIFF, HEIF, etc.).
    Image,
    /// Video file (MP4, MOV, etc.).
    Video,
    /// Audio file (MP3, AAC, WAV, etc.).
    Audio,
    /// PDF document.
    Pdf,
    /// Unknown or unsupported media type.
    Unknown,
}

impl MediaType {
    /// Detect a media type from a filename extension.
    pub fn from_extension(ext: &str) -> Self {
        match ext.to_ascii_lowercase().as_str() {
            "png" | "jpg" | "jpeg" | "gif" | "tiff" | "tif" | "bmp" | "heic" | "heif" | "webp"
            | "svg" => Self::Image,
            "mp4" | "mov" | "m4v" | "avi" | "mkv" => Self::Video,
            "mp3" | "aac" | "m4a" | "wav" | "aiff" | "aif" | "ogg" => Self::Audio,
            "pdf" => Self::Pdf,
            _ => Self::Unknown,
        }
    }

    /// Sniff common media signatures without trusting the filename.
    pub fn from_bytes(data: &[u8]) -> Self {
        if data.starts_with(b"\x89PNG\r\n\x1a\n")
            || data.starts_with(b"\xff\xd8\xff")
            || data.starts_with(b"GIF87a")
            || data.starts_with(b"GIF89a")
            || data.starts_with(b"II*\0")
            || data.starts_with(b"MM\0*")
            || data.starts_with(b"BM")
            || (data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WEBP")
        {
            return Self::Image;
        }
        if data.starts_with(b"%PDF-") {
            return Self::Pdf;
        }
        if data.starts_with(b"ID3")
            || data
                .get(..2)
                .is_some_and(|prefix| prefix[0] == 0xff && prefix[1] & 0xe0 == 0xe0)
            || (data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WAVE")
            || (data.len() >= 12
                && &data[..4] == b"FORM"
                && matches!(&data[8..12], b"AIFF" | b"AIFC"))
            || data.starts_with(b"OggS")
        {
            return Self::Audio;
        }
        if data.len() >= 12 && &data[4..8] == b"ftyp" {
            return match &data[8..12] {
                b"heic" | b"heix" | b"hevc" | b"hevx" | b"heim" | b"heis" | b"mif1" | b"msf1"
                | b"avif" | b"avis" => Self::Image,
                b"M4A " | b"M4B " | b"M4P " => Self::Audio,
                _ => Self::Video,
            };
        }
        Self::Unknown
    }

    /// Get a human-readable name for this media type.
    pub fn name(self) -> &'static str {
        match self {
            Self::Image => "Image",
            Self::Video => "Video",
            Self::Audio => "Audio",
            Self::Pdf => "PDF Document",
            Self::Unknown => "Unknown",
        }
    }
}

/// Information about a materialized `Data/*` package member.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaAsset {
    /// Relative path within the bundle.
    pub path: PathBuf,
    /// Media type inferred from the filename.
    pub media_type: MediaType,
    /// File size in bytes.
    pub size: u64,
    /// File name without its parent path.
    pub filename: String,
}

impl MediaAsset {
    /// Create a media asset entry.
    pub fn new(path: PathBuf, size: u64) -> Self {
        let filename = path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("unknown")
            .to_owned();
        let media_type = path
            .extension()
            .and_then(|extension| extension.to_str())
            .map(MediaType::from_extension)
            .unwrap_or(MediaType::Unknown);
        Self {
            path,
            media_type,
            size,
            filename,
        }
    }

    pub fn is_image(&self) -> bool {
        self.media_type == MediaType::Image
    }

    pub fn is_video(&self) -> bool {
        self.media_type == MediaType::Video
    }

    pub fn is_audio(&self) -> bool {
        self.media_type == MediaType::Audio
    }
}

#[derive(Debug, Clone)]
enum MediaSource {
    Directory(PathBuf),
    File(PathBuf),
    Package(IWorkPackage),
}

/// Read-only media access for directory bundles, package files, and bytes.
#[derive(Debug, Clone)]
pub struct MediaManager {
    source: MediaSource,
    assets: HashMap<String, MediaAsset>,
}

impl MediaManager {
    /// Open a directory bundle or a single-file `.pages`, `.numbers`, or `.key` package.
    pub fn new<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref().to_path_buf();
        if path.is_dir() {
            let mut assets = HashMap::new();
            Self::scan_directory_bundle(&path, &mut assets)?;
            Ok(Self {
                source: MediaSource::Directory(path),
                assets,
            })
        } else if path.is_file() {
            let package = IWorkPackage::open(&path)?;
            let assets = Self::scan_package(&package)?;
            Ok(Self {
                source: MediaSource::File(path),
                assets,
            })
        } else {
            Err(Error::Bundle(format!(
                "Media source does not exist: {}",
                path.display()
            )))
        }
    }

    /// Open read-only media access over an in-memory iWork package.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(IWorkPackage::from_bytes(bytes)?)
    }

    /// Create read-only media access from an already parsed package.
    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        let assets = Self::scan_package(&package)?;
        Ok(Self {
            source: MediaSource::Package(package),
            assets,
        })
    }

    fn scan_directory_bundle(
        bundle_path: &Path,
        assets: &mut HashMap<String, MediaAsset>,
    ) -> Result<()> {
        let data_dir = bundle_path.join("Data");
        if data_dir.is_dir() {
            Self::scan_directory_recursive(&data_dir, bundle_path, assets)?;
        }
        Ok(())
    }

    fn scan_directory_recursive(
        directory: &Path,
        bundle_root: &Path,
        assets: &mut HashMap<String, MediaAsset>,
    ) -> Result<()> {
        for entry in fs::read_dir(directory)? {
            let entry = entry?;
            let path = entry.path();
            if path.is_dir() {
                Self::scan_directory_recursive(&path, bundle_root, assets)?;
            } else if path.is_file() {
                let relative = path
                    .strip_prefix(bundle_root)
                    .unwrap_or(&path)
                    .to_path_buf();
                let asset = MediaAsset::new(relative, entry.metadata()?.len());
                insert_unique_asset(assets, asset)?;
            }
        }
        Ok(())
    }

    fn scan_package(package: &IWorkPackage) -> Result<HashMap<String, MediaAsset>> {
        let mut assets = HashMap::new();
        for name in package
            .entry_names()
            .filter(|name| name.starts_with("Data/"))
        {
            let data = package.entry(name).ok_or_else(|| {
                Error::Bundle(format!("Package entry disappeared while scanning: {name}"))
            })?;
            let size = u64::try_from(data.len())
                .map_err(|_| Error::Bundle(format!("Media asset is too large: {name}")))?;
            insert_unique_asset(&mut assets, MediaAsset::new(PathBuf::from(name), size))?;
        }
        Ok(assets)
    }

    pub fn assets(&self) -> &HashMap<String, MediaAsset> {
        &self.assets
    }

    pub fn get(&self, filename: &str) -> Option<&MediaAsset> {
        self.assets.get(filename)
    }

    pub fn assets_by_type(&self, media_type: MediaType) -> Vec<&MediaAsset> {
        self.assets
            .values()
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
        match &self.source {
            MediaSource::Directory(root) => {
                let mut file = fs::File::open(root.join(&asset.path))?;
                let capacity = usize::try_from(asset.size).unwrap_or(0);
                let mut data = Vec::with_capacity(capacity);
                file.read_to_end(&mut data)?;
                Ok(data)
            },
            MediaSource::File(path) => {
                let package = IWorkPackage::open(path)?;
                extract_package_entry(&package, asset)
            },
            MediaSource::Package(package) => extract_package_entry(package, asset),
        }
    }

    pub fn extract_to_file(&self, filename: &str, output_path: &Path) -> Result<()> {
        fs::write(output_path, self.extract(filename)?)?;
        Ok(())
    }

    pub fn stats(&self) -> MediaStats {
        let mut stats = MediaStats {
            total_count: self.assets.len(),
            ..MediaStats::default()
        };
        for asset in self.assets.values() {
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

fn insert_unique_asset(assets: &mut HashMap<String, MediaAsset>, asset: MediaAsset) -> Result<()> {
    if let Some(previous) = assets.insert(asset.filename.clone(), asset.clone()) {
        return Err(Error::Bundle(format!(
            "Media basenames are ambiguous: {} and {}",
            previous.path.display(),
            asset.path.display()
        )));
    }
    Ok(())
}

fn extract_package_entry(package: &IWorkPackage, asset: &MediaAsset) -> Result<Vec<u8>> {
    let name = asset.path.to_str().ok_or_else(|| {
        Error::Bundle(format!(
            "Media path is not valid UTF-8: {}",
            asset.path.display()
        ))
    })?;
    package
        .entry(name)
        .map(<[u8]>::to_vec)
        .ok_or_else(|| Error::Bundle(format!("Media package entry not found: {name}")))
}

/// Metadata-backed view of one `TSP.DataInfo` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedMediaAsset {
    /// Stable `TSP.DataInfo.identifier` used by drawable protobufs.
    pub data_identifier: u64,
    /// Name iWork prefers when materializing or exporting the asset.
    pub preferred_filename: String,
    /// Materialized package member, when present.
    pub package_path: Option<String>,
    /// Type inferred from the preferred/materialized filename.
    pub media_type: MediaType,
    /// Actual materialized byte length.
    pub size: Option<u64>,
    /// Length declared in `TSP.DataInfo`.
    pub declared_size: Option<u64>,
    /// Digest declared in `TSP.DataInfo` (normally SHA-1).
    pub digest: Vec<u8>,
    /// Sum of component-level object reference counts.
    pub component_reference_count: u64,
    /// Number of `ComponentDataReference` records, including zero-count records.
    pub component_reference_record_count: usize,
    /// Number of aggregate `MessageInfo.data_references` occurrences.
    pub message_reference_count: usize,
    /// Whether the package's `DataMetadataMap` contains this identifier.
    pub has_data_metadata: bool,
    /// Object identifiers whose MessageInfo/component records reference this data.
    pub referencing_object_ids: Vec<u64>,
}

impl EmbeddedMediaAsset {
    pub fn is_materialized(&self) -> bool {
        self.package_path.is_some() && self.size.is_some()
    }

    pub fn is_referenced(&self) -> bool {
        self.component_reference_record_count != 0
            || self.message_reference_count != 0
            || self.has_data_metadata
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

fn validate_replacement_type(asset: &EmbeddedMediaAsset, replacement: &[u8]) -> Result<()> {
    let detected = MediaType::from_bytes(replacement);
    if asset.media_type != MediaType::Unknown
        && detected != MediaType::Unknown
        && asset.media_type != detected
    {
        return Err(Error::Bundle(format!(
            "Replacement signature is {}, but {} is declared as {}",
            detected.name(),
            asset.preferred_filename,
            asset.media_type.name()
        )));
    }
    Ok(())
}

fn validate_new_media(filename: &str, data: &[u8], maximum_length: usize) -> Result<()> {
    if data.is_empty() {
        return Err(Error::Bundle(
            "A materialized media asset cannot contain empty data".to_owned(),
        ));
    }
    if data.len() > maximum_length {
        return Err(Error::Bundle(format!(
            "Media is {} bytes, exceeding the configured {}-byte limit",
            data.len(),
            maximum_length
        )));
    }
    let path = Path::new(filename);
    if path.file_name().and_then(|name| name.to_str()) != Some(filename) {
        return Err(Error::Bundle(format!(
            "Preferred media filename must be a safe basename: {filename:?}"
        )));
    }
    data_entry_name(filename)?;
    let expected = path
        .extension()
        .and_then(|extension| extension.to_str())
        .map(MediaType::from_extension)
        .unwrap_or(MediaType::Unknown);
    let detected = MediaType::from_bytes(data);
    if expected != MediaType::Unknown && detected != MediaType::Unknown && expected != detected {
        return Err(Error::Bundle(format!(
            "Media signature is {}, but {filename} is declared as {}",
            detected.name(),
            expected.name()
        )));
    }
    Ok(())
}

fn materialized_file_name(preferred_filename: &str, data_identifier: u64) -> Result<String> {
    let path = Path::new(preferred_filename);
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .ok_or_else(|| Error::Bundle("Preferred media filename has no stem".to_owned()))?;
    Ok(
        match path.extension().and_then(|extension| extension.to_str()) {
            Some(extension) if !extension.is_empty() => {
                format!("{stem}-{data_identifier}.{extension}")
            },
            _ => format!("{stem}-{data_identifier}"),
        },
    )
}

fn embedded_assets(package: &IWorkPackage) -> Result<Vec<EmbeddedMediaAsset>> {
    let metadata = decode_package_metadata(package)?;
    let mut component_counts = HashMap::<u64, u64>::new();
    let mut component_record_counts = HashMap::<u64, usize>::new();
    let mut referencing_objects = HashMap::<u64, HashSet<u64>>::new();
    for component in metadata
        .components
        .iter()
        .chain(metadata.versioned_components.iter())
    {
        for reference in &component.data_references {
            let record_count = component_record_counts
                .entry(reference.data_identifier)
                .or_default();
            *record_count = record_count.checked_add(1).ok_or_else(|| {
                Error::Bundle("Component data reference record count overflow".to_owned())
            })?;
            let count = reference
                .object_reference_list
                .iter()
                .try_fold(0u64, |sum, object| sum.checked_add(u64::from(object.count)))
                .ok_or_else(|| {
                    Error::Bundle("Component data reference count overflow".to_owned())
                })?;
            let current = component_counts
                .entry(reference.data_identifier)
                .or_default();
            *current = current.checked_add(count).ok_or_else(|| {
                Error::Bundle("Component data reference count overflow".to_owned())
            })?;
            for object in &reference.object_reference_list {
                referencing_objects
                    .entry(reference.data_identifier)
                    .or_default()
                    .insert(object.object_identifier);
            }
        }
    }

    let mut message_counts = HashMap::<u64, usize>::new();
    let metadata_map_identifier = metadata
        .data_metadata_map
        .as_ref()
        .map(|reference| reference.identifier);
    let mut data_metadata_ids = HashSet::new();
    let mut metadata_map_payloads = 0usize;
    let iwa_names = package
        .iwa_entry_names()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    for name in iwa_names {
        let archive = package.archive(&name)?;
        for object in archive.objects {
            let object_identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("Object in {name} has no identifier"))
            })?;
            if object.archive_info.identifier == metadata_map_identifier {
                for message in &object.messages {
                    if message.type_ == DATA_METADATA_MAP_MESSAGE_TYPE {
                        metadata_map_payloads =
                            metadata_map_payloads.checked_add(1).ok_or_else(|| {
                                Error::Bundle("DataMetadataMap payload count overflow".to_owned())
                            })?;
                        let map = protobuf::tsp::DataMetadataMap::decode(message.data.as_slice())?;
                        for entry in map.data_metadata_entries {
                            data_metadata_ids.insert(entry.data_identifier);
                        }
                    }
                }
            }
            for info in object.archive_info.message_infos {
                for identifier in info.data_references {
                    let count = message_counts.entry(identifier).or_default();
                    *count = count.checked_add(1).ok_or_else(|| {
                        Error::Bundle("Message data reference count overflow".to_owned())
                    })?;
                    referencing_objects
                        .entry(identifier)
                        .or_default()
                        .insert(object_identifier);
                }
            }
        }
    }
    if metadata_map_identifier.is_some() && metadata_map_payloads != 1 {
        return Err(Error::Bundle(format!(
            "Expected one DataMetadataMap payload, found {metadata_map_payloads}"
        )));
    }

    let mut assets = Vec::with_capacity(metadata.datas.len());
    let mut identifiers = std::collections::HashSet::with_capacity(metadata.datas.len());
    for data in metadata.datas {
        if !identifiers.insert(data.identifier) {
            return Err(Error::Bundle(format!(
                "Duplicate DataInfo identifier {}",
                data.identifier
            )));
        }
        let package_path = data
            .file_name
            .as_deref()
            .filter(|file_name| !file_name.is_empty())
            .map(data_entry_name)
            .transpose()?
            .filter(|path| package.contains_entry(path));
        let size = package_path
            .as_deref()
            .and_then(|path| package.entry(path))
            .map(|bytes| u64::try_from(bytes.len()))
            .transpose()
            .map_err(|_| Error::Bundle("Materialized asset length exceeds u64".to_owned()))?;
        let type_name = package_path
            .as_deref()
            .and_then(|path| Path::new(path).file_name())
            .and_then(|name| name.to_str())
            .unwrap_or(&data.preferred_file_name);
        let media_type = Path::new(type_name)
            .extension()
            .and_then(|extension| extension.to_str())
            .map(MediaType::from_extension)
            .unwrap_or(MediaType::Unknown);
        assets.push(EmbeddedMediaAsset {
            data_identifier: data.identifier,
            preferred_filename: data.preferred_file_name,
            package_path,
            media_type,
            size,
            declared_size: data.materialized_length,
            digest: data.digest,
            component_reference_count: component_counts.get(&data.identifier).copied().unwrap_or(0),
            component_reference_record_count: component_record_counts
                .get(&data.identifier)
                .copied()
                .unwrap_or(0),
            message_reference_count: message_counts.get(&data.identifier).copied().unwrap_or(0),
            has_data_metadata: data_metadata_ids.contains(&data.identifier),
            referencing_object_ids: {
                let mut identifiers = referencing_objects
                    .remove(&data.identifier)
                    .unwrap_or_default()
                    .into_iter()
                    .collect::<Vec<_>>();
                identifiers.sort_unstable();
                identifiers
            },
        });
    }
    assets.sort_unstable_by_key(|asset| asset.data_identifier);
    Ok(assets)
}

pub(crate) fn reachable_embedded_assets(
    package: &IWorkPackage,
    roots: impl IntoIterator<Item = u64>,
) -> Result<Vec<EmbeddedMediaAsset>> {
    let assets = embedded_assets(package)?;
    let mut outgoing = HashMap::<u64, Vec<u64>>::new();
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        for object in archive.objects {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("Object in {name} has no identifier"))
            })?;
            let references = outgoing.entry(identifier).or_default();
            for info in object.archive_info.message_infos {
                for reference in info.object_references {
                    if !references.contains(&reference) {
                        references.push(reference);
                    }
                }
            }
        }
    }

    let mut reachable = HashSet::new();
    let mut queue = roots.into_iter().collect::<VecDeque<_>>();
    while let Some(identifier) = queue.pop_front() {
        if !reachable.insert(identifier) {
            continue;
        }
        if let Some(references) = outgoing.get(&identifier) {
            queue.extend(references.iter().copied());
        }
    }
    Ok(assets
        .into_iter()
        .filter(|asset| {
            asset
                .referencing_object_ids
                .iter()
                .any(|identifier| reachable.contains(identifier))
        })
        .collect())
}

fn decode_package_metadata(package: &IWorkPackage) -> Result<protobuf::tsp::PackageMetadata> {
    let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let mut payload = None;
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                && payload.replace(message.data.as_slice()).is_some()
            {
                return Err(Error::Bundle(
                    "Package contains multiple PackageMetadata payloads".to_owned(),
                ));
            }
        }
    }
    protobuf::tsp::PackageMetadata::decode(
        payload.ok_or_else(|| Error::Bundle("PackageMetadata payload was not found".to_owned()))?,
    )
    .map_err(Into::into)
}

fn data_entry_name(file_name: &str) -> Result<String> {
    if file_name.is_empty() || file_name.contains(['\0', '\\']) {
        return Err(Error::Bundle(format!(
            "Unsafe DataInfo filename: {file_name:?}"
        )));
    }
    let path = Path::new(file_name);
    if path.is_absolute()
        || path
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(Error::Bundle(format!(
            "Unsafe DataInfo filename: {file_name:?}"
        )));
    }
    Ok(format!("Data/{file_name}"))
}

#[derive(Debug, Clone)]
struct WireField {
    number: u32,
    wire_type: u8,
    start: usize,
    key_end: usize,
    end: usize,
    payload: Option<Range<usize>>,
}

fn parse_wire_fields(data: &[u8]) -> Result<Vec<WireField>> {
    let mut fields = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let start = offset;
        let (key, key_length) = decode_varint_from_bytes(&data[offset..])
            .map_err(|error| Error::InvalidFormat(format!("Invalid protobuf key: {error}")))?;
        offset = offset
            .checked_add(key_length)
            .ok_or_else(|| Error::InvalidFormat("Protobuf key offset overflow".to_owned()))?;
        let number = key >> 3;
        if number == 0 || number > 0x1fff_ffff {
            return Err(Error::InvalidFormat(format!(
                "Invalid protobuf field number {number}"
            )));
        }
        let wire_type = (key & 7) as u8;
        let key_end = offset;
        let payload = match wire_type {
            0 => {
                let (_, length) = decode_varint_from_bytes(&data[offset..]).map_err(|error| {
                    Error::InvalidFormat(format!("Invalid protobuf varint value: {error}"))
                })?;
                offset = offset.checked_add(length).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf varint offset overflow".to_owned())
                })?;
                None
            },
            1 => {
                offset = offset.checked_add(8).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf fixed64 offset overflow".to_owned())
                })?;
                None
            },
            2 => {
                let (length, prefix_length) =
                    decode_varint_from_bytes(&data[offset..]).map_err(|error| {
                        Error::InvalidFormat(format!("Invalid protobuf length: {error}"))
                    })?;
                offset = offset.checked_add(prefix_length).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf length prefix overflow".to_owned())
                })?;
                let payload_start = offset;
                let length = usize::try_from(length).map_err(|_| {
                    Error::InvalidFormat("Protobuf field length exceeds usize".to_owned())
                })?;
                offset = offset.checked_add(length).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf field range overflow".to_owned())
                })?;
                Some(payload_start..offset)
            },
            5 => {
                offset = offset.checked_add(4).ok_or_else(|| {
                    Error::InvalidFormat("Protobuf fixed32 offset overflow".to_owned())
                })?;
                None
            },
            3 | 4 => {
                return Err(Error::InvalidFormat(
                    "Deprecated protobuf groups are not supported in PackageMetadata".to_owned(),
                ));
            },
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "Invalid protobuf wire type {wire_type}"
                )));
            },
        };
        if offset > data.len() {
            return Err(Error::InvalidFormat(
                "Truncated protobuf field in PackageMetadata".to_owned(),
            ));
        }
        fields.push(WireField {
            number: number as u32,
            wire_type,
            start,
            key_end,
            end: offset,
            payload,
        });
    }
    Ok(fields)
}

fn field_payload<'a>(data: &'a [u8], field: &WireField) -> Result<&'a [u8]> {
    let range = field.payload.clone().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Protobuf field {} is not length-delimited",
            field.number
        ))
    })?;
    data.get(range)
        .ok_or_else(|| Error::InvalidFormat("Protobuf payload range is invalid".to_owned()))
}

fn field_varint(data: &[u8], field: &WireField) -> Result<u64> {
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "Protobuf field {} is not a varint",
            field.number
        )));
    }
    decode_varint_from_bytes(
        data.get(field.key_end..field.end)
            .ok_or_else(|| Error::InvalidFormat("Protobuf varint range is invalid".to_owned()))?,
    )
    .map(|(value, _)| value)
    .map_err(|error| Error::InvalidFormat(format!("Invalid protobuf varint: {error}")))
}

fn data_info_identifier(data: &[u8]) -> Result<u64> {
    let fields = parse_wire_fields(data)?;
    let identifiers = fields
        .iter()
        .filter(|field| field.number == 1)
        .map(|field| field_varint(data, field))
        .collect::<Result<Vec<_>>>()?;
    match identifiers.as_slice() {
        [identifier] => Ok(*identifier),
        [] => Err(Error::InvalidFormat(
            "DataInfo is missing its required identifier".to_owned(),
        )),
        _ => Err(Error::InvalidFormat(
            "DataInfo contains duplicate identifiers".to_owned(),
        )),
    }
}

fn patch_package_metadata(
    metadata: &[u8],
    data_identifier: u64,
    digest: &[u8],
    materialized_length: u64,
) -> Result<Vec<u8>> {
    if digest.len() != 20 {
        return Err(Error::InvalidFormat(format!(
            "iWork materialized data digest must be SHA-1 (20 bytes), got {}",
            digest.len()
        )));
    }
    let fields = parse_wire_fields(metadata)?;
    let mut output = Vec::with_capacity(metadata.len());
    let mut patched_count = 0usize;
    for field in fields {
        if field.number == 4 {
            if field.wire_type != 2 {
                return Err(Error::InvalidFormat(
                    "PackageMetadata.datas has an invalid wire type".to_owned(),
                ));
            }
            let data_info = field_payload(metadata, &field)?;
            if data_info_identifier(data_info)? == data_identifier {
                patched_count = patched_count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Patched DataInfo count overflow".to_owned())
                })?;
                let patched = patch_data_info(data_info, digest, materialized_length)?;
                output.extend_from_slice(&metadata[field.start..field.key_end]);
                output.extend(encode_varint(patched.len() as u64));
                output.extend_from_slice(&patched);
                continue;
            }
        }
        output.extend_from_slice(&metadata[field.start..field.end]);
    }
    match patched_count {
        1 => {},
        0 => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is absent from PackageMetadata"
            )));
        },
        _ => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is duplicated in PackageMetadata"
            )));
        },
    }
    let decoded = protobuf::tsp::PackageMetadata::decode(output.as_slice())?;
    let matches = decoded
        .datas
        .iter()
        .filter(|data| data.identifier == data_identifier)
        .collect::<Vec<_>>();
    if matches.len() != 1
        || matches[0].digest != digest
        || matches[0].materialized_length != Some(materialized_length)
    {
        return Err(Error::InvalidFormat(
            "Patched PackageMetadata did not decode to the requested values".to_owned(),
        ));
    }
    Ok(output)
}

fn append_data_info(
    metadata: &[u8],
    data_identifier: u64,
    digest: &[u8],
    preferred_filename: &str,
    file_name: &str,
    materialized_length: u64,
) -> Result<Vec<u8>> {
    if digest.len() != 20 {
        return Err(Error::InvalidFormat(format!(
            "iWork materialized data digest must be SHA-1 (20 bytes), got {}",
            digest.len()
        )));
    }
    let decoded = protobuf::tsp::PackageMetadata::decode(metadata)?;
    if decoded
        .datas
        .iter()
        .any(|data| data.identifier == data_identifier)
    {
        return Err(Error::Bundle(format!(
            "Data identifier {data_identifier} already exists"
        )));
    }

    let mut data_info = Vec::new();
    append_wire_varint(&mut data_info, 1, data_identifier);
    append_wire_bytes(&mut data_info, 2, digest);
    append_wire_bytes(&mut data_info, 3, preferred_filename.as_bytes());
    append_wire_bytes(&mut data_info, 4, file_name.as_bytes());
    append_wire_varint(&mut data_info, 18, materialized_length);

    // Appending a repeated field is protobuf-canonical and avoids rewriting any
    // pre-existing metadata field, including unknown extensions.
    let mut output = Vec::with_capacity(metadata.len() + data_info.len() + 16);
    output.extend_from_slice(metadata);
    append_wire_bytes(&mut output, 4, &data_info);
    let verified = protobuf::tsp::PackageMetadata::decode(output.as_slice())?;
    let inserted = verified
        .datas
        .iter()
        .filter(|data| data.identifier == data_identifier)
        .collect::<Vec<_>>();
    if inserted.len() != 1
        || inserted[0].digest != digest
        || inserted[0].preferred_file_name != preferred_filename
        || inserted[0].file_name.as_deref() != Some(file_name)
        || inserted[0].materialized_length != Some(materialized_length)
    {
        return Err(Error::InvalidFormat(
            "Appended DataInfo did not decode to the requested values".to_owned(),
        ));
    }
    Ok(output)
}

fn append_wire_varint(output: &mut Vec<u8>, field_number: u64, value: u64) {
    output.extend(encode_varint(field_number << 3));
    output.extend(encode_varint(value));
}

fn append_wire_bytes(output: &mut Vec<u8>, field_number: u64, value: &[u8]) {
    output.extend(encode_varint((field_number << 3) | 2));
    output.extend(encode_varint(value.len() as u64));
    output.extend_from_slice(value);
}

fn remove_data_info(metadata: &[u8], data_identifier: u64) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(metadata)?;
    let mut output = Vec::with_capacity(metadata.len());
    let mut removed_count = 0usize;
    for field in fields {
        if field.number == 4 {
            if field.wire_type != 2 {
                return Err(Error::InvalidFormat(
                    "PackageMetadata.datas has an invalid wire type".to_owned(),
                ));
            }
            if data_info_identifier(field_payload(metadata, &field)?)? == data_identifier {
                removed_count = removed_count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Removed DataInfo count overflow".to_owned())
                })?;
                continue;
            }
        }
        output.extend_from_slice(&metadata[field.start..field.end]);
    }
    match removed_count {
        1 => {},
        0 => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is absent from PackageMetadata"
            )));
        },
        _ => {
            return Err(Error::Bundle(format!(
                "Data identifier {data_identifier} is duplicated in PackageMetadata"
            )));
        },
    }
    let decoded = protobuf::tsp::PackageMetadata::decode(output.as_slice())?;
    if decoded
        .datas
        .iter()
        .any(|data| data.identifier == data_identifier)
    {
        return Err(Error::InvalidFormat(
            "Removed DataInfo still decodes from PackageMetadata".to_owned(),
        ));
    }
    Ok(output)
}

fn patch_data_info(data: &[u8], digest: &[u8], materialized_length: u64) -> Result<Vec<u8>> {
    let fields = parse_wire_fields(data)?;
    let mut output = Vec::with_capacity(data.len());
    let mut digest_count = 0usize;
    let mut length_count = 0usize;
    for field in fields {
        match field.number {
            2 => {
                if field.wire_type != 2 {
                    return Err(Error::InvalidFormat(
                        "DataInfo.digest has an invalid wire type".to_owned(),
                    ));
                }
                digest_count += 1;
                if digest_count > 1 {
                    return Err(Error::InvalidFormat(
                        "DataInfo contains duplicate digests".to_owned(),
                    ));
                }
                output.extend_from_slice(&data[field.start..field.key_end]);
                output.extend(encode_varint(digest.len() as u64));
                output.extend_from_slice(digest);
            },
            18 => {
                if field.wire_type != 0 {
                    return Err(Error::InvalidFormat(
                        "DataInfo.materialized_length has an invalid wire type".to_owned(),
                    ));
                }
                length_count += 1;
                if length_count > 1 {
                    return Err(Error::InvalidFormat(
                        "DataInfo contains duplicate materialized lengths".to_owned(),
                    ));
                }
                output.extend_from_slice(&data[field.start..field.key_end]);
                output.extend(encode_varint(materialized_length));
            },
            _ => output.extend_from_slice(&data[field.start..field.end]),
        }
    }
    if digest_count == 0 {
        output.extend(encode_varint((2 << 3) | 2));
        output.extend(encode_varint(digest.len() as u64));
        output.extend_from_slice(digest);
    }
    if length_count == 0 {
        output.extend(encode_varint(18 << 3));
        output.extend(encode_varint(materialized_length));
    }
    Ok(output)
}

/// Statistics about materialized media assets.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MediaStats {
    pub total_count: usize,
    pub total_size: u64,
    pub image_count: usize,
    pub video_count: usize,
    pub audio_count: usize,
    pub pdf_count: usize,
    pub unknown_count: usize,
}

impl MediaStats {
    pub fn total_size_human(&self) -> String {
        format_bytes(self.total_size)
    }

    pub fn summary(&self) -> String {
        format!(
            "{} files ({}) - {} images, {} videos, {} audio, {} PDFs",
            self.total_count,
            self.total_size_human(),
            self.image_count,
            self.video_count,
            self.audio_count,
            self.pdf_count
        )
    }
}

fn format_bytes(bytes: u64) -> String {
    const UNITS: &[&str] = &["B", "KB", "MB", "GB", "TB"];
    let mut size = bytes as f64;
    let mut unit = 0;
    while size >= 1024.0 && unit < UNITS.len() - 1 {
        size /= 1024.0;
        unit += 1;
    }
    format!("{size:.2} {}", UNITS[unit])
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject};

    const PNG: &[u8] = b"\x89PNG\r\n\x1a\nreplacement";

    fn append_varint_field(output: &mut Vec<u8>, number: u64, value: u64) {
        output.extend(encode_varint(number << 3));
        output.extend(encode_varint(value));
    }

    fn append_bytes_field(output: &mut Vec<u8>, number: u64, value: &[u8]) {
        output.extend(encode_varint((number << 3) | 2));
        output.extend(encode_varint(value.len() as u64));
        output.extend_from_slice(value);
    }

    fn synthetic_metadata(data_bytes: &[u8]) -> Vec<u8> {
        let mut data_info = Vec::new();
        append_varint_field(&mut data_info, 1, 7);
        append_bytes_field(&mut data_info, 2, &[0x11; 20]);
        append_bytes_field(&mut data_info, 3, b"image.png");
        append_bytes_field(&mut data_info, 4, b"image-7.png");
        // DataAttributes is an empty generated message, but Apple writes
        // extension fields inside it. This payload must remain byte-exact.
        append_bytes_field(&mut data_info, 10, &[0x08, 0x96, 0x01, 0x12, 0x01, 0xff]);
        append_varint_field(&mut data_info, 18, data_bytes.len() as u64);

        let mut metadata = Vec::new();
        append_varint_field(&mut metadata, 1, 100);
        append_bytes_field(&mut metadata, 4, &data_info);
        append_bytes_field(&mut metadata, 100, b"outer-unknown");
        metadata
    }

    fn synthetic_package() -> IWorkPackage {
        let original = b"\x89PNG\r\n\x1a\noriginal";
        let mut metadata_object = ArchiveObject::new(
            2,
            vec![RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: synthetic_metadata(original),
            }],
        )
        .unwrap();
        metadata_object.archive_info.message_infos[0].data_references = Vec::new();
        let metadata_archive = Archive {
            objects: vec![metadata_object],
        };

        let mut document_object = ArchiveObject::new(
            50,
            vec![RawMessage {
                type_: 999,
                data: vec![1],
            }],
        )
        .unwrap();
        document_object.archive_info.message_infos[0].data_references = vec![7];
        let document_archive = Archive {
            objects: vec![document_object],
        };

        let mut package = IWorkPackage::new();
        package
            .replace_archive(PACKAGE_METADATA_ENTRY, &metadata_archive)
            .unwrap();
        package
            .replace_archive("Index/Document.iwa", &document_archive)
            .unwrap();
        package
            .insert_entry("Data/image-7.png", original.to_vec())
            .unwrap();
        package
    }

    fn nested_field_bytes(metadata: &[u8], field_number: u32) -> Vec<u8> {
        let outer = parse_wire_fields(metadata).unwrap();
        let data_info = outer
            .iter()
            .find(|field| field.number == 4)
            .map(|field| field_payload(metadata, field).unwrap())
            .unwrap();
        let nested = parse_wire_fields(data_info).unwrap();
        let field = nested
            .iter()
            .find(|field| field.number == field_number)
            .unwrap();
        data_info[field.start..field.end].to_vec()
    }

    #[test]
    fn detects_extensions_and_signatures() {
        assert_eq!(MediaType::from_extension("PNG"), MediaType::Image);
        assert_eq!(MediaType::from_extension("mov"), MediaType::Video);
        assert_eq!(MediaType::from_extension("m4a"), MediaType::Audio);
        assert_eq!(MediaType::from_extension("pdf"), MediaType::Pdf);
        assert_eq!(MediaType::from_bytes(PNG), MediaType::Image);
        assert_eq!(MediaType::from_bytes(b"%PDF-1.7\n"), MediaType::Pdf);
        assert_eq!(MediaType::from_bytes(b"ID3\x04\0\0"), MediaType::Audio);
    }

    #[test]
    fn manager_reads_single_file_and_memory_packages() {
        let package = synthetic_package();
        let bytes = package.to_bytes().unwrap();
        let memory = MediaManager::from_bytes(&bytes).unwrap();
        assert_eq!(memory.assets().len(), 1);
        assert!(memory.get("image-7.png").unwrap().is_image());
        assert_eq!(
            memory.extract("image-7.png").unwrap(),
            b"\x89PNG\r\n\x1a\noriginal"
        );

        let file = tempfile::NamedTempFile::new().unwrap();
        fs::write(file.path(), bytes).unwrap();
        let disk = MediaManager::new(file.path()).unwrap();
        assert_eq!(
            disk.extract("image-7.png").unwrap(),
            memory.extract("image-7.png").unwrap()
        );
    }

    #[test]
    fn replaces_asset_and_preserves_unknown_metadata_fields() {
        let mut editor = IWorkMediaEditor::from_package(synthetic_package()).unwrap();
        let before_archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
        let before = &before_archive.object(2).unwrap().messages[0].data;
        let attributes_before = nested_field_bytes(before, 10);
        let outer_unknown_before = parse_wire_fields(before)
            .unwrap()
            .into_iter()
            .find(|field| field.number == 100)
            .map(|field| before[field.start..field.end].to_vec())
            .unwrap();

        let previous = editor.replace(7, PNG).unwrap();
        assert_eq!(previous, b"\x89PNG\r\n\x1a\noriginal");
        assert_eq!(editor.extract(7).unwrap(), PNG);
        let asset = editor.asset(7).unwrap();
        assert_eq!(asset.digest, Sha1::digest(PNG).to_vec());
        assert_eq!(asset.declared_size, Some(PNG.len() as u64));
        assert_eq!(asset.message_reference_count, 1);

        let after_archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
        let after = &after_archive.object(2).unwrap().messages[0].data;
        assert_eq!(nested_field_bytes(after, 10), attributes_before);
        let outer_unknown_after = parse_wire_fields(after)
            .unwrap()
            .into_iter()
            .find(|field| field.number == 100)
            .map(|field| after[field.start..field.end].to_vec())
            .unwrap();
        assert_eq!(outer_unknown_after, outer_unknown_before);
    }

    #[test]
    fn rejects_type_mismatch_transactionally() {
        let mut editor = IWorkMediaEditor::from_package(synthetic_package()).unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(editor.replace(7, b"%PDF-1.7\nnot-an-image").is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn removes_only_unreferenced_assets_transactionally() {
        let mut referenced = IWorkMediaEditor::from_package(synthetic_package()).unwrap();
        assert!(referenced.remove_unreferenced(7).is_err());
        assert!(referenced.asset(7).is_some());

        let mut package = synthetic_package();
        package
            .update_archive("Index/Document.iwa", |archive| {
                archive.objects[0].archive_info.message_infos[0]
                    .data_references
                    .clear();
                Ok(())
            })
            .unwrap();
        let before_archive = package.archive(PACKAGE_METADATA_ENTRY).unwrap();
        let before = &before_archive.object(2).unwrap().messages[0].data;
        let outer_unknown_before = parse_wire_fields(before)
            .unwrap()
            .into_iter()
            .find(|field| field.number == 100)
            .map(|field| before[field.start..field.end].to_vec())
            .unwrap();

        let mut editor = IWorkMediaEditor::from_package(package).unwrap();
        let removed = editor.remove_unreferenced(7).unwrap().unwrap();
        assert_eq!(removed, b"\x89PNG\r\n\x1a\noriginal");
        assert!(editor.asset(7).is_none());
        assert!(!editor.package().contains_entry("Data/image-7.png"));

        let after_archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
        let after = &after_archive.object(2).unwrap().messages[0].data;
        let outer_unknown_after = parse_wire_fields(after)
            .unwrap()
            .into_iter()
            .find(|field| field.number == 100)
            .map(|field| after[field.start..field.end].to_vec())
            .unwrap();
        assert_eq!(outer_unknown_after, outer_unknown_before);
    }

    #[test]
    fn inserts_and_removes_unreferenced_asset_without_metadata_drift() {
        let package = synthetic_package();
        let initial_metadata = package
            .archive(PACKAGE_METADATA_ENTRY)
            .unwrap()
            .object(2)
            .unwrap()
            .messages[0]
            .data
            .clone();
        let mut editor = IWorkMediaEditor::from_package(package).unwrap();
        let inserted = editor.insert_unreferenced("new.png", PNG).unwrap();
        assert_eq!(inserted.data_identifier, 8);
        assert_eq!(inserted.package_path.as_deref(), Some("Data/new-8.png"));
        assert!(!inserted.is_referenced());
        assert_eq!(editor.extract(8).unwrap(), PNG);

        assert_eq!(editor.remove_unreferenced(8).unwrap().unwrap(), PNG);
        let final_metadata = editor
            .package()
            .archive(PACKAGE_METADATA_ENTRY)
            .unwrap()
            .object(2)
            .unwrap()
            .messages[0]
            .data
            .clone();
        assert_eq!(final_metadata, initial_metadata);
        assert!(editor.asset(8).is_none());
    }

    #[test]
    fn wire_parser_rejects_truncation_and_groups() {
        assert!(parse_wire_fields(&[0x12, 0x05, 1]).is_err());
        assert!(parse_wire_fields(&[0x0b]).is_err());
        assert!(parse_wire_fields(&[0x80]).is_err());
    }

    #[test]
    fn formats_sizes() {
        assert_eq!(format_bytes(0), "0.00 B");
        assert_eq!(format_bytes(1024), "1.00 KB");
        assert_eq!(format_bytes(1536 * 1024), "1.50 MB");
    }
}
