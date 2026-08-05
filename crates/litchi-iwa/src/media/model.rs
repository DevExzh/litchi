//! Semantic media models and resource profiles.

use std::path::PathBuf;

use crate::{Error, Result};

const DEFAULT_MAX_MEDIA_ASSETS: usize = 100_000;
const DEFAULT_MAX_MEDIA_ASSET_BYTES: u64 = 512 * 1024 * 1024;
const DEFAULT_MAX_MEDIA_TOTAL_BYTES: u64 = 2 * 1024 * 1024 * 1024;

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

/// Resource ceilings for media discovery and extraction.
///
/// These limits apply to directory bundles as well as package-backed media.
/// The checked constructor only permits tighter profiles than the format-wide
/// safety ceilings, so callers cannot accidentally disable the allocation
/// guardrails while selecting a smaller workload budget.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MediaLimits {
    pub(super) max_assets: usize,
    pub(super) max_asset_bytes: u64,
    pub(super) max_total_bytes: u64,
}

impl MediaLimits {
    /// Hard ceiling for the number of discovered media members.
    pub const HARD_MAX_ASSETS: usize = DEFAULT_MAX_MEDIA_ASSETS;
    /// Hard ceiling for one materialized media member.
    pub const HARD_MAX_ASSET_BYTES: u64 = DEFAULT_MAX_MEDIA_ASSET_BYTES;
    /// Hard ceiling for the aggregate discovered media size.
    pub const HARD_MAX_TOTAL_BYTES: u64 = DEFAULT_MAX_MEDIA_TOTAL_BYTES;

    /// Construct a checked media resource profile.
    pub fn new(max_assets: usize, max_asset_bytes: u64, max_total_bytes: u64) -> Result<Self> {
        if max_assets == 0 || max_asset_bytes == 0 || max_total_bytes == 0 {
            return Err(Error::Bundle("Media limits must be non-zero".to_owned()));
        }
        if max_assets > Self::HARD_MAX_ASSETS {
            return Err(Error::Bundle(format!(
                "Media asset-count limit exceeds {} entries",
                Self::HARD_MAX_ASSETS
            )));
        }
        if max_asset_bytes > Self::HARD_MAX_ASSET_BYTES {
            return Err(Error::Bundle(format!(
                "Media member limit exceeds {} bytes",
                Self::HARD_MAX_ASSET_BYTES
            )));
        }
        if max_total_bytes > Self::HARD_MAX_TOTAL_BYTES {
            return Err(Error::Bundle(format!(
                "Media total-size limit exceeds {} bytes",
                Self::HARD_MAX_TOTAL_BYTES
            )));
        }
        Ok(Self {
            max_assets,
            max_asset_bytes,
            max_total_bytes,
        })
    }

    /// Maximum number of assets this profile may retain.
    pub const fn max_assets(self) -> usize {
        self.max_assets
    }

    /// Maximum size of one asset this profile may retain or extract.
    pub const fn max_asset_bytes(self) -> u64 {
        self.max_asset_bytes
    }

    /// Maximum aggregate size of all discovered assets.
    pub const fn max_total_bytes(self) -> u64 {
        self.max_total_bytes
    }
}

impl Default for MediaLimits {
    fn default() -> Self {
        Self {
            max_assets: DEFAULT_MAX_MEDIA_ASSETS,
            max_asset_bytes: DEFAULT_MAX_MEDIA_ASSET_BYTES,
            max_total_bytes: DEFAULT_MAX_MEDIA_TOTAL_BYTES,
        }
    }
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

pub(crate) fn format_bytes(bytes: u64) -> String {
    const UNITS: &[&str] = &["B", "KB", "MB", "GB", "TB"];
    let mut size = bytes as f64;
    let mut unit = 0;
    while size >= 1024.0 && unit < UNITS.len() - 1 {
        size /= 1024.0;
        unit += 1;
    }
    format!("{size:.2} {}", UNITS[unit])
}
