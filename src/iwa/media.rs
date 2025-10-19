//! Media Asset Management for iWork Documents
//!
//! iWork documents store media assets (images, videos, audio) in the Data/
//! directory within the bundle. This module provides utilities for extracting
//! and managing these media files.

use std::collections::HashMap;
use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};

use crate::iwa::{Error, Result};

/// Types of media assets that can be found in iWork documents
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MediaType {
    /// Image file (PNG, JPEG, TIFF, etc.)
    Image,
    /// Video file (MP4, MOV, etc.)
    Video,
    /// Audio file (MP3, AAC, WAV, etc.)
    Audio,
    /// PDF document
    Pdf,
    /// Unknown or unsupported media type
    Unknown,
}

impl MediaType {
    /// Detect media type from file extension
    pub fn from_extension(ext: &str) -> Self {
        match ext.to_lowercase().as_str() {
            "png" | "jpg" | "jpeg" | "gif" | "tiff" | "tif" | "bmp" | "heic" | "heif" => {
                MediaType::Image
            }
            "mp4" | "mov" | "m4v" | "avi" | "mkv" => MediaType::Video,
            "mp3" | "aac" | "m4a" | "wav" | "aiff" => MediaType::Audio,
            "pdf" => MediaType::Pdf,
            _ => MediaType::Unknown,
        }
    }

    /// Get a human-readable name for this media type
    pub fn name(&self) -> &'static str {
        match self {
            MediaType::Image => "Image",
            MediaType::Video => "Video",
            MediaType::Audio => "Audio",
            MediaType::Pdf => "PDF Document",
            MediaType::Unknown => "Unknown",
        }
    }
}

/// Information about a media asset
#[derive(Debug, Clone)]
pub struct MediaAsset {
    /// Relative path within the bundle
    pub path: PathBuf,
    /// Media type
    pub media_type: MediaType,
    /// File size in bytes
    pub size: u64,
    /// File name without path
    pub filename: String,
}

impl MediaAsset {
    /// Create a new media asset entry
    pub fn new(path: PathBuf, size: u64) -> Self {
        let filename = path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("unknown")
            .to_string();

        let media_type = path
            .extension()
            .and_then(|e| e.to_str())
            .map(MediaType::from_extension)
            .unwrap_or(MediaType::Unknown);

        Self {
            path,
            media_type,
            size,
            filename,
        }
    }

    /// Check if this is an image asset
    pub fn is_image(&self) -> bool {
        self.media_type == MediaType::Image
    }

    /// Check if this is a video asset
    pub fn is_video(&self) -> bool {
        self.media_type == MediaType::Video
    }

    /// Check if this is an audio asset
    pub fn is_audio(&self) -> bool {
        self.media_type == MediaType::Audio
    }
}

/// Manager for media assets in an iWork bundle
#[derive(Debug, Clone)]
pub struct MediaManager {
    /// Bundle root path
    bundle_path: PathBuf,
    /// Map of media assets by filename
    assets: HashMap<String, MediaAsset>,
}

impl MediaManager {
    /// Create a new media manager for a bundle
    pub fn new<P: AsRef<Path>>(bundle_path: P) -> Result<Self> {
        let bundle_path = bundle_path.as_ref().to_path_buf();
        let mut assets = HashMap::new();

        // Check if this is a directory bundle
        if bundle_path.is_dir() {
            Self::scan_directory_bundle(&bundle_path, &mut assets)?;
        }

        Ok(Self {
            bundle_path,
            assets,
        })
    }

    /// Scan a directory bundle for media assets
    fn scan_directory_bundle(bundle_path: &Path, assets: &mut HashMap<String, MediaAsset>) -> Result<()> {
        let data_dir = bundle_path.join("Data");
        if !data_dir.exists() || !data_dir.is_dir() {
            return Ok(()); // No Data directory is not an error
        }

        Self::scan_directory_recursive(&data_dir, bundle_path, assets)?;
        Ok(())
    }

    /// Recursively scan a directory for media files
    fn scan_directory_recursive(
        dir: &Path,
        bundle_root: &Path,
        assets: &mut HashMap<String, MediaAsset>,
    ) -> Result<()> {
        let entries = fs::read_dir(dir).map_err(Error::Io)?;

        for entry in entries {
            let entry = entry.map_err(Error::Io)?;
            let path = entry.path();

            if path.is_dir() {
                Self::scan_directory_recursive(&path, bundle_root, assets)?;
            } else if path.is_file() {
                if let Ok(metadata) = fs::metadata(&path) {
                    let relative_path = path
                        .strip_prefix(bundle_root)
                        .unwrap_or(&path)
                        .to_path_buf();

                    let asset = MediaAsset::new(relative_path.clone(), metadata.len());
                    let filename = asset.filename.clone();
                    assets.insert(filename, asset);
                }
            }
        }

        Ok(())
    }

    /// Get all media assets
    pub fn assets(&self) -> &HashMap<String, MediaAsset> {
        &self.assets
    }

    /// Get a media asset by filename
    pub fn get(&self, filename: &str) -> Option<&MediaAsset> {
        self.assets.get(filename)
    }

    /// Get all assets of a specific type
    pub fn assets_by_type(&self, media_type: MediaType) -> Vec<&MediaAsset> {
        self.assets
            .values()
            .filter(|asset| asset.media_type == media_type)
            .collect()
    }

    /// Get all image assets
    pub fn images(&self) -> Vec<&MediaAsset> {
        self.assets_by_type(MediaType::Image)
    }

    /// Get all video assets
    pub fn videos(&self) -> Vec<&MediaAsset> {
        self.assets_by_type(MediaType::Video)
    }

    /// Get all audio assets
    pub fn audio(&self) -> Vec<&MediaAsset> {
        self.assets_by_type(MediaType::Audio)
    }

    /// Extract a media asset to a byte vector
    pub fn extract(&self, filename: &str) -> Result<Vec<u8>> {
        let asset = self
            .get(filename)
            .ok_or_else(|| Error::Bundle(format!("Media asset not found: {}", filename)))?;

        let full_path = self.bundle_path.join(&asset.path);
        let mut file = fs::File::open(&full_path).map_err(Error::Io)?;
        let mut data = Vec::new();
        file.read_to_end(&mut data).map_err(Error::Io)?;
        Ok(data)
    }

    /// Extract a media asset to a file
    pub fn extract_to_file(&self, filename: &str, output_path: &Path) -> Result<()> {
        let data = self.extract(filename)?;
        fs::write(output_path, data).map_err(Error::Io)?;
        Ok(())
    }

    /// Get media statistics
    pub fn stats(&self) -> MediaStats {
        let mut stats = MediaStats {
            total_count: self.assets.len(),
            total_size: 0,
            image_count: 0,
            video_count: 0,
            audio_count: 0,
            pdf_count: 0,
            unknown_count: 0,
        };

        for asset in self.assets.values() {
            stats.total_size += asset.size;
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

/// Statistics about media assets in a bundle
#[derive(Debug, Clone, Default)]
pub struct MediaStats {
    /// Total number of media assets
    pub total_count: usize,
    /// Total size of all media assets in bytes
    pub total_size: u64,
    /// Number of image files
    pub image_count: usize,
    /// Number of video files
    pub video_count: usize,
    /// Number of audio files
    pub audio_count: usize,
    /// Number of PDF files
    pub pdf_count: usize,
    /// Number of unknown/unsupported files
    pub unknown_count: usize,
}

impl MediaStats {
    /// Format the total size as a human-readable string
    pub fn total_size_human(&self) -> String {
        format_bytes(self.total_size)
    }

    /// Get a summary string of the media statistics
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

/// Format a byte count as a human-readable string
fn format_bytes(bytes: u64) -> String {
    const UNITS: &[&str] = &["B", "KB", "MB", "GB", "TB"];
    let mut size = bytes as f64;
    let mut unit_index = 0;

    while size >= 1024.0 && unit_index < UNITS.len() - 1 {
        size /= 1024.0;
        unit_index += 1;
    }

    format!("{:.2} {}", size, UNITS[unit_index])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_media_type_detection() {
        assert_eq!(MediaType::from_extension("png"), MediaType::Image);
        assert_eq!(MediaType::from_extension("PNG"), MediaType::Image);
        assert_eq!(MediaType::from_extension("jpg"), MediaType::Image);
        assert_eq!(MediaType::from_extension("jpeg"), MediaType::Image);
        assert_eq!(MediaType::from_extension("mp4"), MediaType::Video);
        assert_eq!(MediaType::from_extension("mov"), MediaType::Video);
        assert_eq!(MediaType::from_extension("mp3"), MediaType::Audio);
        assert_eq!(MediaType::from_extension("pdf"), MediaType::Pdf);
        assert_eq!(MediaType::from_extension("unknown"), MediaType::Unknown);
    }

    #[test]
    fn test_format_bytes() {
        assert_eq!(format_bytes(0), "0.00 B");
        assert_eq!(format_bytes(1024), "1.00 KB");
        assert_eq!(format_bytes(1024 * 1024), "1.00 MB");
        assert_eq!(format_bytes(1536 * 1024), "1.50 MB");
    }

    #[test]
    fn test_media_asset_creation() {
        let path = PathBuf::from("Data/image.png");
        let asset = MediaAsset::new(path, 1024);

        assert_eq!(asset.filename, "image.png");
        assert_eq!(asset.media_type, MediaType::Image);
        assert_eq!(asset.size, 1024);
        assert!(asset.is_image());
        assert!(!asset.is_video());
    }

    #[test]
    fn test_media_manager_with_pages() {
        let bundle_path = std::path::Path::new("test.pages");
        if !bundle_path.exists() {
            return; // Skip test if file doesn't exist
        }

        let manager = MediaManager::new(bundle_path);
        if let Ok(manager) = manager {
            let stats = manager.stats();
            println!("Media stats: {}", stats.summary());

            // Check if we found any media
            if stats.total_count > 0 {
                println!("Found {} media assets", stats.total_count);
                for (filename, asset) in manager.assets() {
                    println!("  - {}: {} ({})", filename, asset.media_type.name(), format_bytes(asset.size));
                }
            }
        }
    }
}

