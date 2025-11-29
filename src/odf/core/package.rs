//! ODF package (ZIP archive) handling functionality.
//!
//! This module provides utilities for working with ODF files as ZIP archives,
//! including reading files, checking existence, and basic package operations.
//!
//! Uses soapberry-zip for high-performance zero-copy ZIP parsing.

use crate::common::{Error, Result};
use soapberry_zip::office::ArchiveReader;
use std::io::Read;

/// An ODF package (ZIP file containing XML documents)
///
/// Uses soapberry-zip for efficient lazy decompression.
pub struct Package<'data> {
    archive: ArchiveReader<'data>,
    #[allow(dead_code)]
    manifest: super::manifest::Manifest,
    mimetype: String,
}

/// Owned version of Package that owns the data buffer.
pub struct OwnedPackage {
    data: Vec<u8>,
}

#[allow(dead_code)]
impl OwnedPackage {
    /// Open an ODF package from a reader
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;

        // Validate the archive can be parsed
        let _ = ArchiveReader::new(&data)
            .map_err(|_| Error::InvalidFormat("Invalid ZIP archive".to_string()))?;

        Ok(Self { data })
    }

    /// Create an ODF package from bytes
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        // Validate the archive can be parsed
        let _ = ArchiveReader::new(&data)
            .map_err(|_| Error::InvalidFormat("Invalid ZIP archive".to_string()))?;

        Ok(Self { data })
    }

    /// Get a borrowed Package for accessing archive contents
    pub fn package(&self) -> Result<Package<'_>> {
        Package::new(&self.data)
    }

    /// Get the underlying data
    pub fn into_inner(self) -> Vec<u8> {
        self.data
    }

    /// Get a reference to the underlying data
    pub fn as_bytes(&self) -> &[u8] {
        &self.data
    }

    // Convenience methods that delegate to Package

    /// Get the MIME type from the mimetype file
    pub fn mimetype(&self) -> Result<String> {
        let package = self.package()?;
        Ok(package.mimetype().to_string())
    }

    /// Get a file from the package by path
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let package = self.package()?;
        package.get_file(path)
    }

    /// Check if a file exists in the package
    pub fn has_file(&self, path: &str) -> Result<bool> {
        let package = self.package()?;
        Ok(package.has_file(path))
    }

    /// List all files in the package
    pub fn files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.files()
    }

    /// Get all embedded media files from the package.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.media_files()
    }
}

impl<'data> Package<'data> {
    /// Create a new Package from a byte slice
    pub fn new(data: &'data [u8]) -> Result<Self> {
        let archive = ArchiveReader::new(data)
            .map_err(|_| Error::InvalidFormat("Invalid ZIP archive".to_string()))?;

        // Read MIME type from mimetype file
        let mimetype = archive
            .read_string("mimetype")
            .map_err(|_| Error::InvalidFormat("No mimetype file found in ODF package".to_string()))?
            .trim()
            .to_string();

        // Parse the manifest
        let manifest = super::manifest::Manifest::from_archive_reader(&archive)?;

        Ok(Self {
            archive,
            manifest,
            mimetype,
        })
    }

    /// Get the MIME type from the mimetype file
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Get a file from the package by path
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        self.archive
            .read(path)
            .map_err(|_| Error::InvalidFormat(format!("File not found: {}", path)))
    }

    /// Check if a file exists in the package
    pub fn has_file(&self, path: &str) -> bool {
        self.archive.contains(path)
    }

    /// Get the manifest
    #[allow(dead_code)]
    pub fn manifest(&self) -> &super::manifest::Manifest {
        &self.manifest
    }

    /// List all files in the package
    pub fn files(&self) -> Result<Vec<String>> {
        Ok(self.archive.file_names().map(String::from).collect())
    }

    /// Get all embedded media files (images, etc.) from the package.
    ///
    /// This returns paths to all files in the Pictures/ directory and other media directories.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let all_files = self.files()?;
        Ok(all_files
            .into_iter()
            .filter(|path| {
                path.starts_with("Pictures/")
                    || path.starts_with("media/")
                    || path.starts_with("Object/")
                    || path.ends_with(".png")
                    || path.ends_with(".jpg")
                    || path.ends_with(".jpeg")
                    || path.ends_with(".gif")
                    || path.ends_with(".svg")
            })
            .collect())
    }

    /// Check if the package contains any media files.
    #[allow(dead_code)] // Reserved for future use
    pub fn has_media(&self) -> bool {
        self.media_files().map(|m| !m.is_empty()).unwrap_or(false)
    }
}
