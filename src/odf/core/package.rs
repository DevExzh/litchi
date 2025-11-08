//! ODF package (ZIP archive) handling functionality.
//!
//! This module provides utilities for working with ODF files as ZIP archives,
//! including reading files, checking existence, and basic package operations.

use crate::common::{Error, Result};
use std::cell::RefCell;
use std::io::{Read, Seek};

/// An ODF package (ZIP file containing XML documents)
pub struct Package<R> {
    archive: RefCell<zip::ZipArchive<R>>,
    #[allow(dead_code)]
    manifest: super::manifest::Manifest,
    mimetype: String,
}

impl<R: Read + Seek> Package<R> {
    /// Open an ODF package from a reader
    pub fn from_reader(reader: R) -> Result<Self> {
        let mut archive = zip::ZipArchive::new(reader)
            .map_err(|_| Error::InvalidFormat("Invalid ZIP archive".to_string()))?;

        // Read MIME type from mimetype file
        let mimetype = Self::read_mimetype(&mut archive)?;

        // Parse the manifest
        let manifest = super::manifest::Manifest::from_archive(&mut archive)?;

        Ok(Self {
            archive: RefCell::new(archive),
            manifest,
            mimetype,
        })
    }

    /// Create an ODF package from an already-parsed ZIP archive.
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been parsed during format detection. It avoids double-parsing.
    pub fn from_zip_archive(mut archive: zip::ZipArchive<R>) -> Result<Self> {
        // Read MIME type from mimetype file
        let mimetype = Self::read_mimetype(&mut archive)?;

        // Parse the manifest
        let manifest = super::manifest::Manifest::from_archive(&mut archive)?;

        Ok(Self {
            archive: RefCell::new(archive),
            manifest,
            mimetype,
        })
    }

    /// Read MIME type from the mimetype file
    fn read_mimetype(archive: &mut zip::ZipArchive<R>) -> Result<String> {
        let mut mimetype_file = archive.by_name("mimetype").map_err(|_| {
            Error::InvalidFormat("No mimetype file found in ODF package".to_string())
        })?;

        let mut content = String::new();
        mimetype_file.read_to_string(&mut content)?;
        Ok(content.trim().to_string())
    }

    /// Get the MIME type from the mimetype file
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Get a file from the package by path
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let mut archive = self.archive.borrow_mut();
        let mut file = archive
            .by_name(path)
            .map_err(|_| Error::InvalidFormat(format!("File not found: {}", path)))?;

        let mut content = Vec::new();
        file.read_to_end(&mut content)?;
        Ok(content)
    }

    /// Check if a file exists in the package
    pub fn has_file(&self, path: &str) -> bool {
        self.archive.borrow_mut().by_name(path).is_ok()
    }

    /// Get the manifest
    #[allow(dead_code)]
    pub fn manifest(&self) -> &super::manifest::Manifest {
        &self.manifest
    }

    /// List all files in the package
    pub fn files(&self) -> Result<Vec<String>> {
        let mut files = Vec::new();
        let mut archive = self.archive.borrow_mut();
        for i in 0..archive.len() {
            let file = archive.by_index(i)?;
            files.push(file.name().to_string());
        }
        Ok(files)
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
