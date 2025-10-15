//! ODF package (ZIP archive) handling functionality.
//!
//! This module provides utilities for working with ODF files as ZIP archives,
//! including reading files, checking existence, and basic package operations.

use crate::common::{Error, Result};
use std::io::{Read, Seek};

/// An ODF package (ZIP file containing XML documents)
pub struct Package<R> {
    archive: zip::ZipArchive<R>,
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
            archive,
            manifest,
            mimetype,
        })
    }

    /// Read MIME type from the mimetype file
    fn read_mimetype(archive: &mut zip::ZipArchive<R>) -> Result<String> {
        let mut mimetype_file = archive.by_name("mimetype")
            .map_err(|_| Error::InvalidFormat("No mimetype file found in ODF package".to_string()))?;

        let mut content = String::new();
        mimetype_file.read_to_string(&mut content)?;
        Ok(content.trim().to_string())
    }

    /// Get the MIME type from the mimetype file
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Get a file from the package by path
    pub fn get_file(&mut self, path: &str) -> Result<Vec<u8>> {
        let mut file = self.archive.by_name(path)
            .map_err(|_| Error::InvalidFormat(format!("File not found: {}", path)))?;

        let mut content = Vec::new();
        file.read_to_end(&mut content)?;
        Ok(content)
    }

    /// Check if a file exists in the package
    pub fn has_file(&mut self, path: &str) -> bool {
        self.archive.by_name(path).is_ok()
    }

    /// Get the manifest
    pub fn manifest(&self) -> &super::manifest::Manifest {
        &self.manifest
    }

    /// List all files in the package
    pub fn files(&mut self) -> Result<Vec<String>> {
        let mut files = Vec::new();
        for i in 0..self.archive.len() {
            let file = self.archive.by_index(i)?;
            files.push(file.name().to_string());
        }
        Ok(files)
    }
}
