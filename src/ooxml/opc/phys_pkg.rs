//! Provides a general interface to a physical OPC package (ZIP file).
//!
//! This module handles the low-level reading of OPC packages from ZIP archives,
//! providing efficient access to package contents with minimal memory allocation.
//!
//! Uses the high-performance soapberry-zip library for zero-copy ZIP parsing.

use crate::ooxml::opc::error::{OpcError, Result};
use crate::ooxml::opc::packuri::PackURI;
use soapberry_zip::office::LazyArchiveReader;
use std::io::Read;
use std::path::Path;

/// Physical package reader that provides access to parts in a ZIP-based OPC package.
///
/// Uses soapberry_zip for high-performance zero-copy ZIP parsing with lazy decompression.
/// File contents are decompressed on-demand and cached for efficiency. This enables
/// pipelining of decompression with XML parsing for better throughput.
pub struct PhysPkgReader<'data> {
    /// The underlying ZIP archive reader (lazy decompression with caching)
    archive: LazyArchiveReader<'data>,
}

/// Owned version of PhysPkgReader that owns the data buffer.
///
/// This is used when reading from files or readers where we need to own the data.
pub struct OwnedPhysPkgReader {
    /// The owned data buffer
    data: Vec<u8>,
}

impl OwnedPhysPkgReader {
    /// Open an OPC package from a file path.
    ///
    /// # Arguments
    /// * `path` - Path to the OPC package file (.docx, .xlsx, .pptx, etc.)
    ///
    /// # Returns
    /// A new OwnedPhysPkgReader instance
    ///
    /// # Errors
    /// Returns an error if the file doesn't exist, isn't a valid ZIP file,
    /// or cannot be opened.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();

        if !path.exists() {
            return Err(OpcError::PackageNotFound(path.display().to_string()));
        }

        let data = std::fs::read(path)?;
        Self::from_bytes(data)
    }

    /// Create a new OwnedPhysPkgReader from owned bytes.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        // Validate the ZIP archive can be parsed
        let _ = LazyArchiveReader::new(&data)?;
        Ok(Self { data })
    }

    /// Create a new OwnedPhysPkgReader from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read
    ///
    /// # Returns
    /// A new OwnedPhysPkgReader instance
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes(data)
    }

    /// Get a borrowed reader for accessing archive contents.
    #[inline]
    pub fn reader(&self) -> Result<PhysPkgReader<'_>> {
        PhysPkgReader::new(&self.data)
    }

    /// Get the binary content for a part by its PackURI.
    #[inline]
    pub fn blob_for(&self, pack_uri: &PackURI) -> Result<Vec<u8>> {
        self.reader()?.blob_for(pack_uri)
    }

    /// Get the [Content_Types].xml content.
    #[inline]
    pub fn content_types_xml(&self) -> Result<Vec<u8>> {
        self.reader()?.content_types_xml()
    }

    /// Get the relationships XML for a specific source URI.
    #[inline]
    pub fn rels_xml_for(&self, source_uri: &PackURI) -> Result<Option<Vec<u8>>> {
        self.reader()?.rels_xml_for(source_uri)
    }

    /// Get the number of files in the package.
    #[inline]
    pub fn len(&self) -> Result<usize> {
        Ok(self.reader()?.len())
    }

    /// Check if the package is empty.
    #[inline]
    pub fn is_empty(&self) -> Result<bool> {
        Ok(self.reader()?.is_empty())
    }

    /// List all member names in the package.
    #[inline]
    pub fn member_names(&self) -> Result<Vec<String>> {
        self.reader()?.member_names()
    }

    /// Check if a specific member exists in the package.
    #[inline]
    pub fn contains(&self, pack_uri: &PackURI) -> Result<bool> {
        Ok(self.reader()?.contains(pack_uri))
    }

    /// Consume self and return the underlying data.
    #[inline]
    pub fn into_inner(self) -> Vec<u8> {
        self.data
    }

    /// Get a reference to the underlying data.
    #[inline]
    pub fn as_bytes(&self) -> &[u8] {
        &self.data
    }
}

impl<'data> PhysPkgReader<'data> {
    /// Create a new PhysPkgReader from a byte slice.
    ///
    /// # Arguments
    /// * `data` - The ZIP archive data as a byte slice
    ///
    /// # Returns
    /// A new PhysPkgReader instance
    pub fn new(data: &'data [u8]) -> Result<Self> {
        let archive = LazyArchiveReader::new(data)?;
        Ok(Self { archive })
    }

    /// Get the binary content for a part by its PackURI.
    ///
    /// Uses efficient lazy decompression. The returned vector contains
    /// the decompressed content.
    ///
    /// # Arguments
    /// * `pack_uri` - The PackURI of the part to read
    ///
    /// # Returns
    /// The binary content of the part
    pub fn blob_for(&self, pack_uri: &PackURI) -> Result<Vec<u8>> {
        let membername = pack_uri.membername();

        self.archive
            .read(membername)
            .map_err(|_| OpcError::PartNotFound(pack_uri.to_string()))
    }

    /// Get the [Content_Types].xml content.
    ///
    /// This is a required part of every OPC package that maps parts to content types.
    pub fn content_types_xml(&self) -> Result<Vec<u8>> {
        let content_types_uri = PackURI::new(crate::ooxml::opc::packuri::CONTENT_TYPES_URI)
            .map_err(OpcError::InvalidPackUri)?;
        self.blob_for(&content_types_uri)
    }

    /// Get the relationships XML for a specific source URI.
    ///
    /// Relationships files are stored in _rels directories and have a .rels extension.
    /// Returns None if the source has no relationships file.
    ///
    /// # Arguments
    /// * `source_uri` - The PackURI of the source (part or package)
    pub fn rels_xml_for(&self, source_uri: &PackURI) -> Result<Option<Vec<u8>>> {
        let rels_uri = source_uri.rels_uri().map_err(OpcError::InvalidPackUri)?;

        match self.blob_for(&rels_uri) {
            Ok(blob) => Ok(Some(blob)),
            Err(OpcError::PartNotFound(_)) => Ok(None),
            Err(e) => Err(e),
        }
    }

    /// Get the number of files in the package (excluding directories).
    pub fn len(&self) -> usize {
        self.archive.len()
    }

    /// Check if the package is empty.
    pub fn is_empty(&self) -> bool {
        self.archive.is_empty()
    }

    /// List all member names in the package.
    ///
    /// Returns all file names in the ZIP archive (excluding directories).
    /// Useful for debugging or exploring package contents.
    pub fn member_names(&self) -> Result<Vec<String>> {
        Ok(self.archive.file_names().map(String::from).collect())
    }

    /// Check if a specific member exists in the package.
    ///
    /// Uses the pre-built index for O(1) lookup.
    ///
    /// # Arguments
    /// * `pack_uri` - The PackURI to check
    pub fn contains(&self, pack_uri: &PackURI) -> bool {
        let membername = pack_uri.membername();
        self.archive.contains(membername)
    }

    /// Read multiple blobs in parallel.
    ///
    /// Uses rayon for parallel decompression, providing significant speedup
    /// when reading many parts at once.
    ///
    /// # Arguments
    /// * `uris` - Slice of PackURIs to read
    ///
    /// # Returns
    /// A HashMap mapping member names to their decompressed contents.
    /// Parts that fail to read are not included in the result.
    pub fn blobs_parallel(&self, uris: &[PackURI]) -> std::collections::HashMap<String, Vec<u8>> {
        let names: Vec<&str> = uris.iter().map(|uri| uri.membername()).collect();
        self.archive.read_many_parallel(&names)
    }

    /// Get a reference to the underlying lazy archive reader.
    ///
    /// The lazy reader decompresses files on-demand and caches results.
    /// This enables pipelining of decompression with parsing.
    #[inline]
    pub fn archive(&self) -> &LazyArchiveReader<'data> {
        &self.archive
    }
}

/// Physical package writer for creating OPC packages.
///
/// Handles the low-level writing of parts to a ZIP archive with optimal compression.
/// Uses soapberry-zip's high-performance writer.
pub struct PhysPkgWriter {
    /// The underlying ZIP archive writer
    archive: soapberry_zip::office::StreamingArchiveWriter<std::io::Cursor<Vec<u8>>>,
}

impl PhysPkgWriter {
    /// Create a new package writer that writes to memory.
    pub fn new() -> Self {
        Self {
            archive: soapberry_zip::office::StreamingArchiveWriter::new(),
        }
    }

    /// Write a part to the package with Deflate compression.
    ///
    /// # Arguments
    /// * `pack_uri` - The PackURI for the part
    /// * `blob` - The binary content to write
    pub fn write(&mut self, pack_uri: &PackURI, blob: &[u8]) -> Result<()> {
        self.archive
            .write_deflated(pack_uri.membername(), blob)
            .map_err(|e| OpcError::ZipError(e.to_string()))
    }

    /// Write a part to the package without compression (stored).
    ///
    /// # Arguments
    /// * `pack_uri` - The PackURI for the part
    /// * `blob` - The binary content to write
    pub fn write_stored(&mut self, pack_uri: &PackURI, blob: &[u8]) -> Result<()> {
        self.archive
            .write_stored(pack_uri.membername(), blob)
            .map_err(|e| OpcError::ZipError(e.to_string()))
    }

    /// Finish writing and return the package bytes.
    ///
    /// Consumes the writer and returns the complete ZIP archive.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.archive
            .finish_to_bytes()
            .map_err(|e| OpcError::ZipError(e.to_string()))
    }
}

impl Default for PhysPkgWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_round_trip() {
        // Create a ZIP archive with soapberry-zip
        let mut writer = PhysPkgWriter::new();
        let pack_uri = PackURI::new("/test.txt").unwrap();
        writer.write(&pack_uri, b"Hello, World!").unwrap();
        let zip_data = writer.finish().unwrap();

        // Read the ZIP archive
        let reader = PhysPkgReader::new(&zip_data).unwrap();
        let content = reader.blob_for(&pack_uri).unwrap();
        assert_eq!(content, b"Hello, World!");
    }

    #[test]
    fn test_multiple_parts() {
        let mut writer = PhysPkgWriter::new();

        let content_types = PackURI::new("/[Content_Types].xml").unwrap();
        let rels = PackURI::new("/_rels/.rels").unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();

        writer.write(&content_types, b"<Types/>").unwrap();
        writer.write(&rels, b"<Relationships/>").unwrap();
        writer.write(&document, b"<document/>").unwrap();

        let zip_data = writer.finish().unwrap();
        let reader = PhysPkgReader::new(&zip_data).unwrap();

        assert!(reader.contains(&content_types));
        assert!(reader.contains(&rels));
        assert!(reader.contains(&document));
        assert_eq!(reader.blob_for(&document).unwrap(), b"<document/>");
    }
}
