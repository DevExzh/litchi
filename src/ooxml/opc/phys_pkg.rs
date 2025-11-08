//! Provides a general interface to a physical OPC package (ZIP file).
//!
//! This module handles the low-level reading of OPC packages from ZIP archives,
//! providing efficient access to package contents with minimal memory allocation.

use crate::ooxml::opc::error::{OpcError, Result};
use crate::ooxml::opc::packuri::PackURI;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;
use zip::ZipArchive;

/// Physical package reader that provides access to parts in a ZIP-based OPC package.
///
/// Uses zip::ZipArchive for efficient random access to package members.
/// Implements buffered reading to minimize system calls and improve performance.
pub struct PhysPkgReader<R: Read + Seek> {
    /// The underlying ZIP archive
    archive: ZipArchive<R>,
}

impl PhysPkgReader<BufReader<File>> {
    /// Open an OPC package from a file path.
    ///
    /// # Arguments
    /// * `path` - Path to the OPC package file (.docx, .xlsx, .pptx, etc.)
    ///
    /// # Returns
    /// A new PhysPkgReader instance
    ///
    /// # Errors
    /// Returns an error if the file doesn't exist, isn't a valid ZIP file,
    /// or cannot be opened.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();

        if !path.exists() {
            return Err(OpcError::PackageNotFound(path.display().to_string()));
        }

        let file = File::open(path)?;
        // Use BufReader for better I/O performance
        let buf_reader = BufReader::with_capacity(8192, file);

        Self::new(buf_reader)
    }
}

impl<R: Read + Seek> PhysPkgReader<R> {
    /// Create a new PhysPkgReader from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read + Seek (e.g., File, BufReader)
    ///
    /// # Returns
    /// A new PhysPkgReader instance
    pub fn new(reader: R) -> Result<Self> {
        let archive = ZipArchive::new(reader)?;
        Ok(Self { archive })
    }

    /// Get the binary content for a part by its PackURI.
    ///
    /// Uses efficient reading with minimal allocation. The returned vector
    /// is pre-allocated to the exact size needed, avoiding reallocation.
    ///
    /// # Arguments
    /// * `pack_uri` - The PackURI of the part to read
    ///
    /// # Returns
    /// The binary content of the part
    pub fn blob_for(&mut self, pack_uri: &PackURI) -> Result<Vec<u8>> {
        let membername = pack_uri.membername();

        // Get the file from the ZIP archive
        let mut file = self
            .archive
            .by_name(membername)
            .map_err(|_| OpcError::PartNotFound(pack_uri.to_string()))?;

        // Pre-allocate vector to exact size to avoid reallocation
        let size = file.size() as usize;
        let mut buffer = Vec::with_capacity(size);

        // Read directly into the vector
        file.read_to_end(&mut buffer)?;

        Ok(buffer)
    }

    /// Get the [Content_Types].xml content.
    ///
    /// This is a required part of every OPC package that maps parts to content types.
    pub fn content_types_xml(&mut self) -> Result<Vec<u8>> {
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
    pub fn rels_xml_for(&mut self, source_uri: &PackURI) -> Result<Option<Vec<u8>>> {
        let rels_uri = source_uri.rels_uri().map_err(OpcError::InvalidPackUri)?;

        match self.blob_for(&rels_uri) {
            Ok(blob) => Ok(Some(blob)),
            Err(OpcError::PartNotFound(_)) => Ok(None),
            Err(e) => Err(e),
        }
    }

    /// Get the number of files in the package.
    pub fn len(&self) -> usize {
        self.archive.len()
    }

    /// Check if the package is empty.
    pub fn is_empty(&self) -> bool {
        self.archive.len() == 0
    }

    /// List all member names in the package.
    ///
    /// Returns an iterator over the names of all files in the ZIP archive.
    /// Useful for debugging or exploring package contents.
    pub fn member_names(&mut self) -> Result<Vec<String>> {
        let mut names = Vec::with_capacity(self.archive.len());

        for i in 0..self.archive.len() {
            let file = self.archive.by_index(i)?;
            names.push(file.name().to_string());
        }

        Ok(names)
    }

    /// Check if a specific member exists in the package.
    ///
    /// Uses fast string searching to locate the member without reading its content.
    ///
    /// # Arguments
    /// * `pack_uri` - The PackURI to check
    pub fn contains(&mut self, pack_uri: &PackURI) -> bool {
        let membername = pack_uri.membername();
        self.archive.by_name(membername).is_ok()
    }
}

/// Physical package writer for creating OPC packages.
///
/// Handles the low-level writing of parts to a ZIP archive with optimal compression.
pub struct PhysPkgWriter<W: std::io::Write + std::io::Seek> {
    /// The underlying ZIP archive writer
    archive: zip::ZipWriter<W>,
}

impl PhysPkgWriter<File> {
    /// Create a new package writer for a file path.
    ///
    /// # Arguments
    /// * `path` - Path where the package should be written
    pub fn create<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::create(path)?;
        let archive = zip::ZipWriter::new(file);
        Ok(Self { archive })
    }
}

impl<W: std::io::Write + std::io::Seek> PhysPkgWriter<W> {
    /// Create a new package writer from a writer stream.
    ///
    /// # Arguments
    /// * `writer` - A writer that implements Write + Seek
    pub fn new(writer: W) -> Self {
        Self {
            archive: zip::ZipWriter::new(writer),
        }
    }

    /// Write a part to the package.
    ///
    /// # Arguments
    /// * `pack_uri` - The PackURI for the part
    /// * `blob` - The binary content to write
    pub fn write(&mut self, pack_uri: &PackURI, blob: &[u8]) -> Result<()> {
        use zip::write::SimpleFileOptions;

        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(6)); // Balanced compression level

        self.archive.start_file(pack_uri.membername(), options)?;
        std::io::copy(&mut std::io::Cursor::new(blob), &mut self.archive)?;

        Ok(())
    }

    /// Finish writing and close the package.
    ///
    /// This must be called to ensure all data is flushed to disk.
    pub fn finish(self) -> Result<()> {
        self.archive.finish()?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_reader_from_cursor() {
        // Create a minimal ZIP archive in memory
        let mut zip_data = Vec::new();
        {
            let cursor = Cursor::new(&mut zip_data);
            let mut writer = zip::ZipWriter::new(cursor);

            use zip::write::SimpleFileOptions;
            let options = SimpleFileOptions::default();
            writer.start_file("test.txt", options).unwrap();
            std::io::Write::write_all(&mut writer, b"Hello, World!").unwrap();
            writer.finish().unwrap();
        }

        // Read the ZIP archive
        let cursor = Cursor::new(zip_data);
        let mut reader = PhysPkgReader::new(cursor).unwrap();

        let pack_uri = PackURI::new("/test.txt").unwrap();
        let content = reader.blob_for(&pack_uri).unwrap();
        assert_eq!(content, b"Hello, World!");
    }
}
