//! ODF manifest parsing functionality.
//!
//! The manifest.xml file contains metadata about all files in the ODF package,
//! including their MIME types, sizes, and encryption status.

use crate::common::{Error, Result};
use soapberry_zip::office::ArchiveReader;
use std::collections::HashMap;

/// ODF manifest (META-INF/manifest.xml)
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct Manifest {
    pub mimetype: String,
    pub entries: HashMap<String, ManifestEntry>,
}

/// Entry in the ODF manifest
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct ManifestEntry {
    pub full_path: String,
    pub media_type: String,
    pub size: Option<u64>,
    pub encrypted: bool,
}

impl Manifest {
    /// Parse manifest from ArchiveReader
    pub fn from_archive_reader(archive: &ArchiveReader<'_>) -> Result<Self> {
        // Try to read manifest.xml from META-INF/manifest.xml first
        let manifest_content = if let Ok(content) = archive.read_string("META-INF/manifest.xml") {
            content
        } else if let Ok(content) = archive.read_string("manifest.xml") {
            // Try alternate manifest location for some ODF files
            content
        } else {
            return Err(Error::InvalidFormat(
                "No manifest.xml found in ODF package".to_string(),
            ));
        };

        Self::parse(&manifest_content)
    }

    /// Parse manifest XML content
    pub fn parse(xml_content: &str) -> Result<Self> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();

        let mut entries = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.name().as_ref() == b"manifest:file-entry"
                        && let Some(entry) = Self::parse_file_entry(e)?
                    {
                        let full_path = entry.full_path.clone();
                        entries.insert(full_path, entry);
                    }
                },
                Ok(Event::Empty(ref e)) => {
                    if e.name().as_ref() == b"manifest:file-entry"
                        && let Some(entry) = Self::parse_file_entry(e)?
                    {
                        let full_path = entry.full_path.clone();
                        entries.insert(full_path, entry);
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::InvalidFormat(format!("XML parsing error: {}", e))),
                _ => {},
            }
            buf.clear();
        }

        // Extract mimetype from root document entry
        // Note: Clone necessary here as we need an owned String for the struct field
        let mimetype = entries
            .get("/")
            .map(|entry| entry.media_type.clone())
            .unwrap_or_else(|| "application/vnd.oasis.opendocument.text".to_string());

        Ok(Self { mimetype, entries })
    }

    /// Parse a single file-entry element
    fn parse_file_entry(e: &quick_xml::events::BytesStart) -> Result<Option<ManifestEntry>> {
        let mut full_path = String::new();
        let mut media_type = String::new();
        let mut size = None;

        for attr_result in e.attributes() {
            let attr = attr_result
                .map_err(|_| Error::InvalidFormat("Invalid attribute in manifest".to_string()))?;
            let value = String::from_utf8(attr.value.to_vec())
                .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in manifest".to_string()))?;

            match attr.key.as_ref() {
                b"manifest:full-path" => full_path = value,
                b"manifest:media-type" => media_type = value,
                b"manifest:size" => {
                    if let Ok(s) = value.parse::<u64>() {
                        size = Some(s);
                    }
                },
                _ => {},
            }
        }

        if !full_path.is_empty() {
            let encrypted = media_type == "application/vnd.sun.star.oleobject"
                || media_type.contains("encrypted");

            Ok(Some(ManifestEntry {
                full_path,
                media_type,
                size,
                encrypted,
            }))
        } else {
            Ok(None)
        }
    }

    /// Get media type for a path
    #[allow(dead_code)]
    pub fn get_media_type(&self, path: &str) -> Option<&str> {
        self.entries
            .get(path)
            .map(|entry| entry.media_type.as_str())
    }

    /// Check if a path exists in manifest
    #[allow(dead_code)]
    pub fn has_path(&self, path: &str) -> bool {
        self.entries.contains_key(path)
    }

    /// Get all paths in manifest
    #[allow(dead_code)]
    pub fn paths(&self) -> impl Iterator<Item = &String> {
        self.entries.keys()
    }

    /// Get entry for a path
    #[allow(dead_code)]
    pub fn get_entry(&self, path: &str) -> Option<&ManifestEntry> {
        self.entries.get(path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEST_MANIFEST: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text" manifest:size="1234"/>
    <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml" manifest:size="5678"/>
    <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml" manifest:size="901"/>
    <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml" manifest:size="234"/>
</manifest:manifest>"#;

    #[test]
    fn test_manifest_parse() {
        let manifest = Manifest::parse(TEST_MANIFEST).unwrap();
        assert_eq!(manifest.mimetype, "application/vnd.oasis.opendocument.text");
        assert_eq!(manifest.entries.len(), 4);
    }

    #[test]
    fn test_manifest_parse_empty() {
        let xml = r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"></manifest:manifest>"#;
        let manifest = Manifest::parse(xml).unwrap();
        assert_eq!(manifest.entries.len(), 0);
        assert_eq!(manifest.mimetype, "application/vnd.oasis.opendocument.text"); // Default
    }

    #[test]
    fn test_manifest_get_media_type() {
        let manifest = Manifest::parse(TEST_MANIFEST).unwrap();
        assert_eq!(
            manifest.get_media_type("/"),
            Some("application/vnd.oasis.opendocument.text")
        );
        assert_eq!(manifest.get_media_type("content.xml"), Some("text/xml"));
        assert_eq!(manifest.get_media_type("nonexistent.xml"), None);
    }

    #[test]
    fn test_manifest_has_path() {
        let manifest = Manifest::parse(TEST_MANIFEST).unwrap();
        assert!(manifest.has_path("/"));
        assert!(manifest.has_path("content.xml"));
        assert!(!manifest.has_path("nonexistent.xml"));
    }

    #[test]
    fn test_manifest_paths() {
        let manifest = Manifest::parse(TEST_MANIFEST).unwrap();
        let paths: Vec<&String> = manifest.paths().collect();
        assert_eq!(paths.len(), 4);
        assert!(paths.contains(&&"/".to_string()));
        assert!(paths.contains(&&"content.xml".to_string()));
    }

    #[test]
    fn test_manifest_get_entry() {
        let manifest = Manifest::parse(TEST_MANIFEST).unwrap();

        let entry = manifest.get_entry("content.xml").unwrap();
        assert_eq!(entry.full_path, "content.xml");
        assert_eq!(entry.media_type, "text/xml");
        assert_eq!(entry.size, Some(5678));
        assert!(!entry.encrypted);

        assert!(manifest.get_entry("nonexistent.xml").is_none());
    }

    #[test]
    fn test_manifest_entry_encrypted() {
        let xml = r#"<?xml version="1.0"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="application/vnd.sun.star.oleobject"/>
</manifest:manifest>"#;
        let manifest = Manifest::parse(xml).unwrap();
        let entry = manifest.get_entry("content.xml").unwrap();
        assert!(entry.encrypted);
    }
}
