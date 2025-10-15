//! ODF manifest parsing functionality.
//!
//! The manifest.xml file contains metadata about all files in the ODF package,
//! including their MIME types, sizes, and encryption status.

use crate::common::{Error, Result};
use std::collections::HashMap;
use std::io::{Read, Seek};

/// ODF manifest (META-INF/manifest.xml)
#[derive(Debug, Clone)]
pub struct Manifest {
    pub mimetype: String,
    pub entries: HashMap<String, ManifestEntry>,
}

/// Entry in the ODF manifest
#[derive(Debug, Clone)]
pub struct ManifestEntry {
    pub full_path: String,
    pub media_type: String,
    pub size: Option<u64>,
    pub encrypted: bool,
}

impl Manifest {
    /// Parse manifest from ZIP archive
    pub fn from_archive<R: Read + Seek>(archive: &mut zip::ZipArchive<R>) -> Result<Self> {
        // Try to read manifest.xml from META-INF/manifest.xml first
        let manifest_content = if let Ok(mut file) = archive.by_name("META-INF/manifest.xml") {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            content
        } else if let Ok(mut file) = archive.by_name("manifest.xml") {
            // Try alternate manifest location for some ODF files
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            content
        } else {
            return Err(Error::InvalidFormat("No manifest.xml found in ODF package".to_string()));
        };

        Self::parse(&manifest_content)
    }

    /// Parse manifest XML content
    pub fn parse(xml_content: &str) -> Result<Self> {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();

        let mut mimetype = String::new();
        let mut entries = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if e.name().as_ref() == b"manifest:file-entry"
                        && let Some(entry) = Self::parse_file_entry(e)? {
                            entries.insert(entry.full_path.clone(), entry);
                        }
                }
                Ok(Event::Empty(ref e)) => {
                    if e.name().as_ref() == b"manifest:file-entry"
                        && let Some(entry) = Self::parse_file_entry(e)? {
                            entries.insert(entry.full_path.clone(), entry);
                        }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::InvalidFormat(format!("XML parsing error: {}", e))),
                _ => {}
            }
            buf.clear();
        }

        // Extract mimetype from root document entry
        mimetype = entries.get("/")
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
            let attr = attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute in manifest".to_string()))?;
            let value = String::from_utf8(attr.value.to_vec())
                .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in manifest".to_string()))?;

            match attr.key.as_ref() {
                b"manifest:full-path" => full_path = value,
                b"manifest:media-type" => media_type = value,
                b"manifest:size" => {
                    if let Ok(s) = value.parse::<u64>() {
                        size = Some(s);
                    }
                }
                _ => {}
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
    pub fn get_media_type(&self, path: &str) -> Option<&str> {
        self.entries.get(path).map(|entry| entry.media_type.as_str())
    }

    /// Check if a path exists in manifest
    pub fn has_path(&self, path: &str) -> bool {
        self.entries.contains_key(path)
    }

    /// Get all paths in manifest
    pub fn paths(&self) -> impl Iterator<Item = &String> {
        self.entries.keys()
    }

    /// Get entry for a path
    pub fn get_entry(&self, path: &str) -> Option<&ManifestEntry> {
        self.entries.get(path)
    }
}
