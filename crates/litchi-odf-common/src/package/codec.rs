//! Archive access and neutral `manifest.xml` codecs.

use super::model::{Archive, Entry, Manifest};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use soapberry_zip::office::ArchiveReader;
use std::collections::HashMap;

const MANIFEST_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const MANIFEST_PATHS: [&str; 2] = ["META-INF/manifest.xml", "manifest.xml"];

impl<'data> Archive<'data> {
    /// Open an ODF ZIP archive from a borrowed byte slice.
    ///
    /// # Errors
    ///
    /// Returns an error when `data` is not a readable ZIP archive.
    pub fn new(data: &'data [u8]) -> Result<Self> {
        let reader = ArchiveReader::new(data)
            .map_err(|error| Error::InvalidFormat(format!("Invalid ZIP archive: {error}")))?;
        Ok(Self { reader })
    }

    /// Read and decode one archive member.
    ///
    /// # Errors
    ///
    /// Returns an error when `path` is absent or its content cannot be read.
    pub fn read(&self, path: &str) -> Result<Vec<u8>> {
        self.reader
            .read(path)
            .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    /// Read one UTF-8 archive member.
    ///
    /// # Errors
    ///
    /// Returns an error when `path` cannot be read or does not contain UTF-8.
    pub fn read_string(&self, path: &str) -> Result<String> {
        let bytes = self.read(path)?;
        String::from_utf8(bytes)
            .map_err(|error| Error::InvalidFormat(format!("Invalid UTF-8 in '{path}': {error}")))
    }

    /// Read the package manifest, accepting both common ODF locations.
    ///
    /// # Errors
    ///
    /// Returns an error when neither conventional manifest location can be
    /// read as UTF-8.
    pub fn read_manifest_xml(&self) -> Result<String> {
        match self.read_string(MANIFEST_PATHS[0]) {
            Ok(manifest) => Ok(manifest),
            Err(first_error) => self.read_string(MANIFEST_PATHS[1]).map_err(|second_error| {
                Error::InvalidFormat(format!(
                    "No manifest.xml found in ODF package: {first_error}; {second_error}"
                ))
            }),
        }
    }

    /// Check whether an archive member exists.
    #[must_use]
    pub fn contains(&self, path: &str) -> bool {
        self.reader.contains(path)
    }

    /// Iterate over archive members in physical order.
    pub fn file_names(&self) -> impl Iterator<Item = &str> {
        self.reader.file_names()
    }

    /// Check whether an archive member uses ZIP Store.
    ///
    /// # Errors
    ///
    /// Returns an error when `path` cannot be inspected in the ZIP archive.
    pub fn is_stored(&self, path: &str) -> Result<bool> {
        self.reader
            .is_stored(path)
            .map_err(|error| Error::InvalidFormat(error.to_string()))
    }
}

/// Read and parse the family-neutral package manifest.
///
/// # Errors
///
/// Returns an error when the manifest member cannot be read or parsed.
pub fn read_manifest(archive: &Archive<'_>) -> Result<Manifest> {
    parse_manifest(&archive.read_manifest_xml()?)
}

/// Read the neutral file-entry model from an archive reader.
///
/// # Errors
///
/// Returns an error when the XML is malformed or contains invalid manifest
/// entries.
pub fn parse_manifest(xml: &str) -> Result<Manifest> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut entries = HashMap::new();
    let mut current_path: Option<String> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("Invalid manifest XML: {error}")))?;
        match event {
            Event::Start(element) if is_manifest_element(&namespace, &element, b"file-entry") => {
                if current_path.is_some() {
                    return Err(Error::InvalidFormat(
                        "Nested manifest file entries are invalid".to_string(),
                    ));
                }
                let (path, entry) = parse_entry(&reader, &element)?.ok_or_else(|| {
                    Error::InvalidFormat("Manifest file entry has no full path".to_string())
                })?;
                if entries.insert(path.clone(), entry).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Duplicate manifest file entry '{path}'"
                    )));
                }
                current_path = Some(path);
            },
            Event::Empty(element) if is_manifest_element(&namespace, &element, b"file-entry") => {
                let (path, entry) = parse_entry(&reader, &element)?.ok_or_else(|| {
                    Error::InvalidFormat("Manifest file entry has no full path".to_string())
                })?;
                if entries.insert(path.clone(), entry).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Duplicate manifest file entry '{path}'"
                    )));
                }
            },
            Event::End(element)
                if namespace_is_manifest(&namespace)
                    && element.local_name().as_ref() == b"file-entry" =>
            {
                current_path = None;
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if current_path.is_some() {
        return Err(Error::InvalidFormat(
            "Incomplete manifest file entry".to_string(),
        ));
    }
    let mimetype = entries
        .get("/")
        .map(|entry| entry.media_type.clone())
        .unwrap_or_default();
    Ok(Manifest { mimetype, entries })
}

/// Return whether a package path is a likely embedded media resource.
#[must_use]
pub fn is_media_path(path: &str) -> bool {
    path.starts_with("Pictures/")
        || path.starts_with("media/")
        || path.starts_with("Object/")
        || has_ascii_extension(path, b".png")
        || has_ascii_extension(path, b".jpg")
        || has_ascii_extension(path, b".jpeg")
        || has_ascii_extension(path, b".gif")
        || has_ascii_extension(path, b".svg")
}

fn is_manifest_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> bool {
    namespace_is_manifest(namespace) && element.local_name().as_ref() == local
}

fn namespace_is_manifest(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == MANIFEST_NAMESPACE)
}

fn manifest_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<HashMap<Vec<u8>, String>> {
    let mut values = HashMap::new();
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            Error::InvalidFormat(format!("Invalid manifest attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NAMESPACE) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("Invalid manifest attribute value: {error}"))
            })?
            .into_owned();
        if values.insert(local.as_ref().to_vec(), value).is_some() {
            return Err(Error::InvalidFormat(
                "Duplicate manifest attribute".to_string(),
            ));
        }
    }
    Ok(values)
}

fn parse_entry(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<(String, Entry)>> {
    let attributes = manifest_attributes(reader, element)?;
    let Some(full_path) = attributes.get(b"full-path".as_slice()).cloned() else {
        return Ok(None);
    };
    if full_path.is_empty() {
        return Ok(None);
    }
    let size = attributes
        .get(b"size".as_slice())
        .map(|value| {
            value.parse::<u64>().map_err(|error| {
                Error::InvalidFormat(format!("Invalid manifest size for '{full_path}': {error}"))
            })
        })
        .transpose()?;
    Ok(Some((
        full_path,
        Entry {
            media_type: attributes
                .get(b"media-type".as_slice())
                .cloned()
                .unwrap_or_default(),
            size,
        },
    )))
}

fn has_ascii_extension(path: &str, extension: &[u8]) -> bool {
    let bytes = path.as_bytes();
    bytes
        .get(bytes.len().saturating_sub(extension.len())..)
        .is_some_and(|suffix| suffix.eq_ignore_ascii_case(extension))
}
