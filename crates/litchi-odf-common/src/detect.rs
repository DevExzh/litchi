//! Best-effort OpenDocument format detection.
//!
//! Detection is inert: it reads the standardized MIME type from a flat XML
//! root or packaged `mimetype` member without constructing a document model.
//!
//! ```rust
//! use litchi_odf_common::detect::{self, Format};
//!
//! assert_eq!(
//!     detect::mime(b"application/vnd.oasis.opendocument.text"),
//!     Some(Format::Odt),
//! );
//! ```

use crate::constants;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::{Read, Seek, SeekFrom};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const ZIP_SIGNATURE: &[u8] = b"PK\x03\x04";
const LOCAL_HEADER_BYTES: usize = 30;
const MIMETYPE_PATH: &str = "mimetype";
const MIMETYPE_NAME: &[u8] = MIMETYPE_PATH.as_bytes();
const MAX_MIMETYPE_BYTES: usize = 256;

/// Neutral file classification returned by the detector.
pub use litchi_core::FileFormat as Format;

/// Classify the raw contents of an ODF `mimetype` member.
///
/// Leading and trailing ASCII whitespace is ignored without allocating.
/// Unknown MIME types and invalid UTF-8 return `None`.
#[inline]
pub fn mime(value: &[u8]) -> Option<Format> {
    match std::str::from_utf8(trim_ascii(value)).ok()? {
        constants::ODF_TEXT | constants::ODF_TEXT_TEMPLATE => Some(Format::Odt),
        constants::ODF_SPREADSHEET | constants::ODF_SPREADSHEET_TEMPLATE => Some(Format::Ods),
        constants::ODF_PRESENTATION | constants::ODF_PRESENTATION_TEMPLATE => Some(Format::Odp),
        constants::ODF_DRAWING | constants::ODF_DRAWING_TEMPLATE => Some(Format::Odg),
        constants::ODF_CHART | constants::ODF_CHART_TEMPLATE => Some(Format::Odc),
        constants::ODF_FORMULA | constants::ODF_FORMULA_TEMPLATE => Some(Format::Odf),
        constants::ODF_IMAGE | constants::ODF_IMAGE_TEMPLATE => Some(Format::Odi),
        constants::ODF_MASTER | constants::ODF_MASTER_TEMPLATE => Some(Format::Odm),
        constants::ODF_WEB => Some(Format::Oth),
        constants::ODF_DATABASE => Some(Format::Odb),
        _ => None,
    }
}

/// Read a recognized `office:mimetype` value from a flat ODF XML root.
///
/// The root must be `office:document` in the ODF office namespace. Namespace
/// prefixes are resolved semantically. The returned value is decoded and XML
/// attribute whitespace is normalized before classification, then copied into
/// an owned string.
pub fn flat_mime(value: &[u8]) -> Option<String> {
    with_flat_mime(value, |value| {
        let value = trim_ascii(value);
        mime(value)?;
        let value = std::str::from_utf8(value).ok()?;
        Some(value.to_owned())
    })
}

fn with_flat_mime<T>(value: &[u8], classify: impl FnOnce(&[u8]) -> Option<T>) -> Option<T> {
    let mut reader = NsReader::from_reader(value);
    loop {
        let (namespace, event) = reader.read_resolved_event().ok()?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                    || element.local_name().as_ref() != b"document"
                {
                    return None;
                }
                for attribute in element.attributes() {
                    let attribute = attribute.ok()?;
                    let (namespace, local_name) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                        && local_name.as_ref() == b"mimetype"
                    {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .ok()?;
                        return classify(value.as_bytes());
                    }
                }
                return None;
            },
            Event::Decl(_) | Event::Comment(_) | Event::DocType(_) | Event::PI(_) => {},
            Event::Text(text) if text.iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => return None,
            _ => return None,
        }
    }
}

/// Detect a flat OpenDocument XML document.
#[inline]
pub fn flat(value: &[u8]) -> Option<Format> {
    with_flat_mime(value, mime)
}

/// Detect a packaged or flat OpenDocument document from complete bytes.
///
/// Conforming ODF packages place an uncompressed `mimetype` entry first. The
/// local ZIP header and payload are checked in place without allocating a
/// decompression buffer. The central archive structure is also validated.
pub fn bytes(value: &[u8]) -> Option<Format> {
    if value.starts_with(ZIP_SIGNATURE) {
        let format = mime(packaged_mime(value)?)?;
        let archive = soapberry_zip::office::ArchiveReader::new(value).ok()?;
        if !archive.is_stored(MIMETYPE_PATH).ok()? {
            return None;
        }
        return Some(format);
    }
    flat(value)
}

/// Detect a packaged or flat OpenDocument stream.
///
/// Detection reads the complete stream from its beginning and restores the
/// caller's original cursor position on every success or failure path. If the
/// original position cannot be restored, this function returns `None`.
pub fn reader<R: Read + Seek>(value: &mut R) -> Option<Format> {
    let original = value.stream_position().ok()?;
    let detected = (|| {
        value.seek(SeekFrom::Start(0)).ok()?;
        let mut data = Vec::new();
        value.read_to_end(&mut data).ok()?;
        bytes(&data)
    })();
    value.seek(SeekFrom::Start(original)).ok()?;
    detected
}

fn packaged_mime(value: &[u8]) -> Option<&[u8]> {
    if value.get(..4)? != ZIP_SIGNATURE {
        return None;
    }
    let flags = little_u16(value, 6)?;
    let compression = little_u16(value, 8)?;
    let expected_crc = little_u32(value, 14)?;
    let compressed = usize::try_from(little_u32(value, 18)?).ok()?;
    let uncompressed = usize::try_from(little_u32(value, 22)?).ok()?;
    let name_len = usize::from(little_u16(value, 26)?);
    let extra_len = usize::from(little_u16(value, 28)?);

    // ODF permits the UTF-8 file-name flag, but its mimetype entry must be
    // stored, sized in the local header, unencrypted, and have no extra field.
    if flags & !(1 << 11) != 0
        || compression != 0
        || compressed != uncompressed
        || uncompressed > MAX_MIMETYPE_BYTES
        || name_len != MIMETYPE_NAME.len()
        || extra_len != 0
    {
        return None;
    }
    let name_start = LOCAL_HEADER_BYTES;
    let name_end = name_start.checked_add(name_len)?;
    if value.get(name_start..name_end)? != MIMETYPE_NAME {
        return None;
    }
    let data_start = name_end.checked_add(extra_len)?;
    let data_end = data_start.checked_add(compressed)?;
    let data = value.get(data_start..data_end)?;
    (soapberry_zip::crc32(data) == expected_crc).then_some(data)
}

fn little_u16(value: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        value.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn little_u32(value: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        value.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn trim_ascii(mut value: &[u8]) -> &[u8] {
    while value.first().is_some_and(u8::is_ascii_whitespace) {
        value = &value[1..];
    }
    while value.last().is_some_and(u8::is_ascii_whitespace) {
        value = &value[..value.len() - 1];
    }
    value
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn classifies_every_packaged_family_and_template_without_lossy_text() {
        for (value, expected) in [
            (constants::ODF_TEXT, Format::Odt),
            (constants::ODF_TEXT_TEMPLATE, Format::Odt),
            (constants::ODF_SPREADSHEET, Format::Ods),
            (constants::ODF_SPREADSHEET_TEMPLATE, Format::Ods),
            (constants::ODF_PRESENTATION, Format::Odp),
            (constants::ODF_PRESENTATION_TEMPLATE, Format::Odp),
            (constants::ODF_DRAWING, Format::Odg),
            (constants::ODF_DRAWING_TEMPLATE, Format::Odg),
            (constants::ODF_CHART, Format::Odc),
            (constants::ODF_CHART_TEMPLATE, Format::Odc),
            (constants::ODF_FORMULA, Format::Odf),
            (constants::ODF_FORMULA_TEMPLATE, Format::Odf),
            (constants::ODF_IMAGE, Format::Odi),
            (constants::ODF_IMAGE_TEMPLATE, Format::Odi),
            (constants::ODF_MASTER, Format::Odm),
            (constants::ODF_MASTER_TEMPLATE, Format::Odm),
            (constants::ODF_WEB, Format::Oth),
            (constants::ODF_DATABASE, Format::Odb),
        ] {
            assert_eq!(mime(value.as_bytes()), Some(expected), "{value}");
        }
        assert_eq!(
            mime(b" \napplication/vnd.oasis.opendocument.text\t"),
            Some(Format::Odt)
        );
        assert_eq!(mime(b"application/pdf"), None);
        assert_eq!(mime(b"\xff"), None);
    }

    #[test]
    fn detects_flat_documents_with_semantic_namespace_resolution() {
        for (body, mimetype, expected) in [
            ("text", constants::ODF_TEXT, Format::Odt),
            ("spreadsheet", constants::ODF_SPREADSHEET, Format::Ods),
            ("presentation", constants::ODF_PRESENTATION, Format::Odp),
            ("drawing", constants::ODF_DRAWING, Format::Odg),
            ("chart", constants::ODF_CHART, Format::Odc),
            ("formula", constants::ODF_FORMULA, Format::Odf),
            ("image", constants::ODF_IMAGE, Format::Odi),
        ] {
            let xml = format!(
                r#"<?xml version="1.0"?><!--flat--><o:document xmlns:o="{}" o:mimetype="{mimetype}"><o:body><o:{body}/></o:body></o:document>"#,
                String::from_utf8_lossy(OFFICE_NAMESPACE),
            );
            assert_eq!(flat_mime(xml.as_bytes()).as_deref(), Some(mimetype));
            assert_eq!(flat(xml.as_bytes()), Some(expected));
            assert_eq!(bytes(xml.as_bytes()), Some(expected));
        }

        let padded = format!(
            r#"<o:document xmlns:o="{}" o:mimetype="  {}  "><o:body><o:text/></o:body></o:document>"#,
            String::from_utf8_lossy(OFFICE_NAMESPACE),
            constants::ODF_TEXT,
        );
        assert_eq!(
            flat_mime(padded.as_bytes()).as_deref(),
            Some(constants::ODF_TEXT)
        );
    }

    #[test]
    fn rejects_non_flat_roots_and_unknown_mimetypes() {
        for xml in [
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text"/>"#,
            r#"<office:document xmlns:office="urn:wrong" office:mimetype="application/vnd.oasis.opendocument.text"/>"#,
            r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/xml"/>"#,
            r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
        ] {
            assert_eq!(flat(xml.as_bytes()), None);
        }
    }

    #[test]
    fn detects_packaged_documents_and_restores_nonzero_reader_position() {
        let mut writer = crate::core::PackageWriter::new();
        writer.set_mimetype(constants::ODF_TEXT).unwrap();
        let package = writer.finish_to_bytes().unwrap();
        assert_eq!(bytes(&package), Some(Format::Odt));

        let mut input = Cursor::new(package);
        input.set_position(7);
        assert_eq!(reader(&mut input), Some(Format::Odt));
        assert_eq!(input.position(), 7);

        let xml = format!(
            r#"<o:document xmlns:o="{}" o:mimetype="{}"><o:body><o:text/></o:body></o:document>"#,
            String::from_utf8_lossy(OFFICE_NAMESPACE),
            constants::ODF_TEXT,
        );
        let mut flat_input = Cursor::new(xml);
        flat_input.set_position(9);
        assert_eq!(reader(&mut flat_input), Some(Format::Odt));
        assert_eq!(flat_input.position(), 9);

        let mut invalid = Cursor::new(b"not an OpenDocument file".to_vec());
        invalid.set_position(4);
        assert_eq!(reader(&mut invalid), None);
        assert_eq!(invalid.position(), 4);
    }

    #[test]
    fn packaged_detection_rejects_nonconforming_local_mimetype_entries() {
        let mut writer = crate::core::PackageWriter::new();
        writer.set_mimetype(constants::ODF_TEXT).unwrap();
        let package = writer.finish_to_bytes().unwrap();

        let mut compressed = package.clone();
        compressed[8..10].copy_from_slice(&8_u16.to_le_bytes());
        assert_eq!(bytes(&compressed), None);

        let mut corrupt = package.clone();
        let payload = LOCAL_HEADER_BYTES + MIMETYPE_NAME.len();
        corrupt[payload] ^= 1;
        assert_eq!(bytes(&corrupt), None);

        assert_eq!(bytes(&package[..LOCAL_HEADER_BYTES]), None);
        let local_entry_end = LOCAL_HEADER_BYTES + MIMETYPE_NAME.len() + constants::ODF_TEXT.len();
        assert_eq!(bytes(&package[..local_entry_end]), None);
    }
}
