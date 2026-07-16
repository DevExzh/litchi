//! Opaque WordprocessingML alternative-format import anchors and payloads.

use crate::docx::namespace::is_wordprocessing_namespace;
use crate::error::{OoxmlError, Result};
use litchi_opc::constants::relationship_type;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;

const TRANSITIONAL_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// A block-level `<w:altChunk>` import anchor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AltChunk {
    relationship_id: String,
    match_source: Option<bool>,
}

impl AltChunk {
    /// Relationship ID identifying the alternative-format import part.
    #[inline]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Whether imported formatting should match the source formatting.
    ///
    /// `None` means `<w:matchSrc>` was absent. `Some(true)` includes the
    /// empty-element form, while `Some(false)` represents an explicit false.
    #[inline]
    pub fn match_source(&self) -> Option<bool> {
        self.match_source
    }
}

/// Recognized MIME family for an opaque alternative-format import part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AlternativeFormatKind {
    WordprocessingMl,
    Html,
    Xhtml,
    Rtf,
    PlainText,
    Xml,
    MimeMessage,
    Unknown,
}

impl AlternativeFormatKind {
    fn classify(content_type: &str) -> Self {
        match content_type.split(';').next().unwrap_or(content_type).trim() {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" => {
                Self::WordprocessingMl
            },
            "text/html" => Self::Html,
            "application/xhtml+xml" => Self::Xhtml,
            "application/rtf" | "text/rtf" => Self::Rtf,
            "text/plain" => Self::PlainText,
            "application/xml" | "text/xml" => Self::Xml,
            "message/rfc822" => Self::MimeMessage,
            _ => Self::Unknown,
        }
    }
}

/// A borrowed, opaque alternative-format import payload.
///
/// Access never parses the foreign format, opens nested packages, fetches
/// resources, or performs filesystem or network I/O.
pub struct AlternativeFormatPart<'a> {
    part: &'a dyn Part,
    kind: AlternativeFormatKind,
}

impl<'a> AlternativeFormatPart<'a> {
    pub(crate) fn new(part: &'a dyn Part) -> Self {
        Self {
            kind: AlternativeFormatKind::classify(part.content_type()),
            part,
        }
    }

    #[inline]
    pub fn part_name(&self) -> &PackURI {
        self.part.partname()
    }

    #[inline]
    pub fn content_type(&self) -> &str {
        self.part.content_type()
    }

    #[inline]
    pub fn kind(&self) -> AlternativeFormatKind {
        self.kind
    }

    /// Return the raw OPC part bytes without interpreting them.
    #[inline]
    pub fn bytes(&self) -> &[u8] {
        self.part.blob()
    }
}

pub(crate) fn is_alternative_format_relationship(value: &str) -> bool {
    matches!(
        value,
        relationship_type::ALTERNATIVE_FORMAT_IMPORT
            | relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
    )
}

struct PendingAltChunk {
    root_depth: usize,
    start: u32,
    relationship_id: String,
    match_source: Option<bool>,
    saw_properties: bool,
    properties_depth: Option<usize>,
}

/// Parse every altChunk anchor against the full namespace context.
pub(crate) fn scan_alt_chunks(xml: &[u8]) -> Result<BTreeMap<u32, AltChunk>> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut pending: Option<PendingAltChunk> = None;
    let mut chunks = BTreeMap::new();

    loop {
        let event_start = u32::try_from(reader.buffer_position()).map_err(|_| {
            OoxmlError::InvalidFormat("altChunk XML offset does not fit u32".into())
        })?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                let event_depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("altChunk XML nesting is too deep".into())
                })?;
                if pending.is_none()
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    pending = Some(PendingAltChunk {
                        root_depth: event_depth,
                        start: event_start,
                        relationship_id: relationship_id(&element, decoder, &resolver)?,
                        match_source: None,
                        saw_properties: false,
                        properties_depth: None,
                    });
                } else if let Some(chunk) = pending.as_mut() {
                    parse_child(
                        chunk,
                        event_depth,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        false,
                    )?;
                }
                depth = event_depth;
            },
            Event::Empty(element) => {
                let event_depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("altChunk XML nesting is too deep".into())
                })?;
                if pending.is_none()
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"altChunk"
                {
                    let chunk = AltChunk {
                        relationship_id: relationship_id(&element, decoder, &resolver)?,
                        match_source: None,
                    };
                    insert_chunk(&mut chunks, event_start, chunk)?;
                } else if let Some(chunk) = pending.as_mut() {
                    parse_child(
                        chunk,
                        event_depth,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        true,
                    )?;
                }
            },
            Event::End(_) => {
                if let Some(chunk) = pending.as_mut()
                    && chunk.properties_depth == Some(depth)
                {
                    chunk.properties_depth = None;
                }
                if pending
                    .as_ref()
                    .is_some_and(|chunk| chunk.root_depth == depth)
                {
                    let chunk = pending.take().ok_or_else(|| {
                        OoxmlError::InvalidFormat("missing pending altChunk".into())
                    })?;
                    insert_chunk(
                        &mut chunks,
                        chunk.start,
                        AltChunk {
                            relationship_id: chunk.relationship_id,
                            match_source: chunk.match_source,
                        },
                    )?;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("unexpected altChunk XML end element".into())
                })?;
            },
            Event::Text(text) if pending.is_some() => {
                if text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                    return Err(OoxmlError::InvalidFormat(
                        "altChunk contains unexpected text".into(),
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if pending.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "altChunk contains unexpected character data".into(),
                ));
            },
            Event::Eof => {
                if pending.is_some() {
                    return Err(OoxmlError::InvalidFormat("unterminated altChunk".into()));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(chunks)
}

fn parse_child(
    chunk: &mut PendingAltChunk,
    depth: usize,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
    empty: bool,
) -> Result<()> {
    let is_word = is_wordprocessing_namespace(namespace);
    if depth == chunk.root_depth + 1
        && is_word
        && element.local_name().as_ref() == b"altChunkPr"
        && !chunk.saw_properties
    {
        chunk.saw_properties = true;
        if !empty {
            chunk.properties_depth = Some(depth);
        }
        return Ok(());
    }
    if depth == chunk.root_depth + 2
        && chunk.properties_depth == Some(chunk.root_depth + 1)
        && is_word
        && element.local_name().as_ref() == b"matchSrc"
        && chunk.match_source.is_none()
    {
        chunk.match_source = Some(parse_on_off(element, decoder, resolver)?);
        return Ok(());
    }
    Err(OoxmlError::InvalidFormat(
        "altChunk has invalid child content".into(),
    ))
}

fn relationship_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let valid_namespace = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri))
                if uri.as_ref() == TRANSITIONAL_RELATIONSHIP_NAMESPACE
                    || uri.as_ref() == STRICT_RELATIONSHIP_NAMESPACE
        );
        if !valid_namespace {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "altChunk has duplicate relationship IDs".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    value
        .filter(|value| !value.is_empty())
        .ok_or_else(|| OoxmlError::InvalidFormat("altChunk lacks a relationship ID".into()))
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_wordprocessing_namespace(&namespace)
            && !matches!(namespace, ResolveResult::Unbound)
        {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "matchSrc has duplicate values".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    match value.as_deref() {
        None | Some("true" | "1" | "on") => Ok(true),
        Some("false" | "0" | "off") => Ok(false),
        Some(value) => Err(OoxmlError::InvalidFormat(format!(
            "invalid matchSrc value '{value}'"
        ))),
    }
}

fn insert_chunk(
    chunks: &mut BTreeMap<u32, AltChunk>,
    start: u32,
    chunk: AltChunk,
) -> Result<()> {
    if chunks.insert(start, chunk).is_some() {
        return Err(OoxmlError::InvalidFormat(
            "duplicate altChunk XML position".into(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::{DocumentBlock, Package};
    use litchi_opc::constants::{content_type, relationship_type};
    use litchi_opc::part::{BlobPart, Part};
    use litchi_opc::{OpcPackage, PackURI};
    use std::io::Cursor;

    const HTML_FIXTURE: &[u8] = include_bytes!(
        "../../../../3rdparty/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-html.docx"
    );
    const DOCX_FIXTURE: &[u8] = include_bytes!(
        "../../../../3rdparty/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk.docx"
    );
    const HEADER_FIXTURE: &[u8] = include_bytes!(
        "../../../../3rdparty/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx"
    );

    #[test]
    fn libreoffice_html_fixture_is_ordered_and_borrowed() {
        let package = Package::from_reader(Cursor::new(HTML_FIXTURE)).unwrap();
        let document = package.document().unwrap();
        let blocks = document.blocks().unwrap();
        assert_eq!(blocks.len(), 3);
        assert!(matches!(blocks[0], DocumentBlock::Paragraph(_)));
        let DocumentBlock::AltChunk(chunk) = &blocks[1] else {
            panic!("missing ordered altChunk")
        };
        assert!(matches!(blocks[2], DocumentBlock::Paragraph(_)));
        let payload = document.resolve_alt_chunk(chunk).unwrap();
        assert_eq!(payload.kind(), AlternativeFormatKind::Html);
        assert_eq!(payload.content_type(), "text/html");
        assert_eq!(
            payload.bytes(),
            b"<html><body><p>HTML AltChunk</p></body></html>"
        );
    }

    #[test]
    fn libreoffice_nested_docx_remains_opaque() {
        let package = Package::from_reader(Cursor::new(DOCX_FIXTURE)).unwrap();
        let document = package.document().unwrap();
        let chunk = document.alt_chunks().unwrap().remove(0);
        let payload = document.resolve_alt_chunk(&chunk).unwrap();
        assert_eq!(payload.kind(), AlternativeFormatKind::WordprocessingMl);
        assert!(payload.bytes().starts_with(b"PK"));
    }

    #[test]
    fn libreoffice_absolute_internal_target_resolves() {
        let package = Package::from_reader(Cursor::new(HEADER_FIXTURE)).unwrap();
        let document = package.document().unwrap();
        let chunk = document.alt_chunks().unwrap().remove(0);
        let payload = document.resolve_alt_chunk(&chunk).unwrap();
        assert_eq!(payload.part_name().as_str(), "/word/afchunk2.docx");
        assert_eq!(payload.kind(), AlternativeFormatKind::WordprocessingMl);
    }

    #[test]
    fn validates_namespaces_order_match_source_and_mce() {
        let xml = br#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:q="http://purl.oclc.org/ooxml/officeDocument/relationships"><s:body><s:p/><s:altChunk q:id="chunk"><s:altChunkPr><s:matchSrc s:val="off"/></s:altChunkPr></s:altChunk><s:tbl/></s:body></s:document>"#;
        let chunks = scan_alt_chunks(xml).unwrap();
        let chunk = chunks.values().next().unwrap();
        assert_eq!(chunk.relationship_id(), "chunk");
        assert_eq!(chunk.match_source(), Some(false));

        let mce = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p/></mc:Choice><mc:Fallback><w:altChunk r:id="fallback"/></mc:Fallback></mc:AlternateContent></w:body></w:document>"#;
        let part = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            content_type::WML_DOCUMENT_MAIN.into(),
            mce.to_vec(),
        );
        let document = crate::docx::parts::DocumentPart::from_part(&part).unwrap();
        let blocks = document.blocks().unwrap();
        assert!(matches!(blocks.as_slice(), [DocumentBlock::AltChunk(_)]));
    }

    #[test]
    fn rejects_missing_wrong_namespace_duplicate_id_and_invalid_children() {
        let wrapper = |anchor: &str| {
            format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:bad="urn:bad"><w:body>{anchor}</w:body></w:document>"#
            )
        };
        for anchor in [
            r#"<w:altChunk/>"#,
            r#"<w:altChunk bad:id="x"/>"#,
            r#"<w:altChunk r:id="x" q:id="y"/>"#,
            r#"<w:altChunk r:id="x"><w:altChunkPr/><w:altChunkPr/></w:altChunk>"#,
            r#"<w:altChunk r:id="x"><w:altChunkPr><w:matchSrc w:val="maybe"/></w:altChunkPr></w:altChunk>"#,
        ] {
            assert!(scan_alt_chunks(wrapper(anchor).as_bytes()).is_err(), "{anchor}");
        }
    }

    #[test]
    fn validates_relationship_type_mode_and_target_without_importing() {
        for relationship in [
            relationship_type::ALTERNATIVE_FORMAT_IMPORT,
            relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT,
        ] {
            let package = synthetic_package(relationship, false, true);
            let document = package.document().unwrap();
            let chunk = document.alt_chunks().unwrap().remove(0);
            assert_eq!(
                document.resolve_alt_chunk(&chunk).unwrap().kind(),
                AlternativeFormatKind::Html
            );
        }

        let package = synthetic_package(relationship_type::IMAGE, false, true);
        let document = package.document().unwrap();
        let chunk = document.alt_chunks().unwrap().remove(0);
        assert!(document.resolve_alt_chunk(&chunk).is_err());

        let package = synthetic_package(relationship_type::ALTERNATIVE_FORMAT_IMPORT, true, true);
        let document = package.document().unwrap();
        let chunk = document.alt_chunks().unwrap().remove(0);
        assert!(document.resolve_alt_chunk(&chunk).is_err());

        let package = synthetic_package(relationship_type::ALTERNATIVE_FORMAT_IMPORT, false, false);
        let document = package.document().unwrap();
        let chunk = document.alt_chunks().unwrap().remove(0);
        assert!(document.resolve_alt_chunk(&chunk).is_err());
    }

    fn synthetic_package(relationship: &str, external: bool, include_target: bool) -> Package {
        let mut opc = OpcPackage::new();
        opc.relate_to("word/document.xml", relationship_type::OFFICE_DOCUMENT);
        let mut document = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            content_type::WML_DOCUMENT_MAIN.into(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="chunk"/></w:body></w:document>"#.to_vec(),
        );
        document.rels_mut().add_relationship(
            relationship.into(),
            if external { "https://example.invalid/chunk" } else { "chunk.html" }.into(),
            "chunk".into(),
            external,
        );
        opc.add_part(Box::new(document));
        if include_target && !external {
            opc.add_part(Box::new(BlobPart::new(
                PackURI::new("/word/chunk.html").unwrap(),
                "text/html".into(),
                b"<p>opaque</p>".to_vec(),
            )));
        }
        Package::from_opc_package(opc).unwrap()
    }
}
