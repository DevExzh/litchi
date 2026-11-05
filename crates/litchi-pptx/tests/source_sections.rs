#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the in-memory package fixtures and assertions intentionally panic"
)]

//! Source-backed presentation-section projections.

use std::io::{self, Cursor};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{Position, ReadAt, SourceVersion};
use litchi_opc::{OpcPackage, PackURI, PackageWriter, SourceLimits};
use litchi_pptx::presentation::sections::raw;
use litchi_pptx::{Error, SourceBackedPresentation};

const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

fn section_extension(slide_ids: &str) -> String {
    format!(
        r#"<p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section name="Opening authored" id="{{11111111-1111-1111-1111-111111111111}}"><p14:sldIdLst>{slide_ids}</p14:sldIdLst><p:extLst><p:ext uri="urn:litchi:opaque-section"><x:payload xmlns:x="urn:litchi:opaque"/></p:ext></p:extLst></p14:section></p14:sectionLst></p:ext>"#
    )
}

fn source_bytes(slide_ids: &str) -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    let presentation = builder.presentation().unwrap();
    for index in 0..3 {
        presentation
            .add_slide()
            .unwrap()
            .set_title(&format!("Slide {index}"));
    }
    let package = builder.build().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let uri = PackURI::new("/ppt/presentation.xml").unwrap();
    let part = opc.get_part_mut(&uri).unwrap();
    let xml = String::from_utf8(part.blob().to_vec()).unwrap();
    let extension = section_extension(slide_ids);
    let replacement = if xml.contains("</p:extLst>") {
        xml.replacen("</p:extLst>", &format!("{extension}</p:extLst>"), 1)
    } else {
        xml.replacen(
            "</p:presentation>",
            &format!("<p:extLst>{extension}</p:extLst></p:presentation>"),
            1,
        )
    };
    assert_ne!(replacement, xml);
    part.set_blob(replacement.into_bytes());
    PackageWriter::to_bytes(&opc).unwrap()
}

fn valid_source_bytes() -> Vec<u8> {
    source_bytes(r#"<p14:sldId id="256"/><p14:sldId id="258"/>"#)
}

#[test]
fn source_sections_project_positions_and_keep_native_values_raw() {
    let source = SourceBackedPresentation::from_reader(Cursor::new(valid_source_bytes())).unwrap();
    let sections = source.sections().unwrap();

    assert_eq!(sections.len(), 1);
    assert!(!sections.is_empty());
    assert_eq!(sections.entries().len(), 1);
    let section = sections.section(Position::new(0)).unwrap();
    assert_eq!(section.position(), Position::new(0));
    assert_eq!(section.name(), Some("Opening authored"));
    assert_eq!(section.slides(), [Position::new(0), Position::new(2)]);
    assert_eq!(
        raw::id(section),
        Some("{11111111-1111-1111-1111-111111111111}")
    );
    assert_eq!(raw::native_slide_ids(section), [256, 258]);
    assert!(raw::extension_xml(section).is_some());
    assert!(sections.section(Position::new(1)).is_none());

    let debug = format!("{sections:?} {:?}", section);
    assert!(!debug.contains("Opening authored"));
    assert!(!debug.contains("11111111"));
}

#[test]
fn source_sections_are_detached_and_clone_cheaply() {
    let source = SourceBackedPresentation::from_reader(Cursor::new(valid_source_bytes())).unwrap();
    let sections = source.sections().unwrap();
    let clone = sections.clone();
    drop(source);

    assert_eq!(
        clone.entries()[0].slides(),
        [Position::new(0), Position::new(2)]
    );
    assert_eq!(sections.entries()[0].name(), Some("Opening authored"));
}

#[test]
fn unknown_section_slide_id_is_refused_without_native_id_fallback() {
    let source = SourceBackedPresentation::from_reader(Cursor::new(source_bytes(
        r#"<p14:sldId id="999999"/>"#,
    )))
    .unwrap();
    let error = source.sections().unwrap_err();
    assert!(matches!(error, Error::Invalid(_)));
    assert!(!error.to_string().contains("999999"));
}

#[test]
fn wrong_section_namespace_is_refused_strictly() {
    let bytes = valid_source_bytes();
    let mut opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let uri = PackURI::new("/ppt/presentation.xml").unwrap();
    let part = opc.get_part_mut(&uri).unwrap();
    let xml = String::from_utf8(part.blob().to_vec()).unwrap();
    part.set_blob(
        xml.replace(P14, "urn:litchi:wrong-section-namespace")
            .into_bytes(),
    );
    let source =
        SourceBackedPresentation::from_reader(Cursor::new(PackageWriter::to_bytes(&opc).unwrap()))
            .unwrap();
    assert!(matches!(source.sections(), Err(Error::Invalid(_))));
}

#[test]
fn source_change_refuses_new_acquisition_but_keeps_old_snapshot() {
    let source = Arc::new(VersionedSource::new(valid_source_bytes()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let sections = presentation.sections().unwrap();
    source.changed();

    assert!(matches!(
        presentation.sections(),
        Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
    ));
    assert_eq!(
        sections.entries()[0].slides(),
        [Position::new(0), Position::new(2)]
    );
}

#[test]
fn selected_source_limits_are_not_bypassed_by_sections() {
    let bytes = valid_source_bytes();
    let limits = SourceLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let error = match SourceBackedPresentation::from_reader_with_options(
        Cursor::new(bytes),
        litchi_pptx::OpenOptions::default().source_limits(limits),
    ) {
        Ok(_) => panic!("an over-tight source limit unexpectedly opened the package"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(litchi_opc::OpcError::Limit { .. })
    ));
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(61, self.revision.load(Ordering::SeqCst)))
    }
}
