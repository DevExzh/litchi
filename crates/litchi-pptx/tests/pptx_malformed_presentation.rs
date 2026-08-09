#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::PackURI;
use litchi_pptx::parts::SlideReference;
use litchi_pptx::{Error, Package, Result};
use tempfile::NamedTempFile;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const RML: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

type SlideReferences = Result<Vec<SlideReference>>;
type SlideSize = Result<(i64, i64)>;

/// Replace one canonical part in a valid package without asking the typed
/// presentation reader to accept the replacement during the edit itself.
fn with_presentation_blob(blob: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)?.set_blob(blob.to_vec());
            Ok(())
        })
        .unwrap();
    package
}

fn presentation_accessors(blob: &[u8]) -> (SlideReferences, SlideSize) {
    let package = with_presentation_blob(blob);
    let presentation = match package.presentation() {
        Ok(presentation) => presentation,
        Err(error) => {
            return (
                Err(error),
                Err(Error::Invalid("presentation root was rejected".into())),
            );
        },
    };
    (presentation.slide_references(), presentation.slide_size())
}

#[test]
fn malformed_presentation_roots_return_typed_errors() {
    let cases = [
        Vec::new(),
        b"<p:presentation".to_vec(),
        format!(r#"<p:deck xmlns:p="{PML}"/>"#).into_bytes(),
        format!(r#"<x/><p:presentation xmlns:p="{PML}"/>"#).into_bytes(),
        format!(r#"<p:presentation xmlns:p="{PML}"><p:sldIdLst><p:sldId id="256"/>"#).into_bytes(),
        b"\xff\xfe\x00not xml\x00".to_vec(),
    ];

    for blob in cases {
        let (references, size) = presentation_accessors(&blob);
        assert!(references.is_err(), "accepted malformed blob: {blob:?}");
        assert!(size.is_err(), "accepted malformed blob: {blob:?}");
    }
}

#[test]
fn malformed_presentation_children_are_reported_by_their_owner() {
    let missing_size = format!(r#"<p:presentation xmlns:p="{PML}"/>"#);
    let (references, size) = presentation_accessors(missing_size.as_bytes());
    assert!(references.is_ok());
    assert!(size.is_err());

    let malformed_reference = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{RML}"><p:sldIdLst><p:sldId id="not-a-number" r:id="rId1"/></p:sldIdLst><p:sldSz cx="1" cy="1"/></p:presentation>"#
    );
    let (references, size) = presentation_accessors(malformed_reference.as_bytes());
    assert!(references.is_err());
    assert!(size.is_ok());
}

#[test]
fn malformed_slide_accessors_do_not_panic() {
    let cases = [
        Vec::new(),
        b"<p:sld".to_vec(),
        b"\x00\x01\x02".to_vec(),
        format!(r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree><p:sp"#).into_bytes(),
        format!(r#"<p:sld xmlns:p="{PML}"><p:bg><p:bgPr"#).into_bytes(),
    ];

    for blob in cases {
        let output = NamedTempFile::with_suffix(".pptx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .presentation_mut()
            .unwrap()
            .add_slide()
            .unwrap()
            .set_title("seed");
        package.save(output.path()).unwrap();

        let mut package = Package::open(output.path()).unwrap();
        let slide_uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        package
            .edit_opc(|opc| {
                opc.get_part_mut(&slide_uri)?.set_blob(blob.clone());
                Ok(())
            })
            .unwrap();

        match package.presentation() {
            Err(_) => {},
            Ok(presentation) => match presentation.slides() {
                Err(_) => {},
                Ok(slides) => {
                    let slide = &slides[0];
                    let _ = slide.text();
                    let _ = slide.shapes();
                },
            },
        }
    }
}
