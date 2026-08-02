use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const RML: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Build a package whose `/ppt/presentation.xml` has the given blob and run
/// the presentation-level accessor battery over it. Any panic inside the
/// accessors fails the test; typed errors are returned for inspection.
fn with_presentation_blob(blob: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(blob.to_vec());
    package
}

/// The accessor battery every malformed blob is pushed through.
type AccessorResults = (
    litchi_ooxml::error::Result<Vec<u32>>,
    litchi_ooxml::error::Result<litchi_ooxml::pptx::sections::SectionList>,
    litchi_ooxml::error::Result<Option<(i64, i64)>>,
);

fn exercise(blob: &[u8]) -> AccessorResults {
    let package = with_presentation_blob(blob);
    let presentation = package.presentation().unwrap();
    (
        presentation.slide_ids(),
        presentation.sections(),
        presentation.slide_size(),
    )
}

#[test]
fn truncated_presentation_xml_fails_with_typed_errors() {
    let cases: Vec<Vec<u8>> = vec![
        Vec::new(),
        b"<p:presentation".to_vec(),
        format!(r#"<p:presentation xmlns:p="{PML}">"#).into_bytes(),
        format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{RML}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/>"#).into_bytes(),
        format!(r#"<p:presentation xmlns:p="{PML}"><p:sldIdLst></p:presentation>"#).into_bytes(),
        format!(r#"<p:presentation xmlns:p="{PML}"><!-- unterminated comment"#).into_bytes(),
        format!(r#"<p:presentation xmlns:p="{PML}"><p:sldSz cx="9144000" cy="#).into_bytes(),
        b"\xff\xfe\x00not xml\x00".to_vec(),
    ];
    for blob in cases {
        let (ids, sections, size) = exercise(&blob);
        assert!(ids.is_err(), "accepted truncated blob: {blob:?}");
        assert!(sections.is_err(), "accepted truncated blob: {blob:?}");
        assert!(size.is_err(), "accepted truncated blob: {blob:?}");
    }
}

#[test]
fn structural_violations_fail_with_typed_errors() {
    let cases: Vec<Vec<u8>> = vec![
        // Wrong root element.
        format!(r#"<p:deck xmlns:p="{PML}"/>"#).into_bytes(),
        // Content before the root element.
        format!(r#"<x/><p:presentation xmlns:p="{PML}"/>"#).into_bytes(),
        // Duplicate slide ID list.
        format!(
            r#"<p:presentation xmlns:p="{PML}"><p:sldIdLst/><p:sldIdLst/></p:presentation>"#
        )
        .into_bytes(),
        // Slide ID below the 256 minimum (ECMA-376 §19.2.1.33).
        format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{RML}"><p:sldIdLst><p:sldId id="12" r:id="rId2"/></p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
        // Slide ID overflowing u32.
        format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{RML}"><p:sldIdLst><p:sldId id="99999999999999999999" r:id="rId2"/></p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
        // Duplicate slide IDs in one list.
        format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{RML}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="256" r:id="rId3"/></p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
        // Missing relationship ID on a slide reference.
        format!(
            r#"<p:presentation xmlns:p="{PML}"><p:sldIdLst><p:sldId id="256"/></p:sldIdLst></p:presentation>"#
        )
        .into_bytes(),
    ];
    for blob in cases {
        let (ids, _, _) = exercise(&blob);
        assert!(ids.is_err(), "accepted invalid blob: {blob:?}");
    }
}

#[test]
fn entity_and_doctype_payloads_are_inert() {
    let blob = format!(
        r#"<!DOCTYPE p:presentation [<!ENTITY laugh "ha">]><p:presentation xmlns:p="{PML}"><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
    );
    let (ids, _, size) = exercise(blob.as_bytes());
    // No entity expansion happens; the document itself is well-formed.
    assert!(ids.is_ok());
    assert!(size.is_ok());
}

#[test]
fn excessive_nesting_depth_is_rejected() {
    let mut blob = format!(r#"<p:presentation xmlns:p="{PML}">"#).into_bytes();
    blob.extend_from_slice(b"<p:ext>".repeat(100_000).as_slice());
    blob.extend_from_slice(b"</p:ext>".repeat(100_000).as_slice());
    blob.extend_from_slice(b"</p:presentation>");
    let (ids, _, _) = exercise(&blob);
    assert!(ids.is_err(), "deeply nested presentation XML was accepted");
}

#[test]
fn excessive_element_count_is_rejected() {
    let mut blob = format!(r#"<p:presentation xmlns:p="{PML}"><p:extLst>"#).into_bytes();
    blob.extend_from_slice(b"<p:ext/>".repeat(2_000_000).as_slice());
    blob.extend_from_slice(b"</p:extLst></p:presentation>");
    let (ids, _, _) = exercise(&blob);
    assert!(ids.is_err(), "oversized presentation XML was accepted");
}

#[test]
fn slide_part_parsers_reject_malformed_xml_without_panics() {
    let cases: Vec<Vec<u8>> = vec![
        Vec::new(),
        b"<p:sld".to_vec(),
        b"\x00\x01\x02".to_vec(),
        format!(r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree><p:sp"#).into_bytes(),
        format!(r#"<p:sld xmlns:p="{PML}"><p:transition><x:bogus/></p:transition></p:sld>"#)
            .into_bytes(),
        format!(r#"<p:sld xmlns:p="{PML}"><p:bg><p:bgPr"#).into_bytes(),
    ];
    for blob in cases {
        // Corrupt the slide part in an otherwise valid package and run the
        // slide-level accessor battery: Ok or typed Err, never a panic.
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
            .opc_package_mut()
            .get_part_mut(&slide_uri)
            .unwrap()
            .set_blob(blob.clone());
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        let slide = &slides[0];
        let _ = slide.text();
        let _ = slide.shapes();
        let _ = slide.transition();
        let _ = slide.background();
        let _ = slide.animations();
    }
}

/// Corrupt the slide part in an otherwise valid one-slide package and run
/// the slide-level accessor battery over the given blob.
fn exercise_slide(
    blob: &[u8],
) -> (
    litchi_ooxml::error::Result<String>,
    litchi_ooxml::error::Result<usize>,
    litchi_ooxml::error::Result<Option<litchi_pptx::transition::Transition>>,
) {
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
        .opc_package_mut()
        .get_part_mut(&slide_uri)
        .unwrap()
        .set_blob(blob.to_vec());
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let slide = &slides[0];
    (
        slide.text(),
        slide.shapes().map(litchi_pptx::shape::Scene::len),
        slide.transition(),
    )
}

#[test]
fn slide_part_excessive_nesting_depth_is_rejected() {
    let mut blob = format!(r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree>"#).into_bytes();
    blob.extend_from_slice(b"<p:grpSp>".repeat(100_000).as_slice());
    blob.extend_from_slice(b"</p:grpSp>".repeat(100_000).as_slice());
    blob.extend_from_slice(b"</p:spTree></p:cSld></p:sld>");
    let (_, shapes, _) = exercise_slide(&blob);
    assert!(shapes.is_err(), "deeply nested slide XML was accepted");
}

#[test]
fn slide_part_excessive_element_count_is_rejected() {
    let mut blob = format!(r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree>"#).into_bytes();
    blob.extend_from_slice(b"<p:sp/>".repeat(2_000_000).as_slice());
    blob.extend_from_slice(b"</p:spTree></p:cSld></p:sld>");
    let (_, shapes, _) = exercise_slide(&blob);
    assert!(shapes.is_err(), "oversized slide XML was accepted");
}

#[test]
fn transition_excessive_nesting_depth_is_rejected() {
    let mut blob = format!(r#"<p:sld xmlns:p="{PML}"><p:transition>"#).into_bytes();
    blob.extend_from_slice(b"<p:ext>".repeat(100_000).as_slice());
    blob.extend_from_slice(b"</p:ext>".repeat(100_000).as_slice());
    blob.extend_from_slice(b"</p:transition></p:sld>");
    let (_, _, transition) = exercise_slide(&blob);
    assert!(
        transition.is_err(),
        "deeply nested transition XML was accepted"
    );
}
