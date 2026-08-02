use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use tempfile::NamedTempFile;

#[test]
fn reopened_deck_removes_notes_by_name_and_checked_index() {
    let source = authored_two_slide_deck();
    let mut package = Package::open(source.path()).unwrap();
    name_slide(&mut package, "/ppt/slides/slide1.xml", "Overview");
    name_slide(&mut package, "/ppt/slides/slide2.xml", "Appendix");

    let before_text = slide_text(&package);
    let before_slide_xml = slide_xml(&package);
    assert_eq!(
        slide_notes(&package),
        vec!["Overview secret", "Appendix secret"]
    );

    assert!(package.remove_notes("Overview").unwrap());
    assert!(!package.remove_notes("Overview").unwrap());
    let slides = package.presentation().unwrap().slides().unwrap();
    assert!(slides[0].notes().unwrap().is_none());
    assert_eq!(
        slides[1].notes().unwrap().as_deref(),
        Some("Appendix secret")
    );
    assert_eq!(slide_text(&package), before_text);
    assert_eq!(slide_xml(&package), before_slide_xml);
    let graph = package.notes().unwrap().unwrap();
    assert_eq!(graph.slides().len(), 1);
    assert_eq!(graph.slides()[0].owner(), "/ppt/slides/slide2.xml");

    assert!(package.remove_notes(1usize).unwrap());
    assert!(!package.remove_notes(1usize).unwrap());
    assert_eq!(package.clear_notes().unwrap(), 0);
    assert_no_speaker_notes(&package);
    assert_eq!(slide_text(&package), before_text);
    assert_eq!(slide_xml(&package), before_slide_xml);

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    assert_no_speaker_notes(&reopened);
    assert_eq!(slide_text(&reopened), before_text);
    assert_eq!(slide_xml(&reopened), before_slide_xml);
}

#[test]
fn reopened_deck_clears_all_notes_idempotently() {
    let source = authored_two_slide_deck();
    let mut package = Package::open(source.path()).unwrap();
    let before_text = slide_text(&package);

    assert_eq!(package.clear_notes().unwrap(), 2);
    assert_eq!(package.clear_notes().unwrap(), 0);
    assert_no_speaker_notes(&package);
    assert_eq!(slide_text(&package), before_text);

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    assert_no_speaker_notes(&reopened);
    assert_eq!(slide_text(&reopened), before_text);
}

#[test]
fn selector_failures_do_not_mutate_the_notes_graph() {
    let source = authored_two_slide_deck();
    let mut package = Package::open(source.path()).unwrap();
    name_slide(&mut package, "/ppt/slides/slide1.xml", "Overview");
    name_slide(&mut package, "/ppt/slides/slide2.xml", "Appendix");
    let before = package.notes().unwrap().unwrap();
    let before_parts = package.opc_package().part_count();

    assert!(matches!(
        package.remove_notes(2usize),
        Err(OoxmlError::Pptx(
            litchi_pptx::Error::SlideIndexOutOfBounds { index: 2, len: 2 }
        ))
    ));
    assert!(matches!(
        package.remove_notes("Missing"),
        Err(OoxmlError::Pptx(litchi_pptx::Error::SlideNameNotFound(name)))
            if name == "Missing"
    ));
    assert_eq!(package.opc_package().part_count(), before_parts);
    assert_eq!(package.notes().unwrap().unwrap(), before);

    name_slide(&mut package, "/ppt/slides/slide2.xml", "Overview");
    let ambiguous_before = package.notes().unwrap().unwrap();
    assert!(matches!(
        package.remove_notes("Overview"),
        Err(OoxmlError::Pptx(litchi_pptx::Error::AmbiguousSlideName {
            name,
            matches: 2
        })) if name == "Overview"
    ));
    assert_eq!(package.notes().unwrap().unwrap(), ambiguous_before);
}

#[test]
fn dirty_legacy_writer_is_rejected_without_reading_or_editing_notes() {
    let source = authored_two_slide_deck();
    let clean = Package::open(source.path()).unwrap();
    let graph = clean.notes().unwrap().unwrap();
    let mut package = Package::new().unwrap();
    package
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap()
        .set_notes("Keep me");

    assert!(matches!(
        package.notes(),
        Err(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation: "notes",
            reason,
        }) if reason.contains("reading or editing notes")
    ));
    assert!(matches!(
        package.presentation().unwrap().notes(),
        Err(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation: "notes",
            reason,
        }) if reason.contains("reading or editing notes")
    ));
    assert!(matches!(
        package.put_notes(graph),
        Err(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation: "put_notes",
            reason,
        }) if reason.contains("reading or editing notes")
    ));
    assert!(matches!(
        package.remove_notes(0usize),
        Err(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation: "remove_notes",
            ..
        })
    ));
    assert!(matches!(
        package.clear_notes(),
        Err(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation: "clear_notes",
            ..
        })
    ));
    assert_eq!(
        package
            .presentation_mut()
            .unwrap()
            .slide_mut(0)
            .unwrap()
            .notes(),
        Some("Keep me")
    );
}

fn authored_two_slide_deck() -> NamedTempFile {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    let presentation = package.presentation_mut().unwrap();
    let overview = presentation.add_slide().unwrap();
    overview.set_title("Overview body");
    overview.set_notes("Overview secret");
    let appendix = presentation.add_slide().unwrap();
    appendix.set_title("Appendix body");
    appendix.set_notes("Appendix secret");
    package.save(output.path()).unwrap();
    output
}

fn name_slide(package: &mut Package, slide_name: &str, name: &str) {
    let slide_name = PackURI::new(slide_name).unwrap();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let named = if xml.contains("<p:cSld>") {
        xml.replacen("<p:cSld>", &format!(r#"<p:cSld name="{name}">"#), 1)
    } else {
        let marker = " name=\"";
        let root = xml.find("<p:cSld ").unwrap();
        let end = root + xml[root..].find('>').unwrap();
        let value_start = root + xml[root..end].find(marker).unwrap() + marker.len();
        let value_end = value_start + xml[value_start..end].find('"').unwrap();
        let mut named = xml.to_owned();
        named.replace_range(value_start..value_end, name);
        named
    };
    slide.set_blob(named.into_bytes());
}

fn slide_text(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.text().unwrap())
        .collect()
}

fn slide_xml(package: &Package) -> Vec<Vec<u8>> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.part().part().blob().to_vec())
        .collect()
}

fn slide_notes(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.notes().unwrap().unwrap())
        .collect()
}

fn assert_no_speaker_notes(package: &Package) {
    let graph = package.notes().unwrap().unwrap();
    assert!(graph.slides().is_empty());
    assert!(
        package
            .opc_package()
            .contains_part(&PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap())
    );
    assert!(
        package
            .opc_package()
            .iter_parts()
            .all(|part| part.content_type() != ct::PML_NOTES_SLIDE)
    );
    assert!(package.opc_package().iter_parts().all(|part| {
        part.rels()
            .iter()
            .all(|relationship| relationship.reltype() != rt::NOTES_SLIDE)
    }));
    for slide in package.presentation().unwrap().slides().unwrap() {
        assert!(slide.notes().unwrap().is_none());
    }
}
