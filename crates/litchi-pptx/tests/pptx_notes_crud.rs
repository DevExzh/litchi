#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::notes;
use litchi_pptx::{Error, Package};
use tempfile::NamedTempFile;

#[test]
fn reopened_deck_removes_notes_by_semantic_slide_and_preserves_slide_content() {
    let source = authored_two_slide_deck();
    let mut package = Package::open(source.path()).unwrap();
    package = name_slide(package, "/ppt/slides/slide1.xml", "Overview");
    package = name_slide(package, "/ppt/slides/slide2.xml", "Appendix");

    let before_text = slide_text(&package);
    let before_slide_xml = slide_xml(&package);
    assert_eq!(
        slide_notes(&package),
        vec!["Overview secret", "Appendix secret"]
    );

    let (next, removed) = remove_notes(package, "Overview");
    package = next;
    assert!(removed.unwrap());
    let (next, removed) = remove_notes(package, "Overview");
    package = next;
    assert!(!removed.unwrap());
    assert_eq!(slide_notes(&package), vec!["", "Appendix secret"]);

    let graph = notes_graph(&package).unwrap();
    assert_eq!(graph.slides().len(), 1);
    assert_eq!(graph.slides()[0].owner(), "/ppt/slides/slide2.xml");

    let (next, removed) = remove_notes(package, "Appendix");
    package = next;
    assert!(removed.unwrap());
    let (next, removed) = clear_notes(package);
    package = next;
    assert_eq!(removed.unwrap(), 0);
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
    let package = Package::open(source.path()).unwrap();
    let before_text = slide_text(&package);

    let (next, removed) = clear_notes(package);
    let package = next;
    assert_eq!(removed.unwrap(), 2);
    let (next, removed) = clear_notes(package);
    let package = next;
    assert_eq!(removed.unwrap(), 0);
    assert_no_speaker_notes(&package);
    assert_eq!(slide_text(&package), before_text);

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = package;
    package.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    assert_no_speaker_notes(&reopened);
    assert_eq!(slide_text(&reopened), before_text);
}

#[test]
fn selector_failures_do_not_mutate_the_notes_graph() {
    let source = authored_two_slide_deck();
    let mut package = Package::open(source.path()).unwrap();
    package = name_slide(package, "/ppt/slides/slide1.xml", "Overview");
    package = name_slide(package, "/ppt/slides/slide2.xml", "Appendix");
    let before = notes_graph(&package).unwrap();
    let before_parts = package.opc().unwrap().part_count();

    // Slide selection belongs to the semantic Presentation facade; the notes
    // owner intentionally receives the already-resolved physical slide URI.
    let presentation = package.presentation().unwrap();
    assert!(presentation.find_slide(2usize).unwrap().is_none());
    assert!(presentation.find_slide("Missing").unwrap().is_none());
    assert_eq!(package.opc().unwrap().part_count(), before_parts);
    assert_eq!(notes_graph(&package).unwrap(), before);

    package = name_slide(package, "/ppt/slides/slide2.xml", "Overview");
    let ambiguous = match package.presentation().unwrap().find_slide("Overview") {
        Ok(_) => panic!("duplicate slide names must be rejected"),
        Err(error) => error,
    };
    assert!(matches!(
        ambiguous,
        Error::AmbiguousSlideName { name, matches: 2 } if name == "Overview"
    ));
    assert_eq!(notes_graph(&package).unwrap(), before);
}

#[test]
fn dirty_writer_is_rejected_before_direct_notes_owner_access() {
    let source = authored_two_slide_deck();
    let clean = Package::open(source.path()).unwrap();
    assert!(notes_graph(&clean).is_some());

    let mut package = Package::new().unwrap();
    package
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap()
        .set_notes("Keep me");

    // The standalone notes owner consumes a canonical OPC snapshot. The
    // package facade therefore rejects access while its mutable writer is
    // dirty, without adding a compatibility-specific notes API.
    let error = match package.opc() {
        Ok(_) => panic!("dirty writer unexpectedly exposed an OPC snapshot"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::UnsafeEdit { operation: "opc", reason } if reason.contains("unflushed")
    ));
    let error = match package.presentation() {
        Ok(_) => panic!("dirty writer unexpectedly exposed a presentation graph"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::UnsafeEdit { operation: "presentation", reason } if reason.contains("unflushed")
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

fn name_slide(package: Package, slide_name: &str, name: &str) -> Package {
    let slide_name = PackURI::new(slide_name).unwrap();
    let (package, result) = edit_package(package, |opc| {
        let slide = opc.get_part_mut(&slide_name)?;
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
        Ok(())
    });
    result.unwrap();
    package
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
    let graph = package.notes().unwrap().unwrap();
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            graph
                .slides()
                .iter()
                .find(|notes| notes.owner() == slide.part().part().partname().as_str())
                .and_then(|slide| slide.text().unwrap())
                .unwrap_or_default()
        })
        .collect()
}

fn notes_graph(package: &Package) -> Option<notes::Graph> {
    package.notes().unwrap()
}

fn assert_no_speaker_notes(package: &Package) {
    let graph = notes_graph(package).unwrap();
    assert!(graph.slides().is_empty());
    assert!(
        package
            .opc()
            .unwrap()
            .contains_part(&PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap())
    );
    assert!(
        package
            .opc()
            .unwrap()
            .iter_parts()
            .all(|part| part.content_type() != ct::PML_NOTES_SLIDE)
    );
    assert!(package.opc().unwrap().iter_parts().all(|part| {
        part.rels()
            .iter()
            .all(|relationship| relationship.reltype() != rt::NOTES_SLIDE)
    }));
}

fn remove_notes(package: Package, slide_name: &str) -> (Package, litchi_pptx::Result<bool>) {
    let mut package = package;
    let result = package.remove_notes(slide_name);
    (package, result)
}

fn clear_notes(package: Package) -> (Package, litchi_pptx::Result<usize>) {
    let mut package = package;
    let result = package.clear_notes();
    (package, result)
}

fn edit_package<T>(
    mut package: Package,
    edit: impl FnOnce(&mut OpcPackage) -> litchi_pptx::Result<T>,
) -> (Package, litchi_pptx::Result<T>) {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    let result = edit(&mut opc);
    let package = Package::from_opc_package(opc).unwrap();
    (package, result)
}
