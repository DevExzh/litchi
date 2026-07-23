use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

#[test]
fn presentation_exposes_the_default_notes_graph() {
    let package = Package::new().unwrap();
    let notes = package
        .presentation()
        .unwrap()
        .notes_graph()
        .unwrap()
        .unwrap();

    assert_eq!(notes.master.part_name, "/ppt/notesMasters/notesMaster1.xml");
    assert_eq!(notes.master.theme.part_name, "/ppt/theme/theme1.xml");
    assert!(notes.slides.is_empty());
}

#[test]
fn default_notes_graph_survives_save_and_reopen() {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let reopened = Package::open(output.path()).unwrap();
    let notes = reopened
        .presentation()
        .unwrap()
        .notes_graph()
        .unwrap()
        .unwrap();
    assert_eq!(notes.master.part_name, "/ppt/notesMasters/notesMaster1.xml");
    assert!(notes.slides.is_empty());
}
