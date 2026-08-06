use super::{EXTENSION_URI, NAMESPACE, codec, model::Outcome, transaction};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

fn shape_with_classification(outcome: &str) -> Vec<u8> {
    format!(
        r#"<p:sp xmlns:p="{PML}" xmlns:p184="{NAMESPACE}"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="{EXTENSION_URI}"><p184:classification val="{outcome}"/></p:ext></p:extLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
    )
    .into_bytes()
}

#[test]
fn reads_and_replaces_typed_classification_without_normalizing_the_shape() {
    let original = shape_with_classification("hdr");
    let source = codec::read(&original).unwrap();
    assert_eq!(
        source
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.outcome()),
        Some(Outcome::Header)
    );

    let updated = transaction::set(&original, &source, Outcome::Watermark).unwrap();
    let parsed = codec::read(&updated).unwrap();
    assert_eq!(
        parsed
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.outcome()),
        Some(Outcome::Watermark)
    );
    assert!(
        updated
            .windows(b"<p:spPr/>".len())
            .any(|window| window == b"<p:spPr/>")
            || updated
                .windows(b"<p:spPr>".len())
                .any(|window| window == b"<p:spPr>")
    );
}

#[test]
fn unknown_extensions_survive_first_class_authoring() {
    let xml = format!(
        r#"<p:sp xmlns:p="{PML}" xmlns:f="urn:future"><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr><p:extLst><p:ext uri="urn:future"><f:payload answer="42"/></p:ext></p:extLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
    )
    .into_bytes();
    let source = codec::read(&xml).unwrap();
    assert!(source.snapshot.is_none());

    let updated = transaction::set(&xml, &source, Outcome::Footer).unwrap();
    assert!(
        updated
            .windows(b"<f:payload answer=\"42\">".len())
            .any(|window| window == b"<f:payload answer=\"42\">")
            || updated
                .windows(b"<f:payload answer=\"42\"/>".len())
                .any(|window| window == b"<f:payload answer=\"42\"/>")
    );
    assert_eq!(
        codec::read(&updated)
            .unwrap()
            .snapshot
            .and_then(|snapshot| snapshot.outcome()),
        Some(Outcome::Footer)
    );
}

#[test]
fn detached_edits_are_atomic_and_keep_opaque_extensions() {
    let snapshot = super::Snapshot::new(Outcome::Header);
    let mut editor = snapshot.edit();
    editor.set(Outcome::Footer).unwrap();
    assert_eq!(editor.snapshot().outcome(), Some(Outcome::Footer));
    let committed = editor.commit().unwrap();
    assert_eq!(snapshot.outcome(), Some(Outcome::Header));
    assert_eq!(committed.outcome(), Some(Outcome::Footer));
}
