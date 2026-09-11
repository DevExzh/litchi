use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;
use litchi_odg::{Drawing, Transition};

fn inherited_drawing(indirect: bool) -> Result<Drawing> {
    let content = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Parent page" draw:style-name="parent"/><draw:page draw:name="Child page" draw:style-name="child"/></office:drawing></office:body></office:document-content>"#;
    let intermediate = if indirect {
        r#"<style:style style:name="middle" style:family="drawing-page" style:parent-style-name="parent"/>"#
    } else {
        ""
    };
    let parent = if indirect { "middle" } else { "parent" };
    let styles = format!(
        r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:styles><style:style style:name="parent" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual"/></style:style>{intermediate}<style:style style:name="child" style:family="drawing-page" style:parent-style-name="{parent}"/></office:styles></office:document-styles>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.graphics")?;
    writer.add_file("content.xml", content.as_bytes())?;
    writer.add_file("styles.xml", styles.as_bytes())?;
    Drawing::from_bytes(writer.finish_to_bytes()?)
}

#[test]
fn refuses_transition_edit_that_would_change_an_inheriting_page() -> Result<()> {
    for indirect in [false, true] {
        let source = inherited_drawing(indirect)?;
        let mut transition = Transition::new();
        transition.set_transition_type(Some("automatic"))?;
        let mut edit = source.edit();
        assert!(matches!(
            edit.set_page_transition(0, Some(transition)),
            Err(litchi_core::Error::Unsupported(_))
        ));
        let unchanged = edit.commit()?;
        assert!(!unchanged.changed());
        assert_eq!(unchanged.snapshot().as_bytes(), source.as_bytes());
        assert_eq!(
            source.pages()[1]
                .transition()
                .and_then(Transition::transition_type),
            Some("manual")
        );
    }
    Ok(())
}

#[test]
fn inherited_shared_transition_noop_remains_exact() -> Result<()> {
    let source = inherited_drawing(true)?;
    let mut edit = source.edit();
    edit.set_page_transition(0, source.pages()[0].transition().cloned())?;
    let unchanged = edit.commit()?;
    assert!(!unchanged.changed());
    assert_eq!(unchanged.snapshot().as_bytes(), source.as_bytes());
    Ok(())
}

#[test]
fn child_transition_override_does_not_change_its_parent_page() -> Result<()> {
    let source = inherited_drawing(true)?;
    let mut transition = Transition::new();
    transition.set_transition_type(Some("automatic"))?;
    let mut edit = source.edit();
    edit.set_page_transition(1, Some(transition))?;
    let output = edit.commit()?.into_snapshot();
    assert_eq!(
        output.pages()[0]
            .transition()
            .and_then(Transition::transition_type),
        Some("manual")
    );
    assert_eq!(
        output.pages()[1]
            .transition()
            .and_then(Transition::transition_type),
        Some("automatic")
    );
    Ok(())
}
