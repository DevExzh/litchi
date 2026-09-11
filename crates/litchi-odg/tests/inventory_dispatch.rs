use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;
use litchi_odg::Drawing;

fn drawing(body: &str) -> Result<Drawing> {
    let content = format!(
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:drawing><d:page d:name="Page">{body}</d:page></office:drawing></office:body></office:document-content>"#
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.graphics")?;
    writer.add_file("content.xml", content.as_bytes())?;
    Drawing::from_bytes(writer.finish_to_bytes()?)
}

#[test]
fn optional_inventory_observes_local_aliases_and_default_namespaces() -> Result<()> {
    let source = drawing(
        r#"<d:custom-shape><enhanced-geometry xmlns="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" d:type="rectangle"><equation d:name="f0" d:formula="width/2"/></enhanced-geometry></d:custom-shape><d:rect><g:glue-point xmlns:g="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" g:id="1" g:escape-direction="auto" s:x="1cm" s:y="2cm"/></d:rect>"#,
    )?;
    let shapes = source.pages()[0].shapes();
    assert_eq!(
        shapes[0]
            .enhanced_geometry()
            .map(|value| value.children().len()),
        Some(1)
    );
    assert_eq!(shapes[1].glue_points().len(), 1);
    assert_eq!(shapes[1].glue_points()[0].id(), "1");
    Ok(())
}

#[test]
fn optional_inventory_still_rejects_misplaced_aliased_owners() {
    for body in [
        r#"<g:glue-point xmlns:g="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" g:id="1" g:escape-direction="auto" s:x="1cm" s:y="2cm"/>"#,
        r#"<d:rect><image-map xmlns="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"/></d:rect>"#,
        r#"<d:rect><q:contour-path xmlns:q="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" q:recreate-on-edit="false" s:viewBox="0 0 10 10" s:d="M0 0"/></d:rect>"#,
        r#"<d:custom-shape><d:enhanced-geometry><d:equation d:name="f0" d:formula="width"><d:rect/></d:equation></d:enhanced-geometry></d:custom-shape>"#,
    ] {
        assert!(
            drawing(body).is_err(),
            "misplaced owner was accepted: {body}"
        );
    }
}

#[test]
fn foreign_namespace_markers_remain_opaque_and_exact() -> Result<()> {
    let source = drawing(
        r#"<d:rect><d:glue-point xmlns:d="urn:example:foreign" value="opaque"/></d:rect><!-- enhanced-geometry image-map dr3d:scene -->"#,
    )?;
    assert!(source.pages()[0].shapes()[0].glue_points().is_empty());
    let committed = source.edit().commit()?;
    assert!(!committed.changed());
    assert_eq!(committed.snapshot().as_bytes(), source.as_bytes());
    Ok(())
}
