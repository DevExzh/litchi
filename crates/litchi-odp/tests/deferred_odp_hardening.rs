use litchi_core::Error;
use litchi_odp::FlatPresentation;

const SEED: &str = r#"<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" office:mimetype="application/vnd.oasis.opendocument.presentation"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" smil:type="fade"/></style:style></office:automatic-styles><office:body><office:presentation><draw:page draw:name="s1" draw:style-name="dp1"><draw:g draw:name="group"><draw:custom-shape><draw:enhanced-geometry draw:type="rectangle"><draw:equation draw:name="f1" draw:formula="width/2"/><draw:handle draw:handle-position="$0 0"/></draw:enhanced-geometry><text:p>custom</text:p></draw:custom-shape></draw:g><presentation:notes><text:p>speaker notes</text:p></presentation:notes></draw:page></office:presentation></office:body></office:document>"#;

fn exercise(bytes: Vec<u8>) -> Option<Error> {
    match FlatPresentation::from_bytes(bytes) {
        Ok(flat) => {
            let _ = flat.slides();
            for result in [
                flat.declarations().map(|_| ()),
                flat.layouts().map(|_| ()),
                flat.pages().map(|_| ()),
                flat.settings().map(|_| ()),
            ] {
                if let Err(error) = result {
                    return Some(error);
                }
            }
            None
        },
        Err(error) => Some(error),
    }
}

fn document(styles: &str, body: &str) -> Vec<u8> {
    format!(r#"<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:mimetype="application/vnd.oasis.opendocument.presentation"><office:automatic-styles>{styles}</office:automatic-styles><office:body><office:presentation>{body}</office:presentation></office:body></office:document>"#).into_bytes()
}

fn assert_typed(bytes: Vec<u8>) {
    assert!(matches!(exercise(bytes), Some(Error::InvalidFormat(_))));
}

#[test]
fn flat_odp_truncation_and_mutation_never_panic() {
    for end in 0..SEED.len() {
        exercise(SEED.as_bytes()[..end].to_vec());
    }
    for position in 0..SEED.len() {
        let mut bytes = SEED.as_bytes().to_vec();
        bytes[position] ^= 1;
        exercise(bytes);
    }
    assert!(exercise(SEED.as_bytes().to_vec()).is_none());
}

#[test]
fn flat_odp_malformed_inputs_are_typed_errors() {
    assert_typed(document(
        r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="not-a-type"/></style:style>"#,
        r#"<draw:page draw:name="s" draw:style-name="dp1"/>"#,
    ));
    assert_typed(document(
        r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-speed="warp"/></style:style>"#,
        r#"<draw:page draw:name="s" draw:style-name="dp1"/>"#,
    ));
    assert_typed(document(
        r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual"><presentation:sound/></style:drawing-page-properties></style:style>"#,
        r#"<draw:page draw:name="s" draw:style-name="dp1"/>"#,
    ));
    assert_typed(document("", r#"<draw:page><draw:g></draw:page></draw:g>"#));
    assert_typed(document("", r#"</office:presentation><draw:page/>"#));
    assert_typed(document("", r#"<draw:page>&undefined;</draw:page>"#));
    assert_typed(
        document("", r#"<draw:page/>"#)
            .into_iter()
            .chain([0xff])
            .collect(),
    );
}

#[test]
fn flat_odp_wellformed_foreign_content_is_inert() {
    let bytes = document(
        "",
        r#"<draw:page draw:name="s"><ext:item xmlns:ext="urn:example" ext:value="1"/></draw:page>"#,
    );
    let _ = exercise(bytes);
}
