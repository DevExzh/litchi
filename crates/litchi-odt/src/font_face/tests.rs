//! Unit tests for the font-face model and XML codec.

use super::*;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";

fn document(body: &str) -> String {
    format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:v="{SVG}" xmlns:x="{XLINK}">{body}<o:body/></o:document-content>"#
    )
}

#[test]
fn parses_and_round_trips_complete_font_face_metadata() {
    // ODF 1.2/1.3 style:font-face grammar; values mirror the declarations
    // emitted by LibreOffice and modeled by odfpy's style/svg constructors.
    let xml = document(
        r#"<o:font-face-decls><s:font-face s:name="Body &amp; Text" s:font-adornments="Regular" s:font-family-generic="swiss" s:font-pitch="variable" s:font-charset="UTF-8" v:font-family="'Liberation Sans'" v:font-style="italic" v:font-variant="small-caps" v:font-weight="700" v:font-stretch="semi-expanded" v:font-size="10.5pt" v:unicode-range="U+0-10FFFF" v:units-per-em="2048" v:ascent="1854" v:descent="-434" v:panose-1="2 11 6 4" v:widths="1 2" v:bbox="0 -434 2000 1854"><v:font-face-src><v:font-face-uri x:type="simple" x:href="Fonts/body.ttf" x:actuate="onRequest"><v:font-face-format v:string="truetype"/><v:font-face-format/></v:font-face-uri><v:font-face-name v:name="Liberation Sans"/></v:font-face-src><v:definition-src x:type="simple" x:href="Fonts/body.svg"/></s:font-face><s:font-face s:name="Empty"/></o:font-face-decls>"#,
    );
    let declarations = parse(&xml).unwrap().unwrap();
    assert_eq!(declarations.faces.len(), 2);
    let face = declarations.face("Body & Text").unwrap();
    assert_eq!(face.generic_family, Some(Family::Swiss));
    assert_eq!(face.weight, Some(Weight::Weight700));
    assert_eq!(face.size.as_ref().unwrap().as_str(), "10.5pt");
    assert_eq!(face.metrics.len(), 3);
    assert_eq!(face.sources.len(), 2);
    assert!(matches!(
        &face.sources[0],
        Source::Uri { formats, .. } if formats == &[Some("truetype".to_string()), None]
    ));

    let serialized = declarations.to_xml().unwrap();
    let reparsed = parse(&format!(
        r#"<office:document-content xmlns:office="{OFFICE}">{serialized}<office:body/></office:document-content>"#
    ))
    .unwrap()
    .unwrap();
    assert_eq!(reparsed, declarations);
}

#[test]
fn rejects_malformed_font_face_grammar() {
    for body in [
        r#"<o:font-face-decls><s:font-face/></o:font-face-decls>"#,
        r#"<o:font-face-decls><s:font-face s:name="A"/><s:font-face s:name="A"/></o:font-face-decls>"#,
        r#"<o:font-face-decls><s:font-face s:name="A" v:font-weight="950"/></o:font-face-decls>"#,
        r#"<o:font-face-decls><s:font-face s:name="A" v:font-size="0pt"/></o:font-face-decls>"#,
        r#"<o:font-face-decls><s:font-face s:name="A"><v:font-face-src/></s:font-face></o:font-face-decls>"#,
        r#"<o:font-face-decls><s:font-face s:name="A"><v:font-face-src><v:font-face-uri x:type="extended" x:href="x"/></v:font-face-src></s:font-face></o:font-face-decls>"#,
        r#"<o:font-face-decls><s:font-face s:name="A">active text</s:font-face></o:font-face-decls>"#,
    ] {
        assert!(parse(&document(body)).is_err(), "{body}");
    }
}

#[test]
fn rejects_misplaced_or_duplicate_containers() {
    assert!(parse(&document(r#"<o:body><o:font-face-decls/></o:body>"#)).is_err());
    assert!(parse(&document(r#"<o:font-face-decls/><o:font-face-decls/>"#)).is_err());
}
