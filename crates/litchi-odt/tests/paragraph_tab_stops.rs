use litchi_odt::style::paragraph::tab_stop::{
    LeaderColor, LeaderStyle, LeaderType, LeaderWidth, Position, Stop, Stops, Style, Type, parse,
};
use litchi_odt::{Builder, Document};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

fn wrap(body: &str) -> String {
    format!(r#"<o:styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}">{body}</o:styles>"#)
}

#[test]
fn parses_aliases_all_values_and_deterministic_round_trip() {
    let xml = wrap(
        r##"<s:default-style s:family="paragraph"><s:paragraph-properties><s:tab-stops><s:tab-stop s:position="1cm"/></s:tab-stops></s:paragraph-properties></s:default-style><s:style s:name="Parent" s:family="paragraph"><s:paragraph-properties><s:tab-stops><s:tab-stop s:position="8.5cm" s:type="center"/><s:tab-stop s:position="17cm" s:type="char" s:char="," s:leader-type="double" s:leader-style="dot-dash" s:leader-width="25%" s:leader-color="#a0B1c2" s:leader-text="." s:leader-text-style="Leader"/></s:tab-stops></s:paragraph-properties></s:style><s:style s:name="Child" s:family="paragraph" s:parent-style-name="Parent"/><s:style s:name="Clear" s:family="paragraph" s:parent-style-name="Parent"><s:paragraph-properties><s:tab-stops/></s:paragraph-properties></s:style>"##,
    );
    let parsed = parse(&xml).unwrap();
    assert_eq!(parsed.styles.len(), 4);
    let parent = parsed.get("Parent").unwrap();
    let stops = parent.tab_stops.as_ref().unwrap();
    assert_eq!(stops.len(), 2);
    assert_eq!(stops.as_slice()[0].tab_type, Type::Center);
    assert_eq!(stops.as_slice()[1].tab_type, Type::Character(','));
    assert_eq!(
        stops.as_slice()[1].leader_color,
        Some(LeaderColor::Rgb(160, 177, 194))
    );
    assert_eq!(parsed.resolved_tab_stops("Child").unwrap().unwrap(), stops);
    assert_eq!(
        parsed.resolved_tab_stops("Clear").unwrap().unwrap().len(),
        0
    );
    assert_eq!(
        parsed.resolved_tab_stops("Missing").unwrap().unwrap().len(),
        1
    );

    let fragment = parent.to_xml_fragment().unwrap();
    assert_eq!(
        fragment,
        concat!(
            r#"<style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:family="paragraph" style:name="Parent"><style:paragraph-properties><style:tab-stops>"#,
            r#"<style:tab-stop style:position="8.5cm" style:type="center"/>"#,
            r##"<style:tab-stop style:leader-color="#A0B1C2" style:leader-style="dot-dash" style:leader-text="." style:leader-text-style="Leader" style:leader-type="double" style:leader-width="25%" style:position="17cm" style:char="," style:type="char"/>"##,
            "</style:tab-stops></style:paragraph-properties></style:style>"
        )
    );
    let reparsed = parse(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.get("Parent"), Some(parent));
}

#[test]
fn parses_odfpy_and_libreoffice_reference_documents() {
    let odfpy = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/odfpy/xml2odf/definitionlists.xml"
    ));
    let parsed = parse(odfpy).unwrap();
    assert!(parsed.styles.iter().any(|style| {
        style
            .tab_stops
            .as_ref()
            .is_some_and(|stops| stops.iter().any(|stop| stop.position.as_str() == "0cm"))
    }));

    let libreoffice = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/extras/source/autotext/lang/cs/template/HLC/styles.xml"
    ));
    let parsed = parse(libreoffice).unwrap();
    assert!(parsed.styles.iter().any(|style| {
        style.tab_stops.as_ref().is_some_and(|stops| {
            stops
                .iter()
                .any(|stop| stop.position.as_str() == "8.5cm" && stop.tab_type == Type::Center)
        })
    }));
}

#[test]
fn rejects_malformed_structure_values_and_overflow() {
    let invalid = [
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:tab-stops><s:tab-stop/></s:tab-stops></s:paragraph-properties></s:style>"#,
        ),
        wrap(r#"<s:style s:name="x" s:family="paragraph"><s:tab-stops/></s:style>"#),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:tab-stops/><s:tab-stops/></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:tab-stops><s:tab-stop s:position="1em"/></s:tab-stops></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:tab-stops><s:tab-stop s:position="1cm" s:type="char"/></s:tab-stops></s:paragraph-properties></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:tab-stops><s:tab-stop s:position="1cm"><s:tab-stop s:position="2cm"/></s:tab-stop></s:tab-stops></s:paragraph-properties></s:style>"#,
        ),
        format!(
            r#"<o:styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:x="urn:wrong"><s:style s:name="x" s:family="paragraph"><s:paragraph-properties><x:tab-stops/></s:paragraph-properties></s:style></o:styles>"#
        ),
        format!(
            r#"<!DOCTYPE x><o:styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}"><s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:tab-stops/></s:paragraph-properties></s:style></o:styles>"#
        ),
    ];
    for xml in invalid {
        assert!(parse(&xml).is_err(), "accepted {xml}");
    }

    let stops = (0..65)
        .map(|index| format!(r#"<s:tab-stop s:position="{index}cm"/>"#))
        .collect::<String>();
    let overflow = wrap(&format!(
        r#"<s:style s:name="x" s:family="paragraph"><s:paragraph-properties><s:tab-stops>{stops}</s:tab-stops></s:paragraph-properties></s:style>"#
    ));
    assert!(parse(&overflow).is_err());
}

#[test]
fn builder_package_and_mutable_document_preserve_typed_styles() {
    let mut first = Stop::new(Position::new("2.5cm").unwrap());
    first.tab_type = Type::Right;
    first.leader_type = Some(LeaderType::Single);
    first.leader_style = Some(LeaderStyle::Dotted);
    first.leader_width = Some(LeaderWidth::new("thin").unwrap());
    first.leader_color = Some(LeaderColor::FontColor);
    let stops = Stops::try_from_vec(vec![first]).unwrap();
    let mut style = Style::named("ReportTabs", Some(stops)).unwrap();
    style.parent_style_name = Some("Standard".to_owned());

    let mut builder = Builder::new();
    builder.add_paragraph_tab_style(style.clone()).unwrap();
    builder.add_paragraph("tabbed").unwrap();
    let bytes = builder.build().unwrap();
    let package = litchi_odt::generic::OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
    assert_eq!(
        package
            .paragraph_style_tab_stops()
            .unwrap()
            .get("ReportTabs"),
        Some(&style)
    );

    let document = Document::from_bytes(bytes).unwrap();
    let mutable = litchi_odt::mutable::MutableDocument::from_document(document).unwrap();
    let round_trip = mutable.to_bytes().unwrap();
    let package = litchi_odt::generic::OpenDocumentPackage::from_bytes(round_trip).unwrap();
    assert_eq!(
        package
            .paragraph_style_tab_stops()
            .unwrap()
            .get("ReportTabs"),
        Some(&style)
    );
}
