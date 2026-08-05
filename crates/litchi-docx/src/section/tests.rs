use super::model::MAX_XML_BYTES;
use super::*;
use crate::header_footer::Kind;

#[test]
fn emu_conversions_and_checked_domain() {
    let inch = Emu::from_inches(1.0);
    assert_eq!(inch.get(), 914_400);
    assert_eq!(inch.to_twips(), 1440);
    assert_eq!(inch.try_to_twips().unwrap(), 1440);
    assert_eq!(Emu::from_twips(1440), inch);
    assert_eq!(Emu::from_cm(2.54), inch);
    assert_eq!(Emu::from_pt(72.0), inch);
    assert!(Emu::try_from_twips(i64::MAX).is_err());
    assert!(Emu::try_from_inches(f64::INFINITY).is_err());
    assert!(Emu::from_inches(0.5).try_to_twips().is_ok());
}

#[test]
fn enum_lexemes_and_defaults_are_contextual() {
    assert_eq!(Orientation::default(), Orientation::Portrait);
    assert_eq!(Orientation::Landscape.to_xml(), "landscape");
    assert_eq!(
        Orientation::from_xml("portrait"),
        Some(Orientation::Portrait)
    );
    assert_eq!(Orientation::from_xml("unknown"), None);
    assert_eq!(Start::default(), Start::NewPage);
    assert_eq!(Start::Continuous.to_xml(), "continuous");
    assert_eq!(Start::from_xml("nextColumn"), Some(Start::NewColumn));
    assert_eq!(Start::OddPage.to_string(), "Odd Page");
}

#[test]
fn reads_geometry_columns_references_and_preserves_unknown_content() {
    let xml = br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example">
  <w:headerReference w:type="default" r:id="rId1"/>
  <w:type w:val="continuous"/>
  <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
  <w:pgMar w:top="1440" w:right="1200" w:bottom="1440" w:left="1200" w:header="720" w:footer="720" w:gutter="0"/>
  <w:cols w:equalWidth="0" w:num="2"><w:col w:w="5000" w:space="240"/><w:col w:w="5000"/></w:cols>
  <!-- preserve this comment -->
  <x:extension x:value="keep"/>
</w:sectPr>"#;
    let mut section = Section::from_xml_bytes(xml.to_vec()).unwrap();
    assert_eq!(section.page_width(), Some(Emu::from_twips(12240)));
    assert_eq!(section.page_height(), Some(Emu::from_twips(15840)));
    assert_eq!(section.orientation(), Orientation::Portrait);
    assert_eq!(section.top_margin(), Some(Emu::from_twips(1440)));
    assert_eq!(section.right_margin(), Some(Emu::from_twips(1200)));
    assert_eq!(section.start_type(), Start::Continuous);

    let columns = section.columns().unwrap().unwrap();
    assert!(!columns.equal_width);
    assert_eq!(columns.count, 2);
    assert_eq!(columns.columns.len(), 2);
    assert_eq!(columns.columns[0].width, Emu::from_twips(5000));

    assert_eq!(
        section.headers().unwrap(),
        vec![Reference {
            kind: Kind::Primary,
            relationship_id: "rId1".into(),
        }]
    );
    assert_eq!(section.to_xml_bytes().unwrap(), xml);
}

#[test]
fn mutations_keep_foreign_children_and_round_trip() {
    let xml = br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:example"><x:before/><w:pgSz w:w="12240" w:h="15840"/><x:after/></w:sectPr>"#;
    let mut section = Section::from_xml_bytes(xml.to_vec()).unwrap();
    section
        .set_page_size(PageSize {
            width: Some(Emu::from_twips(11906)),
            height: Some(Emu::from_twips(16838)),
            orientation: Orientation::Landscape,
        })
        .unwrap();
    section
        .set_margins(Margins {
            top: Some(Emu::from_twips(720)),
            bottom: Some(Emu::from_twips(720)),
            left: Some(Emu::from_twips(1080)),
            right: Some(Emu::from_twips(1080)),
            header: Some(Emu::from_twips(360)),
            footer: Some(Emu::from_twips(360)),
            gutter: None,
        })
        .unwrap();
    let output = section.to_xml().unwrap();
    assert!(output.contains("<x:before/>"));
    assert!(output.contains("<x:after/>"));
    assert!(output.contains("w:w=\"11906\""));
    assert!(output.contains("w:orient=\"landscape\""));
    assert!(output.contains("w:top=\"720\""));

    let mut reparsed = Section::from_xml_bytes(output.into_bytes()).unwrap();
    assert_eq!(reparsed.page_width(), Some(Emu::from_twips(11906)));
    assert_eq!(reparsed.page_height(), Some(Emu::from_twips(16838)));
    assert_eq!(reparsed.orientation(), Orientation::Landscape);
    assert_eq!(reparsed.left_margin(), Some(Emu::from_twips(1080)));

    reparsed.clear_page_size().unwrap();
    assert!(reparsed.page_width().is_none());
    assert!(!reparsed.to_xml().unwrap().contains("pgSz"));
}

#[test]
fn section_inheritance_and_collection_crud_are_ordered() {
    let first = Section::from_xml_bytes(
        br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="720"/></w:sectPr>"#.to_vec(),
    )
    .unwrap();
    let second = Section::from_xml_bytes(
        br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgMar w:right="720"/></w:sectPr>"#.to_vec(),
    )
    .unwrap();
    let third = Section::default();
    let mut sections = Sections::new(vec![first.clone(), second, third]);
    let margins = sections.effective_margins(1).unwrap();
    assert_eq!(margins.left, Some(Emu::from_twips(720)));
    assert_eq!(margins.right, Some(Emu::from_twips(720)));
    assert_eq!(
        sections.effective_page_size(2).unwrap().width,
        Some(Emu::from_twips(12240))
    );

    assert_eq!(
        sections.remove(0).unwrap().page_width(),
        Some(Emu::from_twips(12240))
    );
    sections.insert(0, first).unwrap();
    sections.push(Section::default());
    assert_eq!(sections.len(), 4);
    assert!(sections.remove(10).is_none());
}

#[test]
fn malformed_and_unbounded_fragments_are_rejected() {
    assert!(Section::from_xml_bytes(b"<w:notSectPr/>".to_vec()).is_err());
    assert!(Section::from_xml_bytes(b"<w:sectPr><w:pgSz></w:sectPr>".to_vec()).is_err());
    assert!(Section::from_xml_bytes(vec![b' '; MAX_XML_BYTES + 1]).is_err());
    assert!(Section::from_xml_bytes(
        br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="31681"/></w:sectPr>"#.to_vec()
    )
    .is_ok());
    assert!(Section::from_xml_bytes(
        br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/word/2006/main"><w:pgSz w:w="31681"/></w:sectPr>"#.to_vec(),
    )
    .is_err());
}
