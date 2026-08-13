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

#[test]
fn standalone_sections_validate_qnames_without_rejecting_inherited_prefixes() {
    for xml in [
        br#"<w:sectPr xmlns:1bad="urn:invalid"/>"#.to_vec(),
        br#"<w:sectPr><1bad:x/></w:sectPr>"#.to_vec(),
        br#"<w:sectPr foo:bad:name="x"/>"#.to_vec(),
        br#"<w:sectPr><xmlns:p/></w:sectPr>"#.to_vec(),
        br#"<w:sectPr xmlns:foo="http://www.w3.org/XML/1998/namespace"/>"#.to_vec(),
        br#"<w:sectPr xmlns:xml="urn:wrong"/>"#.to_vec(),
        br#"<w:sectPr xmlns:foo="http://www.w3.org/2000/xmlns/"/>"#.to_vec(),
        br#"<w:sectPr xmlns="http://www.w3.org/XML/1998/namespace"/>"#.to_vec(),
    ] {
        assert!(
            Section::from_xml_bytes(xml).is_err(),
            "accepted invalid standalone section QName or binding"
        );
    }

    assert!(
        Section::from_xml_bytes(
            br#"<w:sectPr><x:opaque x:value="inherited"/></w:sectPr>"#.to_vec()
        )
        .is_ok()
    );
}

#[test]
fn inventory_reports_zero_one_and_multiple_logical_sections() {
    let empty = Inventory::parse(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#,
    )
    .unwrap();
    assert_eq!(empty.sections().len(), 1);
    assert_eq!(empty.sections()[0].ownership(), Ownership::Implicit);
    assert!(empty.sections()[0].paragraphs().is_empty());

    let one = Inventory::parse(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
    )
    .unwrap();
    assert_eq!(one.sections().len(), 1);
    assert_eq!(one.sections()[0].ownership(), Ownership::BodyFinal);
    assert_eq!(one.sections()[0].paragraphs().len(), 1);
    assert_eq!(
        one.property(0, Property::PageSize),
        Some(PropertyValue::PageSize(Some(PageSize {
            width: Some(Emu::from_twips(12240)),
            height: Some(Emu::from_twips(15840)),
            orientation: Orientation::Portrait,
        })))
    );

    let multiple = Inventory::parse(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/><w:sectPr/></w:body></w:document>"#,
    )
    .unwrap();
    assert_eq!(multiple.sections().len(), 2);
    assert_eq!(
        multiple.sections()[0].ownership(),
        Ownership::Paragraph(litchi_core::Position::new(0))
    );
    assert_eq!(multiple.sections()[0].paragraphs().len(), 1);
    assert_eq!(multiple.sections()[1].paragraphs().len(), 2);
    assert_eq!(
        multiple
            .section(Selector::paragraph(litchi_core::Position::new(0)))
            .unwrap()
            .start(),
        Some(Start::Continuous)
    );
    assert!(multiple.section(9).is_none());
    assert!(multiple.section(Ownership::Implicit).is_none());
}

#[test]
fn inventory_selects_mce_and_keeps_header_footer_references_inert() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><w:body><w:p/><mc:AlternateContent><mc:Choice Requires="x"><w:sectPr><w:headerReference w:type="default" r:id="active-unsafe"/></w:sectPr></mc:Choice><mc:Fallback><w:sectPr><w:headerReference w:type="first" r:id="rHeader"/><w:footerReference w:type="even" r:id="rFooter"/><x:opaque/></w:sectPr></mc:Fallback></mc:AlternateContent></w:body></w:document>"#;
    let inventory = Inventory::parse(xml).unwrap();
    assert_eq!(inventory.sections().len(), 1);
    assert_eq!(inventory.sections()[0].ownership(), Ownership::BodyFinal);
    assert_eq!(
        inventory.sections()[0].headers()[0].relationship_id,
        "rHeader"
    );
    assert_eq!(
        inventory.sections()[0].footers()[0].relationship_id,
        "rFooter"
    );
    let clone = inventory.clone();
    assert!(inventory.shares_allocation_with(&clone));
}

#[test]
fn inventory_ignores_table_cell_section_properties_for_main_story_boundaries() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:sectPr/></w:pPr></w:p><w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r/></w:p></w:tc></w:tr></w:tbl><w:p/><w:sectPr/></w:body></w:document>"#;
    let inventory = Inventory::parse(xml).unwrap();

    assert_eq!(inventory.paragraph_count(), 3);
    assert_eq!(inventory.sections().len(), 2);
    assert_eq!(
        inventory.sections()[0].ownership(),
        Ownership::Paragraph(litchi_core::Position::new(0))
    );
    assert_eq!(inventory.sections()[0].paragraphs().len(), 1);
    assert_eq!(inventory.sections()[1].ownership(), Ownership::BodyFinal);
    assert_eq!(inventory.sections()[1].paragraphs().len(), 2);
}

#[test]
fn inventory_materializes_inherited_namespaces_after_quoted_gt_and_refuses_malformed_opening() {
    let quoted = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:opaque"><w:body><w:sectPr x:opaque="a>b" x:single='c>d'><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>"#;
    let inventory = Inventory::parse(quoted).unwrap();
    assert_eq!(inventory.sections()[0].headers()[0].relationship_id, "rId1");

    for malformed in [
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><w:body><w:sectPr x:opaque="a>b></w:sectPr></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><w:body><w:sectPr x:opaque='a>b></w:sectPr></w:body></w:document>"#.as_slice(),
    ] {
        assert!(Inventory::parse(malformed).is_err());
    }
}

#[test]
fn inventory_rejects_unbound_prefixes_and_non_xml_relationship_ids() {
    let unbound_relationship = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>"#;
    assert!(Inventory::parse(unbound_relationship).is_err());

    let unbound_element = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><x:opaque/></w:body></w:document>"#;
    assert!(Inventory::parse(unbound_element).is_err());

    for relationship_id in ["1rId", "r Id", "r:id", "r/id"] {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:headerReference w:type="default" r:id="{relationship_id}"/></w:sectPr></w:body></w:document>"#
        );
        assert!(
            Inventory::parse(xml.as_bytes()).is_err(),
            "accepted invalid relationship ID {relationship_id:?}"
        );
    }

    let valid = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:headerReference w:type="default" r:id="_header-1.2"/></w:sectPr></w:body></w:document>"#;
    assert!(Inventory::parse(valid).is_ok());
}

#[test]
fn inventory_rejects_invalid_qnames_and_reserved_namespace_bindings() {
    for xml in [
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:1bad="urn:invalid"><w:body/></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:foo="urn:invalid"><w:body foo:bad:name="x"/></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:1bad="urn:invalid"><w:body><1bad:p/></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xmlns="urn:invalid"><w:body><xmlns:p/></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:foo="http://www.w3.org/XML/1998/namespace"><w:body/></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="urn:wrong"><w:body/></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:foo="http://www.w3.org/2000/xmlns/"><w:body/></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns="http://www.w3.org/XML/1998/namespace"><w:body/></w:document>"#.as_slice(),
    ] {
        assert!(Inventory::parse(xml).is_err());
    }
}

#[test]
fn inventory_refuses_unsupported_mce_branch_and_body_final_trailing_content() {
    let malformed_mce = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><w:body><mc:AlternateContent><mc:Fallback/><mc:Choice Requires="x"/></mc:AlternateContent></w:body></w:document>"#;
    assert!(Inventory::parse(malformed_mce).is_err());

    let trailing = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><w:body><w:p/><w:sectPr/><x:tail/></w:body></w:document>"#;
    assert!(Inventory::parse(trailing).is_err());
}

#[test]
fn inventory_accepts_strict_and_refuses_bad_topology_and_section_grammar() {
    let strict = Inventory::parse(
        br#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:body><s:p><s:pPr><s:sectPr/></s:pPr></s:p><s:sectPr><s:type s:val="oddPage"/></s:sectPr></s:body></s:document>"#,
    )
    .unwrap();
    assert_eq!(strict.sections().len(), 2);
    assert_eq!(strict.sections()[1].start(), Some(Start::OddPage));

    for malformed in [
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr/><w:p/></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr/><w:sectPr/></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:sectPr/></w:p></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:sectPr/><w:spacing/></w:pPr></w:p></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgMar/><w:pgSz/></w:sectPr></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz/><w:pgSz/></w:sectPr></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:type w:val="hostile"/></w:sectPr></w:body></w:document>"#.as_slice(),
    ] {
        assert!(Inventory::parse(malformed).is_err());
    }
}

#[test]
fn inventory_enforces_caller_limits_and_snapshot_selection() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr/></w:body></w:document>"#;
    let inventory = Inventory::parse(xml).unwrap();
    let mut limits = Limits {
        max_sections: 1,
        max_paragraphs: 1,
        max_input_bytes: xml.len(),
        ..Limits::default()
    };
    assert!(Inventory::parse_with_limits(xml, &limits).is_ok());

    limits.max_paragraphs = 0;
    assert!(Inventory::parse_with_limits(xml, &limits).is_err());
    limits.max_paragraphs = 1;
    limits.max_input_bytes -= 1;
    assert!(matches!(
        Inventory::parse_with_limits(xml, &limits),
        Err(crate::Error::SectionInventoryLimit { .. })
    ));

    assert!(inventory.property(8, Property::Margins).is_none());
    let snapshot = Snapshot::from_xml(xml.to_vec()).unwrap();
    assert!(snapshot.source_version().is_none());
    assert!(snapshot.section(Selector::body_final()).is_some());
    let clone = snapshot.clone();
    assert!(snapshot.shares_allocation_with(&clone));
}

#[test]
fn inventory_rejects_spoofed_attributes_and_malformed_document_structure() {
    let spoofed = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:foreign"><w:body><w:sectPr><w:headerReference x:type="default" x:id="spoof"/><w:pgSz x:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#;
    assert!(Inventory::parse(spoofed).is_err());

    let foreign_width = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:body><w:sectPr><w:pgSz x:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#;
    let inventory = Inventory::parse(foreign_width).unwrap();
    assert_eq!(inventory.sections()[0].page_size().unwrap().width, None);

    for malformed in [
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p></w:body></w:document>"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>trailing"#.as_slice(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></x:document>"#.as_slice(),
    ] {
        assert!(Inventory::parse(malformed).is_err());
    }
}

#[test]
fn inventory_preflights_reference_limits_before_semantic_retention() {
    // The invalid type after the large relationship would win during semantic
    // decoding. Receiving the limit error proves the original-resolver
    // preflight rejects the ID before constructing the Section/Reference.
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:sectPr><w:headerReference w:type="default" r:id="relationship-id"/><w:type w:val="invalid"/></w:sectPr></w:body></w:document>"#;
    let limits = Limits {
        max_reference_bytes: "relationship-id".len() - 1,
        ..Limits::default()
    };
    assert!(matches!(
        Inventory::parse_with_limits(xml, &limits),
        Err(crate::Error::SectionInventoryLimit {
            resource: "header/footer reference bytes",
            ..
        })
    ));
}

#[test]
fn inventory_bounds_inherited_namespace_materialization() {
    let section = br#"<w:sectPr><w:headerReference w:type="default" rel:id="header"/></w:sectPr>"#;
    let mut xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>"#.to_vec();
    xml.extend_from_slice(section);
    xml.extend_from_slice(b"</w:body></w:document>");
    let limits = Limits {
        max_section_bytes: section.len(),
        ..Limits::default()
    };
    assert!(matches!(
        Inventory::parse_with_limits(&xml, &limits),
        Err(crate::Error::SectionInventoryLimit {
            resource: "section bytes",
            ..
        })
    ));
}

#[test]
fn inventory_uses_original_namespace_aliases_and_ignores_nested_story_paragraphs() {
    let xml = br#"<word:document xmlns:word="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><word:body><word:p><word:r><word:drawing><word:txbxContent><word:p><word:pPr><word:sectPr/></word:pPr></word:p></word:txbxContent></word:drawing></word:r></word:p><word:p><word:pPr><word:sectPr><word:headerReference word:type="default" rel:id="aliased-header"/></word:sectPr></word:pPr></word:p><word:sectPr/></word:body></word:document>"#;
    let inventory = Inventory::parse(xml).unwrap();
    assert_eq!(inventory.paragraph_count(), 2);
    assert_eq!(inventory.sections().len(), 2);
    assert_eq!(
        inventory.sections()[0].ownership(),
        Ownership::Paragraph(litchi_core::Position::new(1))
    );
    assert_eq!(inventory.sections()[0].paragraphs().len(), 2);
    assert_eq!(
        inventory.sections()[0].headers()[0].relationship_id,
        "aliased-header"
    );
    assert!(inventory.sections()[1].paragraphs().is_empty());
}
