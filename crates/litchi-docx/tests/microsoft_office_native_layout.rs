//! Native Word layout extensions must survive public source projections.
use litchi_docx::{Package, paragraph::Collapsed, section::footnote_columns::Layout};
fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
#[test]
fn native_collapsed_heading_and_footnote_columns_reach_public_views() {
    for name in ["word-layout-extensions.docx", "word-layout-extensions.dotx"] {
        let package = native(name);
        let document = package.document().unwrap();
        let paragraphs = document.paragraphs().unwrap();
        let heading = paragraphs
            .iter()
            .find(|p| p.text().unwrap() == "COLLAPSED_HEADING_016")
            .unwrap();
        assert_eq!(
            heading.collapsed().unwrap(),
            Some(Collapsed::Enabled),
            "{name}"
        );
        let mut sections = document.sections().unwrap();
        assert_eq!(sections.len(), 1);
        assert_eq!(
            sections
                .iter_mut()
                .next()
                .unwrap()
                .footnote_columns()
                .unwrap(),
            Some(Layout::new(2).unwrap()),
            "{name}"
        );
        let notes = document.footnotes().unwrap();
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].id(), 1);
        assert!(
            notes[0]
                .text()
                .unwrap()
                .contains("FOOTNOTE_COLUMN_016: native two-column footnote-area specimen.")
        );
    }
}

#[test]
fn native_section_extension_mutations_keep_page_and_note_settings() {
    for name in ["word-layout-extensions.docx", "word-layout-extensions.dotx"] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let document = package.document().unwrap();
        let mut sections = document.sections().unwrap();
        let section = sections.iter_mut().next().unwrap();
        let original = section.to_xml_bytes().unwrap();
        let page = section.page_size_checked().unwrap();
        let margins = section.margins_checked().unwrap();
        let columns = section.columns().unwrap();
        section
            .set_footnote_columns(Some(Layout::new(3).unwrap()))
            .unwrap();
        let changed = section.to_xml_bytes().unwrap();
        let mut reopened = litchi_docx::section::Section::from_xml_bytes(changed.clone()).unwrap();
        assert_eq!(
            reopened.footnote_columns().unwrap(),
            Some(Layout::new(3).unwrap())
        );
        assert_eq!(reopened.page_size_checked().unwrap(), page);
        assert_eq!(reopened.margins_checked().unwrap(), margins);
        assert_eq!(reopened.columns().unwrap(), columns);
        assert!(
            changed
                == String::from_utf8(original.clone())
                    .unwrap()
                    .replacen(
                        "<w15:footnoteColumns w:val=\"2\"",
                        "<w15:footnoteColumns w:val=\"3\"",
                        1
                    )
                    .as_bytes()
        );
        reopened.clear_footnote_columns().unwrap();
        assert_eq!(reopened.footnote_columns().unwrap(), None);
        reopened
            .set_footnote_columns(Some(Layout::new(0).unwrap()))
            .unwrap();
        assert!(
            reopened
                .footnote_columns()
                .unwrap()
                .unwrap()
                .follows_page_layout()
        );
        assert!(package.to_bytes().unwrap() == before);
    }
}

#[test]
fn native_borrowed_heading_can_change_collapse_state_without_losing_text_or_style() {
    for name in ["word-layout-extensions.docx", "word-layout-extensions.dotx"] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let document = package.document().unwrap();
        let mut paragraph = document
            .paragraphs()
            .unwrap()
            .into_iter()
            .find(|p| p.text().unwrap() == "COLLAPSED_HEADING_016")
            .unwrap();
        let properties = paragraph.direct_properties().unwrap();
        paragraph.set_collapsed(Some(Collapsed::Disabled)).unwrap();
        assert_eq!(paragraph.collapsed().unwrap(), Some(Collapsed::Disabled));
        assert_eq!(paragraph.text().unwrap(), "COLLAPSED_HEADING_016");
        assert_eq!(paragraph.direct_properties().unwrap(), properties);
        paragraph.set_collapsed(None).unwrap();
        assert_eq!(paragraph.collapsed().unwrap(), None);
        paragraph.set_collapsed(Some(Collapsed::Enabled)).unwrap();
        assert_eq!(paragraph.collapsed().unwrap(), Some(Collapsed::Enabled));
        assert!(package.to_bytes().unwrap() == before);
    }
}

#[test]
fn native_layout_source_text_edits_preserve_extensions_and_secondary_parts() {
    use litchi_core::Position;
    for name in ["word-layout-extensions.docx", "word-layout-extensions.dotx"] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let mut edit = package.edit().unwrap();
        edit.replace_run_text(Position::new(0), Position::new(0), "LAYOUT_SOURCE_REVIEWED")
            .unwrap();
        let commit = edit.commit().unwrap();
        let after = commit.package().to_bytes().unwrap();
        let output = tempfile::NamedTempFile::with_suffix(if name.ends_with("dotx") {
            ".dotx"
        } else {
            ".docx"
        })
        .unwrap();
        commit.package().save(output.path()).unwrap();
        let reopened = Package::open(output.path()).unwrap();
        let document = reopened.document().unwrap();
        assert_eq!(
            document.paragraphs().unwrap()[0].text().unwrap(),
            "LAYOUT_SOURCE_REVIEWED"
        );
        assert_eq!(
            document
                .paragraphs()
                .unwrap()
                .iter()
                .find(|p| p.text().unwrap() == "COLLAPSED_HEADING_016")
                .unwrap()
                .collapsed()
                .unwrap(),
            Some(Collapsed::Enabled)
        );
        assert_eq!(
            document
                .sections()
                .unwrap()
                .iter_mut()
                .next()
                .unwrap()
                .footnote_columns()
                .unwrap(),
            Some(Layout::new(2).unwrap())
        );
        let archive = |bytes: Vec<u8>| {
            litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
                bytes,
                litchi_opc::SourceLimits::default(),
            )
            .unwrap()
        };
        let a = archive(before.clone());
        let b = archive(after);
        assert_eq!(a.member_names().unwrap(), b.member_names().unwrap());
        for name in a.member_names().unwrap() {
            if name != "word/document.xml" {
                assert!(
                    a.read_member(&name).unwrap() == b.read_member(&name).unwrap(),
                    "changed {name}"
                );
            }
        }
        assert!(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == before
        );
        assert!(package.to_bytes().unwrap() == before);
    }
}

fn synthetic(xml: String) -> Package {
    let package = Package::new().unwrap();
    let mut opc = package.opc_package().clone();
    let main = opc.main_document_part().unwrap().partname().clone();
    opc.get_part_mut(&main).unwrap().set_blob(xml.into_bytes());
    Package::from_opc_package(opc).unwrap()
}
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const E: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[test]
fn selected_sections_keep_aliases_and_ignore_inactive_invalid_extensions() {
    let xml = format!(
        r#"<q:document xmlns:q="{W}" xmlns:e="{E}" xmlns:c="{MC}" xmlns:future="urn:future" xmlns:mc="urn:foreign" c:Ignorable="e"><q:body><q:p><q:pPr><e:collapsed/></q:pPr><q:r><q:t>Heading</q:t></q:r></q:p><c:AlternateContent><c:Choice Requires="future"><q:sectPr><e:footnoteColumns q:val="-1"/></q:sectPr></c:Choice><c:Fallback><q:sectPr><q:pgSz q:w="12240" q:h="15840"/><e:footnoteColumns q:val="02"/></q:sectPr></c:Fallback></c:AlternateContent></q:body></q:document>"#
    );
    for xml in [
        xml.clone(),
        xml.replace(W, "http://purl.oclc.org/ooxml/wordprocessingml/main"),
    ] {
        let package = synthetic(xml);
        let document = package.document().unwrap();
        assert_eq!(
            document.paragraphs().unwrap()[0].collapsed().unwrap(),
            Some(Collapsed::Enabled)
        );
        let mut sections = document.sections().unwrap();
        assert_eq!(sections.len(), 1);
        let section = sections.iter_mut().next().unwrap();
        assert_eq!(
            section.footnote_columns().unwrap(),
            Some(Layout::new(2).unwrap())
        );
        let unchanged = section.to_xml_bytes().unwrap();
        section
            .set_footnote_columns(Some(Layout::new(2).unwrap()))
            .unwrap();
        assert!(section.to_xml_bytes().unwrap() == unchanged);
        section
            .set_footnote_columns(Some(Layout::new(5).unwrap()))
            .unwrap();
        assert_eq!(
            section.footnote_columns().unwrap(),
            Some(Layout::new(5).unwrap())
        );
    }
}

#[test]
fn section_projection_does_not_invent_missing_ignorable_or_accept_invalid_values() {
    for suffix in [
        r#"<e:footnoteColumns q:val="-1"/>"#,
        r#"<e:footnoteColumns q:val="2"/><e:footnoteColumns q:val="3"/>"#,
    ] {
        let xml = format!(
            r#"<q:document xmlns:q="{W}" xmlns:e="{E}" xmlns:c="{MC}" c:Ignorable="e"><q:body><q:sectPr>{suffix}</q:sectPr></q:body></q:document>"#
        );
        assert!(synthetic(xml).document().unwrap().sections().is_err());
    }
    let xml = format!(
        r#"<q:document xmlns:q="{W}" xmlns:e="{E}"><q:body><q:sectPr><e:footnoteColumns q:val="2"/></q:sectPr></q:body></q:document>"#
    );
    assert!(synthetic(xml).document().unwrap().sections().is_err());
    let foreign = format!(
        r#"<q:document xmlns:q="{W}" xmlns:e="urn:foreign" xmlns:c="{MC}" c:Ignorable="e"><q:body><q:p><q:pPr><e:collapsed/></q:pPr><q:r><q:t>Foreign</q:t></q:r></q:p><q:sectPr><e:footnoteColumns q:val="2"/></q:sectPr></q:body></q:document>"#
    );
    let package = synthetic(foreign);
    let document = package.document().unwrap();
    assert_eq!(document.paragraphs().unwrap()[0].collapsed().unwrap(), None);
    assert_eq!(
        document
            .sections()
            .unwrap()
            .iter_mut()
            .next()
            .unwrap()
            .footnote_columns()
            .unwrap(),
        None
    );
}

#[test]
fn selected_section_projection_applies_compatibility_inside_the_section_owner() {
    let xml = format!(
        r#"<q:document xmlns:q="{W}" xmlns:e="{E}" xmlns:c="{MC}" xmlns:future="urn:future" c:Ignorable="e"><q:body><q:sectPr><c:AlternateContent><c:Choice Requires="future"><q:pgSz q:w="1" q:h="1"/><e:footnoteColumns q:val="-1"/></c:Choice><c:Fallback><q:pgSz q:w="12240" q:h="15840"/><e:footnoteColumns q:val="3"/></c:Fallback></c:AlternateContent></q:sectPr></q:body></q:document>"#
    );
    let package = synthetic(xml);
    let mut sections = package.document().unwrap().sections().unwrap();
    let section = sections.iter_mut().next().unwrap();
    assert_eq!(
        section
            .page_size_checked()
            .unwrap()
            .width
            .unwrap()
            .to_twips(),
        12240
    );
    assert_eq!(
        section.footnote_columns().unwrap(),
        Some(Layout::new(3).unwrap())
    );
}

#[test]
fn native_caption_fields_keep_sequence_identifiers_and_cached_results() {
    for name in ["word-layout-extensions.docx", "word-layout-extensions.dotx"] {
        let package = native(name);
        let document = package.document().unwrap();
        let fields = document.fields().unwrap();
        assert_eq!(fields.len(), 2);
        for (field, identifier) in fields.iter().zip(["Figure", "Table"]) {
            assert_eq!(field.instruction(), format!(r"SEQ {identifier} \* ARABIC"));
            assert_eq!(field.result(), Some("1"));
            let sequence = field.sequence_field().unwrap().unwrap();
            assert_eq!(sequence.identifier(), identifier);
            assert_eq!(sequence.bookmark(), None);
            assert_eq!(sequence.cached_result(), Some("1"));
            assert!(!sequence.is_dirty());
            assert!(!sequence.is_locked());
        }
        let table = document.table(0).unwrap().unwrap();
        assert_eq!(
            (table.row_count().unwrap(), table.column_count().unwrap()),
            (2, 2)
        );
        for (row, expected) in [["AUTO_TABLE_016", "Amount"], ["Native AutoCaption", "42"]]
            .into_iter()
            .enumerate()
        {
            for (column, text) in expected.into_iter().enumerate() {
                assert_eq!(
                    table.cell(row, column).unwrap().unwrap().text().unwrap(),
                    text
                );
            }
        }
    }
}
