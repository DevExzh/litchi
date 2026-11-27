//! Document page defaults are distinct from per-section overrides.
use litchi_rtf::{Package, PageOrientation, SectionDocumentGridType, SectionProperties};
fn properties(package: &Package) -> Vec<SectionProperties> {
    package
        .sections()
        .iter()
        .map(|s| s.properties.clone())
        .collect()
}
#[test]
fn sectd_restores_document_geometry_after_section_overrides() {
    let package=Package::parse(r"{\rtf1\paperw15000\paperh20000\margl1000\margr1100\margt1200\margb1300\landscape\sectd A\sect\sectd\pgwsxn17000\marglsxn1900 B\sect\sectd C}").unwrap();
    let sections = properties(&package);
    assert_eq!(sections.len(), 3);
    assert_eq!(
        (sections[0].page_width, sections[0].page_height),
        (15000, 20000)
    );
    assert_eq!(
        (
            sections[0].margin_left,
            sections[0].margin_right,
            sections[0].margin_top,
            sections[0].margin_bottom
        ),
        (1000, 1100, 1200, 1300)
    );
    assert_eq!(
        (sections[1].page_width, sections[1].margin_left),
        (17000, 1900)
    );
    assert_eq!(sections[2], sections[0]);
    assert!(
        sections
            .iter()
            .all(|s| s.orientation == PageOrientation::Landscape)
    );
    let mut output = Vec::new();
    litchi_rtf::write::Writer::new(&mut output)
        .write_document(package.raw())
        .unwrap();
    assert_eq!(properties(&Package::from_bytes(&output).unwrap()), sections);
}
#[test]
fn document_updates_leave_explicit_section_overrides_intact() {
    let package=Package::parse(r"{\rtf1\paperw12000\margl1000\sectd A\sect\sectd\pgwsxn14000\marglsxn2000 B\paperw16000\margl900\sect\sectd C}").unwrap();
    assert_eq!(
        package
            .sections()
            .iter()
            .map(|s| (s.properties.page_width, s.properties.margin_left))
            .collect::<Vec<_>>(),
        [(16000, 900), (14000, 2000), (16000, 900)]
    );
}
#[test]
fn default_grid_and_snap_grid_do_not_reset_columns_or_line_pitch() {
    for (word, expected) in [
        ("sectdefaultcl", SectionDocumentGridType::Default),
        ("sectspecifygenN", SectionDocumentGridType::SnapToGrid),
        ("sectspecifygen", SectionDocumentGridType::SnapToGrid),
    ] {
        let package = Package::parse(&format!(
            "{{\\rtf1\\sectd\\cols2\\colsx360\\sectlinegrid480\\sectspecifyl\\{word} X}}"
        ))
        .unwrap();
        let p = &package.sections()[0].properties;
        assert_eq!((p.columns.count, p.columns.default_spacing), (2, 360));
        assert_eq!(p.document_grid.line_grid, Some(480));
        assert_eq!(p.document_grid.grid_type, Some(expected));
        let mut output = Vec::new();
        litchi_rtf::write::Writer::new(&mut output)
            .write_document(package.raw())
            .unwrap();
        assert_eq!(
            properties(&Package::from_bytes(&output).unwrap()),
            properties(&package)
        );
    }
}
#[test]
fn grouped_section_boundaries_survive_writing_and_ignore_opaque_destinations() {
    let package = Package::parse(r"{\rtf1\sectd A{\b\sect}\sectd B}").unwrap();
    let breaks = package.raw().section_breaks().copied().collect::<Vec<_>>();
    assert_eq!(breaks, [litchi_rtf::SectionBreak::new(1, Some(1))]);
    let mut output = Vec::new();
    litchi_rtf::write::Writer::new(&mut output)
        .write_document(package.raw())
        .unwrap();
    assert_eq!(
        Package::from_bytes(&output)
            .unwrap()
            .raw()
            .section_breaks()
            .copied()
            .collect::<Vec<_>>(),
        breaks
    );
    let opaque = Package::parse(r"{\rtf1\sectd A{\*\future\sect}{\b\sect}\sectd B}").unwrap();
    assert_eq!(
        opaque.raw().section_breaks().copied().collect::<Vec<_>>(),
        breaks
    );
}
#[test]
fn section_edits_follow_definition_identity_across_inherited_grouped_breaks() {
    let package =
        Package::parse(r"{\rtf1\sectd\pgwsxn12000 A{\b\sect}B{\i\sect}\sectd\pgwsxn16000 C}")
            .unwrap();
    assert_eq!(package.sections().len(), 2);
    assert_eq!(
        package
            .raw()
            .section_breaks()
            .map(|b| b.next_section)
            .collect::<Vec<_>>(),
        [None, Some(1)]
    );
    let mut edit = package.edit();
    edit.set_section_values(1, &[litchi_rtf::edit::SectionValue::MarginLeft(1000)])
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package();
    assert_eq!(
        after.sections()[0].properties,
        package.sections()[0].properties
    );
    assert_eq!(
        (
            after.sections()[1].properties.page_width,
            after.sections()[1].properties.margin_left
        ),
        (16000, 1000)
    );
    assert_eq!(after.text(), package.text());
    assert!(
        after
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == package.to_bytes().unwrap()
    );
}
#[test]
fn header_section_controls_do_not_change_document_page_defaults() {
    let package=Package::parse(r"{\rtf1\paperw15000\margl1000\sectd{\header H\sectd\paperw100\pgwsxn200\margl300\marglsxn400}X\sect\sectd Y}").unwrap();
    assert_eq!(package.sections().len(), 2);
    for s in package.sections() {
        assert_eq!(
            (s.properties.page_width, s.properties.margin_left),
            (15000, 1000)
        );
    }
    assert_eq!(package.sections()[0].headers_footers[0].text(), "H");
}
#[test]
fn note_edits_allow_scoped_script_selection_and_preserve_the_reference_group() {
    let bytes = br"{\rtf1 A{\footnote{\super\chftn}{\rtlch\fcs1 x}{\ltrch\fcs0 y}}B}";
    let package = Package::from_bytes(bytes).unwrap();
    let mut edit = package.edit();
    edit.set_note_text(0, "changed").unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.package().to_bytes().unwrap();
    assert!(String::from_utf8_lossy(&output).contains(r"{\super\chftn}"));
    assert_eq!(commit.package().notes()[0].content, "changed");
    assert_eq!(commit.package().notes()[0].reference, "1");
    assert_eq!(commit.package().text(), package.text());
    assert!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    let scoped = Package::from_bytes(
        br"{\rtf1 A{\footnote x\fcs1 y}B\par\trowd\cellx1000\pard\intbl T\cell\row}",
    )
    .unwrap();
    let mut edit = scoped.edit();
    edit.set_note_text(0, "changed").unwrap();
    assert_eq!(scoped.tables().len(), 1);
    assert!(
        matches!(edit.commit(),Err(litchi_rtf::edit::Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent && r.reason()=="destination text contains a structural or state-changing control")
    );
}
