use litchi_ooxml::xlsx::{
    ExternalDdeLink, ExternalLinkConformance, ExternalLinkKind, ExternalWorkbookLink,
    ExternalWorkbookTarget, Workbook,
};
use litchi_opc::constants::relationship_type as rt;

fn dde(topic: &str) -> ExternalLinkKind {
    ExternalLinkKind::Dde(ExternalDdeLink {
        service: "excel".into(),
        topic: topic.into(),
        items: Vec::new(),
    })
}

fn book(target: &str) -> ExternalLinkKind {
    ExternalLinkKind::Workbook(ExternalWorkbookLink {
        target: ExternalWorkbookTarget {
            relationship_id: "rIdPath1".into(),
            target: target.into(),
            relationship_type: rt::EXTERNAL_LINK_PATH.into(),
        },
        sheet_names: vec!["Data".into()],
        defined_names: Vec::new(),
        cached_sheets: Vec::new(),
    })
}

#[test]
fn add_find_replace_reorder_remove_survive_save_without_fetching() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("links.xlsx");
    let mut workbook = Workbook::create().unwrap();
    assert_eq!(
        workbook
            .add_external_link(book("https://127.0.0.1:9/never.xlsx"))
            .unwrap(),
        1
    );
    assert_eq!(workbook.add_external_link(dde("second")).unwrap(), 2);
    assert_eq!(
        workbook
            .find_external_links_by_target("https://127.0.0.1:9/never.xlsx")
            .len(),
        1
    );
    workbook.update_external_link(1, dde("first")).unwrap();
    workbook.reorder_external_links(&[2, 1]).unwrap();
    assert!(
        matches!(&workbook.external_link(1).unwrap().kind, ExternalLinkKind::Dde(link) if link.topic == "second")
    );
    workbook.remove_external_link(2).unwrap();
    workbook.save(&path).unwrap();
    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(reopened.external_links().len(), 1);
    assert!(
        matches!(&reopened.external_link(1).unwrap().kind, ExternalLinkKind::Dde(link) if link.topic == "second")
    );
}

#[test]
fn malformed_target_rolls_back_and_strict_part_is_supported() {
    let mut workbook = Workbook::create().unwrap();
    let invalid = book("https://example.test/bad\u{0}target.xlsx");
    assert!(workbook.add_external_link(invalid).is_err());
    assert!(workbook.external_links().is_empty());
    assert_eq!(
        workbook
            .add_external_link_with_conformance(dde("strict"), ExternalLinkConformance::Strict,)
            .unwrap(),
        1
    );
}

#[test]
fn reorder_validation_is_atomic() {
    let mut workbook = Workbook::create().unwrap();
    workbook.add_external_link(dde("a")).unwrap();
    workbook.add_external_link(dde("b")).unwrap();
    assert!(workbook.reorder_external_links(&[1, 1]).is_err());
    assert!(
        matches!(&workbook.external_link(2).unwrap().kind, ExternalLinkKind::Dde(link) if link.topic == "b")
    );
}
