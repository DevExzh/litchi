#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use litchi_opc::PackURI;
use litchi_xlsx::{Hyperlink, HyperlinkReference, Package, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn workbook() -> Workbook {
    Package::create().unwrap().into_workbook().unwrap()
}

fn internal(reference: &str, location: &str) -> Hyperlink {
    Hyperlink::internal(
        HyperlinkReference::new(reference).unwrap(),
        location.to_owned(),
    )
    .unwrap()
}

fn external(reference: &str, target: &str) -> Hyperlink {
    Hyperlink::external(
        HyperlinkReference::new(reference).unwrap(),
        target.to_owned(),
    )
    .unwrap()
}

fn links(workbook: &Workbook) -> Vec<Hyperlink> {
    workbook
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .hyperlinks()
        .unwrap()
}

#[test]
fn hyperlinks_add_replace_remove_and_reopen() {
    let base = workbook();
    let mut edit = base.edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet.put_hyperlink(internal("A1", "Sheet1!B2")).unwrap();
        sheet
            .put_hyperlink(external("B2", "https://example.test/inert"))
            .unwrap();
    }
    let after = edit.commit().unwrap().into_workbook();

    let initial_links = links(&after);
    assert_eq!(initial_links.len(), 2);
    let internal_link = initial_links
        .iter()
        .find(|link| link.reference().as_str() == "A1");
    assert_eq!(internal_link.unwrap().location(), Some("Sheet1!B2"));
    let external_link = initial_links
        .iter()
        .find(|link| link.reference().as_str() == "B2");
    assert_eq!(
        external_link.unwrap().external_target(),
        Some("https://example.test/inert")
    );

    let mut edit = after.edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        let removed = sheet
            .remove_hyperlink(HyperlinkReference::new("B2").unwrap())
            .unwrap()
            .unwrap();
        assert_eq!(
            removed.external_target(),
            Some("https://example.test/inert")
        );
        sheet
            .replace_hyperlink(external("A1", "https://example.test/replaced"))
            .unwrap();
    }
    let after_replace = edit.commit().unwrap().into_workbook();
    let links = links(&after_replace);
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].reference().as_str(), "A1");
    assert_eq!(links[0].location(), None);
    assert_eq!(
        links[0].external_target(),
        Some("https://example.test/replaced")
    );
}

#[test]
fn hyperlink_patch_inverse_restores_source_snapshot() {
    let base = workbook();
    let mut edit = base.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .put_hyperlink(external("C3:D4", "https://example.test/undo"))
        .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();

    let restored = after.apply(&patch.inverse()).unwrap().into_workbook();
    assert!(links(&restored).is_empty());
    assert_eq!(
        restored.to_plain_bytes().unwrap(),
        base.to_plain_bytes().unwrap()
    );
}

#[test]
fn hyperlink_add_then_remove_is_an_exact_noop() {
    let base = workbook();
    let mut edit = base.edit().unwrap();
    let href = internal("$A$1", "Sheet1!B2");
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet.put_hyperlink(href.clone()).unwrap();
        let removed = sheet.remove_hyperlink(href.reference().clone()).unwrap();
        assert_eq!(removed, Some(href));
    }
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_empty());
    let unchanged = commit.into_workbook();
    assert_eq!(
        unchanged.to_plain_bytes().unwrap(),
        base.to_plain_bytes().unwrap()
    );
}

#[test]
fn hyperlink_joins_keep_disjoint_ranges_and_reject_same_semantic_range() {
    let base = workbook();
    let mut left = base.edit().unwrap();
    left.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .put_hyperlink(internal("A1", "Sheet1!B2"))
        .unwrap();
    let mut right = base.edit().unwrap();
    right
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .put_hyperlink(external("C3", "https://example.test/right"))
        .unwrap();
    left.join(right).unwrap();
    let joined = left.commit().unwrap().into_workbook();
    assert_eq!(links(&joined).len(), 2);

    let mut same_left = base.edit().unwrap();
    same_left
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .put_hyperlink(internal("A1", "Sheet1!B2"))
        .unwrap();
    let mut same_right = base.edit().unwrap();
    same_right
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .put_hyperlink(external("$A$1", "https://example.test/conflict"))
        .unwrap();
    assert!(same_left.join(same_right).is_err());
}

#[test]
fn hyperlink_edit_refuses_unknown_owner_markup() {
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let mut raw = Package::create().unwrap().into_plain_opc();
    raw.get_part_mut(&worksheet_uri).unwrap().set_blob(
        format!(
            r#"<worksheet xmlns="{MAIN}"><sheetData/><hyperlinks><hyperlink ref="A1" location="Sheet1!B2"><vendor/></hyperlink></hyperlinks></worksheet>"#
        )
        .into_bytes(),
    );
    let base = Package::from_opc(raw).unwrap().into_workbook().unwrap();
    let mut edit = base.edit().unwrap();
    let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
    assert!(sheet.put_hyperlink(internal("A1", "Sheet1!C3")).is_err());
}

#[test]
fn hyperlink_values_reject_non_stable_xml_whitespace() {
    for location in ["Sheet1!A\t1", "Sheet1!A\n1", "Sheet1!A\r1"] {
        assert!(Hyperlink::internal(HyperlinkReference::new("A1").unwrap(), location).is_err());
    }
}
