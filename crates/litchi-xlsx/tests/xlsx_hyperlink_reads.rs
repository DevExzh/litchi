use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{PackURI, TargetMode};
use litchi_xlsx::Package;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[test]
fn worksheet_facade_exposes_internal_and_inert_external_hyperlinks() {
    let worksheet_uri =
        PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet part URI should be valid");
    let worksheet = format!(
        r#"<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData/><hyperlinks><hyperlink ref="A1" location="Sheet1!B2" display="local"/><hyperlink ref="C3:D4" r:id="rIdHyperlink1" tooltip="inert"/></hyperlinks></worksheet>"#
    )
    .into_bytes();
    let mut raw = Package::create()
        .expect("minimal XLSX package should be created")
        .into_plain_opc();
    let part = raw
        .get_part_mut(&worksheet_uri)
        .expect("worksheet part should exist");
    part.set_blob(worksheet);
    part.rels_mut()
        .try_add_relationship(
            rt::HYPERLINK.to_string(),
            "https://127.0.0.1:9/never-open?q=1#fragment".to_string(),
            "rIdHyperlink1".to_string(),
            TargetMode::External,
        )
        .expect("external hyperlink relationship should be added");

    let package = Package::from_opc(raw).expect("hyperlink package should validate");
    let workbook = package.workbook().expect("workbook should open");
    let sheet = workbook
        .active_sheet()
        .expect("active worksheet should exist");
    let hyperlinks = sheet.hyperlinks().expect("typed hyperlinks should parse");

    assert_eq!(hyperlinks.len(), 2);
    assert_eq!(hyperlinks[0].reference().as_str(), "A1");
    assert_eq!(hyperlinks[0].location(), Some("Sheet1!B2"));
    assert_eq!(hyperlinks[0].display(), Some("local"));
    assert_eq!(hyperlinks[1].reference().range().a1(), "C3:D4");
    assert_eq!(
        hyperlinks[1].external_target(),
        Some("https://127.0.0.1:9/never-open?q=1#fragment")
    );
    assert_eq!(hyperlinks[1].tooltip(), Some("inert"));
}
