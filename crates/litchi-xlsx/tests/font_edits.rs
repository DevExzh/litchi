//! Native XLSX font authoring and reversible patch regressions.
//!
//! These tests deliberately inspect the package's styles part as well as the
//! semantic format view.  A font edit changes a workbook-global resource, so
//! checking only the selected cell can miss shared-font fan-out, duplicate
//! resource allocation, stale guards, or loss of opaque stylesheet markup.
#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use litchi_opc::PackURI;
use litchi_xlsx::{Error, FontChange, Package};

const STYLES_URI: &str = "/xl/styles.xml";
const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14AC_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

fn fixture() -> Package {
    let workbook = Package::new().expect("create workbook");
    let mut edit = workbook.edit().expect("begin fixture edit");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("find fixture sheet")
        .expect("fixture worksheet");
    sheet
        .set("A1", "shared font target")
        .and_then(|sheet| sheet.set("B1", "shared font neighbour"))
        .and_then(|sheet| sheet.set("C1", "shared font reuse"))
        .expect("seed fixture cells");
    edit.commit().expect("commit fixture").into_package()
}

fn styles_xml(workbook: &Package) -> Vec<u8> {
    let package = Package::from_bytes(&workbook.to_plain_bytes().expect("serialize workbook"))
        .expect("reopen workbook package")
        .into_plain_opc();
    package
        .get_part(&PackURI::new(STYLES_URI).expect("styles URI"))
        .expect("styles part")
        .blob()
        .to_vec()
}

fn typical_excel_styles_fixture() -> Package {
    let package = Package::new().expect("create package").into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            &format!(r#"<styleSheet xmlns="{MAIN_NS}">"#),
            &format!(
                r#"<styleSheet xmlns="{MAIN_NS}" xmlns:mc="{MCE_NS}" xmlns:x14ac="{X14AC_NS}" mc:Ignorable="x14ac">"#
            ),
        )
        .replace(
            r#"<fonts count="1">"#,
            r#"<fonts count="1" x14ac:knownFonts="1">"#,
        )
        .replace(
            r#"<font/>"#,
            r#"<font x14ac:fontMarker="keep"><x14ac:fontExtension marker="keep"/></font>"#,
        )
        .replace(
            r#"<cellXfs count="1"><xf "#,
            r#"<cellXfs count="1"><xf x14ac:xfMarker="keep" "#,
        );
    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    let workbook = Package::from_opc(package).expect("adopt typical Excel styles");
    let mut edit = workbook.edit().expect("begin fixture edit");
    edit.sheet("Sheet1")
        .expect("find fixture sheet")
        .expect("fixture worksheet")
        .set("A1", "typical Excel styles")
        .expect("seed fixture cell");
    edit.commit().expect("commit fixture").into_package()
}

fn direct_tag_count(xml: &[u8], tag: &str) -> usize {
    let text = std::str::from_utf8(xml).expect("styles XML is UTF-8");
    let prefix = format!("<{tag}");
    text.match_indices(&prefix)
        .filter(|(offset, _)| {
            text.as_bytes()
                .get(offset + prefix.len())
                .is_none_or(|byte| !byte.is_ascii_alphanumeric())
        })
        .count()
}

fn bold(workbook: &Package, cell: &str) -> Option<bool> {
    workbook
        .workbook_ref()
        .sheet("Sheet1")
        .expect("find sheet")
        .expect("worksheet")
        .style(cell)
        .expect("read cell style")
        .and_then(|style| style.format().expect("read cell format").font().cloned())
        .and_then(|font| font.bold())
}

fn italic(workbook: &Package, cell: &str) -> Option<bool> {
    workbook
        .workbook_ref()
        .sheet("Sheet1")
        .expect("find sheet")
        .expect("worksheet")
        .style(cell)
        .expect("read cell style")
        .and_then(|style| style.format().expect("read cell format").font().cloned())
        .and_then(|font| font.italic())
}

fn apply_font(workbook: &Package, cell: &str, change: FontChange) -> litchi_xlsx::Commit {
    let mut edit = workbook.edit().expect("begin font edit");
    edit.sheet("Sheet1")
        .expect("find sheet")
        .expect("worksheet")
        .font(cell, &change)
        .expect("stage font edit");
    edit.commit().expect("commit font edit")
}

#[test]
fn changing_one_cell_does_not_mutate_a_shared_font_neighbour() {
    let source = fixture();
    let source_bytes = source.to_plain_bytes().expect("source bytes");
    assert_ne!(bold(&source, "A1"), Some(true));
    assert_ne!(bold(&source, "B1"), Some(true));

    let committed = apply_font(&source, "A1", FontChange::new().bold(Some(true)));
    let target = committed.package();

    assert_eq!(bold(target, "A1"), Some(true));
    assert_ne!(bold(target, "B1"), Some(true));
    assert_ne!(bold(target, "C1"), Some(true));
    assert_eq!(
        source.to_plain_bytes().expect("source remains unchanged"),
        source_bytes
    );
}

#[test]
fn one_font_and_one_cell_format_are_reused_for_multiple_targets() {
    let source = fixture();
    let before = styles_xml(&source);
    let mut edit = source.edit().expect("begin shared font edit");
    let change = FontChange::new().bold(Some(true)).italic(Some(true));
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("find sheet")
        .expect("worksheet");
    sheet
        .font("A1", &change)
        .and_then(|sheet| sheet.font("B1", &change))
        .and_then(|sheet| sheet.font("C1", &change))
        .expect("stage shared font edits");
    let target = edit
        .commit()
        .expect("commit shared font edits")
        .into_package();
    let after = styles_xml(&target);

    assert_eq!(
        direct_tag_count(&after, "font"),
        direct_tag_count(&before, "font") + 1
    );
    assert_eq!(
        direct_tag_count(&after, "xf"),
        direct_tag_count(&before, "xf") + 1
    );
    for cell in ["A1", "B1", "C1"] {
        assert_eq!(bold(&target, cell), Some(true), "{cell} bold");
        assert_eq!(italic(&target, cell), Some(true), "{cell} italic");
    }
}

#[test]
fn repeating_the_same_font_request_is_an_exact_noop() {
    let source = fixture();
    let first = apply_font(&source, "A1", FontChange::new().bold(Some(true)));
    let first_bytes = first.package().to_plain_bytes().expect("first bytes");
    let first_styles = styles_xml(first.package());

    let repeated = apply_font(first.package(), "A1", FontChange::new().bold(Some(true)));
    assert!(repeated.patch().is_noop());
    assert_eq!(
        repeated.package().to_plain_bytes().expect("repeat bytes"),
        first_bytes
    );
    assert_eq!(styles_xml(repeated.package()), first_styles);
}

#[test]
fn opaque_stylesheet_markup_survives_font_authoring() {
    let package = Package::new().expect("create package").into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            "</styleSheet>",
            r#"<extLst><ext uri="urn:litchi:font-test"><opaque xmlns="urn:litchi:test" marker="keep-me"/></ext></extLst></styleSheet>"#,
        );
    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    let source = Package::from_opc(package).expect("adopt opaque stylesheet");

    let source = {
        let mut edit = source.edit().expect("seed opaque fixture");
        edit.sheet("Sheet1")
            .expect("find sheet")
            .expect("worksheet")
            .set("A1", "opaque style")
            .expect("seed cell");
        edit.commit().expect("seed commit").into_package()
    };
    let committed = apply_font(&source, "A1", FontChange::new().bold(Some(true)));
    let after_styles = styles_xml(committed.package());
    let after = std::str::from_utf8(&after_styles).expect("styles XML");
    assert!(after.contains("urn:litchi:font-test"));
    assert!(after.contains("marker=\"keep-me\""));
    assert_eq!(bold(committed.package(), "A1"), Some(true));
}

#[test]
fn typical_excel_ignorable_font_markup_survives_the_patch_lifecycle() {
    let source = typical_excel_styles_fixture();
    let before = source.to_plain_bytes().expect("source bytes");
    let before_styles = styles_xml(&source);
    let committed = apply_font(&source, "A1", FontChange::new().bold(Some(true)));
    let after_styles = styles_xml(committed.package());
    let after_text = std::str::from_utf8(&after_styles).expect("edited styles XML");

    assert_eq!(bold(committed.package(), "A1"), Some(true));
    assert!(after_text.contains(r#"mc:Ignorable="x14ac""#));
    assert!(after_text.contains(r#"x14ac:knownFonts="1""#));
    assert!(after_text.contains(r#"x14ac:fontMarker="keep""#));
    assert!(after_text.contains(r#"<x14ac:fontExtension marker="keep"/>"#));
    assert!(after_text.contains(r#"x14ac:xfMarker="keep""#));

    let replayed = source
        .apply(committed.patch())
        .expect("forward replay with ignorable styles")
        .into_package();
    assert_eq!(
        replayed.to_plain_bytes().expect("replayed bytes"),
        committed
            .package()
            .to_plain_bytes()
            .expect("committed bytes")
    );

    let durable = committed.patch().durable().expect("durable font patch");
    let durable_forward = durable.apply(&source).expect("durable forward replay");
    let durable_inverse = durable
        .inverse()
        .apply(&durable_forward)
        .expect("durable inverse replay");
    assert_eq!(
        durable_inverse
            .to_plain_bytes()
            .expect("durable inverse bytes"),
        before
    );

    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse replay with ignorable styles")
        .into_package();
    assert_eq!(styles_xml(&restored), before_styles);
    assert_eq!(restored.to_plain_bytes().expect("restored bytes"), before);

    let repeated = apply_font(
        committed.package(),
        "A1",
        FontChange::new().bold(Some(true)),
    );
    assert!(repeated.patch().is_noop());
    assert_eq!(
        repeated.package().to_plain_bytes().expect("repeat bytes"),
        committed.package().to_plain_bytes().expect("first bytes")
    );
}

#[test]
fn alternate_content_in_styles_refuses_atomically() {
    let package = Package::new().expect("create package").into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace(
            &format!(r#"<styleSheet xmlns="{MAIN_NS}">"#),
            &format!(
                r#"<styleSheet xmlns="{MAIN_NS}" xmlns:mc="{MCE_NS}" xmlns:x14ac="{X14AC_NS}" mc:Ignorable="x14ac">"#
            ),
        )
        .replace(
            r#"</fonts>"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x14ac"><x14ac:fontFuture/></mc:Choice><mc:Fallback/></mc:AlternateContent></fonts>"#,
        );
    let mut package = package;
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(styles.into_bytes());
    let source = Package::from_opc(package).expect("adopt AlternateContent styles");
    let before = source.to_plain_bytes().expect("source bytes");
    let mut edit = source.edit().expect("begin rejected edit");
    edit.sheet("Sheet1")
        .expect("find sheet")
        .expect("worksheet")
        .font("A1", &FontChange::new().bold(Some(true)))
        .expect("stage rejected edit");
    assert!(edit.commit().is_err(), "MCE AlternateContent must refuse");
    assert_eq!(source.to_plain_bytes().expect("unchanged source"), before);
}

#[test]
fn font_patch_replays_and_inverts_with_style_lineage_rebinding() {
    let source = fixture();
    let source_bytes = source.to_plain_bytes().expect("source bytes");
    let source_style = source
        .workbook_ref()
        .sheet("Sheet1")
        .expect("find sheet")
        .expect("worksheet")
        .style("A1")
        .expect("read source style")
        .expect("source style");
    let committed = apply_font(&source, "A1", FontChange::new().bold(Some(true)));
    let target_style = committed
        .package()
        .workbook_ref()
        .sheet("Sheet1")
        .expect("find target sheet")
        .expect("target worksheet")
        .style("A1")
        .expect("read target style")
        .expect("target style");
    assert!(!source_style.same_workbook(&target_style));
    assert!(!source_style.same(&target_style));

    let replayed = source
        .apply(committed.patch())
        .expect("forward replay")
        .into_package();
    assert_eq!(bold(&replayed, "A1"), Some(true));
    assert!(
        !replayed
            .workbook_ref()
            .sheet("Sheet1")
            .expect("find replayed sheet")
            .expect("replayed worksheet")
            .style("A1")
            .expect("read replayed style")
            .expect("replayed style")
            .same(&target_style)
    );

    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse replay")
        .into_package();
    assert_eq!(
        restored.to_plain_bytes().expect("restored bytes"),
        source_bytes
    );
}

#[test]
fn stale_stylesheet_bytes_refuse_a_font_patch_before_publication() {
    let source = fixture();
    let committed = apply_font(&source, "A1", FontChange::new().bold(Some(true)));
    let stale_bytes = source.to_plain_bytes().expect("stale source bytes");
    let mut package = Package::from_bytes(&stale_bytes)
        .expect("open stale candidate")
        .into_plain_opc();
    let styles_uri = PackURI::new(STYLES_URI).expect("styles URI");
    let stale_styles = std::str::from_utf8(package.get_part(&styles_uri).expect("styles").blob())
        .expect("styles XML")
        .replace("</styleSheet>", "<!-- stale stylesheet --></styleSheet>");
    package
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(stale_styles.into_bytes());
    let stale = Package::from_opc(package).expect("adopt stale candidate");
    let before = stale.to_plain_bytes().expect("stale bytes before apply");
    assert!(matches!(
        stale.apply(committed.patch()),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(
        stale.to_plain_bytes().expect("stale bytes after refusal"),
        before
    );
}

#[test]
fn disjoint_font_requests_and_orthogonal_payloads_join_but_overlaps_refuse() {
    let source = fixture();
    let mut left = source.edit().expect("left edit");
    left.sheet("Sheet1")
        .expect("left sheet")
        .expect("left worksheet")
        .font("A1", &FontChange::new().bold(Some(true)))
        .expect("left font");
    let mut right = source.edit().expect("right edit");
    right
        .sheet("Sheet1")
        .expect("right sheet")
        .expect("right worksheet")
        .font("B1", &FontChange::new().italic(Some(true)))
        .expect("right font");
    left.join(right).expect("disjoint font requests join");
    let joined = left.commit().expect("disjoint font commit").into_package();
    assert_eq!(bold(&joined, "A1"), Some(true));
    assert_eq!(italic(&joined, "B1"), Some(true));

    let mut bold_edit = source.edit().expect("same-cell bold edit");
    bold_edit
        .sheet("Sheet1")
        .expect("bold sheet")
        .expect("bold worksheet")
        .font("A1", &FontChange::new().bold(Some(true)))
        .expect("same-cell bold");
    let mut italic_edit = source.edit().expect("same-cell italic edit");
    italic_edit
        .sheet("Sheet1")
        .expect("italic sheet")
        .expect("italic worksheet")
        .font("A1", &FontChange::new().italic(Some(true)))
        .expect("same-cell italic");
    bold_edit
        .join(italic_edit)
        .expect("orthogonal same-cell font fields join");
    let joined = bold_edit
        .commit()
        .expect("same-cell font commit")
        .into_package();
    assert_eq!(bold(&joined, "A1"), Some(true));
    assert_eq!(italic(&joined, "A1"), Some(true));

    let mut font = source.edit().expect("font edit");
    font.sheet("Sheet1")
        .expect("font sheet")
        .expect("font worksheet")
        .font("A1", &FontChange::new().bold(Some(true)))
        .expect("font change");
    let mut payload = source.edit().expect("payload edit");
    payload
        .sheet("Sheet1")
        .expect("payload sheet")
        .expect("payload worksheet")
        .set("A1", "joined payload")
        .expect("payload change");
    font.join(payload).expect("font and payload join");
    let joined = font.commit().expect("joined commit").into_package();
    assert_eq!(bold(&joined, "A1"), Some(true));
    assert!(matches!(
        joined
            .workbook_ref().sheet("Sheet1")
            .expect("joined sheet")
            .expect("joined worksheet")
            .cell("A1")
            .expect("joined A1")
            .stored(),
        Some(litchi_xlsx::Cell::Value(litchi_xlsx::Value::Text(value)))
            if value.as_str() == "joined payload"
    ));

    let mut overlapping_left = source.edit().expect("overlapping left edit");
    overlapping_left
        .sheet("Sheet1")
        .expect("overlapping left sheet")
        .expect("overlapping left worksheet")
        .font("A1", &FontChange::new().bold(Some(true)))
        .expect("overlapping left font");
    let mut overlapping_right = source.edit().expect("overlapping right edit");
    overlapping_right
        .sheet("Sheet1")
        .expect("overlapping right sheet")
        .expect("overlapping right worksheet")
        .font("A1", &FontChange::new().bold(Some(false)))
        .expect("overlapping right font");
    assert!(overlapping_left.join(overlapping_right).is_err());
}

#[test]
fn font_and_cell_removal_conflict_in_both_join_orders() {
    let source = fixture();
    for font_first in [false, true] {
        let mut font = source.edit().expect("font edit");
        font.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .font("A1", &FontChange::new().bold(Some(true)))
            .unwrap();
        let mut removal = source.edit().expect("removal edit");
        removal
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .remove("A1")
            .unwrap();
        let result = if font_first {
            font.join(removal)
        } else {
            removal.join(font)
        };
        assert!(result.is_err());
    }
}

#[test]
fn durable_font_patch_replays_forward_and_inverse() {
    let source = fixture();
    let source_bytes = source.to_plain_bytes().expect("source bytes");
    let committed = apply_font(&source, "A1", FontChange::new().bold(Some(true)));
    let durable = committed.patch().durable().expect("durable font patch");
    let applied = durable.apply(&source).expect("durable forward replay");
    assert_eq!(bold(&applied, "A1"), Some(true));
    let restored = durable
        .inverse()
        .apply(&applied)
        .expect("durable inverse replay");
    assert_eq!(
        restored.to_plain_bytes().expect("durable restored bytes"),
        source_bytes
    );
}
