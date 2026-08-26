use litchi_core::OwnedSource;
use litchi_odf_common::{constants, core::PackageWriter, package::raw_identical_members};
use litchi_ods::{
    Actuate, Builder, Cell, CellValue, CellView, FlatSnapshot, Link, Show, SourceBackedSpreadsheet,
    Spreadsheet,
    worksheet::{CellChange, Snapshot},
};
use std::sync::Arc;

const LINK_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    r#"<office:body><office:spreadsheet><table:table table:name="Data">"#,
    r#"<table:table-row><table:table-cell office:value-type="string">"#,
    r#"<text:p>前置<text:a xlink:type="simple" xlink:href="https://example.test/jp" "#,
    r#"xlink:show="new" xlink:actuate="onRequest" office:name="jp-link" "#,
    r#"office:title="Japanese" office:target-frame-name="_blank" "#,
    r#"text:style-name="Internet_20_link" "#,
    r#"text:visited-style-name="Visited_20_Internet_20_link">日本語</text:a>"#,
    r##" and <text:a xlink:type="simple" xlink:href="#Sheet2.A1">link</text:a>尾</text:p>"##,
    r#"</table:table-cell><table:table-cell office:value-type="string">"#,
    r#"<text:p>untouched</text:p></table:table-cell></table:table-row>"#,
    r#"<table:table-row><table:table-cell office:value-type="string">"#,
    r#"<text:p>untouched row</text:p></table:table-cell></table:table-row>"#,
    r#"</table:table></office:spreadsheet></office:body></office:document-content>"#,
);

const NESTED_LINK_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    r#"<office:body><office:spreadsheet><table:table table:name="Data"><table:table-row>"#,
    r#"<table:table-cell office:value-type="string"><text:p><text:span>outer<text:a xlink:type="simple" "#,
    r#"xlink:href="https://example.test/inner">inner</text:a></text:span></text:p>"#,
    r#"</table:table-cell></table:table-row></table:table></office:spreadsheet>"#,
    r#"</office:body></office:document-content>"#,
);

const FOREIGN_INLINE_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:vendor="urn:example:vendor" "#,
    r#"office:version="1.4"><office:body><office:spreadsheet>"#,
    r#"<table:table table:name="Data"><table:table-row><table:table-cell "#,
    r#"office:value-type="string"><text:p><text:a xlink:type="simple" "#,
    r#"xlink:href="https://example.test/foreign">before<vendor:foreign/>after</text:a></text:p>"#,
    r#"</table:table-cell></table:table-row></table:table></office:spreadsheet>"#,
    r#"</office:body></office:document-content>"#,
);

fn package(content_xml: &str) -> litchi_core::Result<Vec<u8>> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET)?;
    writer.add_file("content.xml", content_xml.as_bytes())?;
    writer.add_file_with_media_type(
        "Pictures/opaque.bin",
        &[0x5a; 128 * 1024],
        "application/octet-stream",
    )?;
    writer.finish_to_bytes()
}

fn text_cell(text: &str) -> Cell {
    Cell::new(CellValue::Text(text.to_string()), text)
}

fn link(href: &str, text: &str) -> litchi_core::Result<Link> {
    Link::with_text(href, text)
}

fn assert_unchanged_after(
    cell: &mut Cell,
    operation: impl FnOnce(&mut Cell) -> litchi_core::Result<()>,
) {
    let before = cell.clone();
    assert!(operation(cell).is_err());
    assert_eq!(*cell, before);
}

#[test]
fn parses_unicode_multiple_direct_links_and_all_link_metadata() -> litchi_core::Result<()> {
    let source = Builder::new().content_xml(LINK_CONTENT).build()?;
    let spreadsheet = Spreadsheet::from_bytes(source)?;
    let Some(CellView::Stored(cell)) = spreadsheet.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "link fixture cell is missing".to_string(),
        ));
    };

    assert_eq!(cell.text, "前置日本語 and link尾");
    assert_eq!(cell.hyperlinks.len(), 2);
    let first = &cell.hyperlinks[0];
    assert_eq!(first.href(), "https://example.test/jp");
    assert_eq!(first.text(), "日本語");
    assert_eq!(first.range(), "前置".len().."前置日本語".len());
    assert_eq!(first.name.as_deref(), Some("jp-link"));
    assert_eq!(first.title.as_deref(), Some("Japanese"));
    assert_eq!(first.target_frame_name.as_deref(), Some("_blank"));
    assert_eq!(first.show, Some(Show::New));
    assert_eq!(first.actuate, Some(Actuate::OnRequest));
    assert_eq!(first.style_name.as_deref(), Some("Internet_20_link"));
    assert_eq!(
        first.visited_style_name.as_deref(),
        Some("Visited_20_Internet_20_link")
    );

    let second = &cell.hyperlinks[1];
    let second_start = "前置日本語 and ".len();
    assert_eq!(second.href(), "#Sheet2.A1");
    assert_eq!(second.text(), "link");
    assert_eq!(second.range(), second_start..second_start + "link".len());
    assert!(second.name.is_none());
    assert!(second.show.is_none());
    assert!(cell.hyperlink_at(2).is_none());
    assert!(cell.has_hyperlinks());
    assert_eq!(cell.links(), cell.hyperlinks());
    Ok(())
}

#[test]
fn cell_link_crud_preserves_ranges_and_text() -> litchi_core::Result<()> {
    let text = "日本語 and link";
    let first_range = 0.."日本語".len();
    let second_start = "日本語 and ".len();
    let second_range = second_start..second_start + "link".len();
    let mut cell = text_cell(text);

    let mut first = link("https://example.test/one", "日本語")?;
    first.name = Some("first".to_string());
    cell.add_hyperlink(first_range.clone(), first.clone())?;
    cell.add_link(
        second_range.clone(),
        link("https://example.test/two", "link")?,
    )?;
    assert_eq!(cell.hyperlinks().len(), 2);
    assert_eq!(cell.hyperlink(), Some(&first));
    assert_eq!(cell.hyperlinks()[0].range(), first_range);
    assert_eq!(cell.hyperlinks()[1].range(), second_range);

    let replacement = link("https://example.test/replaced", "日本語")?;
    assert_eq!(cell.replace_hyperlink(0, replacement.clone())?, Some(first));
    assert_eq!(cell.hyperlink_at(0), Some(&replacement));

    let replacement_two = link("https://example.test/two-replaced", "link")?;
    let mut expected_previous_two = link("https://example.test/two", "link")?;
    expected_previous_two.set_range(second_range.clone());
    let mut expected_replacement_two = replacement_two.clone();
    expected_replacement_two.set_range(second_range.clone());
    assert_eq!(
        cell.replace_hyperlink_at(second_range.clone(), replacement_two.clone())?,
        Some(expected_previous_two)
    );
    assert_eq!(cell.link(1), Some(&expected_replacement_two));

    assert_eq!(cell.remove_link(0), Some(replacement));
    assert_eq!(cell.remove_hyperlink(9), None);
    assert_eq!(cell.clear_links(), vec![expected_replacement_two]);
    assert!(cell.hyperlinks().is_empty());
    assert_eq!(cell.text, text);

    cell.set_hyperlink(link("https://example.test/all", "全部")?)?;
    assert_eq!(cell.text, "全部");
    assert_eq!(cell.hyperlinks().len(), 1);
    assert_eq!(cell.hyperlinks()[0].range(), 0.."全部".len());
    Ok(())
}

#[test]
fn invalid_link_authoring_is_failure_atomic() -> litchi_core::Result<()> {
    let mut invalid_range = text_cell("éclair");
    assert_unchanged_after(&mut invalid_range, |cell| {
        cell.add_hyperlink(1..2, link("https://example.test", "é")?)
    });

    let mut invalid_text = text_cell("anchor");
    assert_unchanged_after(&mut invalid_text, |cell| {
        cell.add_hyperlink(0..6, link("https://example.test", "other")?)
    });

    let mut overlap = text_cell("first second");
    overlap.add_hyperlink(0..5, link("https://example.test/first", "first")?)?;
    assert_unchanged_after(&mut overlap, |cell| {
        cell.add_hyperlink(3..9, link("https://example.test/overlap", "st sec")?)
    });

    let mut covered = text_cell("covered");
    covered.set_covered(true);
    assert_unchanged_after(&mut covered, |cell| {
        cell.add_hyperlink(0..7, link("https://example.test/covered", "covered")?)
    });

    let mut unsafe_href = text_cell("unsafe");
    let mut unsafe_link = Link::new("javascript:alert(1)");
    unsafe_link.text = "unsafe".to_string();
    unsafe_link.set_range(0..6);
    assert_unchanged_after(&mut unsafe_href, |cell| {
        cell.add_hyperlink(0..6, unsafe_link)
    });

    let mut backwards = text_cell("range");
    assert_unchanged_after(&mut backwards, |cell| {
        cell.add_hyperlink(
            std::ops::Range { start: 4, end: 1 },
            link("https://example.test", "")?,
        )
    });
    Ok(())
}

#[test]
fn owned_and_unified_link_commits_reopen_preserve_opaque_members_and_inverse()
-> litchi_core::Result<()> {
    let source = package(LINK_CONTENT)?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let original = snapshot.sheets()[0]
        .cell(0, 0)
        .ok_or_else(|| litchi_core::Error::InvalidFormat("link cell missing".to_string()))?
        .clone();

    let mut no_op = snapshot.edit();
    assert_eq!(no_op.set_cell("Data", 0, 0, original.clone())?, Some(()));
    let no_op_commit = no_op.commit()?;
    assert!(!no_op_commit.changed());
    assert_eq!(no_op_commit.snapshot().as_bytes(), source);

    let mut changed = original.clone();
    assert!(
        changed
            .replace_hyperlink_at(
                changed.hyperlinks()[0].range(),
                link("https://example.test/edited", "日本語")?,
            )?
            .is_some()
    );
    let mut edit = snapshot.edit();
    assert_eq!(edit.set_cell("Data", 0, 0, changed.clone())?, Some(()));
    let commit = edit.commit()?;
    assert!(commit.changed());
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    let Some(CellView::Stored(cell)) = reopened.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "reopened link cell is missing".to_string(),
        ));
    };
    assert_eq!(cell.hyperlinks()[0].href(), "https://example.test/edited");
    assert_eq!(cell.hyperlinks()[1].href(), "#Sheet2.A1");
    let Some(CellView::Stored(untouched)) = reopened.cell("Data", 1, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "untouched link row is missing".to_string(),
        ));
    };
    assert_eq!(untouched.text, "untouched row");

    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), source);
    let identical = raw_identical_members(&source, commit.snapshot().as_bytes())
        .ok_or_else(|| litchi_core::Error::InvalidFormat("raw comparison failed".to_string()))?;
    assert!(identical.contains("mimetype"));
    assert!(identical.contains("META-INF/manifest.xml"));
    assert!(identical.contains("Pictures/opaque.bin"));
    assert!(!identical.contains("content.xml"));

    let unified_source = litchi_ods::document::Snapshot::from_bytes(source.clone())?;
    let mut unified_noop_edit = unified_source.edit();
    unified_noop_edit.worksheets(|worksheets| {
        assert_eq!(
            worksheets.set_cell("Data", 0, 0, original.clone())?,
            Some(())
        );
        Ok(())
    })?;
    let unified_noop_commit = unified_noop_edit.commit()?;
    assert!(!unified_noop_commit.changed());
    assert_eq!(unified_noop_commit.snapshot().as_bytes(), source);

    let mut unified_edit = unified_source.edit();
    unified_edit.worksheets(|worksheets| {
        worksheets
            .set_cell("Data", 0, 0, changed.clone())?
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat("unified link sheet missing".to_string())
            })?;
        Ok(())
    })?;
    let unified_commit = unified_edit.commit()?;
    let unified_reopened = Spreadsheet::from_bytes(unified_commit.snapshot().as_bytes().to_vec())?;
    let Some(CellView::Stored(cell)) = unified_reopened.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "unified reopened link cell is missing".to_string(),
        ));
    };
    assert_eq!(cell.hyperlinks()[0].href(), "https://example.test/edited");
    let unified_restored = unified_commit
        .patch()
        .inverse()
        .apply(unified_commit.snapshot())?;
    assert_eq!(unified_restored.snapshot().as_bytes(), source);
    Ok(())
}

#[test]
fn invalid_owned_link_batch_does_not_mutate_the_draft() -> litchi_core::Result<()> {
    let source = Builder::new().content_xml(LINK_CONTENT).build()?;
    let snapshot = Snapshot::from_bytes(source)?;
    let mut edit = snapshot.edit();
    let before = edit.sheets().to_vec();
    let first = text_cell("one");
    let second = Cell::repeated(CellValue::Text("two".to_string()), "two", 2)?;
    let error = edit
        .set_cells(
            "Data",
            vec![CellChange::new(0, 0, first), CellChange::new(0, 1, second)],
        )
        .unwrap_err();
    assert!(error.to_string().contains("non-repeated"));
    assert_eq!(edit.sheets(), before);
    Ok(())
}

#[test]
fn nested_and_foreign_inline_markup_respects_lossless_read_boundary() -> litchi_core::Result<()> {
    let nested_source = package(NESTED_LINK_CONTENT)?;
    let nested = Spreadsheet::from_bytes(nested_source.clone())?;
    let Some(CellView::Stored(cell)) = nested.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "nested link cell is missing".to_string(),
        ));
    };
    assert_eq!(cell.text, "outerinner");
    assert!(cell.hyperlinks().is_empty());

    let nested_snapshot = Snapshot::from_bytes(nested_source.clone())?;
    let mut nested_edit = nested_snapshot.edit();
    assert_eq!(
        nested_edit.set_cell("Data", 0, 0, text_cell("changed"))?,
        Some(())
    );
    let error = match nested_edit.commit() {
        Ok(_) => panic!("unsupported nested inline row unexpectedly published"),
        Err(error) => error,
    };
    assert!(error.to_string().contains("unsupported"));
    assert_eq!(nested_snapshot.as_bytes(), nested_source);

    let foreign_source = package(FOREIGN_INLINE_CONTENT)?;
    assert!(Spreadsheet::from_bytes(foreign_source).is_err());
    Ok(())
}

#[test]
fn owned_publication_rejects_new_unsafe_links_but_preserves_legacy_ones() -> litchi_core::Result<()>
{
    let source = package(LINK_CONTENT)?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let original = snapshot.sheets()[0]
        .cell(0, 0)
        .ok_or_else(|| litchi_core::Error::InvalidFormat("link cell missing".to_string()))?
        .clone();
    let mut bypass = original;
    let mut dangerous = Link::new("javascript:new-dangerous-link");
    dangerous.text = "前置".to_string();
    dangerous.set_range(0.."前置".len());
    bypass.hyperlinks.insert(0, dangerous);

    let mut edit = snapshot.edit();
    let before_sheets = edit.sheets().to_vec();
    let error = edit.set_cell("Data", 0, 0, bypass).unwrap_err();
    assert!(matches!(error, litchi_core::Error::InvalidFormat(_)));
    assert_eq!(edit.sheets(), before_sheets);
    assert_eq!(snapshot.as_bytes(), source);

    let legacy_content = LINK_CONTENT.replace(
        "https://example.test/jp",
        "javascript:legacy-dangerous-link",
    );
    let legacy_source = package(&legacy_content)?;
    let legacy_snapshot = Snapshot::from_bytes(legacy_source)?;
    let mut legacy_edit = legacy_snapshot.edit();
    assert_eq!(
        legacy_edit.set_cell("Data", 0, 1, text_cell("safe replacement"))?,
        Some(())
    );
    let legacy_commit = legacy_edit.commit()?;
    let reopened = Spreadsheet::from_bytes(legacy_commit.snapshot().as_bytes().to_vec())?;
    let Some(CellView::Stored(link_cell)) = reopened.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "legacy link cell is missing".to_string(),
        ));
    };
    assert_eq!(
        link_cell.hyperlinks()[0].href(),
        "javascript:legacy-dangerous-link"
    );
    let Some(CellView::Stored(safe_cell)) = reopened.cell("Data", 0, 1) else {
        return Err(litchi_core::Error::InvalidFormat(
            "safe replacement cell is missing".to_string(),
        ));
    };
    assert_eq!(safe_cell.text, "safe replacement");
    Ok(())
}

#[test]
fn flat_transaction_rejects_new_unsafe_public_link_without_staging() -> litchi_core::Result<()> {
    let flat_xml = LINK_CONTENT
        .replace(
            "<office:document-content ",
            "<office:document office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\" ",
        )
        .replace(" office:version=\"1.4\">", ">")
        .replace("</office:document-content>", "</office:document>");
    let flat = FlatSnapshot::from_bytes(flat_xml.into_bytes())?;
    let before = flat.as_bytes().to_vec();
    let Some(CellView::Stored(source_cell)) = flat.cell("Data", 0, 0)? else {
        return Err(litchi_core::Error::InvalidFormat(
            "flat link cell is missing".to_string(),
        ));
    };
    let mut bypass = source_cell.clone();
    let mut dangerous = Link::new("javascript:flat-dangerous-link");
    dangerous.text = "前置".to_string();
    dangerous.set_range(0.."前置".len());
    bypass.hyperlinks.insert(0, dangerous);

    let mut edit = flat.transaction()?;
    assert_eq!(edit.set_cell("Data", 0, 0, bypass)?, Some(()));
    let error = match edit.commit() {
        Ok(_) => panic!("unsafe flat hyperlink commit unexpectedly succeeded"),
        Err(error) => error,
    };
    assert!(matches!(error, litchi_core::Error::InvalidFormat(_)));
    assert_eq!(flat.as_bytes(), before.as_slice());
    Ok(())
}

#[test]
fn source_edit_refuses_linked_row_without_publication() -> litchi_core::Result<()> {
    let source = package(LINK_CONTENT)?;
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone())))?;
    let before = owner.cell_snapshot()?;
    let original_content = before.content_xml().to_string();
    let mut edit = owner.edit_cells()?;
    let error = edit
        .set_cell("Data", 0, 1, text_cell("changed"))
        .unwrap_err();
    assert!(error.to_string().contains("hyperlinks"));
    assert_eq!(edit.changed_cells(), 0);
    assert!(edit.is_no_op());
    assert_eq!(edit.before().content_xml(), original_content.as_str());

    let commit = edit.commit()?;
    let mut output = Vec::new();
    let report = commit
        .write_to(&mut output)
        .expect("no-op source-backed commit should publish successfully");
    assert!(report.is_no_op());
    assert_eq!(output, source);
    Ok(())
}

#[test]
fn link_metadata_authoring_rejects_unsafe_target_without_cell_mutation() -> litchi_core::Result<()>
{
    let mut cell = text_cell("safe");
    let before = cell.clone();
    let mut unsafe_link = Link::new("data:text/html,<script>alert(1)</script>");
    unsafe_link.text = "safe".to_string();
    unsafe_link.name = Some("metadata".to_string());
    assert!(cell.add_hyperlink(0..4, unsafe_link).is_err());
    assert_eq!(cell, before);
    Ok(())
}
