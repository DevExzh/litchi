use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_odf_common::{constants, core::PackageWriter};
use litchi_ods::{
    Cell, CellValue, HyperlinkChange, HyperlinkOperation, Link, SourceBackedSpreadsheet,
    Spreadsheet,
};

const CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
    r#"<office:body><office:spreadsheet><table:table table:name="Data">"#,
    r#"<table:table-row><table:table-cell office:value-type="string">"#,
    r#"<text:p>prefix<text:a xlink:type="simple" xlink:href="https://example.test/old">"#,
    r#"linked</text:a> suffix</text:p></table:table-cell></table:table-row>"#,
    r#"</table:table></office:spreadsheet></office:body></office:document-content>"#,
);

fn package(content: &str) -> litchi_core::Result<Vec<u8>> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET)?;
    writer.add_file("content.xml", content.as_bytes())?;
    writer.finish_to_bytes()
}

fn text(value: &str) -> Cell {
    Cell::new(CellValue::Text(value.to_string()), value)
}

fn first_link(bytes: Vec<u8>) -> litchi_core::Result<Link> {
    let spreadsheet = Spreadsheet::from_bytes(bytes)?;
    let Some(litchi_ods::CellView::Stored(cell)) = spreadsheet.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "source hyperlink fixture cell is missing".to_string(),
        ));
    };
    cell.hyperlinks().first().cloned().ok_or_else(|| {
        litchi_core::Error::InvalidFormat(
            "source hyperlink fixture contains no direct link".to_string(),
        )
    })
}

#[test]
fn source_hyperlink_batch_rewrites_direct_links_and_supports_exact_no_op() -> litchi_core::Result<()>
{
    let source = package(CONTENT)?;
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone())))?;
    let mut no_op = owner.edit_cells()?;
    assert_eq!(
        no_op.set_hyperlinks(
            "Data",
            vec![HyperlinkChange::new(
                0,
                0,
                vec![HyperlinkOperation::remove(99)],
            )],
        )?,
        Some(0)
    );
    let no_op_commit = no_op.commit()?;
    assert!(!no_op_commit.changed());
    assert_eq!(no_op_commit.snapshot().content_xml(), CONTENT);
    let mut no_op_output = Vec::new();
    no_op_commit.write_to(&mut no_op_output).unwrap();
    assert_eq!(no_op_output, source);

    let mut edit = owner.edit_cells()?;
    let replacement = Link::with_text("https://example.test/new", "linked")?;
    let added = Link::with_text("https://example.test/added", "suffix")?;
    assert_eq!(
        edit.set_hyperlinks(
            "Data",
            vec![HyperlinkChange::new(
                0,
                0,
                vec![
                    HyperlinkOperation::replace_at(6..12, replacement),
                    HyperlinkOperation::add(13..19, added),
                ],
            )],
        )?,
        Some(1)
    );
    let commit = edit.commit()?;
    let mut output = Vec::new();
    commit.write_to(&mut output).unwrap();
    let reopened = Spreadsheet::from_bytes(output.clone())?;
    let Some(litchi_ods::CellView::Stored(cell)) = reopened.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "source hyperlink cell is missing after publication".to_string(),
        ));
    };
    assert_eq!(cell.hyperlinks().len(), 2);
    assert_eq!(cell.hyperlinks()[0].href(), "https://example.test/new");
    assert_eq!(cell.hyperlinks()[1].href(), "https://example.test/added");

    let chained_owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(output)))?;
    let mut chained = chained_owner.edit_cells()?;
    let chained_link = Link::with_text("https://example.test/chained", "linked")?;
    assert_eq!(
        chained.set_hyperlinks(
            "Data",
            vec![HyperlinkChange::replace(0, 0, 0, chained_link)],
        )?,
        Some(1)
    );
    let chained_commit = chained.commit()?;
    let mut chained_output = Vec::new();
    chained_commit.write_to(&mut chained_output).unwrap();
    let chained_reopen = Spreadsheet::from_bytes(chained_output)?;
    let Some(litchi_ods::CellView::Stored(cell)) = chained_reopen.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "chained source hyperlink cell is missing after publication".to_string(),
        ));
    };
    assert_eq!(cell.hyperlinks()[0].href(), "https://example.test/chained");
    Ok(())
}

#[test]
fn source_hyperlink_operations_are_ordered_and_failure_atomic() -> litchi_core::Result<()> {
    let source = package(CONTENT)?;
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone())))?;
    let mut edit = owner.edit_cells()?;
    let set = Link::with_text("https://example.test/set", "whole")?;
    let replacement = Link::with_text("https://example.test/replaced", "whole")?;
    let final_link = Link::with_text("https://example.test/final", "whole")?;
    assert_eq!(
        edit.set_hyperlinks(
            "Data",
            vec![HyperlinkChange::new(
                0,
                0,
                vec![
                    HyperlinkOperation::clear(),
                    HyperlinkOperation::set(set),
                    HyperlinkOperation::replace(0, replacement),
                    HyperlinkOperation::remove(0),
                    HyperlinkOperation::add(0..5, final_link),
                ],
            )],
        )?,
        Some(1)
    );
    let commit = edit.commit()?;
    let mut output = Vec::new();
    commit.write_to(&mut output).unwrap();
    let reopened = Spreadsheet::from_bytes(output)?;
    let Some(litchi_ods::CellView::Stored(cell)) = reopened.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "ordered source hyperlink cell is missing after publication".to_string(),
        ));
    };
    assert_eq!(cell.text, "whole");
    assert_eq!(cell.hyperlinks().len(), 1);
    assert_eq!(cell.hyperlinks()[0].href(), "https://example.test/final");

    let unsafe_content = CONTENT.replace("https://example.test/old", "JaVaScRiPt:alert%281%29");
    let unsafe_link = first_link(package(&unsafe_content)?)?;
    let mut rejected = owner.edit_cells()?;
    assert!(
        rejected
            .set_hyperlinks(
                "Data",
                vec![HyperlinkChange::new(
                    0,
                    0,
                    vec![
                        HyperlinkOperation::clear(),
                        HyperlinkOperation::add(6..12, unsafe_link),
                    ],
                )],
            )
            .is_err()
    );
    assert_eq!(rejected.changed_cells(), 0);
    assert!(rejected.is_no_op());
    Ok(())
}

#[test]
fn source_hyperlink_preserves_unchanged_legacy_unsafe_links() -> litchi_core::Result<()> {
    let unsafe_content = CONTENT.replace("https://example.test/old", "JaVaScRiPt:alert%281%29");
    let source = package(&unsafe_content)?;
    let legacy = first_link(source.clone())?;
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone())))?;

    let mut no_op = owner.edit_cells()?;
    assert_eq!(
        no_op.set_hyperlinks("Data", vec![HyperlinkChange::replace(0, 0, 0, legacy)],)?,
        Some(0)
    );
    let no_op_commit = no_op.commit()?;
    let mut no_op_output = Vec::new();
    no_op_commit.write_to(&mut no_op_output).unwrap();
    assert_eq!(no_op_output, source);

    let mut edit = owner.edit_cells()?;
    let safe = Link::with_text("https://example.test/safe", "suffix")?;
    assert_eq!(
        edit.set_hyperlinks("Data", vec![HyperlinkChange::add(0, 0, 13..19, safe)])?,
        Some(1)
    );
    let commit = edit.commit()?;
    let mut output = Vec::new();
    commit.write_to(&mut output).unwrap();
    let reopened = Spreadsheet::from_bytes(output)?;
    let Some(litchi_ods::CellView::Stored(cell)) = reopened.cell("Data", 0, 0) else {
        return Err(litchi_core::Error::InvalidFormat(
            "legacy source hyperlink cell is missing after publication".to_string(),
        ));
    };
    assert_eq!(cell.hyperlinks().len(), 2);
    assert_eq!(cell.hyperlinks()[0].href(), "JaVaScRiPt:alert%281%29");
    assert_eq!(cell.hyperlinks()[1].href(), "https://example.test/safe");
    Ok(())
}

#[test]
fn source_hyperlink_batch_preserves_scalar_refusal_and_rejects_repeated_rows()
-> litchi_core::Result<()> {
    let source = package(CONTENT)?;
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone())))?;
    let mut scalar = owner.edit_cells()?;
    assert!(scalar.set_cell("Data", 0, 0, text("changed")).is_err());
    assert_eq!(scalar.changed_cells(), 0);

    let repeated_content = CONTENT.replace(
        "<table:table-row>",
        "<table:table-row table:number-rows-repeated=\"2\">",
    );
    let repeated = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package(
        &repeated_content,
    )?)))?;
    let mut edit = repeated.edit_cells()?;
    assert!(
        edit.set_hyperlinks(
            "Data",
            vec![HyperlinkChange::new(
                1,
                0,
                vec![HyperlinkOperation::clear()],
            )],
        )
        .is_err()
    );
    assert_eq!(edit.changed_cells(), 0);
    Ok(())
}
