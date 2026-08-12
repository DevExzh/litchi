mod support;

use litchi_core::Result;
use litchi_odf_common::{constants, core::PackageWriter, package::raw_identical_members};
use litchi_ods::{
    Cell, CellValue, Row, Spreadsheet,
    document::{LogicalRowEdit, Patch, Snapshot},
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn package_with_content(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype(constants::ODF_SPREADSHEET)
        .expect("fixed ODS mimetype");
    writer
        .add_file("content.xml", content.as_bytes())
        .expect("fixed content.xml");
    writer
        .add_file_with_media_type(
            "Pictures/untouched.bin",
            b"opaque-row-test-payload",
            "application/octet-stream",
        )
        .expect("fixed opaque member");
    writer.finish_to_bytes().expect("fixed ODS package")
}

fn ordinary_content(rows: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}" office:version="1.3"><office:body><office:spreadsheet><t:table t:name="Data">{rows}</t:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

fn text_row(text: &str) -> Row {
    let mut row = Row::new();
    row.push_cell(Cell::new(
        CellValue::Text(text.to_string()),
        text.to_string(),
    ))
    .expect("fixed row fixture");
    row
}

fn logical_texts(bytes: &[u8]) -> Result<Vec<String>> {
    let spreadsheet = Spreadsheet::from_bytes(bytes.to_vec())?;
    let sheet = spreadsheet.sheet("Data").ok_or_else(|| {
        litchi_core::Error::InvalidFormat("logical-row test sheet is missing".to_string())
    })?;
    let mut values = Vec::new();
    for index in 0..sheet.logical_row_count() {
        values.push(
            sheet
                .cell(index, 0)
                .map_or_else(String::new, |cell| cell.text.clone()),
        );
    }
    Ok(values)
}

#[test]
fn sequential_batch_splits_repeats_moves_rows_and_is_durable() -> Result<()> {
    let rows = concat!(
        r#"<t:table-column t:number-columns-repeated="2" t:visibility="collapse"/>"#,
        r#"<t:table-row><t:table-cell office:value-type="string"><x:p>A</x:p></t:table-cell></t:table-row>"#,
        r#"<t:table-row t:visibility="collapse" t:number-rows-repeated='3'><t:table-cell office:value-type="string"><x:p>B</x:p></t:table-cell></t:table-row>"#,
        r#"<t:table-row><t:table-cell office:value-type="string"><x:p>C</x:p></t:table-cell></t:table-row>"#,
    );
    let source = package_with_content(&ordinary_content(rows));
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.edit_logical_rows(
        "Data",
        &[
            LogicalRowEdit::Insert {
                at: 0,
                rows: vec![text_row("X")],
            },
            LogicalRowEdit::Remove { at: 3, count: 1 },
            LogicalRowEdit::Move {
                at: 0,
                count: 1,
                to: 4,
            },
        ],
    )?;
    let commit = edit.commit()?;
    assert_eq!(
        logical_texts(commit.snapshot().as_bytes())?,
        ["A", "B", "B", "C", "X"]
    );
    let content = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
        .content_xml()
        .to_string();
    assert!(content.contains("t:number-rows-repeated='2'"));
    assert!(content.contains("<t:table-column t:number-columns-repeated=\"2\""));
    assert!(content.contains("t:visibility=\"collapse\""));
    assert_eq!(
        commit.patch().apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    let wire = commit.patch().to_deterministic_json()?;
    let decoded = Patch::from_deterministic_json(&wire, snapshot.limits())?;
    assert_eq!(
        decoded.apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    let inverse_wire = decoded.inverse().to_deterministic_json()?;
    let inverse = Patch::from_deterministic_json(&inverse_wire, snapshot.limits())?;
    assert_eq!(
        inverse.apply(commit.snapshot())?.snapshot().as_bytes(),
        source
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );
    let stale = Snapshot::from_bytes(package_with_content(&ordinary_content(r#"<t:table-row/>"#)))?;
    assert!(commit.patch().apply(&stale).is_err());
    let identical =
        raw_identical_members(&source, commit.snapshot().as_bytes()).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("raw member comparison failed".to_string())
        })?;
    assert!(identical.contains("Pictures/untouched.bin"));
    Ok(())
}

#[test]
fn first_middle_last_only_and_large_repeat_stay_physically_bounded() -> Result<()> {
    let rows = r#"<t:table-row t:number-rows-repeated="1000000"><t:table-cell/></t:table-row>"#;
    let source = package_with_content(&ordinary_content(rows));
    let snapshot = Snapshot::from_bytes(source)?;
    let mut edit = snapshot.edit();
    edit.edit_logical_rows(
        "Data",
        &[LogicalRowEdit::Insert {
            at: 500_000,
            rows: vec![text_row("middle")],
        }],
    )?;
    let middle = edit.commit()?.into_snapshot();
    let sheet = Spreadsheet::from_bytes(middle.as_bytes().to_vec())?
        .sheet("Data")
        .cloned()
        .ok_or_else(|| litchi_core::Error::InvalidFormat("Data is missing".to_string()))?;
    assert_eq!(sheet.logical_row_count(), 1_000_001);
    assert_eq!(sheet.rows.len(), 3);
    assert_eq!(sheet.rows[0].repeat(), 500_000);
    assert_eq!(sheet.rows[2].repeat(), 500_000);

    let mut edges = middle.edit();
    edges.edit_logical_rows(
        "Data",
        &[
            LogicalRowEdit::Insert {
                at: 0,
                rows: vec![text_row("first")],
            },
            LogicalRowEdit::Insert {
                at: 1_000_002,
                rows: vec![text_row("last")],
            },
            LogicalRowEdit::Remove { at: 0, count: 1 },
            LogicalRowEdit::Remove {
                at: 1_000_001,
                count: 1,
            },
        ],
    )?;
    assert_eq!(edges.commit()?.snapshot().as_bytes(), middle.as_bytes());

    let only = package_with_content(&ordinary_content(r#"<t:table-row/>"#));
    let only_snapshot = Snapshot::from_bytes(only)?;
    let mut remove_only = only_snapshot.edit();
    remove_only.edit_logical_rows("Data", &[LogicalRowEdit::Remove { at: 0, count: 1 }])?;
    let removed = remove_only.commit()?;
    assert_eq!(
        Spreadsheet::from_bytes(removed.snapshot().as_bytes().to_vec())?
            .sheet("Data")
            .map(|sheet| sheet.logical_row_count()),
        Some(0)
    );
    Ok(())
}

#[test]
fn namespace_aliased_self_closing_table_opens_without_losing_its_start_tag() -> Result<()> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" office:version="1.3"><office:body><office:spreadsheet><t:table t:name='Data'/></office:spreadsheet></office:body></office:document-content>"#
    );
    let snapshot = Snapshot::from_bytes(package_with_content(&content))?;
    let mut edit = snapshot.edit();
    edit.edit_logical_rows(
        "Data",
        &[LogicalRowEdit::Insert {
            at: 0,
            rows: vec![text_row("created")],
        }],
    )?;
    let commit = edit.commit()?;
    let output = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
        .content_xml()
        .to_string();
    assert!(output.contains("<t:table t:name='Data'>"));
    assert!(output.contains("</t:table>"));
    assert_eq!(logical_texts(commit.snapshot().as_bytes())?, ["created"]);

    let explicit = content.replace(
        " t:name='Data'/>",
        " t:name='Data'><t:table-column t:visibility=\"collapse\"/></t:table>",
    );
    let explicit_snapshot = Snapshot::from_bytes(package_with_content(&explicit))?;
    let mut explicit_edit = explicit_snapshot.edit();
    explicit_edit.edit_logical_rows(
        "Data",
        &[LogicalRowEdit::Insert {
            at: 0,
            rows: vec![text_row("explicit")],
        }],
    )?;
    let explicit_commit = explicit_edit.commit()?;
    assert_eq!(
        logical_texts(explicit_commit.snapshot().as_bytes())?,
        ["explicit"]
    );
    assert!(
        Spreadsheet::from_bytes(explicit_commit.snapshot().as_bytes().to_vec())?
            .content_xml()
            .contains("<t:table-column t:visibility=\"collapse\"/>")
    );
    Ok(())
}

#[test]
fn no_op_and_late_batch_failure_leave_exact_candidate_unchanged() -> Result<()> {
    let source = package_with_content(&ordinary_content(
        r#"<t:table-row><t:table-cell office:value-type="string"><x:p>A</x:p></t:table-cell></t:table-row>"#,
    ));
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut noop = snapshot.edit();
    noop.edit_logical_rows(
        "Data",
        &[LogicalRowEdit::Move {
            at: 0,
            count: 1,
            to: 0,
        }],
    )?;
    assert_eq!(noop.as_bytes(), source);
    assert!(!noop.commit()?.patch().changed());

    let mut failed = snapshot.edit();
    let before = failed.as_bytes().to_vec();
    assert!(
        failed
            .edit_logical_rows(
                "Data",
                &[
                    LogicalRowEdit::Insert {
                        at: 1,
                        rows: vec![text_row("temporary")],
                    },
                    LogicalRowEdit::Remove {
                        at: usize::MAX,
                        count: 1,
                    },
                ],
            )
            .is_err()
    );
    assert!(
        failed
            .edit_logical_rows(
                "Data",
                &[LogicalRowEdit::Remove {
                    at: usize::MAX,
                    count: 0,
                }]
            )
            .is_err()
    );
    assert_eq!(failed.as_bytes(), before);
    assert!(
        failed
            .edit_logical_rows("Missing", &[LogicalRowEdit::Remove { at: 0, count: 1 }])
            .is_err()
    );
    Ok(())
}

#[test]
fn logical_row_edit_count_accepts_below_and_exact_limit_and_rejects_above() -> Result<()> {
    const LIMIT: usize = 4_096;
    let source = package_with_content(&ordinary_content(r#"<t:table-row/>"#));
    let snapshot = Snapshot::from_bytes(source.clone())?;
    for count in [LIMIT - 1, LIMIT] {
        let operations = vec![LogicalRowEdit::Remove { at: 0, count: 0 }; count];
        let mut edit = snapshot.edit();
        edit.edit_logical_rows("Data", &operations)?;
        assert_eq!(edit.as_bytes(), source);
    }
    let operations = vec![LogicalRowEdit::Remove { at: 0, count: 0 }; LIMIT + 1];
    let mut refused = snapshot.edit();
    assert!(refused.edit_logical_rows("Data", &operations).is_err());
    assert_eq!(refused.as_bytes(), source);
    Ok(())
}

#[test]
fn source_and_inserted_style_references_are_refused_atomically() -> Result<()> {
    let table_style = ordinary_content(r#"<t:table-row/>"#).replace(
        r#"<t:table t:name="Data">"#,
        r#"<t:table t:name="Data" t:style-name="TableStyle">"#,
    );
    let styled_noop_source = package_with_content(&table_style);
    let source_cases = [
        table_style,
        ordinary_content(
            r#"<t:table-column t:default-cell-style-name="ColumnStyle"/><t:table-row/>"#,
        ),
        ordinary_content(r#"<t:table-row t:style-name="RowStyle"/>"#),
        ordinary_content(r#"<t:table-row><t:table-cell t:style-name="CellStyle"/></t:table-row>"#),
    ];
    for content in source_cases {
        let source = package_with_content(&content);
        let snapshot = Snapshot::from_bytes(source.clone())?;
        let mut edit = snapshot.edit();
        assert!(
            edit.edit_logical_rows(
                "Data",
                &[LogicalRowEdit::Insert {
                    at: 0,
                    rows: vec![text_row("unsafe")],
                }]
            )
            .is_err()
        );
        assert_eq!(edit.as_bytes(), source);
    }

    let styled_noop = Snapshot::from_bytes(styled_noop_source.clone())?;
    let mut noop = styled_noop.edit();
    noop.edit_logical_rows(
        "Data",
        &[LogicalRowEdit::Move {
            at: 0,
            count: 1,
            to: 0,
        }],
    )?;
    assert_eq!(noop.as_bytes(), styled_noop_source);

    let source = package_with_content(&ordinary_content(r#"<t:table-row/>"#));
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut styled_row = text_row("row-style");
    styled_row.style_name = Some("RowStyle".to_string());
    let mut styled_cell_row = Row::new();
    let mut styled_cell = Cell::new(CellValue::Text("cell-style".to_string()), "cell-style");
    styled_cell.style_name = Some("CellStyle".to_string());
    styled_cell_row.push_cell(styled_cell)?;
    for row in [styled_row, styled_cell_row] {
        let mut edit = snapshot.edit();
        assert!(
            edit.edit_logical_rows(
                "Data",
                &[LogicalRowEdit::Insert {
                    at: 0,
                    rows: vec![row],
                }]
            )
            .is_err()
        );
        assert_eq!(edit.as_bytes(), source);
    }
    Ok(())
}

#[test]
fn every_unrewritten_dependency_domain_is_refused_atomically() -> Result<()> {
    let cases = [
        (
            "formula",
            r#"<t:table-row><t:table-cell t:formula="of:=A1" office:value-type="float" office:value="1"/></t:table-row>"#,
        ),
        (
            "merge",
            r#"<t:table-row><t:table-cell t:number-rows-spanned="2" t:number-columns-spanned="1"/></t:table-row>"#,
        ),
        ("validation", r#"<t:content-validations/><t:table-row/>"#),
        (
            "named range",
            r#"<t:table-row/><t:named-expressions><t:named-range t:name="n" t:base-cell-address="Data.A1" t:cell-range-address="Data.A1:Data.A2"/></t:named-expressions>"#,
        ),
        ("tracked changes", r#"<t:tracked-changes/><t:table-row/>"#),
        (
            "annotation",
            r#"<t:table-row><t:table-cell><o:annotation><x:p>note</x:p></o:annotation></t:table-cell></t:table-row>"#,
        ),
        (
            "rich text",
            r#"<t:table-row><t:table-cell office:value-type="string"><x:p><x:span>rich</x:span></x:p></t:table-cell></t:table-row>"#,
        ),
        (
            "foreign extension",
            r#"<t:table-row xmlns:v="urn:vendor" v:owner="opaque"><t:table-cell/></t:table-row>"#,
        ),
    ];
    for (label, rows) in cases {
        let source = package_with_content(&ordinary_content(rows));
        let snapshot = Snapshot::from_bytes(source.clone())?;
        let mut edit = snapshot.edit();
        assert!(
            edit.edit_logical_rows(
                "Data",
                &[LogicalRowEdit::Insert {
                    at: 0,
                    rows: vec![text_row("unsafe")],
                }]
            )
            .is_err(),
            "dependency case unexpectedly accepted: {label}"
        );
        assert_eq!(edit.as_bytes(), source, "atomicity failed for {label}");
    }

    let style_source = support::raw_package(&[
        (
            "content.xml",
            ordinary_content(r#"<t:table-row/>"#).as_bytes(),
            "text/xml",
        ),
        (
            "styles.xml",
            format!(r#"<office:document-styles xmlns:office="{OFFICE}"/>"#).as_bytes(),
            "text/xml",
        ),
    ]);
    let style_snapshot = Snapshot::from_bytes(style_source.clone())?;
    let mut style_edit = style_snapshot.edit();
    assert!(
        style_edit
            .edit_logical_rows("Data", &[LogicalRowEdit::Remove { at: 0, count: 1 }])
            .is_err()
    );
    assert_eq!(style_edit.as_bytes(), style_source);
    Ok(())
}
