#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::chart::ChartClass;
use litchi_odf_common::chart::authoring::{
    CachedCell, CachedRow, CachedTable, CachedValue, Definition, SeriesSpec,
};
use litchi_odp::content::{
    Cell, ControlKind, FormControl, Paragraph, RichText, Run, Table, TextBox,
};
use litchi_odp::{Builder, Presentation, edit};

fn chart() -> Definition {
    let mut definition = Definition::new(ChartClass::line());
    definition.plot_area.series.push(SeriesSpec {
        values_cell_range_address: Some("local.B2:.B2".to_string()),
        ..SeriesSpec::default()
    });
    let mut table = CachedTable::new("local", 2);
    table.header_rows.push(CachedRow::new(vec![
        CachedCell::new(CachedValue::String("Label".to_string())),
        CachedCell::new(CachedValue::String("Value".to_string())),
    ]));
    table.rows.push(CachedRow::new(vec![
        CachedCell::new(CachedValue::String("A".to_string())),
        CachedCell::new(CachedValue::Float(1.0)),
    ]));
    definition.cached_table = Some(table);
    definition
}

#[test]
fn rich_text_table_form_and_fine_chart_data_share_one_durable_root() {
    let source = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.add("Rich objects", "Atomic content").unwrap();

    let rich = RichText::new(vec![
        Paragraph::new(vec![
            Run::new("Revenue ").unwrap(),
            Run::new("rose").unwrap().with_style("Emphasis").unwrap(),
        ])
        .unwrap(),
        Paragraph::plain("Second paragraph").unwrap(),
    ])
    .unwrap();
    transaction
        .add_text_box(0usize, &TextBox::new("Rich Box", rich.clone()).unwrap())
        .unwrap();

    let table = Table::new(
        "Summary Table",
        vec![
            vec![
                Cell::new(RichText::plain("Quarter").unwrap()),
                Cell::new(RichText::plain("Revenue").unwrap()),
            ],
            vec![Cell::new(RichText::plain("Q1").unwrap()), Cell::new(rich)],
        ],
    )
    .unwrap();
    transaction.add_table(0usize, &table).unwrap();

    let control = FormControl::new("Approved", ControlKind::Checkbox)
        .unwrap()
        .with_label("Approved")
        .unwrap()
        .with_value("true")
        .unwrap();
    transaction.add_form_control(0usize, &control).unwrap();

    transaction
        .add_chart_definition(
            0usize,
            "Fine Data",
            litchi_odp::charts::Storage::InlineXml,
            &chart(),
        )
        .unwrap();
    transaction
        .replace_chart_cached_cell(
            "Fine Data",
            1,
            1,
            &CachedCell::new(CachedValue::Float(42.0)),
        )
        .unwrap();
    transaction
        .replace_chart_series(
            "Fine Data",
            0,
            &SeriesSpec {
                values_cell_range_address: Some("local.B2:.B2".to_string()),
                label_cell_address: Some("local.B1".to_string()),
                ..SeriesSpec::default()
            },
        )
        .unwrap();
    transaction
        .add_chart_series(
            "Fine Data",
            &SeriesSpec {
                values_cell_range_address: Some("local.B2:.B2".to_string()),
                ..SeriesSpec::default()
            },
        )
        .unwrap();
    transaction.remove_chart_series("Fine Data", 1).unwrap();

    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit.patch().domains(),
        &[
            edit::Domain::Slides,
            edit::Domain::Charts,
            edit::Domain::Content
        ]
    );
    let content = commit.snapshot().to_presentation().unwrap();
    assert_eq!(content.slides().unwrap().len(), 1);
    assert!(content.content_xml().contains("Rich Box"));
    assert!(content.content_xml().contains("Summary Table"));
    assert!(content.content_xml().contains("form:checkbox"));
    assert!(content.content_xml().contains("office:value=\"42\""));
    assert!(!content.content_xml().contains("> <"));

    let durable =
        edit::Patch::from_durable_bytes(&commit.patch().to_durable_bytes().unwrap()).unwrap();
    let applied = durable.apply(&source).unwrap();
    let reopened = Presentation::from_bytes(applied.bytes().to_vec()).unwrap();
    assert!(reopened.content_xml().contains("Rich Box"));
    assert_eq!(
        durable.inverse().apply(&applied).unwrap().bytes(),
        source.bytes()
    );

    let budget = source.bytes().len() + commit.snapshot().bytes().len();
    let mut history = edit::History::new(source.clone(), 2, budget).unwrap();
    history.record(&commit).unwrap();
    assert_eq!(history.undo().unwrap().bytes(), source.bytes());
    assert_eq!(history.redo().unwrap().bytes(), commit.snapshot().bytes());
}

#[test]
fn rich_content_selectors_and_values_fail_without_partial_publication() {
    assert!(Table::new("ragged", vec![vec![], vec![]]).is_err());
    assert!(Run::new("bad\0text").is_err());

    let mut builder = Builder::new();
    builder.add_slide_with_title("One", "body").unwrap();
    let source = edit::Snapshot::from_bytes(builder.build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    let box_one = TextBox::new("duplicate", RichText::plain("one").unwrap()).unwrap();
    transaction.add_text_box(0usize, &box_one).unwrap();
    assert!(transaction.add_text_box(0usize, &box_one).is_err());
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .to_presentation()
            .unwrap()
            .slides()
            .unwrap()
            .len(),
        1
    );
}
