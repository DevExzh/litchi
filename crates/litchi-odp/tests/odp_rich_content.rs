#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odf_common::chart::ChartClass;
use litchi_odf_common::chart::authoring::{
    CachedCell, CachedRow, CachedTable, CachedValue, Definition, SeriesSpec,
};
use litchi_odf_common::core::{OwnedPackage, PackageWriter};
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

#[test]
fn rich_content_objects_transfer_through_durable_history() {
    let source_base = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut source_transaction = source_base.transaction().unwrap();
    source_transaction.add("Source", "transfer owners").unwrap();
    source_transaction
        .add_text_box(
            0usize,
            &TextBox::new("Source Box", RichText::plain("copied text").unwrap()).unwrap(),
        )
        .unwrap();
    source_transaction
        .add_table(
            0usize,
            &Table::new(
                "Source Table",
                vec![vec![Cell::new(RichText::plain("copied cell").unwrap())]],
            )
            .unwrap(),
        )
        .unwrap();
    source_transaction
        .add_form_control(
            0usize,
            &FormControl::new("Source Control", ControlKind::Button)
                .unwrap()
                .with_label("copied control")
                .unwrap(),
        )
        .unwrap();
    let source = source_transaction.commit().unwrap().snapshot().clone();

    let destination_base = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut destination_transaction = destination_base.transaction().unwrap();
    destination_transaction
        .add("Destination", "transfer target")
        .unwrap();
    destination_transaction
        .transfer_text_box_from(&source, "Source Box", 0usize, "Copied Box")
        .unwrap();
    destination_transaction
        .transfer_table_from(&source, "Source Table", 0usize, "Copied Table")
        .unwrap();
    destination_transaction
        .transfer_form_control_from(&source, "Source Control", 0usize, "Copied Control")
        .unwrap();
    let commit = destination_transaction.commit().unwrap();
    let content = commit.snapshot().to_presentation().unwrap();
    assert!(content.content_xml().contains("Copied Box"));
    assert!(content.content_xml().contains("copied text"));
    assert!(content.content_xml().contains("Copied Table"));
    assert!(content.content_xml().contains("copied cell"));
    assert!(content.content_xml().contains("Copied Control"));
    assert!(content.content_xml().contains("copied control"));
    assert!(!content.content_xml().contains("> <"));

    let durable =
        edit::Patch::from_durable_bytes(&commit.patch().to_durable_bytes().unwrap()).unwrap();
    let applied = durable.apply(&destination_base).unwrap();
    assert_eq!(applied.bytes(), commit.snapshot().bytes());
    assert_eq!(
        durable.inverse().apply(&applied).unwrap().bytes(),
        destination_base.bytes()
    );

    let budget = destination_base.bytes().len() + commit.snapshot().bytes().len();
    let mut history = edit::History::new(destination_base.clone(), 2, budget).unwrap();
    history.record(&commit).unwrap();
    let mut history =
        edit::History::from_durable_bytes(&history.to_durable_bytes().unwrap()).unwrap();
    assert_eq!(history.current().bytes(), commit.snapshot().bytes());
    assert_eq!(history.undo().unwrap().bytes(), destination_base.bytes());
    assert_eq!(history.redo().unwrap().bytes(), commit.snapshot().bytes());
}

#[test]
fn source_backed_story_table_and_form_models_edit_without_flattening() {
    let source_base = edit::Snapshot::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut authored = source_base.transaction().unwrap();
    authored.add("Models", "source backed").unwrap();
    authored
        .add_text_box(
            0usize,
            &TextBox::new("Story", RichText::plain("before story").unwrap()).unwrap(),
        )
        .unwrap();
    authored
        .add_table(
            0usize,
            &Table::new(
                "Grid",
                vec![vec![Cell::new(RichText::plain("before cell").unwrap())]],
            )
            .unwrap(),
        )
        .unwrap();
    authored
        .add_form_control(
            0usize,
            &FormControl::new("Choice", ControlKind::Button)
                .unwrap()
                .with_label("before label")
                .unwrap(),
        )
        .unwrap();
    let source = authored.commit().unwrap().snapshot().clone();

    let inventory = source.rich_content().unwrap();
    let mut transaction = source.transaction().unwrap();
    let mut story = inventory
        .text_boxes()
        .iter()
        .find(|model| model.name() == "Story")
        .unwrap()
        .clone();
    story
        .set_xml(story.xml().replace(
            "<text:p>before story</text:p>",
            "<text:list><text:list-item><text:p>after story</text:p></text:list-item></text:list>",
        ))
        .unwrap();
    assert_eq!(story.list_count(), 1);
    transaction.replace_text_box_model("Story", &story).unwrap();

    let mut table = inventory
        .tables()
        .iter()
        .find(|model| model.name() == "Grid")
        .unwrap()
        .clone();
    table
        .set_xml(table.xml().replace("before cell", "after cell").replace(
            "office:value-type=\"string\"",
            "office:value-type=\"string\" table:number-columns-spanned=\"2\"",
        ))
        .unwrap();
    transaction.replace_table_model("Grid", &table).unwrap();

    let mut control = inventory
        .form_controls()
        .iter()
        .find(|model| model.name() == "Choice")
        .unwrap()
        .clone();
    control
        .set_xml(
            control
                .declaration_xml()
                .replace("before label", "after label"),
            control.visual_xml().to_string(),
        )
        .unwrap();
    transaction
        .replace_form_control_model("Choice", &control)
        .unwrap();

    let commit = transaction.commit().unwrap();
    let inventory = commit.snapshot().rich_content().unwrap();
    let story = inventory
        .text_boxes()
        .iter()
        .find(|model| model.name() == "Story")
        .unwrap();
    let table = inventory
        .tables()
        .iter()
        .find(|model| model.name() == "Grid")
        .unwrap();
    let control = inventory
        .form_controls()
        .iter()
        .find(|model| model.name() == "Choice")
        .unwrap();
    assert_eq!(story.list_count(), 1);
    assert_eq!(story.paragraph_count(), 1);
    assert!(story.xml().contains("after story"));
    assert!(table.xml().contains("after cell"));
    assert!(table.xml().contains("table:number-columns-spanned=\"2\""));
    assert!(control.declaration_xml().contains("after label"));
}

#[test]
fn producer_named_drawing_resources_transfer_with_their_payload() {
    const CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles><style:style style:name="ProducerStyle" style:family="graphic"><style:graphic-properties draw:fill="bitmap" draw:fill-image-name="ProducerImage"/></style:style><draw:fill-image draw:name="ProducerImage" xlink:type="simple" xlink:href="Pictures/producer.png"/></office:automatic-styles><office:body><office:presentation><draw:page draw:name="Source"><draw:frame draw:name="Producer Box" draw:style-name="ProducerStyle"><draw:text-box><text:p>producer closure</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer
        .add_file_with_media_type("Pictures/producer.png", b"producer-pixels", "image/png")
        .unwrap();
    let source = edit::Snapshot::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

    let mut destination_builder = Builder::new();
    destination_builder.add_slide("Destination").unwrap();
    let destination = edit::Snapshot::from_bytes(destination_builder.build().unwrap()).unwrap();
    let mut transaction = destination.transaction().unwrap();
    transaction
        .transfer_text_box_from(&source, "Producer Box", 0usize, "Copied Producer Box")
        .unwrap();
    let commit = transaction.commit().unwrap();
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    assert_eq!(
        package.get_file("Pictures/producer.png").unwrap(),
        b"producer-pixels"
    );
    let content = commit.snapshot().to_presentation().unwrap();
    assert!(content.content_xml().contains("ProducerStyle"));
    assert!(content.content_xml().contains("ProducerImage"));
    assert!(content.content_xml().contains("producer closure"));
    assert!(!content.content_xml().contains("> <"));
}
