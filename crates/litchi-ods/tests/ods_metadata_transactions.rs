use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_ods::{
    Builder, CellView, MutableSpreadsheet, SourceBackedSpreadsheet, Spreadsheet,
    model::source::CellRange,
    model::structure::StyleUsage,
    styles::table_template::{Region, Style, Template},
};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Data"><table:title>Revenue &amp; Plan</table:title><table:desc>Imported source</table:desc><table:table-row><table:table-cell office:value-type="string"><table:cell-range-source table:name="DataRange" table:last-column-spanned="2" table:last-row-spanned="3" xlink:type="simple" xlink:href="source.ods#Data.A1:B3" xlink:actuate="onRequest" table:filter-name="csv"/><text:p>cached</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;

#[test]
fn table_metadata_and_cell_range_source_round_trip_through_owned_and_source_open() {
    let bytes = Builder::new().content_xml(CONTENT).build().expect("build");
    let spreadsheet = Spreadsheet::from_bytes(bytes.clone()).expect("open");
    let sheet = &spreadsheet.sheets()[0];
    assert_eq!(sheet.title(), Some("Revenue & Plan"));
    assert_eq!(sheet.description(), Some("Imported source"));
    let cell = spreadsheet.cell("Data", 0, 0).expect("cell");
    let CellView::Stored(cell) = cell else {
        panic!("stored cell");
    };
    let source = cell.range_source().expect("range source");
    assert_eq!(source.name(), "DataRange");
    assert_eq!(source.rows(), 3);
    assert_eq!(source.columns(), 2);
    assert!(source.actuate_on_request());
    assert_eq!(source.filter_name(), Some("csv"));

    let source_open = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("source open");
    assert!(
        source_open
            .table_templates()
            .expect("source templates")
            .templates()
            .is_empty()
    );
    assert_eq!(
        source_open.sheets().expect("source sheets")[0].title(),
        Some("Revenue & Plan")
    );
    let Some(CellView::Stored(cell)) = source_open.cell("Data", 0, 0).expect("source cell") else {
        panic!("stored source cell");
    };
    assert_eq!(
        cell.range_source().expect("source range").href(),
        "source.ods#Data.A1:B3"
    );
}

#[test]
fn worksheet_metadata_edit_is_failure_atomic_and_reopenable() {
    let bytes = Builder::new().content_xml(CONTENT).build().expect("build");
    let mut mutable = MutableSpreadsheet::from_bytes(bytes).expect("open");
    let mut sheet = mutable.sheets()[0].clone();
    sheet.set_title(Some("Edited".to_string())).expect("title");
    sheet
        .set_description(Some("Edited description".to_string()))
        .expect("description");
    mutable.set_sheets(vec![sheet]).expect("publish");
    let reopened = mutable.spreadsheet();
    assert_eq!(reopened.sheets()[0].title(), Some("Edited"));
    assert_eq!(
        reopened.sheets()[0].description(),
        Some("Edited description")
    );
    assert!(
        reopened
            .content_xml()
            .contains("<table:title>Edited</table:title>")
    );
    assert!(reopened.content_xml().contains(
        r#"<table:cell-range-source table:name="DataRange" table:last-column-spanned="2" table:last-row-spanned="3" xlink:type="simple" xlink:href="source.ods#Data.A1:B3" xlink:actuate="onRequest" table:filter-name="csv"/>"#
    ));
}

#[test]
fn table_template_builder_and_package_lifecycle_reopen() {
    let template = Template::new("Bands").with_region(Region::Body, Style::new("Body"));
    let mut builder = Builder::new();
    builder
        .edit_table_templates(|edit| edit.add(template.clone()))
        .expect("builder template");
    let bytes = builder.build().expect("build with styles");
    let mut spreadsheet = Spreadsheet::from_bytes(bytes).expect("open styles");
    assert_eq!(
        spreadsheet
            .table_templates()
            .expect("templates")
            .templates(),
        &[template]
    );
    spreadsheet
        .edit_table_templates(|edit| {
            edit.replace_named("Bands", {
                Template::new("Bands").with_region(Region::Body, Style::new("Body2"))
            })
        })
        .expect("edit template");
    assert_eq!(
        spreadsheet
            .table_templates()
            .expect("edited templates")
            .templates()[0]
            .body
            .as_ref()
            .expect("body")
            .style_name,
        "Body2"
    );
}

#[test]
fn table_template_edit_creates_missing_styles_owner() {
    let bytes = Builder::new().build().expect("minimal package");
    let mut spreadsheet = Spreadsheet::from_bytes(bytes).expect("open");
    let template = Template::new("Created").with_region(Region::Body, Style::new("Body"));
    spreadsheet
        .edit_table_templates(|edit| edit.add(template.clone()))
        .expect("create styles owner");
    assert_eq!(
        spreadsheet
            .table_templates()
            .expect("templates")
            .templates(),
        &[template]
    );
    let snapshot = spreadsheet.table_templates().expect("snapshot");
    let mut edit = snapshot.edit();
    edit.clear().expect("clear");
    let commit = edit.commit().expect("remove template");
    assert!(commit.changed());
    spreadsheet
        .apply_table_template_patch(commit.patch())
        .expect("apply clear");
    assert!(
        spreadsheet
            .table_templates()
            .expect("empty templates")
            .templates()
            .is_empty()
    );
}

#[test]
fn fresh_table_template_builder_writes_normative_owner_and_axes() {
    let template = Template::new("Created").with_region(Region::Body, Style::new("Body"));
    let mut builder = Builder::new();
    builder
        .edit_table_templates(|edit| edit.add(template))
        .expect("template");
    let spreadsheet = Spreadsheet::from_bytes(builder.build().expect("build")).expect("open");
    let styles = spreadsheet.styles_xml().expect("styles owner");
    assert!(styles.contains("<table:table-template"));
    for attribute in [
        "table:first-row-start-column=",
        "table:first-row-end-column=",
        "table:last-row-start-column=",
        "table:last-row-end-column=",
    ] {
        assert!(styles.contains(attribute), "missing {attribute}");
    }
    assert!(styles.contains("<table:body table:style-name=\"Body\"/"));
    assert!(!styles.contains("table:use-first-row-styles"));
}

#[test]
fn metadata_rewrite_refuses_comment_or_pi_loss() {
    let source = CONTENT.replace(
        "<table:title>Revenue &amp; Plan</table:title>",
        "<table:title><!-- retained? -->Revenue &amp; Plan<?producer keep?></table:title>",
    );
    let bytes = Builder::new()
        .content_xml(source)
        .build()
        .expect("build source");
    let mut mutable = MutableSpreadsheet::from_bytes(bytes).expect("open");
    assert!(
        mutable
            .set_sheet_title("Data", Some("Edited".to_string()))
            .is_err()
    );
}

#[test]
fn row_and_table_metadata_changes_use_the_full_table_fallback() {
    let bytes = Builder::new().content_xml(CONTENT).build().expect("build");
    let mut mutable = MutableSpreadsheet::from_bytes(bytes).expect("open");
    let mut sheet = mutable.sheets()[0].clone();
    sheet.set_title(Some("Edited".to_string())).expect("title");
    sheet
        .set_cell(
            0,
            0,
            litchi_ods::Cell::new(
                litchi_ods::CellValue::Text("changed".to_string()),
                "changed",
            ),
        )
        .expect("cell");
    mutable.set_sheets(vec![sheet]).expect("combined edit");
    let reopened = Spreadsheet::from_bytes(mutable.to_bytes()).expect("reopen");
    assert_eq!(reopened.sheets()[0].title(), Some("Edited"));
    let CellView::Stored(cell) = reopened.cell("Data", 0, 0).expect("cell") else {
        panic!("stored cell");
    };
    assert_eq!(cell.text, "changed");
}

#[test]
fn cell_range_source_builder_rejects_zero_dimensions() {
    assert!(CellRange::new("x", "href", 0, 1).is_err());
}

#[test]
fn cell_range_source_rejects_duplicate_actuate_attributes() {
    let content = CONTENT.replace(
        "xlink:actuate=\"onRequest\"",
        "xlink:actuate=\"onRequest\" xlink:actuate=\"onRequest\"",
    );
    assert!(Builder::new().content_xml(content).build().is_err());

    let content = CONTENT
        .replace(
            "xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
            "xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:alias=\"http://www.w3.org/1999/xlink\"",
        )
        .replace(
            "xlink:actuate=\"onRequest\"",
            "xlink:actuate=\"onRequest\" alias:actuate=\"onRequest\"",
        );
    assert!(Builder::new().content_xml(content).build().is_err());
}

#[test]
fn normative_table_template_usage_round_trips_through_sheet_edit() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Data" table:template-name="Bands" table:use-first-row-styles="true" table:use-last-row-styles="false" table:use-first-column-styles="1" table:use-last-column-styles="0" table:use-banding-rows-styles="true" table:use-banding-columns-styles="false"><table:table-row/></table:table></office:spreadsheet></office:body></office:document-content>"#;
    let bytes = Builder::new().content_xml(content).build().expect("build");
    let spreadsheet = Spreadsheet::from_bytes(bytes.clone()).expect("open");
    let sheet = &spreadsheet.sheets()[0];
    assert_eq!(sheet.template_name.as_deref(), Some("Bands"));
    assert_eq!(
        sheet.style_usage,
        StyleUsage {
            use_first_row_styles: Some(true),
            use_last_row_styles: Some(false),
            use_first_column_styles: Some(true),
            use_last_column_styles: Some(false),
            use_banding_row_styles: Some(true),
            use_banding_column_styles: Some(false),
        }
    );

    let mut metadata_only = MutableSpreadsheet::from_bytes(bytes.clone()).expect("metadata open");
    let mut metadata_sheet = metadata_only.sheets()[0].clone();
    metadata_sheet
        .set_title(Some("Accessible data".to_string()))
        .expect("title");
    metadata_only
        .set_sheets(vec![metadata_sheet])
        .expect("metadata edit");
    let metadata_xml = metadata_only.spreadsheet().content_xml();
    assert!(metadata_xml.contains(
        r#"<table:table table:name="Data" table:template-name="Bands" table:use-first-row-styles="true" table:use-last-row-styles="false" table:use-first-column-styles="1" table:use-last-column-styles="0" table:use-banding-rows-styles="true" table:use-banding-columns-styles="false">"#
    ));
    assert!(metadata_xml.contains("<table:title>Accessible data</table:title>"));

    let mut edited = sheet.clone();
    edited
        .set_template_name(Some("Other Bands".to_string()))
        .expect("template name");
    edited.style_usage.use_last_column_styles = Some(true);
    let mut mutable = MutableSpreadsheet::from_bytes(bytes).expect("mutable open");
    mutable.set_sheets(vec![edited]).expect("sheet edit");
    let xml = mutable.spreadsheet().content_xml();
    assert!(xml.contains(r#"table:template-name="Other Bands""#));
    assert!(xml.contains(r#"table:use-last-column-styles="true""#));

    let reopened = Spreadsheet::from_bytes(mutable.to_bytes()).expect("reopen");
    let sheet = &reopened.sheets()[0];
    assert_eq!(sheet.template_name.as_deref(), Some("Other Bands"));
    assert_eq!(sheet.style_usage.use_last_column_styles, Some(true));
}
