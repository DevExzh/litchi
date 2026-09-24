#![allow(clippy::unwrap_used, reason = "focused semantic projection assertions")]

use litchi_oth::{Builder, Template, change, frame, index, note, table};

const CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:foreign="urn:example:foreign" office:version="1.4">"#,
    r#"<office:body><office:text>"#,
    r#"<text:tracked-changes><text:changed-region xml:id="r1" text:id="r1"><text:insertion><office:change-info><dc:creator>Ada</dc:creator><dc:date>2026-09-02T03:04:05Z</dc:date><text:p>inserted</text:p></office:change-info></text:insertion></text:changed-region></text:tracked-changes>"#,
    r#"<text:section text:name="Intro" text:protected="true"><text:p>section text</text:p></text:section>"#,
    r#"<text:note text:note-class="footnote" text:id="fn1"><text:note-citation text:label="1">1</text:note-citation><text:note-body><text:p>note text</text:p></text:note-body></text:note>"#,
    r#"<office:annotation office:name="comment"><dc:creator>Ada</dc:creator><text:p>review text</text:p></office:annotation>"#,
    r#"<text:change-start text:change-id="r1"/><text:change-end text:change-id="r1"/>"#,
    r#"<text:table-of-content text:name="Contents" text:protected="true"><text:table-of-content-source/><text:index-body><text:p>cached entry</text:p></text:index-body></text:table-of-content>"#,
    r#"<draw:frame draw:name="Box" text:anchor-type="paragraph" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><draw:text-box><text:p>frame text</text:p></draw:text-box></draw:frame>"#,
    r#"<text:ruby text:style-name="Ruby"><text:ruby-base>漢</text:ruby-base><text:ruby-text text:style-name="RubyText">kan</text:ruby-text></text:ruby>"#,
    r#"<table:table table:name="Data"><table:table-column table:number-columns-repeated="2"/><table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="2"><text:p>cell</text:p></table:table-cell><table:covered-table-cell/></table:table-row></table:table>"#,
    r#"</office:text></office:body></office:document-content>"#,
);

#[test]
fn projects_all_oth_body_structure_families() {
    let template =
        Template::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();

    let sections = body.sections().unwrap();
    assert_eq!(sections.len(), 1);
    assert_eq!(sections[0].name(), Some("Intro"));
    assert!(sections[0].is_protected());
    let notes = body.notes().unwrap();
    assert_eq!(notes.len(), 1);
    assert!(matches!(notes[0].class(), note::NoteClass::Footnote));
    assert_eq!(notes[0].body(), "note text");
    let annotations = body.annotations().unwrap();
    assert_eq!(annotations[0].creator(), Some("Ada"));
    assert_eq!(annotations[0].text(), "review text");
    let changes = body.changes().unwrap();
    assert_eq!(changes.len(), 3);
    assert!(matches!(changes[0].kind(), change::Kind::Insertion));
    assert_eq!(changes[0].text(), "");
    assert_eq!(changes[0].info().unwrap().creator(), "Ada");
    assert_eq!(changes[0].info().unwrap().date(), "2026-09-02T03:04:05Z");
    assert_eq!(
        changes[0].info().unwrap().paragraphs()[0].text(),
        "inserted"
    );
    assert!(matches!(changes[1].kind(), change::Kind::Start));
    let indexes = body.indexes().unwrap();
    assert_eq!(indexes[0].name(), Some("Contents"));
    assert_eq!(indexes[0].body(), "cached entry");
    assert!(matches!(indexes[0].kind(), index::Kind::TableOfContents));
    let frames = body.frames().unwrap();
    assert_eq!(frames[0].name(), Some("Box"));
    assert!(matches!(frames[0].kind(), frame::Kind::TextBox));
    let rubies = body.rubies().unwrap();
    assert_eq!(rubies[0].base(), "漢");
    assert_eq!(rubies[0].text(), "kan");
    let tables = body.tables().unwrap();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].declared_column_count(), 2);
    assert_eq!(tables[0].column_count(), 3);
    assert_eq!(tables[0].rows()[0].repeat_count(), 3);
    assert_eq!(tables[0].rows()[0].cells()[0].repeat_count(), 2);
    assert!(matches!(
        tables[0].rows()[0].cells()[1].kind(),
        table::CellKind::Covered
    ));
}

#[test]
fn office_forms_may_precede_tracking_prelude() {
    let content = CONTENT.replace(
        r#"<office:body><office:text><text:tracked-changes>"#,
        r#"<office:body><office:text><office:forms/><text:tracked-changes>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();
    assert!(body.forms().is_empty());
    assert!(body.change_tracking().unwrap().is_some());
    assert_eq!(body.changes().unwrap().len(), 3);
}

#[test]
fn structure_projection_ignores_same_named_foreign_content() {
    let content = CONTENT
        .replace(
            r#"<office:body>"#,
            r#"<office:styles><table:table table:name="styles-spoof"/></office:styles><office:body>"#,
        )
        .replace(
            r#"</text:tracked-changes>"#,
            r#"</text:tracked-changes><foreign:wrapper><table:table table:name="foreign-spoof"/></foreign:wrapper>"#,
        );
    let content = content.replace(
        r#"<text:p>section text</text:p>"#,
        r#"<text:p>section text</text:p><foreign:wrapper><table:table table:name="nested-foreign-spoof"/></foreign:wrapper>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    assert_eq!(template.text_body().unwrap().tables().unwrap().len(), 1);
}

#[test]
fn nested_sections_and_tables_remain_projected_without_expansion() {
    let content = CONTENT.replace(
        r#"<text:section text:name="Intro" text:protected="true"><text:p>section text</text:p></text:section>"#,
        r#"<text:section text:name="Intro" text:protected="true"><text:p>section text</text:p><text:section text:name="Nested"><table:table table:name="NestedTable"><table:table-row><table:table-cell><text:p>nested cell</text:p></table:table-cell></table:table-row></table:table></text:section></text:section>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();
    let sections = body.sections().unwrap();
    assert_eq!(sections.len(), 2);
    assert_eq!(sections[1].depth(), 2);
    let tables = body.tables().unwrap();
    assert_eq!(tables.len(), 2);
    assert!(
        tables
            .iter()
            .any(|table| table.name() == Some("NestedTable"))
    );
}

#[test]
fn grouped_table_rows_and_columns_are_projected() {
    let content = CONTENT.replace(
        r#"</office:text>"#,
        r#"<table:table table:name="Grouped"><table:table-column-group><table:table-column table:number-columns-repeated="4"/></table:table-column-group><table:table-row-group><table:table-row><table:table-cell><text:p>grouped</text:p></table:table-cell></table:table-row></table:table-row-group></table:table></office:text>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();
    let tables = body.tables().unwrap();
    let grouped = tables
        .iter()
        .find(|table| table.name() == Some("Grouped"))
        .unwrap();
    assert_eq!(grouped.columns().len(), 1);
    assert_eq!(grouped.declared_column_count(), 4);
    assert_eq!(grouped.column_count(), 4);
    assert_eq!(grouped.rows().len(), 1);
}

#[test]
fn structure_projection_is_deferred_until_typed_body_access() {
    let content = CONTENT.replace(
        r#"table:number-columns-repeated="2""#,
        r#"table:number-columns-repeated="bogus""#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();
    assert!(body.tables().is_err());
    assert!(body.tables().is_err());
}

#[test]
fn huge_table_repeats_stay_compact_and_checked() {
    let content = CONTENT
        .replacen(
            r#"table:number-columns-repeated="2""#,
            r#"table:number-columns-repeated="1000000""#,
            1,
        )
        .replace(
            r#"table:number-rows-repeated="3""#,
            r#"table:number-rows-repeated="1000000""#,
        );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let tables = template.text_body().unwrap().tables().unwrap().to_vec();
    let table = tables
        .iter()
        .find(|table| table.name() == Some("Data"))
        .unwrap();
    assert_eq!(table.columns().len(), 1);
    assert_eq!(table.columns()[0].repeat_count(), 1_000_000);
    assert_eq!(table.rows().len(), 1);
    assert_eq!(table.rows()[0].repeat_count(), 1_000_000);

    let overflowing = CONTENT.replace(
        r#"<table:table table:name="Data">"#,
        r#"<table:table table:name="Data"><table:table-column table:number-columns-repeated="1000000"/><table:table-column table:number-columns-repeated="2"/>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(overflowing).build().unwrap()).unwrap();
    assert!(template.text_body().unwrap().tables().is_err());
}

#[test]
fn change_markers_are_inert_without_tracked_changes_and_text_change_is_projected() {
    let without_tracking = CONTENT
        .replace(
            r#"<text:tracked-changes><text:changed-region xml:id="r1" text:id="r1"><text:insertion><office:change-info><dc:creator>Ada</dc:creator><dc:date>2026-09-02T03:04:05Z</dc:date><text:p>inserted</text:p></office:change-info></text:insertion></text:changed-region></text:tracked-changes>"#,
            "",
        );
    let template = Template::from_bytes(
        Builder::new()
            .content_xml(without_tracking)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert!(template.text_body().unwrap().changes().unwrap().is_empty());

    let with_point_marker = CONTENT.replace(
        r#"<text:change-end text:change-id="r1"/>"#,
        r#"<text:change text:change-id="r1"/>"#,
    );
    let template = Template::from_bytes(
        Builder::new()
            .content_xml(with_point_marker)
            .build()
            .unwrap(),
    )
    .unwrap();
    let body = template.text_body().unwrap();
    let changes = body.changes().unwrap();
    assert_eq!(changes.len(), 3);
    assert!(matches!(changes[2].kind(), change::Kind::Other(value) if value == "change"));
}

#[test]
fn foreign_wrappers_keep_typed_descendants_inert() {
    let content = CONTENT.replace(
        r#"</text:tracked-changes>"#,
        r#"</text:tracked-changes><foreign:wrapper><text:section text:name="foreign-section"><text:p>foreign</text:p></text:section><text:change text:change-id="foreign"/></foreign:wrapper>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();
    assert_eq!(body.sections().unwrap().len(), 1);
    assert_eq!(body.changes().unwrap().len(), 3);
}

#[test]
fn additional_index_and_frame_families_remain_inert_projections() {
    let content = CONTENT.replace(
        r#"</office:text>"#,
        r#"<text:illustration-index text:name="ExtraIllustrations"><text:illustration-index-source/><text:index-body/></text:illustration-index><text:alphabetical-index text:name="ExtraAlphabetical"><text:alphabetical-index-source/><text:index-body/></text:alphabetical-index><text:bibliography text:name="ExtraBibliography"><text:bibliography-source/><text:index-body/></text:bibliography><draw:frame draw:name="Image"><draw:image xlink:href="Pictures/image"/></draw:frame><draw:frame draw:name="Object"><draw:object xlink:href="Object 1"/></draw:frame></office:text>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();
    let indexes = body.indexes().unwrap();
    assert_eq!(indexes.len(), 4);
    assert!(
        indexes
            .iter()
            .any(|index| matches!(index.kind(), index::Kind::Illustration))
    );
    assert!(
        indexes
            .iter()
            .any(|index| matches!(index.kind(), index::Kind::Alphabetical))
    );
    assert!(
        indexes
            .iter()
            .any(|index| matches!(index.kind(), index::Kind::Bibliography))
    );
    let frames = body.frames().unwrap();
    assert_eq!(frames.len(), 4);
    assert!(
        frames
            .iter()
            .any(|frame| matches!(frame.kind(), frame::Kind::Image))
    );
    assert!(
        frames
            .iter()
            .any(|frame| matches!(frame.kind(), frame::Kind::Object))
    );
}

#[test]
fn malformed_text_space_repeat_fails_as_a_bounded_projection() {
    let content = CONTENT.replace(
        r#"<text:p>cell</text:p>"#,
        r#"<text:p><text:s text:c="not-a-count"/></text:p>"#,
    );
    assert!(Builder::new().content_xml(content).build().is_err());
}
