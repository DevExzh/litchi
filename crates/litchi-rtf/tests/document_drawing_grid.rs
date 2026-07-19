use litchi_rtf::{
    DocumentDrawingGrid, DrawingGridLineInterval, DrawingGridSpacing, RtfDocument, RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_a_complete_producer_style_grid() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\dgmargin\dghspace180\dgvspace180\dghorigin1440"#,
        r#"\dgvorigin1440\dghshow1\dgvshow1 Body}"#,
    ))
    .unwrap();
    assert_eq!(
        *document.drawing_grid(),
        DocumentDrawingGrid {
            follows_margins: true,
            snap_to_grid: false,
            horizontal_spacing: DrawingGridSpacing::new(180),
            vertical_spacing: DrawingGridSpacing::new(180),
            horizontal_origin_twips: Some(1440),
            vertical_origin_twips: Some(1440),
            horizontal_line_interval: DrawingGridLineInterval::new(1),
            vertical_line_interval: DrawingGridLineInterval::new(1),
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_is_distinct_from_all_effective_defaults() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    let grid = document.drawing_grid();
    assert!(grid.is_empty());
    assert_eq!(grid.effective_horizontal_spacing().get(), 120);
    assert_eq!(grid.effective_vertical_spacing().get(), 120);
    assert_eq!(grid.effective_horizontal_origin_twips(), 1701);
    assert_eq!(grid.effective_vertical_origin_twips(), 1984);
    assert_eq!(grid.effective_horizontal_line_interval().get(), 3);
    assert_eq!(grid.effective_vertical_line_interval().get(), 0);
    let serialized = String::from_utf8(write(&document)).unwrap();
    for control in [
        "dghspace",
        "dgvspace",
        "dghorigin",
        "dgvorigin",
        "dghshow",
        "dgvshow",
    ] {
        assert!(!serialized.contains(control));
    }
}

#[test]
fn typed_api_round_trips_boundaries_flags_and_canonical_order() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_drawing_grid(DocumentDrawingGrid {
        follows_margins: true,
        snap_to_grid: true,
        horizontal_spacing: DrawingGridSpacing::new(0),
        vertical_spacing: DrawingGridSpacing::new(32767),
        horizontal_origin_twips: Some(-32768),
        vertical_origin_twips: Some(32767),
        horizontal_line_interval: DrawingGridLineInterval::new(0),
        vertical_line_interval: DrawingGridLineInterval::new(32767),
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    let controls = [
        "\\dghspace0",
        "\\dgvspace32767",
        "\\dghorigin-32768",
        "\\dgvorigin32767",
        "\\dghshow0",
        "\\dgvshow32767",
        "\\dgsnap",
        "\\dgmargin",
    ];
    for pair in controls.windows(2) {
        assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
    }
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.drawing_grid(), document.drawing_grid());
    assert_eq!(reparsed.text(), "Body");

    document.clear_drawing_grid();
    assert!(document.drawing_grid().is_empty());
    assert_eq!(document.text(), "Body");
    assert!(DrawingGridSpacing::new(32768).is_none());
    assert!(DrawingGridLineInterval::new(32768).is_none());
}

#[test]
fn coexists_with_rendering_and_processing_properties_without_side_effects() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\dgvshow2\jcompress\grfdocevents0\dgmargin"#,
        r#"\dghspace120\horzdoc\dgsnap\fracwidth Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.drawing_grid(), document.drawing_grid());
    assert_eq!(reparsed.rendering_settings(), document.rendering_settings());
    assert_eq!(
        reparsed.processing_settings(),
        document.processing_settings()
    );
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_parameters_missing_values_ranges_overflow_and_duplicates() {
    for source in [
        r#"{\rtf1\dgmargin0 Body}"#,
        r#"{\rtf1\dgsnap1 Body}"#,
        r#"{\rtf1\dghspace Body}"#,
        r#"{\rtf1\dghspace-1 Body}"#,
        r#"{\rtf1\dghspace32768 Body}"#,
        r#"{\rtf1\dgvspace Body}"#,
        r#"{\rtf1\dgvspace-1 Body}"#,
        r#"{\rtf1\dgvspace32768 Body}"#,
        r#"{\rtf1\dghorigin Body}"#,
        r#"{\rtf1\dghorigin-32769 Body}"#,
        r#"{\rtf1\dghorigin32768 Body}"#,
        r#"{\rtf1\dgvorigin Body}"#,
        r#"{\rtf1\dgvorigin-32769 Body}"#,
        r#"{\rtf1\dgvorigin32768 Body}"#,
        r#"{\rtf1\dghshow Body}"#,
        r#"{\rtf1\dghshow-1 Body}"#,
        r#"{\rtf1\dghshow32768 Body}"#,
        r#"{\rtf1\dgvshow Body}"#,
        r#"{\rtf1\dgvshow-1 Body}"#,
        r#"{\rtf1\dgvshow32768 Body}"#,
        r#"{\rtf1\dghspace99999999999 Body}"#,
        r#"{\rtf1\dgvorigin-99999999999 Body}"#,
        r#"{\rtf1\dgmargin\dgmargin Body}"#,
        r#"{\rtf1\dgsnap\dgsnap Body}"#,
        r#"{\rtf1\dghspace1\dghspace2 Body}"#,
        r#"{\rtf1\dgvspace1\dgvspace2 Body}"#,
        r#"{\rtf1\dghorigin1\dghorigin2 Body}"#,
        r#"{\rtf1\dgvorigin1\dgvorigin2 Body}"#,
        r#"{\rtf1\dghshow1\dghshow2 Body}"#,
        r#"{\rtf1\dgvshow1\dgvshow2 Body}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn rejects_every_starred_grouped_and_late_grid_control() {
    for control in [
        r#"\dgmargin"#,
        r#"\dgsnap"#,
        r#"\dghspace120"#,
        r#"\dgvspace120"#,
        r#"\dghorigin1701"#,
        r#"\dgvorigin1984"#,
        r#"\dghshow3"#,
        r#"\dgvshow0"#,
    ] {
        for source in [
            format!(r#"{{\rtf1{{\*{control}}}Body}}"#),
            format!(r#"{{\rtf1{{{control}}}Body}}"#),
            format!(r#"{{\rtf1 Body{control}}}"#),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
}
