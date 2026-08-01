use litchi_rtf::{BorderStyle, RtfDocument, RtfWriter, TableDistanceUnit, TablePreferredWidthUnit};

#[test]
fn parses_row_cell_defaults_and_resets_per_row() {
    let source = r#"{\rtf1\trowd\tsbrdrt\brdrs\brdrw20\brdrcf2\tsbrdrh\brdrdot\brdrw10\tscellpaddl100\tscellpaddfl3\tscellspct20\tscellspcft3\tscellwidth2500\tscellwidthfts2\cellx2500\intbl A\cell\row\trowd\tscellpaddr40\tscellpaddfr3\cellx2500\intbl B\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();

    let defaults = rows[0].cell_defaults();
    let top = defaults.borders.top.unwrap();
    assert_eq!(top.style, BorderStyle::Single);
    assert_eq!(top.width, 20);
    assert_eq!(top.color_ref, 2);
    let horizontal = defaults.borders.horizontal_inside.unwrap();
    assert_eq!(horizontal.style, BorderStyle::Dotted);
    assert_eq!(horizontal.width, 10);
    assert!(defaults.borders.vertical_inside.is_none());
    assert!(
        defaults
            .borders
            .diagonal_upper_left_to_lower_right
            .is_none()
    );
    assert_eq!(defaults.padding.left.value, Some(100));
    assert_eq!(defaults.padding.left.unit, Some(TableDistanceUnit::Twips));
    assert_eq!(defaults.spacing.top.value, Some(20));
    assert_eq!(defaults.spacing.top.unit, Some(TableDistanceUnit::Twips));
    let width = defaults.preferred_cell_width.unwrap();
    assert_eq!(width.unit(), TablePreferredWidthUnit::Percent);
    assert_eq!(width.value(), Some(2500));

    // A new row definition drops the previous defaults.
    let reset = rows[1].cell_defaults();
    assert!(reset.borders.is_empty());
    assert_eq!(reset.padding.right.value, Some(40));
    assert!(reset.preferred_cell_width.is_none());
}

#[test]
fn parses_default_cell_width_units() {
    for (source, unit, value) in [
        (
            r#"{\rtf1\trowd\tscellwidthfts1\cellx1000\intbl A\cell\row}"#,
            TablePreferredWidthUnit::Auto,
            None,
        ),
        (
            r#"{\rtf1\trowd\tscellwidth1440\tscellwidthfts3\cellx1440\intbl A\cell\row}"#,
            TablePreferredWidthUnit::Twips,
            Some(1440),
        ),
    ] {
        let document = RtfDocument::parse(source).unwrap();
        let width = document.tables()[0].rows()[0]
            .cell_defaults()
            .preferred_cell_width
            .unwrap();
        assert_eq!(width.unit(), unit, "misparsed {source}");
        assert_eq!(width.value(), value, "misparsed {source}");
    }
}

#[test]
fn writer_round_trips_deterministically() {
    let document = RtfDocument::parse(
        r#"{\rtf1\trowd\tsbrdrb\brdrs\brdrw30\brdrcf1\brsp4\tsbrdrv\brdrdash\brdrw10\tscellpaddl57\tscellpaddfl3\tscellspcb12\tscellspcfb3\tscellwidth1000\tscellwidthfts2\cellx2000\intbl Cell\cell\row}"#,
    )
    .unwrap();
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let first = String::from_utf8(first).unwrap();
    assert!(first.contains("\\tsbrdrb\\brdrs\\brdrw30\\brdrcf1\\brsp4"));
    assert!(first.contains("\\tsbrdrv\\brdrdash\\brdrw10"));
    assert!(first.contains("\\tscellpaddl57\\tscellpaddfl3"));
    assert!(first.contains("\\tscellspcb12\\tscellspcfb3"));
    assert!(first.contains("\\tscellwidthfts2\\tscellwidth1000"));
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(
        reparsed.tables()[0].rows()[0].cell_defaults(),
        document.tables()[0].rows()[0].cell_defaults()
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, String::from_utf8(second).unwrap());
}

#[test]
fn rejects_malformed_default_controls() {
    for source in [
        r#"{\rtf1\trowd\tscellpaddl X}"#,
        r#"{\rtf1\trowd\tscellpaddl-1 X}"#,
        r#"{\rtf1\trowd\tscellpaddl31681 X}"#,
        r#"{\rtf1\trowd\tscellpaddfl1 X}"#,
        r#"{\rtf1\trowd\tscellspcft9 X}"#,
        r#"{\rtf1\trowd\tscellwidth X}"#,
        r#"{\rtf1\trowd\tscellwidth2500\cellx1000\intbl A\cell\row}"#,
        r#"{\rtf1\trowd\tscellwidthfts2\cellx1000\intbl A\cell\row}"#,
        r#"{\rtf1\trowd\tscellwidthfts4 X}"#,
        r#"{\rtf1\trowd\tscellwidth100\tscellwidthfts2\tscellwidthfts3 X}"#,
        r#"{\rtf1\trowd\tscellwidth100\tscellwidth200\tscellwidthfts3 X}"#,
        r#"{\rtf1\trowd\tsbrdrt\brdrw20 X}"#,
        r#"{\rtf1\trowd\tsbrdrt\brdrs\brdrw80 X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
    assert!(
        RtfDocument::parse(r#"{\rtf1{\*\unknown\tscellpaddl-1\tscellwidthfts9 bad}Visible}"#)
            .is_ok()
    );
}
