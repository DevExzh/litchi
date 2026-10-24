//! Worksheet features of opened workbooks (record 0327): reads on producer
//! fixtures, and every edit family on a fresh and an opened workbook, read
//! back through this reader and the eager worksheet reader.

use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::Package;
use litchi_xls::sheet_features::edit::{
    Anchor, DrawingChange, Format, FormatRule, Hyperlink, PageBreak, Picture, Shape, ShapeGeometry,
    Style, Threshold, Validation,
};
use litchi_xls::sheet_features::{Area, DrawingKind, FeatureSet, RuleType};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(name: &str) -> Package {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls")
        .join(name);
    Package::from_bytes(&std::fs::read(path).unwrap()).unwrap()
}

/// The Workbook stream of a package: what the editor keeps exact (the
/// compound file around it is re-rendered).
fn stream(package: &Package) -> Vec<u8> {
    let bytes = package.to_bytes().unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    ole.open_stream(&["Workbook"]).unwrap()
}

fn fresh() -> Package {
    Package::new().unwrap()
}

fn signed() -> Package {
    let source = fresh().to_bytes().unwrap();
    let mut ole = OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default()).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer
        .create_stream(&["DigitalSignature"], b"opaque signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    Package::from_bytes(&output.into_inner()).unwrap()
}

/// The package reopened from its bytes (eager), so the result is checked by
/// a full parse.
fn reopen(package: &Package) -> Package {
    let bytes = package.to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    // The eager reader parses every worksheet.
    package.reader_ref().unwrap().xls_worksheet(0).unwrap();
    package
}

const PNG: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xf8, 0xcf, 0xc0, 0xf0,
    0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99, 0x3d, 0x1d, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
];

#[test]
fn reads_producer_fixtures() {
    let comments = fixture("SimpleWithComments.xls")
        .sheet_features(0, FeatureSet::NOTES)
        .unwrap();
    assert_eq!(comments.notes.len(), 3);
    let images = fixture("SimpleWithImages.xls")
        .sheet_features(0, FeatureSet::DRAWINGS)
        .unwrap();
    assert_eq!(images.drawings.len(), 4);
    assert!(
        images
            .drawings
            .iter()
            .all(|drawing| drawing.kind == DrawingKind::Picture)
    );
    assert!(
        images
            .drawings
            .iter()
            .all(|drawing| drawing.image.is_some())
    );
    let filter = fixture("SimpleWithAutofilter.xls")
        .sheet_features(0, FeatureSet::AUTO_FILTER | FeatureSet::DRAWINGS)
        .unwrap();
    let auto_filter = filter.auto_filter.unwrap();
    assert_eq!(auto_filter.columns, 2);
    assert_eq!(auto_filter.area, Some(Area::new(0, 0, 3, 1)));
    // The drop-downs are not pictures or shapes.
    assert!(filter.drawings.is_empty());
    let breaks = fixture("SimpleWithPageBreaks.xls")
        .sheet_features(0, FeatureSet::PAGE_BREAKS)
        .unwrap();
    assert!(!breaks.row_breaks.is_empty() || !breaks.column_breaks.is_empty());
    let formats = fixture("WithConditionalFormatting.xls")
        .sheet_features(0, FeatureSet::FORMATS)
        .unwrap();
    assert!(!formats.formats.is_empty());
    let links = fixture("WithHyperlink.xls")
        .sheet_features(0, FeatureSet::HYPERLINKS)
        .unwrap();
    assert!(!links.hyperlinks.is_empty());
}

#[test]
fn page_breaks_round_trip_exactly() {
    let source = fresh();
    let before = stream(&source);
    let mut edit = source.edit_sheet_features().unwrap();
    edit.set_row_breaks(0, &[PageBreak::new(9), PageBreak::new(4)])
        .unwrap();
    edit.set_column_breaks(0, &[PageBreak::new(3)]).unwrap();
    let added = reopen(&edit.commit().unwrap().into_package());
    let features = added.sheet_features(0, FeatureSet::PAGE_BREAKS).unwrap();
    assert_eq!(
        features
            .row_breaks
            .iter()
            .map(|item| item.index)
            .collect::<Vec<_>>(),
        [4, 9]
    );
    assert_eq!(features.column_breaks[0].index, 3);
    assert_eq!(features.column_breaks[0].max, 65_535);
    let mut edit = added.edit_sheet_features().unwrap();
    edit.set_row_breaks(0, &[]).unwrap();
    edit.set_column_breaks(0, &[]).unwrap();
    let removed = edit.commit().unwrap().into_package();
    assert_eq!(stream(&removed), before);
}

#[test]
fn hyperlinks_add_set_remove() {
    let source = fixture("SimpleWithComments.xls");
    let before = stream(&source);
    let mut edit = source.edit_sheet_features().unwrap();
    let index = edit
        .add_hyperlink(
            0,
            &Hyperlink {
                area: Area::cell(8, 0),
                url: Some("https://example.com/".into()),
                location: None,
                display: Some("Example".into()),
                tooltip: Some("Open".into()),
            },
        )
        .unwrap();
    assert_eq!(index, 0);
    edit.add_hyperlink(
        0,
        &Hyperlink {
            area: Area::cell(1, 1),
            url: None,
            location: Some("Sheet2!A1".into()),
            display: None,
            tooltip: None,
        },
    )
    .unwrap();
    let added = reopen(&edit.commit().unwrap().into_package());
    let links = added
        .sheet_features(0, FeatureSet::HYPERLINKS)
        .unwrap()
        .hyperlinks;
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].address(), Some("https://example.com/"));
    assert_eq!(links[0].tooltip(), Some("Open"));
    assert_eq!(links[0].display_name(), Some("Example"));
    assert_eq!(links[1].location(), Some("Sheet2!A1"));
    // The eager reader sees the same links.
    let eager = added
        .reader_ref()
        .unwrap()
        .xls_worksheet(0)
        .unwrap()
        .hyperlinks()
        .len();
    assert_eq!(eager, 2);
    let mut edit = added.edit_sheet_features().unwrap();
    edit.set_hyperlink(
        0,
        1,
        &Hyperlink {
            area: Area::cell(2, 2),
            url: Some("mailto:a@example.com".into()),
            location: None,
            display: None,
            tooltip: None,
        },
    )
    .unwrap();
    let set = edit.commit().unwrap().into_package();
    let links = set
        .sheet_features(0, FeatureSet::HYPERLINKS)
        .unwrap()
        .hyperlinks;
    assert_eq!(links[1].address(), Some("mailto:a@example.com"));
    let mut edit = set.edit_sheet_features().unwrap();
    edit.remove_hyperlink(0, 1).unwrap();
    edit.remove_hyperlink(0, 0).unwrap();
    assert_eq!(stream(&edit.commit().unwrap().into_package()), before);
}

fn number(value: f64) -> Vec<u8> {
    let mut out = vec![0x1f];
    out.extend_from_slice(&value.to_le_bytes());
    out
}

#[test]
fn validations_add_set_remove() {
    let source = fresh();
    let before = stream(&source);
    let spec = Validation {
        ranges: vec![Area::new(0, 0, 9, 0)],
        kind: 1,
        operator: 0,
        error_style: 0,
        formula1: number(1.0),
        formula2: number(10.0),
        list: false,
        allow_blank: true,
        suppress_dropdown: false,
        show_input: true,
        show_error: true,
        prompt_title: Some("Count".into()),
        prompt: Some("1 to 10".into()),
        error_title: None,
        error: None,
    };
    let mut edit = source.edit_sheet_features().unwrap();
    assert_eq!(edit.add_validation(0, &spec).unwrap(), 0);
    let mut list = spec.clone();
    list.ranges = vec![Area::new(0, 1, 9, 1)];
    list.kind = 3;
    list.list = true;
    list.formula1 = {
        let mut out = vec![0x17, 3, 0];
        out.extend_from_slice(b"a\0b");
        out
    };
    list.formula2 = Vec::new();
    assert_eq!(edit.add_validation(0, &list).unwrap(), 1);
    let added = reopen(&edit.commit().unwrap().into_package());
    let validations = added
        .sheet_features(0, FeatureSet::VALIDATIONS)
        .unwrap()
        .validations;
    assert_eq!(validations.len(), 2);
    assert_eq!(validations[0].formula1.as_deref(), Some("1"));
    assert_eq!(validations[0].formula2.as_deref(), Some("10"));
    assert_eq!(validations[0].prompt.as_deref(), Some("1 to 10"));
    assert_eq!(validations[1].formula1.as_deref(), Some("\"a,b\""));
    let mut edit = added.edit_sheet_features().unwrap();
    edit.remove_validation(0, 1).unwrap();
    edit.remove_validation(0, 0).unwrap();
    assert_eq!(stream(&edit.commit().unwrap().into_package()), before);
}

#[test]
fn conditional_formats_of_every_kind() {
    let source = fresh();
    let before = stream(&source);
    let range = vec![Area::new(0, 0, 9, 0)];
    let style = Style {
        fill: Some(10),
        bold: Some(true),
        ..Style::default()
    };
    let threshold = |kind| Threshold {
        kind,
        number: None,
        formula: None,
    };
    let rules = vec![
        (
            FormatRule::CellIs {
                operator: 5,
                formulas: [number(5.0), Vec::new()],
            },
            RuleType::CellIs,
        ),
        (FormatRule::Expression(number(1.0)), RuleType::Expression),
        (
            FormatRule::TopN {
                top: true,
                percent: false,
                rank: 3,
            },
            RuleType::TopN,
        ),
        (
            FormatRule::ColorScale(vec![(threshold(2), 0xf8_696b), (threshold(3), 0x63_be7b)]),
            RuleType::ColorScale,
        ),
        (
            FormatRule::DataBar {
                color: 0x63_8ec6,
                min: threshold(2),
                max: threshold(3),
                min_length: 0,
                max_length: 100,
                show_value: true,
            },
            RuleType::DataBar,
        ),
        (
            FormatRule::IconSet {
                set: 0,
                thresholds: vec![
                    Threshold {
                        kind: 4,
                        number: Some(0.0),
                        formula: None,
                    },
                    Threshold {
                        kind: 4,
                        number: Some(33.0),
                        formula: None,
                    },
                    Threshold {
                        kind: 4,
                        number: Some(67.0),
                        formula: None,
                    },
                ],
                reverse: false,
                show_value: true,
            },
            RuleType::IconSet,
        ),
        (
            FormatRule::Template {
                template: 0x1b,
                params: [0; 16],
                formula: number(1.0),
            },
            RuleType::Expression,
        ),
    ];
    let mut edit = source.edit_sheet_features().unwrap();
    for (rule, _) in &rules {
        edit.add_format(
            0,
            &Format {
                ranges: range.clone(),
                rule: rule.clone(),
                style: style.clone(),
                stop_if_true: false,
            },
        )
        .unwrap();
    }
    let added = reopen(&edit.commit().unwrap().into_package());
    let formats = added
        .sheet_features(0, FeatureSet::FORMATS)
        .unwrap()
        .formats;
    assert_eq!(formats.len(), rules.len());
    for (format, (_, kind)) in formats.iter().zip(&rules) {
        assert_eq!(format.rule.as_ref().unwrap().rule_type, *kind);
        assert_eq!(format.ranges, range);
    }
    let first = formats[0].rule.as_ref().unwrap();
    assert_eq!(first.operator, Some(5));
    assert_eq!(first.formulas, ["5"]);
    assert_eq!(first.style.bold, Some(true));
    assert!(first.style.fill.is_some());
    assert_eq!(formats[2].rule.as_ref().unwrap().parameter, 3);
    assert_eq!(
        formats[3]
            .rule
            .as_ref()
            .unwrap()
            .color_scale
            .as_ref()
            .unwrap()
            .points[1]
            .1,
        Some(0x63_be7b)
    );
    assert_eq!(formats[6].rule.as_ref().unwrap().template, 0x1b);
    // Ranges only.
    let mut edit = added.edit_sheet_features().unwrap();
    edit.set_format_ranges(0, 0, &[Area::new(1, 1, 4, 2)])
        .unwrap();
    let set = edit.commit().unwrap().into_package();
    let formats = set.sheet_features(0, FeatureSet::FORMATS).unwrap().formats;
    assert_eq!(formats[0].ranges, [Area::new(1, 1, 4, 2)]);
    let mut edit = set.edit_sheet_features().unwrap();
    for index in (0..rules.len()).rev() {
        edit.remove_format(0, index).unwrap();
    }
    assert_eq!(stream(&edit.commit().unwrap().into_package()), before);
}

#[test]
fn notes_add_set_remove() {
    let sources = [
        fresh(),
        fixture("SimpleWithComments.xls"),
        fixture("SimpleWithImages.xls"),
    ];
    for (position, source) in sources.into_iter().enumerate() {
        let before = stream(&source);
        let notes_before = source
            .sheet_features(0, FeatureSet::NOTES)
            .unwrap()
            .notes
            .len();
        let mut edit = source.edit_sheet_features().unwrap();
        edit.add_note(0, 20, 3, "Ann", "Check this")
            .unwrap()
            .unwrap();
        assert!(edit.add_note(0, 20, 3, "Ann", "again").unwrap().is_none());
        let added = reopen(&edit.commit().unwrap().into_package());
        let notes = added.sheet_features(0, FeatureSet::NOTES).unwrap().notes;
        assert_eq!(notes.len(), notes_before + 1);
        let note = notes
            .iter()
            .find(|note| (note.row, note.column) == (20, 3))
            .unwrap();
        assert_eq!(note.author, "Ann");
        assert_eq!(note.text, "Check this");
        let eager = added
            .reader_ref()
            .unwrap()
            .xls_worksheet(0)
            .unwrap()
            .comments()
            .len();
        assert_eq!(eager, notes_before + 1);
        let mut edit = added.edit_sheet_features().unwrap();
        edit.set_note(0, (20, 3), (21, 4), "Bob", "Done").unwrap();
        let set = reopen(&edit.commit().unwrap().into_package());
        let notes = set.sheet_features(0, FeatureSet::NOTES).unwrap().notes;
        let note = notes
            .iter()
            .find(|note| (note.row, note.column) == (21, 4))
            .unwrap();
        assert_eq!((note.author.as_str(), note.text.as_str()), ("Bob", "Done"));
        let mut edit = set.edit_sheet_features().unwrap();
        edit.remove_note(0, 21, 4).unwrap();
        let removed = reopen(&edit.commit().unwrap().into_package());
        assert_eq!(
            removed
                .sheet_features(0, FeatureSet::NOTES)
                .unwrap()
                .notes
                .len(),
            notes_before
        );
        // A workbook without a drawing group gets its stream back exactly;
        // an existing group keeps its raised shape-id counters.
        if position == 0 {
            assert_eq!(stream(&removed), before);
        }
    }
}

fn anchor() -> Anchor {
    Anchor {
        first_row: 1,
        first_column: 1,
        last_row: 5,
        last_column: 4,
        behavior: litchi_xls::AnchorBehavior::MoveAndResize,
    }
}

#[test]
fn pictures_and_shapes() {
    for source in [
        fresh(),
        fixture("SimpleWithComments.xls"),
        fixture("SimpleWithImages.xls"),
    ] {
        let before = stream(&source);
        let drawings_before = source
            .sheet_features(0, FeatureSet::DRAWINGS)
            .unwrap()
            .drawings;
        let mut edit = source.edit_sheet_features().unwrap();
        edit.add_picture(
            0,
            &Picture {
                data: PNG.to_vec(),
                anchor: anchor(),
                name: Some("Logo".into()),
                description: Some("A dot".into()),
                rotation: 0.0,
                flip_h: false,
                flip_v: false,
            },
        )
        .unwrap();
        edit.add_shape(
            0,
            &Shape {
                geometry: ShapeGeometry::Rectangle,
                anchor: anchor(),
                text: Some("Box".into()),
                fill: Some(Some(0xff_0000)),
                line: Some(Some((0x00_00ff, 12_700))),
                name: Some("Shape 1".into()),
                description: None,
                rotation: 0.0,
                flip_h: false,
                flip_v: false,
            },
        )
        .unwrap();
        let added = reopen(&edit.commit().unwrap().into_package());
        let drawings = added
            .sheet_features(0, FeatureSet::DRAWINGS)
            .unwrap()
            .drawings;
        assert_eq!(drawings.len(), drawings_before.len() + 2);
        let picture = drawings
            .iter()
            .find(|drawing| drawing.name.as_deref() == Some("Logo"))
            .unwrap();
        assert_eq!(picture.kind, DrawingKind::Picture);
        assert_eq!(picture.description.as_deref(), Some("A dot"));
        assert_eq!(picture.image.unwrap().0, "image/png");
        let shape = drawings
            .iter()
            .find(|drawing| drawing.kind == DrawingKind::Shape)
            .unwrap();
        assert_eq!(shape.text.as_deref(), Some("Box"));
        assert_eq!(
            shape.fill,
            Some(litchi_xls::sheet_features::Fill::Solid(0xff_0000))
        );
        let shapes_before = drawings_before
            .iter()
            .filter(|drawing| drawing.kind == DrawingKind::Shape)
            .count();
        let pictures_before = drawings_before.len() - shapes_before;
        // Set: anchor, name, rotation and flips.
        let mut edit = added.edit_sheet_features().unwrap();
        edit.set_drawing(
            0,
            true,
            pictures_before,
            &DrawingChange {
                anchor: Some(Anchor {
                    first_row: 3,
                    ..anchor()
                }),
                name: Some(Some("Renamed picture".into())),
                description: Some(None),
                rotation: Some(90.0),
                flip: Some((true, false)),
            },
        )
        .unwrap();
        let set = reopen(&edit.commit().unwrap().into_package());
        let drawings = set
            .sheet_features(0, FeatureSet::DRAWINGS)
            .unwrap()
            .drawings;
        let picture = drawings
            .iter()
            .find(|drawing| drawing.name.as_deref() == Some("Renamed picture"))
            .unwrap();
        assert_eq!(picture.description, None);
        assert_eq!(picture.rotation, 90.0);
        assert!(picture.flip_h);
        assert_eq!(picture.anchor.unwrap().top_left().row(), 3);
        let mut edit = set.edit_sheet_features().unwrap();
        edit.remove_drawing(0, false, shapes_before).unwrap();
        edit.remove_drawing(0, true, pictures_before).unwrap();
        let removed = reopen(&edit.commit().unwrap().into_package());
        assert_eq!(
            removed
                .sheet_features(0, FeatureSet::DRAWINGS)
                .unwrap()
                .drawings,
            drawings_before
        );
        if drawings_before.is_empty()
            && source
                .sheet_features(0, FeatureSet::NOTES)
                .unwrap()
                .notes
                .is_empty()
        {
            assert_eq!(stream(&removed), before);
        }
    }
}

#[test]
fn auto_filter_add_set_remove() {
    let source = fresh();
    let mut edit = source.edit_sheet_features().unwrap();
    edit.set_auto_filter(0, Some(Area::new(0, 0, 9, 2)))
        .unwrap();
    let added = reopen(&edit.commit().unwrap().into_package());
    let filter = added
        .sheet_features(0, FeatureSet::AUTO_FILTER)
        .unwrap()
        .auto_filter
        .unwrap();
    assert_eq!(filter.columns, 3);
    assert_eq!(filter.area, Some(Area::new(0, 0, 9, 2)));
    let eager = added.reader_ref().unwrap();
    assert!(eager.filter_database_name(0).is_some());
    let mut edit = added.edit_sheet_features().unwrap();
    edit.set_auto_filter(0, Some(Area::new(1, 1, 5, 1)))
        .unwrap();
    let set = reopen(&edit.commit().unwrap().into_package());
    let filter = set
        .sheet_features(0, FeatureSet::AUTO_FILTER)
        .unwrap()
        .auto_filter
        .unwrap();
    assert_eq!(
        (filter.columns, filter.area),
        (1, Some(Area::new(1, 1, 5, 1)))
    );
    let mut edit = set.edit_sheet_features().unwrap();
    edit.set_auto_filter(0, None).unwrap();
    let removed = reopen(&edit.commit().unwrap().into_package());
    assert!(
        removed
            .sheet_features(0, FeatureSet::AUTO_FILTER)
            .unwrap()
            .auto_filter
            .is_none()
    );
    // The producer fixture's filter is replaced and removed too.
    let producer = fixture("SimpleWithAutofilter.xls");
    let mut edit = producer.edit_sheet_features().unwrap();
    edit.set_auto_filter(0, None).unwrap();
    let removed = reopen(&edit.commit().unwrap().into_package());
    assert!(
        removed
            .sheet_features(0, FeatureSet::AUTO_FILTER)
            .unwrap()
            .auto_filter
            .is_none()
    );
}

#[test]
fn an_exact_no_op_keeps_the_package() {
    let source = fixture("SimpleWithComments.xls");
    let edit = source.edit_sheet_features().unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn signed_feature_reads_and_noops_preserve_bytes_but_changes_refuse() {
    let source = signed();
    let bytes = source.to_bytes().unwrap();
    let mut edit = source.edit_sheet_features().unwrap();
    let _ = edit.features(0, FeatureSet::PAGE_BREAKS).unwrap();
    let noop = edit.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.package().to_bytes().unwrap(), bytes);
    assert_eq!(
        source
            .apply_sheet_features(noop.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );

    let mut changed = source.edit_sheet_features().unwrap();
    changed.set_row_breaks(0, &[PageBreak::new(5)]).unwrap();
    assert!(matches!(
        changed.commit(),
        Err(litchi_xls::sheet_features::edit::Error::Signed)
    ));
}

#[test]
fn patches_apply_and_invert() {
    let source = fresh();
    let mut edit = source.edit_sheet_features().unwrap();
    edit.set_row_breaks(0, &[PageBreak::new(5)]).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let replayed = source.apply_sheet_features(&patch).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        after.to_bytes().unwrap()
    );
    let undone = after.apply_sheet_features(&patch.inverse()).unwrap();
    assert_eq!(
        undone.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
    assert!(matches!(
        after.apply_sheet_features(&patch),
        Err(litchi_xls::sheet_features::edit::Error::PatchConflict)
    ));
}

/// Writes one workbook with every feature to `LITCHI_XLS_FEATURES_OUT`, for
/// an external reader to check (record 0327 used LibreOffice).
#[test]
#[ignore = "writes an evidence workbook"]
fn write_every_feature() {
    let Ok(out) = std::env::var("LITCHI_XLS_FEATURES_OUT") else {
        return;
    };
    let source =
        std::env::var("LITCHI_XLS_FEATURES_SOURCE").map_or_else(|_| fresh(), |name| fixture(&name));
    let mut edit = source.edit_sheet_features().unwrap();
    edit.add_hyperlink(
        0,
        &Hyperlink {
            area: Area::cell(0, 0),
            url: Some("https://example.org/".into()),
            location: None,
            display: None,
            tooltip: None,
        },
    )
    .unwrap();
    edit.set_row_breaks(0, &[PageBreak::new(9)]).unwrap();
    edit.set_column_breaks(0, &[PageBreak::new(3)]).unwrap();
    edit.add_hyperlink(
        0,
        &Hyperlink {
            area: Area::cell(8, 0),
            url: Some("https://example.com/".into()),
            location: None,
            display: Some("Example".into()),
            tooltip: Some("Open".into()),
        },
    )
    .unwrap();
    edit.add_validation(
        0,
        &Validation {
            ranges: vec![Area::new(0, 5, 9, 5)],
            kind: 1,
            operator: 0,
            error_style: 0,
            formula1: number(1.0),
            formula2: number(10.0),
            list: false,
            allow_blank: true,
            suppress_dropdown: false,
            show_input: true,
            show_error: true,
            prompt_title: Some("Count".into()),
            prompt: Some("1 to 10".into()),
            error_title: None,
            error: None,
        },
    )
    .unwrap();
    edit.add_format(
        0,
        &Format {
            ranges: vec![Area::new(0, 0, 9, 0)],
            rule: FormatRule::CellIs {
                operator: 5,
                formulas: [number(5.0), Vec::new()],
            },
            style: Style {
                fill: Some(10),
                bold: Some(true),
                ..Style::default()
            },
            stop_if_true: false,
        },
    )
    .unwrap();
    edit.add_format(
        0,
        &Format {
            ranges: vec![Area::new(0, 1, 9, 1)],
            rule: FormatRule::DataBar {
                color: 0x63_8ec6,
                min: Threshold {
                    kind: 2,
                    number: None,
                    formula: None,
                },
                max: Threshold {
                    kind: 3,
                    number: None,
                    formula: None,
                },
                min_length: 0,
                max_length: 100,
                show_value: true,
            },
            style: Style::default(),
            stop_if_true: false,
        },
    )
    .unwrap();
    edit.add_note(0, 2, 2, "Ann", "Check this").unwrap();
    edit.add_picture(
        0,
        &Picture {
            data: PNG.to_vec(),
            anchor: anchor(),
            name: Some("Logo".into()),
            description: Some("A dot".into()),
            rotation: 0.0,
            flip_h: false,
            flip_v: false,
        },
    )
    .unwrap();
    edit.add_shape(
        0,
        &Shape {
            geometry: ShapeGeometry::Ellipse,
            anchor: Anchor {
                first_row: 10,
                first_column: 1,
                last_row: 14,
                last_column: 4,
                behavior: litchi_xls::AnchorBehavior::MoveAndResize,
            },
            text: Some("Box".into()),
            fill: Some(Some(0xff_0000)),
            line: Some(Some((0x00_00ff, 12_700))),
            name: Some("Shape 1".into()),
            description: None,
            rotation: 0.0,
            flip_h: false,
            flip_v: false,
        },
    )
    .unwrap();
    edit.set_auto_filter(0, Some(Area::new(0, 0, 9, 2)))
        .unwrap();
    let package = reopen(&edit.commit().unwrap().into_package());
    std::fs::write(out, package.to_bytes().unwrap()).unwrap();
}

#[test]
fn formulas_compile_with_the_built_in_functions_only() {
    use litchi_xls::sheet_features::edit::compile_formula;
    // NOT and ISERROR take one argument (PtgFunc), SEARCH two or three
    // (PtgFuncVar).
    let tokens = compile_formula("NOT(ISERROR(SEARCH(\"x\",A1)))").unwrap();
    assert_eq!(tokens.iter().filter(|byte| **byte == 0x41).count(), 2);
    assert!(compile_formula("AND(A1>=DATE(YEAR(TODAY()),1,1),A1<5)").is_ok());
    for refused in [
        "CALL(\"x\")",
        "REGISTER.ID(\"x\")",
        "ISODD(A1)",
        "NOT(1,2)",
        "Sheet2!A1",
    ] {
        assert!(compile_formula(refused).is_err(), "{refused}");
    }
}
