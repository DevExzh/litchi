#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure assertions"
)]

use litchi_core::sheet::{CellValue, traits::WorkbookTrait};
use litchi_xlsb::cell_values::{
    AuthoredStyle, CellFormula, DrawingTransferRefusal, Reference, TransferLimits, Value,
    WorkbookHistory, WorkbookPatch,
};
use litchi_xlsb::chart::{Chart, ExternalDataPart};
use litchi_xlsb::named_ranges::Definition as NamedRange;
use litchi_xlsb::package::table::{Column, Range as TableRange, Table, Type as TableType};
use litchi_xlsb::package::{SharedString, SharedStringRun};
use litchi_xlsb::shapes::{
    Anchor as ShapeAnchor, CellMarker, EditAs, Emu, EmuExtent, EmuOffset, Object as ShapeObject,
    Preset,
};
use litchi_xlsb::styles::{Alignment, Fill, Font, HorizontalAlignment, VerticalAlignment};
use litchi_xlsb::writer::{ChartAnchor, Image, ImageFormat, MutableWorksheet, WorkbookWriter};
use litchi_xlsb::writer::{ConnectionEndSpec, ConnectionShapeSpec, GroupSpec, ShapeSpec};
use litchi_xlsb::{Package, Workbook};
use litchi_xlsb::{Parser, TableColumns, TableDataType, TableReference, TableRowType, Token};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn producer_workbook(name: &str, initial_value: Option<&str>) -> Workbook {
    let mut sheet = MutableWorksheet::new(name);
    if let Some(value) = initial_value {
        sheet.set_cell(0, 0, value);
    }
    let mut producer = WorkbookWriter::new();
    producer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    producer.save(&mut bytes).expect("producer save");
    Workbook::new(Cursor::new(bytes.into_inner())).expect("producer reopen")
}

fn shape_marker(column: u32, row: u32) -> CellMarker {
    CellMarker {
        column,
        row,
        column_offset: Emu(0),
        row_offset: Emu(0),
    }
}

fn shape_graph_workbook(name: &str) -> Workbook {
    let child_anchor = ShapeAnchor::TwoCell {
        from: shape_marker(0, 0),
        to: shape_marker(1, 1),
        edit_as: EditAs::TwoCell,
    };
    let group = GroupSpec::new(
        "Pair",
        ShapeAnchor::OneCell {
            from: shape_marker(2, 1),
            extent: EmuExtent {
                width: Emu(4_000_000),
                height: Emu(2_000_000),
            },
        },
    )
    .with_child(ShapeSpec::shape("Left", child_anchor, Preset::Rect, "L").into())
    .with_child(ShapeSpec::shape("Right", child_anchor, Preset::Ellipse, "R").into());
    let connection = ConnectionShapeSpec::new(
        "Bridge",
        ShapeAnchor::Absolute {
            position: EmuOffset {
                x: Emu(500_000),
                y: Emu(500_000),
            },
            extent: EmuExtent {
                width: Emu(1_000_000),
                height: Emu(1_000_000),
            },
        },
        Preset::StraightConnector1,
        ConnectionEndSpec {
            shape_name: "Left".to_string(),
            site: 1,
        },
        ConnectionEndSpec {
            shape_name: "Right".to_string(),
            site: 2,
        },
    );
    let mut sheet = MutableWorksheet::new(name);
    sheet
        .add_shape(ShapeSpec::shape(
            "Standalone",
            child_anchor,
            Preset::RoundRect,
            "S",
        ))
        .expect("standalone shape");
    sheet.add_group(group).expect("shape group");
    sheet.add_connection(connection).expect("shape connection");
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes).expect("shape graph save");
    Workbook::new(Cursor::new(bytes.into_inner())).expect("shape graph reopen")
}

fn collect_object_ids(object: &ShapeObject, output: &mut Vec<u32>) {
    match object {
        ShapeObject::Shape(shape) => output.extend(shape.non_visual.id),
        ShapeObject::ConnectionShape(connection) => output.extend(connection.non_visual.id),
        ShapeObject::Group(group) => {
            output.extend(group.non_visual.id);
            for child in &group.children {
                collect_object_ids(child, output);
            }
        },
        ShapeObject::OleObject(object) => output.extend(object.non_visual.id),
        ShapeObject::Unknown(_) => {},
    }
}

fn collect_object_names(object: &ShapeObject, output: &mut Vec<String>) {
    match object {
        ShapeObject::Shape(shape) => output.extend(shape.non_visual.name.clone()),
        ShapeObject::ConnectionShape(connection) => {
            output.extend(connection.non_visual.name.clone());
        },
        ShapeObject::Group(group) => {
            output.extend(group.non_visual.name.clone());
            for child in &group.children {
                collect_object_names(child, output);
            }
        },
        ShapeObject::OleObject(object) => output.extend(object.non_visual.name.clone()),
        ShapeObject::Unknown(_) => {},
    }
}

fn standard_drawing_xml(workbook: &Workbook) -> Vec<u8> {
    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).expect("workbook package save");
    let package = Package::from_slice(bytes.get_ref()).expect("package reopen");
    package
        .opc_package()
        .iter_parts()
        .find(|part| {
            part.content_type() == "application/vnd.openxmlformats-officedocument.drawing+xml"
        })
        .expect("standard drawing part")
        .blob()
        .to_vec()
}

fn formula_dependency_workbook(source_order: bool) -> Workbook {
    let mut data = MutableWorksheet::new("Data");
    data.set_cell(0, 0, "Region");
    data.set_cell(0, 1, "Amount");
    data.set_cell(1, 0, "North");
    data.set_cell(1, 1, 42.5);
    data.add_table(Table {
        id: if source_order { 3 } else { 9 },
        name: Some("SalesTable".to_string()),
        display_name: Some("SalesTable".to_string()),
        range: TableRange {
            first_row: 0,
            last_row: 1,
            first_column: 0,
            last_column: 1,
        },
        table_type: TableType::Range,
        header_row_count: 1,
        columns: vec![
            Column {
                id: 1,
                name: Some("Region".to_string()),
                ..Column::default()
            },
            Column {
                id: 2,
                name: Some("Amount".to_string()),
                ..Column::default()
            },
        ],
        ..Table::default()
    })
    .expect("valid structured table");
    let mut summary = MutableWorksheet::new("Summary");
    summary.set_cell(
        0,
        0,
        CellValue::Formula {
            formula: "Rate+Data!A1".to_string(),
            cached_value: Some(Box::new(CellValue::Float(1.0))),
            is_array: false,
            array_range: None,
        },
    );
    let mut producer = WorkbookWriter::new();
    if source_order {
        producer.add_worksheet(data);
        producer.add_worksheet(summary);
        producer.add_named_range(
            NamedRange::new("Rate".to_string(), None).with_formula(vec![0x1e, 1, 0]),
        );
        producer.add_named_range(
            NamedRange::new("Other".to_string(), None).with_formula(vec![0x1e, 2, 0]),
        );
    } else {
        producer.add_worksheet(summary);
        producer.add_worksheet(data);
        producer.add_named_range(
            NamedRange::new("Other".to_string(), None).with_formula(vec![0x1e, 2, 0]),
        );
        producer.add_named_range(
            NamedRange::new("Rate".to_string(), None).with_formula(vec![0x1e, 1, 0]),
        );
    }
    let mut bytes = Cursor::new(Vec::new());
    producer.save(&mut bytes).expect("dependency producer save");
    Workbook::new(Cursor::new(bytes.into_inner())).expect("dependency producer reopen")
}

#[test]
fn producer_reopen_transfers_sst_and_renames_on_a_durable_root() {
    let source = producer_workbook("Source", Some("shared transfer value"));
    let mut target = producer_workbook("Target", None);
    let destination = Reference::new(4, 3).expect("destination");

    let mut edit = target
        .edit_workbook_structure()
        .expect("detached workbook edit");
    edit.rename_sheet(0, "Imported".to_string())
        .expect("rename");
    edit.transfer_cell(
        &source,
        0,
        Reference::new(0, 0).expect("source reference"),
        0,
        destination,
    )
    .expect("dependency-managed transfer");
    let commit = edit.commit().expect("workbook commit");
    let encoded = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable patch");
    let decoded = WorkbookPatch::from_bytes(&encoded, TransferLimits::DEFAULT)
        .expect("validated durable patch");
    decoded.apply(&mut target).expect("atomic publication");

    assert_eq!(target.worksheet_names(), &["Imported".to_string()]);
    let snapshot = target.cell_values(0).expect("target cells");
    let cell = snapshot
        .cell(destination)
        .expect("unique destination")
        .expect("transferred destination");
    let Value::SharedStringIndex(index) = cell.value() else {
        panic!("writer-produced string should remain SST-backed")
    };
    assert_eq!(
        target.shared_strings()[usize::try_from(*index).expect("SST index")].text,
        "shared transfer value"
    );

    let mut bytes = Cursor::new(Vec::new());
    target.save(&mut bytes).expect("save transferred workbook");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("consumer reopen");
    assert_eq!(reopened.worksheet_names(), &["Imported".to_string()]);
    assert!(
        reopened
            .cell_values(0)
            .expect("reopened cells")
            .cell(destination)
            .expect("unique destination")
            .is_some()
    );
}

#[test]
fn ordinary_root_authors_sst_rich_style_and_formula_resources() {
    const GIF_1X1: &[u8] = &[
        0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00, 0x00, 0x00,
        0x00, 0xff, 0xff, 0xff, 0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, 0x2c, 0x00, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x44, 0x01, 0x00, 0x3b,
    ];
    let mut workbook = producer_workbook("Target", None);
    let style = AuthoredStyle {
        font: Font {
            name: "Arial".to_string(),
            size: 12.0,
            color: Some(0xff_20_40_80),
            bold: true,
            italic: false,
            underline: false,
            strike: false,
        },
        fill: Fill {
            pattern_type: 1,
            fg_color: Some(0xff_f0_e0_20),
            bg_color: None,
        },
        number_format: Some("0.0000".to_string()),
        alignment: Some(Alignment {
            horizontal: HorizontalAlignment::Center,
            vertical: VerticalAlignment::Center,
            rotation: 0,
            indent: 0,
            text_direction: 0,
            wrap_text: true,
            shrink_to_fit: false,
        }),
        ..AuthoredStyle::default()
    };
    let shared_ref = Reference::new(1, 1).expect("shared reference");
    let rich_ref = Reference::new(1, 2).expect("rich reference");
    let formula_ref = Reference::new(1, 3).expect("formula reference");
    let formula_string_ref = Reference::new(1, 4).expect("string-formula reference");
    let string = SharedString {
        text: "authored rich value".to_string(),
        runs: vec![SharedStringRun {
            character_index: 0,
            font_id: u16::MAX,
        }],
        phonetic: None,
    };
    let mut edit = workbook.edit_workbook_structure().expect("root edit");
    edit.insert_shared_string(0, shared_ref, string.clone(), &style)
        .expect("SST authoring");
    edit.insert_rich_string(0, rich_ref, string, &style)
        .expect("rich-string authoring");
    edit.insert_formula_number(
        0,
        formula_ref,
        2.0,
        CellFormula::new(0, vec![0x1e, 2, 0], vec![]).expect("constant formula"),
        &style,
    )
    .expect("formula authoring");
    edit.insert_formula(
        0,
        formula_string_ref,
        Value::FormulaStringCache("two".to_string()),
        CellFormula::new(0, vec![0x1e, 2, 0], vec![]).expect("constant string-cache formula"),
        &style,
    )
    .expect("string-cache formula authoring");
    let image = Image::new(
        GIF_1X1.to_vec(),
        ImageFormat::Gif,
        ChartAnchor::new(0, 3, 2, 7),
    )
    .expect("valid image")
    .with_description("authored one-pixel image")
    .expect("valid image description");
    edit.insert_image(0, image).expect("image authoring");
    let commit = edit.commit().expect("root commit");
    let bytes = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable resources");
    let patch = WorkbookPatch::from_bytes(&bytes, TransferLimits::DEFAULT)
        .expect("durable resource replay");
    patch.apply(&mut workbook).expect("publish resources");

    let snapshot = workbook.cell_values(0).expect("authored snapshot");
    assert!(matches!(
        snapshot
            .cell(shared_ref)
            .expect("lookup")
            .expect("shared")
            .value(),
        Value::SharedStringIndex(_)
    ));
    assert!(matches!(
        snapshot
            .cell(rich_ref)
            .expect("lookup")
            .expect("rich")
            .value(),
        Value::RichString(_)
    ));
    assert_eq!(
        snapshot
            .cell(formula_ref)
            .expect("lookup")
            .expect("formula")
            .value(),
        &Value::FormulaNumberCache(2.0)
    );
    assert_eq!(
        snapshot
            .cell(formula_string_ref)
            .expect("lookup")
            .expect("string formula")
            .value(),
        &Value::FormulaStringCache("two".to_string())
    );
    assert!(workbook.styles().get_font(1).is_some());
    let drawing = workbook.sheet_drawing(0).expect("authored drawing");
    assert_eq!(drawing.images.len(), 1);
    assert_eq!(drawing.images[0].data.as_ref(), GIF_1X1);
    let mut reopened_bytes = Cursor::new(Vec::new());
    workbook.save(&mut reopened_bytes).expect("save authored");
    let reopened =
        Workbook::new(Cursor::new(reopened_bytes.into_inner())).expect("reopen authored");
    assert_eq!(
        reopened.cell_values(0).expect("reopen cells").cells().len(),
        4
    );
    assert_eq!(
        reopened
            .sheet_drawing(0)
            .expect("reopened drawing")
            .images
            .len(),
        1
    );

    let mut transfer_target = producer_workbook("Image target", None);
    let mut transfer = transfer_target
        .edit_workbook_structure()
        .expect("image transfer edit");
    transfer
        .transfer_image(&reopened, 0, 0, 0, ChartAnchor::new(3, 4, 6, 9))
        .expect("dependency-closed image transfer");
    let transfer_commit = transfer.commit().expect("image transfer commit");
    transfer_target
        .apply_workbook_structure(&transfer_commit)
        .expect("publish transferred image");
    let transferred = transfer_target
        .sheet_drawing(0)
        .expect("transferred drawing");
    assert_eq!(transferred.images[0].data.as_ref(), GIF_1X1);
    assert_eq!(
        transferred.images[0].description.as_deref(),
        Some("authored one-pixel image")
    );

    let appended = Image::new(
        GIF_1X1.to_vec(),
        ImageFormat::Gif,
        ChartAnchor::new(7, 4, 9, 8),
    )
    .expect("valid appended image")
    .with_description("appended image")
    .expect("valid appended description");
    let mut append_edit = transfer_target
        .edit_workbook_structure()
        .expect("existing drawing edit");
    append_edit
        .insert_image(0, appended)
        .expect("append to existing drawing");
    let append_commit = append_edit.commit().expect("append image commit");
    let append_bytes = append_commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable appended image");
    WorkbookPatch::from_bytes(&append_bytes, TransferLimits::DEFAULT)
        .expect("decode appended image patch")
        .apply(&mut transfer_target)
        .expect("publish appended image");
    let appended_drawing = transfer_target.sheet_drawing(0).expect("appended drawing");
    assert_eq!(appended_drawing.images.len(), 2);
    assert_eq!(
        appended_drawing.images[1].description.as_deref(),
        Some("appended image")
    );
}

#[test]
fn durable_drawing_transfer_closes_connector_graph_and_remaps_collisions() {
    let source = shape_graph_workbook("Source shapes");
    let mut empty_target = producer_workbook("Empty target", None);
    let mut empty_edit = empty_target
        .edit_workbook_structure()
        .expect("empty drawing edit");
    empty_edit
        .transfer_drawing_object(&source, 0, 2, 0)
        .expect("new connector drawing transfer");
    let empty_commit = empty_edit.commit().expect("new drawing commit");
    empty_target
        .apply_workbook_structure(&empty_commit)
        .expect("publish new drawing");
    assert_eq!(
        empty_target
            .sheet_drawing(0)
            .expect("new target drawing")
            .drawing
            .anchors
            .len(),
        2
    );

    let mut target = shape_graph_workbook("Target shapes");
    let before = target.sheet_drawing(0).expect("target drawing");
    assert_eq!(before.drawing.anchors.len(), 3);

    let mut drawing_edit = target.edit_workbook_structure().expect("drawing edit");
    drawing_edit
        .transfer_drawing_object(&source, 0, 2, 0)
        .expect("connector closure transfer");
    let drawing_commit = drawing_edit.commit().expect("drawing commit");
    let durable = drawing_commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable drawing patch");
    let drawing_patch =
        WorkbookPatch::from_bytes(&durable, TransferLimits::DEFAULT).expect("drawing patch replay");

    let mut rename_edit = target.edit_workbook_structure().expect("rename edit");
    rename_edit
        .rename_sheet(0, "Transferred graphs".to_string())
        .expect("rename");
    let rename_commit = rename_edit.commit().expect("rename commit");
    let merged = drawing_patch
        .merge_three_way(rename_commit.patch())
        .expect("drawing/name merge");
    assert!(merged.conflicts().is_empty());
    let merged = merged.patch().expect("merged drawing patch").clone();
    merged.apply(&mut target).expect("publish drawing graph");

    let drawing = target.sheet_drawing(0).expect("transferred drawing");
    assert_eq!(drawing.drawing.anchors.len(), 5);
    assert_eq!(drawing.shapes.len(), 5);
    let mut ids = Vec::new();
    for anchored in &drawing.shapes {
        collect_object_ids(&anchored.object, &mut ids);
    }
    let mut unique = ids.clone();
    unique.sort_unstable();
    unique.dedup();
    assert_eq!(unique.len(), ids.len(), "all remapped IDs must be unique");
    let mut names = Vec::new();
    for anchored in &drawing.shapes {
        collect_object_names(&anchored.object, &mut names);
    }
    let mut unique_names = names
        .iter()
        .map(|name| name.to_lowercase())
        .collect::<Vec<_>>();
    unique_names.sort_unstable();
    unique_names.dedup();
    assert_eq!(
        unique_names.len(),
        names.len(),
        "all colliding imported names must be unique"
    );
    let ShapeObject::ConnectionShape(connection) = &drawing.shapes[4].object else {
        panic!("transferred closure should end with its connector")
    };
    assert!(connection.start.is_some());
    assert!(connection.end.is_some());

    let mut history = WorkbookHistory::new(TransferLimits::DEFAULT).expect("drawing history");
    history.push(merged).expect("retain drawing patch");
    history.undo(&mut target).expect("undo drawing graph");
    assert_eq!(
        target
            .sheet_drawing(0)
            .expect("undone target drawing")
            .drawing
            .anchors
            .len(),
        3
    );
    history.redo(&mut target).expect("redo drawing graph");
    assert_eq!(
        target.worksheet_names(),
        &["Transferred graphs".to_string()]
    );
}

#[test]
fn ordinary_chart_graph_transfer_retains_existing_drawing_and_resource_payloads() {
    const GIF_1X1: &[u8] = &[
        0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00, 0x00, 0x00,
        0x00, 0xff, 0xff, 0xff, 0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, 0x2c, 0x00, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x44, 0x01, 0x00, 0x3b,
    ];
    let chart = Chart::bar_chart(
        "Transferred chart",
        "Source!$A$1:$A$2",
        "Source!$B$1:$B$2",
        ChartAnchor::new(0, 0, 7, 12),
    )
    .expect("chart")
    .with_external_data_part(
        ExternalDataPart::embedded_workbook(b"PK exact chart workbook".to_vec()),
        Some(false),
    );
    let mut source_sheet = MutableWorksheet::new("Source");
    source_sheet.add_chart(chart).expect("source chart");
    let mut source_writer = WorkbookWriter::new();
    source_writer.add_worksheet(source_sheet);
    let mut source_bytes = Cursor::new(Vec::new());
    source_writer.save(&mut source_bytes).expect("source save");
    let source = Workbook::new(Cursor::new(source_bytes.into_inner())).expect("source reopen");

    let mut target_sheet = MutableWorksheet::new("Target");
    target_sheet
        .add_image(
            Image::new(
                GIF_1X1.to_vec(),
                ImageFormat::Gif,
                ChartAnchor::new(8, 1, 10, 4),
            )
            .expect("target image")
            .with_description("retained target image")
            .expect("target description"),
        )
        .expect("add target image");
    let mut target_writer = WorkbookWriter::new();
    target_writer.add_worksheet(target_sheet);
    let mut target_bytes = Cursor::new(Vec::new());
    target_writer.save(&mut target_bytes).expect("target save");
    let mut target = Workbook::new(Cursor::new(target_bytes.into_inner())).expect("target reopen");
    let existing_drawing_xml = standard_drawing_xml(&target);

    let mut edit = target
        .edit_workbook_structure()
        .expect("chart transfer edit");
    edit.transfer_drawing_object(&source, 0, 0, 0)
        .expect("chart graph transfer");
    let commit = edit.commit().expect("chart transfer commit");
    let durable = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable chart graph");
    WorkbookPatch::from_bytes(&durable, TransferLimits::DEFAULT)
        .expect("chart graph replay")
        .apply(&mut target)
        .expect("publish chart graph");

    let transferred_drawing_xml = standard_drawing_xml(&target);
    let root_close = existing_drawing_xml
        .windows(b"</xdr:wsDr>".len())
        .rposition(|window| window == b"</xdr:wsDr>")
        .expect("drawing root close");
    assert_eq!(
        &transferred_drawing_xml[..root_close],
        &existing_drawing_xml[..root_close],
        "all existing drawing bytes before the root close stay exact"
    );
    let retained_suffix = &existing_drawing_xml[root_close..];
    assert!(transferred_drawing_xml.ends_with(retained_suffix));

    let drawing = target.sheet_drawing(0).expect("target chart drawing");
    assert_eq!(drawing.drawing.anchors.len(), 2);
    assert_eq!(drawing.images.len(), 1);
    assert_eq!(drawing.images[0].data.as_ref(), GIF_1X1);
    assert_eq!(
        drawing.images[0].description.as_deref(),
        Some("retained target image")
    );
    assert_eq!(drawing.charts.len(), 1);
    let external = drawing.charts[0]
        .external_data_part
        .as_ref()
        .expect("transferred embedded workbook");
    let litchi_xlsb::chart::ExternalDataTarget::Embedded { data, .. } = &external.target else {
        panic!("chart external data should stay embedded")
    };
    assert_eq!(data, b"PK exact chart workbook");

    let mut picture_edit = target
        .edit_workbook_structure()
        .expect("picture refusal edit");
    let error = picture_edit
        .transfer_drawing_object(&target, 0, 0, 0)
        .expect_err("pictures use the image API");
    assert!(matches!(
        error,
        litchi_xlsb::cell_values::Error::DrawingTransfer(
            DrawingTransferRefusal::PictureUsesImageTransfer
        )
    ));
}

#[test]
fn durable_transfer_remaps_name_sheet_and_table_formula_dependencies() {
    let mut source = formula_dependency_workbook(true);
    let source_summary = 1;
    let contextual_ref = Reference::new(0, 0).expect("contextual formula reference");
    let source_cells = source.cell_values(source_summary).expect("source cells");
    let contextual = source_cells
        .cell(contextual_ref)
        .expect("unique contextual formula")
        .expect("contextual formula");
    let contextual_formula = contextual.formula().expect("formula payload");
    let source_xti =
        Parser::with_extra(contextual_formula.tokens(), contextual_formula.ancillary())
            .parse()
            .expect("parse contextual formula")
            .into_iter()
            .find_map(|token| match token {
                Token::CellRef3d { sheet_index, .. } => Some(sheet_index),
                _ => None,
            })
            .expect("Data XTI");
    let (table_tokens, table_extra) = Token::TableReference(TableReference {
        sheet_index: source_xti,
        row_type: Some(TableRowType::Data),
        columns: Some(TableColumns::One(1)),
        square_bracket_space: false,
        comma_space: false,
        data_type: TableDataType::Reference,
        invalid: false,
        list_index: Some(3),
        external: None,
    })
    .to_extended_binary()
    .expect("structured-reference encoding");
    let table_formula_ref = Reference::new(0, 1).expect("table formula reference");
    let rich_ref = Reference::new(0, 2).expect("rich resource reference");
    let rich_style = AuthoredStyle {
        font: Font {
            name: "Aptos".to_string(),
            size: 11.0,
            bold: true,
            ..Font::default()
        },
        number_format: Some("0.00".to_string()),
        ..AuthoredStyle::default()
    };
    let rich = SharedString {
        text: "resource-bearing transfer".to_string(),
        runs: vec![SharedStringRun {
            character_index: 0,
            font_id: u16::MAX,
        }],
        phonetic: None,
    };
    let mut source_edit = source.edit_workbook_structure().expect("source root edit");
    source_edit
        .insert_formula(
            source_summary,
            table_formula_ref,
            Value::FormulaNumberCache(42.5),
            CellFormula::new(0, table_tokens, table_extra).expect("table cell formula"),
            &rich_style,
        )
        .expect("table formula authoring");
    source_edit
        .insert_rich_string(source_summary, rich_ref, rich, &rich_style)
        .expect("rich resource authoring");
    let source_commit = source_edit.commit().expect("source dependency commit");
    source
        .apply_workbook_structure(&source_commit)
        .expect("publish source dependencies");

    let mut target = formula_dependency_workbook(false);
    let name_target_ref = Reference::new(4, 0).expect("name target");
    let table_target_ref = Reference::new(4, 1).expect("table target");
    let rich_target_ref = Reference::new(4, 2).expect("rich target");
    let mut transfer = target.edit_workbook_structure().expect("transfer edit");
    transfer
        .transfer_cell(&source, source_summary, contextual_ref, 0, name_target_ref)
        .expect("name and sheet dependency remap");
    transfer
        .transfer_cell(
            &source,
            source_summary,
            table_formula_ref,
            0,
            table_target_ref,
        )
        .expect("table dependency remap");
    transfer
        .transfer_cell(&source, source_summary, rich_ref, 0, rich_target_ref)
        .expect("rich style dependency transfer");
    let commit = transfer.commit().expect("dependency transfer commit");
    let durable = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable dependency patch");
    let patch = WorkbookPatch::from_bytes(&durable, TransferLimits::DEFAULT)
        .expect("decode dependency patch");
    patch.apply(&mut target).expect("publish dependency patch");

    let cells = target.cell_values(0).expect("target cells");
    let name_formula = cells
        .cell(name_target_ref)
        .expect("unique name target")
        .expect("name target")
        .formula()
        .expect("name target formula");
    let name_tokens = Parser::with_extra(name_formula.tokens(), name_formula.ancillary())
        .parse()
        .expect("parse remapped name formula");
    assert!(
        name_tokens
            .iter()
            .any(|token| matches!(token, Token::Name(2)))
    );
    let table_formula = cells
        .cell(table_target_ref)
        .expect("unique table target")
        .expect("table target")
        .formula()
        .expect("table target formula");
    let table_tokens = Parser::with_extra(table_formula.tokens(), table_formula.ancillary())
        .parse()
        .expect("parse remapped table formula");
    assert!(table_tokens.iter().any(|token| {
        matches!(
            token,
            Token::TableReference(TableReference {
                list_index: Some(9),
                ..
            })
        )
    }));
    assert!(matches!(
        cells
            .cell(rich_target_ref)
            .expect("unique rich target")
            .expect("rich target")
            .value(),
        Value::RichString(_)
    ));
    let mut bytes = Cursor::new(Vec::new());
    target.save(&mut bytes).expect("save dependency target");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen dependencies");
    assert_eq!(
        reopened
            .cell_values(0)
            .expect("reopened dependency cells")
            .cells()
            .len(),
        4
    );
    assert_eq!(reopened.tables()[0].table_id(), 9);
}

#[test]
fn third_party_fixture_cell_transfers_with_its_style_dependency_closure() {
    let fixtures = [
        "test-data/ooxml/xlsb/sample.xlsb",
        "test-data/ooxml/xlsb/Simple.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/51519.xlsb",
    ];
    let (source, sheet, source_reference) = fixtures
        .iter()
        .find_map(|relative| {
            let workbook = Workbook::new(File::open(fixture(relative)).ok()?).ok()?;
            let selected = (0..workbook.worksheet_count()).find_map(|sheet| {
                workbook.cell_values(sheet).ok().and_then(|snapshot| {
                    snapshot
                        .cells()
                        .find(|cell| cell.formula().is_none())
                        .map(|cell| (sheet, cell.reference()))
                })
            });
            selected.map(|(sheet, reference)| (workbook, sheet, reference))
        })
        .expect("a third-party fixture contains a direct stored cell");
    let destination = Reference::new(10, 7).expect("destination");
    let mut target = producer_workbook("Target", None);
    let mut edit = target.edit_workbook_structure().expect("root edit");
    edit.transfer_cell(&source, sheet, source_reference, 0, destination)
        .expect("third-party transfer plan");
    let commit = edit.commit().expect("third-party transfer commit");
    target
        .apply_workbook_structure(&commit)
        .expect("publish third-party transfer");

    let transferred_snapshot = target.cell_values(0).expect("transferred snapshot");
    let transferred = transferred_snapshot
        .cell(destination)
        .expect("unique destination")
        .expect("transferred cell");
    assert!(
        target
            .styles()
            .get_cell_format(usize::try_from(transferred.style().get()).expect("style index"))
            .is_some()
    );
    let mut bytes = Cursor::new(Vec::new());
    target.save(&mut bytes).expect("save transferred fixture");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen transfer");
    assert!(
        reopened
            .cell_values(0)
            .expect("reopened snapshot")
            .cell(destination)
            .expect("unique destination")
            .is_some()
    );
}

#[test]
fn three_way_history_and_adversarial_envelope_remain_atomic() {
    let source = producer_workbook("Source", Some("merge value"));
    let mut target = producer_workbook("Target", None);
    let destination = Reference::new(8, 2).expect("destination");

    let mut left_edit = target.edit_workbook_structure().expect("left plan");
    left_edit
        .rename_sheet(0, "Merged".to_string())
        .expect("rename");
    let left_commit = left_edit.commit().expect("left commit");

    let mut right_edit = target.edit_workbook_structure().expect("right plan");
    right_edit
        .transfer_cell(
            &source,
            0,
            Reference::new(0, 0).expect("source reference"),
            0,
            destination,
        )
        .expect("transfer");
    let right_commit = right_edit.commit().expect("right commit");
    let outcome = left_commit
        .patch()
        .merge_three_way(right_commit.patch())
        .expect("three-way merge");
    assert!(outcome.conflicts().is_empty());
    let merged = outcome.patch().expect("merged patch").clone();

    let mut corrupt = merged
        .to_bytes(TransferLimits::DEFAULT)
        .expect("encoded merge");
    let last = corrupt.len().checked_sub(1).expect("nonempty patch");
    corrupt[last] ^= 0x5a;
    assert!(WorkbookPatch::from_bytes(&corrupt, TransferLimits::DEFAULT).is_err());
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);

    merged.apply(&mut target).expect("apply merge");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
    assert!(
        merged.apply(&mut target).is_err(),
        "stale reapply must fail"
    );

    let mut history = WorkbookHistory::new(TransferLimits::DEFAULT).expect("history");
    history.push(merged).expect("retain patch");
    history.undo(&mut target).expect("undo");
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);
    assert!(
        target
            .cell_values(0)
            .expect("undone cells")
            .cell(destination)
            .expect("unique destination")
            .is_none()
    );
    history.redo(&mut target).expect("redo");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
}
