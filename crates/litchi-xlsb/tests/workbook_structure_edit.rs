#![allow(
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::wildcard_enum_match_arm,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure assertions"
)]

use litchi_core::sheet::{CellValue, traits::WorkbookTrait};
use litchi_xlsb::builder::chart::{Anchor, Chart, ExternalDataPart};
use litchi_xlsb::builder::{self, Builder, Image, ImageFormat};
use litchi_xlsb::builder::{ConnectionEndSpec, ConnectionShapeSpec, GroupSpec, ShapeSpec};
use litchi_xlsb::cell_values::structure::{AuthoredStyle, History, LimitKind, Patch};
use litchi_xlsb::cell_values::{
    CellFormula, DrawingTransferRefusal, Reference, TransferLimits, Value,
};
use litchi_xlsb::named_ranges::Definition as NamedRange;
use litchi_xlsb::package::table::{Column, Range as TableRange, Table, Type as TableType};
use litchi_xlsb::package::{PackageError, SharedString, SharedStringRun};
use litchi_xlsb::raw::{Encoder, Records, kind};
use litchi_xlsb::shapes::{
    Anchor as ShapeAnchor, CellMarker, EditAs, Emu, EmuExtent, EmuOffset, Object as ShapeObject,
    Preset,
};
use litchi_xlsb::styles::{Alignment, Fill, Font, HorizontalAlignment, VerticalAlignment};
use litchi_xlsb::{Package, Workbook};
use litchi_xlsb::{Parser, TableColumns, TableDataType, TableReference, TableRowType, Token};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

#[path = "support/projection.rs"]
mod projection;
use projection::{InPlace as _, OwnedWorkbook as _};

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn producer_workbook(name: &str, initial_value: Option<&str>) -> Workbook {
    let mut sheet = builder::Worksheet::new(name);
    if let Some(value) = initial_value {
        sheet.set_cell(0, 0, value);
    }
    let mut producer = Builder::new();
    producer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    producer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("producer save");
    Workbook::new(Cursor::new(bytes.into_inner())).expect("producer reopen")
}

fn two_operation_patch() -> Patch {
    let source = producer_workbook("Source", Some("transferred value"));
    let target = producer_workbook("Target", None);
    let mut edit = target.to_package().edit().expect("two-operation edit");
    edit.set_sheet_name(0, "Renamed".to_owned())
        .expect("rename operation")
        .expect("resolved worksheet selector");
    edit.transfer_cell(
        &source,
        0,
        Reference::new(0, 0).expect("source reference"),
        0,
        Reference::new(1, 0).expect("target reference"),
    )
    .expect("transfer operation");
    edit.commit().expect("two-operation commit").patch().clone()
}

fn workbook_with_bundle_suffix(workbook: &Workbook, suffix: &[u8]) -> Workbook {
    let mut package = workbook.to_package().into_opc();
    let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let mut rewritten = Vec::new();
    {
        let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
        for record in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.expect("workbook record");
            let mut payload = record.payload().to_vec();
            if record.kind() == kind::BUNDLE_SH {
                payload.extend_from_slice(suffix);
            }
            writer
                .write_record(record.kind(), &payload)
                .expect("rewrite workbook record");
        }
    }
    package
        .get_part_mut(&workbook_uri)
        .expect("workbook part")
        .set_blob(rewritten);
    Package::from_opc(package)
        .expect("suffix-tolerant package")
        .owned_workbook()
        .expect("suffix-tolerant workbook")
}

fn workbook_with_unpaired_bundle_name(workbook: &Workbook) -> Workbook {
    let mut package = workbook.to_package().into_opc();
    let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let mut rewritten = Vec::new();
    {
        let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
        for record in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.expect("workbook record");
            let mut payload = record.payload().to_vec();
            if record.kind() == kind::BUNDLE_SH {
                let current_id = u32::from_le_bytes(payload[4..8].try_into().expect("tab ID"));
                let strings_offset = if (1..=0xffff).contains(&current_id) {
                    8
                } else {
                    12
                };
                let relationship_units = u32::from_le_bytes(
                    payload[strings_offset..strings_offset + 4]
                        .try_into()
                        .expect("relationship length"),
                );
                let name_offset = strings_offset
                    .checked_add(4)
                    .and_then(|offset| {
                        usize::try_from(relationship_units)
                            .ok()?
                            .checked_mul(2)
                            .and_then(|bytes| offset.checked_add(bytes))
                    })
                    .expect("name offset")
                    .checked_add(4)
                    .expect("name length offset");
                let name_units = u32::from_le_bytes(
                    payload[name_offset - 4..name_offset]
                        .try_into()
                        .expect("name length"),
                );
                assert_eq!(name_units, 1, "fixture name must contain one UTF-16 unit");
                payload[name_offset..name_offset + 2].copy_from_slice(&0xd800_u16.to_le_bytes());
            }
            writer
                .write_record(record.kind(), &payload)
                .expect("rewrite workbook record");
        }
    }
    package
        .get_part_mut(&workbook_uri)
        .expect("workbook part")
        .set_blob(rewritten);
    Package::from_opc(package)
        .expect("unpaired-name package")
        .owned_workbook()
        .expect("unpaired-name workbook")
}

fn assert_limit(error: PackageError, expected: LimitKind) {
    let PackageError::Limit { kind, .. } = error else {
        panic!("expected typed limit error, got {error:?}");
    };
    assert_eq!(kind, expected.into());
}

fn publish_structure_patch(workbook: &mut Workbook, patch: &Patch) {
    *workbook = workbook
        .to_package()
        .apply(patch)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("publish structure patch")
        .owned_workbook()
        .expect("reopen published structure patch");
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
    let mut sheet = builder::Worksheet::new(name);
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
    let mut writer = Builder::new();
    writer.add_worksheet(sheet);
    let mut bytes = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("shape graph save");
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
    workbook
        .to_package()
        .write_to(&mut bytes)
        .expect("workbook package save");
    let package = Package::from_bytes(bytes.get_ref()).expect("package reopen");
    package
        .opc_package()
        .try_iter_parts()
        .find_map(|part| {
            let part = part.expect("drawing part decodes");
            (part.content_type() == "application/vnd.openxmlformats-officedocument.drawing+xml")
                .then_some(part)
        })
        .expect("standard drawing part")
        .blob()
        .to_vec()
}

fn standard_drawing_external_hyperlinks(workbook: &Workbook) -> Vec<(String, String)> {
    let mut bytes = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut bytes)
        .expect("workbook package save");
    let package = Package::from_bytes(bytes.get_ref()).expect("package reopen");
    let drawing = package
        .opc_package()
        .iter_parts()
        .find(|part| {
            part.content_type() == "application/vnd.openxmlformats-officedocument.drawing+xml"
        })
        .expect("standard drawing part");
    let mut hyperlinks = drawing
        .rels()
        .iter()
        .filter(|relationship| {
            relationship.is_external() && relationship.reltype().ends_with("/hyperlink")
        })
        .map(|relationship| {
            (
                relationship.r_id().to_string(),
                relationship.target_ref().to_string(),
            )
        })
        .collect::<Vec<_>>();
    hyperlinks.sort_unstable();
    hyperlinks
}

fn formula_dependency_workbook(source_order: bool) -> Workbook {
    let mut data = builder::Worksheet::new("Data");
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
    let mut summary = builder::Worksheet::new("Summary");
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
    let mut producer = Builder::new();
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
    producer
        .build()
        .expect("build package")
        .write_to(&mut bytes)
        .expect("dependency producer save");
    Workbook::new(Cursor::new(bytes.into_inner())).expect("dependency producer reopen")
}

#[test]
fn producer_reopen_transfers_sst_and_renames_on_a_durable_root() {
    let source = producer_workbook("Source", Some("shared transfer value"));
    let mut target = producer_workbook("Target", None);
    let destination = Reference::new(4, 3).expect("destination");

    let mut edit = target.to_package().edit().expect("detached workbook edit");
    edit.set_sheet_name(0, "Imported".to_string())
        .expect("rename")
        .expect("resolved worksheet selector");
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
    let decoded = Patch::from_bytes(
        &encoded,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("validated durable patch");
    target
        .apply_structure(&decoded)
        .expect("atomic publication");

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
    target
        .to_package()
        .write_to(&mut bytes)
        .expect("save transferred workbook");
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
    let mut edit = workbook.to_package().edit().expect("root edit");
    edit.insert_shared_string(0, shared_ref, string.clone(), &style)
        .expect("SST authoring");
    edit.insert_rich_string(0, rich_ref, string, &style)
        .expect("rich-string authoring");
    edit.insert_formula_number(
        0,
        formula_ref,
        2.0,
        CellFormula::new(
            0,
            vec![0x1e, 2, 0],
            vec![],
            litchi_xlsb::raw::Limits::DEFAULT,
        )
        .expect("constant formula"),
        &style,
    )
    .expect("formula authoring");
    edit.insert_formula(
        0,
        formula_string_ref,
        Value::FormulaStringCache("two".to_string()),
        CellFormula::new(
            0,
            vec![0x1e, 2, 0],
            vec![],
            litchi_xlsb::raw::Limits::DEFAULT,
        )
        .expect("constant string-cache formula"),
        &style,
    )
    .expect("string-cache formula authoring");
    let image = Image::new(
        GIF_1X1.to_vec(),
        ImageFormat::Gif,
        Anchor::new(0, 3, 2, 7),
        &litchi_xlsb::package::drawing_image::Limits::DEFAULT,
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
    let patch = Patch::from_bytes(
        &bytes,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("durable resource replay");
    workbook.apply_structure(&patch).expect("publish resources");

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
    workbook
        .to_package()
        .write_to(&mut reopened_bytes)
        .expect("save authored");
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
        .to_package()
        .edit()
        .expect("image transfer edit");
    transfer
        .transfer_image(&reopened, 0, 0, 0, Anchor::new(3, 4, 6, 9))
        .expect("dependency-closed image transfer");
    let transfer_commit = transfer.commit().expect("image transfer commit");
    publish_structure_patch(&mut transfer_target, transfer_commit.patch());
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
        Anchor::new(7, 4, 9, 8),
        &litchi_xlsb::package::drawing_image::Limits::DEFAULT,
    )
    .expect("valid appended image")
    .with_description("appended image")
    .expect("valid appended description");
    let mut append_edit = transfer_target
        .to_package()
        .edit()
        .expect("existing drawing edit");
    append_edit
        .insert_image(0, appended)
        .expect("append to existing drawing");
    let append_commit = append_edit.commit().expect("append image commit");
    let append_bytes = append_commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable appended image");
    publish_structure_patch(
        &mut transfer_target,
        &Patch::from_bytes(
            &append_bytes,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT,
        )
        .expect("decode appended image patch"),
    );
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
        .to_package()
        .edit()
        .expect("empty drawing edit");
    empty_edit
        .transfer_drawing_object(&source, 0, 2, 0)
        .expect("new connector drawing transfer");
    let empty_commit = empty_edit.commit().expect("new drawing commit");
    publish_structure_patch(&mut empty_target, empty_commit.patch());
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

    let mut drawing_edit = target.to_package().edit().expect("drawing edit");
    drawing_edit
        .transfer_drawing_object(&source, 0, 2, 0)
        .expect("connector closure transfer");
    let drawing_commit = drawing_edit.commit().expect("drawing commit");
    let durable = drawing_commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable drawing patch");
    let drawing_patch = Patch::from_bytes(
        &durable,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("drawing patch replay");

    let mut rename_edit = target.to_package().edit().expect("rename edit");
    rename_edit
        .set_sheet_name(0, "Transferred graphs".to_string())
        .expect("rename")
        .expect("resolved worksheet selector");
    let rename_commit = rename_edit.commit().expect("rename commit");
    let merged = drawing_patch
        .merge_three_way(rename_commit.patch())
        .expect("drawing/name merge");
    assert!(merged.conflicts().is_empty());
    let merged = merged.patch().expect("merged drawing patch").clone();
    target
        .apply_structure(&merged)
        .expect("publish drawing graph");

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

    let mut history = History::new(TransferLimits::DEFAULT).expect("drawing history");
    history.push(merged).expect("retain drawing patch");
    target.undo_with(&mut history).expect("undo drawing graph");
    assert_eq!(
        target
            .sheet_drawing(0)
            .expect("undone target drawing")
            .drawing
            .anchors
            .len(),
        3
    );
    target.redo_with(&mut history).expect("redo drawing graph");
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
        Anchor::new(0, 0, 7, 12),
    )
    .expect("chart")
    .with_external_data_part(
        ExternalDataPart::embedded_workbook(b"PK exact chart workbook".to_vec()),
        Some(false),
    );
    let mut source_sheet = builder::Worksheet::new("Source");
    source_sheet.add_chart(chart).expect("source chart");
    let mut source_writer = Builder::new();
    source_writer.add_worksheet(source_sheet);
    let mut source_bytes = Cursor::new(Vec::new());
    source_writer
        .build()
        .expect("build package")
        .write_to(&mut source_bytes)
        .expect("source save");
    let source = Workbook::new(Cursor::new(source_bytes.into_inner())).expect("source reopen");

    let mut target_sheet = builder::Worksheet::new("Target");
    target_sheet
        .add_image(
            Image::new(
                GIF_1X1.to_vec(),
                ImageFormat::Gif,
                Anchor::new(8, 1, 10, 4),
                &litchi_xlsb::package::drawing_image::Limits::DEFAULT,
            )
            .expect("target image")
            .with_description("retained target image")
            .expect("target description"),
        )
        .expect("add target image");
    let mut target_writer = Builder::new();
    target_writer.add_worksheet(target_sheet);
    let mut target_bytes = Cursor::new(Vec::new());
    target_writer
        .build()
        .expect("build package")
        .write_to(&mut target_bytes)
        .expect("target save");
    let mut target = Workbook::new(Cursor::new(target_bytes.into_inner())).expect("target reopen");
    let existing_drawing_xml = standard_drawing_xml(&target);

    let mut edit = target.to_package().edit().expect("chart transfer edit");
    edit.transfer_drawing_object(&source, 0, 0, 0)
        .expect("chart graph transfer");
    let commit = edit.commit().expect("chart transfer commit");
    let durable = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable chart graph");
    publish_structure_patch(
        &mut target,
        &Patch::from_bytes(
            &durable,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT,
        )
        .expect("chart graph replay"),
    );

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
    let builder::chart::ExternalDataTarget::Embedded { data, .. } = &external.target else {
        panic!("chart external data should stay embedded")
    };
    assert_eq!(data, b"PK exact chart workbook");

    let mut picture_edit = target.to_package().edit().expect("picture refusal edit");
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
fn genuine_shape_hyperlink_transfer_is_durable_mergeable_and_reversible() {
    let source = Workbook::new(
        File::open(fixture(
            "test-data/poi/test-data/spreadsheet/testVarious.xlsb",
        ))
        .expect("genuine hyperlink-shape fixture"),
    )
    .expect("genuine hyperlink-shape workbook");
    assert_eq!(
        standard_drawing_external_hyperlinks(&source),
        vec![("rId1".to_string(), "http://lucene.apache.org/".to_string())]
    );
    let mut target = producer_workbook("Hyperlink target", None);
    let mut drawing_edit = target.to_package().edit().expect("hyperlink edit");
    drawing_edit
        .transfer_drawing_object(&source, 0, 0, 0)
        .expect("genuine shape hyperlink transfer");
    let drawing_commit = drawing_edit.commit().expect("hyperlink commit");
    let durable = drawing_commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable hyperlink patch");
    let drawing_patch = Patch::from_bytes(
        &durable,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("durable hyperlink replay");

    let mut rename_edit = target.to_package().edit().expect("rename edit");
    rename_edit
        .set_sheet_name(0, "Hyperlinked shape".to_string())
        .expect("rename")
        .expect("resolved worksheet selector");
    let rename_commit = rename_edit.commit().expect("rename commit");
    let merged = drawing_patch
        .merge_three_way(rename_commit.patch())
        .expect("hyperlink/name merge");
    assert!(merged.conflicts().is_empty());
    let merged = merged.patch().expect("merged hyperlink patch").clone();
    target
        .apply_structure(&merged)
        .expect("publish hyperlink shape");
    assert_eq!(
        standard_drawing_external_hyperlinks(&target),
        vec![("rId1".to_string(), "http://lucene.apache.org/".to_string())]
    );

    let mut history = History::new(TransferLimits::DEFAULT).expect("hyperlink history");
    history.push(merged).expect("retain hyperlink patch");
    target
        .undo_with(&mut history)
        .expect("inverse hyperlink patch");
    assert!(target.sheet_drawing(0).is_none());
    target
        .redo_with(&mut history)
        .expect("redo hyperlink patch");

    let mut reopened_bytes = Cursor::new(Vec::new());
    target
        .to_package()
        .write_to(&mut reopened_bytes)
        .expect("save transferred hyperlink workbook");
    let reopened = Workbook::new(Cursor::new(reopened_bytes.into_inner()))
        .expect("full reopen transferred hyperlink workbook");
    assert_eq!(
        reopened.worksheet_names(),
        &["Hyperlinked shape".to_string()]
    );
    assert_eq!(
        standard_drawing_external_hyperlinks(&reopened),
        vec![("rId1".to_string(), "http://lucene.apache.org/".to_string())]
    );
}

#[test]
fn checked_in_unique_standard_drawing_corpus_transfers_every_anchor() {
    let corpus: &[(&str, &[usize], usize, usize, Option<usize>)] = &[
        (
            "test-data/libreoffice-core/sc/qa/unit/data/xlsb/tdf108017_calcProtection.xlsb",
            &[0],
            1,
            0,
            Some(1),
        ),
        (
            "test-data/ooxml/xlsb/universal-content.xlsb",
            &[0, 1],
            1,
            1,
            None,
        ),
        (
            "test-data/poi/test-data/spreadsheet/WithTextBox.xlsb",
            &[0, 1, 2],
            1,
            1,
            None,
        ),
        (
            "test-data/poi/test-data/spreadsheet/testVarious.xlsb",
            &[0],
            3,
            3,
            None,
        ),
    ];
    let mut parsed = 0usize;
    let mut transferred = 0usize;
    for &(
        relative,
        worksheet_catalog_positions,
        expected_parsed,
        expected_transfers,
        chart_catalog_position,
    ) in corpus
    {
        let source = Workbook::new(File::open(fixture(relative)).expect("drawing corpus fixture"))
            .expect("drawing corpus workbook");

        assert_eq!(
            source.worksheet_count(),
            worksheet_catalog_positions.len(),
            "worksheet catalog map changed for {relative}"
        );
        let parsed_for_fixture = source
            .sheet_drawings()
            .iter()
            .map(|drawing| drawing.drawing.anchors.len())
            .sum::<usize>();
        assert_eq!(
            parsed_for_fixture, expected_parsed,
            "parsed drawing anchor census changed for {relative}"
        );
        if let Some(catalog_position) = chart_catalog_position {
            assert!(
                source.chart_sheet(catalog_position).is_some(),
                "expected chart-sheet catalog position {catalog_position} for {relative}"
            );
            assert_eq!(
                source
                    .sheet_drawing(catalog_position)
                    .expect("chart-sheet drawing inventory")
                    .drawing
                    .anchors
                    .len(),
                1,
                "chart-sheet anchor census changed for {relative}"
            );
        }

        let mut transferred_for_fixture = 0usize;
        for (worksheet_ordinal, &catalog_position) in worksheet_catalog_positions.iter().enumerate()
        {
            let Some(drawing) = source.sheet_drawing(catalog_position) else {
                continue;
            };
            for anchor in 0..drawing.drawing.anchors.len() {
                let mut target = producer_workbook("Corpus target", None);
                let mut edit = target.to_package().edit().expect("corpus edit");
                edit.transfer_drawing_object(&source, worksheet_ordinal, anchor, 0)
                    .expect("corpus drawing transfer");
                let commit = edit.commit().expect("corpus commit");
                publish_structure_patch(&mut target, commit.patch());
                let mut bytes = Cursor::new(Vec::new());
                target
                    .to_package()
                    .write_to(&mut bytes)
                    .expect("save corpus transfer");
                let reopened = Workbook::new(Cursor::new(bytes.into_inner()))
                    .expect("full reopen corpus transfer");
                let reopened_drawing = reopened
                    .sheet_drawing(0)
                    .expect("full reopen corpus transfer drawing");
                assert_eq!(
                    reopened_drawing.drawing.anchors.len(),
                    1,
                    "one selected corpus anchor must survive reopen for {relative} worksheet {worksheet_ordinal} anchor {anchor}"
                );
                transferred_for_fixture = transferred_for_fixture.saturating_add(1);
                transferred = transferred.saturating_add(1);
            }
        }
        assert_eq!(
            transferred_for_fixture, expected_transfers,
            "worksheet transfer anchor census changed for {relative}"
        );
        parsed = parsed.saturating_add(parsed_for_fixture);
    }
    assert_eq!(parsed, 6, "the parsed corpus anchor census changed");
    assert_eq!(
        transferred, 5,
        "the supported worksheet transfer anchor census changed"
    );
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
    let source_xti = Parser::with_extra(
        contextual_formula.tokens(),
        contextual_formula.ancillary(),
        litchi_xlsb::raw::Limits::DEFAULT,
    )
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
    .to_extended_binary(litchi_xlsb::raw::Limits::DEFAULT)
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
    let mut source_edit = source.to_package().edit().expect("source root edit");
    source_edit
        .insert_formula(
            source_summary,
            table_formula_ref,
            Value::FormulaNumberCache(42.5),
            CellFormula::new(
                0,
                table_tokens,
                table_extra,
                litchi_xlsb::raw::Limits::DEFAULT,
            )
            .expect("table cell formula"),
            &rich_style,
        )
        .expect("table formula authoring");
    source_edit
        .insert_rich_string(source_summary, rich_ref, rich, &rich_style)
        .expect("rich resource authoring");
    let source_commit = source_edit.commit().expect("source dependency commit");
    publish_structure_patch(&mut source, source_commit.patch());

    let mut target = formula_dependency_workbook(false);
    let name_target_ref = Reference::new(4, 0).expect("name target");
    let table_target_ref = Reference::new(4, 1).expect("table target");
    let rich_target_ref = Reference::new(4, 2).expect("rich target");
    let mut transfer = target.to_package().edit().expect("transfer edit");
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
    let patch = Patch::from_bytes(
        &durable,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode dependency patch");
    target
        .apply_structure(&patch)
        .expect("publish dependency patch");

    let cells = target.cell_values(0).expect("target cells");
    let name_formula = cells
        .cell(name_target_ref)
        .expect("unique name target")
        .expect("name target")
        .formula()
        .expect("name target formula");
    let name_tokens = Parser::with_extra(
        name_formula.tokens(),
        name_formula.ancillary(),
        litchi_xlsb::raw::Limits::DEFAULT,
    )
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
    let table_tokens = Parser::with_extra(
        table_formula.tokens(),
        table_formula.ancillary(),
        litchi_xlsb::raw::Limits::DEFAULT,
    )
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
    target
        .to_package()
        .write_to(&mut bytes)
        .expect("save dependency target");
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
    let mut edit = target.to_package().edit().expect("root edit");
    edit.transfer_cell(&source, sheet, source_reference, 0, destination)
        .expect("third-party transfer plan");
    let commit = edit.commit().expect("third-party transfer commit");
    publish_structure_patch(&mut target, commit.patch());

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
    target
        .to_package()
        .write_to(&mut bytes)
        .expect("save transferred fixture");
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

    let mut left_edit = target.to_package().edit().expect("left plan");
    left_edit
        .set_sheet_name(0, "Merged".to_string())
        .expect("rename")
        .expect("resolved worksheet selector");
    let left_commit = left_edit.commit().expect("left commit");

    let mut right_edit = target.to_package().edit().expect("right plan");
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
    assert!(
        Patch::from_bytes(
            &corrupt,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);

    target.apply_structure(&merged).expect("apply merge");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
    let error = target
        .apply_structure(&merged)
        .expect_err("stale reapply must fail");
    assert!(matches!(
        error,
        PackageError::PatchConflict { part: "workbook" }
    ));

    let mut history = History::new(TransferLimits::DEFAULT).expect("history");
    history.push(merged).expect("retain patch");
    target.undo_with(&mut history).expect("undo");
    assert_eq!(target.worksheet_names(), &["Target".to_string()]);
    assert!(
        target
            .cell_values(0)
            .expect("undone cells")
            .cell(destination)
            .expect("unique destination")
            .is_none()
    );
    target.redo_with(&mut history).expect("redo");
    assert_eq!(target.worksheet_names(), &["Merged".to_string()]);
}

#[test]
fn structural_diagnostics_distinguish_images_from_operation_counts() {
    let mut workbook = producer_workbook("Original", None);
    let empty = workbook
        .to_package()
        .edit()
        .expect("edit")
        .commit()
        .expect("commit");
    assert!(!empty.diagnostics().changed());
    assert!(empty.patch().is_noop());
    assert!(empty.patch().inverse().is_noop());

    let mut edit = workbook.to_package().edit().expect("edit");
    edit.set_sheet_name(0, "Original".to_owned())
        .expect("same name")
        .expect("resolved worksheet selector");
    let same = edit.commit().expect("same-name commit");
    assert_eq!(same.patch().operation_count(), 1);
    assert!(!same.diagnostics().changed());
    assert!(same.patch().is_noop());
    let wire = same
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("wire");
    let decoded = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode");
    assert!(decoded.is_noop());
    workbook.apply_structure(&decoded).expect("no-op apply");
    assert_eq!(workbook.worksheet_names(), &["Original".to_owned()]);

    let mut edit = workbook.to_package().edit().expect("edit");
    edit.set_sheet_name(0, "Changed".to_owned())
        .expect("rename")
        .expect("resolved worksheet selector");
    let changed = edit.commit().expect("changed commit");
    assert!(changed.diagnostics().changed());
    assert!(!changed.patch().is_noop());
    workbook.apply_structure(changed.patch()).expect("apply");
    assert!(workbook.apply_structure(empty.patch()).is_err());
    assert!(workbook.apply_structure(same.patch()).is_err());
    assert_eq!(workbook.worksheet_names(), &["Changed".to_owned()]);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.operation_count(), 0);
    assert!(!inverse.is_noop());
    let wire = inverse
        .to_bytes(TransferLimits::DEFAULT)
        .expect("inverse wire");
    let decoded = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("inverse decode");
    assert!(!decoded.is_noop());
    workbook.apply_structure(&decoded).expect("inverse apply");
    assert_eq!(workbook.worksheet_names(), &["Original".to_owned()]);
    assert_eq!(changed.patch().before(), empty.patch().before());
    assert_eq!(changed.clone().diagnostics(), changed.diagnostics());
}

#[test]
fn rename_preserves_tolerated_bundle_sheet_suffix() {
    const SUFFIX: &[u8] = &[0xde, 0xad, 0xbe, 0xef];

    let workbook = workbook_with_bundle_suffix(&producer_workbook("Original", None), SUFFIX);
    let mut edit = workbook.to_package().edit().expect("suffix edit");
    edit.set_sheet_name(0, "Renamed".to_owned())
        .expect("rename with suffix")
        .expect("resolved worksheet selector");
    let commit = edit.commit().expect("rename commit");
    let wire = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("encode suffix patch");
    let decoded = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode suffix patch");
    let replayed = workbook
        .to_package()
        .apply(&decoded)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("replay suffix patch");
    assert_eq!(
        replayed.to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = replayed
        .apply(&decoded.inverse())
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("inverse suffix patch");
    assert_eq!(
        restored.to_bytes().unwrap(),
        workbook.to_package().to_bytes().unwrap()
    );
    let renamed = commit
        .into_package()
        .owned_workbook()
        .expect("renamed workbook");

    let package = renamed.to_package().into_opc();
    let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let bundle = Records::new(
        package
            .get_part(&workbook_uri)
            .expect("workbook part")
            .blob(),
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .map(|record| record.expect("workbook record"))
    .find(|record| record.kind() == kind::BUNDLE_SH)
    .expect("bundle sheet");
    assert_eq!(renamed.worksheet_names(), &["Renamed".to_owned()]);
    assert!(bundle.payload().ends_with(SUFFIX));
}

#[test]
fn same_tolerated_unpaired_name_is_an_exact_durable_noop() {
    let workbook = workbook_with_unpaired_bundle_name(&producer_workbook("X", None));
    assert_eq!(workbook.worksheet_names(), &["\u{fffd}".to_owned()]);
    let source_bytes = workbook.to_package().to_bytes().expect("source bytes");

    let mut edit = workbook.to_package().edit().expect("no-op edit");
    edit.set_sheet_name(0, "\u{fffd}".to_owned())
        .expect("same tolerated name")
        .expect("resolved worksheet selector");
    let commit = edit.commit().expect("no-op commit");
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().expect("commit bytes"),
        source_bytes
    );

    let wire = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable no-op");
    let decoded = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode no-op");
    assert!(decoded.is_noop());
    let applied = workbook
        .to_package()
        .apply(&decoded)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("apply no-op");
    assert_eq!(applied.to_bytes().expect("applied bytes"), source_bytes);
}

#[test]
fn structural_quotas_return_typed_limit_errors() {
    let source = producer_workbook("Source", Some("transferred value"));
    let target = producer_workbook("Target", None);
    let limits = TransferLimits::new(
        TransferLimits::DEFAULT.bytes(),
        1,
        TransferLimits::DEFAULT.history_entries(),
        TransferLimits::DEFAULT.history_bytes(),
    );
    let mut edit = target
        .to_package()
        .edit_with_options(
            litchi_xlsb::cell_values::structure::EditOptions::new().transfer_limits(limits),
        )
        .expect("limited edit");
    edit.set_sheet_name(0, "Renamed".to_owned())
        .expect("first operation")
        .expect("resolved worksheet selector");
    let error = edit
        .transfer_cell(
            &source,
            0,
            Reference::new(0, 0).expect("source reference"),
            0,
            Reference::new(1, 0).expect("target reference"),
        )
        .expect_err("second operation quota");
    assert_limit(error, LimitKind::Changes);

    let patch = two_operation_patch();
    let wire = patch
        .to_bytes(TransferLimits::DEFAULT)
        .expect("default patch encoding");

    let operation_limits = TransferLimits::new(
        TransferLimits::DEFAULT.bytes(),
        1,
        TransferLimits::DEFAULT.history_entries(),
        TransferLimits::DEFAULT.history_bytes(),
    );
    assert_limit(
        patch
            .to_bytes(operation_limits)
            .expect_err("patch operation quota"),
        LimitKind::Changes,
    );
    assert_limit(
        Patch::from_bytes(&wire, operation_limits, litchi_xlsb::raw::Limits::DEFAULT)
            .expect_err("decoded operation quota"),
        LimitKind::Changes,
    );

    let byte_limits = TransferLimits::new(
        wire.len().saturating_sub(1),
        TransferLimits::DEFAULT.changes(),
        TransferLimits::DEFAULT.history_entries(),
        TransferLimits::DEFAULT.history_bytes(),
    );
    assert_limit(
        patch.to_bytes(byte_limits).expect_err("patch byte quota"),
        LimitKind::PatchBytes,
    );
    assert_limit(
        Patch::from_bytes(&wire, byte_limits, litchi_xlsb::raw::Limits::DEFAULT)
            .expect_err("decoded byte quota"),
        LimitKind::PatchBytes,
    );

    let history_limits = TransferLimits::new(
        TransferLimits::DEFAULT.bytes(),
        TransferLimits::DEFAULT.changes(),
        TransferLimits::DEFAULT.history_entries(),
        1,
    );
    let mut history = History::new(history_limits).expect("limited history");
    assert_limit(
        history.push(patch).expect_err("history byte quota"),
        LimitKind::HistoryBytes,
    );
}
