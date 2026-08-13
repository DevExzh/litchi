use super::compatibility_document_from_bytes;
use super::*;
use crate::archive::{Archive, ArchiveObject, FieldInfo, FieldPath};
use crate::numbers::{NumbersDocumentBuilder, SemanticTableCellAssertions};
use crate::package_metadata::{PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE};
use crate::protobuf::tn;
use crate::protobuf::tsp::{ComponentInfo, ObjectUuidMapEntry, PackageMetadata, Reference, Uuid};
use crate::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::table::cell::{
    BorderSide,
    layout::{Inset, Insets, Layout, TextWrap, VerticalAlignment},
};
use litchi_numbers::table::lock::State as FocusedTableLockState;
use litchi_numbers::{Package as FocusedNumbersPackage, SheetSelector, TableSelector};

fn cell_number(value: f64) -> CellValue {
    CellValue::number(value).expect("finite test number")
}

fn cached_number(value: f64) -> FormulaCachedValue {
    FormulaCachedValue::number(value).expect("finite cached test number")
}

fn cached_scalar_number(value: f64) -> crate::numbers::bnc::CachedScalar {
    crate::numbers::bnc::CachedScalar::Number(
        litchi_iwa_common::formula::FiniteF64::new(value)
            .expect("finite cached scalar test number"),
    )
}

#[test]
fn ordinary_text_box_crud_is_guarded_and_byte_exact() {
    let mut editor = NumbersEditor::from_package(test_package_with_text_box()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let text_boxes = editor.sheet_text_boxes(2).unwrap();
    assert_eq!(text_boxes.len(), 1);
    assert_eq!(text_boxes[0].drawable_object_id, 50);
    assert_eq!(text_boxes[0].storage.id, TextStorageId::new(53).unwrap());
    assert_eq!(text_boxes[0].storage.storage.text(), "Source");

    editor
        .replace_sheet_text_box_text(2, 50, 0..6, "Edited 🚀")
        .unwrap();
    assert_eq!(
        editor.sheet_text_boxes(2).unwrap()[0]
            .storage
            .storage
            .text(),
        "Edited 🚀"
    );
    editor.set_sheet_text_box_text(2, 50, "Source").unwrap();

    let original_geometry = editor.sheet_text_box_geometry(2, 50).unwrap();
    let changed_geometry = DrawableGeometry {
        position: Some(DrawablePoint { x: 140.0, y: 90.0 }),
        size: Some(DrawableSize {
            width: 260.0,
            height: 70.0,
        }),
        flags: Some(3),
        angle: Some(15.0),
    };
    editor
        .set_sheet_text_box_geometry(2, 50, changed_geometry)
        .unwrap();
    assert_eq!(
        editor.sheet_text_box_geometry(2, 50).unwrap(),
        changed_geometry
    );
    editor
        .set_sheet_text_box_geometry(2, 50, original_geometry)
        .unwrap();

    let original_properties = editor.sheet_text_box_properties(2, 50).unwrap();
    let changed_properties = DrawableProperties {
        hyperlink_url: Some("https://example.test/numbers-text-box".to_owned()),
        locked: Some(true),
        aspect_ratio_locked: Some(true),
        accessibility_description: Some("Accessible Numbers text box".to_owned()),
    };
    editor
        .set_sheet_text_box_properties(2, 50, changed_properties.clone())
        .unwrap();
    assert_eq!(
        editor.sheet_text_box_properties(2, 50).unwrap(),
        changed_properties
    );
    editor
        .set_sheet_text_box_properties(2, 50, original_properties)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(editor.set_sheet_text_box_text(999, 50, "no").is_err());
    assert!(editor.set_sheet_text_box_text(2, 3, "no").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn sheet_owned_drawable_comment_crud_is_guarded_and_byte_exact() {
    let mut editor = NumbersEditor::from_package(test_package_with_text_box()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert_eq!(
        editor
            .sheet_drawables(2)
            .unwrap()
            .into_iter()
            .map(|drawable| drawable.id.get())
            .collect::<Vec<_>>(),
        vec![50]
    );
    assert!(editor.sheet_drawable_comment(2, 50).unwrap().is_none());

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_sheet_drawable_comment(999, 50, "Wrong sheet")
            .is_err()
    );
    assert!(
        editor
            .set_sheet_drawable_comment(2, 3, "Unsupported table")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    editor
        .set_sheet_drawable_comment(2, 50, "Sheet annotation")
        .unwrap();
    let comment = editor.sheet_drawable_comment(2, 50).unwrap().unwrap();
    assert_eq!(comment.drawable_id.get(), 50);
    assert_eq!(comment.comment.text, "Sheet annotation");
    assert_eq!(
        editor.sheet_text_boxes(2).unwrap()[0]
            .storage
            .storage
            .text(),
        "Source"
    );
    let bytes = editor.to_bytes().unwrap();
    editor
        .set_sheet_drawable_comment(2, 50, "Sheet annotation")
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), bytes);

    let mut reparsed = NumbersEditor::from_bytes(&bytes).unwrap();
    assert_eq!(
        reparsed
            .sheet_drawable_comment(2, 50)
            .unwrap()
            .unwrap()
            .comment
            .text,
        "Sheet annotation"
    );
    reparsed.clear_sheet_drawable_comment(2, 50).unwrap();
    assert!(reparsed.sheet_drawable_comment(2, 50).unwrap().is_none());
    assert_eq!(reparsed.to_bytes().unwrap(), baseline);
}

#[test]
fn duplicate_and_cross_sheet_drawable_ownership_fail_transactionally() {
    let mut duplicate_owner = test_package_with_text_box();
    duplicate_owner
        .update_archive("Index/Document.iwa", |archive| {
            let sheet = archive.object_mut(2).unwrap();
            let mut decoded = tn::SheetArchive::decode(sheet.messages[0].data.as_slice())?;
            decoded.drawable_infos.push(Reference {
                identifier: 50,
                ..Default::default()
            });
            sheet.replace_message(
                0,
                RawMessage {
                    type_: 2,
                    data: decoded.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(duplicate_owner).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.sheet_drawables(2).is_err());
    assert!(
        editor
            .set_sheet_drawable_comment(2, 50, "Rejected")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut cross_sheet = test_package_with_text_box();
    cross_sheet
        .update_archive("Index/Document.iwa", |archive| {
            let root = archive.object_mut(1).unwrap();
            let mut document = tn::DocumentArchive::decode(root.messages[0].data.as_slice())?;
            document.sheets.push(Reference {
                identifier: 60,
                ..Default::default()
            });
            root.replace_message(
                0,
                RawMessage {
                    type_: 1,
                    data: document.encode_to_vec(),
                },
            )?;
            archive.insert_object(ArchiveObject::new(
                60,
                vec![RawMessage {
                    type_: 2,
                    data: tn::SheetArchive {
                        name: "Second".to_owned(),
                        drawable_infos: vec![Reference {
                            identifier: 50,
                            ..Default::default()
                        }],
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(cross_sheet).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.sheet_drawables(2).is_err());
    assert!(editor.clear_sheet_drawable_comment(60, 50).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn ordinary_text_box_duplicate_delete_is_independent_and_exact() {
    let mut editor = NumbersEditor::from_package(test_package_with_text_box()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let source_geometry = editor.sheet_text_box_geometry(2, 50).unwrap();
    let created = editor
        .duplicate_sheet_text_box(2, 50, "Independent clone")
        .unwrap();
    assert_ne!(created.drawable_object_id, 50);
    assert_ne!(created.storage.id, TextStorageId::new(53).unwrap());
    assert_eq!(created.storage.storage.text(), "Independent clone");
    assert_eq!(editor.sheet_text_boxes(2).unwrap().len(), 2);
    assert_eq!(
        editor.sheet_text_boxes(2).unwrap()[0]
            .storage
            .storage
            .text(),
        "Source"
    );
    let clone_geometry = editor
        .sheet_text_box_geometry(2, created.drawable_object_id)
        .unwrap();
    assert_eq!(
        clone_geometry.position,
        source_geometry.position.map(|position| DrawablePoint {
            x: position.x + DRAWABLE_DUPLICATE_OFFSET,
            y: position.y + DRAWABLE_DUPLICATE_OFFSET,
        })
    );

    let removed = editor
        .remove_sheet_text_box(2, created.drawable_object_id)
        .unwrap();
    assert_eq!(removed.text_box.storage.storage.text(), "Independent clone");
    assert_eq!(editor.sheet_text_boxes(2).unwrap().len(), 1);
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn populated_sheet_duplicate_is_ordered_and_independent() {
    const SOURCE_SHEET_ID: u64 = 2;
    const SOURCE_TABLE_ID: u64 = 10;
    const SOURCE_TEXT_BOX_ID: u64 = 50;
    const SOURCE_CELL_ROW: usize = 0;
    const SOURCE_CELL_COLUMN: usize = 1;

    let mut editor = NumbersEditor::from_package(test_package_with_text_box()).unwrap();
    let created = editor
        .duplicate_sheet(test_sheet_selector(&editor, SOURCE_SHEET_ID))
        .unwrap();
    assert_ne!(created.object_id, SOURCE_SHEET_ID);
    assert_eq!(created.index, 1);
    assert_eq!(created.name, "Sheet 1-1");
    assert_eq!(
        editor
            .sheets()
            .unwrap()
            .into_iter()
            .map(|sheet| sheet.name)
            .collect::<Vec<_>>(),
        vec!["Sheet 1", "Sheet 1-1"]
    );

    let (_, _, source_sheet) = numbers_sheet(editor.package(), SOURCE_SHEET_ID).unwrap();
    let (_, _, copied_sheet) = numbers_sheet(editor.package(), created.object_id).unwrap();
    assert_eq!(source_sheet.drawable_infos.len(), 2);
    assert_eq!(copied_sheet.drawable_infos.len(), 2);
    assert_ne!(
        source_sheet.drawable_infos[0].identifier,
        copied_sheet.drawable_infos[0].identifier
    );
    assert_ne!(
        source_sheet.drawable_infos[1].identifier,
        copied_sheet.drawable_infos[1].identifier
    );

    let copied_table_id = editor
        .tables()
        .unwrap()
        .into_iter()
        .find(|table| table.object_id != SOURCE_TABLE_ID)
        .unwrap()
        .object_id;
    assert_eq!(
        find_table_owner(editor.package(), copied_table_id)
            .unwrap()
            .sheet_id,
        created.object_id
    );
    let copied_text_box = editor
        .sheet_text_boxes(created.object_id)
        .unwrap()
        .remove(0);
    assert_ne!(copied_text_box.drawable_object_id, SOURCE_TEXT_BOX_ID);
    assert_eq!(copied_text_box.storage.storage.text(), "Source");

    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        copied_table_id,
        SOURCE_CELL_ROW,
        SOURCE_CELL_COLUMN,
        CellValue::Text("Copied cell".to_owned()),
    )
    .unwrap();
    editor
        .set_sheet_text_box_text(
            created.object_id,
            copied_text_box.drawable_object_id,
            "Copied box",
        )
        .unwrap();
    assert_eq!(
        editor.sheet_text_boxes(SOURCE_SHEET_ID).unwrap()[0]
            .storage
            .storage
            .text(),
        "Source"
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets();
    assert_eq!(
        sheets[0]
            .tables()
            .next()
            .unwrap()
            .get_cell(SOURCE_CELL_ROW, SOURCE_CELL_COLUMN),
        Some(&CellValue::Text("Original".to_owned()))
    );
    assert_eq!(
        sheets[1]
            .tables()
            .next()
            .unwrap()
            .get_cell(SOURCE_CELL_ROW, SOURCE_CELL_COLUMN),
        Some(&CellValue::Text("Copied cell".to_owned()))
    );
}

#[test]
fn unsupported_sheet_drawable_duplicate_fails_transactionally() {
    const SOURCE_SHEET_ID: u64 = 2;
    const UNSUPPORTED_DRAWABLE_ID: u64 = 90;
    const SHEET_ARCHIVE_MESSAGE_TYPE: u32 = 2;

    let mut package = test_package_with_text_box();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let sheet = archive.object_mut(SOURCE_SHEET_ID).unwrap();
            let mut decoded = tn::SheetArchive::decode(sheet.messages[0].data.as_slice())?;
            decoded.drawable_infos.push(Reference {
                identifier: UNSUPPORTED_DRAWABLE_ID,
                ..Default::default()
            });
            sheet.replace_message(
                0,
                RawMessage {
                    type_: SHEET_ARCHIVE_MESSAGE_TYPE,
                    data: decoded.encode_to_vec(),
                },
            )?;
            archive.insert_object(ArchiveObject::new(
                UNSUPPORTED_DRAWABLE_ID,
                vec![RawMessage {
                    type_: SHAPE_INFO_MESSAGE_TYPE,
                    data: tswp::ShapeInfoArchive {
                        is_text_box: Some(false),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .duplicate_sheet(test_sheet_selector(&editor, SOURCE_SHEET_ID))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn cross_table_dependency_sheet_duplicate_fails_transactionally() {
    const SOURCE_SHEET_ID: u64 = 2;
    const FORMULA_OWNER_OBJECT_ID: u64 = 101;
    const EXTERNAL_OWNER_ID: u32 = 777;
    const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;

    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let object = archive.object_mut(FORMULA_OWNER_OBJECT_ID).unwrap();
            let mut owner =
                tsce::FormulaOwnerDependenciesArchive::decode(object.messages[0].data.as_slice())?;
            owner
                .cell_dependencies
                .get_or_insert_default()
                .cell_record
                .push(tsce::CellRecordExpandedArchive {
                    column: 0,
                    row: 0,
                    expanded_edges: Some(tsce::ExpandedEdgesArchive {
                        edge_with_owner_rows: vec![0],
                        edge_with_owner_columns: vec![0],
                        internal_owner_id_for_edge: vec![EXTERNAL_OWNER_ID],
                        ..Default::default()
                    }),
                    ..Default::default()
                });
            object.replace_message(
                0,
                RawMessage {
                    type_: FORMULA_OWNER_MESSAGE_TYPE,
                    data: owner.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .duplicate_sheet(test_sheet_selector(&editor, SOURCE_SHEET_ID))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn malformed_text_box_ownership_and_external_references_fail_transactionally() {
    let mut duplicate_owner = test_package_with_text_box();
    duplicate_owner
        .update_archive("Index/Document.iwa", |archive| {
            let sheet = archive.object_mut(2).unwrap();
            let mut decoded = tn::SheetArchive::decode(sheet.messages[0].data.as_slice())?;
            decoded.drawable_infos.push(Reference {
                identifier: 50,
                ..Default::default()
            });
            sheet.replace_message(
                0,
                RawMessage {
                    type_: 2,
                    data: decoded.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(duplicate_owner).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_sheet_text_box_text(2, 50, "rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut externally_referenced = test_package_with_text_box();
    externally_referenced
        .update_archive("Index/Document.iwa", |archive| {
            let mut owner = ArchiveObject::new(
                60,
                vec![RawMessage {
                    type_: 999,
                    data: Vec::new(),
                }],
            )?;
            owner.archive_info.message_infos[0]
                .object_references
                .push(51);
            Ok(archive.insert_object(owner)?)
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(externally_referenced).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_sheet_text_box(2, 50).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn text_box_duplicate_tracks_document_component_uuids_and_highwater() {
    let mut editor = NumbersEditor::from_package(test_package_with_text_box_metadata()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let created = editor
        .duplicate_sheet_text_box(2, 50, "Metadata clone")
        .unwrap();
    let graph = numbers_text_box_graph(editor.package(), 2, created.drawable_object_id).unwrap();
    assert_eq!(graph.object_ids, graph.uuid_object_ids);
    assert_eq!(
        crate::package_metadata::package_last_object_identifier(editor.package()).unwrap(),
        graph.object_ids.iter().copied().max()
    );
    editor
        .remove_sheet_text_box(2, created.drawable_object_id)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn table_data_list_wire_mutations_preserve_unknown_fields_and_restore_exactly() {
    let baseline = TableDataList {
        list_type: tst::table_data_list::ListType::String as i32,
        next_list_id: 2,
        entries: vec![tst::table_data_list::ListEntry {
            key: 1,
            refcount: 2,
            string: Some("first".to_owned()),
            ..Default::default()
        }],
        segments: vec![Reference {
            identifier: 100,
            ..Default::default()
        }],
        is_new_for_bnc: Some(true),
    };
    let mut original = crate::wire::transform_length_delimited_fields_at_path(
        &baseline.encode_to_vec(),
        &[3],
        |entry| {
            let mut entry = entry.to_vec();
            append_unknown_varint(&mut entry, 98, 980);
            Ok(entry)
        },
    )
    .unwrap();
    original =
        crate::wire::transform_length_delimited_fields_at_path(&original, &[4], |reference| {
            let mut reference = reference.to_vec();
            append_unknown_varint(&mut reference, 97, 970);
            Ok(reference)
        })
        .unwrap();
    append_unknown_varint(&mut original, 99, 990);
    let previous = TableDataList::decode(original.as_slice()).unwrap();
    let mut current = previous.clone();
    current.next_list_id = 3;
    current.entries[0].refcount = 3;
    current.entries.push(tst::table_data_list::ListEntry {
        key: 2,
        refcount: 1,
        string: Some("second".to_owned()),
        ..Default::default()
    });
    current.segments.push(Reference {
        identifier: 101,
        ..Default::default()
    });

    let changed = rewrite_table_data_list_wire(&original, &previous, &current).unwrap();
    assert_eq!(TableDataList::decode(changed.as_slice()).unwrap(), current);
    let restored = rewrite_table_data_list_wire(&changed, &current, &previous).unwrap();
    assert_eq!(restored, original);
}

#[test]
fn table_data_list_segments_preserve_unknown_fields_and_reject_duplicates() {
    let baseline = TableDataListSegment {
        list_type: tst::table_data_list::ListType::String as i32,
        key_range: crate::protobuf::tsp::Range {
            location: 1,
            length: 1,
        },
        entries: vec![tst::table_data_list::ListEntry {
            key: 1,
            refcount: 1,
            string: Some("value".to_owned()),
            ..Default::default()
        }],
    };
    let mut original = crate::wire::transform_length_delimited_fields_at_path(
        &baseline.encode_to_vec(),
        &[2],
        |range| {
            let mut range = range.to_vec();
            append_unknown_varint(&mut range, 98, 980);
            Ok(range)
        },
    )
    .unwrap();
    original = crate::wire::transform_length_delimited_fields_at_path(&original, &[3], |entry| {
        let mut entry = entry.to_vec();
        append_unknown_varint(&mut entry, 97, 970);
        Ok(entry)
    })
    .unwrap();
    append_unknown_varint(&mut original, 99, 990);
    let previous = TableDataListSegment::decode(original.as_slice()).unwrap();
    let mut current = previous.clone();
    current.key_range.length = 2;
    current.entries[0].refcount = 2;
    let changed = rewrite_table_data_list_segment_wire(&original, &previous, &current).unwrap();
    let restored = rewrite_table_data_list_segment_wire(&changed, &current, &previous).unwrap();
    assert_eq!(restored, original);

    let duplicate =
        crate::wire::transform_length_delimited_fields_at_path(&original, &[3], |entry| {
            let mut entry = entry.to_vec();
            entry.extend(litchi_iwa_common::varint::encode_varint(16));
            entry.extend(litchi_iwa_common::varint::encode_varint(1));
            Ok(entry)
        })
        .unwrap();
    let duplicate_previous = TableDataListSegment::decode(duplicate.as_slice()).unwrap();
    let mut duplicate_current = duplicate_previous.clone();
    duplicate_current.entries[0].refcount = 2;
    assert!(
        rewrite_table_data_list_segment_wire(&duplicate, &duplicate_previous, &duplicate_current)
            .is_err()
    );
}

#[test]
fn tile_wire_mutations_preserve_unknown_rows_and_restore_exactly() {
    let baseline = Tile {
        max_column: 2,
        max_row: 1,
        num_cells: 1,
        numrows: 1,
        row_infos: vec![TileRowInfo {
            tile_row_index: 0,
            cell_count: 1,
            cell_storage_buffer_pre_bnc: Vec::new(),
            cell_offsets_pre_bnc: Vec::new(),
            storage_version: Some(5),
            cell_storage_buffer: Some(vec![1, 2, 3]),
            cell_offsets: Some(vec![0, 0]),
            has_wide_offsets: None,
        }],
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        should_use_wide_rows: None,
    };
    let mut original = crate::wire::transform_length_delimited_fields_at_path(
        &baseline.encode_to_vec(),
        &[5],
        |row| {
            let mut row = row.to_vec();
            append_unknown_varint(&mut row, 98, 980);
            Ok(row)
        },
    )
    .unwrap();
    append_unknown_varint(&mut original, 99, 990);
    let previous = Tile::decode(original.as_slice()).unwrap();
    let mut current = previous.clone();
    current.numrows = 2;
    current.row_infos[0].cell_storage_buffer = Some(vec![4, 5, 6]);
    current.row_infos.push(TileRowInfo {
        tile_row_index: 1,
        cell_count: 1,
        cell_storage_buffer_pre_bnc: Vec::new(),
        cell_offsets_pre_bnc: Vec::new(),
        storage_version: Some(5),
        cell_storage_buffer: Some(vec![7, 8]),
        cell_offsets: Some(vec![0, 0]),
        has_wide_offsets: Some(false),
    });

    let changed = rewrite_tile_wire(&original, &previous, &current).unwrap();
    assert_eq!(Tile::decode(changed.as_slice()).unwrap(), current);
    let restored = rewrite_tile_wire(&changed, &current, &previous).unwrap();
    assert_eq!(restored, original);

    let duplicate =
        crate::wire::transform_length_delimited_fields_at_path(&original, &[5], |row| {
            let mut row = row.to_vec();
            row.extend(litchi_iwa_common::varint::encode_varint(8));
            row.extend(litchi_iwa_common::varint::encode_varint(0));
            Ok(row)
        })
        .unwrap();
    let duplicate_previous = Tile::decode(duplicate.as_slice()).unwrap();
    let mut duplicate_current = duplicate_previous.clone();
    duplicate_current.row_infos[0].cell_count = 2;
    assert!(rewrite_tile_wire(&duplicate, &duplicate_previous, &duplicate_current).is_err());
}

#[test]
fn header_bucket_wire_mutations_preserve_unknown_entries_and_restore_exactly() {
    let baseline = tst::HeaderStorageBucket {
        bucket_hash_function: 1,
        headers: vec![tst::header_storage_bucket::Header {
            index: 0,
            size: 20.0,
            hiding_state: 0,
            number_of_cells: 1,
            cell_style: Some(Reference {
                identifier: 10,
                ..Default::default()
            }),
            text_style: None,
        }],
    };
    let mut original = crate::wire::transform_length_delimited_fields_at_path(
        &baseline.encode_to_vec(),
        &[2],
        |header| {
            let mut header = header.to_vec();
            append_unknown_varint(&mut header, 98, 980);
            Ok(header)
        },
    )
    .unwrap();
    original =
        crate::wire::transform_length_delimited_fields_at_path(&original, &[2, 5], |reference| {
            let mut reference = reference.to_vec();
            append_unknown_varint(&mut reference, 97, 970);
            Ok(reference)
        })
        .unwrap();
    append_unknown_varint(&mut original, 99, 990);
    let previous = tst::HeaderStorageBucket::decode(original.as_slice()).unwrap();
    let mut current = previous.clone();
    current.headers[0].number_of_cells = 2;
    current.headers.push(tst::header_storage_bucket::Header {
        index: 1,
        size: 0.0,
        hiding_state: 0,
        number_of_cells: 1,
        cell_style: None,
        text_style: None,
    });

    let changed = rewrite_header_bucket_wire(&original, &previous, &current).unwrap();
    let restored = rewrite_header_bucket_wire(&changed, &current, &previous).unwrap();
    assert_eq!(restored, original);

    let duplicate =
        crate::wire::transform_length_delimited_fields_at_path(&original, &[2], |header| {
            let mut header = header.to_vec();
            header.extend(litchi_iwa_common::varint::encode_varint(8));
            header.extend(litchi_iwa_common::varint::encode_varint(0));
            Ok(header)
        })
        .unwrap();
    let duplicate_previous = tst::HeaderStorageBucket::decode(duplicate.as_slice()).unwrap();
    let mut duplicate_current = duplicate_previous.clone();
    duplicate_current.headers[0].number_of_cells = 2;
    assert!(
        rewrite_header_bucket_wire(&duplicate, &duplicate_previous, &duplicate_current).is_err()
    );
}

#[test]
fn uid_map_and_stroke_sidecar_mutations_preserve_deep_unknown_fields_exactly() {
    let uid_baseline = tst::ColumnRowUidMapArchive {
        sorted_column_uids: vec![crate::protobuf::tsp::Uuid { lower: 1, upper: 2 }],
        column_index_for_uid: vec![0],
        column_uid_for_index: vec![0],
        sorted_row_uids: vec![crate::protobuf::tsp::Uuid { lower: 3, upper: 4 }],
        row_index_for_uid: vec![0],
        row_uid_for_index: vec![0],
    };
    let mut uid_original = crate::wire::transform_length_delimited_fields_at_path(
        &uid_baseline.encode_to_vec(),
        &[1],
        |uuid| {
            let mut uuid = uuid.to_vec();
            append_unknown_varint(&mut uuid, 98, 980);
            Ok(uuid)
        },
    )
    .unwrap();
    uid_original =
        crate::wire::transform_length_delimited_fields_at_path(&uid_original, &[4], |uuid| {
            let mut uuid = uuid.to_vec();
            append_unknown_varint(&mut uuid, 97, 970);
            Ok(uuid)
        })
        .unwrap();
    append_unknown_varint(&mut uid_original, 99, 990);
    let uid_previous = tst::ColumnRowUidMapArchive::decode(uid_original.as_slice()).unwrap();
    let mut uid_current = uid_previous.clone();
    uid_current
        .sorted_column_uids
        .push(crate::protobuf::tsp::Uuid { lower: 5, upper: 6 });
    uid_current.column_index_for_uid.push(1);
    uid_current.column_uid_for_index.push(1);
    uid_current
        .sorted_row_uids
        .push(crate::protobuf::tsp::Uuid { lower: 7, upper: 8 });
    uid_current.row_index_for_uid.push(1);
    uid_current.row_uid_for_index.push(1);
    let uid_changed = rewrite_uid_map_wire(&uid_original, &uid_previous, &uid_current).unwrap();
    let uid_restored = rewrite_uid_map_wire(&uid_changed, &uid_current, &uid_previous).unwrap();
    assert_eq!(uid_restored, uid_original);

    let sidecar_baseline = tst::StrokeSidecarArchive {
        max_order: Some(7),
        column_count: Some(4),
        row_count: Some(5),
        left_column_stroke_layers: vec![Reference {
            identifier: 100,
            ..Default::default()
        }],
        ..Default::default()
    };
    let mut sidecar_original = crate::wire::transform_length_delimited_fields_at_path(
        &sidecar_baseline.encode_to_vec(),
        &[4],
        |reference| {
            let mut reference = reference.to_vec();
            append_unknown_varint(&mut reference, 98, 980);
            Ok(reference)
        },
    )
    .unwrap();
    append_unknown_varint(&mut sidecar_original, 99, 990);
    let sidecar_previous = tst::StrokeSidecarArchive::decode(sidecar_original.as_slice()).unwrap();
    let mut sidecar_current = sidecar_previous.clone();
    sidecar_current.column_count = Some(6);
    sidecar_current.row_count = Some(7);
    let sidecar_changed =
        rewrite_stroke_sidecar_wire(&sidecar_original, &sidecar_previous, &sidecar_current)
            .unwrap();
    let sidecar_restored =
        rewrite_stroke_sidecar_wire(&sidecar_changed, &sidecar_current, &sidecar_previous).unwrap();
    assert_eq!(sidecar_restored, sidecar_original);
}

#[test]
fn comment_table_reference_mutation_preserves_nested_unknown_fields_exactly() {
    let baseline = TableModelArchive {
        table_id: "table".to_owned(),
        table_name: "Table 1".to_owned(),
        number_of_rows: 1,
        number_of_columns: 1,
        base_data_store: tst::DataStore::default(),
        ..Default::default()
    };
    let mut original = crate::wire::transform_length_delimited_fields_at_path(
        &baseline.encode_to_vec(),
        &[4],
        |store| {
            let mut store = store.to_vec();
            append_unknown_varint(&mut store, 98, 980);
            Ok(store)
        },
    )
    .unwrap();
    append_unknown_varint(&mut original, 99, 990);
    let previous = TableModelArchive::decode(original.as_slice()).unwrap();
    let mut current = previous.clone();
    current.base_data_store.comment_storage_table = Some(Reference {
        identifier: 123,
        ..Default::default()
    });

    let changed = rewrite_table_model_comment_table_wire(&original, &previous, &current).unwrap();
    let restored = rewrite_table_model_comment_table_wire(&changed, &current, &previous).unwrap();
    assert_eq!(restored, original);
}

#[test]
fn rich_text_table_reference_mutation_is_strict_and_preserves_unknown_fields() {
    let baseline = TableModelArchive {
        table_id: "table".to_owned(),
        table_name: "Table 1".to_owned(),
        number_of_rows: 1,
        number_of_columns: 1,
        base_data_store: tst::DataStore::default(),
        ..Default::default()
    };
    let mut original = crate::wire::transform_length_delimited_fields_at_path(
        &baseline.encode_to_vec(),
        &[4],
        |store| {
            let mut store = store.to_vec();
            append_unknown_varint(&mut store, 98, 980);
            Ok(store)
        },
    )
    .unwrap();
    append_unknown_varint(&mut original, 99, 990);
    let previous = TableModelArchive::decode(original.as_slice()).unwrap();
    let mut current = previous.clone();
    current.base_data_store.rich_text_table = Some(Reference {
        identifier: 123,
        ..Default::default()
    });

    let changed = rewrite_table_model_rich_text_table_wire(&original, &previous, &current).unwrap();
    let restored = rewrite_table_model_rich_text_table_wire(&changed, &current, &previous).unwrap();
    assert_eq!(restored, original);

    let mut malformed = current;
    malformed.number_of_rows = 2;
    assert!(rewrite_table_model_rich_text_table_wire(&original, &previous, &malformed).is_err());
}

#[test]
fn formula_dependency_wire_mutations_preserve_deep_unknown_fields_exactly() {
    let record = tsce::CellRecordExpandedArchive {
        column: 1,
        row: 2,
        expanded_edges: Some(tsce::ExpandedEdgesArchive {
            edge_without_owner_rows: vec![0],
            edge_without_owner_columns: vec![0],
            ..Default::default()
        }),
        ..Default::default()
    };
    let owner_baseline = tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid: crate::protobuf::tsp::Uuid { lower: 1, upper: 2 },
        internal_formula_owner_id: 3,
        cell_dependencies: Some(tsce::CellDependenciesExpandedArchive {
            cell_record: vec![record.clone()],
        }),
        tiled_cell_dependencies: Some(tsce::CellDependenciesTiledArchive {
            cell_record_tiles: vec![Reference {
                identifier: 100,
                ..Default::default()
            }],
        }),
        formula_owner: Some(Reference {
            identifier: 50,
            ..Default::default()
        }),
        ..Default::default()
    };
    let mut owner_original = crate::wire::transform_length_delimited_fields_at_path(
        &owner_baseline.encode_to_vec(),
        &[4, 1],
        |record| {
            let mut record = record.to_vec();
            append_unknown_varint(&mut record, 97, 970);
            Ok(record)
        },
    )
    .unwrap();
    owner_original = crate::wire::transform_length_delimited_fields_at_path(
        &owner_original,
        &[4],
        |dependencies| {
            let mut dependencies = dependencies.to_vec();
            append_unknown_varint(&mut dependencies, 98, 980);
            Ok(dependencies)
        },
    )
    .unwrap();
    owner_original = crate::wire::transform_length_delimited_fields_at_path(
        &owner_original,
        &[13, 1],
        |reference| {
            let mut reference = reference.to_vec();
            append_unknown_varint(&mut reference, 95, 950);
            Ok(reference)
        },
    )
    .unwrap();
    owner_original = crate::wire::transform_length_delimited_fields_at_path(
        &owner_original,
        &[13],
        |dependencies| {
            let mut dependencies = dependencies.to_vec();
            append_unknown_varint(&mut dependencies, 96, 960);
            Ok(dependencies)
        },
    )
    .unwrap();
    append_unknown_varint(&mut owner_original, 99, 990);
    let owner_previous =
        tsce::FormulaOwnerDependenciesArchive::decode(owner_original.as_slice()).unwrap();
    let mut owner_current = owner_previous.clone();
    owner_current
        .cell_dependencies
        .as_mut()
        .unwrap()
        .cell_record
        .push(tsce::CellRecordExpandedArchive {
            column: 3,
            row: 4,
            ..Default::default()
        });
    owner_current
        .tiled_cell_dependencies
        .as_mut()
        .unwrap()
        .cell_record_tiles
        .push(Reference {
            identifier: 101,
            ..Default::default()
        });
    let owner_changed =
        rewrite_formula_owner_dependencies_wire(&owner_original, &owner_previous, &owner_current)
            .unwrap();
    let owner_restored =
        rewrite_formula_owner_dependencies_wire(&owner_changed, &owner_current, &owner_previous)
            .unwrap();
    assert_eq!(owner_restored, owner_original);

    let tile_baseline = tsce::CellRecordTileArchive {
        internal_owner_id: 3,
        tile_column_begin: 0,
        tile_row_begin: 0,
        cell_records: vec![record],
    };
    let mut tile_original = crate::wire::transform_length_delimited_fields_at_path(
        &tile_baseline.encode_to_vec(),
        &[4],
        |record| {
            let mut record = record.to_vec();
            append_unknown_varint(&mut record, 98, 980);
            Ok(record)
        },
    )
    .unwrap();
    append_unknown_varint(&mut tile_original, 99, 990);
    let tile_previous = tsce::CellRecordTileArchive::decode(tile_original.as_slice()).unwrap();
    let mut tile_current = tile_previous.clone();
    tile_current
        .cell_records
        .push(tsce::CellRecordExpandedArchive {
            column: 5,
            row: 6,
            ..Default::default()
        });
    let tile_changed =
        rewrite_dependency_tile_wire(&tile_original, &tile_previous, &tile_current).unwrap();
    let tile_restored =
        rewrite_dependency_tile_wire(&tile_changed, &tile_current, &tile_previous).unwrap();
    assert_eq!(tile_restored, tile_original);

    let engine_baseline = tsce::CalculationEngineArchive {
        dependency_tracker: tsce::DependencyTrackerArchive {
            number_of_formulas: Some(0),
            ..Default::default()
        },
        ..Default::default()
    };
    let mut engine_original = crate::wire::transform_length_delimited_fields_at_path(
        &engine_baseline.encode_to_vec(),
        &[2],
        |tracker| {
            let mut tracker = tracker.to_vec();
            append_unknown_varint(&mut tracker, 98, 980);
            Ok(tracker)
        },
    )
    .unwrap();
    append_unknown_varint(&mut engine_original, 99, 990);
    let engine_previous =
        tsce::CalculationEngineArchive::decode(engine_original.as_slice()).unwrap();
    let mut engine_current = engine_previous.clone();
    engine_current.dependency_tracker.number_of_formulas = Some(1);
    let engine_changed = rewrite_calculation_engine_formula_count_wire(
        &engine_original,
        &engine_previous,
        &engine_current,
    )
    .unwrap();
    let engine_restored = rewrite_calculation_engine_formula_count_wire(
        &engine_changed,
        &engine_current,
        &engine_previous,
    )
    .unwrap();
    assert_eq!(engine_restored, engine_original);
}

#[test]
fn sparse_row_round_trip_and_wide_promotion() {
    let mut cells = vec![None; 4];
    cells[1] = Some(vec![5; 24]);
    cells[3] = Some(vec![6; 30]);
    let (storage, offsets, wide) = encode_row(&cells, false).unwrap();
    let row = TileRowInfo {
        tile_row_index: 0,
        cell_count: 2,
        cell_storage_buffer_pre_bnc: Vec::new(),
        cell_offsets_pre_bnc: Vec::new(),
        storage_version: Some(5),
        cell_storage_buffer: Some(storage),
        cell_offsets: Some(offsets),
        has_wide_offsets: Some(wide),
    };
    assert_eq!(split_row(&row).unwrap(), cells);

    let huge = vec![Some(vec![0; 70_000]), Some(vec![1; 12])];
    let (_, _, wide) = encode_row(&huge, false)
        .or_else(|_| encode_row(&huge, true))
        .unwrap();
    assert!(wide);
}

#[test]
fn edit_promotes_a_complete_legacy_tile_mirror_to_bnc() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(30).unwrap();
            let message_type = object.messages[0].type_;
            let mut tile = Tile::decode(object.messages[0].data.as_slice())?;
            tile.last_saved_in_bnc = Some(false);
            tile.row_infos[0].cell_storage_buffer_pre_bnc = vec![1, 2, 3];
            tile.row_infos[0].cell_offsets_pre_bnc = vec![4, 5];
            object.replace_message(
                0,
                RawMessage {
                    type_: message_type,
                    data: tile.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 0, 1, cell_number(7.5)).unwrap();
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let tile = Tile::decode(archive.object(30).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(tile.last_saved_in_bnc, Some(true));
    assert_eq!(tile.storage_version, Some(5));
    assert!(tile.row_infos.iter().all(|row| {
        row.storage_version == Some(5)
            && row.cell_storage_buffer.is_some()
            && row.cell_offsets.is_some()
            && row.cell_storage_buffer_pre_bnc.is_empty()
            && row.cell_offsets_pre_bnc.is_empty()
    }));
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0]
            .tables()
            .next()
            .unwrap()
            .get_cell(0, 1)
            .unwrap()
            .as_number(),
        Some(7.5)
    );
}

#[test]
fn segmented_formula_entries_are_reused_and_released() {
    let expression = FormulaExpression::function(
        "SUM",
        [
            FormulaExpression::Number(1.0),
            FormulaExpression::Number(2.0),
        ],
    );
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    editor.set_formula(10, 0, 0, expression.clone()).unwrap();
    let mut package = editor.into_package();
    move_table_data_list_entries_to_segment(&mut package, 21, 61);

    let document = compatibility_document_from_bytes(&package.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(0, 0),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor.set_formula(10, 1, 0, expression).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 0, 0, cell_number(7.0)).unwrap();

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let segment =
        TableDataListSegment::decode(archive.object(61).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(segment.entries.len(), 1);
    assert_eq!(segment.entries[0].refcount, 1);
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(table.get_cell(0, 0).unwrap().as_number(), Some(7.0));
    assert_eq!(
        table.get_cell(1, 0),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
}

#[test]
fn formula_error_cells_release_root_and_segmented_list_entries() {
    for segmented in [false, true] {
        let mut package = test_package_with_formula_error();
        if segmented {
            move_table_data_list_entries_to_segment(&mut package, 22, 60);
        }
        let before = compatibility_document_from_bytes(&package.to_bytes().unwrap()).unwrap();
        assert_eq!(
            before.sheets()[0].tables().next().unwrap().get_cell(0, 1),
            Some(&CellValue::Error("Syntax Error".to_owned()))
        );

        let mut editor = NumbersEditor::from_package(package).unwrap();
        crate::numbers::editor::set_cell_fixture(&mut editor, 10, 0, 1, cell_number(4.5)).unwrap();
        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let errors =
            TableDataList::decode(archive.object(22).unwrap().messages[0].data.as_slice()).unwrap();
        assert!(errors.entries.is_empty());
        assert!(errors.segments.is_empty());
        if segmented {
            assert!(archive.object(60).is_none());
        }
        let tile = Tile::decode(archive.object(30).unwrap().messages[0].data.as_slice()).unwrap();
        let cells = split_row(&tile.row_infos[0]).unwrap();
        let cell = BncCell::parse(cells[1].as_deref().unwrap()).unwrap();
        assert_eq!(cell.formula_error_identifier(), None);
        assert_eq!(cell.stored_value(), StoredValue::Number);
        let after = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            after.sheets()[0]
                .tables()
                .next()
                .unwrap()
                .get_cell(0, 1)
                .unwrap()
                .as_number(),
            Some(4.5)
        );
    }
}

#[test]
fn cell_comment_crud_preserves_value_and_comment_metadata() {
    let mut editor = NumbersEditor::from_package(test_package_with_comments(false)).unwrap();
    let original = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    assert_eq!(original.comment.text, "Original comment");
    assert_eq!(original.comment.creation_date_seconds, Some(123.5));
    assert_eq!(
        original.comment.reply_ids.as_ref(),
        [StorageId::new(70).unwrap()]
    );
    assert_eq!(original.comment.storage_uuid.unwrap().lower(), 61);
    let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let reader_comment = reopened.cell_comment(10, 0, 1).unwrap().unwrap();
    assert_eq!(reader_comment.comment, original.comment);

    editor
        .set_cell_comment(10, 0, 1, "Updated comment")
        .unwrap();
    let updated = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    assert_eq!(updated.storage_id, original.storage_id);
    assert_eq!(updated.comment.text, "Updated comment");
    assert_eq!(updated.comment.creation_date_seconds, Some(123.5));
    assert_eq!(updated.comment.storage_uuid, original.comment.storage_uuid);

    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 0, 1, cell_number(8.5)).unwrap();
    assert_eq!(
        editor.cell_comment(10, 0, 1).unwrap().unwrap().comment.text,
        "Updated comment"
    );
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 0, 1, CellValue::Empty).unwrap();
    assert!(editor.cell_comment(10, 0, 1).unwrap().is_some());

    editor.clear_cell_comment(10, 0, 1).unwrap();
    assert!(editor.cell_comment(10, 0, 1).unwrap().is_none());
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    assert!(archive.object(61).is_none());
    assert!(archive.object(70).is_none());
    let list =
        TableDataList::decode(archive.object(60).unwrap().messages[0].data.as_slice()).unwrap();
    assert!(list.entries.is_empty());

    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(table.get_cell(0, 1), Some(&CellValue::Empty));
    let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(reopened.cell_comment(10, 0, 1).unwrap().is_none());
}

#[test]
fn cell_comment_reply_crud_is_copy_on_write_and_transactional() {
    let mut editor = NumbersEditor::from_package(test_package_with_comments(false)).unwrap();
    let original_root = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    let original_replies = editor.cell_comment_replies(10, 0, 1).unwrap();
    assert_eq!(original_replies.len(), 1);
    assert_eq!(original_replies[0].comment.text, "Reply");

    let added_id = editor
        .add_cell_comment_reply(10, 0, 1, "Second reply")
        .unwrap();
    let after_add = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    assert_ne!(after_add.storage_id, original_root.storage_id);
    assert_eq!(
        after_add.comment.storage_uuid,
        original_root.comment.storage_uuid
    );
    assert_eq!(
        after_add.comment.reply_ids.as_ref(),
        [
            StorageId::new(70).unwrap(),
            StorageId::new(added_id).unwrap()
        ]
    );
    let replies = editor.cell_comment_replies(10, 0, 1).unwrap();
    assert_eq!(replies[1].comment.text, "Second reply");
    assert!(replies[1].comment.creation_date_seconds.is_some());
    assert!(replies[1].comment.storage_uuid.is_some());

    let stable = editor.to_bytes().unwrap();
    assert_eq!(
        editor
            .set_cell_comment_reply(10, 0, 1, added_id, "Second reply")
            .unwrap(),
        added_id
    );
    assert_eq!(editor.to_bytes().unwrap(), stable);

    let updated_id = editor
        .set_cell_comment_reply(10, 0, 1, added_id, "Updated reply")
        .unwrap();
    assert_ne!(updated_id, added_id);
    let replies = editor.cell_comment_replies(10, 0, 1).unwrap();
    assert_eq!(
        replies
            .iter()
            .map(|reply| reply.storage_id.get())
            .collect::<Vec<_>>(),
        [70, updated_id]
    );
    assert_eq!(replies[1].comment.text, "Updated reply");
    assert!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .object(added_id)
            .is_none()
    );

    editor
        .remove_cell_comment_reply(10, 0, 1, updated_id)
        .unwrap();
    let replies = editor.cell_comment_replies(10, 0, 1).unwrap();
    assert_eq!(replies.len(), 1);
    assert_eq!(replies[0].storage_id.get(), 70);
    assert_eq!(
        editor.cell_comment(10, 0, 1).unwrap().unwrap().comment.text,
        "Original comment"
    );

    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_cell_comment_reply(10, 0, 1, 999_999).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);

    let reparsed = NumbersEditor::from_bytes(&before).unwrap();
    assert_eq!(
        reparsed.cell_comment_replies(10, 0, 1).unwrap()[0]
            .comment
            .text,
        "Reply"
    );
}

#[test]
fn shared_segmented_comments_use_copy_on_write_and_cleanup() {
    let mut package = test_package_with_comments(true);
    move_table_data_list_entries_to_segment(&mut package, 60, 62);
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let original_storage = editor
        .cell_comment(10, 0, 1)
        .unwrap()
        .unwrap()
        .storage_id
        .get();

    editor
        .set_cell_comment(10, 0, 1, "Independent comment")
        .unwrap();
    let first = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    let second = editor.cell_comment(10, 0, 2).unwrap().unwrap();
    assert_ne!(first.storage_id.get(), original_storage);
    assert_eq!(second.storage_id.get(), original_storage);
    assert_eq!(first.comment.text, "Independent comment");
    assert_eq!(second.comment.text, "Original comment");

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let segment =
        TableDataListSegment::decode(archive.object(62).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(segment.entries[0].refcount, 1);
    let root =
        TableDataList::decode(archive.object(60).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(root.entries.len(), 1);

    editor.clear_cell_comment(10, 0, 2).unwrap();
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let root =
        TableDataList::decode(archive.object(60).unwrap().messages[0].data.as_slice()).unwrap();
    assert!(root.segments.is_empty());
    assert!(archive.object(62).is_none());
    assert!(archive.object(original_storage).is_none());
    assert!(archive.object(70).is_some());
    assert_eq!(
        editor.cell_comment(10, 0, 1).unwrap().unwrap().comment.text,
        "Independent comment"
    );
}

#[test]
fn comment_updates_and_copy_on_write_preserve_unknown_storage_fields() {
    let mut package = test_package_with_comments(false);
    let unknown = add_unknown_comment_storage_field(&mut package, 61);
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .set_cell_comment(10, 0, 1, "Wire-safe update")
        .unwrap();
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    assert!(
        archive.object(61).unwrap().messages[0]
            .data
            .ends_with(&unknown)
    );

    let mut package = test_package_with_comments(true);
    add_unknown_comment_storage_field(&mut package, 61);
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor.set_cell_comment(10, 0, 1, "Wire-safe copy").unwrap();
    let cloned = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    assert_ne!(cloned.storage_id.get(), 61);
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    assert!(
        archive.object(cloned.storage_id.get()).unwrap().messages[0]
            .data
            .ends_with(&unknown)
    );
    assert!(
        archive.object(61).unwrap().messages[0]
            .data
            .ends_with(&unknown)
    );
}

#[test]
fn creates_comment_table_and_comment_only_cell_when_missing() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    editor
        .set_cell_comment(10, 1, 2, "Created comment")
        .unwrap();
    let info = editor.cell_comment(10, 1, 2).unwrap().unwrap();
    assert_eq!(info.comment.text, "Created comment");
    assert_eq!(info.list_id.get(), 2);
    assert!(info.comment.creation_date_seconds.is_some());
    assert!(info.comment.storage_uuid.is_some());

    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(table.get_cell(1, 2), Some(&CellValue::Empty));
    let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .cell_comment(10, 1, 2)
            .unwrap()
            .unwrap()
            .comment
            .text,
        "Created comment"
    );

    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 2, cell_number(42.0)).unwrap();
    editor.clear_cell_comment(10, 1, 2).unwrap();
    assert!(editor.cell_comment(10, 1, 2).unwrap().is_none());
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0]
            .tables()
            .next()
            .unwrap()
            .get_cell(1, 2)
            .unwrap()
            .as_number(),
        Some(42.0)
    );
    assert!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .object(info.storage_id.get())
            .is_none()
    );
}

#[test]
fn empty_native_author_storage_supports_cell_comment_creation() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            Ok(archive.insert_object(ArchiveObject::new(
                90,
                vec![RawMessage {
                    type_: 213,
                    data: crate::protobuf::tsk::AnnotationAuthorStorageArchive::default()
                        .encode_to_vec(),
                }],
            )?)?)
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .set_cell_comment(10, 1, 2, "Generated local author")
        .unwrap();
    let comment = editor.cell_comment(10, 1, 2).unwrap().unwrap();
    assert_eq!(comment.comment.text, "Generated local author");
    assert!(comment.comment.author_id.is_some());

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let model =
        TableModelArchive::decode(archive.object(10).unwrap().messages[0].data.as_slice()).unwrap();
    let comment_table_id = model
        .base_data_store
        .comment_storage_table
        .unwrap()
        .identifier;
    let comment_table = archive.object(comment_table_id).unwrap();
    assert_eq!(
        comment_table.archive_info.message_infos[0].versions,
        [1, 0, 5]
    );
    let list = TableDataList::decode(comment_table.messages[0].data.as_slice()).unwrap();
    assert_eq!(list.is_new_for_bnc, None);
    assert_eq!(list.entries.len(), 1);
}

#[test]
fn malformed_comment_storage_fails_transactionally() {
    let mut package = test_package_with_comments(false);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(61).unwrap();
            object.archive_info.message_infos[0].type_ = 9999;
            object.messages[0].type_ = 9999;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_cell_comment(10, 0, 1, "Rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    // Semantic Numbers data is validated at ingress, so malformed native
    // comment storage is rejected before an archive-backed document exists.
    assert!(compatibility_document_from_bytes(&before).is_err());
}

#[test]
fn public_reader_applies_tile_row_origins() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let model_object = archive.object_mut(10).unwrap();
            let model_type = model_object.messages[0].type_;
            let mut model = TableModelArchive::decode(model_object.messages[0].data.as_slice())?;
            model.number_of_rows = 300;
            model
                .base_data_store
                .tiles
                .tiles
                .push(tst::tile_storage::Tile {
                    tileid: 1,
                    tile: Reference {
                        identifier: 43,
                        ..Default::default()
                    },
                });
            model_object.replace_message(
                0,
                RawMessage {
                    type_: model_type,
                    data: model.encode_to_vec(),
                },
            )?;

            let mut cell = BncCell::minimal();
            cell.set_number(99.0)?;
            let mut row = TileRowInfo {
                tile_row_index: 0,
                cell_count: 0,
                cell_storage_buffer_pre_bnc: Vec::new(),
                cell_offsets_pre_bnc: Vec::new(),
                storage_version: Some(5),
                cell_storage_buffer: Some(Vec::new()),
                cell_offsets: Some(Vec::new()),
                has_wide_offsets: Some(false),
            };
            rebuild_row(&mut row, &[Some(cell.encode())])?;
            archive.insert_object(ArchiveObject::new(
                43,
                vec![RawMessage {
                    type_: 6002,
                    data: Tile {
                        numrows: 1,
                        row_infos: vec![row],
                        storage_version: Some(5),
                        last_saved_in_bnc: Some(true),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();

    let document = compatibility_document_from_bytes(&package.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(table.get_cell(0, 1).unwrap().as_text(), "Original");
    assert_eq!(table.get_cell(256, 0).unwrap().as_number(), Some(99.0));
}

#[test]
fn formula_writes_intern_validate_and_release_references() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let expression = FormulaExpression::function(
        "SUM",
        [
            FormulaExpression::Number(1.0),
            FormulaExpression::Number(2.0),
        ],
    );

    editor
        .set_formula(table_id, 0, 0, expression.clone())
        .unwrap();
    editor
        .set_formula(table_id, 1, 0, expression.clone())
        .unwrap();

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let formula_object = archive.object(21).unwrap();
    let formulas = TableDataList::decode(formula_object.messages[0].data.as_slice()).unwrap();
    assert_eq!(formulas.entries.len(), 1);
    assert_eq!(formulas.entries[0].refcount, 2);

    let bytes = editor.to_bytes().unwrap();
    let document = compatibility_document_from_bytes(&bytes).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(0, 0),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    assert_eq!(table.get_cell(1, 0), table.get_cell(0, 0));

    crate::numbers::editor::set_cell_fixture(&mut editor, table_id, 0, 0, cell_number(42.0))
        .unwrap();
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let formulas =
        TableDataList::decode(archive.object(21).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(formulas.entries.len(), 1);
    assert_eq!(formulas.entries[0].refcount, 1);

    editor
        .set_formula(
            table_id,
            2,
            0,
            FormulaExpression::binary(
                super::super::FormulaBinaryOperator::GreaterThanOrEqual,
                FormulaExpression::Percent(Box::new(FormulaExpression::Number(50.0))),
                FormulaExpression::Number(0.5),
            ),
        )
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(2, 0),
        Some(&CellValue::Formula("=((50)%>=0.5)".to_owned()))
    );

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_formula(
                table_id,
                0,
                0,
                FormulaExpression::function("IF", [FormulaExpression::Boolean(true)]),
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn overwrites_formula_with_tiled_only_app_dependency_storage() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    editor
        .set_formula(
            table_id,
            0,
            0,
            FormulaExpression::function("SUM", [FormulaExpression::Number(1.0)]),
        )
        .unwrap();
    editor
        .package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let owner = archive.object_mut(101).unwrap();
            let mut dependencies =
                tsce::FormulaOwnerDependenciesArchive::decode(owner.messages[0].data.as_slice())?;
            dependencies.cell_dependencies = None;
            owner.replace_message(
                0,
                RawMessage {
                    type_: 4008,
                    data: dependencies.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();

    editor
        .set_formula(
            table_id,
            0,
            0,
            FormulaExpression::function("SUM", [FormulaExpression::Number(2.0)]),
        )
        .unwrap();
    let archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        archive.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert!(owner.cell_dependencies.is_none());
    let tile_id = owner.tiled_cell_dependencies.unwrap().cell_record_tiles[0].identifier;
    let tile = tsce::CellRecordTileArchive::decode(
        archive.object(tile_id).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(tile.cell_records.len(), 1);
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(0, 0),
        Some(&CellValue::Formula("=SUM(2)".to_owned()))
    );
}

#[test]
fn local_reference_formulas_write_exact_calculation_engine_edges() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let expression = FormulaExpression::function(
        "SUM",
        [FormulaExpression::range(
            crate::numbers::FormulaCellReference::relative(0, 0),
            crate::numbers::FormulaCellReference::mixed(1, 1, true, false),
        )],
    );

    editor.set_formula(table_id, 3, 2, expression).unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(3, 2),
        Some(&CellValue::Formula("=SUM(A1:B$2)".to_owned()))
    );

    let archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = archive.object(101).unwrap();
    let dependencies =
        tsce::FormulaOwnerDependenciesArchive::decode(owner.messages[0].data.as_slice()).unwrap();
    let record = &dependencies.cell_dependencies.as_ref().unwrap().cell_record[0];
    let edges = record.expanded_edges.as_ref().unwrap();
    assert_eq!(edges.edge_without_owner_rows, [0, 0, 1, 1]);
    assert_eq!(edges.edge_without_owner_columns, [0, 1, 0, 1]);
    let tile_id = dependencies
        .tiled_cell_dependencies
        .as_ref()
        .unwrap()
        .cell_record_tiles[0]
        .identifier;
    let tile = tsce::CellRecordTileArchive::decode(
        archive.object(tile_id).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(tile.cell_records[0].expanded_edges.as_ref(), Some(edges));

    crate::numbers::editor::set_cell_fixture(&mut editor, table_id, 3, 2, CellValue::Empty)
        .unwrap();
    let archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let dependencies = tsce::FormulaOwnerDependenciesArchive::decode(
        archive.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert!(
        dependencies
            .cell_dependencies
            .unwrap()
            .cell_record
            .is_empty()
    );
    let engine = tsce::CalculationEngineArchive::decode(
        archive.object(100).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(engine.dependency_tracker.number_of_formulas, Some(0));
}

#[test]
fn cell_write_refreshes_aggregate_range_formula_cache() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(2.0)).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 2, 1, cell_number(3.0)).unwrap();
    editor
        .set_formula_with_cached_value(
            10,
            3,
            2,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    crate::numbers::FormulaCellReference::relative(1, 1),
                    crate::numbers::FormulaCellReference::relative(2, 1),
                )],
            ),
            cached_number(5.0),
        )
        .unwrap();

    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(4.0)).unwrap();

    assert_eq!(
        cached_formula_scalar(&editor, 10, 3, 2),
        cached_scalar_number(7.0)
    );
}

#[test]
fn cell_write_refreshes_typed_boolean_formula_cache() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(1.0)).unwrap();
    editor
        .set_formula_with_cached_value(
            10,
            1,
            2,
            FormulaExpression::binary(
                crate::numbers::FormulaBinaryOperator::GreaterThan,
                FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
                FormulaExpression::Number(1.0),
            ),
            FormulaCachedValue::Boolean(false),
        )
        .unwrap();

    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(2.0)).unwrap();

    assert_eq!(
        cached_formula_scalar(&editor, 10, 1, 2),
        crate::numbers::bnc::CachedScalar::Boolean(true)
    );
}

#[test]
fn cell_write_refreshes_all_supported_numeric_aggregates() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(2.0)).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 2, 1, cell_number(4.0)).unwrap();
    for (row, function, cached) in [
        (0, "AVERAGE", 3.0),
        (1, "COUNT", 2.0),
        (2, "MIN", 2.0),
        (3, "MAX", 4.0),
    ] {
        editor
            .set_formula_with_cached_value(
                10,
                row,
                2,
                FormulaExpression::function(
                    function,
                    [FormulaExpression::range(
                        crate::numbers::FormulaCellReference::relative(1, 1),
                        crate::numbers::FormulaCellReference::relative(2, 1),
                    )],
                ),
                cached_number(cached),
            )
            .unwrap();
    }

    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(6.0)).unwrap();

    for (row, expected) in [(0, 5.0), (1, 2.0), (2, 4.0), (3, 6.0)] {
        assert_eq!(
            cached_formula_scalar(&editor, 10, row, 2),
            cached_scalar_number(expected)
        );
    }
}

#[test]
fn cell_write_refreshes_cross_table_formula_cache() {
    let mut editor = NumbersEditor::from_package(test_package_with_cross_table_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 11, 0, 1, cell_number(2.0)).unwrap();
    editor
        .set_formula_with_cached_value(
            10,
            3,
            2,
            FormulaExpression::table_cell(11, crate::numbers::FormulaCellReference::relative(0, 1)),
            cached_number(2.0),
        )
        .unwrap();

    crate::numbers::editor::set_cell_fixture(&mut editor, 11, 0, 1, cell_number(4.0)).unwrap();

    assert_eq!(
        cached_formula_scalar(&editor, 10, 3, 2),
        cached_scalar_number(4.0)
    );
}

fn cached_formula_scalar(
    editor: &NumbersEditor,
    table_id: u64,
    row: usize,
    column: usize,
) -> crate::numbers::bnc::CachedScalar {
    let location = locate_attached_cell(editor.package(), table_id, row, column).unwrap();
    let bytes = read_tile_cell(
        editor.package(),
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )
    .unwrap()
    .unwrap();
    BncCell::parse(&bytes)
        .unwrap()
        .cached_scalar()
        .unwrap()
        .unwrap()
}

#[test]
fn formula_dependency_tiles_use_app_dimensions_and_global_record_coordinates() {
    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message_type = object.messages[0].type_;
            let mut model = TableModelArchive::decode(object.messages[0].data.as_slice())?;
            model.number_of_rows = 200;
            model.number_of_columns = 40;
            object.replace_message(
                0,
                RawMessage {
                    type_: message_type,
                    data: model.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();

    editor
        .set_formula(
            10,
            130,
            33,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(129, 31)),
        )
        .unwrap();

    let archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        archive.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    let tile_id = owner.tiled_cell_dependencies.unwrap().cell_record_tiles[0].identifier;
    let tile = tsce::CellRecordTileArchive::decode(
        archive.object(tile_id).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(tile.tile_column_begin, 32);
    assert_eq!(tile.tile_row_begin, 128);
    assert_eq!(
        (tile.cell_records[0].column, tile.cell_records[0].row),
        (33, 130)
    );
    let edges = tile.cell_records[0].expanded_edges.as_ref().unwrap();
    assert_eq!(edges.edge_without_owner_columns, [31]);
    assert_eq!(edges.edge_without_owner_rows, [129]);
}

#[test]
fn cross_table_formula_cells_write_owner_uid_ast_and_external_edges() {
    const VERSIONED_ENGINE_ENTRY: &str = "Index/CalculationEngine-77.iwa";
    let mut package = test_package_with_cross_table_engine();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry(VERSIONED_ENGINE_ENTRY, engine)
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .set_formula(
            10,
            3,
            2,
            FormulaExpression::function(
                "SUM",
                [
                    FormulaExpression::relative_cell(0, 0),
                    FormulaExpression::table_cell(
                        11,
                        crate::numbers::FormulaCellReference::relative(0, 1),
                    ),
                ],
            ),
        )
        .unwrap();

    let document_archive = editor.package().archive("Index/Document.iwa").unwrap();
    let formulas = TableDataList::decode(
        document_archive.object(21).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let node = &formulas.entries[0]
        .formula
        .as_ref()
        .unwrap()
        .ast_node_array
        .ast_node[1];
    let table_id = &node
        .ast_cross_table_reference_extra_info
        .as_ref()
        .unwrap()
        .table_id;
    assert_eq!(table_id.uuid_w0, Some(0x3170_bbd8));
    assert_eq!(table_id.uuid_w3, Some(0xbd57_cadc));

    let calculation_archive = editor.package().archive(VERSIONED_ENGINE_ENTRY).unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        calculation_archive.object(101).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let edges = owner.cell_dependencies.unwrap().cell_record[0]
        .expanded_edges
        .clone()
        .unwrap();
    assert_eq!(edges.edge_without_owner_columns, [0]);
    assert_eq!(edges.edge_without_owner_rows, [0]);
    assert_eq!(edges.edge_with_owner_columns, [1]);
    assert_eq!(edges.edge_with_owner_rows, [0]);
    assert_eq!(edges.internal_owner_id_for_edge, [7]);

    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets();
    let host = sheets[0]
        .tables()
        .find(|table| table.name() == "Table 1")
        .unwrap();
    assert_eq!(
        host.get_cell(3, 2),
        Some(&CellValue::Formula(
            "=SUM(A1,Sheet 1::Target::B1)".to_owned()
        ))
    );
}

#[test]
fn cross_table_formula_ranges_expand_external_edges_in_row_major_order() {
    let mut editor = NumbersEditor::from_package(test_package_with_cross_table_engine()).unwrap();
    editor
        .set_formula(
            10,
            3,
            2,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::table_range(
                    11,
                    crate::numbers::FormulaCellReference::relative(0, 1),
                    crate::numbers::FormulaCellReference::relative(1, 2),
                )],
            ),
        )
        .unwrap();

    let document_archive = editor.package().archive("Index/Document.iwa").unwrap();
    let formulas = TableDataList::decode(
        document_archive.object(21).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let node = &formulas.entries[0]
        .formula
        .as_ref()
        .unwrap()
        .ast_node_array
        .ast_node[0];
    assert!(node.ast_colon_tract.is_some());
    let table_id = &node
        .ast_cross_table_reference_extra_info
        .as_ref()
        .unwrap()
        .table_id;
    assert_eq!(table_id.uuid_w0, Some(0x3170_bbd8));
    assert_eq!(table_id.uuid_w3, Some(0xbd57_cadc));

    let calculation_archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        calculation_archive.object(101).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let edges = owner.cell_dependencies.unwrap().cell_record[0]
        .expanded_edges
        .clone()
        .unwrap();
    assert!(edges.edge_without_owner_columns.is_empty());
    assert!(edges.edge_without_owner_rows.is_empty());
    assert_eq!(edges.edge_with_owner_columns, [1, 2, 1, 2]);
    assert_eq!(edges.edge_with_owner_rows, [0, 0, 1, 1]);
    assert_eq!(edges.internal_owner_id_for_edge, [7, 7, 7, 7]);
}

#[test]
fn whole_row_formula_ranges_round_trip_with_complete_external_edges() {
    let mut editor = NumbersEditor::from_package(test_package_with_cross_table_engine()).unwrap();
    editor
        .set_formula(
            10,
            3,
            2,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::table_rows(
                    11,
                    crate::numbers::FormulaAxisReference::relative(0),
                    crate::numbers::FormulaAxisReference::relative(1),
                )],
            ),
        )
        .unwrap();

    let document_archive = editor.package().archive("Index/Document.iwa").unwrap();
    let formulas = TableDataList::decode(
        document_archive.object(21).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let node = &formulas.entries[0]
        .formula
        .as_ref()
        .unwrap()
        .ast_node_array
        .ast_node[0];
    let tract = node.ast_colon_tract.as_ref().unwrap();
    assert_eq!(tract.absolute_column[0].range_begin, i16::MAX as u32);
    assert!(tract.relative_column.is_empty());

    let calculation_archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        calculation_archive.object(101).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let edges = owner.cell_dependencies.unwrap().cell_record[0]
        .expanded_edges
        .clone()
        .unwrap();
    assert!(edges.edge_without_owner_columns.is_empty());
    assert!(edges.edge_without_owner_rows.is_empty());
    assert_eq!(edges.edge_with_owner_columns, [0, 1, 2, 3, 0, 1, 2, 3]);
    assert_eq!(edges.edge_with_owner_rows, [0, 0, 0, 0, 1, 1, 1, 1]);
    assert_eq!(edges.internal_owner_id_for_edge, [7; 8]);

    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let host = document.sheets()[0]
        .tables()
        .find(|table| table.name() == "Table 1")
        .unwrap()
        .clone();
    assert_eq!(
        host.get_cell(3, 2),
        Some(&CellValue::Formula("=SUM(Sheet 1::Target::1:2)".to_owned()))
    );
}

#[test]
fn focused_table_lock_edits_round_trip_through_legacy_host_creation() {
    let editor = NumbersDocumentBuilder::new()
        .table_name("Locked Table")
        .table_dimensions(3, 2)
        .build()
        .unwrap();
    let table = editor.tables().unwrap().remove(0);
    let baseline = editor.to_bytes().unwrap();
    let focused = FocusedNumbersPackage::from_bytes(&baseline).unwrap();
    assert_eq!(
        focused
            .table_lock(SheetSelector::index(0), TableSelector::name(&table.name))
            .unwrap(),
        FocusedTableLockState::Unlocked
    );
    let mut lock = focused
        .edit_table_lock(SheetSelector::index(0), TableSelector::name(&table.name))
        .unwrap();
    lock.lock();
    let locked = lock.commit().unwrap();

    let mut locked_bytes = Vec::new();
    locked.package().write_to(&mut locked_bytes).unwrap();
    let mut editor = NumbersEditor::from_bytes(&locked_bytes).unwrap();
    let duplicate = editor
        .duplicate_table(test_table_selector(&editor, table.object_id))
        .unwrap();
    let focused = FocusedNumbersPackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        focused
            .table_lock(SheetSelector::index(0), TableSelector::index(1))
            .unwrap(),
        FocusedTableLockState::Locked
    );

    let mut unlock_duplicate = focused
        .edit_table_lock(SheetSelector::index(0), TableSelector::index(1))
        .unwrap();
    unlock_duplicate.unlock();
    let unlocked_duplicate = unlock_duplicate.commit().unwrap();
    let mut unlocked_duplicate_bytes = Vec::new();
    unlocked_duplicate
        .package()
        .write_to(&mut unlocked_duplicate_bytes)
        .unwrap();
    let mut editor = NumbersEditor::from_bytes(&unlocked_duplicate_bytes).unwrap();
    let focused = FocusedNumbersPackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        focused
            .table_lock(SheetSelector::index(0), TableSelector::index(0))
            .unwrap(),
        FocusedTableLockState::Locked
    );
    assert_eq!(
        focused
            .table_lock(SheetSelector::index(0), TableSelector::index(1))
            .unwrap(),
        FocusedTableLockState::Unlocked
    );

    editor
        .remove_table(test_table_selector(&editor, duplicate.object_id))
        .unwrap();
    let focused = FocusedNumbersPackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let mut unlock_original = focused
        .edit_table_lock(SheetSelector::index(0), TableSelector::name(&table.name))
        .unwrap();
    unlock_original.unlock();
    let restored = unlock_original.commit().unwrap();
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes).unwrap();
    assert_eq!(restored_bytes, baseline);
}

#[test]
fn focused_cell_edit_round_trips_through_legacy_host_reader() {
    let fixture = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    let mut editor = NumbersEditor::open(fixture).unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;

    crate::numbers::editor::test_set_cell(
        &mut editor,
        table_id,
        1,
        2,
        CellValue::Text("Focused host interop".to_owned()),
    )
    .unwrap();

    assert_eq!(editor.tables().unwrap()[0].object_id, table_id);
    let focused = FocusedNumbersPackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        focused
            .table_cell(
                SheetSelector::index(0),
                TableSelector::index(0),
                litchi_numbers::table::CellPosition::new(1, 2),
            )
            .unwrap()
            .storage()
            .value(),
        Some(&CellValue::Text("Focused host interop".to_owned()))
    );
}

#[test]
fn builder_empty_table_accepts_focused_commit_apply_and_inverse() {
    fn bytes(package: &FocusedNumbersPackage) -> Vec<u8> {
        let mut output = Vec::new();
        package.write_to(&mut output).unwrap();
        output
    }

    let editor = NumbersDocumentBuilder::new()
        .table_name("Focused Builder")
        .table_dimensions(2, 3)
        .build()
        .unwrap();
    let source = FocusedNumbersPackage::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let before = bytes(&source);
    let position = litchi_numbers::table::CellPosition::new(1, 1);
    let commit = source
        .edit_table_cells(SheetSelector::index(0), TableSelector::index(0))
        .unwrap()
        .set(position, litchi_numbers::table::cells::Input::boolean(true))
        .unwrap()
        .commit()
        .unwrap();
    assert!(matches!(
        commit
            .package()
            .table_cell(SheetSelector::index(0), TableSelector::index(0), position)
            .unwrap()
            .storage()
            .value(),
        Some(CellValue::Boolean(true))
    ));

    let replay = source.apply_table_cells(commit.patch()).unwrap();
    assert_eq!(bytes(replay.package()), bytes(commit.package()));
    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())
        .unwrap();
    assert_eq!(bytes(restored.package()), before);
}

#[test]
fn duplicates_populated_table_with_independent_storage() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let created = editor
        .duplicate_table(test_table_selector(&editor, 10))
        .unwrap();

    assert_ne!(created.object_id, 10);
    assert_eq!(created.name, "Table 1 copy");
    assert_eq!((created.rows, created.columns), (4, 4));

    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let mut tables = document.sheets()[0].tables();
    assert_eq!(tables.len(), 2);
    assert_eq!(
        tables.next().unwrap().get_cell(0, 1).unwrap().as_text(),
        "Original"
    );
    assert_eq!(
        tables.next().unwrap().get_cell(0, 1).unwrap().as_text(),
        "Original"
    );

    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        created.object_id,
        0,
        1,
        CellValue::Text("Independent".to_owned()),
    )
    .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let mut tables = document.sheets()[0].tables();
    assert_eq!(
        tables.next().unwrap().get_cell(0, 1).unwrap().as_text(),
        "Original"
    );
    assert_eq!(
        tables.next().unwrap().get_cell(0, 1).unwrap().as_text(),
        "Independent"
    );

    assert_eq!(
        editor
            .duplicate_table(test_table_selector(&editor, 10))
            .unwrap()
            .name,
        "Table 1 copy 2"
    );
}

#[test]
fn duplicates_formula_table_with_independent_dependency_owner() {
    const VERSIONED_ENGINE_ENTRY: &str = "Index/CalculationEngine-78-2.iwa";
    let mut package = test_package_with_calculation_engine();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry(VERSIONED_ENGINE_ENTRY, engine)
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let expression = FormulaExpression::function(
        "SUM",
        [
            FormulaExpression::Number(1.0),
            FormulaExpression::Number(2.0),
        ],
    );
    editor.set_formula(10, 1, 1, expression).unwrap();

    let created = editor
        .duplicate_table(test_table_selector(&editor, 10))
        .unwrap();
    let owner = find_table_owner(editor.package(), created.object_id).unwrap();
    let cloned_table_info_id = owner.table_info_id;
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(document.sheets()[0].table_count(), 2);
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    assert_eq!(
        document.sheets()[0].tables().nth(1).unwrap().get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );

    let calculation = editor.package().archive(VERSIONED_ENGINE_ENTRY).unwrap();
    let owners = calculation
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == 4008)
        .filter_map(|message| {
            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).ok()
        })
        .collect::<Vec<_>>();
    let original_owner = owners
        .iter()
        .find(|candidate| {
            candidate
                .formula_owner
                .as_ref()
                .is_some_and(|owner| owner.identifier == 3)
        })
        .unwrap();
    let cloned_owner = owners
        .iter()
        .find(|candidate| {
            candidate
                .formula_owner
                .as_ref()
                .is_some_and(|candidate| candidate.identifier == owner.table_info_id)
        })
        .unwrap();
    assert_ne!(
        original_owner.internal_formula_owner_id,
        cloned_owner.internal_formula_owner_id
    );
    assert_ne!(
        original_owner.formula_owner_uid,
        cloned_owner.formula_owner_uid
    );
    let engine = calculation
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find_map(|message| {
            (message.type_ == 4000)
                .then(|| tsce::CalculationEngineArchive::decode(message.data.as_slice()).ok())
                .flatten()
        })
        .unwrap();
    assert_eq!(engine.dependency_tracker.number_of_formulas, Some(2));
    assert!(
        engine
            .dependency_tracker
            .owner_id_map
            .unwrap()
            .map_entry
            .iter()
            .any(|entry| entry.internal_owner_id == cloned_owner.internal_formula_owner_id)
    );

    editor
        .set_formula(
            created.object_id,
            1,
            1,
            FormulaExpression::function("SUM", [FormulaExpression::Number(9.0)]),
        )
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    assert_eq!(
        document.sheets()[0].tables().nth(1).unwrap().get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(9)".to_owned()))
    );

    editor
        .remove_table(test_table_selector(&editor, created.object_id))
        .unwrap();
    assert_eq!(editor.tables().unwrap().len(), 1);
    let calculation = editor.package().archive(VERSIONED_ENGINE_ENTRY).unwrap();
    let owners = calculation
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == 4008)
        .filter_map(|message| {
            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).ok()
        })
        .collect::<Vec<_>>();
    assert!(owners.iter().all(|owner| {
        owner
            .formula_owner
            .as_ref()
            .is_none_or(|owner| owner.identifier != cloned_table_info_id)
    }));
    let engine = calculation
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find_map(|message| {
            (message.type_ == 4000)
                .then(|| tsce::CalculationEngineArchive::decode(message.data.as_slice()).ok())
                .flatten()
        })
        .unwrap();
    assert_eq!(engine.dependency_tracker.number_of_formulas, Some(1));
}

#[test]
fn duplicates_app_normalized_range_graph_and_removes_it_without_orphans() {
    const FORMULA_OWNER_ID: u64 = 101;
    const RANGE_TILE_ID: u64 = 10_001;
    const RANGE_TILE_MESSAGE_TYPE: u32 = 4_010;

    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_formula(
            10,
            3,
            1,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    crate::numbers::FormulaCellReference::relative(1, 1),
                    crate::numbers::FormulaCellReference::relative(2, 1),
                )],
            ),
        )
        .unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let owner_object = archive.object_mut(FORMULA_OWNER_ID).unwrap();
            let message = owner_object.messages[0].clone();
            let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            let internal_owner_id = owner.internal_formula_owner_id;
            let owner_uuid = owner.formula_owner_uid;
            owner.cell_dependencies.as_mut().unwrap().cell_record[0].expanded_edges =
                Some(tsce::ExpandedEdgesArchive::default());
            owner.range_dependencies = Some(tsce::RangeDependenciesArchive {
                back_dependency: vec![tsce::RangeBackDependencyArchive {
                    cell_coord_row: 3,
                    cell_coord_column: 1,
                    internal_range_reference: Some(tsce::InternalRangeReferenceArchive {
                        owner_id: internal_owner_id,
                        range: tsce::RangeCoordinateArchive {
                            top_left_column: 1,
                            top_left_row: 1,
                            bottom_right_column: 1,
                            bottom_right_row: 2,
                        },
                    }),
                    ..Default::default()
                }],
            });
            owner.uuid_references = Some(tsce::UuidReferencesArchive {
                table_refs: vec![tsce::uuid_references_archive::TableRef {
                    owner_uuid,
                    coord_set: None,
                }],
                table_uuid_refs: vec![tsce::uuid_references_archive::TableWithUuidRef {
                    owner_uuid,
                    uuid_refs: vec![tsce::uuid_references_archive::UuidRef {
                        uuid: tsp::Uuid {
                            lower: 111,
                            upper: 222,
                        },
                        coord_set: Some(tsce::CellCoordSetArchive {
                            column_entries: vec![tsce::cell_coord_set_archive::ColumnEntry {
                                column: 1,
                                row_set: tsce::IndexSetArchive {
                                    entries: vec![tsce::index_set_archive::IndexSetEntry {
                                        range_begin: 3,
                                        range_end: None,
                                    }],
                                },
                            }],
                        }),
                    }],
                }],
            });
            owner.tiled_range_dependencies = Some(tsce::RangeDependenciesTiledArchive {
                range_precedents_tile: vec![Reference {
                    identifier: RANGE_TILE_ID,
                    ..Default::default()
                }],
            });
            let mut data = owner.encode_to_vec();
            data = crate::wire::transform_length_delimited_fields_at_path(
                &data,
                &[5, 2],
                |payload| {
                    let mut payload = payload.to_vec();
                    append_unknown_varint(&mut payload, 99, 990);
                    Ok(payload)
                },
            )?;
            data = crate::wire::transform_length_delimited_fields_at_path(
                &data,
                &[14, 2, 2, 2, 1, 2, 1],
                |payload| {
                    let mut payload = payload.to_vec();
                    append_unknown_varint(&mut payload, 98, 980);
                    Ok(payload)
                },
            )?;
            owner_object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            owner_object.archive_info.message_infos[0]
                .object_references
                .push(RANGE_TILE_ID);

            let mut tile_data = tsce::RangePrecedentsTileArchive {
                to_owner_id: internal_owner_id,
                from_to_range: vec![tsce::range_precedents_tile_archive::FromToRangeArchive {
                    from_coord: tsce::CellCoordinateArchive {
                        column: Some(1),
                        row: Some(3),
                        ..Default::default()
                    },
                    refers_to_rect: tsce::CellRectArchive {
                        origin: tsce::CellCoordinateArchive {
                            column: Some(1),
                            row: Some(1),
                            ..Default::default()
                        },
                        size: tsce::ColumnRowSize {
                            num_rows: Some(2),
                            ..Default::default()
                        },
                    },
                }],
            }
            .encode_to_vec();
            for (path, field, value) in [(vec![2], 97, 970), (vec![2, 2], 96, 960)] {
                tile_data = crate::wire::transform_length_delimited_fields_at_path(
                    &tile_data,
                    &path,
                    |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    },
                )?;
            }
            Ok(archive.insert_object(ArchiveObject::new(
                RANGE_TILE_ID,
                vec![RawMessage {
                    type_: RANGE_TILE_MESSAGE_TYPE,
                    data: tile_data,
                }],
            )?)?)
        })
        .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    let created = editor
        .duplicate_table(test_table_selector(&editor, 10))
        .unwrap();
    let cloned_table_info_id = find_table_owner(editor.package(), created.object_id)
        .unwrap()
        .table_info_id;
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(document.sheets()[0].table_count(), 2);
    assert!(
        document.sheets()[0].tables().all(
            |table| table.get_cell(3, 1) == Some(&CellValue::Formula("=SUM(B2:B3)".to_owned()))
        )
    );

    let calculation = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owners = calculation
        .objects
        .iter()
        .filter_map(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == 4_008)
                .map(|message| {
                    (
                        object.archive_info.identifier.unwrap(),
                        message,
                        tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                            .unwrap(),
                    )
                })
        })
        .collect::<Vec<_>>();
    let original = owners
        .iter()
        .find(|(_, _, owner)| owner.formula_owner.as_ref().unwrap().identifier == 3)
        .unwrap();
    let cloned = owners
        .iter()
        .find(|(_, _, owner)| {
            owner.formula_owner.as_ref().unwrap().identifier == cloned_table_info_id
        })
        .unwrap();
    assert_ne!(
        original.2.internal_formula_owner_id,
        cloned.2.internal_formula_owner_id
    );
    assert_ne!(original.2.formula_owner_uid, cloned.2.formula_owner_uid);
    let cloned_range = &cloned
        .2
        .range_dependencies
        .as_ref()
        .unwrap()
        .back_dependency[0];
    assert_eq!(
        cloned_range
            .internal_range_reference
            .as_ref()
            .unwrap()
            .owner_id,
        cloned.2.internal_formula_owner_id
    );
    let cloned_uuids = cloned.2.uuid_references.as_ref().unwrap();
    assert_eq!(
        cloned_uuids.table_refs[0].owner_uuid,
        cloned.2.formula_owner_uid
    );
    assert_eq!(
        cloned_uuids.table_uuid_refs[0].owner_uuid,
        cloned.2.formula_owner_uid
    );
    assert_eq!(
        cloned_uuids.table_uuid_refs[0].uuid_refs[0].uuid,
        tsp::Uuid {
            lower: 111,
            upper: 222
        }
    );
    let original_tile_id = original
        .2
        .tiled_range_dependencies
        .as_ref()
        .unwrap()
        .range_precedents_tile[0]
        .identifier;
    let cloned_tile_id = cloned
        .2
        .tiled_range_dependencies
        .as_ref()
        .unwrap()
        .range_precedents_tile[0]
        .identifier;
    assert_eq!(original_tile_id, RANGE_TILE_ID);
    assert_ne!(cloned_tile_id, original_tile_id);
    let cloned_tile_object = calculation.object(cloned_tile_id).unwrap();
    let cloned_tile_message = &cloned_tile_object.messages[0];
    let cloned_tile =
        tsce::RangePrecedentsTileArchive::decode(cloned_tile_message.data.as_slice()).unwrap();
    assert_eq!(cloned_tile.to_owner_id, cloned.2.internal_formula_owner_id);
    assert_eq!(cloned_tile.from_to_range[0].from_coord.row, Some(3));
    let suffix = |field: u32, value: u64| {
        let mut data = litchi_iwa_common::varint::encode_varint(u64::from(field) << 3);
        data.extend(litchi_iwa_common::varint::encode_varint(value));
        data
    };
    let range_suffix = suffix(99, 990);
    let owner_ranges = crate::wire::repeated_length_delimited_payloads(&cloned.1.data, 5).unwrap();
    let range_records =
        crate::wire::repeated_length_delimited_payloads(owner_ranges[0], 2).unwrap();
    assert!(range_records[0].ends_with(&range_suffix));
    let mut uuid_payload = cloned.1.data.as_slice();
    for field in [14, 2, 2, 2, 1, 2, 1] {
        let nested = crate::wire::repeated_length_delimited_payloads(uuid_payload, field).unwrap();
        uuid_payload = nested[0];
    }
    assert!(uuid_payload.ends_with(&suffix(98, 980)));
    let tile_suffix = suffix(97, 970);
    let tile_records =
        crate::wire::repeated_length_delimited_payloads(&cloned_tile_message.data, 2).unwrap();
    assert!(tile_records[0].ends_with(&tile_suffix));
    let rect_suffix = suffix(96, 960);
    let rects = crate::wire::repeated_length_delimited_payloads(tile_records[0], 2).unwrap();
    assert!(rects[0].ends_with(&rect_suffix));

    editor
        .remove_table(test_table_selector(&editor, created.object_id))
        .unwrap();
    let calculation = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    assert!(calculation.object(RANGE_TILE_ID).is_some());
    assert!(calculation.object(cloned_tile_id).is_none());
    assert!(calculation.object(cloned.0).is_none());
}

#[test]
fn formula_table_duplicate_rejects_unsupported_dependencies_transactionally() {
    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let object = archive.object_mut(101).unwrap();
            let message = object.messages[0].clone();
            let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            owner.spill_range_sizes = Some(tsce::CellSpillSizesArchive {
                spills: vec![tsce::cell_spill_sizes_archive::SpillForCell {
                    coordinate: tsce::CellCoordinateArchive {
                        column: Some(1),
                        row: Some(1),
                        ..Default::default()
                    },
                    spill_size: tsce::ColumnRowSize {
                        num_columns: Some(2),
                        num_rows: Some(2),
                    },
                }],
            });
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: owner.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();

    assert!(
        editor
            .duplicate_table(test_table_selector(&editor, 10))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn resize_preserves_unknown_wire_and_restores_exact_component() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            for (identifier, field_number) in
                [(2, 99), (10, 98), (30, 97), (40, 96), (41, 95), (42, 94)]
            {
                let object = archive.object_mut(identifier).unwrap();
                let mut message = object.messages[0].clone();
                append_unknown_varint(&mut message.data, field_number, 900 + identifier);
                object.replace_message(0, message)?;
            }
            for (identifier, paths) in [
                (30, vec![vec![5]]),
                (40, vec![vec![1], vec![4]]),
                (42, vec![vec![2]]),
            ] {
                let object = archive.object_mut(identifier).unwrap();
                let mut message = object.messages[0].clone();
                for path in paths {
                    message.data = crate::wire::transform_length_delimited_fields_at_path(
                        &message.data,
                        &path,
                        |nested| {
                            let mut nested = nested.to_vec();
                            append_unknown_varint(&mut nested, 93, 930 + identifier);
                            Ok(nested)
                        },
                    )?;
                }
                object.replace_message(0, message)?;
            }
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let baseline = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();

    editor
        .resize_table(test_table_selector(&editor, 10), 6, 6)
        .unwrap();
    editor
        .resize_table(test_table_selector(&editor, 10), 4, 4)
        .unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        baseline
    );
}

#[test]
fn grows_and_truncates_blank_table_edges_with_uid_maps() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    editor
        .resize_table(test_table_selector(&editor, 10), 6, 6)
        .unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        5,
        5,
        CellValue::Text("edge".to_owned()),
    )
    .unwrap();
    assert!(
        editor
            .resize_table(test_table_selector(&editor, 10), 4, 4)
            .is_err()
    );
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 5, 5, CellValue::Empty).unwrap();
    editor
        .resize_table(test_table_selector(&editor, 10), 3, 3)
        .unwrap();
    let table = editor.tables().unwrap().remove(0);
    assert_eq!((table.rows, table.columns), (3, 3));

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let uid_object = archive.object(40).unwrap();
    let uid_map = uid_object
        .messages
        .iter()
        .find_map(|message| tst::ColumnRowUidMapArchive::decode(message.data.as_slice()).ok())
        .unwrap();
    assert_eq!(uid_map.sorted_row_uids.len(), 3);
    assert_eq!(uid_map.sorted_column_uids.len(), 3);
    let sidecar = archive
        .object(41)
        .unwrap()
        .messages
        .iter()
        .find_map(|message| tst::StrokeSidecarArchive::decode(message.data.as_slice()).ok())
        .unwrap();
    assert_eq!(sidecar.row_count, Some(3));
    assert_eq!(sidecar.column_count, Some(3));
}

#[test]
fn table_cell_border_crud_allocates_sparse_layers_and_round_trips() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let side = BorderSide::Bottom;
    assert_eq!(editor.table_cell_borders(10, 1, 2).unwrap().get(side), None);

    let stroke = crate::shapes::Stroke::new(
        crate::shapes::RgbaColor::new(0.8, 0.1, 0.2, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
        crate::shapes::Width::new(2.5).unwrap(),
        crate::shapes::Pattern::MediumDash,
    );
    editor
        .set_table_cell_border(10, 1, 2, side, stroke)
        .unwrap();
    assert_eq!(
        editor.table_cell_borders(10, 1, 2).unwrap().get(side),
        Some(stroke)
    );
    assert_eq!(editor.table_cell_borders(10, 1, 1).unwrap().get(side), None);

    let sidecar = test_stroke_sidecar(editor.package());
    assert_eq!(sidecar.max_order, Some(1));
    assert_eq!(sidecar.bottom_row_stroke_layers.len(), 1);
    let layer_id = sidecar.bottom_row_stroke_layers[0].identifier;
    let layer = test_stroke_layer(editor.package(), layer_id);
    assert_eq!(layer.row_column_index, Some(1));
    assert_eq!(
        layer
            .stroke_runs
            .iter()
            .map(|run| (run.origin, run.length, run.order))
            .collect::<Vec<_>>(),
        [(Some(2), Some(1), Some(1))]
    );

    editor.clear_table_cell_border(10, 1, 2, side).unwrap();
    assert_eq!(editor.table_cell_borders(10, 1, 2).unwrap().get(side), None);
    let sidecar = test_stroke_sidecar(editor.package());
    assert_eq!(sidecar.max_order, Some(2));
    assert_eq!(sidecar.bottom_row_stroke_layers[0].identifier, layer_id);
    let layer = test_stroke_layer(editor.package(), layer_id);
    assert_eq!(layer.stroke_runs.len(), 1);
    assert_eq!(layer.stroke_runs[0].order, Some(2));
    assert_eq!(
        crate::shapes::stroke_from_native(layer.stroke_runs[0].stroke.as_ref().unwrap()).unwrap(),
        None
    );

    let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_borders(10, 1, 2).unwrap().get(side),
        None
    );
}

#[test]
fn table_cell_fill_crud_round_trips_and_reuses_private_style() {
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Fill")
        .table_name("Colors")
        .table_dimensions(3, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let original = editor.table_cell_fill(table_id, 1, 1).unwrap();
    let red = crate::shapes::ShapeFill::Solid(
        crate::shapes::RgbaColor::new(0.85, 0.1, 0.2, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
    );
    let blue = crate::shapes::ShapeFill::Solid(
        crate::shapes::RgbaColor::new(0.1, 0.25, 0.9, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
    );

    editor.set_table_cell_fill(table_id, 1, 1, &red).unwrap();
    assert_eq!(editor.table_cell_fill(table_id, 1, 1).unwrap(), red);
    assert_eq!(editor.table_cell_fill(table_id, 1, 2).unwrap(), original);
    let after_first = storage::object_locations(editor.package()).unwrap().len();

    editor.set_table_cell_fill(table_id, 1, 1, &blue).unwrap();
    assert_eq!(
        storage::object_locations(editor.package()).unwrap().len(),
        after_first
    );
    assert_eq!(editor.table_cell_fill(table_id, 1, 1).unwrap(), blue);

    let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(reopened.reset_table_cell_fill(table_id, 1, 1).unwrap());
    assert_eq!(reopened.table_cell_fill(table_id, 1, 1).unwrap(), original);
    assert!(!reopened.reset_table_cell_fill(table_id, 1, 1).unwrap());
}

#[test]
fn table_cell_layout_composes_with_fill_and_resets_independently() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Layout")
        .table_dimensions(3, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let original_layout = editor.table_cell_layout(table_id, 1, 1).unwrap();
    let original_object_count = storage::object_locations(editor.package()).unwrap().len();
    let fill = crate::shapes::ShapeFill::Solid(
        crate::shapes::RgbaColor::new(0.8, 0.2, 0.1, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
    );
    editor.set_table_cell_fill(table_id, 1, 1, &fill).unwrap();
    let styled_object_count = storage::object_locations(editor.package()).unwrap().len();
    assert_eq!(styled_object_count, original_object_count + 1);

    let layout = Layout::default()
        .with_text_wrap(TextWrap::Wrapped)
        .with_vertical_alignment(VerticalAlignment::Middle)
        .with_insets(Insets::new(
            Inset::from_points(1.0).unwrap(),
            Inset::from_points(2.0).unwrap(),
            Inset::from_points(3.0).unwrap(),
            Inset::from_points(4.0).unwrap(),
        ));
    editor
        .set_table_cell_layout(table_id, 1, 1, layout)
        .unwrap();
    assert_eq!(
        storage::object_locations(editor.package()).unwrap().len(),
        styled_object_count
    );

    let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.table_cell_layout(table_id, 1, 1).unwrap(), layout);
    assert_eq!(reopened.table_cell_fill(table_id, 1, 1).unwrap(), fill);
    assert!(reopened.reset_table_cell_layout(table_id, 1, 1).unwrap());
    assert_eq!(
        reopened.table_cell_layout(table_id, 1, 1).unwrap(),
        original_layout
    );
    assert_eq!(reopened.table_cell_fill(table_id, 1, 1).unwrap(), fill);
    assert!(!reopened.reset_table_cell_layout(table_id, 1, 1).unwrap());
    assert!(reopened.reset_table_cell_fill(table_id, 1, 1).unwrap());
    assert_eq!(
        storage::object_locations(reopened.package()).unwrap().len(),
        original_object_count
    );
}

#[test]
fn table_cell_layout_rejects_invalid_coordinates_transactionally() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_table_cell_layout(table_id, 2, 0, Layout::default(),)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn table_cell_border_rejects_invalid_coordinates_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let before = editor.to_bytes().unwrap();
    let stroke = crate::shapes::Stroke::new(
        crate::shapes::RgbaColor::black(),
        crate::shapes::Width::ONE,
        crate::shapes::Pattern::Solid,
    );
    assert!(
        editor
            .set_table_cell_border(10, usize::MAX, 0, BorderSide::Top, stroke,)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn table_cell_border_update_preserves_unknown_sidecar_layer_and_run_fields() {
    let mut package = test_package();
    add_test_stroke_layer(&mut package, 50, TestStrokeLayerSide::Top, 1, &[(1, 1)]);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let sidecar = archive.object_mut(41).unwrap();
            let message = sidecar.messages[0].clone();
            let mut data = message.data;
            append_unknown_varint(&mut data, 99, 9_999);
            sidecar.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;

            let layer = archive.object_mut(50).unwrap();
            let message = layer.messages[0].clone();
            let mut data = crate::wire::transform_length_delimited_fields_at_path(
                &message.data,
                &[2],
                |run| {
                    let mut run = run.to_vec();
                    append_unknown_varint(&mut run, 97, 9_797);
                    Ok(run)
                },
            )?;
            append_unknown_varint(&mut data, 98, 9_898);
            Ok(layer
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )
                .map(drop)?)
        })
        .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .set_table_cell_border(
            10,
            1,
            1,
            BorderSide::Top,
            crate::shapes::Stroke::new(
                crate::shapes::RgbaColor::black(),
                crate::shapes::Width::ONE,
                crate::shapes::Pattern::Solid,
            ),
        )
        .unwrap();

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let sidecar = &archive.object(41).unwrap().messages[0].data;
    assert!(
        crate::wire::parse_wire_fields(sidecar)
            .unwrap()
            .iter()
            .any(|field| field.number() == 99)
    );
    let layer = &archive.object(50).unwrap().messages[0].data;
    assert!(
        crate::wire::parse_wire_fields(layer)
            .unwrap()
            .iter()
            .any(|field| field.number() == 98)
    );
    let runs = crate::wire::repeated_length_delimited_payloads(layer, 2).unwrap();
    assert_eq!(runs.len(), 1);
    assert!(
        crate::wire::parse_wire_fields(runs[0])
            .unwrap()
            .iter()
            .any(|field| field.number() == 97)
    );
}

#[test]
fn table_row_insertion_preserves_explicit_stroke_layers_on_original_cells() {
    let mut package = test_package();
    add_test_stroke_layer(&mut package, 50, TestStrokeLayerSide::Top, 1, &[(1, 1)]);
    add_test_stroke_layer(&mut package, 51, TestStrokeLayerSide::Left, 1, &[(1, 3)]);

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(1))
        .unwrap();

    let sidecar = test_stroke_sidecar(editor.package());
    assert_eq!(sidecar.row_count, Some(5));
    assert_eq!(
        sidecar
            .top_row_stroke_layers
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [50]
    );
    assert_eq!(
        sidecar
            .left_column_stroke_layers
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [51]
    );
    assert_eq!(
        test_stroke_layer(editor.package(), 50).row_column_index,
        Some(2)
    );
    let left = test_stroke_layer(editor.package(), 51);
    assert_eq!(left.stroke_runs[0].origin, Some(2));
    assert_eq!(left.stroke_runs[0].length, Some(3));
}

#[test]
fn table_row_insertion_splits_crossing_stroke_runs_without_normalizing_unknown_fields() {
    let mut package = test_package();
    add_test_stroke_layer(&mut package, 50, TestStrokeLayerSide::Left, 1, &[(1, 3)]);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(50).unwrap();
            let message = object.messages[0].clone();
            let data = crate::wire::transform_length_delimited_fields_at_path(
                &message.data,
                &[2],
                |run| {
                    let mut run = run.to_vec();
                    append_unknown_varint(&mut run, 99, 990);
                    Ok(run)
                },
            )?;
            Ok(object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )
                .map(|_| ())?)
        })
        .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(2))
        .unwrap();

    let layer = test_stroke_layer(editor.package(), 50);
    assert_eq!(
        layer
            .stroke_runs
            .iter()
            .map(|run| (run.origin, run.length))
            .collect::<Vec<_>>(),
        [(Some(1), Some(1)), (Some(3), Some(2))]
    );
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let raw_runs = crate::wire::repeated_length_delimited_payloads(
        archive.object(50).unwrap().messages[0].data.as_slice(),
        2,
    )
    .unwrap();
    assert_eq!(raw_runs.len(), 2);
    for run in raw_runs {
        assert!(
            crate::wire::parse_wire_fields(run)
                .unwrap()
                .iter()
                .any(|field| field.number() == 99)
        );
    }
}

#[test]
fn table_row_deletion_removes_or_compacts_explicit_stroke_layers_and_metadata_references() {
    let mut package = test_package();
    add_test_stroke_layer(&mut package, 50, TestStrokeLayerSide::Top, 1, &[(1, 1)]);
    add_test_stroke_layer(&mut package, 51, TestStrokeLayerSide::Left, 1, &[(1, 3)]);
    package
        .update_archive("Index/Document.iwa", |archive| {
            archive.object_mut(41).unwrap().archive_info.message_infos[0]
                .object_references
                .push(999);
            Ok(())
        })
        .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(1))
        .unwrap();

    let sidecar = test_stroke_sidecar(editor.package());
    assert_eq!(sidecar.row_count, Some(3));
    assert!(sidecar.top_row_stroke_layers.is_empty());
    assert_eq!(
        sidecar
            .left_column_stroke_layers
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [51]
    );
    let left = test_stroke_layer(editor.package(), 51);
    assert_eq!(left.stroke_runs[0].origin, Some(1));
    assert_eq!(left.stroke_runs[0].length, Some(2));
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let references = &archive.object(41).unwrap().archive_info.message_infos[0].object_references;
    assert!(!references.contains(&50));
    assert!(references.contains(&51));
    assert!(references.contains(&999));
}

#[test]
fn table_column_insertion_preserves_explicit_stroke_layers_on_original_cells() {
    let mut package = test_package();
    add_test_stroke_layer(&mut package, 50, TestStrokeLayerSide::Left, 1, &[(1, 1)]);
    add_test_stroke_layer(&mut package, 51, TestStrokeLayerSide::Top, 1, &[(1, 3)]);

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(1))
        .unwrap();

    let sidecar = test_stroke_sidecar(editor.package());
    assert_eq!(sidecar.column_count, Some(5));
    assert_eq!(
        test_stroke_layer(editor.package(), 50).row_column_index,
        Some(2)
    );
    let top = test_stroke_layer(editor.package(), 51);
    assert_eq!(top.stroke_runs[0].origin, Some(2));
    assert_eq!(top.stroke_runs[0].length, Some(3));
}

#[test]
fn inserts_blank_table_row_and_shifts_cells_uids_headers_and_formulas() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("Apples".to_owned()),
    )
    .unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(1))
        .unwrap();

    let bytes = editor.to_bytes().unwrap();
    let document = compatibility_document_from_bytes(&bytes).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!((table.row_count(), table.column_count()), (5, 4));
    assert_eq!(table.get_cell(1, 1), None);
    assert_eq!(
        table.get_cell(2, 1),
        Some(&CellValue::Text("Apples".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 2),
        Some(&CellValue::Formula("=B3".to_owned()))
    );

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let uid_map = tst::ColumnRowUidMapArchive::decode(
        archive.object(40).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(uid_map.sorted_row_uids.len(), 5);
    assert_eq!(uid_map.row_index_for_uid, [0, 2, 3, 4, 1]);
    assert_eq!(uid_map.row_uid_for_index, [0, 4, 1, 2, 3]);
    let sidecar =
        tst::StrokeSidecarArchive::decode(archive.object(41).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(sidecar.row_count, Some(5));
    let headers =
        tst::HeaderStorageBucket::decode(archive.object(42).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(
        headers
            .headers
            .iter()
            .map(|header| (header.index, header.number_of_cells))
            .collect::<Vec<_>>(),
        [(0, 1), (2, 1), (3, 1)]
    );

    let engine = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        engine.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    let record = &owner.cell_dependencies.unwrap().cell_record[0];
    assert_eq!((record.row, record.column), (3, 2));
    assert_eq!(
        record
            .expanded_edges
            .as_ref()
            .unwrap()
            .edge_without_owner_rows,
        [2]
    );
    let tile_id = owner.tiled_cell_dependencies.unwrap().cell_record_tiles[0].identifier;
    let tile = tsce::CellRecordTileArchive::decode(
        engine.object(tile_id).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(
        (tile.cell_records[0].row, tile.cell_records[0].column),
        (3, 2)
    );
}

#[test]
fn appends_blank_table_row_without_allocating_storage() {
    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let object = archive.object_mut(101).unwrap();
            let message = object.messages[0].clone();
            let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            let spanning = tsce::SpanningDependenciesExpandedArchive {
                total_range_for_table: Some(tsce::RangeCoordinateArchive {
                    top_left_column: 0,
                    top_left_row: 0,
                    bottom_right_column: 3,
                    bottom_right_row: 3,
                }),
                body_range_for_table: Some(tsce::RangeCoordinateArchive {
                    top_left_column: 0,
                    top_left_row: 0,
                    bottom_right_column: 3,
                    bottom_right_row: 3,
                }),
                ..Default::default()
            };
            owner.spanning_column_dependencies = Some(spanning.clone());
            owner.spanning_row_dependencies = Some(spanning);
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: owner.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(4))
        .unwrap();
    assert_eq!(editor.tables().unwrap()[0].rows, 5);
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(document.sheets()[0].tables().next().unwrap().row_count(), 5);
    let archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        archive.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(
        owner
            .spanning_column_dependencies
            .unwrap()
            .total_range_for_table
            .unwrap()
            .bottom_right_row,
        4
    );
}

#[test]
fn row_insert_rejects_missing_destination_tile_transactionally() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message = object.messages[0].clone();
            let mut model = TableModelArchive::decode(message.data.as_slice())?;
            model.base_data_store.tiles.tile_size = Some(4);
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: model.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 3, 0, cell_number(9.0)).unwrap();
    let before = editor.to_bytes().unwrap();

    assert!(
        editor
            .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(3))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_rejects_out_of_bounds_index_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(5))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_rewrites_relative_formula_ast_losslessly() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    let before = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(2))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(3, 2),
        Some(&CellValue::Formula("=B2".to_owned()))
    );
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(2))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn column_insert_rewrites_absolute_formula_ast_losslessly() {
    let mut editor =
        NumbersEditor::from_package(test_package_with_column_headers_and_engine()).unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::absolute(1, 1)),
        )
        .unwrap();
    let before = editor.to_bytes().unwrap();

    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(1))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(2, 3),
        Some(&CellValue::Formula("=$C$2".to_owned()))
    );
    editor
        .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::body(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_rewrites_range_ast_and_preserves_unknown_formula_wire() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_formula(
            10,
            3,
            2,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    crate::numbers::FormulaCellReference::relative(0, 1),
                    crate::numbers::FormulaCellReference::relative(1, 1),
                )],
            ),
        )
        .unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(21).unwrap();
            let message = object.messages[0].clone();
            let mut data = message.data;
            for (path, field, value) in [
                (vec![3], 99, 990),
                (vec![3, 5], 98, 980),
                (vec![3, 5, 1], 97, 970),
                (vec![3, 5, 1, 1], 96, 960),
                (vec![3, 5, 1, 1, 27], 95, 950),
            ] {
                data = transform_length_delimited_fields_at_path(&data, &path, |payload| {
                    let mut payload = payload.to_vec();
                    append_unknown_varint(&mut payload, field, value);
                    Ok(payload)
                })?;
            }
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(2))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(4, 2),
        Some(&CellValue::Formula("=SUM(B1:B2)".to_owned()))
    );
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(2))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_rewrites_segmented_formula_ast_losslessly() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    let mut package = editor.into_package();
    move_table_data_list_entries_to_segment(&mut package, 21, 61);
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(2))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(3, 2),
        Some(&CellValue::Formula("=B2".to_owned()))
    );
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(2))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_copy_on_writes_shared_formula_ast_and_remerges_on_delete() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    editor
        .set_formula(
            10,
            3,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(2, 1)),
        )
        .unwrap();
    let before = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(2))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(3, 2),
        Some(&CellValue::Formula("=B2".to_owned()))
    );
    assert_eq!(
        table.get_cell(4, 2),
        Some(&CellValue::Formula("=B4".to_owned()))
    );
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(2))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_expands_footer_aggregate_and_delete_restores_exact_bytes() {
    let mut package = test_package_with_calculation_engine();
    set_table_header_settings_in_package(
        &mut package,
        10,
        HeaderSettings {
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .set_formula(
            10,
            3,
            1,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    crate::numbers::FormulaCellReference::relative(1, 1),
                    crate::numbers::FormulaCellReference::relative(2, 1),
                )],
            ),
        )
        .unwrap();
    let mut package = editor.into_package();
    let dependency_tile_id = {
        let archive = package.archive("Index/CalculationEngine.iwa").unwrap();
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(
            archive.object(101).unwrap().messages[0].data.as_slice(),
        )
        .unwrap();
        owner.tiled_cell_dependencies.unwrap().cell_record_tiles[0].identifier
    };
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            for (object_id, path, field, value) in [
                (101, vec![4, 1, 6], 99, 990),
                (dependency_tile_id, vec![4, 6], 98, 980),
            ] {
                let object = archive.object_mut(object_id).unwrap();
                let message = object.messages[0].clone();
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    &message.data,
                    &path,
                    |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(3))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(4, 1),
        Some(&CellValue::Formula("=SUM(B2:B4)".to_owned()))
    );
    let archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        archive.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    let record = &owner.cell_dependencies.unwrap().cell_record[0];
    let edges = record.expanded_edges.as_ref().unwrap();
    assert_eq!((record.row, record.column), (4, 1));
    assert_eq!(edges.edge_without_owner_rows, [1, 2, 3]);
    assert_eq!(edges.edge_without_owner_columns, [1, 1, 1]);

    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(3))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_roundtrips_app_normalized_footer_range_dependencies() {
    const FORMULA_OWNER_ID: u64 = 101;
    const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
    const RANGE_TILE_ID: u64 = 10_001;
    const RANGE_TILE_MESSAGE_TYPE: u32 = 4_010;
    const VERSIONED_ENGINE_ENTRY: &str = "Index/CalculationEngine-10-2.iwa";

    let mut package = test_package_with_calculation_engine();
    set_table_header_settings_in_package(
        &mut package,
        10,
        HeaderSettings {
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(2.0)).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 2, 1, cell_number(3.0)).unwrap();
    editor
        .set_formula_with_cached_value(
            10,
            3,
            1,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    crate::numbers::FormulaCellReference::relative(1, 1),
                    crate::numbers::FormulaCellReference::relative(2, 1),
                )],
            ),
            cached_number(5.0),
        )
        .unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            use tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive;
            use tsce::ast_node_array_archive::{
                AstColonTractArchive, AstNodeArchive, AstNodeType, AstStickyBits,
            };

            let object = archive.object_mut(21).unwrap();
            let mut formulas = TableDataList::decode(object.messages[0].data.as_slice())?;
            let formula = formulas.entries[0].formula.as_mut().unwrap();
            let function = formula.ast_node_array.ast_node.pop().unwrap();
            formula.ast_node_array.ast_node = vec![
                AstNodeArchive {
                    ast_node_type: AstNodeType::ColonTractNode as i32,
                    ast_sticky_bits: Some(AstStickyBits {
                        begin_row_is_absolute: false,
                        begin_column_is_absolute: false,
                        end_row_is_absolute: false,
                        end_column_is_absolute: false,
                    }),
                    ast_colon_tract: Some(AstColonTractArchive {
                        relative_column: vec![AstColonTractRelativeRangeArchive {
                            range_begin: 0,
                            range_end: None,
                        }],
                        relative_row: vec![AstColonTractRelativeRangeArchive {
                            range_begin: -2,
                            range_end: Some(-1),
                        }],
                        preserve_rectangular: Some(true),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                function,
            ];
            object.replace_message(
                0,
                RawMessage {
                    type_: object.messages[0].type_,
                    data: formulas.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let object = archive.object_mut(FORMULA_OWNER_ID).unwrap();
            let mut owner =
                tsce::FormulaOwnerDependenciesArchive::decode(object.messages[0].data.as_slice())?;
            let internal_owner_id = owner.internal_formula_owner_id;
            let owner_uid = owner.formula_owner_uid;
            let cell_tile_id = owner
                .tiled_cell_dependencies
                .as_ref()
                .unwrap()
                .cell_record_tiles[0]
                .identifier;
            owner.cell_dependencies.as_mut().unwrap().cell_record[0].expanded_edges =
                Some(tsce::ExpandedEdgesArchive::default());
            owner.range_dependencies = Some(tsce::RangeDependenciesArchive {
                back_dependency: vec![tsce::RangeBackDependencyArchive {
                    cell_coord_row: 3,
                    cell_coord_column: 1,
                    internal_range_reference: Some(tsce::InternalRangeReferenceArchive {
                        owner_id: internal_owner_id,
                        range: tsce::RangeCoordinateArchive {
                            top_left_column: 1,
                            top_left_row: 1,
                            bottom_right_column: 1,
                            bottom_right_row: 2,
                        },
                    }),
                    ..Default::default()
                }],
            });
            owner.uuid_references = Some(tsce::UuidReferencesArchive {
                table_uuid_refs: vec![tsce::uuid_references_archive::TableWithUuidRef {
                    owner_uuid: owner_uid,
                    uuid_refs: vec![tsce::uuid_references_archive::UuidRef {
                        uuid: tsp::Uuid {
                            lower: 111,
                            upper: 222,
                        },
                        coord_set: Some(tsce::CellCoordSetArchive {
                            column_entries: vec![tsce::cell_coord_set_archive::ColumnEntry {
                                column: 1,
                                row_set: tsce::IndexSetArchive {
                                    entries: vec![tsce::index_set_archive::IndexSetEntry {
                                        range_begin: 3,
                                        range_end: None,
                                    }],
                                },
                            }],
                        }),
                    }],
                }],
                ..Default::default()
            });
            owner.tiled_range_dependencies = Some(tsce::RangeDependenciesTiledArchive {
                range_precedents_tile: vec![Reference {
                    identifier: RANGE_TILE_ID,
                    ..Default::default()
                }],
            });
            object.replace_message(
                0,
                RawMessage {
                    type_: FORMULA_OWNER_MESSAGE_TYPE,
                    data: owner.encode_to_vec(),
                },
            )?;
            let tile_object = archive.object_mut(cell_tile_id).unwrap();
            let mut tile =
                tsce::CellRecordTileArchive::decode(tile_object.messages[0].data.as_slice())?;
            tile.cell_records[0].expanded_edges = Some(tsce::ExpandedEdgesArchive::default());
            tile_object.replace_message(
                0,
                RawMessage {
                    type_: 4_009,
                    data: tile.encode_to_vec(),
                },
            )?;
            archive.insert_object(ArchiveObject::new(
                RANGE_TILE_ID,
                vec![RawMessage {
                    type_: RANGE_TILE_MESSAGE_TYPE,
                    data: tsce::RangePrecedentsTileArchive {
                        to_owner_id: internal_owner_id,
                        from_to_range: vec![
                            tsce::range_precedents_tile_archive::FromToRangeArchive {
                                from_coord: tsce::CellCoordinateArchive {
                                    column: Some(1),
                                    row: Some(3),
                                    ..Default::default()
                                },
                                refers_to_rect: tsce::CellRectArchive {
                                    origin: tsce::CellCoordinateArchive {
                                        column: Some(1),
                                        row: Some(1),
                                        ..Default::default()
                                    },
                                    size: tsce::ColumnRowSize {
                                        num_rows: Some(2),
                                        ..Default::default()
                                    },
                                },
                            },
                        ],
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            for (object_id, path, field, value) in [
                (FORMULA_OWNER_ID, vec![5, 2], 99, 990),
                (FORMULA_OWNER_ID, vec![14, 2, 2, 2, 1, 2, 1], 98, 980),
                (RANGE_TILE_ID, vec![2], 97, 970),
                (RANGE_TILE_ID, vec![2, 2], 96, 960),
            ] {
                let object = archive.object_mut(object_id).unwrap();
                let message = object.messages[0].clone();
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    &message.data,
                    &path,
                    |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry(VERSIONED_ENGINE_ENTRY, engine)
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(4.0)).unwrap();
    assert_eq!(
        cached_formula_scalar(&editor, 10, 3, 1),
        cached_scalar_number(7.0)
    );
    let before = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(3))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(4, 1),
        Some(&CellValue::Formula("=SUM(B2:B4)".to_owned()))
    );
    let archive = editor.package().archive(VERSIONED_ENGINE_ENTRY).unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        archive.object(FORMULA_OWNER_ID).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let dependency = &owner.range_dependencies.as_ref().unwrap().back_dependency[0];
    assert_eq!(
        (dependency.cell_coord_row, dependency.cell_coord_column),
        (4, 1)
    );
    assert_eq!(
        dependency
            .internal_range_reference
            .as_ref()
            .unwrap()
            .range
            .bottom_right_row,
        3
    );
    let uuid_row = &owner.uuid_references.as_ref().unwrap().table_uuid_refs[0].uuid_refs[0]
        .coord_set
        .as_ref()
        .unwrap()
        .column_entries[0]
        .row_set
        .entries[0];
    assert_eq!(uuid_row.range_begin, 4);
    let range_tile = tsce::RangePrecedentsTileArchive::decode(
        archive.object(RANGE_TILE_ID).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    assert_eq!(range_tile.from_to_range[0].from_coord.row, Some(4));
    assert_eq!(
        range_tile.from_to_range[0].refers_to_rect.size.num_rows,
        Some(3)
    );

    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(3))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_rejects_incoming_cross_table_formula_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package_with_cross_table_engine()).unwrap();
    editor
        .set_formula(
            11,
            0,
            0,
            FormulaExpression::table_cell(10, crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    let before = editor.to_bytes().unwrap();

    assert!(
        editor
            .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(1))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_preserves_unknown_tile_header_and_dependency_record_fields() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("opaque".to_owned()),
    )
    .unwrap();
    editor
        .set_formula(10, 2, 2, FormulaExpression::Number(7.0))
        .unwrap();
    let mut package = editor.into_package();
    let dependency_tile_id = {
        let archive = package.archive("Index/CalculationEngine.iwa").unwrap();
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(
            archive.object(101).unwrap().messages[0].data.as_slice(),
        )
        .unwrap();
        owner.tiled_cell_dependencies.unwrap().cell_record_tiles[0].identifier
    };
    package
        .update_archive("Index/Document.iwa", |archive| {
            for (object_id, path, field, value) in [(30, vec![5], 99, 990), (42, vec![2], 98, 980)]
            {
                let object = archive.object_mut(object_id).unwrap();
                let message = object.messages[0].clone();
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    &message.data,
                    &path,
                    |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            for (object_id, path, field, value) in [
                (101, vec![4, 1], 97, 970),
                (dependency_tile_id, vec![4], 96, 960),
            ] {
                let object = archive.object_mut(object_id).unwrap();
                let message = object.messages[0].clone();
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    &message.data,
                    &path,
                    |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(1))
        .unwrap();

    let document = editor.package().archive("Index/Document.iwa").unwrap();
    let tile_rows =
        repeated_length_delimited_payloads(&document.object(30).unwrap().messages[0].data, 5)
            .unwrap();
    let headers =
        repeated_length_delimited_payloads(&document.object(42).unwrap().messages[0].data, 2)
            .unwrap();
    let engine = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner_dependencies = repeated_length_delimited_payloads(
        repeated_length_delimited_payloads(&engine.object(101).unwrap().messages[0].data, 4)
            .unwrap()[0],
        1,
    )
    .unwrap();
    let tiled_dependencies = repeated_length_delimited_payloads(
        &engine.object(dependency_tile_id).unwrap().messages[0].data,
        4,
    )
    .unwrap();
    let suffix = |field, value| {
        let mut bytes = Vec::new();
        append_unknown_varint(&mut bytes, field, value);
        bytes
    };
    let tile_suffix = suffix(99, 990);
    let header_suffix = suffix(98, 980);
    let owner_suffix = suffix(97, 970);
    let tiled_suffix = suffix(96, 960);
    assert!(tile_rows.iter().all(|row| row.ends_with(&tile_suffix)));
    assert!(
        headers
            .iter()
            .all(|header| header.ends_with(&header_suffix))
    );
    assert_eq!(owner_dependencies.len(), 1);
    assert!(owner_dependencies[0].ends_with(&owner_suffix));
    assert_eq!(tiled_dependencies.len(), 1);
    assert!(tiled_dependencies[0].ends_with(&tiled_suffix));
}

#[test]
fn inserts_blank_table_column_and_shifts_cells_uids_headers_and_formulas() {
    let mut editor =
        NumbersEditor::from_package(test_package_with_column_headers_and_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("Apples".to_owned()),
    )
    .unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();

    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(1))
        .unwrap();

    let bytes = editor.to_bytes().unwrap();
    let document = compatibility_document_from_bytes(&bytes).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!((table.row_count(), table.column_count()), (4, 5));
    assert_eq!(table.get_cell(1, 1), None);
    assert_eq!(
        table.get_cell(1, 2),
        Some(&CellValue::Text("Apples".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 3),
        Some(&CellValue::Formula("=C2".to_owned()))
    );

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let uid_map = tst::ColumnRowUidMapArchive::decode(
        archive.object(40).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(uid_map.sorted_column_uids.len(), 5);
    assert_eq!(uid_map.column_index_for_uid, [0, 2, 3, 4, 1]);
    assert_eq!(uid_map.column_uid_for_index, [0, 4, 1, 2, 3]);
    let sidecar =
        tst::StrokeSidecarArchive::decode(archive.object(41).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(sidecar.column_count, Some(5));
    let headers =
        tst::HeaderStorageBucket::decode(archive.object(43).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(
        headers
            .headers
            .iter()
            .map(|header| (header.index, header.number_of_cells))
            .collect::<Vec<_>>(),
        [(0, 1), (2, 1), (3, 1)]
    );

    let engine = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        engine.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    let record = &owner.cell_dependencies.as_ref().unwrap().cell_record[0];
    assert_eq!((record.row, record.column), (2, 3));
    assert_eq!(
        record
            .expanded_edges
            .as_ref()
            .unwrap()
            .edge_without_owner_columns,
        [2]
    );
    let tile_id = owner.tiled_cell_dependencies.unwrap().cell_record_tiles[0].identifier;
    let tile = tsce::CellRecordTileArchive::decode(
        engine.object(tile_id).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(
        (tile.cell_records[0].row, tile.cell_records[0].column),
        (2, 3)
    );
}

#[test]
fn appends_blank_table_column_and_grows_dependency_ranges() {
    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let object = archive.object_mut(101).unwrap();
            let message = object.messages[0].clone();
            let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            let spanning = tsce::SpanningDependenciesExpandedArchive {
                total_range_for_table: Some(tsce::RangeCoordinateArchive {
                    top_left_column: 0,
                    top_left_row: 0,
                    bottom_right_column: 3,
                    bottom_right_row: 3,
                }),
                body_range_for_table: Some(tsce::RangeCoordinateArchive {
                    top_left_column: 0,
                    top_left_row: 0,
                    bottom_right_column: 3,
                    bottom_right_row: 3,
                }),
                ..Default::default()
            };
            owner.spanning_column_dependencies = Some(spanning.clone());
            owner.spanning_row_dependencies = Some(spanning);
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: owner.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();

    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(4))
        .unwrap();

    assert_eq!(editor.tables().unwrap()[0].columns, 5);
    let archive = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner = tsce::FormulaOwnerDependenciesArchive::decode(
        archive.object(101).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(
        owner
            .spanning_column_dependencies
            .unwrap()
            .total_range_for_table
            .unwrap()
            .bottom_right_column,
        4
    );
}

#[test]
fn column_insert_rejects_out_of_bounds_and_incoming_formulas_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(5))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut editor = NumbersEditor::from_package(test_package_with_cross_table_engine()).unwrap();
    editor
        .set_formula(
            11,
            0,
            0,
            FormulaExpression::table_cell(10, crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(1))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn column_insert_rejects_short_cell_offset_tables_transactionally() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(30).unwrap();
            let message = object.messages[0].clone();
            let mut tile = Tile::decode(message.data.as_slice())?;
            tile.row_infos[0].cell_offsets = Some(vec![0, 0]);
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: tile.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();

    assert!(
        editor
            .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(1))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn column_insert_preserves_unknown_tile_header_and_dependency_record_fields() {
    let mut editor =
        NumbersEditor::from_package(test_package_with_column_headers_and_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("opaque".to_owned()),
    )
    .unwrap();
    editor
        .set_formula(10, 2, 2, FormulaExpression::Number(7.0))
        .unwrap();
    let mut package = editor.into_package();
    let dependency_tile_id = {
        let archive = package.archive("Index/CalculationEngine.iwa").unwrap();
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(
            archive.object(101).unwrap().messages[0].data.as_slice(),
        )
        .unwrap();
        owner.tiled_cell_dependencies.unwrap().cell_record_tiles[0].identifier
    };
    package
        .update_archive("Index/Document.iwa", |archive| {
            for (object_id, path, field, value) in [(30, vec![5], 99, 990), (43, vec![2], 98, 980)]
            {
                let object = archive.object_mut(object_id).unwrap();
                let message = object.messages[0].clone();
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    &message.data,
                    &path,
                    |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            for (object_id, path, field, value) in [
                (101, vec![4, 1], 97, 970),
                (dependency_tile_id, vec![4], 96, 960),
            ] {
                let object = archive.object_mut(object_id).unwrap();
                let message = object.messages[0].clone();
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    &message.data,
                    &path,
                    |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(1))
        .unwrap();

    let document = editor.package().archive("Index/Document.iwa").unwrap();
    let tile_rows =
        repeated_length_delimited_payloads(&document.object(30).unwrap().messages[0].data, 5)
            .unwrap();
    let headers =
        repeated_length_delimited_payloads(&document.object(43).unwrap().messages[0].data, 2)
            .unwrap();
    let engine = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
    let owner_dependencies = repeated_length_delimited_payloads(
        repeated_length_delimited_payloads(&engine.object(101).unwrap().messages[0].data, 4)
            .unwrap()[0],
        1,
    )
    .unwrap();
    let tiled_dependencies = repeated_length_delimited_payloads(
        &engine.object(dependency_tile_id).unwrap().messages[0].data,
        4,
    )
    .unwrap();
    let suffix = |field, value| {
        let mut bytes = Vec::new();
        append_unknown_varint(&mut bytes, field, value);
        bytes
    };
    assert!(tile_rows.iter().all(|row| row.ends_with(&suffix(99, 990))));
    assert!(
        headers
            .iter()
            .all(|header| header.ends_with(&suffix(98, 980)))
    );
    assert!(owner_dependencies[0].ends_with(&suffix(97, 970)));
    assert!(tiled_dependencies[0].ends_with(&suffix(96, 960)));
}

#[test]
fn row_insert_then_delete_restores_exact_package_bytes() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("Apples".to_owned()),
    )
    .unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::body(2))
        .unwrap();
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(2))
        .unwrap();

    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn column_insert_then_delete_restores_exact_package_bytes() {
    let mut editor =
        NumbersEditor::from_package(test_package_with_column_headers_and_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("Apples".to_owned()),
    )
    .unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::body(2))
        .unwrap();
    editor
        .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::body(2))
        .unwrap();

    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn section_relative_header_insertions_shift_formulas_and_restore_exactly() {
    let mut package = test_package_with_calculation_engine();
    set_table_header_settings_in_package(
        &mut package,
        10,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(2.0)).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 2, 1, cell_number(3.0)).unwrap();
    editor
        .set_formula(
            10,
            3,
            1,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    crate::numbers::FormulaCellReference::relative(1, 1),
                    crate::numbers::FormulaCellReference::relative(2, 1),
                )],
            ),
        )
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::header(1))
        .unwrap();
    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::header(1))
        .unwrap();

    let settings = table_header_settings_in_package(editor.package(), 10).unwrap();
    assert_eq!(settings.header_row_count(), 2);
    assert_eq!(settings.header_column_count(), 2);
    assert_eq!(settings.footer_row_count(), 1);
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(4, 2),
        Some(&CellValue::Formula("=SUM(C3:C4)".to_owned()))
    );

    editor
        .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::header(1))
        .unwrap();
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::header(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn footer_insertions_do_not_expand_body_formula_ranges() {
    let mut package = test_package_with_calculation_engine();
    set_table_header_settings_in_package(
        &mut package,
        10,
        HeaderSettings {
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 1, 1, cell_number(2.0)).unwrap();
    crate::numbers::editor::set_cell_fixture(&mut editor, 10, 2, 1, cell_number(3.0)).unwrap();
    editor
        .set_formula(
            10,
            3,
            1,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    crate::numbers::FormulaCellReference::relative(1, 1),
                    crate::numbers::FormulaCellReference::relative(2, 1),
                )],
            ),
        )
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::footer(0))
        .unwrap();
    assert_eq!(
        table_header_settings_in_package(editor.package(), 10)
            .unwrap()
            .footer_row_count(),
        2
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(4, 1),
        Some(&CellValue::Formula("=SUM(B2:B3)".to_owned()))
    );
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::footer(0))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::footer(1))
        .unwrap();
    assert_eq!(
        table_header_settings_in_package(editor.package(), 10)
            .unwrap()
            .footer_row_count(),
        2
    );
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::footer(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn section_insertions_create_first_fixed_regions_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::header(0))
        .unwrap();
    editor
        .insert_table_row(test_table_selector(&editor, 10), RowInsertion::footer(0))
        .unwrap();
    editor
        .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::header(0))
        .unwrap();
    let settings = table_header_settings_in_package(editor.package(), 10).unwrap();
    assert_eq!(settings.header_row_count(), 1);
    assert_eq!(settings.footer_row_count(), 1);
    assert_eq!(settings.header_column_count(), 1);

    editor
        .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::header(0))
        .unwrap();
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::footer(0))
        .unwrap();
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::header(0))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .insert_table_row(test_table_selector(&editor, 10), RowInsertion::header(1))
            .is_err()
    );
    assert!(
        editor
            .insert_table_row(test_table_selector(&editor, 10), RowInsertion::footer(1))
            .is_err()
    );
    assert!(
        editor
            .insert_table_column(test_table_selector(&editor, 10), ColumnInsertion::header(1))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn section_relative_deletions_target_fixed_regions_transactionally() {
    let mut package = test_package();
    set_table_header_settings_in_package(
        &mut package,
        10,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            header_columns: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        0,
        0,
        CellValue::Text("Header".to_owned()),
    )
    .unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("Body".to_owned()),
    )
    .unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        3,
        2,
        CellValue::Text("Footer".to_owned()),
    )
    .unwrap();

    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::header(0))
        .unwrap();
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::footer(0))
        .unwrap();
    editor
        .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::header(0))
        .unwrap();

    let settings = table_header_settings_in_package(editor.package(), 10).unwrap();
    assert_eq!(settings.header_row_count(), 0);
    assert_eq!(settings.footer_row_count(), 0);
    assert_eq!(settings.header_column_count(), 0);
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!((table.row_count(), table.column_count()), (2, 3));
    assert_eq!(
        table.get_cell(0, 0),
        Some(&CellValue::Text("Body".to_owned()))
    );
    assert!(!table.iter_cells().any(|cell| matches!(
        cell.value(),
        CellValue::Text(text) if text == "Header" || text == "Footer"
    )));

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .remove_table_row(test_table_selector(&editor, 10), RowDeletion::header(0))
            .is_err()
    );
    assert!(
        editor
            .remove_table_row(test_table_selector(&editor, 10), RowDeletion::footer(0))
            .is_err()
    );
    assert!(
        editor
            .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::header(0))
            .is_err()
    );
    assert!(
        editor
            .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(2))
            .is_err()
    );
    assert!(
        editor
            .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::body(3))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn removes_populated_table_axes_with_reference_cleanup_and_formula_shifts() {
    let mut editor =
        NumbersEditor::from_package(test_package_with_column_headers_and_engine()).unwrap();
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("discarded row".to_owned()),
    )
    .unwrap();
    editor
        .set_formula(10, 2, 2, FormulaExpression::Number(7.0))
        .unwrap();
    editor
        .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(1))
        .unwrap();

    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!((table.row_count(), table.column_count()), (3, 4));
    assert_eq!(table.get_cell(1, 1), None);
    assert_eq!(
        table.get_cell(1, 2),
        Some(&CellValue::Formula("=7".to_owned()))
    );

    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        1,
        1,
        CellValue::Text("discarded column".to_owned()),
    )
    .unwrap();
    editor
        .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::body(1))
        .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!((table.row_count(), table.column_count()), (3, 3));
    assert_eq!(
        table.get_cell(1, 1),
        Some(&CellValue::Formula("=7".to_owned()))
    );
}

#[test]
fn table_axis_delete_releases_comment_graphs() {
    for remove_row in [true, false] {
        let mut editor = NumbersEditor::from_package(test_package_with_comments(false)).unwrap();
        if remove_row {
            editor
                .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(0))
                .unwrap();
        } else {
            editor
                .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::body(1))
                .unwrap();
        }

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        assert!(archive.object(61).is_none());
        assert!(archive.object(70).is_none());
        let list =
            TableDataList::decode(archive.object(60).unwrap().messages[0].data.as_slice()).unwrap();
        assert!(list.entries.is_empty());
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(reopened.cell_comment(10, 0, 1).unwrap().is_none());
    }
}

#[test]
fn table_axis_delete_rejects_live_formula_references_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    assert!(
        editor
            .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(1))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert!(
        editor
            .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::body(1))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert!(
        editor
            .remove_table_row(test_table_selector(&editor, 10), RowDeletion::body(4))
            .is_err()
    );
    assert!(
        editor
            .remove_table_column(test_table_selector(&editor, 10), ColumnDeletion::body(4))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn table_sort_order_is_typed_transactional_and_native_clear_compatible() {
    let mut sort_unknown = Vec::new();
    append_unknown_varint(&mut sort_unknown, 98, 980);
    let mut rule_unknown = Vec::new();
    append_unknown_varint(&mut rule_unknown, 97, 970);
    let mut native = tst::TableSortOrderArchive {
        r#type: tst::table_sort_order_archive::SortType::EntireTable as i32,
        rules: vec![tst::table_sort_order_archive::SortRuleArchive {
            index: 1,
            direction: tst::table_sort_order_archive::sort_rule_archive::Direction::Ascending
                as i32,
        }],
    }
    .encode_to_vec();
    native.extend_from_slice(&sort_unknown);
    native = crate::wire::transform_length_delimited_fields_at_path(&native, &[2], |rule| {
        let mut rule = rule.to_vec();
        rule.extend_from_slice(&rule_unknown);
        Ok(rule)
    })
    .unwrap();
    let tracker = tst::SortRuleReferenceTrackerArchive {
        reference_tracker: Reference {
            identifier: 91,
            ..Default::default()
        },
    }
    .encode_to_vec();
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message = object.messages[0].clone();
            let mut data = message.data;
            append_unknown_varint(&mut data, 99, 990);
            crate::wire::append_length_delimited_field(&mut data, 44, &native)?;
            crate::wire::append_length_delimited_field(&mut data, 45, &tracker)?;
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert_eq!(
        editor.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(1).unwrap(),
                NumbersTableSortDirection::Ascending,
            )])
            .unwrap()
        )
    );

    let order = NumbersTableSortOrder::new([
        NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(2).unwrap(),
            NumbersTableSortDirection::Descending,
        ),
        NumbersTableSortRule::new(
            NumbersTableSortColumnIndex::new(1).unwrap(),
            NumbersTableSortDirection::Ascending,
        ),
    ])
    .unwrap();
    editor
        .set_table_sort_order(TableSelector::index(0), order.clone())
        .unwrap();
    assert_eq!(
        editor.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(order.clone())
    );
    assert_eq!(order.rules()[0].column().get(), 2);
    assert_eq!(
        order.rules()[0].direction(),
        NumbersTableSortDirection::Descending
    );
    let changed = editor.to_bytes().unwrap();
    let reparsed = NumbersEditor::from_bytes(&changed).unwrap();
    assert_eq!(
        reparsed.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(order.clone())
    );

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let model =
        TableModelArchive::decode(archive.object(10).unwrap().messages[0].data.as_slice()).unwrap();
    let native = model.sort_order.unwrap();
    assert_eq!(
        native.r#type,
        tst::table_sort_order_archive::SortType::EntireTable as i32
    );
    assert_eq!(native.rules.len(), 2);
    assert_eq!(native.rules[0].index, 2);
    assert_eq!(
        native.rules[0].direction,
        tst::table_sort_order_archive::sort_rule_archive::Direction::Descending as i32
    );
    let sort_payload = crate::wire::repeated_length_delimited_payloads(
        archive.object(10).unwrap().messages[0].data.as_slice(),
        44,
    )
    .unwrap()
    .pop()
    .unwrap();
    assert!(sort_payload.ends_with(&sort_unknown));
    let sort_rules = crate::wire::repeated_length_delimited_payloads(sort_payload, 2).unwrap();
    assert_eq!(sort_rules.len(), 2);
    assert!(sort_rules[0].ends_with(&rule_unknown));
    assert_eq!(
        crate::wire::repeated_length_delimited_payloads(
            archive.object(10).unwrap().messages[0].data.as_slice(),
            45,
        )
        .unwrap(),
        vec![tracker.as_slice()]
    );

    editor
        .set_table_sort_order(TableSelector::index(0), order.clone())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), changed);

    let out_of_bounds = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(4).unwrap(),
        NumbersTableSortDirection::Ascending,
    )])
    .unwrap();
    assert!(
        editor
            .set_table_sort_order(TableSelector::index(0), out_of_bounds)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), changed);

    editor
        .clear_table_sort_order(TableSelector::index(0))
        .unwrap();
    assert_eq!(
        editor.table_sort_order(TableSelector::index(0)).unwrap(),
        None
    );
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let model =
        TableModelArchive::decode(archive.object(10).unwrap().messages[0].data.as_slice()).unwrap();
    assert!(model.sort_order.is_some());
    assert!(model.sort_order.unwrap().rules.is_empty());
    let sort_payload = crate::wire::repeated_length_delimited_payloads(
        archive.object(10).unwrap().messages[0].data.as_slice(),
        44,
    )
    .unwrap()
    .pop()
    .unwrap();
    assert!(sort_payload.ends_with(&sort_unknown));
    assert!(
        crate::wire::repeated_length_delimited_payloads(sort_payload, 2)
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        crate::wire::repeated_length_delimited_payloads(
            archive.object(10).unwrap().messages[0].data.as_slice(),
            45,
        )
        .unwrap(),
        vec![tracker.as_slice()]
    );
    let cleared = editor.to_bytes().unwrap();
    editor
        .clear_table_sort_order(TableSelector::index(0))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), cleared);
    assert_ne!(cleared, baseline);
}

#[test]
fn table_sort_order_rejects_duplicate_wire_fields_transactionally() {
    let order = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(1).unwrap(),
        NumbersTableSortDirection::Ascending,
    )])
    .unwrap();
    let native = tst::TableSortOrderArchive {
        r#type: tst::table_sort_order_archive::SortType::EntireTable as i32,
        rules: vec![tst::table_sort_order_archive::SortRuleArchive {
            index: 1,
            direction: tst::table_sort_order_archive::sort_rule_archive::Direction::Ascending
                as i32,
        }],
    }
    .encode_to_vec();
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message = object.messages[0].clone();
            let mut data = message.data;
            crate::wire::append_length_delimited_field(&mut data, 44, &native)?;
            crate::wire::append_length_delimited_field(&mut data, 44, &native)?;
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.table_sort_order(TableSelector::index(0)).is_err());
    assert!(
        editor
            .set_table_sort_order(TableSelector::index(0), order)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn table_sort_order_rejects_malformed_nested_wire_transactionally() {
    const SORT_TYPE_FIELD: u32 = 1;
    let order = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(1).unwrap(),
        NumbersTableSortDirection::Ascending,
    )])
    .unwrap();
    let mut native = tst::TableSortOrderArchive {
        r#type: tst::table_sort_order_archive::SortType::EntireTable as i32,
        rules: vec![tst::table_sort_order_archive::SortRuleArchive {
            index: 1,
            direction: tst::table_sort_order_archive::sort_rule_archive::Direction::Ascending
                as i32,
        }],
    }
    .encode_to_vec();
    append_unknown_varint(
        &mut native,
        SORT_TYPE_FIELD,
        tst::table_sort_order_archive::SortType::EntireTable as u64,
    );
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message = object.messages[0].clone();
            let mut data = message.data;
            crate::wire::append_length_delimited_field(&mut data, 44, &native)?;
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.table_sort_order(TableSelector::index(0)).is_err());
    assert!(
        editor
            .set_table_sort_order(TableSelector::index(0), order)
            .is_err()
    );
    assert!(
        editor
            .clear_table_sort_order(TableSelector::index(0))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn source_created_sparse_boundary_supports_formula_and_comment_crud() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(257, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    editor
        .set_formula_with_cached_value(
            table_id,
            256,
            0,
            FormulaExpression::Number(42.0),
            cached_number(42.0),
        )
        .unwrap();
    editor
        .set_cell_comment(table_id, 256, 1, "Boundary comment")
        .unwrap();

    assert_eq!(
        editor
            .cell_comment(table_id, 256, 1)
            .unwrap()
            .unwrap()
            .comment
            .text,
        "Boundary comment"
    );
    let descriptor = attached_table_descriptor(editor.package(), table_id).unwrap();
    assert_eq!(
        descriptor
            .model
            .base_data_store
            .tiles
            .tiles
            .iter()
            .map(|tile| tile.tileid)
            .collect::<Vec<_>>(),
        vec![0, 1]
    );

    let bytes = editor.to_bytes().unwrap();
    let mut reopened = NumbersEditor::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened
            .cell_comment(table_id, 256, 1)
            .unwrap()
            .unwrap()
            .comment
            .text,
        "Boundary comment"
    );
    reopened.clear_cell_comment(table_id, 256, 1).unwrap();
    assert!(reopened.cell_comment(table_id, 256, 1).unwrap().is_none());
}

#[test]
fn source_created_large_table_allocates_header_buckets_only_when_needed() {
    const SECOND_HEADER_BUCKET_ROW: usize = 65_536;

    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(SECOND_HEADER_BUCKET_ROW + 1, 1)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        table_id,
        SECOND_HEADER_BUCKET_ROW,
        0,
        CellValue::Text("Second header bucket".to_owned()),
    )
    .unwrap();

    let descriptor = attached_table_descriptor(editor.package(), table_id).unwrap();
    assert_eq!(
        descriptor.model.base_data_store.row_headers.buckets.len(),
        2
    );
    assert_eq!(
        descriptor
            .model
            .base_data_store
            .tiles
            .tiles
            .iter()
            .map(|tile| tile.tileid)
            .collect::<Vec<_>>(),
        vec![0, 256]
    );
    assert_eq!(
        descriptor
            .model
            .base_data_store
            .row_tile_tree
            .nodes
            .iter()
            .map(|node| (node.key, node.value))
            .collect::<Vec<_>>(),
        vec![(0, 0), (65_536, 1)]
    );
    assert_eq!(descriptor.model.base_data_store.next_row_strip_id, 257);
    assert!(NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).is_ok());
}

#[test]
fn source_created_table_supports_sort_order_configuration_crud() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(4, 3)
        .build()
        .unwrap();
    let baseline = editor.to_bytes().unwrap();
    editor
        .clear_table_sort_order(TableSelector::index(0))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    let order = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(1).unwrap(),
        NumbersTableSortDirection::Ascending,
    )])
    .unwrap();
    editor
        .set_table_sort_order(TableSelector::index(0), order.clone())
        .unwrap();
    let reparsed = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reparsed.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(order)
    );

    editor
        .clear_table_sort_order(TableSelector::index(0))
        .unwrap();
    assert_eq!(
        editor.table_sort_order(TableSelector::index(0)).unwrap(),
        None
    );
}

#[test]
fn selected_row_sort_roundtrips_scope_and_moves_only_the_explicit_body_range() {
    let editor = NumbersDocumentBuilder::new()
        .table_dimensions(6, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let mut package = editor.into_package();
    set_table_header_settings_in_package(
        &mut package,
        table_id,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Region".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Q1".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("Outside".to_owned())),
            TableCellUpdate::new(1, 1, cell_number(50.0)),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, cell_number(98.0)),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, cell_number(105.0)),
            TableCellUpdate::new(4, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(4, 1, cell_number(120.0)),
            TableCellUpdate::new(5, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(5, 1, cell_number(323.0)),
        ],
    )
    .unwrap();
    editor
        .set_cell_comment(table_id, 2, 1, "Selected South comment")
        .unwrap();
    let reply_id = editor
        .add_cell_comment_reply(table_id, 2, 1, "Selected South reply")
        .unwrap();
    let comment_id = editor
        .cell_comment(table_id, 2, 1)
        .unwrap()
        .unwrap()
        .storage_id
        .get();
    let order = NumbersTableSortOrder::selected_rows([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(1).unwrap(),
        NumbersTableSortDirection::Descending,
    )])
    .unwrap();
    assert_eq!(order.scope(), NumbersTableSortScope::SelectedRows);
    editor
        .set_table_sort_order(TableSelector::index(0), order.clone())
        .unwrap();
    assert_eq!(
        NumbersEditor::from_bytes(&editor.to_bytes().unwrap())
            .unwrap()
            .table_sort_order(TableSelector::index(0))
            .unwrap(),
        Some(order.clone())
    );

    let before_wrong_executor = editor.to_bytes().unwrap();
    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_wrong_executor);
    let outside = NumbersTableSortRowRange::new(1, 5).unwrap();
    assert!(
        editor
            .apply_table_sort_order_to_rows(TableSelector::index(0), outside)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_wrong_executor);

    let selected = NumbersTableSortRowRange::new(1, 4).unwrap();
    assert!(
        editor
            .apply_table_sort_order_to_rows(TableSelector::index(0), selected)
            .unwrap()
    );
    assert_eq!(
        editor.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(order)
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(1, 0),
        Some(&CellValue::Text("Outside".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 0),
        Some(&CellValue::Text("North".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 0),
        Some(&CellValue::Text("Central".to_owned()))
    );
    assert_eq!(
        table.get_cell(4, 0),
        Some(&CellValue::Text("South".to_owned()))
    );
    assert_eq!(
        table.get_cell(5, 0),
        Some(&CellValue::Text("Total".to_owned()))
    );
    assert!(editor.cell_comment(table_id, 2, 1).unwrap().is_none());
    assert_eq!(
        editor
            .cell_comment(table_id, 4, 1)
            .unwrap()
            .unwrap()
            .storage_id
            .get(),
        comment_id
    );
    assert_eq!(
        editor.cell_comment_replies(table_id, 4, 1).unwrap()[0]
            .storage_id
            .get(),
        reply_id
    );

    let sorted = editor.to_bytes().unwrap();
    assert!(
        !editor
            .apply_table_sort_order_to_rows(TableSelector::index(0), selected)
            .unwrap()
    );
    assert_eq!(editor.to_bytes().unwrap(), sorted);
}

#[test]
fn table_sort_order_executes_stable_body_sort_and_remaps_row_uids() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        10,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("C first".to_owned())),
            TableCellUpdate::new(0, 1, cell_number(3.0)),
            TableCellUpdate::new(1, 0, CellValue::Text("A".to_owned())),
            TableCellUpdate::new(1, 1, cell_number(1.0)),
            TableCellUpdate::new(2, 0, CellValue::Text("C second".to_owned())),
            TableCellUpdate::new(2, 1, cell_number(3.0)),
            TableCellUpdate::new(3, 0, CellValue::Text("B".to_owned())),
            TableCellUpdate::new(3, 1, cell_number(2.0)),
        ],
    )
    .unwrap();
    let order = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(1).unwrap(),
        NumbersTableSortDirection::Ascending,
    )])
    .unwrap();
    editor
        .set_table_sort_order(TableSelector::index(0), order.clone())
        .unwrap();

    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    assert_eq!(
        editor.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(order)
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(table.get_cell(0, 0), Some(&CellValue::Text("A".to_owned())));
    assert_eq!(table.get_cell(0, 1), Some(&cell_number(1.0)));
    assert_eq!(table.get_cell(1, 0), Some(&CellValue::Text("B".to_owned())));
    assert_eq!(table.get_cell(1, 1), Some(&cell_number(2.0)));
    assert_eq!(
        table.get_cell(2, 0),
        Some(&CellValue::Text("C first".to_owned()))
    );
    assert_eq!(table.get_cell(2, 1), Some(&cell_number(3.0)));
    assert_eq!(
        table.get_cell(3, 0),
        Some(&CellValue::Text("C second".to_owned()))
    );
    assert_eq!(table.get_cell(3, 1), Some(&cell_number(3.0)));

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let uid_map = tst::ColumnRowUidMapArchive::decode(
        archive.object(40).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(uid_map.row_uid_for_index, [1, 3, 0, 2]);
    assert_eq!(uid_map.row_index_for_uid, [2, 0, 3, 1]);

    let sorted = editor.to_bytes().unwrap();
    assert!(
        !editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    assert_eq!(editor.to_bytes().unwrap(), sorted);
}

#[test]
fn source_created_table_executes_stable_plain_text_sort() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(5, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Name".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Marker".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("zebra".to_owned())),
            TableCellUpdate::new(1, 1, CellValue::Text("last".to_owned())),
            TableCellUpdate::new(2, 0, CellValue::Text("apple".to_owned())),
            TableCellUpdate::new(2, 1, CellValue::Text("first apple".to_owned())),
            TableCellUpdate::new(3, 0, CellValue::Text("banana".to_owned())),
            TableCellUpdate::new(3, 1, CellValue::Text("middle".to_owned())),
            TableCellUpdate::new(4, 0, CellValue::Text("apple".to_owned())),
            TableCellUpdate::new(4, 1, CellValue::Text("second apple".to_owned())),
        ],
    )
    .unwrap();
    let order = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(0).unwrap(),
        NumbersTableSortDirection::Ascending,
    )])
    .unwrap();
    editor
        .set_table_sort_order(TableSelector::index(0), order.clone())
        .unwrap();

    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    assert_eq!(
        editor.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(order)
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(0, 0),
        Some(&CellValue::Text("Name".to_owned()))
    );
    assert_eq!(
        table.get_cell(0, 1),
        Some(&CellValue::Text("Marker".to_owned()))
    );
    assert_eq!(
        table.get_cell(1, 0),
        Some(&CellValue::Text("apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(1, 1),
        Some(&CellValue::Text("first apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 0),
        Some(&CellValue::Text("apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 1),
        Some(&CellValue::Text("second apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 0),
        Some(&CellValue::Text("banana".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 1),
        Some(&CellValue::Text("middle".to_owned()))
    );
    assert_eq!(
        table.get_cell(4, 0),
        Some(&CellValue::Text("zebra".to_owned()))
    );
    assert_eq!(
        table.get_cell(4, 1),
        Some(&CellValue::Text("last".to_owned()))
    );

    let sorted = editor.to_bytes().unwrap();
    assert!(
        !editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    assert_eq!(editor.to_bytes().unwrap(), sorted);
}

#[test]
fn table_sort_resolves_plain_text_keys_from_segmented_string_storage() {
    const TABLE_ID: u64 = 10;
    const STRING_LIST_ID: u64 = 20;
    const STRING_SEGMENT_ID: u64 = 60;

    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        TABLE_ID,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("zebra".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("apple".to_owned())),
            TableCellUpdate::new(2, 0, CellValue::Text("banana".to_owned())),
            TableCellUpdate::new(3, 0, CellValue::Text("apple".to_owned())),
        ],
    )
    .unwrap();
    editor
        .set_table_sort_order(
            TableSelector::index(0),
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(0).unwrap(),
                NumbersTableSortDirection::Ascending,
            )])
            .unwrap(),
        )
        .unwrap();
    let mut package = editor.into_package();
    move_table_data_list_entries_to_segment(&mut package, STRING_LIST_ID, STRING_SEGMENT_ID);
    let mut editor = NumbersEditor::from_package(package).unwrap();

    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(0, 0),
        Some(&CellValue::Text("apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(1, 0),
        Some(&CellValue::Text("apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 0),
        Some(&CellValue::Text("banana".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 0),
        Some(&CellValue::Text("zebra".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 1),
        Some(&CellValue::Text("Original".to_owned()))
    );
}

#[test]
fn table_sort_rejects_missing_plain_text_storage_transactionally() {
    const TABLE_ID: u64 = 10;
    const STRING_LIST_ID: u64 = 20;

    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        TABLE_ID,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("zebra".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("apple".to_owned())),
            TableCellUpdate::new(2, 0, CellValue::Text("banana".to_owned())),
            TableCellUpdate::new(3, 0, CellValue::Text("apple".to_owned())),
        ],
    )
    .unwrap();
    editor
        .set_table_sort_order(
            TableSelector::index(0),
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(0).unwrap(),
                NumbersTableSortDirection::Ascending,
            )])
            .unwrap(),
        )
        .unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(STRING_LIST_ID).unwrap();
            let message = object.messages[0].clone();
            let mut strings = TableDataList::decode(message.data.as_slice())?;
            strings.entries.clear();
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: strings.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();

    let error = editor
        .apply_table_sort_order(TableSelector::index(0))
        .unwrap_err();
    assert!(error.to_string().contains("references missing string"));
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn source_created_table_executes_sort_order_without_moving_headers_or_footers() {
    let editor = NumbersDocumentBuilder::new()
        .table_dimensions(5, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let mut package = editor.into_package();
    set_table_header_settings_in_package(
        &mut package,
        table_id,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Region".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Q1".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(1, 1, cell_number(120.0)),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, cell_number(98.0)),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, cell_number(105.0)),
            TableCellUpdate::new(4, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(4, 1, cell_number(323.0)),
        ],
    )
    .unwrap();
    editor
        .set_cell_comment(table_id, 2, 1, "South comment follows row")
        .unwrap();
    let reply_id = editor
        .add_cell_comment_reply(table_id, 2, 1, "Reply follows row too")
        .unwrap();
    let original_comment = editor.cell_comment(table_id, 2, 1).unwrap().unwrap();
    let original_reply = editor.cell_comment_replies(table_id, 2, 1).unwrap()[0].clone();
    assert_eq!(original_reply.storage_id.get(), reply_id);
    let order = NumbersTableSortOrder::new([NumbersTableSortRule::new(
        NumbersTableSortColumnIndex::new(1).unwrap(),
        NumbersTableSortDirection::Ascending,
    )])
    .unwrap();
    editor
        .set_table_sort_order(TableSelector::index(0), order.clone())
        .unwrap();

    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(0, 0),
        Some(&CellValue::Text("Region".to_owned()))
    );
    assert_eq!(
        table.get_cell(0, 1),
        Some(&CellValue::Text("Q1".to_owned()))
    );
    assert_eq!(
        table.get_cell(1, 0),
        Some(&CellValue::Text("South".to_owned()))
    );
    assert_eq!(table.get_cell(1, 1), Some(&cell_number(98.0)));
    assert_eq!(
        table.get_cell(2, 0),
        Some(&CellValue::Text("Central".to_owned()))
    );
    assert_eq!(table.get_cell(2, 1), Some(&cell_number(105.0)));
    assert_eq!(
        table.get_cell(3, 0),
        Some(&CellValue::Text("North".to_owned()))
    );
    assert_eq!(table.get_cell(3, 1), Some(&cell_number(120.0)));
    assert_eq!(
        table.get_cell(4, 0),
        Some(&CellValue::Text("Total".to_owned()))
    );
    assert_eq!(table.get_cell(4, 1), Some(&cell_number(323.0)));
    assert!(editor.cell_comment(table_id, 2, 1).unwrap().is_none());
    let moved_comment = editor.cell_comment(table_id, 1, 1).unwrap().unwrap();
    assert_eq!(moved_comment.row, 1);
    assert_eq!(moved_comment.column, original_comment.column);
    assert_eq!(moved_comment.storage_id, original_comment.storage_id);
    assert_eq!(moved_comment.comment, original_comment.comment);
    let moved_replies = editor.cell_comment_replies(table_id, 1, 1).unwrap();
    assert_eq!(moved_replies.len(), 1);
    assert_eq!(
        moved_replies[0].root_storage_id,
        original_reply.root_storage_id
    );
    assert_eq!(moved_replies[0].storage_id, original_reply.storage_id);
    assert_eq!(moved_replies[0].comment, original_reply.comment);
    let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_sort_order(TableSelector::index(0)).unwrap(),
        Some(order)
    );
    assert!(reopened.cell_comment(table_id, 2, 1).unwrap().is_none());
    assert_eq!(
        reopened
            .cell_comment(table_id, 1, 1)
            .unwrap()
            .unwrap()
            .storage_id,
        original_comment.storage_id
    );
    assert_eq!(
        reopened.cell_comment_replies(table_id, 1, 1).unwrap()[0]
            .storage_id
            .get(),
        reply_id
    );
}

#[test]
fn table_sort_keeps_user_hidden_axes_at_their_physical_positions() {
    use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};

    let editor = NumbersDocumentBuilder::new()
        .table_dimensions(5, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let mut package = editor.into_package();
    set_table_header_settings_in_package(
        &mut package,
        table_id,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            footer_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        table_id,
        [
            TableCellUpdate::new(0, 0, CellValue::Text("Region".to_owned())),
            TableCellUpdate::new(0, 1, CellValue::Text("Q1".to_owned())),
            TableCellUpdate::new(1, 0, CellValue::Text("North".to_owned())),
            TableCellUpdate::new(1, 1, cell_number(120.0)),
            TableCellUpdate::new(2, 0, CellValue::Text("South".to_owned())),
            TableCellUpdate::new(2, 1, cell_number(98.0)),
            TableCellUpdate::new(3, 0, CellValue::Text("Central".to_owned())),
            TableCellUpdate::new(3, 1, cell_number(105.0)),
            TableCellUpdate::new(4, 0, CellValue::Text("Total".to_owned())),
            TableCellUpdate::new(4, 1, cell_number(323.0)),
        ],
    )
    .unwrap();
    let hidden = HiddenAxes::new([AxisIndex::row(2)]).unwrap();
    editor
        .set_table_hidden_axes(test_table_selector(&editor, table_id), &hidden)
        .unwrap();
    editor
        .set_table_sort_order(
            TableSelector::index(0),
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(1).unwrap(),
                NumbersTableSortDirection::Descending,
            )])
            .unwrap(),
        )
        .unwrap();

    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    assert_eq!(
        editor
            .table_hidden_axes(test_table_selector(&editor, table_id))
            .unwrap(),
        hidden
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(1, 0),
        Some(&CellValue::Text("North".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 0),
        Some(&CellValue::Text("Central".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 0),
        Some(&CellValue::Text("South".to_owned()))
    );
    let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table_hidden_axes(test_table_selector(&reopened, table_id))
            .unwrap(),
        hidden
    );
}

#[test]
fn table_sort_moves_rows_across_tile_boundaries() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let model_object = archive.object_mut(10).unwrap();
            let model_message = model_object.messages[0].clone();
            let mut model = TableModelArchive::decode(model_message.data.as_slice())?;
            model.base_data_store.tiles.tile_size = Some(2);
            model
                .base_data_store
                .tiles
                .tiles
                .push(tst::tile_storage::Tile {
                    tileid: 1,
                    tile: Reference {
                        identifier: 31,
                        ..Default::default()
                    },
                });
            model_object.replace_message(
                0,
                RawMessage {
                    type_: model_message.type_,
                    data: model.encode_to_vec(),
                },
            )?;
            Ok(archive.insert_object(ArchiveObject::new(
                31,
                vec![RawMessage {
                    type_: 6002,
                    data: Tile {
                        max_column: 0,
                        max_row: 3,
                        num_cells: 0,
                        numrows: 0,
                        row_infos: Vec::new(),
                        storage_version: Some(5),
                        last_saved_in_bnc: Some(true),
                        should_use_wide_rows: None,
                    }
                    .encode_to_vec(),
                }],
            )?)?)
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        10,
        [
            TableCellUpdate::new(0, 0, cell_number(3.0)),
            TableCellUpdate::new(1, 0, cell_number(2.0)),
            TableCellUpdate::new(2, 0, cell_number(1.0)),
            TableCellUpdate::new(3, 0, cell_number(0.0)),
        ],
    )
    .unwrap();
    editor
        .set_table_sort_order(
            TableSelector::index(0),
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(0).unwrap(),
                NumbersTableSortDirection::Ascending,
            )])
            .unwrap(),
        )
        .unwrap();

    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets()[0].tables().next().unwrap();
    assert_eq!(table.get_cell(0, 0), Some(&cell_number(0.0)));
    assert_eq!(table.get_cell(1, 0), Some(&cell_number(1.0)));
    assert_eq!(table.get_cell(2, 0), Some(&cell_number(2.0)));
    assert_eq!(table.get_cell(3, 0), Some(&cell_number(3.0)));

    let uid_map = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .object(40)
        .unwrap()
        .messages
        .iter()
        .find_map(|message| tst::ColumnRowUidMapArchive::decode(message.data.as_slice()).ok())
        .unwrap();
    assert_eq!(uid_map.row_uid_for_index, [3, 2, 1, 0]);
}

#[test]
fn table_sort_execution_keeps_explicit_border_layers_attached_to_cells() {
    let mut package = test_package();
    add_test_stroke_layer(&mut package, 88, TestStrokeLayerSide::Top, 1, &[(0, 4)]);
    add_test_stroke_layer(
        &mut package,
        89,
        TestStrokeLayerSide::Left,
        1,
        &[(0, 2), (2, 2)],
    );
    package
        .update_archive("Index/Document.iwa", |archive| {
            for identifier in [88, 89] {
                let object = archive.object_mut(identifier).unwrap();
                let message = object.messages[0].clone();
                let mut data = crate::wire::transform_length_delimited_fields_at_path(
                    &message.data,
                    &[2],
                    |run| {
                        let mut run = run.to_vec();
                        append_unknown_varint(&mut run, 98, 980);
                        Ok(run)
                    },
                )?;
                append_unknown_varint(&mut data, 99, 990);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();
    set_table_header_settings_in_package(
        &mut package,
        10,
        HeaderSettings {
            header_rows: Some(HeaderCount::ONE),
            ..Default::default()
        },
    )
    .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut editor,
        10,
        [
            TableCellUpdate::new(0, 1, CellValue::Text("Q1".to_owned())),
            TableCellUpdate::new(1, 1, cell_number(3.0)),
            TableCellUpdate::new(2, 1, cell_number(1.0)),
            TableCellUpdate::new(3, 1, cell_number(2.0)),
        ],
    )
    .unwrap();
    editor
        .set_table_sort_order(
            TableSelector::index(0),
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(1).unwrap(),
                NumbersTableSortDirection::Ascending,
            )])
            .unwrap(),
        )
        .unwrap();
    assert!(
        editor
            .apply_table_sort_order(TableSelector::index(0))
            .unwrap()
    );

    let top = test_stroke_layer(editor.package(), 88);
    assert_eq!(top.row_column_index, Some(3));
    assert_eq!(
        top.stroke_runs
            .iter()
            .map(|run| (run.origin, run.length))
            .collect::<Vec<_>>(),
        [(Some(0), Some(4))]
    );
    let left = test_stroke_layer(editor.package(), 89);
    assert_eq!(left.row_column_index, Some(1));
    assert_eq!(
        left.stroke_runs
            .iter()
            .map(|run| (run.origin, run.length))
            .collect::<Vec<_>>(),
        [(Some(0), Some(1)), (Some(3), Some(1)), (Some(1), Some(2)),]
    );
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    for identifier in [88, 89] {
        let data = archive.object(identifier).unwrap().messages[0]
            .data
            .as_slice();
        assert!(
            crate::wire::parse_wire_fields(data)
                .unwrap()
                .iter()
                .any(|field| field.number() == 99)
        );
        for run in crate::wire::repeated_length_delimited_payloads(data, 2).unwrap() {
            assert!(
                crate::wire::parse_wire_fields(run)
                    .unwrap()
                    .iter()
                    .any(|field| field.number() == 98)
            );
        }
    }
}

#[test]
fn table_sort_distinguishes_empty_and_populated_conditional_style_storage() {
    for (has_entries, should_sort) in [(false, true), (true, false)] {
        let mut package = test_package();
        add_test_app_native_topology_allocations(&mut package);
        add_test_conditional_style_storage(&mut package, has_entries);
        let mut editor = NumbersEditor::from_package(package).unwrap();
        crate::numbers::editor::apply_numbers_fixture(
            &mut editor,
            10,
            [
                TableCellUpdate::new(0, 1, cell_number(3.0)),
                TableCellUpdate::new(1, 1, cell_number(1.0)),
                TableCellUpdate::new(2, 1, cell_number(4.0)),
                TableCellUpdate::new(3, 1, cell_number(2.0)),
            ],
        )
        .unwrap();
        editor
            .set_table_sort_order(
                TableSelector::index(0),
                NumbersTableSortOrder::new([NumbersTableSortRule::new(
                    NumbersTableSortColumnIndex::new(1).unwrap(),
                    NumbersTableSortDirection::Ascending,
                )])
                .unwrap(),
            )
            .unwrap();
        let before = editor.to_bytes().unwrap();
        if should_sort {
            assert!(
                editor
                    .apply_table_sort_order(TableSelector::index(0))
                    .unwrap()
            );
        } else {
            assert!(
                editor
                    .apply_table_sort_order(TableSelector::index(0))
                    .is_err()
            );
            assert_eq!(editor.to_bytes().unwrap(), before);
        }
    }
}

#[test]
fn cell_conditional_highlighting_is_detected_and_deleted_without_changing_value() {
    let mut package = test_package();
    add_test_conditional_style_storage(&mut package, true);
    let location = locate_cell(&package, 10, 0, 1).unwrap();
    let cell_count = update_tile(
        &mut package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        1,
        location.descriptor.model.number_of_columns as usize,
        EncodedValue::ConditionalStyle {
            identifier: Some(1),
            applied_rule: Some(CONDITIONAL_STYLE_NO_APPLIED_RULE),
        },
    )
    .unwrap();
    update_row_header(
        &mut package,
        &location.object_locations,
        &location.descriptor.model,
        0,
        cell_count,
    )
    .unwrap();

    let mut editor = NumbersEditor::from_package(package).unwrap();
    let original_value = compatibility_document_from_bytes(&editor.to_bytes().unwrap())
        .unwrap()
        .sheets()[0]
        .tables()
        .next()
        .unwrap()
        .get_cell(0, 1)
        .cloned();
    let info = editor
        .cell_conditional_highlighting(10, 0, 1)
        .unwrap()
        .unwrap();
    assert_eq!(info.list_identifier, 1);
    assert_eq!(info.style_set_object_id, 91);
    assert_eq!(info.rule_count, 1);

    editor
        .clear_cell_conditional_highlighting(10, 0, 1)
        .unwrap();
    assert!(
        editor
            .cell_conditional_highlighting(10, 0, 1)
            .unwrap()
            .is_none()
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets()[0].tables().next().unwrap().get_cell(0, 1),
        original_value.as_ref()
    );
    assert!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .object(91)
            .is_none()
    );
}

#[test]
fn table_sort_execution_rejects_unsupported_state_transactionally() {
    let mut no_order = NumbersEditor::from_package(test_package()).unwrap();
    let before = no_order.to_bytes().unwrap();
    assert!(
        no_order
            .apply_table_sort_order(TableSelector::index(0))
            .is_err()
    );
    assert_eq!(no_order.to_bytes().unwrap(), before);

    let mut spill_package = test_package_with_calculation_engine();
    add_test_spill_dependency(&mut spill_package);
    let mut spill_editor = NumbersEditor::from_package(spill_package).unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut spill_editor,
        10,
        [
            TableCellUpdate::new(0, 1, cell_number(3.0)),
            TableCellUpdate::new(1, 1, cell_number(1.0)),
            TableCellUpdate::new(2, 1, cell_number(4.0)),
            TableCellUpdate::new(3, 1, cell_number(2.0)),
        ],
    )
    .unwrap();
    spill_editor
        .set_table_sort_order(
            TableSelector::index(0),
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(1).unwrap(),
                NumbersTableSortDirection::Ascending,
            )])
            .unwrap(),
        )
        .unwrap();
    let before = spill_editor.to_bytes().unwrap();
    assert!(
        spill_editor
            .apply_table_sort_order(TableSelector::index(0))
            .is_err()
    );
    assert_eq!(spill_editor.to_bytes().unwrap(), before);

    let mut formula_editor =
        NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    formula_editor
        .set_formula(10, 0, 1, FormulaExpression::Number(3.0))
        .unwrap();
    crate::numbers::editor::apply_numbers_fixture(
        &mut formula_editor,
        10,
        [
            TableCellUpdate::new(1, 1, cell_number(1.0)),
            TableCellUpdate::new(2, 1, cell_number(4.0)),
            TableCellUpdate::new(3, 1, cell_number(2.0)),
        ],
    )
    .unwrap();
    formula_editor
        .set_table_sort_order(
            TableSelector::index(0),
            NumbersTableSortOrder::new([NumbersTableSortRule::new(
                NumbersTableSortColumnIndex::new(1).unwrap(),
                NumbersTableSortDirection::Ascending,
            )])
            .unwrap(),
        )
        .unwrap();
    let before = formula_editor.to_bytes().unwrap();
    assert!(
        formula_editor
            .apply_table_sort_order(TableSelector::index(0))
            .is_err()
    );
    assert_eq!(formula_editor.to_bytes().unwrap(), before);
}

#[test]
fn table_dimension_sizes_are_typed_transactional_and_wire_exact() {
    assert!(Points::new(0.0).is_err());
    assert!(Points::new(-1.0).is_err());
    assert!(Points::new(f32::INFINITY).is_err());
    assert!(Points::new(f32::NAN).is_err());

    let mut editor =
        NumbersEditor::from_package(test_package_with_column_headers_and_engine()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert_eq!(
        editor
            .table_row_height(test_table_selector(&editor, 10), 1)
            .unwrap(),
        Size::Default
    );
    assert_eq!(
        editor
            .table_column_width(test_table_selector(&editor, 10), 2)
            .unwrap(),
        Size::Default
    );
    let row_height = Size::points(32.0).unwrap();
    let column_width = Size::points(124.0).unwrap();

    editor
        .set_table_row_height(test_table_selector(&editor, 10), 1, row_height)
        .unwrap();
    editor
        .set_table_column_width(test_table_selector(&editor, 10), 2, column_width)
        .unwrap();

    assert_eq!(
        editor
            .table_row_height(test_table_selector(&editor, 10), 1)
            .unwrap(),
        row_height
    );
    assert_eq!(
        editor
            .table_column_width(test_table_selector(&editor, 10), 2)
            .unwrap(),
        column_width
    );
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let rows =
        tst::HeaderStorageBucket::decode(archive.object(42).unwrap().messages[0].data.as_slice())
            .unwrap();
    let columns =
        tst::HeaderStorageBucket::decode(archive.object(43).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(
        rows.headers
            .iter()
            .find(|header| header.index == 1)
            .unwrap()
            .size,
        32.0
    );
    assert_eq!(columns.headers[2].size, 124.0);

    let bytes = editor.to_bytes().unwrap();
    let reparsed = NumbersEditor::from_bytes(&bytes).unwrap();
    assert_eq!(
        reparsed
            .table_row_height(test_table_selector(&reparsed, 10), 1)
            .unwrap(),
        row_height
    );
    assert_eq!(
        reparsed
            .table_column_width(test_table_selector(&reparsed, 10), 2)
            .unwrap(),
        column_width
    );

    editor
        .set_table_row_height(test_table_selector(&editor, 10), 1, Size::Default)
        .unwrap();
    editor
        .set_table_column_width(test_table_selector(&editor, 10), 2, Size::Default)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_table_row_height(test_table_selector(&editor, 10), 4, row_height)
            .is_err()
    );
    assert!(
        editor
            .set_table_column_width(test_table_selector(&editor, 10), 4, column_width)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn table_dimension_size_preserves_unknown_header_fields() {
    let mut package = test_package_with_column_headers_and_engine();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(43).unwrap();
            let message = object.messages[0].clone();
            let data = transform_length_delimited_fields_at_path(&message.data, &[2], |payload| {
                let mut payload = payload.to_vec();
                append_unknown_varint(&mut payload, 99, 990);
                Ok(payload)
            })?;
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = repeated_length_delimited_payloads(
        &editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .object(43)
            .unwrap()
            .messages[0]
            .data,
        2,
    )
    .unwrap()
    .into_iter()
    .map(ToOwned::to_owned)
    .collect::<Vec<_>>();

    editor
        .set_table_column_width(
            test_table_selector(&editor, 10),
            2,
            Size::points(124.0).unwrap(),
        )
        .unwrap();

    let document = editor.package().archive("Index/Document.iwa").unwrap();
    let after =
        repeated_length_delimited_payloads(&document.object(43).unwrap().messages[0].data, 2)
            .unwrap();
    assert_eq!(after[0], before[0]);
    assert_eq!(after[1], before[1]);
    let mut unknown = Vec::new();
    append_unknown_varint(&mut unknown, 99, 990);
    assert!(after[2].ends_with(&unknown));
    assert_eq!(
        tst::header_storage_bucket::Header::decode(after[2])
            .unwrap()
            .size,
        124.0
    );
}

#[test]
fn table_dimension_size_rejects_malformed_headers_transactionally() {
    let mut package = test_package_with_column_headers_and_engine();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(43).unwrap();
            let message = object.messages[0].clone();
            let mut bucket = tst::HeaderStorageBucket::decode(message.data.as_slice())?;
            bucket.headers.push(bucket.headers[1]);
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: bucket.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .table_column_width(test_table_selector(&editor, 10), 1)
            .is_err()
    );
    assert!(
        editor
            .set_table_column_width(
                test_table_selector(&editor, 10),
                1,
                Size::points(80.0).unwrap(),
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn removes_table_from_owning_sheet_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let removed = editor
        .remove_table(test_table_selector(&editor, 10))
        .unwrap();
    assert_eq!(removed.name, "Table 1");
    assert!(editor.tables().unwrap().is_empty());
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets();
    assert_eq!(sheets.len(), 1);
    assert!(sheets[0].is_empty());
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .remove_table(test_table_selector(&editor, 10))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn removes_sheets_transactionally() {
    let mut editor = NumbersEditor::from_package(two_sheet_package()).unwrap();
    assert_eq!(
        editor
            .sheets()
            .unwrap()
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["Sheet 1", "Second"]
    );
    let removed = editor
        .remove_sheet(test_sheet_selector(&editor, 50))
        .unwrap();
    assert_eq!(removed.name, "Second");
    assert_eq!(editor.sheets().unwrap()[0].object_id, 2);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .remove_sheet(test_sheet_selector(&editor, 2))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn moves_populated_table_between_sheets_losslessly() {
    let mut package = two_sheet_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let source_sheet = archive.object_mut(2).unwrap();
            let source_message = source_sheet.messages[0].clone();
            let data = crate::wire::transform_length_delimited_fields_at_path(
                &source_message.data,
                &[2],
                |reference| {
                    let mut reference = reference.to_vec();
                    append_unknown_varint(&mut reference, 98, 980);
                    Ok(reference)
                },
            )?;
            source_sheet.replace_message(
                0,
                RawMessage {
                    type_: source_message.type_,
                    data,
                },
            )?;
            source_sheet.archive_info.message_infos[0].object_references = vec![3];

            let target_sheet = archive.object_mut(50).unwrap();
            target_sheet.archive_info.message_infos[0]
                .object_references
                .clear();

            let table_info = archive.object_mut(3).unwrap();
            let mut info = tst::TableInfoArchive::decode(table_info.messages[0].data.as_slice())?;
            info.super_.parent = Some(Reference {
                identifier: 2,
                ..Default::default()
            });
            let mut data = info.encode_to_vec();
            data = crate::wire::transform_length_delimited_fields_at_path(
                &data,
                &[1, 2],
                |reference| {
                    let mut reference = reference.to_vec();
                    append_unknown_varint(&mut reference, 97, 970);
                    Ok(reference)
                },
            )?;
            append_unknown_varint(&mut data, 99, 990);
            table_info.replace_message(0, RawMessage { type_: 6003, data })?;
            table_info.archive_info.message_infos[0].object_references = vec![2, 10];
            Ok(())
        })
        .unwrap();
    let baseline = package.to_bytes().unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();

    let moved = editor
        .move_table(TableSelector::index(0), SheetSelector::index(1))
        .unwrap();
    assert_eq!(moved.name, "Table 1");
    assert_eq!(find_table_owner(editor.package(), 10).unwrap().sheet_id, 50);
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let (_, source) = decode_sheet(archive.object(2).unwrap()).unwrap();
    let (_, target) = decode_sheet(archive.object(50).unwrap()).unwrap();
    assert!(source.drawable_infos.is_empty());
    assert_eq!(
        target
            .drawable_infos
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [3]
    );
    let target_reference = crate::wire::repeated_length_delimited_payloads(
        &archive.object(50).unwrap().messages[0].data,
        2,
    )
    .unwrap()[0];
    assert!(
        target_reference
            .windows(2)
            .any(|window| window == [0x90, 0x06])
    );
    let table_data = &archive.object(3).unwrap().messages[0].data;
    let table_info = tst::TableInfoArchive::decode(table_data.as_slice()).unwrap();
    assert_eq!(table_info.super_.parent.unwrap().identifier, 50);
    let mut table_unknown = Vec::new();
    append_unknown_varint(&mut table_unknown, 99, 990);
    assert!(table_data.ends_with(&table_unknown));
    assert!(table_data.windows(2).any(|window| window == [0x88, 0x06]));

    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        10,
        0,
        0,
        CellValue::Text("Moved cell".to_owned()),
    )
    .unwrap();
    assert_eq!(
        compatibility_document_from_bytes(&editor.to_bytes().unwrap())
            .unwrap()
            .sheets()[1]
            .tables()
            .next()
            .unwrap()
            .get_cell(0, 0),
        Some(&CellValue::Text("Moved cell".to_owned()))
    );

    let mut editor = NumbersEditor::from_bytes(&baseline).unwrap();
    editor
        .move_table(TableSelector::index(0), SheetSelector::index(1))
        .unwrap();
    editor
        .move_table(TableSelector::index(0), SheetSelector::index(0))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .move_table(TableSelector::index(1), SheetSelector::index(1))
            .is_err()
    );
    assert!(
        editor
            .move_table(TableSelector::index(0), SheetSelector::index(2))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn sheet_add_remove_preserves_raw_references_and_restores_exact_component() {
    let mut package = two_sheet_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(1).unwrap();
            let message = object.messages[0].clone();
            let replacements = crate::wire::repeated_length_delimited_payloads(&message.data, 1)?
                .into_iter()
                .enumerate()
                .map(|(index, payload)| {
                    let mut payload = payload.to_vec();
                    append_unknown_varint(&mut payload, 98, 980 + index as u64);
                    payload
                })
                .collect::<Vec<_>>();
            let mut data = crate::wire::rewrite_repeated_length_delimited_fields(
                &message.data,
                1,
                &replacements,
            )?;
            append_unknown_varint(&mut data, 99, 990);
            Ok(object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )
                .map(|_| ())?)
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let baseline = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();

    let created = editor.add_empty_sheet("Temporary").unwrap();
    editor
        .remove_sheet(test_sheet_selector(&editor, created.object_id))
        .unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        baseline
    );
}

#[test]
fn creates_empty_sheet_with_unique_object_id() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let created = editor.add_empty_sheet("Created 東京").unwrap();
    assert_eq!(created.name, "Created 東京");
    assert_eq!(created.index, 1);
    let sheets = editor.sheets().unwrap();
    assert_eq!(sheets.len(), 2);
    assert_eq!(sheets[1].object_id, created.object_id);
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(document.sheets()[1].name(), "Created 東京");
}

#[test]
fn detached_table_models_are_not_exposed_or_writable() {
    let mut editor = NumbersEditor::from_package(two_sheet_package()).unwrap();
    editor
        .remove_sheet(test_sheet_selector(&editor, 2))
        .unwrap();
    assert!(editor.tables().unwrap().is_empty());

    let before = editor.to_bytes().unwrap();
    assert!(
        crate::numbers::editor::test_set_cell(&mut editor, 10, 0, 0, cell_number(1.0)).is_err()
    );
    assert!(
        editor
            .resize_table(test_table_selector(&editor, 10), 5, 5)
            .is_err()
    );
    assert!(
        editor
            .remove_table(test_table_selector(&editor, 10))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn creates_independent_empty_table_on_an_existing_sheet() {
    let mut package = two_sheet_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            object.archive_info.message_infos[0].object_references = vec![20, 999];
            object.archive_info.message_infos[0]
                .field_infos
                .push(FieldInfo {
                    path: FieldPath { path: vec![4, 4] },
                    object_references: vec![20, 999],
                    ..Default::default()
                });
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package.clone()).unwrap();
    let created = editor
        .add_empty_table(test_sheet_selector(&editor, 50), "Created Table", 3, 2)
        .unwrap();
    let mut repeated = NumbersEditor::from_package(package).unwrap();
    repeated
        .add_empty_table(test_sheet_selector(&repeated, 50), "Created Table", 3, 2)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), repeated.to_bytes().unwrap());
    assert_ne!(created.object_id, 10);
    assert_eq!((created.rows, created.columns), (3, 2));
    assert_eq!(created.name, "Created Table");
    assert_eq!(editor.tables().unwrap().len(), 2);
    let locations = object_locations(editor.package()).unwrap();
    let component_name = locations[&created.object_id].clone();
    let component = editor.package().archive(&component_name).unwrap();
    let model_object = component.object(created.object_id).unwrap();
    let cloned_model = TableModelArchive::decode(model_object.messages[0].data.as_slice()).unwrap();
    assert_eq!(
        model_object.archive_info.message_infos[0].field_infos[0].object_references,
        [cloned_model.base_data_store.string_table.identifier]
    );

    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        created.object_id,
        0,
        0,
        CellValue::Text("Independent".to_owned()),
    )
    .unwrap();
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets();
    assert_eq!(
        sheets[0]
            .tables()
            .next()
            .unwrap()
            .get_cell(0, 1)
            .unwrap()
            .as_text(),
        "Original"
    );
    assert_eq!(
        sheets[1]
            .tables()
            .next()
            .unwrap()
            .get_cell(0, 0)
            .unwrap()
            .as_text(),
        "Independent"
    );

    editor
        .remove_table(test_table_selector(&editor, created.object_id))
        .unwrap();
    assert_eq!(editor.tables().unwrap().len(), 1);
    assert!(
        editor
            .package()
            .archive(&component_name)
            .unwrap()
            .object(created.object_id)
            .is_none()
    );
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .add_empty_table(test_sheet_selector(&editor, 999), "Missing", 2, 2)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn recreates_first_table_after_removing_the_last_scratch_table() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 2)
        .build()
        .unwrap();
    let sheet_id = editor.sheets().unwrap()[0].object_id;
    let original_table_id = editor.tables().unwrap()[0].object_id;
    editor
        .remove_table(test_table_selector(&editor, original_table_id))
        .unwrap();
    let tableless = editor.to_bytes().unwrap();

    let created = editor
        .add_empty_table(
            test_sheet_selector(&editor, sheet_id),
            "First runtime",
            3,
            2,
        )
        .unwrap();
    assert_eq!((created.rows, created.columns), (3, 2));
    assert_eq!(created.name, "First runtime");
    crate::numbers::editor::set_cell_fixture(
        &mut editor,
        created.object_id,
        2,
        1,
        CellValue::Text("bootstrapped".to_owned()),
    )
    .unwrap();

    let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let document = compatibility_document_from_bytes(&reopened.to_bytes().unwrap()).unwrap();
    let table = document.sheets()[0].tables().next().unwrap();
    assert_eq!(
        table.get_cell(2, 1),
        Some(&CellValue::Text("bootstrapped".to_owned()))
    );
    reopened
        .remove_table(test_table_selector(&reopened, created.object_id))
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), tableless);
}

#[test]
fn first_table_bootstrap_uses_the_target_sheet() {
    let mut editor = NumbersDocumentBuilder::new().build().unwrap();
    let original_table_id = editor.tables().unwrap()[0].object_id;
    editor
        .remove_table(test_table_selector(&editor, original_table_id))
        .unwrap();
    let target = editor.add_empty_sheet("Target").unwrap();

    let created = editor
        .add_empty_table(
            test_sheet_selector(&editor, target.object_id),
            "Target table",
            2,
            3,
        )
        .unwrap();
    assert_eq!(
        find_table_owner(editor.package(), created.object_id)
            .unwrap()
            .sheet_id,
        target.object_id
    );
    let document = compatibility_document_from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets();
    assert!(sheets[0].is_empty());
    assert_eq!(sheets[1].tables().next().unwrap().name(), "Target table");
}

#[test]
fn first_table_bootstrap_rejects_a_missing_theme_preset_transactionally() {
    let mut editor = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = editor.sheets().unwrap()[0].object_id;
    let original_table_id = editor.tables().unwrap()[0].object_id;
    editor
        .remove_table(test_table_selector(&editor, original_table_id))
        .unwrap();
    editor
        .package
        .update_archive("Index/Document.iwa", |archive| {
            archive
                .remove_object(20)
                .ok_or_else(|| Error::InvalidFormat("test table preset is missing".to_owned()))?;
            Ok(())
        })
        .unwrap();
    let before = editor.to_bytes().unwrap();

    assert!(
        editor
            .add_empty_table(
                test_sheet_selector(&editor, sheet_id),
                "Missing preset",
                2,
                2
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn form_sheet_table_create_delete_restores_unknown_reference_bytes() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(2).unwrap();
            let sheet = tn::SheetArchive::decode(object.messages[0].data.as_slice()).unwrap();
            let mut data = tn::FormBasedSheetArchive {
                super_: sheet,
                ..Default::default()
            }
            .encode_to_vec();
            data = crate::wire::transform_length_delimited_field(&data, 1, |nested| {
                let payloads = crate::wire::repeated_length_delimited_payloads(nested, 2)?;
                let replacements = payloads
                    .into_iter()
                    .map(|payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, 97, 970);
                        payload
                    })
                    .collect::<Vec<_>>();
                let mut nested = crate::wire::rewrite_repeated_length_delimited_fields(
                    nested,
                    2,
                    &replacements,
                )?;
                append_unknown_varint(&mut nested, 98, 980);
                Ok(nested)
            })?;
            append_unknown_varint(&mut data, 99, 990);
            Ok(object
                .replace_message(0, RawMessage { type_: 3, data })
                .map(|_| ())?)
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let baseline = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    let baseline_entries = editor
        .package()
        .entry_names()
        .map(str::to_owned)
        .collect::<Vec<_>>();

    let created = editor
        .add_empty_table(test_sheet_selector(&editor, 2), "Temporary", 2, 2)
        .unwrap();
    editor
        .remove_table(test_table_selector(&editor, created.object_id))
        .unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        baseline
    );
    assert_eq!(
        editor
            .package()
            .entry_names()
            .map(str::to_owned)
            .collect::<Vec<_>>(),
        baseline_entries
    );
}

fn move_table_data_list_entries_to_segment(
    package: &mut IWorkPackage,
    table_id: u64,
    segment_id: u64,
) {
    package
        .update_archive("Index/Document.iwa", |archive| {
            let (list_type, entries, references) = {
                let object = archive.object_mut(table_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == 6005 || message.type_ == 6201)
                    .unwrap();
                let message_type = object.messages[message_index].type_;
                let mut list =
                    TableDataList::decode(object.messages[message_index].data.as_slice())?;
                let entries = std::mem::take(&mut list.entries);
                let references = entries
                    .iter()
                    .flat_map(entry_object_references)
                    .collect::<HashSet<_>>();
                list.segments.push(Reference {
                    identifier: segment_id,
                    ..Default::default()
                });
                let list_type = list.list_type;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data: list.encode_to_vec(),
                    },
                )?;
                for reference in &references {
                    remove_message_object_reference(object, message_index, *reference);
                }
                object.archive_info.message_infos[message_index]
                    .object_references
                    .push(segment_id);
                (list_type, entries, references)
            };
            let key_range = segment_key_range(&entries)?;
            let mut segment_object = ArchiveObject::new(
                segment_id,
                vec![RawMessage {
                    type_: 6011,
                    data: TableDataListSegment {
                        list_type,
                        key_range,
                        entries,
                    }
                    .encode_to_vec(),
                }],
            )?;
            segment_object.archive_info.message_infos[0]
                .object_references
                .extend(references);
            Ok(archive.insert_object(segment_object)?)
        })
        .unwrap();
}

fn test_package_with_formula_error() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let model_object = archive.object_mut(10).unwrap();
            let model_type = model_object.messages[0].type_;
            let mut model = TableModelArchive::decode(model_object.messages[0].data.as_slice())?;
            model.base_data_store.formula_error_table = Some(Reference {
                identifier: 22,
                ..Default::default()
            });
            model_object.replace_message(
                0,
                RawMessage {
                    type_: model_type,
                    data: model.encode_to_vec(),
                },
            )?;
            model_object.archive_info.message_infos[0]
                .object_references
                .push(22);

            let errors = TableDataList {
                list_type: tst::table_data_list::ListType::FormulaError as i32,
                next_list_id: 8,
                entries: vec![tst::table_data_list::ListEntry {
                    key: 7,
                    refcount: 1,
                    string: Some("Syntax Error".to_owned()),
                    ..Default::default()
                }],
                segments: Vec::new(),
                is_new_for_bnc: Some(true),
            };
            archive.insert_object(ArchiveObject::new(
                22,
                vec![RawMessage {
                    type_: 6005,
                    data: errors.encode_to_vec(),
                }],
            )?)?;

            let tile_object = archive.object_mut(30).unwrap();
            let tile_type = tile_object.messages[0].type_;
            let mut tile = Tile::decode(tile_object.messages[0].data.as_slice())?;
            let mut cells = split_row(&tile.row_infos[0])?;
            cells[1] = Some(vec![
                5, 8, 0, 0, 0, 0, 0, 0, // version and formula-error type
                0, 8, 0, 0, // formula-error identifier flag
                7, 0, 0, 0,
            ]);
            rebuild_row(&mut tile.row_infos[0], &cells)?;
            tile_object.replace_message(
                0,
                RawMessage {
                    type_: tile_type,
                    data: tile.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    package
}

fn test_package_with_comments(shared: bool) -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let model_object = archive.object_mut(10).unwrap();
            let model_type = model_object.messages[0].type_;
            let mut model = TableModelArchive::decode(model_object.messages[0].data.as_slice())?;
            model.base_data_store.comment_storage_table = Some(Reference {
                identifier: 60,
                ..Default::default()
            });
            model_object.replace_message(
                0,
                RawMessage {
                    type_: model_type,
                    data: model.encode_to_vec(),
                },
            )?;
            model_object.archive_info.message_infos[0]
                .object_references
                .push(60);

            let tile_object = archive.object_mut(30).unwrap();
            let tile_type = tile_object.messages[0].type_;
            let mut tile = Tile::decode(tile_object.messages[0].data.as_slice())?;
            let mut cells = split_row(&tile.row_infos[0])?;
            let mut first = BncCell::parse(cells[1].as_deref().unwrap())?;
            first.set_comment_identifier(Some(4));
            cells[1] = Some(first.encode());
            if shared {
                let mut second = BncCell::minimal();
                second.set_number(7.0)?;
                second.set_comment_identifier(Some(4));
                cells[2] = Some(second.encode());
            }
            rebuild_row(&mut tile.row_infos[0], &cells)?;
            tile_object.replace_message(
                0,
                RawMessage {
                    type_: tile_type,
                    data: tile.encode_to_vec(),
                },
            )?;

            let list = TableDataList {
                list_type: tst::table_data_list::ListType::CommentStorage as i32,
                next_list_id: 5,
                entries: vec![tst::table_data_list::ListEntry {
                    key: 4,
                    refcount: if shared { 2 } else { 1 },
                    comment_storage: Some(Reference {
                        identifier: 61,
                        ..Default::default()
                    }),
                    ..Default::default()
                }],
                segments: Vec::new(),
                is_new_for_bnc: Some(true),
            };
            let mut list_object = ArchiveObject::new(
                60,
                vec![RawMessage {
                    type_: 6005,
                    data: list.encode_to_vec(),
                }],
            )?;
            list_object.archive_info.message_infos[0]
                .object_references
                .push(61);
            archive.insert_object(list_object)?;
            archive.insert_object(ArchiveObject::new(
                61,
                vec![RawMessage {
                    type_: 3056,
                    data: tsd::CommentStorageArchive {
                        text: Some("Original comment".to_owned()),
                        creation_date: Some(crate::protobuf::tsp::Date { seconds: 123.5 }),
                        replies: vec![Reference {
                            identifier: 70,
                            ..Default::default()
                        }],
                        storage_uuid: Some(crate::protobuf::tsp::Uuid {
                            lower: 61,
                            upper: 62,
                        }),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            archive.object_mut(61).unwrap().archive_info.message_infos[0]
                .object_references
                .push(70);
            archive.insert_object(ArchiveObject::new(
                70,
                vec![RawMessage {
                    type_: 3056,
                    data: tsd::CommentStorageArchive {
                        text: Some("Reply".to_owned()),
                        storage_uuid: Some(crate::protobuf::tsp::Uuid {
                            lower: 70,
                            upper: 71,
                        }),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    package
}

fn add_unknown_comment_storage_field(package: &mut IWorkPackage, storage_id: u64) -> Vec<u8> {
    let mut unknown = litchi_iwa_common::varint::encode_varint(99 << 3);
    unknown.extend(litchi_iwa_common::varint::encode_varint(999));
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(storage_id).unwrap();
            let mut data = object.messages[0].data.clone();
            data.extend_from_slice(&unknown);
            object.replace_message(0, RawMessage { type_: 3056, data })?;
            Ok(())
        })
        .unwrap();
    unknown
}

fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
    data.extend(litchi_iwa_common::varint::encode_varint(
        u64::from(field_number) << 3,
    ));
    data.extend(litchi_iwa_common::varint::encode_varint(value));
}

#[allow(deprecated)]
fn test_package_with_text_box() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let sheet = archive.object_mut(2).unwrap();
            let mut decoded = tn::SheetArchive::decode(sheet.messages[0].data.as_slice())?;
            decoded.drawable_infos.push(Reference {
                identifier: 50,
                ..Default::default()
            });
            sheet.replace_message(
                0,
                RawMessage {
                    type_: 2,
                    data: decoded.encode_to_vec(),
                },
            )?;
            sheet.archive_info.message_infos[0]
                .object_references
                .push(50);

            let shape = tswp::ShapeInfoArchive {
                super_: tsd::ShapeArchive {
                    super_: tsd::DrawableArchive {
                        geometry: Some(tsd::GeometryArchive {
                            position: Some(crate::protobuf::tsp::Point { x: 100.0, y: 60.0 }),
                            size: Some(crate::protobuf::tsp::Size {
                                width: 200.0,
                                height: 60.0,
                            }),
                            flags: Some(0),
                            angle: Some(0.0),
                        }),
                        parent: Some(Reference {
                            identifier: 2,
                            ..Default::default()
                        }),
                        locked: Some(false),
                        aspect_ratio_locked: Some(false),
                        title: Some(Reference {
                            identifier: 52,
                            ..Default::default()
                        }),
                        caption: Some(Reference {
                            identifier: 51,
                            ..Default::default()
                        }),
                        ..Default::default()
                    },
                    ..Default::default()
                },
                deprecated_storage: Some(Reference {
                    identifier: 53,
                    ..Default::default()
                }),
                owned_storage: Some(Reference {
                    identifier: 53,
                    ..Default::default()
                }),
                is_text_box: Some(true),
                ..Default::default()
            };
            let mut shape_object = ArchiveObject::new(
                50,
                vec![RawMessage {
                    type_: SHAPE_INFO_MESSAGE_TYPE,
                    data: shape.encode_to_vec(),
                }],
            )?;
            shape_object.archive_info.message_infos[0].object_references = vec![2, 51, 52, 53];
            archive.insert_object(shape_object)?;
            for identifier in [51, 52] {
                archive.insert_object(ArchiveObject::new(
                    identifier,
                    vec![RawMessage {
                        type_: STANDIN_CAPTION_MESSAGE_TYPE,
                        data: Vec::new(),
                    }],
                )?)?;
            }
            archive.insert_object(ArchiveObject::new(
                53,
                vec![RawMessage {
                    type_: 2_001,
                    data: tswp::StorageArchive {
                        text: vec!["Source".to_owned()],
                        in_document: Some(true),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    package
}

#[test]
fn numbers_object_catalog_is_bounded_and_measured() {
    let package = test_package_with_text_box();
    let archive_count = package.iwa_entry_names().count();
    let mut catalog = NumbersObjectCatalog::build(&package).unwrap();
    let storage = catalog
        .text_storage_info(&package, TextStorageId::new(53).unwrap())
        .unwrap();
    assert_eq!(storage.storage.text(), "Source");
    let stats = catalog.stats();

    assert_eq!(stats.archives_scanned, archive_count);
    assert_eq!(stats.archive_reads, archive_count + 2);
    assert_eq!(stats.sheet_objects_scanned, 1);
    assert!(stats.drawable_objects_scanned >= 1);
    assert_eq!(stats.drawable_objects_scanned, stats.reference_edges);
    assert_eq!(stats.semantic_decodes, 3);
    assert_eq!(stats.peak_live_archives, 1);
    assert_eq!(stats.retained_payload_bytes, 0);

    let graph = catalog.text_box_graph(&package, 2, 50).unwrap();
    assert_eq!(graph.storage_id, TextStorageId::new(53).unwrap());
    assert_eq!(catalog.stats(), stats);
}

#[test]
fn numbers_object_catalog_rejects_later_duplicate_ids_and_malformed_archives() {
    let mut duplicate = test_package_with_text_box();
    let original = duplicate.archive("Index/Document.iwa").unwrap();
    duplicate
        .replace_archive("Index/Later.iwa", &original)
        .unwrap();
    let error = NumbersObjectCatalog::build(&duplicate).unwrap_err();
    assert!(
        error
            .to_string()
            .contains("appears in both Numbers archives")
    );
    assert!(
        NumbersEditor::from_package(duplicate)
            .unwrap()
            .sheet_text_boxes(2)
            .is_err()
    );

    let mut malformed = test_package_with_text_box();
    let compressed = crate::snappy::SnappyStream::compress(&[0x80]).unwrap();
    malformed
        .insert_entry("Index/Later.iwa", compressed)
        .unwrap();
    assert!(NumbersObjectCatalog::build(&malformed).is_err());
    assert!(
        NumbersEditor::from_package(malformed)
            .unwrap()
            .sheet_text_boxes(2)
            .is_err()
    );
}

#[test]
fn numbers_object_catalog_rejects_stale_copy_on_write_generations() {
    let package = test_package_with_text_box();
    let catalog = NumbersObjectCatalog::build(&package).unwrap();
    let mut edited = package.clone();
    edited
        .insert_entry("Metadata/catalog-revision", Vec::new())
        .unwrap();

    assert!(catalog.text_box_graph(&package, 2, 50).is_ok());
    let error = catalog.text_box_graph(&edited, 2, 50).unwrap_err();
    assert!(error.to_string().contains("catalog is stale"));
}

#[test]
fn numbers_object_catalog_enforces_operation_archive_budget() {
    let package = test_package_with_text_box();
    let limits = NumbersObjectCatalogLimits {
        max_archive_reads: 1,
        ..Default::default()
    };
    let error = NumbersObjectCatalog::build_with_limits(&package, limits).unwrap_err();
    assert!(error.to_string().contains("archive reads"));
}

fn test_package_with_text_box_metadata() -> IWorkPackage {
    let mut package = test_package_with_text_box();
    let metadata = PackageMetadata {
        last_object_identifier: 70,
        components: vec![ComponentInfo {
            identifier: DOCUMENT_COMPONENT_IDENTIFIER,
            preferred_locator: "Document".to_owned(),
            object_uuid_map_entries: [50, 51, 52, 53]
                .into_iter()
                .map(|identifier| ObjectUuidMapEntry {
                    identifier,
                    uuid: Uuid {
                        lower: identifier,
                        upper: identifier + 1_000,
                    },
                })
                .collect(),
            ..Default::default()
        }],
        ..Default::default()
    };
    package
        .replace_archive(
            PACKAGE_METADATA_ENTRY,
            &Archive {
                objects: vec![
                    ArchiveObject::new(
                        70,
                        vec![RawMessage {
                            type_: PACKAGE_METADATA_MESSAGE_TYPE,
                            data: metadata.encode_to_vec(),
                        }],
                    )
                    .unwrap(),
                ],
            },
        )
        .unwrap();
    package
}

#[derive(Clone, Copy)]
enum TestStrokeLayerSide {
    Left,
    Top,
}

impl TestStrokeLayerSide {
    fn references_mut(self, sidecar: &mut tst::StrokeSidecarArchive) -> &mut Vec<Reference> {
        match self {
            Self::Left => &mut sidecar.left_column_stroke_layers,
            Self::Top => &mut sidecar.top_row_stroke_layers,
        }
    }
}

fn add_test_stroke_layer(
    package: &mut IWorkPackage,
    identifier: u64,
    side: TestStrokeLayerSide,
    row_column_index: u32,
    runs: &[(i32, u32)],
) {
    const STROKE_LAYER_MESSAGE_TYPE: u32 = 6_306;
    package
        .update_archive("Index/Document.iwa", |archive| {
            {
                let sidecar = archive.object_mut(41).unwrap();
                let message_index = sidecar
                    .messages
                    .iter()
                    .position(|message| {
                        tst::StrokeSidecarArchive::decode(message.data.as_slice()).is_ok()
                    })
                    .unwrap();
                let message = sidecar.messages[message_index].clone();
                let mut payload =
                    tst::StrokeSidecarArchive::decode(message.data.as_slice()).unwrap();
                side.references_mut(&mut payload).push(Reference {
                    identifier,
                    ..Default::default()
                });
                sidecar.replace_message(
                    message_index,
                    RawMessage {
                        type_: message.type_,
                        data: payload.encode_to_vec(),
                    },
                )?;
                sidecar.archive_info.message_infos[message_index]
                    .object_references
                    .push(identifier);
            }
            Ok(archive.insert_object(ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: STROKE_LAYER_MESSAGE_TYPE,
                    data: tst::StrokeLayerArchive {
                        row_column_index: Some(row_column_index),
                        stroke_runs: runs
                            .iter()
                            .map(
                                |&(origin, length)| tst::stroke_layer_archive::StrokeRunArchive {
                                    origin: Some(origin),
                                    length: Some(length),
                                    order: Some(1),
                                    ..Default::default()
                                },
                            )
                            .collect(),
                    }
                    .encode_to_vec(),
                }],
            )?)?)
        })
        .unwrap();
}

fn test_stroke_sidecar(package: &IWorkPackage) -> tst::StrokeSidecarArchive {
    let archive = package.archive("Index/Document.iwa").unwrap();
    archive
        .object(41)
        .unwrap()
        .messages
        .iter()
        .find_map(|message| tst::StrokeSidecarArchive::decode(message.data.as_slice()).ok())
        .unwrap()
}

fn test_stroke_layer(package: &IWorkPackage, identifier: u64) -> tst::StrokeLayerArchive {
    let archive = package.archive("Index/Document.iwa").unwrap();
    archive
        .object(identifier)
        .unwrap()
        .messages
        .iter()
        .find_map(|message| tst::StrokeLayerArchive::decode(message.data.as_slice()).ok())
        .unwrap()
}

fn add_test_conditional_style_storage(package: &mut IWorkPackage, has_entries: bool) {
    const CONDITIONAL_STYLE_TABLE_ID: u64 = 90;
    const CONDITIONAL_STYLE_SET_ID: u64 = 91;
    package
        .update_archive("Index/Document.iwa", |archive| {
            let model_object = archive.object_mut(10).unwrap();
            let message = model_object.messages[0].clone();
            let mut model = TableModelArchive::decode(message.data.as_slice())?;
            model.conditional_style_formula_owner_id = Some(tsp::CfuuidArchive {
                uuid_w0: Some(1),
                uuid_w1: Some(2),
                uuid_w2: Some(3),
                uuid_w3: Some(4),
                ..Default::default()
            });
            model.base_data_store.conditionalstyletable = Some(Reference {
                identifier: CONDITIONAL_STYLE_TABLE_ID,
                ..Default::default()
            });
            model_object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: model.encode_to_vec(),
                },
            )?;

            let entries = has_entries
                .then_some(tst::table_data_list::ListEntry {
                    key: 1,
                    refcount: 1,
                    reference: Some(Reference {
                        identifier: CONDITIONAL_STYLE_SET_ID,
                        ..Default::default()
                    }),
                    ..Default::default()
                })
                .into_iter()
                .collect();
            archive.insert_object(ArchiveObject::new(
                CONDITIONAL_STYLE_TABLE_ID,
                vec![RawMessage {
                    type_: 6_005,
                    data: TableDataList {
                        list_type: tst::table_data_list::ListType::ConditionalStyle as i32,
                        next_list_id: if has_entries { 2 } else { 1 },
                        entries,
                        segments: Vec::new(),
                        is_new_for_bnc: Some(true),
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            if has_entries {
                archive.insert_object(ArchiveObject::new(
                    CONDITIONAL_STYLE_SET_ID,
                    vec![RawMessage {
                        type_: 6_010,
                        data: tst::ConditionalStyleSetArchive {
                            rule_count: 1,
                            ..Default::default()
                        }
                        .encode_to_vec(),
                    }],
                )?)?;
            }
            Ok(())
        })
        .unwrap();
}

fn add_test_app_native_topology_allocations(package: &mut IWorkPackage) {
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message = object.messages[0].clone();
            let mut model = TableModelArchive::decode(message.data.as_slice())?;
            model.category_owner_deprecated = Some(tst::CategoryOwnerArchive {
                owner_uid: tsp::Uuid { lower: 1, upper: 2 },
                group_by: vec![tst::GroupByArchive {
                    group_by_uid: tsp::Uuid { lower: 3, upper: 4 },
                    is_enabled: false,
                    ..Default::default()
                }],
            });
            model.spill_owner = Some(tsce::SpillOwnerArchive {
                owner_uid: tsp::Uuid { lower: 5, upper: 6 },
            });
            Ok(object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data: model.encode_to_vec(),
                    },
                )
                .map(|_| ())?)
        })
        .unwrap();
}

fn add_test_spill_dependency(package: &mut IWorkPackage) {
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let object = archive.object_mut(101).unwrap();
            let message = object.messages[0].clone();
            let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            owner.spill_range_sizes = Some(tsce::CellSpillSizesArchive {
                spills: vec![tsce::cell_spill_sizes_archive::SpillForCell {
                    coordinate: tsce::CellCoordinateArchive {
                        column: Some(1),
                        row: Some(1),
                        ..Default::default()
                    },
                    spill_size: tsce::ColumnRowSize {
                        num_columns: Some(1),
                        num_rows: Some(2),
                    },
                }],
            });
            Ok(object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data: owner.encode_to_vec(),
                    },
                )
                .map(|_| ())?)
        })
        .unwrap();
}

fn test_package() -> IWorkPackage {
    let mut cell = BncCell::minimal();
    cell.set_string(1);
    let mut cells = vec![None; 4];
    cells[1] = Some(cell.encode());
    let (storage, offsets, wide) = encode_row(&cells, false).unwrap();
    let tile = Tile {
        max_column: 0,
        max_row: 0,
        num_cells: 0,
        numrows: 1,
        row_infos: vec![TileRowInfo {
            tile_row_index: 0,
            cell_count: 1,
            cell_storage_buffer_pre_bnc: Vec::new(),
            cell_offsets_pre_bnc: Vec::new(),
            storage_version: Some(5),
            cell_storage_buffer: Some(storage),
            cell_offsets: Some(offsets),
            has_wide_offsets: Some(wide),
        }],
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        should_use_wide_rows: None,
    };

    let strings = TableDataList {
        list_type: tst::table_data_list::ListType::String as i32,
        next_list_id: 2,
        entries: vec![tst::table_data_list::ListEntry {
            key: 1,
            refcount: 1,
            string: Some("Original".to_string()),
            reference: None,
            formula: None,
            format: None,
            custom_format: None,
            rich_text_payload: None,
            comment_storage: None,
            import_warning_set: None,
            cell_spec: None,
        }],
        segments: Vec::new(),
        is_new_for_bnc: Some(true),
    };
    let formulas = TableDataList {
        list_type: tst::table_data_list::ListType::Formula as i32,
        next_list_id: 1,
        entries: Vec::new(),
        segments: Vec::new(),
        is_new_for_bnc: Some(true),
    };

    let model = TableModelArchive {
        table_id: "00000000-0000-0002-0000-000000000001".to_string(),
        table_name: "Table 1".to_string(),
        number_of_rows: 4,
        number_of_columns: 4,
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                buckets: vec![Reference {
                    identifier: 42,
                    ..Default::default()
                }],
            },
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: Reference {
                        identifier: 30,
                        ..Default::default()
                    },
                }],
                tile_size: Some(256),
                should_use_wide_rows: None,
            },
            string_table: Reference {
                identifier: 20,
                ..Default::default()
            },
            formula_table: Reference {
                identifier: 21,
                ..Default::default()
            },
            ..Default::default()
        },
        base_column_row_uids: Some(Reference {
            identifier: 40,
            ..Default::default()
        }),
        stroke_sidecar: Some(Reference {
            identifier: 41,
            ..Default::default()
        }),
        ..Default::default()
    };

    let uid_map = tst::ColumnRowUidMapArchive {
        sorted_column_uids: (1..=4)
            .map(|value| crate::protobuf::tsp::Uuid {
                lower: value,
                upper: value,
            })
            .collect(),
        column_index_for_uid: vec![0, 1, 2, 3],
        column_uid_for_index: vec![0, 1, 2, 3],
        sorted_row_uids: (11..=14)
            .map(|value| crate::protobuf::tsp::Uuid {
                lower: value,
                upper: value,
            })
            .collect(),
        row_index_for_uid: vec![0, 1, 2, 3],
        row_uid_for_index: vec![0, 1, 2, 3],
    };
    let stroke_sidecar = tst::StrokeSidecarArchive {
        row_count: Some(4),
        column_count: Some(4),
        ..Default::default()
    };
    let row_headers = tst::HeaderStorageBucket {
        bucket_hash_function: 1,
        headers: vec![tst::header_storage_bucket::Header {
            index: 0,
            size: 0.0,
            hiding_state: 0,
            number_of_cells: 1,
            cell_style: None,
            text_style: None,
        }],
    };

    let object = |identifier, message_type, data| {
        ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: message_type,
                data,
            }],
        )
        .unwrap()
    };
    let archive = Archive {
        objects: vec![
            object(
                1,
                1,
                tn::DocumentArchive {
                    sheets: vec![Reference {
                        identifier: 2,
                        ..Default::default()
                    }],
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            object(
                2,
                2,
                tn::SheetArchive {
                    name: "Sheet 1".to_owned(),
                    drawable_infos: vec![Reference {
                        identifier: 3,
                        ..Default::default()
                    }],
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            object(
                3,
                6003,
                tst::TableInfoArchive {
                    table_model: Reference {
                        identifier: 10,
                        ..Default::default()
                    },
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            object(10, 6000, model.encode_to_vec()),
            object(20, 6005, strings.encode_to_vec()),
            object(21, 6005, formulas.encode_to_vec()),
            object(30, 6002, tile.encode_to_vec()),
            object(40, 6200, uid_map.encode_to_vec()),
            object(41, 6200, stroke_sidecar.encode_to_vec()),
            object(42, 6004, row_headers.encode_to_vec()),
        ],
    };
    let mut package = IWorkPackage::new();
    package
        .replace_archive("Index/Document.iwa", &archive)
        .unwrap();
    package
}

fn test_package_with_calculation_engine() -> IWorkPackage {
    let mut package = test_package();
    let owner = tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid: crate::protobuf::tsp::Uuid {
            lower: 0x0200_0000_0000_0000,
            upper: 0x0100_0000_0000_0000,
        },
        internal_formula_owner_id: 0,
        owner_kind: Some(1),
        cell_dependencies: Some(tsce::CellDependenciesExpandedArchive::default()),
        formula_owner: Some(Reference {
            identifier: 3,
            ..Default::default()
        }),
        tiled_cell_dependencies: Some(tsce::CellDependenciesTiledArchive::default()),
        ..Default::default()
    };
    let engine = tsce::CalculationEngineArchive {
        dependency_tracker: tsce::DependencyTrackerArchive {
            number_of_formulas: Some(0),
            formula_owner_dependencies: vec![Reference {
                identifier: 101,
                ..Default::default()
            }],
            ..Default::default()
        },
        ..Default::default()
    };
    let mut archive = Archive {
        objects: vec![
            ArchiveObject::new(
                100,
                vec![RawMessage {
                    type_: 4000,
                    data: engine.encode_to_vec(),
                }],
            )
            .unwrap(),
            ArchiveObject::new(
                101,
                vec![RawMessage {
                    type_: 4008,
                    data: owner.encode_to_vec(),
                }],
            )
            .unwrap(),
        ],
    };
    archive.objects[0].archive_info.message_infos[0]
        .object_references
        .push(101);
    package
        .replace_archive("Index/CalculationEngine.iwa", &archive)
        .unwrap();
    package
}

fn test_package_with_column_headers_and_engine() -> IWorkPackage {
    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message = object.messages[0].clone();
            let mut model = TableModelArchive::decode(message.data.as_slice())?;
            model.base_data_store.column_headers = Reference {
                identifier: 43,
                ..Default::default()
            };
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data: model.encode_to_vec(),
                },
            )?;
            let headers = tst::HeaderStorageBucket {
                bucket_hash_function: 1,
                headers: (0..3)
                    .map(|index| tst::header_storage_bucket::Header {
                        index,
                        size: 0.0,
                        hiding_state: 0,
                        number_of_cells: 1,
                        cell_style: None,
                        text_style: None,
                    })
                    .collect(),
            };
            archive.insert_object(ArchiveObject::new(
                43,
                vec![RawMessage {
                    type_: 6004,
                    data: headers.encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    package
}

fn test_package_with_cross_table_engine() -> IWorkPackage {
    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let sheet = archive.object_mut(2).unwrap();
            let sheet_type = sheet.messages[0].type_;
            let mut sheet_archive = tn::SheetArchive::decode(sheet.messages[0].data.as_slice())?;
            sheet_archive.drawable_infos.push(Reference {
                identifier: 4,
                ..Default::default()
            });
            sheet.replace_message(
                0,
                RawMessage {
                    type_: sheet_type,
                    data: sheet_archive.encode_to_vec(),
                },
            )?;

            let mut target =
                TableModelArchive::decode(archive.object(10).unwrap().messages[0].data.as_slice())?;
            target.table_id = "BD57CADC-10F6-658A-E24E-294B3170BBD8".to_owned();
            target.table_name = "Target".to_owned();
            archive.insert_object(ArchiveObject::new(
                4,
                vec![RawMessage {
                    type_: 6003,
                    data: tst::TableInfoArchive {
                        table_model: Reference {
                            identifier: 11,
                            ..Default::default()
                        },
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            archive.insert_object(ArchiveObject::new(
                11,
                vec![RawMessage {
                    type_: 6000,
                    data: target.encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let root = archive.object_mut(100).unwrap();
            let root_type = root.messages[0].type_;
            let mut engine =
                tsce::CalculationEngineArchive::decode(root.messages[0].data.as_slice())?;
            engine
                .dependency_tracker
                .formula_owner_dependencies
                .push(Reference {
                    identifier: 102,
                    ..Default::default()
                });
            root.replace_message(
                0,
                RawMessage {
                    type_: root_type,
                    data: engine.encode_to_vec(),
                },
            )?;
            root.archive_info.message_infos[0]
                .object_references
                .push(102);
            archive.insert_object(ArchiveObject::new(
                102,
                vec![RawMessage {
                    type_: 4008,
                    data:
                        tsce::FormulaOwnerDependenciesArchive {
                            formula_owner_uid: crate::protobuf::tsp::Uuid {
                                upper: 0xbd57_cadc_10f6_658a,
                                lower: 0xe24e_294b_3170_bbd8,
                            },
                            internal_formula_owner_id: 7,
                            owner_kind: Some(1),
                            cell_dependencies: Some(
                                tsce::CellDependenciesExpandedArchive::default(),
                            ),
                            formula_owner: Some(Reference {
                                identifier: 4,
                                ..Default::default()
                            }),
                            tiled_cell_dependencies: Some(
                                tsce::CellDependenciesTiledArchive::default(),
                            ),
                            ..Default::default()
                        }
                        .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    package
}

fn two_sheet_package() -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let root = archive.object_mut(1).unwrap();
            let mut document = tn::DocumentArchive::decode(root.messages[0].data.as_slice())?;
            document.sheets.push(Reference {
                identifier: 50,
                ..Default::default()
            });
            root.replace_message(
                0,
                RawMessage {
                    type_: 1,
                    data: document.encode_to_vec(),
                },
            )?;
            archive.insert_object(ArchiveObject::new(
                50,
                vec![RawMessage {
                    type_: 2,
                    data: tn::SheetArchive {
                        name: "Second".to_owned(),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                }],
            )?)?;
            Ok(())
        })
        .unwrap();
    package
}
