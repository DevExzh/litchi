use super::*;
use crate::archive::{Archive, ArchiveObject};
use crate::numbers::NumbersDocument;
use crate::package_metadata::{PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE};
use crate::protobuf::tn;
use crate::protobuf::tsp::{ComponentInfo, ObjectUuidMapEntry, PackageMetadata, Reference, Uuid};
use crate::shapes::{DrawablePoint, DrawableSize};

#[test]
fn ordinary_text_box_crud_is_guarded_and_byte_exact() {
    let mut editor = NumbersEditor::from_package(test_package_with_text_box()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let text_boxes = editor.sheet_text_boxes(2).unwrap();
    assert_eq!(text_boxes.len(), 1);
    assert_eq!(text_boxes[0].drawable_object_id, 50);
    assert_eq!(text_boxes[0].storage.object_id, 53);
    assert_eq!(text_boxes[0].storage.text, "Source");

    editor
        .replace_sheet_text_box_text(2, 50, 0..6, "Edited 🚀")
        .unwrap();
    assert_eq!(
        editor.sheet_text_boxes(2).unwrap()[0].storage.text,
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
            .map(|drawable| drawable.object_id)
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
    assert_eq!(comment.drawable_object_id, 50);
    assert_eq!(comment.comment.text, "Sheet annotation");
    assert_eq!(
        editor.sheet_text_boxes(2).unwrap()[0].storage.text,
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
    assert_ne!(created.storage.object_id, 53);
    assert_eq!(created.storage.text, "Independent clone");
    assert_eq!(editor.sheet_text_boxes(2).unwrap().len(), 2);
    assert_eq!(
        editor.sheet_text_boxes(2).unwrap()[0].storage.text,
        "Source"
    );
    let clone_geometry = editor
        .sheet_text_box_geometry(2, created.drawable_object_id)
        .unwrap();
    assert_eq!(
        clone_geometry.position,
        source_geometry.position.map(|position| DrawablePoint {
            x: position.x + TEXT_BOX_DUPLICATE_OFFSET,
            y: position.y + TEXT_BOX_DUPLICATE_OFFSET,
        })
    );

    let removed = editor
        .remove_sheet_text_box(2, created.drawable_object_id)
        .unwrap();
    assert_eq!(removed.text_box.storage.text, "Independent clone");
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
    let created = editor.duplicate_sheet(SOURCE_SHEET_ID).unwrap();
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
    assert_eq!(copied_text_box.storage.text, "Source");

    editor
        .set_cell(
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
            .text,
        "Source"
    );
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets().unwrap();
    assert_eq!(
        sheets[0].tables[0].get_cell(SOURCE_CELL_ROW, SOURCE_CELL_COLUMN),
        Some(&CellValue::Text("Original".to_owned()))
    );
    assert_eq!(
        sheets[1].tables[0].get_cell(SOURCE_CELL_ROW, SOURCE_CELL_COLUMN),
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
    assert!(editor.duplicate_sheet(SOURCE_SHEET_ID).is_err());
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
    assert!(editor.duplicate_sheet(SOURCE_SHEET_ID).is_err());
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
            archive.insert_object(owner)
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
            entry.extend(crate::varint::encode_varint(16));
            entry.extend(crate::varint::encode_varint(1));
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
            row.extend(crate::varint::encode_varint(8));
            row.extend(crate::varint::encode_varint(0));
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
            header.extend(crate::varint::encode_varint(8));
            header.extend(crate::varint::encode_varint(0));
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
fn semantic_edits_round_trip_through_public_reader() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let table = editor.tables().unwrap().remove(0);

    editor
        .set_cell(
            table.object_id,
            0,
            1,
            CellValue::Text("Updated".to_string()),
        )
        .unwrap();
    editor
        .set_cell(table.object_id, 1, 2, CellValue::Number(12.5))
        .unwrap();
    editor
        .set_cell(table.object_id, 2, 3, CellValue::Boolean(true))
        .unwrap();
    editor
        .set_cell(table.object_id, 3, 0, CellValue::Date(123_456.25))
        .unwrap();
    editor
        .set_cell(table.object_id, 3, 1, CellValue::Duration(3_600.5))
        .unwrap();

    let bytes = editor.to_bytes().unwrap();
    let document = NumbersDocument::from_bytes(&bytes).unwrap();
    let sheets = document.sheets().unwrap();
    let table = &sheets[0].tables[0];
    assert_eq!(table.get_cell(0, 1).unwrap().as_text(), "Updated");
    assert_eq!(table.get_cell(1, 2).unwrap().as_number(), Some(12.5));
    assert_eq!(table.get_cell(2, 3).unwrap().as_boolean(), Some(true));
    assert_eq!(table.get_cell(3, 0), Some(&CellValue::Date(123_456.25)));
    assert_eq!(table.get_cell(3, 1), Some(&CellValue::Duration(3_600.5)));
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
    editor.set_cell(10, 0, 1, CellValue::Number(7.5)).unwrap();
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
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0]
            .get_cell(0, 1)
            .unwrap()
            .as_number(),
        Some(7.5)
    );
}

#[test]
fn rich_text_cell_updates_preserve_the_payload_reference() {
    let mut editor = NumbersEditor::from_package(test_package_with_rich_text(false)).unwrap();
    editor
        .set_cell(10, 0, 1, CellValue::Text("Updated 🔬".to_owned()))
        .unwrap();

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(0, 1),
        Some(&CellValue::Text("Updated 🔬".to_owned()))
    );
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let tile = Tile::decode(archive.object(30).unwrap().messages[0].data.as_slice()).unwrap();
    let cells = split_row(&tile.row_infos[0]).unwrap();
    assert_eq!(
        BncCell::parse(cells[1].as_deref().unwrap())
            .unwrap()
            .stored_value(),
        StoredValue::RichText(2)
    );
    let list =
        TableDataList::decode(archive.object(50).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(list.entries[0].key, 2);
    assert_eq!(list.entries[0].refcount, 1);
    assert_eq!(
        list.entries[0]
            .rich_text_payload
            .as_ref()
            .unwrap()
            .identifier,
        51
    );
    let storage =
        tswp::StorageArchive::decode(archive.object(52).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(storage.text.concat(), "Updated 🔬");
}

#[test]
fn shared_rich_text_cell_update_uses_copy_on_write() {
    let mut editor = NumbersEditor::from_package(test_package_with_rich_text(true)).unwrap();
    editor
        .set_cell(10, 0, 1, CellValue::Text("Independent".to_owned()))
        .unwrap();

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(
        table.get_cell(0, 1),
        Some(&CellValue::Text("Independent".to_owned()))
    );
    assert_eq!(
        table.get_cell(0, 2),
        Some(&CellValue::Text("Original Rich".to_owned()))
    );

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let tile = Tile::decode(archive.object(30).unwrap().messages[0].data.as_slice()).unwrap();
    let cells = split_row(&tile.row_infos[0]).unwrap();
    assert_eq!(
        BncCell::parse(cells[1].as_deref().unwrap())
            .unwrap()
            .stored_value(),
        StoredValue::RichText(3)
    );
    assert_eq!(
        BncCell::parse(cells[2].as_deref().unwrap())
            .unwrap()
            .stored_value(),
        StoredValue::RichText(2)
    );
    let list =
        TableDataList::decode(archive.object(50).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(
        list.entries
            .iter()
            .map(|entry| (
                entry.key,
                entry.refcount,
                entry.rich_text_payload.as_ref().unwrap().identifier
            ))
            .collect::<Vec<_>>(),
        [(2, 1, 51), (3, 1, 54)]
    );
    let payload = tst::RichTextPayloadArchive::decode(
        archive.object(54).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(payload.storage.identifier, 53);
    assert_eq!(payload.cellid.packed_data, 1 << 16);
    let storage =
        tswp::StorageArchive::decode(archive.object(53).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(storage.text.concat(), "Independent");
}

#[test]
fn replacing_rich_text_releases_list_and_payload_objects() {
    let mut editor = NumbersEditor::from_package(test_package_with_rich_text(false)).unwrap();
    editor.set_cell(10, 0, 1, CellValue::Number(42.25)).unwrap();

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0]
            .get_cell(0, 1)
            .unwrap()
            .as_number(),
        Some(42.25)
    );
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let list =
        TableDataList::decode(archive.object(50).unwrap().messages[0].data.as_slice()).unwrap();
    assert!(list.entries.is_empty());
    assert!(archive.object(51).is_none());
    assert!(archive.object(52).is_none());
}

#[test]
fn segmented_string_entries_round_trip_and_remain_interned() {
    let mut package = test_package();
    move_table_data_list_entries_to_segment(&mut package, 20, 60);
    let before = NumbersDocument::from_bytes(&package.to_bytes().unwrap()).unwrap();
    assert_eq!(
        before.sheets().unwrap()[0].tables[0].get_cell(0, 1),
        Some(&CellValue::Text("Original".to_owned()))
    );

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .set_cell(10, 0, 2, CellValue::Text("Original".to_owned()))
        .unwrap();
    editor
        .set_cell(10, 0, 1, CellValue::Text("Updated".to_owned()))
        .unwrap();

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let root =
        TableDataList::decode(archive.object(20).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(root.segments[0].identifier, 60);
    assert_eq!(root.entries.len(), 1);
    assert_eq!(root.entries[0].string.as_deref(), Some("Updated"));
    let segment =
        TableDataListSegment::decode(archive.object(60).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(segment.entries.len(), 1);
    assert_eq!(segment.entries[0].string.as_deref(), Some("Original"));
    assert_eq!(segment.entries[0].refcount, 1);
    assert_eq!(segment.key_range.location, 1);
    assert_eq!(segment.key_range.length, 1);

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(
        table.get_cell(0, 1),
        Some(&CellValue::Text("Updated".to_owned()))
    );
    assert_eq!(
        table.get_cell(0, 2),
        Some(&CellValue::Text("Original".to_owned()))
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

    let document = NumbersDocument::from_bytes(&package.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(0, 0),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor.set_formula(10, 1, 0, expression).unwrap();
    editor.set_cell(10, 0, 0, CellValue::Number(7.0)).unwrap();

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let segment =
        TableDataListSegment::decode(archive.object(61).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(segment.entries.len(), 1);
    assert_eq!(segment.entries[0].refcount, 1);
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(table.get_cell(0, 0).unwrap().as_number(), Some(7.0));
    assert_eq!(
        table.get_cell(1, 0),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
}

#[test]
fn segmented_shared_rich_text_uses_copy_on_write_and_cleans_up() {
    let mut package = test_package_with_rich_text(true);
    move_table_data_list_entries_to_segment(&mut package, 50, 60);
    let document = NumbersDocument::from_bytes(&package.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(0, 1),
        Some(&CellValue::Text("Original Rich".to_owned()))
    );

    let mut editor = NumbersEditor::from_package(package).unwrap();
    editor
        .set_cell(10, 0, 1, CellValue::Text("Independent".to_owned()))
        .unwrap();
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let root =
        TableDataList::decode(archive.object(50).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(root.entries.len(), 1);
    assert_eq!(root.entries[0].key, 3);
    let segment =
        TableDataListSegment::decode(archive.object(60).unwrap().messages[0].data.as_slice())
            .unwrap();
    assert_eq!(segment.entries[0].refcount, 1);

    editor.set_cell(10, 0, 2, CellValue::Number(9.0)).unwrap();
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let root =
        TableDataList::decode(archive.object(50).unwrap().messages[0].data.as_slice()).unwrap();
    assert!(root.segments.is_empty());
    assert!(archive.object(60).is_none());
    assert!(archive.object(51).is_none());
    assert!(archive.object(52).is_none());
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(
        table.get_cell(0, 1),
        Some(&CellValue::Text("Independent".to_owned()))
    );
    assert_eq!(table.get_cell(0, 2).unwrap().as_number(), Some(9.0));
}

#[test]
fn formula_error_cells_release_root_and_segmented_list_entries() {
    for segmented in [false, true] {
        let mut package = test_package_with_formula_error();
        if segmented {
            move_table_data_list_entries_to_segment(&mut package, 22, 60);
        }
        let before = NumbersDocument::from_bytes(&package.to_bytes().unwrap()).unwrap();
        assert_eq!(
            before.sheets().unwrap()[0].tables[0].get_cell(0, 1),
            Some(&CellValue::Error("Syntax Error".to_owned()))
        );

        let mut editor = NumbersEditor::from_package(package).unwrap();
        editor.set_cell(10, 0, 1, CellValue::Number(4.5)).unwrap();
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
        let after = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            after.sheets().unwrap()[0].tables[0]
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
    assert_eq!(original.comment.reply_object_ids, [70]);
    assert_eq!(original.comment.storage_uuid.unwrap().lower, 61);
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let reader_comment = document.sheets().unwrap()[0].tables[0]
        .get_comment(0, 1)
        .unwrap()
        .clone();
    assert_eq!(reader_comment, original.comment);

    editor
        .set_cell_comment(10, 0, 1, "Updated comment")
        .unwrap();
    let updated = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    assert_eq!(updated.storage_object_id, original.storage_object_id);
    assert_eq!(updated.comment.text, "Updated comment");
    assert_eq!(updated.comment.creation_date_seconds, Some(123.5));
    assert_eq!(updated.comment.storage_uuid, original.comment.storage_uuid);

    editor.set_cell(10, 0, 1, CellValue::Number(8.5)).unwrap();
    assert_eq!(
        editor.cell_comment(10, 0, 1).unwrap().unwrap().comment.text,
        "Updated comment"
    );
    editor.clear_cell(10, 0, 1).unwrap();
    assert!(editor.cell_comment(10, 0, 1).unwrap().is_some());

    editor.clear_cell_comment(10, 0, 1).unwrap();
    assert!(editor.cell_comment(10, 0, 1).unwrap().is_none());
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    assert!(archive.object(61).is_none());
    assert!(archive.object(70).is_none());
    let list =
        TableDataList::decode(archive.object(60).unwrap().messages[0].data.as_slice()).unwrap();
    assert!(list.entries.is_empty());

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(table.get_cell(0, 1), Some(&CellValue::Empty));
    assert!(table.get_comment(0, 1).is_none());
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
    assert_ne!(after_add.storage_object_id, original_root.storage_object_id);
    assert_eq!(
        after_add.comment.storage_uuid,
        original_root.comment.storage_uuid
    );
    assert_eq!(after_add.comment.reply_object_ids, [70, added_id]);
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
            .map(|reply| reply.storage_object_id)
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
    assert_eq!(replies[0].storage_object_id, 70);
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
        .storage_object_id;

    editor
        .set_cell_comment(10, 0, 1, "Independent comment")
        .unwrap();
    let first = editor.cell_comment(10, 0, 1).unwrap().unwrap();
    let second = editor.cell_comment(10, 0, 2).unwrap().unwrap();
    assert_ne!(first.storage_object_id, original_storage);
    assert_eq!(second.storage_object_id, original_storage);
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
    assert_ne!(cloned.storage_object_id, 61);
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    assert!(
        archive.object(cloned.storage_object_id).unwrap().messages[0]
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
    assert_eq!(info.list_identifier, 2);
    assert!(info.comment.creation_date_seconds.is_some());
    assert!(info.comment.storage_uuid.is_some());

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(table.get_cell(1, 2), Some(&CellValue::Empty));
    assert_eq!(table.get_comment(1, 2).unwrap().text, "Created comment");

    editor.set_cell(10, 1, 2, CellValue::Number(42.0)).unwrap();
    editor.clear_cell_comment(10, 1, 2).unwrap();
    assert!(editor.cell_comment(10, 1, 2).unwrap().is_none());
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0]
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
            .object(info.storage_object_id)
            .is_none()
    );
}

#[test]
fn empty_native_author_storage_rejects_unopenable_cell_comment_creation() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            archive.insert_object(ArchiveObject::new(
                90,
                vec![RawMessage {
                    type_: 213,
                    data: crate::protobuf::tsk::AnnotationAuthorStorageArchive::default()
                        .encode_to_vec(),
                }],
            )?)
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    let error = editor
        .set_cell_comment(10, 1, 2, "Would become a blank draft")
        .unwrap_err();
    assert!(error.to_string().contains("native annotation author"));
    assert_eq!(editor.to_bytes().unwrap(), before);
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
    let document = NumbersDocument::from_bytes(&before).unwrap();
    assert!(document.sheets().is_err());
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

    let document = NumbersDocument::from_bytes(&package.to_bytes().unwrap()).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(table.get_cell(0, 1).unwrap().as_text(), "Original");
    assert_eq!(table.get_cell(256, 0).unwrap().as_number(), Some(99.0));
}

#[test]
fn cell_edits_keep_sparse_row_headers_in_lockstep() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();

    editor.set_cell(10, 3, 0, CellValue::Number(1.0)).unwrap();
    editor.set_cell(10, 3, 2, CellValue::Boolean(true)).unwrap();
    let bucket = row_header_bucket(editor.package(), 42);
    assert_eq!(
        bucket
            .headers
            .iter()
            .map(|header| (header.index, header.number_of_cells))
            .collect::<Vec<_>>(),
        [(0, 1), (3, 2)]
    );

    editor.clear_cell(10, 3, 0).unwrap();
    let bucket = row_header_bucket(editor.package(), 42);
    assert_eq!(bucket.headers[1].number_of_cells, 1);

    editor.clear_cell(10, 3, 2).unwrap();
    let bucket = row_header_bucket(editor.package(), 42);
    assert_eq!(
        bucket
            .headers
            .iter()
            .map(|header| (header.index, header.number_of_cells))
            .collect::<Vec<_>>(),
        [(0, 1)]
    );
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let tile = Tile::decode(archive.object(30).unwrap().messages[0].data.as_slice()).unwrap();
    assert!(tile.row_infos.iter().all(|row| row.tile_row_index != 3));
}

fn row_header_bucket(package: &IWorkPackage, identifier: u64) -> tst::HeaderStorageBucket {
    let archive = package.archive("Index/Document.iwa").unwrap();
    tst::HeaderStorageBucket::decode(
        archive.object(identifier).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap()
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
    let document = NumbersDocument::from_bytes(&bytes).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!(
        table.get_cell(0, 0),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    assert_eq!(table.get_cell(1, 0), table.get_cell(0, 0));

    editor
        .set_cell(table_id, 0, 0, CellValue::Number(42.0))
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
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(2, 0),
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
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(0, 0),
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
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(3, 2),
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

    editor.clear_cell(table_id, 3, 2).unwrap();
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
    let mut editor = NumbersEditor::from_package(test_package_with_cross_table_engine()).unwrap();
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
    assert_eq!(edges.edge_without_owner_columns, [0]);
    assert_eq!(edges.edge_without_owner_rows, [0]);
    assert_eq!(edges.edge_with_owner_columns, [1]);
    assert_eq!(edges.edge_with_owner_rows, [0]);
    assert_eq!(edges.internal_owner_id_for_edge, [7]);

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets().unwrap();
    let host = sheets[0]
        .tables
        .iter()
        .find(|table| table.name == "Table 1")
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

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let host = document.sheets().unwrap()[0]
        .tables
        .iter()
        .find(|table| table.name == "Table 1")
        .unwrap()
        .clone();
    assert_eq!(
        host.get_cell(3, 2),
        Some(&CellValue::Formula("=SUM(Sheet 1::Target::1:2)".to_owned()))
    );
}

#[test]
fn failed_edit_is_transactional() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_cell(table_id, 0, 1, CellValue::Formula("1+1".to_string()))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn formula_cells_can_be_cleared_with_refcount_cleanup() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let formula_object = archive.object_mut(21).unwrap();
            let formula_type = formula_object.messages[0].type_;
            let mut formulas = TableDataList::decode(formula_object.messages[0].data.as_slice())?;
            formulas.next_list_id = 2;
            formulas.entries.push(tst::table_data_list::ListEntry {
                key: 1,
                refcount: 1,
                string: None,
                reference: None,
                formula: Some(crate::protobuf::tsce::FormulaArchive::default()),
                format: None,
                custom_format: None,
                rich_text_payload: None,
                comment_storage: None,
                import_warning_set: None,
                cell_spec: None,
            });
            formula_object.replace_message(
                0,
                RawMessage {
                    type_: formula_type,
                    data: formulas.encode_to_vec(),
                },
            )?;

            let tile_object = archive.object_mut(30).unwrap();
            let tile_type = tile_object.messages[0].type_;
            let mut tile = Tile::decode(tile_object.messages[0].data.as_slice())?;
            let mut cells = split_row(&tile.row_infos[0])?;
            let mut cell = BncCell::minimal();
            cell.set_formula_reference(1);
            cells[0] = Some(cell.encode());
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

    let mut editor = NumbersEditor::from_package(package).unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    editor.clear_cell(table_id, 0, 0).unwrap();
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let formula_object = archive.object(21).unwrap();
    let formulas = TableDataList::decode(formula_object.messages[0].data.as_slice()).unwrap();
    assert!(formulas.entries.is_empty());
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        document.sheets().unwrap()[0].tables[0]
            .get_cell(0, 0)
            .is_none()
    );
}

#[test]
fn renames_root_ordered_sheet_and_table() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    assert_eq!(editor.sheets().unwrap()[0].name, "Sheet 1");
    editor.rename_sheet(2, "Résumé 東京").unwrap();
    editor.rename_table(10, "Inventory 🚀").unwrap();
    assert_eq!(editor.sheets().unwrap()[0].name, "Résumé 東京");
    assert_eq!(editor.tables().unwrap()[0].name, "Inventory 🚀");
    let before = editor.to_bytes().unwrap();
    assert!(editor.rename_sheet(2, "").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn duplicates_populated_table_with_independent_storage() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let created = editor.duplicate_table(10).unwrap();

    assert_ne!(created.object_id, 10);
    assert_eq!(created.name, "Table 1 copy");
    assert_eq!((created.rows, created.columns), (4, 4));

    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let tables = &document.sheets().unwrap()[0].tables;
    assert_eq!(tables.len(), 2);
    assert_eq!(tables[0].get_cell(0, 1).unwrap().as_text(), "Original");
    assert_eq!(tables[1].get_cell(0, 1).unwrap().as_text(), "Original");

    editor
        .set_cell(
            created.object_id,
            0,
            1,
            CellValue::Text("Independent".to_owned()),
        )
        .unwrap();
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let tables = &document.sheets().unwrap()[0].tables;
    assert_eq!(tables[0].get_cell(0, 1).unwrap().as_text(), "Original");
    assert_eq!(tables[1].get_cell(0, 1).unwrap().as_text(), "Independent");

    assert_eq!(editor.duplicate_table(10).unwrap().name, "Table 1 copy 2");
}

#[test]
fn duplicates_formula_table_with_independent_dependency_owner() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    let expression = FormulaExpression::function(
        "SUM",
        [
            FormulaExpression::Number(1.0),
            FormulaExpression::Number(2.0),
        ],
    );
    editor.set_formula(10, 1, 1, expression).unwrap();

    let created = editor.duplicate_table(10).unwrap();
    let owner = find_table_owner(editor.package(), created.object_id).unwrap();
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(document.sheets().unwrap()[0].tables.len(), 2);
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    assert_eq!(
        document.sheets().unwrap()[0].tables[1].get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );

    let calculation = editor
        .package()
        .archive("Index/CalculationEngine.iwa")
        .unwrap();
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
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(1,2)".to_owned()))
    );
    assert_eq!(
        document.sheets().unwrap()[0].tables[1].get_cell(1, 1),
        Some(&CellValue::Formula("=SUM(9)".to_owned()))
    );
}

#[test]
fn formula_table_duplicate_rejects_unsupported_dependencies_transactionally() {
    let mut package = test_package_with_calculation_engine();
    package
        .update_archive("Index/CalculationEngine.iwa", |archive| {
            let object = archive.object_mut(101).unwrap();
            let message = object.messages[0].clone();
            let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            owner.uuid_references = Some(tsce::UuidReferencesArchive {
                table_refs: vec![tsce::uuid_references_archive::TableRef {
                    owner_uuid: owner.formula_owner_uid,
                    coord_set: None,
                }],
                table_uuid_refs: Vec::new(),
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

    assert!(editor.duplicate_table(10).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn rename_and_resize_preserve_unknown_wire_and_restore_exact_component() {
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

    editor.rename_sheet(2, "Temporary Sheet").unwrap();
    editor.rename_table(10, "Temporary Table").unwrap();
    editor.rename_table(10, "Table 1").unwrap();
    editor.rename_sheet(2, "Sheet 1").unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        baseline
    );

    editor.resize_table(10, 6, 6).unwrap();
    editor.resize_table(10, 4, 4).unwrap();
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
fn form_sheet_rename_preserves_unknown_outer_and_nested_fields() {
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
                let mut nested = nested.to_vec();
                append_unknown_varint(&mut nested, 98, 980);
                Ok(nested)
            })?;
            append_unknown_varint(&mut data, 99, 990);
            object
                .replace_message(0, RawMessage { type_: 3, data })
                .map(|_| ())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let baseline = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    editor.rename_sheet(2, "Form Temporary").unwrap();
    assert_eq!(editor.sheets().unwrap()[0].name, "Form Temporary");
    editor.rename_sheet(2, "Sheet 1").unwrap();
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
fn table_rename_rejects_duplicate_name_fields_transactionally() {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            let message = object.messages[0].clone();
            let data = crate::wire::append_repeated_length_delimited_field(
                &message.data,
                8,
                b"Duplicate",
            )?;
            object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )
                .map(|_| ())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.rename_table(10, "Rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn grows_and_truncates_blank_table_edges_with_uid_maps() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    editor.resize_table(10, 6, 6).unwrap();
    editor
        .set_cell(10, 5, 5, CellValue::Text("edge".to_owned()))
        .unwrap();
    assert!(editor.resize_table(10, 4, 4).is_err());
    editor.clear_cell(10, 5, 5).unwrap();
    editor.resize_table(10, 3, 3).unwrap();
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
fn inserts_blank_table_row_and_shifts_cells_uids_headers_and_formulas() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_cell(10, 1, 1, CellValue::Text("Apples".to_owned()))
        .unwrap();
    editor
        .set_formula(
            10,
            2,
            2,
            FormulaExpression::cell(crate::numbers::FormulaCellReference::relative(1, 1)),
        )
        .unwrap();

    editor.insert_table_row(10, 1).unwrap();

    let bytes = editor.to_bytes().unwrap();
    let document = NumbersDocument::from_bytes(&bytes).unwrap();
    let table = &document.sheets().unwrap()[0].tables[0];
    assert_eq!((table.row_count, table.column_count), (5, 4));
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
    editor.insert_table_row(10, 4).unwrap();
    assert_eq!(editor.tables().unwrap()[0].rows, 5);
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(document.sheets().unwrap()[0].tables[0].row_count, 5);
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
    editor.set_cell(10, 3, 0, CellValue::Number(9.0)).unwrap();
    let before = editor.to_bytes().unwrap();

    assert!(editor.insert_table_row(10, 3).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_rejects_out_of_bounds_index_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.insert_table_row(10, 5).is_err());
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

    assert!(editor.insert_table_row(10, 1).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn row_insert_preserves_unknown_tile_header_and_dependency_record_fields() {
    let mut editor = NumbersEditor::from_package(test_package_with_calculation_engine()).unwrap();
    editor
        .set_cell(10, 1, 1, CellValue::Text("opaque".to_owned()))
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
    editor.insert_table_row(10, 1).unwrap();

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
fn removes_table_from_owning_sheet_transactionally() {
    let mut editor = NumbersEditor::from_package(test_package()).unwrap();
    let removed = editor.remove_table(10).unwrap();
    assert_eq!(removed.name, "Table 1");
    assert!(editor.tables().unwrap().is_empty());
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets().unwrap();
    assert_eq!(sheets.len(), 1);
    assert!(sheets[0].tables.is_empty());
    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_table(10).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn reorders_and_removes_sheets_transactionally() {
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
    editor.move_sheet(1, 0).unwrap();
    assert_eq!(editor.sheets().unwrap()[0].object_id, 50);
    let removed = editor.remove_sheet(50).unwrap();
    assert_eq!(removed.name, "Second");
    assert_eq!(editor.sheets().unwrap()[0].object_id, 2);
    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_sheet(2).is_err());
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

    let moved = editor.move_table(10, 50).unwrap();
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

    editor
        .set_cell(10, 0, 0, CellValue::Text("Moved cell".to_owned()))
        .unwrap();
    assert_eq!(
        NumbersDocument::from_bytes(&editor.to_bytes().unwrap())
            .unwrap()
            .sheets()
            .unwrap()[1]
            .tables[0]
            .get_cell(0, 0),
        Some(&CellValue::Text("Moved cell".to_owned()))
    );

    let mut editor = NumbersEditor::from_bytes(&baseline).unwrap();
    editor.move_table(10, 50).unwrap();
    editor.move_table(10, 2).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(editor.move_table(999, 50).is_err());
    assert!(editor.move_table(10, 999).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn sheet_list_crud_preserves_raw_references_and_restores_exact_component() {
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
            object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )
                .map(|_| ())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let baseline = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();

    editor.move_sheet(0, 1).unwrap();
    editor.move_sheet(1, 0).unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        baseline
    );

    let created = editor.add_empty_sheet("Temporary").unwrap();
    editor.remove_sheet(created.object_id).unwrap();
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
fn duplicate_sheet_references_fail_transactionally() {
    let mut package = two_sheet_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(1).unwrap();
            let message = object.messages[0].clone();
            let first = crate::wire::repeated_length_delimited_payloads(&message.data, 1)?[0];
            let data =
                crate::wire::append_repeated_length_delimited_field(&message.data, 1, first)?;
            object
                .replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )
                .map(|_| ())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.move_sheet(0, 1).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
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
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(document.sheets().unwrap()[1].name, "Created 東京");
}

#[test]
fn detached_table_models_are_not_exposed_or_writable() {
    let mut editor = NumbersEditor::from_package(two_sheet_package()).unwrap();
    editor.remove_sheet(2).unwrap();
    assert!(editor.tables().unwrap().is_empty());

    let before = editor.to_bytes().unwrap();
    assert!(editor.rename_table(10, "Detached").is_err());
    assert!(editor.set_cell(10, 0, 0, CellValue::Number(1.0)).is_err());
    assert!(editor.resize_table(10, 5, 5).is_err());
    assert!(editor.remove_table(10).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn creates_independent_empty_table_on_an_existing_sheet() {
    let mut package = two_sheet_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(10).unwrap();
            object.archive_info.message_infos[0].object_references = vec![20, 999];
            object.archive_info.message_infos[0].field_infos.push(
                crate::protobuf::tsp::FieldInfo {
                    path: crate::protobuf::tsp::FieldPath { path: vec![4, 4] },
                    object_references: vec![20, 999],
                    ..Default::default()
                },
            );
            Ok(())
        })
        .unwrap();
    let mut editor = NumbersEditor::from_package(package.clone()).unwrap();
    let created = editor.add_empty_table(50, "Created Table", 3, 2).unwrap();
    let mut repeated = NumbersEditor::from_package(package).unwrap();
    repeated.add_empty_table(50, "Created Table", 3, 2).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), repeated.to_bytes().unwrap());
    assert_ne!(created.object_id, 10);
    assert_eq!((created.rows, created.columns), (3, 2));
    assert_eq!(created.name, "Created Table");
    assert_eq!(editor.tables().unwrap().len(), 2);
    let component_name = editor
        .package()
        .entry_names()
        .find(|name| name.starts_with("Index/Tables/Table-"))
        .unwrap();
    let component = editor.package().archive(component_name).unwrap();
    let model_object = component.object(created.object_id).unwrap();
    let cloned_model = TableModelArchive::decode(model_object.messages[0].data.as_slice()).unwrap();
    assert_eq!(
        model_object.archive_info.message_infos[0].field_infos[0].object_references,
        [cloned_model.base_data_store.string_table.identifier]
    );

    editor
        .set_cell(
            created.object_id,
            0,
            0,
            CellValue::Text("Independent".to_owned()),
        )
        .unwrap();
    let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let sheets = document.sheets().unwrap();
    assert_eq!(
        sheets[0].tables[0].get_cell(0, 1).unwrap().as_text(),
        "Original"
    );
    assert_eq!(
        sheets[1].tables[0].get_cell(0, 0).unwrap().as_text(),
        "Independent"
    );

    let component = editor
        .package()
        .entry_names()
        .find(|name| name.starts_with("Index/Tables/Table-"))
        .unwrap()
        .to_owned();
    editor.remove_table(created.object_id).unwrap();
    assert_eq!(editor.tables().unwrap().len(), 1);
    assert!(!editor.package().contains_entry(&component));
    let before = editor.to_bytes().unwrap();
    assert!(editor.add_empty_table(999, "Missing", 2, 2).is_err());
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
            object
                .replace_message(0, RawMessage { type_: 3, data })
                .map(|_| ())
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

    let created = editor.add_empty_table(2, "Temporary", 2, 2).unwrap();
    editor.remove_table(created.object_id).unwrap();
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
            archive.insert_object(segment_object)
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

fn test_package_with_rich_text(shared: bool) -> IWorkPackage {
    let mut package = test_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let model_object = archive.object_mut(10).unwrap();
            let model_type = model_object.messages[0].type_;
            let mut model = TableModelArchive::decode(model_object.messages[0].data.as_slice())?;
            model.base_data_store.rich_text_table = Some(Reference {
                identifier: 50,
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
                .push(50);

            let tile_object = archive.object_mut(30).unwrap();
            let tile_type = tile_object.messages[0].type_;
            let mut tile = Tile::decode(tile_object.messages[0].data.as_slice())?;
            let mut cells = split_row(&tile.row_infos[0])?;
            let mut rich = BncCell::minimal();
            rich.set_rich_text(2);
            cells[1] = Some(rich.encode());
            if shared {
                cells[2] = cells[1].clone();
            }
            rebuild_row(&mut tile.row_infos[0], &cells)?;
            tile.max_column = if shared { 2 } else { 1 };
            tile.max_row = 0;
            tile.num_cells = if shared { 2 } else { 1 };
            tile_object.replace_message(
                0,
                RawMessage {
                    type_: tile_type,
                    data: tile.encode_to_vec(),
                },
            )?;

            let header_object = archive.object_mut(42).unwrap();
            let header_type = header_object.messages[0].type_;
            let mut headers =
                tst::HeaderStorageBucket::decode(header_object.messages[0].data.as_slice())?;
            headers.headers[0].number_of_cells = if shared { 2 } else { 1 };
            header_object.replace_message(
                0,
                RawMessage {
                    type_: header_type,
                    data: headers.encode_to_vec(),
                },
            )?;

            let rich_text_list = TableDataList {
                list_type: tst::table_data_list::ListType::RichTextPayload as i32,
                next_list_id: 3,
                entries: vec![tst::table_data_list::ListEntry {
                    key: 2,
                    refcount: if shared { 2 } else { 1 },
                    rich_text_payload: Some(Reference {
                        identifier: 51,
                        ..Default::default()
                    }),
                    ..Default::default()
                }],
                segments: Vec::new(),
                is_new_for_bnc: Some(true),
            };
            let mut list_object = ArchiveObject::new(
                50,
                vec![RawMessage {
                    type_: 6005,
                    data: rich_text_list.encode_to_vec(),
                }],
            )?;
            list_object.archive_info.message_infos[0]
                .object_references
                .push(51);
            archive.insert_object(list_object)?;

            let payload = tst::RichTextPayloadArchive {
                storage: Reference {
                    identifier: 52,
                    ..Default::default()
                },
                range: None,
                cellid: tst::CellId {
                    packed_data: 1 << 16,
                    expanded_coord: None,
                },
            };
            let mut payload_object = ArchiveObject::new(
                51,
                vec![RawMessage {
                    type_: 6218,
                    data: payload.encode_to_vec(),
                }],
            )?;
            payload_object.archive_info.message_infos[0]
                .object_references
                .push(52);
            archive.insert_object(payload_object)?;

            archive.insert_object(ArchiveObject::new(
                52,
                vec![RawMessage {
                    type_: 2001,
                    data: tswp::StorageArchive {
                        kind: Some(tswp::storage_archive::KindType::Cell as i32),
                        text: vec!["Original Rich".to_owned()],
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
    let mut unknown = crate::varint::encode_varint(99 << 3);
    unknown.extend(crate::varint::encode_varint(999));
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
    data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
    data.extend(crate::varint::encode_varint(value));
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
