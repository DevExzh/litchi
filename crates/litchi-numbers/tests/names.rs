use std::io;
use std::sync::Arc;

use litchi_iwa_archive::{Limits, package::Catalog, package::EntryEdit};
use litchi_iwa_common::{
    decode_varint_from_bytes, encode_varint_into,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsd, tsp, tst};
use litchi_numbers::{Package, SheetSelector, TableSelector, cell::Value, names};
use prost::Message as _;

const DOCUMENT: &str = "Index/Document.iwa";
const VIEW_STATE: &str = "Index/ViewState.iwa";
const FIRST_SHEET: u64 = 2;
const SECOND_SHEET: u64 = 3;
const FIRST_INFO: u64 = 10;
const SECOND_INFO: u64 = 11;
const THIRD_INFO: u64 = 12;
const FIRST_MODEL: u64 = 20;
const SECOND_MODEL: u64 = 21;
const THIRD_MODEL: u64 = 22;
const SIDECARS: u64 = 90;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn table_model(name: &str, identifier: u64, pivot: bool) -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: format!("table-{identifier}"),
        table_name: name.to_owned(),
        number_of_rows: 1,
        number_of_columns: 1,
        base_data_store: tst::DataStore {
            string_table: reference(SIDECARS),
            formula_table: reference(SIDECARS),
            ..Default::default()
        },
        pivot_owner: pivot.then(|| reference(777)),
        ..Default::default()
    }
}

fn table_info_payload(model: u64, locked: bool) -> TestResult<Vec<u8>> {
    let drawable = tsd::DrawableArchive {
        locked: locked.then_some(true),
        ..Default::default()
    }
    .encode_to_vec();
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable)?;
    append_length_delimited_field(&mut payload, 2, &reference(model).encode_to_vec())?;
    Ok(payload)
}

fn table_info(identifier: u64, model: u64, locked: bool) -> TestResult<ArchiveObject> {
    let mut value = object(identifier, 6_000, table_info_payload(model, locked)?)?;
    value.archive_info.message_infos[0].object_references = vec![model];
    Ok(value)
}

fn sidecars() -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        SIDECARS,
        [
            tst::table_data_list::ListType::String,
            tst::table_data_list::ListType::Formula,
        ]
        .into_iter()
        .map(|list_type| RawMessage {
            type_: 6_005,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        })
        .collect(),
    )?)
}

fn compressed(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn package_bytes(locked: bool, pivot: bool, previews: usize) -> TestResult<Vec<u8>> {
    let mut root = object(
        1,
        1,
        tn::DocumentArchive {
            sheets: vec![reference(FIRST_SHEET), reference(SECOND_SHEET)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    root.archive_info.message_infos[0].object_references = vec![FIRST_SHEET, SECOND_SHEET];

    let mut first_sheet = object(
        FIRST_SHEET,
        2,
        tn::SheetArchive {
            name: "Alpha".to_owned(),
            drawable_infos: vec![reference(FIRST_INFO), reference(SECOND_INFO)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    first_sheet.archive_info.message_infos[0].object_references = vec![FIRST_INFO, SECOND_INFO];
    let mut second_sheet = object(
        SECOND_SHEET,
        2,
        tn::SheetArchive {
            name: "Beta".to_owned(),
            drawable_infos: vec![reference(THIRD_INFO)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    second_sheet.archive_info.message_infos[0].object_references = vec![THIRD_INFO];

    let document = compressed(vec![
        root,
        first_sheet,
        second_sheet,
        table_info(FIRST_INFO, FIRST_MODEL, locked)?,
        table_info(SECOND_INFO, SECOND_MODEL, false)?,
        table_info(THIRD_INFO, THIRD_MODEL, false)?,
        object(
            FIRST_MODEL,
            6_001,
            table_model("One", FIRST_MODEL, pivot).encode_to_vec(),
        )?,
        object(
            SECOND_MODEL,
            6_001,
            table_model("Two", SECOND_MODEL, false).encode_to_vec(),
        )?,
        object(
            THIRD_MODEL,
            6_001,
            table_model("One", THIRD_MODEL, false).encode_to_vec(),
        )?,
        sidecars()?,
    ])?;
    let view = compressed(vec![object(
        900,
        777,
        b"view state is byte exact".to_vec(),
    )?])?;
    let mut entries: Vec<(&str, &[u8])> = vec![
        ("Data/sentinel.bin", b"unrelated zip member"),
        (DOCUMENT, &document),
        (VIEW_STATE, &view),
    ];
    let preview_values = [
        ("preview.jpg", b"preview one".as_slice()),
        ("preview-micro.jpg", b"preview two".as_slice()),
        ("preview-web.jpg", b"preview three".as_slice()),
    ];
    entries.extend(preview_values.into_iter().take(previews));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut result = Vec::new();
    package.write_to(&mut result)?;
    Ok(result)
}

fn sheet_name(package: &Package, position: usize) -> &str {
    package.document().sheets().get(position).unwrap().name()
}

fn table_name(package: &Package, sheet: usize, table: usize) -> &str {
    package
        .document()
        .sheets()
        .get(sheet)
        .unwrap()
        .tables()
        .nth(table)
        .unwrap()
        .name()
}

fn rewrite_document(
    source: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT)
        .unwrap();
    let mut archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    mutate(&mut archive)?;
    let payload = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(DOCUMENT, &payload)], Limits::default())?)
}

fn form_first_sheet(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(source, |archive| {
        let sheet = archive.object_mut(FIRST_SHEET).unwrap();
        let original = sheet.messages[0].data.clone();
        sheet.replace_message_preserving_header(
            0,
            RawMessage {
                type_: 3,
                data: tn::FormBasedSheetArchive {
                    super_: tn::SheetArchive::decode(original.as_slice())?,
                    ..Default::default()
                }
                .encode_to_vec(),
            },
        )?;
        let info = &mut sheet.archive_info.message_infos[0];
        info.field_infos.clear();
        let mut drawables = FieldInfo::new(vec![1, 2]);
        drawables.object_references = vec![FIRST_INFO, SECOND_INFO];
        info.field_infos.push(drawables);
        Ok(())
    })
}

fn legacy_model(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(source, |archive| {
        let model = archive.object_mut(FIRST_MODEL).unwrap();
        model.replace_message_preserving_header(
            0,
            RawMessage {
                type_: 6_000,
                data: model.messages[0].data.clone(),
            },
        )?;
        Ok(())
    })
}

fn split_components(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let document = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT)
        .ok_or_else(|| io::Error::other("synthetic document is missing"))?;
    let archive = Archive::parse(&SnappyStream::decompress(document.data())?.into_bytes())?;
    let (document_objects, table_objects): (Vec<_>, Vec<_>) =
        archive.objects.into_iter().partition(|object| {
            object
                .archive_info
                .identifier
                .is_some_and(|identifier| identifier <= SECOND_SHEET)
        });
    let document = compressed(document_objects)?;
    let tables = compressed(table_objects)?;
    let mut entries = catalog
        .iter()
        .filter(|entry| entry.name() != DOCUMENT)
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.insert(1, (DOCUMENT, document.as_slice()));
    entries.insert(2, ("Index/Tables.iwa", tables.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn duplicate_selected_name_field(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(source, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL)
            .ok_or_else(|| io::Error::other("selected model is missing"))?;
        append_length_delimited_field(&mut model.messages[0].data, 8, b"duplicate")?;
        Ok(())
    })
}

fn selected_header_merge(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(source, |archive| {
        archive
            .object_mut(FIRST_MODEL)
            .ok_or_else(|| io::Error::other("selected model is missing"))?
            .archive_info
            .should_merge = Some(true);
        Ok(())
    })
}

fn empty_first_sheet(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(source, |archive| {
        let sheet = archive
            .object_mut(FIRST_SHEET)
            .ok_or_else(|| io::Error::other("selected sheet is missing"))?;
        let mut decoded = tn::SheetArchive::decode(sheet.messages[0].data.as_slice())?;
        decoded.drawable_infos.clear();
        sheet.replace_message_preserving_header(
            0,
            RawMessage {
                type_: 2,
                data: decoded.encode_to_vec(),
            },
        )?;
        sheet.archive_info.message_infos[0]
            .object_references
            .clear();
        Ok(())
    })
}

fn mismatched_selected_message_info(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT)
        .ok_or_else(|| io::Error::other("synthetic document is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("selected object is missing"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&stream[header_offset..])?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("header offset overflow"))?;
    let header_end = header_start
        .checked_add(usize::try_from(header_length)?)
        .ok_or_else(|| io::Error::other("header length overflow"))?;
    assert_eq!(header_end, data_offset);
    let mut header = tsp::ArchiveInfo::decode(&stream[header_start..header_end])?;
    header.message_infos[0].r#type = 99_999;
    let header = header.encode_to_vec();
    let mut rewritten = Vec::with_capacity(stream.len() + header.len());
    rewritten.extend_from_slice(&stream[..header_offset]);
    encode_varint_into(&mut rewritten, u64::try_from(header.len())?);
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(&stream[data_offset..]);
    let compressed = SnappyStream::compress(&rewritten)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(DOCUMENT, &compressed)], Limits::default())?)
}

fn cross_sheet_shared_model(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(source, |archive| {
        let info = archive
            .object_mut(THIRD_INFO)
            .ok_or_else(|| io::Error::other("second-sheet table info is missing"))?;
        info.replace_message_preserving_header(
            0,
            RawMessage {
                type_: 6_000,
                data: table_info_payload(FIRST_MODEL, false)?,
            },
        )?;
        info.archive_info.message_infos[0].object_references = vec![FIRST_MODEL];
        Ok(())
    })
}

fn assert_send_sync<T: Send + Sync>() {}

fn native_fixture() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers")
}

#[test]
fn sheet_table_and_combined_batches_are_atomic_and_preserve_view_state() -> TestResult {
    let source = package_bytes(false, false, 3)?;
    let package = Package::from_bytes(&source)?;
    let sheet = package
        .edit_names()
        .rename_sheet("Alpha", "Gamma")?
        .commit()?;
    assert_eq!(sheet_name(sheet.package(), 0), "Gamma");
    assert_eq!(table_name(sheet.package(), 0, 0), "One");
    assert_eq!(sheet.diagnostics().operations(), 1);
    assert!(sheet.diagnostics().changed());
    assert_eq!(sheet.diagnostics().deleted_previews(), 3);

    let table = package
        .edit_names()
        .rename_table(SheetSelector::index(0), TableSelector::index(1), "Deux")?
        .commit()?;
    assert_eq!(sheet_name(table.package(), 0), "Alpha");
    assert_eq!(table_name(table.package(), 0, 1), "Deux");
    assert!(table.diagnostics().changed());

    let combined = package
        .edit_names()
        .rename_sheet("Alpha", "Líneas 你好")?
        .rename_table("Alpha", "One", "表 Café №42")?
        .rename_table("Beta", "One", "Other sheet can reuse this name")?
        .commit()?;
    assert_eq!(sheet_name(combined.package(), 0), "Líneas 你好");
    assert_eq!(table_name(combined.package(), 0, 0), "表 Café №42");
    assert_eq!(
        table_name(combined.package(), 1, 0),
        "Other sheet can reuse this name"
    );
    assert_eq!(combined.diagnostics().operations(), 3);
    assert!(combined.diagnostics().changed());
    assert_eq!(
        Catalog::from_bytes(&source)?
            .iter()
            .find(|entry| entry.name() == VIEW_STATE)
            .unwrap()
            .data(),
        Catalog::from_bytes(&bytes(combined.package())?)?
            .iter()
            .find(|entry| entry.name() == VIEW_STATE)
            .unwrap()
            .data()
    );
    assert!(
        Catalog::from_bytes(&bytes(combined.package())?)?
            .iter()
            .all(|entry| !entry.name().starts_with("preview"))
    );
    Ok(())
}

#[test]
fn selectors_unicode_invalid_names_and_final_collisions_are_typed() -> TestResult {
    let package = Package::from_bytes(&package_bytes(false, false, 0)?)?;
    let composed = "Café";
    let decomposed = "Cafe\u{301}";
    let commit = package
        .edit_names()
        .rename_sheet(SheetSelector::index(0), composed)?
        .rename_sheet(SheetSelector::index(1), decomposed)?
        .rename_table(SheetSelector::index(0), TableSelector::index(0), "東京😀")?
        .commit()?;
    assert_eq!(sheet_name(commit.package(), 0), composed);
    assert_eq!(sheet_name(commit.package(), 1), decomposed);
    assert!(matches!(
        package.edit_names().rename_sheet("missing", "x"),
        Err(names::Error::SheetNotFound)
    ));
    assert!(matches!(
        package.edit_names().rename_table("Alpha", "missing", "x"),
        Err(names::Error::TableNotFound)
    ));
    assert!(matches!(
        package.edit_names().rename_sheet("Alpha", ""),
        Err(names::Error::InvalidName {
            reason: names::InvalidReason::Empty
        })
    ));
    assert!(matches!(
        package
            .edit_names()
            .rename_table("Alpha", "One", "bad\0name"),
        Err(names::Error::InvalidName {
            reason: names::InvalidReason::ContainsNul
        })
    ));
    assert!(matches!(
        package
            .edit_names()
            .rename_sheet("Alpha", "same")?
            .rename_sheet("Beta", "same"),
        Err(names::Error::DuplicateName {
            path: names::Path::Sheet { position: 1 }
        })
    ));
    assert!(matches!(
        package
            .edit_names()
            .rename_table("Alpha", "One", "same")?
            .rename_table("Alpha", "Two", "same"),
        Err(names::Error::DuplicateName {
            path: names::Path::Table { sheet: 0, table: 1 }
        })
    ));
    assert!(matches!(
        package
            .edit_names()
            .rename_sheet("Alpha", "first")?
            .rename_sheet("Alpha", "second"),
        Err(names::Error::DuplicateTarget {
            path: names::Path::Sheet { position: 0 }
        })
    ));

    let swapped = package
        .edit_names()
        .rename_sheet("Alpha", "Beta")?
        .rename_sheet("Beta", "Alpha")?
        .commit()?;
    assert_eq!(sheet_name(swapped.package(), 0), "Beta");
    assert_eq!(sheet_name(swapped.package(), 1), "Alpha");

    let collision_away = package
        .edit_names()
        .rename_table("Alpha", "One", "Two")?
        .rename_table("Alpha", "Two", "Three")?
        .commit()?;
    assert_eq!(table_name(collision_away.package(), 0, 0), "Two");
    assert_eq!(table_name(collision_away.package(), 0, 1), "Three");
    Ok(())
}

#[test]
fn noops_patches_inverse_conflicts_concurrency_and_redaction() -> TestResult {
    let source = package_bytes(false, false, 2)?;
    let package = Arc::new(Package::from_bytes(&source)?);
    let empty = package.edit_names().commit()?;
    assert!(!empty.diagnostics().changed());
    assert_eq!(empty.diagnostics().operations(), 0);
    assert_eq!(empty.diagnostics().deleted_previews(), 0);
    assert_eq!(bytes(empty.package())?, source);
    let noop = package
        .edit_names()
        .rename_sheet("Alpha", "Alpha")?
        .rename_table("Alpha", "One", "One")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.diagnostics().operations(), 0);
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(bytes(noop.package())?, source);
    assert_eq!(bytes(package.as_ref())?, bytes(noop.package())?);
    assert_eq!(bytes(package.apply_names(noop.patch())?.package())?, source);

    let first = package
        .edit_names()
        .rename_sheet("Alpha", "Changed")?
        .rename_table("Alpha", "One", "Renamed")?
        .commit()?;
    let applied = package.apply_names(first.patch())?;
    assert_eq!(bytes(applied.package())?, bytes(first.package())?);
    let restored = first.package().apply_names(&first.patch().inverse())?;
    assert!(restored.diagnostics().changed());
    assert_eq!(restored.diagnostics().deleted_previews(), 0);
    assert_eq!(bytes(restored.package())?, source);
    assert!(matches!(
        first.package().apply_names(first.patch()),
        Err(names::Error::PatchConflict)
    ));
    assert!(matches!(
        package.apply_names(&first.patch().inverse()),
        Err(names::Error::PatchConflict)
    ));
    let text = format!("{:?} {:?}", first.patch(), names::Error::PatchConflict);
    assert!(!text.contains("Changed"));
    assert!(!text.contains("Renamed"));

    let left = Arc::clone(&package);
    let right = Arc::clone(&package);
    let one = std::thread::spawn(move || left.edit_names().rename_sheet("Alpha", "Left")?.commit());
    let two =
        std::thread::spawn(move || right.edit_names().rename_sheet("Alpha", "Right")?.commit());
    assert_eq!(sheet_name(one.join().unwrap()?.package(), 0), "Left");
    assert_eq!(sheet_name(two.join().unwrap()?.package(), 0), "Right");
    assert_eq!(sheet_name(package.as_ref(), 0), "Alpha");
    assert_send_sync::<Package>();
    assert_send_sync::<names::Edit<'static>>();
    assert_send_sync::<names::Patch>();
    assert_send_sync::<names::Commit>();
    assert_send_sync::<names::Diagnostics>();
    assert_send_sync::<names::Error>();
    Ok(())
}

#[test]
fn form_legacy_lock_and_pivot_dependencies_fail_closed_without_publishing() -> TestResult {
    let form = Package::from_bytes(&form_first_sheet(&package_bytes(false, false, 0)?)?)?;
    let form_commit = form.edit_names().rename_sheet("Alpha", "Form")?.commit()?;
    assert_eq!(sheet_name(form_commit.package(), 0), "Form");

    let legacy = Package::from_bytes(&legacy_model(&package_bytes(false, false, 0)?)?)?;
    let legacy_commit = legacy
        .edit_names()
        .rename_table("Alpha", "One", "Legacy")?
        .commit()?;
    assert_eq!(table_name(legacy_commit.package(), 0, 0), "Legacy");

    let locked = Package::from_bytes(&package_bytes(true, false, 0)?)?;
    assert!(matches!(
        locked
            .edit_names()
            .rename_table("Alpha", "One", "Denied")?
            .commit(),
        Err(names::Error::TableLocked {
            path: names::Path::Table { sheet: 0, table: 0 }
        })
    ));
    assert_eq!(table_name(&locked, 0, 0), "One");

    let pivot = Package::from_bytes(&package_bytes(false, true, 0)?)?;
    assert!(matches!(
        pivot
            .edit_names()
            .rename_table("Alpha", "One", "Denied")?
            .commit(),
        Err(names::Error::UnsupportedDependency {
            path: names::Path::Table { sheet: 0, table: 0 }
        })
    ));
    Ok(())
}

#[test]
fn selected_name_wire_is_strict_and_unknown_bytes_stay_opaque() -> TestResult {
    let source = package_bytes(false, false, 0)?;
    let padded = rewrite_document(&source, |archive| {
        let sheet = archive.object_mut(FIRST_SHEET).unwrap();
        append_varint_field(&mut sheet.messages[0].data, 91, 1234)?;
        let model = archive.object_mut(FIRST_MODEL).unwrap();
        append_length_delimited_field(&mut model.messages[0].data, 92, b"opaque")?;
        Ok(())
    })?;
    let package = Package::from_bytes(&padded)?;
    let commit = package
        .edit_names()
        .rename_table("Alpha", "One", "Known")?
        .commit()?;
    let before = Catalog::from_bytes(&padded)?;
    let after = Catalog::from_bytes(&bytes(commit.package())?)?;
    let before_document = before
        .iter()
        .find(|entry| entry.name() == DOCUMENT)
        .unwrap()
        .data();
    let after_document = after
        .iter()
        .find(|entry| entry.name() == DOCUMENT)
        .unwrap()
        .data();
    assert_ne!(before_document, after_document);
    let parsed = Archive::parse(&SnappyStream::decompress(after_document)?.into_bytes())?;
    let model = parsed.object(FIRST_MODEL).unwrap();
    assert!(
        WireView::parse(&model.messages[0].data)?
            .fields()
            .any(|field| field.number() == 92)
    );
    Ok(())
}

#[test]
fn split_components_and_every_preview_cardinality_publish_once() -> TestResult {
    for previews in 0..=3 {
        let source = split_components(&package_bytes(false, false, previews)?)?;
        let package = Package::from_bytes(&source)?;
        let commit = package
            .edit_names()
            .rename_sheet("Alpha", "Split sheet")?
            .rename_table("Alpha", "One", "Split table")?
            .commit()?;
        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(commit.diagnostics().touched_components(), 2);
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().deleted_previews(), previews);
        assert_eq!(sheet_name(commit.package(), 0), "Split sheet");
        assert_eq!(table_name(commit.package(), 0, 0), "Split table");
        let output = bytes(commit.package())?;
        let catalog = Catalog::from_bytes(&output)?;
        assert_eq!(
            catalog
                .iter()
                .filter(|entry| entry.name().starts_with("preview"))
                .count(),
            0
        );
        assert_eq!(
            catalog
                .iter()
                .find(|entry| entry.name() == VIEW_STATE)
                .unwrap()
                .data(),
            Catalog::from_bytes(&source)?
                .iter()
                .find(|entry| entry.name() == VIEW_STATE)
                .unwrap()
                .data()
        );
    }
    Ok(())
}

#[test]
fn selected_duplicate_name_field_and_header_merge_refuse_atomically() -> TestResult {
    let duplicate = Package::from_bytes(&duplicate_selected_name_field(&package_bytes(
        false, false, 0,
    )?)?)?;
    assert!(matches!(
        duplicate
            .edit_names()
            .rename_table(SheetSelector::index(0), TableSelector::index(0), "Changed")
            .and_then(names::Edit::commit),
        Err(names::Error::InvalidSource)
    ));
    assert_eq!(table_name(&duplicate, 0, 0), "duplicate");

    let merged = Package::from_bytes(&selected_header_merge(&package_bytes(false, false, 0)?)?)?;
    assert!(matches!(
        merged
            .edit_names()
            .rename_table("Alpha", "One", "Changed")?
            .commit(),
        Err(names::Error::InvalidSource)
    ));
    assert_eq!(table_name(&merged, 0, 0), "One");
    Ok(())
}

#[test]
fn native_fixture_combined_rename_reopens_and_inverse_restores_exact_bytes() -> TestResult {
    let source = std::fs::read(native_fixture())?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_names()
        .rename_sheet("Sheet 1", "Líneas 你好 🧪")?
        .rename_table("Sheet 1", "Table 1", "表 Café №42")?
        .commit()?;
    assert_eq!(sheet_name(commit.package(), 0), "Líneas 你好 🧪");
    assert_eq!(table_name(commit.package(), 0, 0), "表 Café №42");
    assert!(matches!(
        commit.package().document().sheets()[0].tables().next().unwrap().get_a1("B2")?,
        Some(Value::Text(value)) if value == "Litchi native Numbers fixture"
    ));
    assert_eq!(
        bytes(
            commit
                .package()
                .apply_names(&commit.patch().inverse())?
                .package()
        )?,
        source
    );
    Ok(())
}

#[test]
fn empty_sheet_name_changes_and_empty_batches_have_exact_noop_diagnostics() -> TestResult {
    let source = empty_first_sheet(&package_bytes(false, false, 1)?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.document().sheets()[0].table_count(), 0);
    let changed = package
        .edit_names()
        .rename_sheet("Alpha", "Empty")?
        .commit()?;
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().operations(), 1);
    assert_eq!(changed.diagnostics().deleted_previews(), 1);
    assert_eq!(sheet_name(changed.package(), 0), "Empty");
    Ok(())
}

#[test]
fn selected_message_metadata_and_cross_sheet_model_aliases_fail_closed() -> TestResult {
    assert!(
        Package::from_bytes(&mismatched_selected_message_info(
            &package_bytes(false, false, 0)?,
            FIRST_SHEET,
        )?)
        .is_err()
    );

    assert!(
        Package::from_bytes(&mismatched_selected_message_info(
            &package_bytes(false, false, 0)?,
            FIRST_MODEL,
        )?)
        .is_err()
    );

    assert!(
        Package::from_bytes(&cross_sheet_shared_model(&package_bytes(true, false, 0,)?)?).is_err()
    );
    Ok(())
}
