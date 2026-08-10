use std::io;
use std::path::PathBuf;
use std::sync::Arc;

use litchi_iwa_archive::{Limits, package::Catalog, package::EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsd, tsp, tst};
use litchi_numbers::table::lock::State as LockState;
use litchi_numbers::{
    Package, SheetSelector, TableLockCommit, TableLockDiagnostics, TableLockEdit, TableLockError,
    TableLockLimitKind, TableLockPatch, TableSelector,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const TABLES_MEMBER: &str = "Index/Tables.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const FIRST_SHEET: u64 = 2;
const SECOND_SHEET: u64 = 3;
const ABSENT_INFO: u64 = 10;
const FALSE_INFO: u64 = 11;
const TRUE_INFO: u64 = 12;
const OTHER_INFO: u64 = 13;
const NON_TABLE_DRAWABLE: u64 = 14;
const ABSENT_MODEL: u64 = 20;
const FALSE_MODEL: u64 = 21;
const TRUE_MODEL: u64 = 22;
const OTHER_MODEL: u64 = 23;
const SIDECARS: u64 = 90;

const FIRST_SHEET_NAME: &str = "Résumé 東京";
const ABSENT_TABLE_NAME: &str = "收入 📊";

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

trait ExactPackageBytes {
    fn exact_bytes(&self) -> &'static [u8];
}

impl ExactPackageBytes for Package {
    fn exact_bytes(&self) -> &'static [u8] {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts every package byte");
        Box::leak(bytes.into_boxed_slice())
    }
}

struct FailsAfter {
    maximum: usize,
    bytes: Vec<u8>,
}

impl io::Write for FailsAfter {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if self.bytes.len() == self.maximum {
            return Err(io::Error::other("injected package sink failure"));
        }
        let amount = buffer.len().min(self.maximum - self.bytes.len());
        self.bytes.extend_from_slice(&buffer[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn checked_native_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

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

fn table_model(name: &str) -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_name: name.to_owned(),
        number_of_rows: 1,
        number_of_columns: 1,
        base_data_store: tst::DataStore {
            string_table: reference(SIDECARS),
            formula_table: reference(SIDECARS),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn table_info_payload(model: u64, locked: Option<bool>, unknowns: bool) -> TestResult<Vec<u8>> {
    let mut drawable = tsd::DrawableArchive {
        locked,
        ..Default::default()
    }
    .encode_to_vec();
    if unknowns {
        litchi_iwa_common::wire::append_varint_field(&mut drawable, 97, 9_997)?;
    }
    let mut payload = Vec::new();
    litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 1, &drawable)?;
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut payload,
        2,
        &reference(model).encode_to_vec(),
    )?;
    if unknowns {
        litchi_iwa_common::wire::append_varint_field(&mut payload, 98, 9_998)?;
    }
    Ok(payload)
}

fn table_info_object(
    identifier: u64,
    model: u64,
    locked: Option<bool>,
    unknowns: bool,
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![
            RawMessage {
                type_: 777,
                data: format!("before-table-info-{identifier}").into_bytes(),
            },
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: table_info_payload(model, locked, unknowns)?,
            },
            RawMessage {
                type_: 778,
                data: format!("after-table-info-{identifier}").into_bytes(),
            },
        ],
    )?;
    object.archive_info.message_infos[1].object_references = vec![model];
    Ok(object)
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
            type_: TABLE_DATA_LIST_MESSAGE_TYPE,
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

fn component_with_unknown_header(
    objects: Vec<ArchiveObject>,
    target_identifier: u64,
) -> TestResult<Vec<u8>> {
    let bytes = Archive { objects }.to_bytes()?;
    let parsed = Archive::parse(&bytes)?;
    let object = parsed
        .object(target_identifier)
        .ok_or_else(|| io::Error::other("synthetic target object is missing"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&bytes[header_offset..])?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflow"))?;
    let header_end = header_start
        .checked_add(usize::try_from(header_length)?)
        .ok_or_else(|| io::Error::other("synthetic header end overflow"))?;
    assert_eq!(header_end, data_offset);
    let mut header = bytes[header_start..header_end].to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut header, 99, 9_999)?;
    let mut rewritten = Vec::with_capacity(bytes.len() + 8);
    rewritten.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut rewritten, u64::try_from(header.len())?);
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(&bytes[data_offset..]);
    assert_eq!(Archive::parse(&rewritten)?.to_bytes()?, rewritten);
    Ok(SnappyStream::compress(&rewritten)?)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    let mut root = object(
        1,
        DOCUMENT_MESSAGE_TYPE,
        tn::DocumentArchive {
            sheets: vec![reference(FIRST_SHEET), reference(SECOND_SHEET)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    root.archive_info.message_infos[0].object_references = vec![FIRST_SHEET, SECOND_SHEET];
    let mut first_sheet = object(
        FIRST_SHEET,
        SHEET_MESSAGE_TYPE,
        tn::SheetArchive {
            name: FIRST_SHEET_NAME.to_owned(),
            drawable_infos: vec![
                reference(NON_TABLE_DRAWABLE),
                reference(ABSENT_INFO),
                reference(FALSE_INFO),
                reference(TRUE_INFO),
            ],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    first_sheet.archive_info.message_infos[0].object_references =
        vec![NON_TABLE_DRAWABLE, ABSENT_INFO, FALSE_INFO, TRUE_INFO];
    let mut second_sheet = object(
        SECOND_SHEET,
        SHEET_MESSAGE_TYPE,
        tn::SheetArchive {
            name: "Other sheet".to_owned(),
            drawable_infos: vec![reference(OTHER_INFO)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    second_sheet.archive_info.message_infos[0].object_references = vec![OTHER_INFO];

    let document = SnappyStream::compress(
        &Archive {
            objects: vec![root, first_sheet, second_sheet],
        }
        .to_bytes()?,
    )?;
    let tables = component_with_unknown_header(
        vec![
            table_info_object(ABSENT_INFO, ABSENT_MODEL, None, true)?,
            table_info_object(FALSE_INFO, FALSE_MODEL, Some(false), false)?,
            table_info_object(TRUE_INFO, TRUE_MODEL, Some(true), false)?,
            table_info_object(OTHER_INFO, OTHER_MODEL, None, false)?,
            object(
                NON_TABLE_DRAWABLE,
                99_999,
                b"opaque non-table drawable".to_vec(),
            )?,
            object(
                ABSENT_MODEL,
                TABLE_MODEL_MESSAGE_TYPE,
                table_model(ABSENT_TABLE_NAME).encode_to_vec(),
            )?,
            object(
                FALSE_MODEL,
                TABLE_MODEL_MESSAGE_TYPE,
                table_model("Explicit false").encode_to_vec(),
            )?,
            object(
                TRUE_MODEL,
                TABLE_MODEL_MESSAGE_TYPE,
                table_model("Explicit true").encode_to_vec(),
            )?,
            object(
                OTHER_MODEL,
                TABLE_MODEL_MESSAGE_TYPE,
                table_model("Other table").encode_to_vec(),
            )?,
            sidecars()?,
        ],
        ABSENT_INFO,
    )?;
    let unrelated = SnappyStream::compress(
        &Archive {
            objects: vec![ArchiveObject::new(
                900,
                vec![RawMessage {
                    type_: 999,
                    data: b"unrelated Numbers component".to_vec(),
                }],
            )?],
        }
        .to_bytes()?,
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document.as_slice()),
            (TABLES_MEMBER, tables.as_slice()),
            (UNRELATED_MEMBER, unrelated.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn legacy_package(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner =
        litchi_iwa_archive::package::to_bytes(inner_entries.iter().copied(), Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.numbers/Index.zip", inner.as_slice()),
            (
                "legacy.numbers/Data/sentinel.bin",
                b"legacy outer sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

fn component_stream(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other("missing synthetic Numbers component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn rewrite_component(
    package: &[u8],
    name: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, name)?;
    let mut archive = Archive::parse(&stream)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(name, &compressed)], Limits::default())?)
}

fn with_selected_unknown_padding(package: &[u8], length: usize) -> TestResult<Vec<u8>> {
    rewrite_component(package, TABLES_MEMBER, |archive| {
        let object = archive
            .object_mut(ABSENT_INFO)
            .ok_or_else(|| io::Error::other("missing selected table info"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-info message"))?;
        let mut payload = object.messages[index].data.clone();
        let padding = (0..length)
            .map(|offset| {
                let value = u64::try_from(offset).unwrap_or(u64::MAX);
                value.wrapping_mul(0x9e37_79b9_7f4a_7c15).rotate_left(17) as u8
            })
            .collect::<Vec<_>>();
        litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 96, &padding)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn with_dense_tables(package: &[u8], count: usize) -> TestResult<Vec<u8>> {
    let with_tables = rewrite_component(package, TABLES_MEMBER, |archive| {
        for index in 0..count {
            let index = u64::try_from(index)?;
            let info_identifier = 1_000 + index;
            let model_identifier = 2_000 + index;
            archive.objects.push(table_info_object(
                info_identifier,
                model_identifier,
                None,
                false,
            )?);
            archive.objects.push(object(
                model_identifier,
                TABLE_MODEL_MESSAGE_TYPE,
                table_model(&format!("Dense table {index}")).encode_to_vec(),
            )?);
        }
        Ok(())
    })?;
    rewrite_component(&with_tables, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(FIRST_SHEET)
            .ok_or_else(|| io::Error::other("missing first sheet"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == SHEET_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing sheet message"))?;
        let mut sheet = tn::SheetArchive::decode(object.messages[message_index].data.as_slice())?;
        for index in 0..count {
            let identifier = 1_000 + u64::try_from(index)?;
            sheet.drawable_infos.push(reference(identifier));
            object.archive_info.message_infos[message_index]
                .object_references
                .push(identifier);
        }
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: SHEET_MESSAGE_TYPE,
                data: sheet.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn with_form_based_first_sheet(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(package, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(FIRST_SHEET)
            .ok_or_else(|| io::Error::other("missing first sheet"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == SHEET_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing sheet message"))?;
        let sheet = tn::SheetArchive::decode(object.messages[message_index].data.as_slice())?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: FORM_BASED_SHEET_MESSAGE_TYPE,
                data: tn::FormBasedSheetArchive {
                    super_: sheet,
                    ..Default::default()
                }
                .encode_to_vec(),
            },
        )?;
        let message_info = &mut object.archive_info.message_infos[message_index];
        message_info
            .object_references
            .retain(|identifier| *identifier != ABSENT_INFO);
        let mut nested_drawables = FieldInfo::new(vec![1, 2]);
        nested_drawables.object_references = vec![ABSENT_INFO];
        message_info.field_infos.push(nested_drawables);
        Ok(())
    })
}

fn with_legacy_table_infos(package: &[u8], identifiers: &[u64]) -> TestResult<Vec<u8>> {
    rewrite_component(package, TABLES_MEMBER, |archive| {
        for identifier in identifiers {
            let object = archive
                .object_mut(*identifier)
                .ok_or_else(|| io::Error::other("missing table info"))?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing canonical table-info message"))?;
            let payload = object.messages[message_index].data.clone();
            object.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: LEGACY_TABLE_INFO_MESSAGE_TYPE,
                    data: payload,
                },
            )?;
        }
        Ok(())
    })
}

fn with_overlong_object_length_prefix(
    package: &[u8],
    component: &str,
    identifier: u64,
) -> TestResult<Vec<u8>> {
    let mut stream = component_stream(package, component)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Numbers object"))?;
    let offset = usize::try_from(object.header_offset)?;
    let (_length, prefix_bytes) = decode_varint_from_bytes(&stream[offset..])?;
    if prefix_bytes != 1 {
        return Err(io::Error::other("synthetic prefix is not one byte").into());
    }
    stream[offset] |= 0x80;
    stream.insert(offset + 1, 0);
    Archive::parse(&stream)?;
    let compressed = SnappyStream::compress(&stream)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(
        catalog
            .reassemble_to_bytes(&[EntryEdit::new(component, &compressed)], Limits::default())?,
    )
}

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&component_stream(package, TABLES_MEMBER)?)?;
    Ok(archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == type_)
        })
        .ok_or_else(|| io::Error::other("missing synthetic Numbers message"))?
        .data
        .clone())
}

fn object_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, TABLES_MEMBER)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Numbers object"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&stream[header_offset..])?;
    let start = header_offset + prefix_length;
    let end = start + usize::try_from(header_length)?;
    assert_eq!(end, data_offset);
    Ok(stream[start..end].to_vec())
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn assert_only_tables_component_changed(before: &[u8], after: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(before)?;
    let after = Catalog::from_bytes(after)?;
    let before_entries = before.iter().collect::<Vec<_>>();
    let after_entries = after.iter().collect::<Vec<_>>();
    assert_eq!(before_entries.len(), after_entries.len());
    for (old, new) in before_entries.into_iter().zip(after_entries) {
        assert_eq!(old.name(), new.name());
        if old.name() == TABLES_MEMBER {
            assert_ne!(old.data(), new.data());
        } else {
            assert_eq!(old.data(), new.data());
            assert_eq!(old.metadata(), new.metadata());
            assert_eq!(
                old.raw_record().local_record(),
                new.raw_record().local_record()
            );
        }
    }
    Ok(())
}

fn assert_send_sync<T: Send + Sync>(_: &T) {}
fn assert_type_send_sync<T: Send + Sync>() {}

#[test]
fn semantic_sheet_and_table_selectors_distinguish_absent_false_and_true() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;

    assert_eq!(
        package.table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Unlocked
    );
    assert_eq!(
        package.table_lock(SheetSelector::index(0), TableSelector::index(0))?,
        LockState::Unlocked
    );
    assert_eq!(
        package.table_lock(FIRST_SHEET_NAME, "Explicit false")?,
        LockState::Unlocked
    );
    assert_eq!(
        package.table_lock(SheetSelector::index(0), TableSelector::index(2))?,
        LockState::Locked
    );
    assert_eq!(
        package.table_lock("Other sheet", "Other table")?,
        LockState::Unlocked
    );

    let dense_bytes = with_dense_tables(&bytes, 32)?;
    let dense = Package::from_bytes(&dense_bytes)?;
    assert_eq!(
        dense.table_lock(SheetSelector::index(0), TableSelector::index(34))?,
        LockState::Unlocked
    );
    assert_eq!(
        dense.table_lock(FIRST_SHEET_NAME, "Dense table 31")?,
        LockState::Unlocked
    );
    assert!(matches!(
        dense.table_lock(SheetSelector::index(0), TableSelector::index(35)),
        Err(TableLockError::TableNotFound)
    ));
    Ok(())
}

#[test]
fn missing_selectors_are_typed_and_duplicate_names_fail_ingress() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    assert!(matches!(
        package.table_lock("Missing sheet", ABSENT_TABLE_NAME),
        Err(TableLockError::SheetNotFound)
    ));
    assert!(matches!(
        package.table_lock(SheetSelector::index(2), TableSelector::index(0)),
        Err(TableLockError::SheetNotFound)
    ));
    assert!(matches!(
        package.table_lock(FIRST_SHEET_NAME, "Missing table"),
        Err(TableLockError::TableNotFound)
    ));
    assert!(matches!(
        package.table_lock(SheetSelector::index(0), TableSelector::index(3)),
        Err(TableLockError::TableNotFound)
    ));

    let duplicate_name = rewrite_component(&bytes, TABLES_MEMBER, |archive| {
        let object = archive
            .object_mut(FALSE_MODEL)
            .ok_or_else(|| io::Error::other("missing false table model"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-model message"))?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: table_model(ABSENT_TABLE_NAME).encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    assert!(Package::from_bytes(&duplicate_name).is_err());
    Ok(())
}

#[test]
fn exact_noops_share_source_bytes_and_preserve_absent_vs_explicit_false() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    let source_snapshot = package.exact_bytes();
    let absent_payload = message_payload(&bytes, ABSENT_INFO, TABLE_INFO_MESSAGE_TYPE)?;
    let false_payload = message_payload(&bytes, FALSE_INFO, TABLE_INFO_MESSAGE_TYPE)?;

    for (table, state) in [
        (TableSelector::index(0), LockState::Unlocked),
        (TableSelector::index(1), LockState::Unlocked),
        (TableSelector::index(2), LockState::Locked),
    ] {
        let mut edit = package.edit_table_lock(SheetSelector::index(0), table)?;
        assert_eq!(edit.state(), state);
        edit.set_state(state);
        let commit = edit.commit()?;
        assert!(commit.patch().is_noop());
        assert_eq!(commit.patch().before(), state);
        assert_eq!(commit.patch().after(), state);
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.package().exact_bytes(), source_snapshot);
        let replay = package.apply_table_lock(commit.patch())?;
        assert!(replay.patch().is_noop());
        assert_eq!(replay.package().exact_bytes(), source_snapshot);
    }

    let absent = package
        .edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?
        .commit()?;
    assert!(absent.patch().is_noop());
    assert_eq!(
        message_payload(
            absent.package().exact_bytes(),
            ABSENT_INFO,
            TABLE_INFO_MESSAGE_TYPE
        )?,
        absent_payload
    );
    assert_eq!(
        message_payload(package.exact_bytes(), FALSE_INFO, TABLE_INFO_MESSAGE_TYPE)?,
        false_payload
    );
    Ok(())
}

#[test]
fn changed_lock_preserves_locality_unknowns_headers_and_exact_inverse() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    let source_snapshot = package.exact_bytes();
    let source_payload = message_payload(&bytes, ABSENT_INFO, TABLE_INFO_MESSAGE_TYPE)?;
    let source_before = message_payload(&bytes, ABSENT_INFO, 777)?;
    let source_after = message_payload(&bytes, ABSENT_INFO, 778)?;
    let source_non_table = message_payload(&bytes, NON_TABLE_DRAWABLE, 99_999)?;
    let source_header = object_header(&bytes, ABSENT_INFO)?;

    let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    assert_eq!(edit.state(), LockState::Unlocked);
    edit.lock();
    let commit = edit.commit()?;
    assert_eq!(commit.patch().before(), LockState::Unlocked);
    assert_eq!(commit.patch().after(), LockState::Locked);
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit
            .package()
            .table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Locked
    );
    assert_eq!(
        commit
            .package()
            .table_lock(FIRST_SHEET_NAME, "Explicit false")?,
        LockState::Unlocked
    );
    assert_eq!(
        commit.package().table_lock("Other sheet", "Other table")?,
        LockState::Unlocked
    );
    assert_only_tables_component_changed(&bytes, commit.package().exact_bytes())?;
    assert_eq!(
        message_payload(commit.package().exact_bytes(), ABSENT_INFO, 777)?,
        source_before
    );
    assert_eq!(
        message_payload(commit.package().exact_bytes(), ABSENT_INFO, 778)?,
        source_after
    );
    assert_eq!(
        message_payload(commit.package().exact_bytes(), NON_TABLE_DRAWABLE, 99_999)?,
        source_non_table
    );
    let target_payload = message_payload(
        commit.package().exact_bytes(),
        ABSENT_INFO,
        TABLE_INFO_MESSAGE_TYPE,
    )?;
    assert_eq!(
        raw_fields(&target_payload, 98)?,
        raw_fields(&source_payload, 98)?
    );
    let source_drawable = WireView::parse(&source_payload)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("source drawable field is missing"))?
        .canonical_payload()?;
    let target_drawable = WireView::parse(&target_payload)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("target drawable field is missing"))?
        .canonical_payload()?;
    assert_eq!(
        raw_fields(target_drawable, 97)?,
        raw_fields(source_drawable, 97)?
    );
    assert_eq!(
        raw_fields(
            &object_header(commit.package().exact_bytes(), ABSENT_INFO)?,
            99
        )?,
        raw_fields(&source_header, 99)?
    );
    assert_eq!(package.exact_bytes(), bytes);
    assert_eq!(package.exact_bytes(), source_snapshot);

    let applied = package.apply_table_lock(commit.patch())?;
    assert_eq!(
        applied.package().exact_bytes(),
        commit.package().exact_bytes()
    );
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), LockState::Locked);
    assert_eq!(inverse.after(), LockState::Unlocked);
    assert_eq!(inverse.inverse(), commit.patch().clone());
    let restored = commit.package().apply_table_lock(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), bytes);
    Ok(())
}

#[test]
fn unlock_method_changes_true_and_inverse_restores_explicit_true() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, "Explicit true")?;
    edit.unlock();
    let commit = edit.commit()?;
    assert_eq!(
        commit
            .package()
            .table_lock(FIRST_SHEET_NAME, "Explicit true")?,
        LockState::Unlocked
    );
    assert_eq!(
        commit
            .package()
            .apply_table_lock(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        bytes
    );
    Ok(())
}

#[test]
fn flat_legacy_table_info_edits_preserve_presence_locality_and_exact_inverse() -> TestResult {
    let base = synthetic_package()?;
    let bytes = with_legacy_table_infos(&base, &[ABSENT_INFO, FALSE_INFO])?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Unlocked
    );
    assert_eq!(
        package.table_lock(FIRST_SHEET_NAME, "Explicit false")?,
        LockState::Unlocked
    );
    let absent_payload = message_payload(&bytes, ABSENT_INFO, LEGACY_TABLE_INFO_MESSAGE_TYPE)?;
    let false_payload = message_payload(&bytes, FALSE_INFO, LEGACY_TABLE_INFO_MESSAGE_TYPE)?;
    let absent_drawable = WireView::parse(&absent_payload)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("legacy absent drawable is missing"))?
        .canonical_payload()?;
    let false_drawable = WireView::parse(&false_payload)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("legacy false drawable is missing"))?
        .canonical_payload()?;
    assert!(raw_fields(absent_drawable, 5)?.is_empty());
    assert_eq!(raw_fields(false_drawable, 5)?.len(), 1);

    for (table, identifier) in [
        (ABSENT_TABLE_NAME, ABSENT_INFO),
        ("Explicit false", FALSE_INFO),
    ] {
        let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, table)?;
        edit.lock();
        let changed = edit.commit()?;
        assert_eq!(changed.patch().before(), LockState::Unlocked);
        assert_eq!(changed.patch().after(), LockState::Locked);
        assert!(changed.diagnostics().changed());
        assert_eq!(changed.diagnostics().touched_components(), 1);
        assert!(changed.diagnostics().full_reparse_performed());
        assert_eq!(
            changed.package().table_lock(FIRST_SHEET_NAME, table)?,
            LockState::Locked
        );
        message_payload(
            changed.package().exact_bytes(),
            identifier,
            LEGACY_TABLE_INFO_MESSAGE_TYPE,
        )?;
        assert_only_tables_component_changed(&bytes, changed.package().exact_bytes())?;
        let restored = changed
            .package()
            .apply_table_lock(&changed.patch().inverse())?;
        assert_eq!(restored.package().exact_bytes(), bytes);
    }

    let dual = rewrite_component(&base, TABLES_MEMBER, |archive| {
        let object = archive
            .object_mut(ABSENT_INFO)
            .ok_or_else(|| io::Error::other("missing selected table info"))?;
        let canonical = object
            .messages
            .iter()
            .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing canonical table-info message"))?;
        object.push_message(RawMessage {
            type_: LEGACY_TABLE_INFO_MESSAGE_TYPE,
            data: canonical.data.clone(),
        })?;
        Ok(())
    })?;
    match Package::from_bytes(&dual) {
        Err(_ingress_error) => {},
        Ok(dual_package) => {
            let mut edit = dual_package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
            edit.lock();
            assert!(matches!(edit.commit(), Err(TableLockError::InvalidSource)));
        },
    }
    Ok(())
}

#[test]
fn stale_tampered_replayed_and_cross_selector_patches_conflict() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    let mut first = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    first.lock();
    let first = first.commit()?;
    let mut second = package.edit_table_lock("Other sheet", "Other table")?;
    second.lock();
    let second = second.commit()?;

    assert!(matches!(
        first.package().apply_table_lock(first.patch()),
        Err(TableLockError::PatchConflict)
    ));
    assert!(matches!(
        first.package().apply_table_lock(second.patch()),
        Err(TableLockError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_table_lock(&first.patch().inverse()),
        Err(TableLockError::PatchConflict)
    ));

    let catalog = Catalog::from_bytes(&bytes)?;
    let tampered_bytes = catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            "Data/sentinel.bin",
            b"tampered unrelated sentinel",
        )],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_bytes)?;
    assert_eq!(
        tampered.table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Unlocked
    );
    assert!(matches!(
        tampered.apply_table_lock(first.patch()),
        Err(TableLockError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn malformed_selected_and_rooted_aliases_fail_closed_while_detached_refs_are_preserved()
-> TestResult {
    let bytes = synthetic_package()?;
    let duplicate_payload = rewrite_component(&bytes, TABLES_MEMBER, |archive| {
        let object = archive
            .object_mut(ABSENT_INFO)
            .ok_or_else(|| io::Error::other("missing selected table info"))?;
        let duplicate = object
            .messages
            .iter()
            .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-info message"))?
            .clone();
        object.push_message(duplicate)?;
        Ok(())
    })?;
    assert!(Package::from_bytes(&duplicate_payload).is_err());

    let missing_metadata = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(FIRST_SHEET)
            .ok_or_else(|| io::Error::other("missing first sheet"))?;
        object.archive_info.message_infos[0]
            .object_references
            .clear();
        Ok(())
    })?;
    let other_sheet_alias = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(SECOND_SHEET)
            .ok_or_else(|| io::Error::other("missing second sheet"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == SHEET_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing sheet message"))?;
        let mut value = tn::SheetArchive::decode(object.messages[index].data.as_slice())?;
        value.drawable_infos.push(reference(ABSENT_INFO));
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: SHEET_MESSAGE_TYPE,
                data: value.encode_to_vec(),
            },
        )?;
        object.archive_info.message_infos[index]
            .object_references
            .push(ABSENT_INFO);
        Ok(())
    })?;
    let detached_payload_alias = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let mut detached = object(
            404,
            SHEET_MESSAGE_TYPE,
            tn::SheetArchive {
                name: "Detached alias".to_owned(),
                drawable_infos: vec![reference(ABSENT_INFO)],
                ..Default::default()
            }
            .encode_to_vec(),
        )?;
        detached.archive_info.message_infos[0].object_references = vec![ABSENT_INFO];
        archive.objects.push(detached);
        Ok(())
    })?;

    assert!(Package::from_bytes(&other_sheet_alias).is_err());

    let package = Package::from_bytes(&missing_metadata)?;
    let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    edit.lock();
    assert!(matches!(edit.commit(), Err(TableLockError::InvalidSource)));
    assert_eq!(package.exact_bytes(), missing_metadata);

    let package = Package::from_bytes(&detached_payload_alias)?;
    let semantics = package.sheets().to_vec();
    let document_component = component_stream(&detached_payload_alias, DOCUMENT_MEMBER)?;
    let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    edit.lock();
    let changed = edit.commit()?;
    assert_eq!(
        changed
            .package()
            .table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Locked
    );
    assert_eq!(changed.package().sheets(), semantics.as_slice());
    assert_eq!(
        component_stream(changed.package().exact_bytes(), DOCUMENT_MEMBER)?,
        document_component
    );
    assert_eq!(package.exact_bytes(), detached_payload_alias);
    Ok(())
}

#[test]
fn rooted_form_based_sheet_nested_reference_path_can_change_and_invert() -> TestResult {
    let base = synthetic_package()?;
    let bytes = with_form_based_first_sheet(&base)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Unlocked
    );

    let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    edit.lock();
    let changed = edit.commit()?;
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert_eq!(
        changed
            .package()
            .table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Locked
    );
    assert_eq!(
        component_stream(changed.package().exact_bytes(), DOCUMENT_MEMBER)?,
        component_stream(&bytes, DOCUMENT_MEMBER)?
    );
    let restored = changed
        .package()
        .apply_table_lock(&changed.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), bytes);
    Ok(())
}

#[test]
fn noncanonical_frames_and_merge_diff_metadata_are_never_normalized() -> TestResult {
    let bytes = synthetic_package()?;
    let noncanonical = with_overlong_object_length_prefix(&bytes, TABLES_MEMBER, OTHER_MODEL)?;
    let merge_metadata = rewrite_component(&bytes, TABLES_MEMBER, |archive| {
        let object = archive
            .object_mut(ABSENT_INFO)
            .ok_or_else(|| io::Error::other("missing selected table info"))?;
        object.archive_info.should_merge = Some(true);
        Ok(())
    })?;
    let diff_metadata = rewrite_component(&bytes, TABLES_MEMBER, |archive| {
        let object = archive
            .object_mut(ABSENT_INFO)
            .ok_or_else(|| io::Error::other("missing selected table info"))?;
        object.archive_info.message_infos[1].base_message_index = Some(0);
        Ok(())
    })?;

    for adversarial in [noncanonical, merge_metadata, diff_metadata] {
        let package = Package::from_bytes(&adversarial)?;
        let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
        edit.lock();
        assert!(matches!(edit.commit(), Err(TableLockError::InvalidSource)));
        assert_eq!(package.exact_bytes(), adversarial);
    }
    Ok(())
}

#[test]
fn legacy_nested_index_reads_and_noops_but_refuses_changed_publication() -> TestResult {
    let flat = synthetic_package()?;
    let legacy = legacy_package(&flat)?;
    let package = Package::from_bytes(&legacy)?;
    assert_eq!(
        package.table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?,
        LockState::Unlocked
    );

    let mut noop = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    noop.unlock();
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), legacy);

    let mut changed = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    changed.lock();
    assert!(matches!(
        changed.commit(),
        Err(TableLockError::UnsupportedSource)
    ));
    assert_eq!(package.exact_bytes(), legacy);
    Ok(())
}

#[test]
fn output_limit_is_typed_and_failure_atomic() -> TestResult {
    let base = synthetic_package()?;
    let mut growth_fixture = None;
    for padding_length in 0..64 {
        let bytes = with_selected_unknown_padding(&base, padding_length)?;
        let package = Package::from_bytes(&bytes)?;
        let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
        edit.lock();
        let changed = edit.commit()?;
        if changed.package().exact_bytes().len() > bytes.len() {
            growth_fixture = Some(bytes);
            break;
        }
    }
    let bytes = growth_fixture
        .ok_or_else(|| io::Error::other("could not construct a growing lock-edit fixture"))?;
    let input_bytes = u64::try_from(bytes.len())?;
    let limits = Limits::new(input_bytes, 32, 1024 * 1024, 1024 * 1024, 1024 * 1024)?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    edit.lock();
    match edit.commit() {
        Err(TableLockError::LimitExceeded {
            kind: TableLockLimitKind::OutputBytes,
            ..
        }) => {},
        Err(error) => return Err(error.into()),
        Ok(_commit) => return Err(io::Error::other("output limit was not enforced").into()),
    }
    assert_eq!(package.exact_bytes(), bytes);
    Ok(())
}

#[test]
fn public_transactions_are_send_sync_and_concurrent_readers_are_deterministic() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Arc::new(Package::from_bytes(&bytes)?);
    let source_snapshot = package.exact_bytes();
    let mut handles = Vec::new();
    for index in 0..8 {
        let package = Arc::clone(&package);
        handles.push(std::thread::spawn(move || {
            if index % 2 == 0 {
                package.table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)
            } else {
                package.table_lock(SheetSelector::index(0), TableSelector::index(2))
            }
        }));
    }
    for (index, handle) in handles.into_iter().enumerate() {
        let state = handle
            .join()
            .map_err(|_panic| io::Error::other("table-lock reader panicked"))??;
        assert_eq!(
            state,
            if index % 2 == 0 {
                LockState::Unlocked
            } else {
                LockState::Locked
            }
        );
    }
    assert_eq!(package.exact_bytes(), bytes);
    assert_eq!(package.exact_bytes(), source_snapshot);
    assert_send_sync(package.as_ref());
    assert_type_send_sync::<TableLockEdit<'static>>();
    assert_type_send_sync::<TableLockCommit>();
    assert_type_send_sync::<TableLockPatch>();
    assert_type_send_sync::<TableLockDiagnostics>();
    assert_type_send_sync::<TableLockError>();
    Ok(())
}

#[test]
fn debug_and_error_text_are_redacted_from_authored_and_physical_names() -> TestResult {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    let mut edit = package.edit_table_lock(FIRST_SHEET_NAME, ABSENT_TABLE_NAME)?;
    edit.lock();
    let edit_debug = format!("{edit:?}");
    let commit = edit.commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    for redacted in [
        FIRST_SHEET_NAME,
        ABSENT_TABLE_NAME,
        DOCUMENT_MEMBER,
        TABLES_MEMBER,
        "unrelated ZIP sentinel",
    ] {
        assert!(!edit_debug.contains(redacted));
        assert!(!patch_debug.contains(redacted));
    }
    for error in [
        TableLockError::SheetNotFound,
        TableLockError::TableNotFound,
        TableLockError::PatchConflict,
        TableLockError::LimitExceeded {
            kind: TableLockLimitKind::OutputBytes,
            observed: 11,
            maximum: 10,
        },
    ] {
        let debug = format!("{error:?}");
        let display = error.to_string();
        for redacted in [FIRST_SHEET_NAME, ABSENT_TABLE_NAME, TABLES_MEMBER] {
            assert!(!debug.contains(redacted));
            assert!(!display.contains(redacted));
        }
    }
    Ok(())
}

#[test]
fn checked_native_fixture_lock_preserves_semantics_and_has_exact_inverse() -> TestResult {
    let bytes = std::fs::read(checked_native_fixture_path())?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.table_lock("Sheet 1", "Table 1")?,
        LockState::Unlocked
    );
    let semantic_before = package.sheets().to_vec();
    let table_before = semantic_before
        .first()
        .and_then(|sheet| sheet.tables().next())
        .ok_or_else(|| io::Error::other("checked Numbers fixture has no first table"))?;
    assert_eq!(table_before.name(), "Table 1");
    assert!(table_before.cell_count() > 0);

    let mut edit = package.edit_table_lock("Sheet 1", "Table 1")?;
    edit.lock();
    let changed = edit.commit()?;
    assert_eq!(
        changed.package().table_lock("Sheet 1", "Table 1")?,
        LockState::Locked
    );
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(changed.package().sheets(), semantic_before.as_slice());
    assert_eq!(package.exact_bytes(), bytes);
    let mut written = Vec::new();
    changed.package().write_to(&mut written)?;
    assert_eq!(written, changed.package().exact_bytes());
    let fail_after = changed.package().exact_bytes().len() / 2;
    let mut failing = FailsAfter {
        maximum: fail_after,
        bytes: Vec::new(),
    };
    let write_error = changed
        .package()
        .write_to(&mut failing)
        .expect_err("injected sink must fail after its checked prefix");
    assert_eq!(write_error.bytes_written(), fail_after);
    assert_eq!(failing.bytes, changed.package().exact_bytes()[..fail_after]);
    assert_eq!(write_error.io_error().kind(), io::ErrorKind::Other);
    assert_eq!(write_error.into_io_error().kind(), io::ErrorKind::Other);

    let restored = changed
        .package()
        .apply_table_lock(&changed.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), bytes);
    assert_eq!(restored.package().sheets(), semantic_before.as_slice());
    Ok(())
}
