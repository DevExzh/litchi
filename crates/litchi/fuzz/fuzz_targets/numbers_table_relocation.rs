#![no_main]

//! Bounded selector-first fuzzing for Numbers table ownership relocation.
//!
//! The one-time fixture setup builds a small strict two-sheet package with the
//! existing archive/core/protobuf dependencies. Every fuzzed operation below
//! uses the focused Numbers package and semantic selectors; native object
//! identifiers, protobuf payloads, and archive member names never cross the
//! fuzz boundary.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageError, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use litchi_iwa_archive::{Limits as ArchiveLimits, package::to_bytes};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsd, tsp, tst};
use prost::Message as _;

const MAX_INPUT_BYTES: u64 = 512 * 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES as usize + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const MAX_OBJECTS: usize = 4 * 1024;
const MAX_SHEETS: usize = 128;
const MAX_TABLES: usize = 512;
const MAX_REFERENCES: usize = 8 * 1024;
const MAX_MATERIALIZED_CELLS: usize = 64 * 1024;
const MAX_TEXT_BYTES: usize = 512 * 1024;
const MAX_COMMAND_BYTES: usize = 64;
const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const TABLES_MEMBER: &str = "Index/Tables.iwa";
const DOCUMENT_ID: u64 = 1;
const SOURCE_SHEET_ID: u64 = 2;
const DESTINATION_SHEET_ID: u64 = 3;
const TABLE_INFO_ID: u64 = 4;
const TABLE_MODEL_ID: u64 = 5;
const SIDECAR_ID: u64 = 6;
const SHEET_MESSAGE_TYPE: u32 = 2;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const SOURCE_SHEET: &str = "Fuzz relocation source";
const DESTINATION_SHEET: &str = "Fuzz relocation destination";
const TABLE_NAME: &str = "Fuzz relocation table";

fuzz_target!(|data: &[u8]| {
    let command = command_input(data);
    exercise_bounded_ingress(data);
    exercise_native_relocation(&command);
    exercise_input_limit();
});

/// The command prefix is intentionally tiny and independent from package
/// bytes:
///
/// * byte 0, low two bits: 0 = same-sheet no-op, 1 = changed move, 2 = move
///   with patch inverse, 3 = invalid semantic selectors;
/// * byte 1, low bit: source sheet selector by position or exact name;
/// * byte 2, low bit: table selector by position or exact name;
/// * byte 3, low bit: destination sheet selector by position or exact name.
///
/// Checked-in `hex:` recipes are decoded here; arbitrary bytes remain the
/// bounded Numbers package-ingress input in `exercise_bounded_ingress`.
fn command_input(data: &[u8]) -> Vec<u8> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded).unwrap_or_default();
    }
    data.get(..data.len().min(MAX_COMMAND_BYTES))
        .unwrap_or(data)
        .to_vec()
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_COMMAND_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_COMMAND_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn options() -> PackageReadOptions {
    static OPTIONS: OnceLock<PackageReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = PackageLimits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid relocation archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| unreachable!("valid relocation semantic limits: {error}"))
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid relocation projection limits: {error}")
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_fixture() -> &'static [u8] {
    static FIXTURE: OnceLock<Box<[u8]>> = OnceLock::new();
    FIXTURE
        .get_or_init(|| {
            let mut document = fixture_object(
                DOCUMENT_ID,
                1,
                tn::DocumentArchive {
                    sheets: vec![reference(SOURCE_SHEET_ID), reference(DESTINATION_SHEET_ID)],
                    ..Default::default()
                }
                .encode_to_vec(),
            );
            document.archive_info.message_infos[0].object_references =
                vec![SOURCE_SHEET_ID, DESTINATION_SHEET_ID];
            let mut sheet_path = FieldInfo::new(vec![1]);
            sheet_path.object_references = vec![SOURCE_SHEET_ID, DESTINATION_SHEET_ID];
            document.archive_info.message_infos[0].field_infos = vec![sheet_path];

            let document_component = SnappyStream::compress(
                &Archive {
                    objects: vec![
                        document,
                        fixture_sheet(SOURCE_SHEET_ID, SOURCE_SHEET, &[TABLE_INFO_ID]),
                        fixture_sheet(DESTINATION_SHEET_ID, DESTINATION_SHEET, &[]),
                    ],
                }
                .to_bytes()
                .unwrap_or_else(|error| panic!("Numbers relocation document fixture: {error}")),
            )
            .unwrap_or_else(|error| panic!("Numbers relocation document compression: {error}"));

            let tables_component = SnappyStream::compress(
                &Archive {
                    objects: vec![
                        fixture_table_info(),
                        fixture_object(
                            TABLE_MODEL_ID,
                            TABLE_MODEL_MESSAGE_TYPE,
                            fixture_table_model().encode_to_vec(),
                        ),
                        fixture_sidecars(),
                    ],
                }
                .to_bytes()
                .unwrap_or_else(|error| panic!("Numbers relocation table fixture: {error}")),
            )
            .unwrap_or_else(|error| panic!("Numbers relocation table compression: {error}"));

            to_bytes(
                [
                    (DOCUMENT_MEMBER, document_component.as_slice()),
                    (TABLES_MEMBER, tables_component.as_slice()),
                    (
                        "Data/relocation-sentinel.bin",
                        b"relocation sentinel".as_slice(),
                    ),
                ],
                ArchiveLimits::default(),
            )
            .unwrap_or_else(|error| panic!("Numbers relocation package fixture: {error}"))
            .into_boxed_slice()
        })
        .as_ref()
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn fixture_object(identifier: u64, message_type: u32, data: Vec<u8>) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )
    .unwrap_or_else(|error| panic!("Numbers relocation fixture object: {error}"))
}

fn fixture_sheet(identifier: u64, name: &str, drawable_identifiers: &[u64]) -> ArchiveObject {
    let mut sheet = fixture_object(
        identifier,
        SHEET_MESSAGE_TYPE,
        tn::SheetArchive {
            name: name.to_owned(),
            drawable_infos: drawable_identifiers
                .iter()
                .copied()
                .map(reference)
                .collect(),
            ..Default::default()
        }
        .encode_to_vec(),
    );
    sheet.archive_info.message_infos[0].object_references = drawable_identifiers.to_vec();
    let mut drawable_path = FieldInfo::new(vec![2]);
    drawable_path.object_references = drawable_identifiers.to_vec();
    sheet.archive_info.message_infos[0].field_infos = vec![drawable_path];
    sheet
}

fn fixture_table_info() -> ArchiveObject {
    let payload = tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            parent: Some(reference(SOURCE_SHEET_ID)),
            ..Default::default()
        },
        table_model: reference(TABLE_MODEL_ID),
        ..Default::default()
    }
    .encode_to_vec();
    let mut table_info = fixture_object(TABLE_INFO_ID, TABLE_INFO_MESSAGE_TYPE, payload);
    table_info.archive_info.message_infos[0].object_references =
        vec![SOURCE_SHEET_ID, TABLE_MODEL_ID];
    let mut parent_path = FieldInfo::new(vec![1, 2]);
    parent_path.object_references = vec![SOURCE_SHEET_ID];
    table_info.archive_info.message_infos[0].field_infos = vec![parent_path];
    table_info
}

fn fixture_table_model() -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: "fuzz-relocation-table".to_owned(),
        table_name: TABLE_NAME.to_owned(),
        number_of_rows: 2,
        number_of_columns: 2,
        base_data_store: tst::DataStore {
            string_table: reference(SIDECAR_ID),
            formula_table: reference(SIDECAR_ID),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn fixture_sidecars() -> ArchiveObject {
    let messages = [
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
    .collect();
    ArchiveObject::new(SIDECAR_ID, messages)
        .unwrap_or_else(|error| panic!("Numbers relocation fixture sidecars: {error}"))
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(native_fixture(), options())
            .unwrap_or_else(|error| panic!("native Numbers relocation seed must open: {error}"))
    })
}

fn exercise_bounded_ingress(data: &[u8]) {
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => {
            black_box(package.sheets().len());
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, options()) {
        Err(PackageError::InputTooLarge { observed, maximum }) => {
            assert_eq!(observed, OVERSIZED_INPUT_BYTES as u64);
            assert_eq!(maximum, MAX_INPUT_BYTES);
            black_box((observed, maximum));
        },
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized Numbers relocation input must be rejected"),
    }
}

fn exercise_native_relocation(command: &[u8]) {
    let package = native_package();
    let source = package
        .sheets()
        .first()
        .unwrap_or_else(|| panic!("native Numbers relocation source sheet is missing"));
    let destination = package
        .sheets()
        .get(1)
        .unwrap_or_else(|| panic!("native Numbers relocation destination is missing"));
    let table = source
        .tables()
        .next()
        .unwrap_or_else(|| panic!("native Numbers relocation table is missing"));
    let source_name = source.name().to_owned();
    let destination_name = destination.name().to_owned();
    let table_name = table.name().to_owned();
    let baseline_bytes = package_bytes(package);

    // Exercise both semantic forms even when the fuzz command selects a
    // different operation.  Same-sheet relocation must retain exact bytes.
    assert_same_sheet_noop(package, &baseline_bytes, &table_name, &source_name);

    match command.first().copied().unwrap_or_default() & 3 {
        0 => assert_same_sheet_noop(package, &baseline_bytes, &table_name, &source_name),
        1 => exercise_changed_move(
            package,
            &baseline_bytes,
            &table_name,
            &source_name,
            &destination_name,
            command,
            false,
        ),
        2 => exercise_changed_move(
            package,
            &baseline_bytes,
            &table_name,
            &source_name,
            &destination_name,
            command,
            true,
        ),
        _ => exercise_invalid_selectors(package, &baseline_bytes, &table_name, &destination_name),
    }
}

fn assert_same_sheet_noop(
    package: &Package,
    baseline_bytes: &[u8],
    table_name: &str,
    source_name: &str,
) {
    let by_index = package
        .move_table(
            SheetSelector::index(0),
            TableSelector::index(0),
            SheetSelector::index(0),
        )
        .unwrap_or_else(|error| panic!("same-sheet index no-op failed: {error}"));
    assert!(by_index.patch().is_noop());
    assert!(!by_index.diagnostics().changed());
    assert_eq!(by_index.diagnostics().touched_components(), 0);
    assert_eq!(package_bytes(by_index.package()), baseline_bytes);
    let replay = package
        .apply_table_relocation(by_index.patch())
        .unwrap_or_else(|error| panic!("same-sheet patch replay failed: {error}"));
    assert!(replay.patch().is_noop());
    assert_eq!(package_bytes(replay.package()), baseline_bytes);

    let by_name = package
        .move_table(
            SheetSelector::name(source_name),
            TableSelector::name(table_name),
            SheetSelector::name(source_name),
        )
        .unwrap_or_else(|error| panic!("same-sheet name no-op failed: {error}"));
    assert!(by_name.patch().is_noop());
    assert_eq!(package_bytes(by_name.package()), baseline_bytes);
}

fn exercise_changed_move(
    package: &Package,
    baseline_bytes: &[u8],
    table_name: &str,
    source_name: &str,
    destination_name: &str,
    command: &[u8],
    with_inverse: bool,
) {
    let source_selector = if control(command, 1) & 1 == 0 {
        SheetSelector::index(0)
    } else {
        SheetSelector::name(source_name)
    };
    let table_selector = if control(command, 2) & 1 == 0 {
        TableSelector::index(0)
    } else {
        TableSelector::name(table_name)
    };
    let destination_selector = if control(command, 3) & 1 == 0 {
        SheetSelector::index(1)
    } else {
        SheetSelector::name(destination_name)
    };
    let commit = package
        .move_table(source_selector, table_selector, destination_selector)
        .unwrap_or_else(|error| panic!("changed Numbers table move failed: {error}"));
    let target_bytes = package_bytes(commit.package());
    assert_ne!(target_bytes, baseline_bytes);
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().touched_components() > 0);
    assert_table_location(commit.package(), source_name, destination_name, table_name);

    let applied = package
        .apply_table_relocation(commit.patch())
        .unwrap_or_else(|error| panic!("fresh relocation patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert!(
        commit
            .package()
            .apply_table_relocation(commit.patch())
            .is_err(),
        "relocation patch must conflict on its target"
    );

    if with_inverse {
        let inverse = commit.patch().inverse();
        assert_eq!(inverse.inverse(), commit.patch().clone());
        assert!(
            package.apply_table_relocation(&inverse).is_err(),
            "relocation inverse must conflict on its source"
        );
        let restored = commit
            .package()
            .apply_table_relocation(&inverse)
            .unwrap_or_else(|error| panic!("relocation inverse must apply: {error}"));
        assert_eq!(package_bytes(restored.package()), baseline_bytes);
        assert_table_location(
            restored.package(),
            destination_name,
            source_name,
            table_name,
        );
    }
}

fn assert_invalid_move(
    package: &Package,
    baseline_bytes: &[u8],
    source: SheetSelector<'_>,
    table: TableSelector<'_>,
    destination: SheetSelector<'_>,
) {
    let error = package
        .move_table(source, table, destination)
        .expect_err("invalid Numbers table move unexpectedly succeeded");
    observe_error(error);
    assert_eq!(package_bytes(package), baseline_bytes);
}

fn exercise_invalid_selectors(
    package: &Package,
    baseline_bytes: &[u8],
    table_name: &str,
    destination_name: &str,
) {
    assert_invalid_move(
        package,
        baseline_bytes,
        SheetSelector::index(usize::MAX),
        TableSelector::index(0),
        SheetSelector::index(1),
    );
    assert_invalid_move(
        package,
        baseline_bytes,
        SheetSelector::index(0),
        TableSelector::index(usize::MAX),
        SheetSelector::index(1),
    );
    assert_invalid_move(
        package,
        baseline_bytes,
        SheetSelector::index(0),
        TableSelector::index(0),
        SheetSelector::index(usize::MAX),
    );
    assert_invalid_move(
        package,
        baseline_bytes,
        SheetSelector::name("__missing_relocation_sheet__"),
        TableSelector::name(table_name),
        SheetSelector::name(destination_name),
    );
}

fn assert_table_location(
    package: &Package,
    source_name: &str,
    destination_name: &str,
    table_name: &str,
) {
    let source = package
        .sheet(SheetSelector::name(source_name))
        .unwrap_or_else(|error| panic!("moved Numbers source lookup failed: {error}"))
        .unwrap_or_else(|| panic!("moved Numbers source sheet disappeared"));
    let destination = package
        .sheet(SheetSelector::name(destination_name))
        .unwrap_or_else(|error| panic!("moved Numbers destination lookup failed: {error}"))
        .unwrap_or_else(|| panic!("moved Numbers destination sheet disappeared"));
    assert!(
        source.tables().next().is_none(),
        "moved Numbers source sheet still owns a table"
    );
    assert!(
        destination.tables().any(|table| table.name() == table_name),
        "moved Numbers destination sheet does not own the table"
    );
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Numbers relocation package failed: {error}"));
    bytes
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
