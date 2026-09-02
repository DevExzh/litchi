#![no_main]
#![allow(
    clippy::too_many_lines,
    reason = "The bounded synthetic package fixture keeps its graph in one auditable function."
)]

//! Bounded selector-first fuzzing for Keynote physical table sorting.
//!
//! The target keeps package ingress and the physical row-reorder transaction
//! on separate, bounded paths. Arbitrary bytes are offered to the real
//! Keynote reader first; a small command prefix is then replayed against the
//! checked-in source-built and locked table packages. The latter makes the
//! selector, persisted-order admission, whole-table/selected-row execution,
//! immutable commit, exact-source apply, inverse, conflict, and failure
//! atomicity paths reachable even when CRC-protected ZIP mutations fail early.
//!
//! Commands use only semantic selectors and archive-free sort values. Native
//! object identifiers, generated protobuf messages, archive names, and raw
//! row IDs never cross this fuzz target's public API boundary.

use std::error::Error as StdError;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadOptions, SemanticLimits, SlideSelector, SlideTablePhysicalSortCommit,
    slide::table::{
        TableSelector,
        sort::{ColumnIndex, Direction, Order, RowRange, Rule, Scope},
    },
};
use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    package::{Catalog, to_bytes as package_to_bytes},
};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsce, tsd, tsk, tsp, tst, tswp};
use prost::Message as _;

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES as usize + 1;
const MAX_ENTRIES: usize = 256;
const MAX_ENTRY_BYTES: u64 = 2 * 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 8 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 2 * 1024 * 1024;
const MAX_OBJECTS: usize = 16 * 1024;
const MAX_SLIDES: usize = 512;
const MAX_REFERENCES: usize = 32 * 1024;
const MAX_TEXT_STORAGES: usize = 8 * 1024;
const MAX_TEXT_FRAGMENTS: usize = 32 * 1024;
const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;
const MAX_COMMAND_BYTES: usize = 1024;
const PRIVATE_SLIDE_NAME: &str = "__litchi_private_physical_sort_slide_missing__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_physical_sort_input_5a2__";
const SOURCE_BUILT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_slide_table_headers/source_built.hex");
const LOCKED_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_slide_table_headers/locked.hex");

const SYNTHETIC_DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SYNTHETIC_METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SYNTHETIC_SENTINEL_MEMBER: &str = "Data/physical-sort-sentinel.bin";
const SYNTHETIC_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const SYNTHETIC_DOCUMENT: u64 = 1;
const SYNTHETIC_SHOW: u64 = 2;
const SYNTHETIC_SLIDE_NODE: u64 = 3;
const SYNTHETIC_SLIDE: u64 = 4;
const SYNTHETIC_TABLE_INFO: u64 = 100;
const SYNTHETIC_TABLE_MODEL: u64 = 101;
const SYNTHETIC_NON_TABLE_DRAWABLE: u64 = 102;
const SYNTHETIC_ROW_HEADERS: u64 = 110;
const SYNTHETIC_COLUMN_HEADERS: u64 = 111;
const SYNTHETIC_TILE: u64 = 112;
const SYNTHETIC_STRINGS: u64 = 113;
const SYNTHETIC_STYLES: u64 = 114;
const SYNTHETIC_FORMULAS: u64 = 115;
const SYNTHETIC_FORMATS: u64 = 116;
const SYNTHETIC_UID_MAP: u64 = 117;
const SYNTHETIC_STROKE_SIDECAR: u64 = 118;
const SYNTHETIC_FORMULA_OWNER_DEPENDENCIES: u64 = 119;
const SYNTHETIC_TITLE_STYLE: u64 = 120;
const SYNTHETIC_SHAPE_STYLE: u64 = 121;
const SYNTHETIC_METADATA_OBJECT: u64 = 900;
const SYNTHETIC_TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const SYNTHETIC_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const SYNTHETIC_TILE_MESSAGE_TYPE: u32 = 6_002;
const SYNTHETIC_TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const SYNTHETIC_HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const SYNTHETIC_COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE: u32 = 6_200;
const SYNTHETIC_COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const SYNTHETIC_STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const SYNTHETIC_FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;
const SYNTHETIC_SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SYNTHETIC_METADATA_MESSAGE_TYPE: u32 = 11_006;
const SYNTHETIC_BASE_UID_FIELD: u32 = 46;
const SYNTHETIC_STROKE_FIELD: u32 = 49;
const SYNTHETIC_ROW_COUNT: u32 = 6;
const SYNTHETIC_COLUMN_COUNT: u32 = 2;
const SYNTHETIC_HEADER_ROWS: u32 = 1;
const SYNTHETIC_FOOTER_ROWS: u32 = 1;
const SYNTHETIC_TILE_UNKNOWN_MARKER: &[u8] = b"physical-sort-tile-unknown";
const SYNTHETIC_HEADER_UNKNOWN_MARKER: &[u8] = b"physical-sort-header-unknown";
const SYNTHETIC_MODEL_FIELD_INVALID_PAYLOAD: &[u8] = &[0x0a, 0x00];
const SYNTHETIC_PRE_BNC_EMPTY_CELL: [u8; 12] = [4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
const SYNTHETIC_PRE_BNC_OFFSETS: [u8; 4] = [0, 0, 12, 0];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SyntheticVariant {
    PaddedOffsets,
    PairedPreBnc,
    StringSidecars,
    UidAliases,
    LegacyUidType,
    UnknownMutableRoot,
    UnknownMutableHeaderRoot,
    ModelField39,
    ModelField45,
    ModelField84,
    ModelField93,
    AggregateCellCountMismatch,
    WidePreBncDisagreement,
    UnknownModelRoot,
}

impl SyntheticVariant {
    const fn expects_rejection(self) -> bool {
        matches!(
            self,
            Self::StringSidecars
                | Self::UidAliases
                | Self::UnknownMutableRoot
                | Self::UnknownMutableHeaderRoot
                | Self::ModelField39
                | Self::ModelField45
                | Self::ModelField84
                | Self::ModelField93
                | Self::AggregateCellCountMismatch
                | Self::WidePreBncDisagreement
                | Self::UnknownModelRoot
        )
    }

    const fn model_field(self) -> Option<u32> {
        match self {
            Self::ModelField39 => Some(39),
            Self::ModelField45 => Some(45),
            Self::ModelField84 => Some(84),
            Self::ModelField93 => Some(93),
            _ => None,
        }
    }
}

struct SyntheticPackage {
    variant: SyntheticVariant,
    package: Package,
}

fuzz_target!(|data: &[u8]| {
    // Keep arbitrary bytes on the real bounded package-ingress path. The
    // command interpretation below is deliberately independent of ZIP
    // mutation survival so that deep transaction coverage is stable.
    let command = command_input(data);
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_untrusted_package(&package, &command),
        Err(error) => observe_error(error),
    }

    for (index, package) in source_packages().iter().enumerate() {
        exercise_package(package, &command, index == 1 || index == 4);
    }
    for synthetic in synthetic_variant_packages() {
        if synthetic.variant.expects_rejection() {
            exercise_rejected_synthetic_variant(&synthetic.package, synthetic.variant, &command);
        } else {
            exercise_package(&synthetic.package, &command, false);
            exercise_synthetic_variant(&synthetic.package, synthetic.variant, &command);
        }
    }
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_archive_limit();
    exercise_semantic_limit(&command);
});

fn fuzz_options() -> ReadOptions {
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid physical-sort archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid physical-sort semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn source_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES.get_or_init(|| {
        [
            source_built_bytes(),
            locked_bytes(),
            synthetic_entire_bytes(),
            synthetic_selected_bytes(),
            synthetic_locked_bytes(),
        ]
        .into_iter()
        .map(|source| {
            Package::from_bytes_with_options(source, fuzz_options()).unwrap_or_else(|error| {
                panic!("source-built physical-sort package must open: {error}")
            })
        })
        .collect::<Vec<_>>()
        .into_boxed_slice()
    })
}

fn synthetic_variant_packages() -> &'static [SyntheticPackage] {
    static PACKAGES: OnceLock<Box<[SyntheticPackage]>> = OnceLock::new();
    PACKAGES.get_or_init(|| {
        [
            SyntheticVariant::PaddedOffsets,
            SyntheticVariant::PairedPreBnc,
            SyntheticVariant::StringSidecars,
            SyntheticVariant::UidAliases,
            SyntheticVariant::LegacyUidType,
            SyntheticVariant::UnknownMutableRoot,
            SyntheticVariant::UnknownMutableHeaderRoot,
            SyntheticVariant::ModelField39,
            SyntheticVariant::ModelField45,
            SyntheticVariant::ModelField84,
            SyntheticVariant::ModelField93,
            SyntheticVariant::AggregateCellCountMismatch,
            SyntheticVariant::WidePreBncDisagreement,
            SyntheticVariant::UnknownModelRoot,
        ]
        .into_iter()
        .map(|variant| {
            let source = synthetic_source_package_with_variant(
                Scope::EntireTable,
                false,
                matches!(variant, SyntheticVariant::UnknownModelRoot),
                Some(variant),
            )
            .unwrap_or_else(|error| {
                panic!("synthetic {variant:?} physical-sort package failed: {error}")
            });
            let package =
                Package::from_bytes_with_options(&source, fuzz_options()).unwrap_or_else(|error| {
                    panic!("synthetic {variant:?} physical-sort package must open: {error}")
                });
            SyntheticPackage { variant, package }
        })
        .collect::<Vec<_>>()
        .into_boxed_slice()
    })
}

fn synthetic_entire_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            synthetic_source_package(Scope::EntireTable, false, false)
                .unwrap_or_else(|error| panic!("synthetic entire-sort package failed: {error}"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn synthetic_selected_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            synthetic_source_package(Scope::SelectedRows, false, false)
                .unwrap_or_else(|error| panic!("synthetic selected-sort package failed: {error}"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn synthetic_locked_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            synthetic_source_package(Scope::EntireTable, true, false)
                .unwrap_or_else(|error| panic!("synthetic locked-sort package failed: {error}"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn source_built_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            decode_hex(
                SOURCE_BUILT_PACKAGE
                    .strip_prefix(b"hex:")
                    .unwrap_or(SOURCE_BUILT_PACKAGE),
            )
            .unwrap_or_else(|| panic!("source-built physical-sort package has invalid hex"))
            .into_boxed_slice()
        })
        .as_ref()
}

fn locked_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            decode_hex(
                LOCKED_PACKAGE
                    .strip_prefix(b"hex:")
                    .unwrap_or(LOCKED_PACKAGE),
            )
            .unwrap_or_else(|| panic!("locked physical-sort package has invalid hex"))
            .into_boxed_slice()
        })
        .as_ref()
}

fn command_input(data: &[u8]) -> Vec<u8> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex_bounded(encoded).unwrap_or_default();
    }
    data.get(..data.len().min(MAX_COMMAND_BYTES))
        .unwrap_or(data)
        .to_vec()
}

fn decode_hex_bounded(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_COMMAND_BYTES.saturating_mul(2).saturating_add(32) {
        return None;
    }
    let decoded = decode_hex(encoded)?;
    (decoded.len() <= MAX_COMMAND_BYTES).then_some(decoded)
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
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

// The source-built header corpus intentionally keeps its package tiny, but it
// does not contain scalar body values. This bounded fixture builder gives the
// physical owner a small, deterministic table with text keys, duplicate-key
// ties, sparse headers, stable row UIDs, a stroke sidecar, and optional opaque
// extensions. It mirrors the canonical graph used by the focused integration
// tests while keeping the fuzz package independent of test-only modules or a
// native `.key` copy.
fn synthetic_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn synthetic_field_reference(path: impl Into<FieldPath>, references: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references.extend_from_slice(references);
    field
}

fn synthetic_object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> Result<ArchiveObject, Box<dyn StdError>> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    object.archive_info.message_infos[0]
        .object_references
        .extend_from_slice(references);
    Ok(object)
}

fn synthetic_data_list(
    list_type: tst::table_data_list::ListType,
    entries: Vec<tst::table_data_list::ListEntry>,
) -> Vec<u8> {
    tst::TableDataList {
        list_type: list_type as i32,
        next_list_id: entries
            .iter()
            .map(|entry| entry.key)
            .max()
            .unwrap_or(0)
            .saturating_add(1),
        entries,
        segments: Vec::new(),
        is_new_for_bnc: Some(true),
    }
    .encode_to_vec()
}

fn synthetic_string_entries_for(
    variant: Option<SyntheticVariant>,
) -> Vec<tst::table_data_list::ListEntry> {
    let with_sidecars = matches!(variant, Some(SyntheticVariant::StringSidecars));
    [
        (1, "Name"),
        (2, "Marker"),
        (3, "zebra"),
        (4, "last"),
        (5, "apple"),
        (6, "first apple"),
        (7, "banana"),
        (8, "middle"),
        (9, "second apple"),
        (10, "Total"),
        (11, "footer"),
    ]
    .into_iter()
    .map(|(key, string)| tst::table_data_list::ListEntry {
        key,
        refcount: 1,
        string: Some(string.to_owned()),
        import_warning_set: (with_sidecars && key == 3).then(|| tst::ImportWarningSetArchive {
            cond_format_expr: Some(true),
            original_data_format: Some("legacy text".to_owned()),
            ..tst::ImportWarningSetArchive::default()
        }),
        cell_spec: (with_sidecars && key == 3).then(|| tst::CellSpecArchive {
            interaction_type: 3,
            ..tst::CellSpecArchive::default()
        }),
        ..tst::table_data_list::ListEntry::default()
    })
    .collect()
}

/// Encode the compact BNC v5 text cell used by the native tile format.
fn synthetic_bnc_text(identifier: u32) -> Vec<u8> {
    let mut bytes = vec![5, 3, 0, 0, 0, 0, 0, 0];
    bytes.extend_from_slice(&8u32.to_le_bytes());
    bytes.extend_from_slice(&identifier.to_le_bytes());
    bytes
}

fn synthetic_encode_row_with_padding(
    cells: &[Option<Vec<u8>>],
    padding_slots: usize,
) -> (Vec<u8>, Vec<u8>) {
    let mut storage = Vec::new();
    let mut offsets =
        Vec::with_capacity(cells.len().saturating_add(padding_slots).saturating_mul(2));
    for cell in cells {
        let Some(cell) = cell else {
            offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            continue;
        };
        offsets.extend_from_slice(
            &u16::try_from(storage.len())
                .expect("synthetic physical-sort BNC rows fit narrow offsets")
                .to_le_bytes(),
        );
        storage.extend_from_slice(cell);
    }
    for _ in 0..padding_slots {
        offsets.extend_from_slice(&u16::MAX.to_le_bytes());
    }
    (storage, offsets)
}

fn synthetic_encode_pre_bnc_sentinels(cells: &[Option<Vec<u8>>]) -> (Vec<u8>, Vec<u8>) {
    let mut storage = Vec::new();
    let mut offsets = Vec::with_capacity(cells.len().saturating_mul(2));
    for cell in cells {
        if cell.is_none() {
            offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            continue;
        }
        offsets.extend_from_slice(
            &u16::try_from(storage.len())
                .expect("synthetic physical-sort pre-BNC rows fit narrow offsets")
                .to_le_bytes(),
        );
        storage.extend_from_slice(&SYNTHETIC_PRE_BNC_EMPTY_CELL);
    }
    (storage, offsets)
}

fn synthetic_row_for_variant(
    cells: [Option<Vec<u8>>; 2],
    index: u32,
    variant: Option<SyntheticVariant>,
) -> tst::TileRowInfo {
    let cell_count = cells.iter().filter(|cell| cell.is_some()).count();
    let padding_slots =
        usize::from(matches!(variant, Some(SyntheticVariant::PaddedOffsets))).saturating_mul(2);
    let (storage, offsets) = synthetic_encode_row_with_padding(&cells, padding_slots);
    let (pre_storage, pre_offsets) = if matches!(
        variant,
        Some(SyntheticVariant::PairedPreBnc | SyntheticVariant::WidePreBncDisagreement)
    ) {
        synthetic_encode_pre_bnc_sentinels(&cells)
    } else {
        (Vec::new(), Vec::new())
    };
    tst::TileRowInfo {
        tile_row_index: index,
        cell_count: u32::try_from(cell_count).expect("synthetic cell count fits u32"),
        cell_storage_buffer_pre_bnc: pre_storage,
        cell_offsets_pre_bnc: pre_offsets,
        storage_version: Some(5),
        cell_storage_buffer: Some(storage),
        cell_offsets: Some(offsets),
        // The disagreement variant deliberately marks the modern row wide
        // while retaining narrow offsets and the canonical narrow pre-BNC
        // mirror. The selected topology remains narrow, so promotion must
        // reject this source before publishing a candidate.
        has_wide_offsets: Some(matches!(
            variant,
            Some(SyntheticVariant::WidePreBncDisagreement)
        )),
    }
}

fn synthetic_tile_payload_for(
    variant: Option<SyntheticVariant>,
) -> Result<Vec<u8>, Box<dyn StdError>> {
    let num_cells = if matches!(variant, Some(SyntheticVariant::AggregateCellCountMismatch)) {
        SYNTHETIC_ROW_COUNT
            .checked_mul(SYNTHETIC_COLUMN_COUNT)
            .and_then(|count| count.checked_add(1))
            .ok_or("synthetic aggregate cell count overflow")?
    } else {
        SYNTHETIC_ROW_COUNT * SYNTHETIC_COLUMN_COUNT
    };
    let mut payload = tst::Tile {
        max_column: SYNTHETIC_COLUMN_COUNT - 1,
        max_row: SYNTHETIC_ROW_COUNT - 1,
        num_cells,
        numrows: SYNTHETIC_ROW_COUNT,
        row_infos: vec![
            synthetic_row_for_variant(
                [Some(synthetic_bnc_text(1)), Some(synthetic_bnc_text(2))],
                0,
                variant,
            ),
            synthetic_row_for_variant(
                [Some(synthetic_bnc_text(3)), Some(synthetic_bnc_text(4))],
                1,
                variant,
            ),
            synthetic_row_for_variant(
                [Some(synthetic_bnc_text(5)), Some(synthetic_bnc_text(6))],
                2,
                variant,
            ),
            synthetic_row_for_variant(
                [Some(synthetic_bnc_text(7)), Some(synthetic_bnc_text(8))],
                3,
                variant,
            ),
            synthetic_row_for_variant(
                [Some(synthetic_bnc_text(5)), Some(synthetic_bnc_text(9))],
                4,
                variant,
            ),
            synthetic_row_for_variant(
                [Some(synthetic_bnc_text(10)), Some(synthetic_bnc_text(11))],
                5,
                variant,
            ),
        ],
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        should_use_wide_rows: Some(false),
    }
    .encode_to_vec();
    if matches!(variant, Some(SyntheticVariant::UnknownMutableRoot)) {
        append_varint_field(&mut payload, 99, 0x0bad_cafe)?;
        append_length_delimited_field(&mut payload, 100, SYNTHETIC_TILE_UNKNOWN_MARKER)?;
    }
    Ok(payload)
}

fn synthetic_row_header(index: u32) -> tst::header_storage_bucket::Header {
    tst::header_storage_bucket::Header {
        index,
        size: f32::from_bits(0x41a0_0000),
        hiding_state: 0,
        number_of_cells: SYNTHETIC_COLUMN_COUNT,
        ..tst::header_storage_bucket::Header::default()
    }
}

fn synthetic_row_header_payload(
    variant: Option<SyntheticVariant>,
) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut payload = tst::HeaderStorageBucket {
        bucket_hash_function: 1,
        // Sparse by design. The owner moves present records without inventing
        // headers for rows that were intentionally absent.
        headers: vec![
            synthetic_row_header(0),
            synthetic_row_header(1),
            synthetic_row_header(3),
            synthetic_row_header(5),
        ],
    }
    .encode_to_vec();
    if matches!(variant, Some(SyntheticVariant::UnknownMutableHeaderRoot)) {
        append_varint_field(&mut payload, 99, 0x0bad_fade)?;
        append_length_delimited_field(&mut payload, 100, SYNTHETIC_HEADER_UNKNOWN_MARKER)?;
    }
    Ok(payload)
}

fn synthetic_column_header_payload() -> Vec<u8> {
    tst::HeaderStorageBucket {
        bucket_hash_function: 1,
        headers: vec![synthetic_row_header(0), synthetic_row_header(1)],
    }
    .encode_to_vec()
}

fn synthetic_uid_map_payload(variant: Option<SyntheticVariant>) -> Vec<u8> {
    let columns = (0..SYNTHETIC_COLUMN_COUNT)
        .map(|index| tsp::Uuid {
            lower: u64::from(index) + 100,
            upper: u64::from(index) + 1_000,
        })
        .collect::<Vec<_>>();
    let mut rows = (0..SYNTHETIC_ROW_COUNT)
        .map(|index| tsp::Uuid {
            lower: u64::from(index) + 200,
            upper: u64::from(index) + 2_000,
        })
        .collect::<Vec<_>>();
    if matches!(variant, Some(SyntheticVariant::UidAliases)) {
        rows[1] = rows[0].clone();
    }
    tst::ColumnRowUidMapArchive {
        sorted_column_uids: columns,
        column_index_for_uid: (0..SYNTHETIC_COLUMN_COUNT).collect(),
        column_uid_for_index: (0..SYNTHETIC_COLUMN_COUNT).collect(),
        sorted_row_uids: rows,
        row_index_for_uid: (0..SYNTHETIC_ROW_COUNT).collect(),
        row_uid_for_index: (0..SYNTHETIC_ROW_COUNT).collect(),
    }
    .encode_to_vec()
}

fn synthetic_data_store() -> tst::DataStore {
    tst::DataStore {
        row_headers: tst::HeaderStorage {
            bucket_hash_function: 1,
            buckets: vec![synthetic_reference(SYNTHETIC_ROW_HEADERS)],
        },
        column_headers: synthetic_reference(SYNTHETIC_COLUMN_HEADERS),
        tiles: tst::TileStorage {
            tiles: vec![tst::tile_storage::Tile {
                tileid: 0,
                tile: synthetic_reference(SYNTHETIC_TILE),
            }],
            tile_size: Some(256),
            should_use_wide_rows: Some(false),
        },
        string_table: synthetic_reference(SYNTHETIC_STRINGS),
        style_table: synthetic_reference(SYNTHETIC_STYLES),
        formula_table: synthetic_reference(SYNTHETIC_FORMULAS),
        format_table_pre_bnc: synthetic_reference(SYNTHETIC_FORMATS),
        next_row_strip_id: 1,
        next_column_strip_id: 1,
        row_tile_tree: tst::TableRbTree {
            nodes: vec![tst::table_rb_tree::Node { key: 0, value: 0 }],
        },
        ..tst::DataStore::default()
    }
}

fn synthetic_table_model(
    scope: Scope,
    with_unknowns: bool,
    variant: Option<SyntheticVariant>,
) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut payload = tst::TableModelArchive {
        table_id: "physical-sort-table".to_owned(),
        table_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        body_text_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        header_row_text_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        header_column_text_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        footer_row_text_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        body_cell_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        header_row_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        header_column_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        footer_row_style: synthetic_reference(SYNTHETIC_TITLE_STYLE),
        table_name_style: Some(synthetic_reference(SYNTHETIC_TITLE_STYLE)),
        table_name_shape_style: Some(synthetic_reference(SYNTHETIC_SHAPE_STYLE)),
        base_data_store: synthetic_data_store(),
        number_of_rows: SYNTHETIC_ROW_COUNT,
        number_of_columns: SYNTHETIC_COLUMN_COUNT,
        table_name: "Cities".to_owned(),
        table_name_enabled: Some(false),
        number_of_header_rows: Some(SYNTHETIC_HEADER_ROWS),
        number_of_footer_rows: Some(SYNTHETIC_FOOTER_ROWS),
        header_rows_frozen: Some(true),
        header_columns_frozen: Some(true),
        default_row_height: 20.0,
        default_column_width: 64.0,
        repeating_header_rows_enabled: Some(true),
        repeating_header_columns_enabled: Some(true),
        sort_order: Some(tst::TableSortOrderArchive {
            r#type: scope.native_value(),
            rules: vec![tst::table_sort_order_archive::SortRuleArchive {
                index: 0,
                direction: Direction::Ascending.native_value(),
            }],
        }),
        base_column_row_uids: Some(synthetic_reference(SYNTHETIC_UID_MAP)),
        stroke_sidecar: Some(synthetic_reference(SYNTHETIC_STROKE_SIDECAR)),
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    if with_unknowns {
        append_varint_field(&mut payload, 99, 0xfeed_beef)?;
        append_length_delimited_field(&mut payload, 100, b"physical-sort-unknown")?;
    }
    if let Some(field) = variant.and_then(SyntheticVariant::model_field) {
        append_length_delimited_field(&mut payload, field, SYNTHETIC_MODEL_FIELD_INVALID_PAYLOAD)?;
    }
    Ok(payload)
}

fn synthetic_table_info_payload(locked: bool) -> Vec<u8> {
    tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            parent: Some(synthetic_reference(SYNTHETIC_SLIDE)),
            locked: Some(locked),
            ..tsd::DrawableArchive::default()
        },
        table_model: synthetic_reference(SYNTHETIC_TABLE_MODEL),
        ..tst::TableInfoArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_formula_owner_dependencies_payload() -> Vec<u8> {
    tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid: tsp::Uuid {
            lower: 0x1010,
            upper: 0x2020,
        },
        internal_formula_owner_id: 1,
        owner_kind: Some(1),
        formula_owner: Some(synthetic_reference(SYNTHETIC_TABLE_INFO)),
        ..tsce::FormulaOwnerDependenciesArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_shape_info_payload() -> Vec<u8> {
    tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
            super_: tsd::DrawableArchive::default(),
            ..tsd::ShapeArchive::default()
        },
        ..tswp::ShapeInfoArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_style_payload() -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: tss_style("physical-sort"),
        ..tswp::ParagraphStyleArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_tss_style(identifier: &str) -> litchi_iwa_protos::tss::StyleArchive {
    litchi_iwa_protos::tss::StyleArchive {
        style_identifier: Some(identifier.to_owned()),
        ..litchi_iwa_protos::tss::StyleArchive::default()
    }
}

fn tss_style(identifier: &str) -> litchi_iwa_protos::tss::StyleArchive {
    synthetic_tss_style(identifier)
}

fn synthetic_model_references() -> Vec<u64> {
    vec![
        SYNTHETIC_ROW_HEADERS,
        SYNTHETIC_COLUMN_HEADERS,
        SYNTHETIC_TILE,
        SYNTHETIC_STRINGS,
        SYNTHETIC_STYLES,
        SYNTHETIC_FORMULAS,
        SYNTHETIC_FORMATS,
        SYNTHETIC_UID_MAP,
        SYNTHETIC_STROKE_SIDECAR,
        SYNTHETIC_FORMULA_OWNER_DEPENDENCIES,
        SYNTHETIC_TITLE_STYLE,
        SYNTHETIC_SHAPE_STYLE,
    ]
}

fn synthetic_model_field_infos() -> Vec<FieldInfo> {
    vec![
        synthetic_field_reference(vec![4, 1, 2], &[SYNTHETIC_ROW_HEADERS]),
        synthetic_field_reference(vec![4, 2], &[SYNTHETIC_COLUMN_HEADERS]),
        synthetic_field_reference(vec![4, 4], &[SYNTHETIC_STRINGS]),
        synthetic_field_reference(vec![4, 5], &[SYNTHETIC_STYLES]),
        synthetic_field_reference(vec![4, 6], &[SYNTHETIC_FORMULAS]),
        synthetic_field_reference(vec![4, 11], &[SYNTHETIC_FORMATS]),
        synthetic_field_reference(vec![SYNTHETIC_BASE_UID_FIELD], &[SYNTHETIC_UID_MAP]),
        synthetic_field_reference(vec![SYNTHETIC_STROKE_FIELD], &[SYNTHETIC_STROKE_SIDECAR]),
    ]
}

fn synthetic_document_payload() -> Vec<u8> {
    kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: synthetic_reference(SYNTHETIC_SHOW),
        ..kn::DocumentArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_show_payload() -> Vec<u8> {
    kn::ShowArchive {
        theme: synthetic_reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![synthetic_reference(SYNTHETIC_SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: synthetic_reference(81),
        ..kn::ShowArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_slide_payload() -> Vec<u8> {
    kn::SlideArchive {
        style: synthetic_reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: vec![
            synthetic_reference(SYNTHETIC_TABLE_INFO),
            synthetic_reference(SYNTHETIC_NON_TABLE_DRAWABLE),
        ],
        drawables_z_order: vec![
            synthetic_reference(SYNTHETIC_TABLE_INFO),
            synthetic_reference(SYNTHETIC_NON_TABLE_DRAWABLE),
        ],
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    }
    .encode_to_vec()
}

fn synthetic_metadata_payload() -> Result<Vec<u8>, Box<dyn StdError>> {
    let identifiers = [
        SYNTHETIC_DOCUMENT,
        SYNTHETIC_SHOW,
        SYNTHETIC_SLIDE_NODE,
        SYNTHETIC_SLIDE,
        SYNTHETIC_TABLE_INFO,
        SYNTHETIC_TABLE_MODEL,
        SYNTHETIC_NON_TABLE_DRAWABLE,
        SYNTHETIC_ROW_HEADERS,
        SYNTHETIC_COLUMN_HEADERS,
        SYNTHETIC_TILE,
        SYNTHETIC_STRINGS,
        SYNTHETIC_STYLES,
        SYNTHETIC_FORMULAS,
        SYNTHETIC_FORMATS,
        SYNTHETIC_UID_MAP,
        SYNTHETIC_STROKE_SIDECAR,
        SYNTHETIC_FORMULA_OWNER_DEPENDENCIES,
        SYNTHETIC_TITLE_STYLE,
        SYNTHETIC_SHAPE_STYLE,
    ];
    Ok(tsp::PackageMetadata {
        last_object_identifier: 900,
        components: vec![tsp::ComponentInfo {
            identifier: 1,
            preferred_locator: SYNTHETIC_DOCUMENT_MEMBER.to_owned(),
            locator: Some(SYNTHETIC_DOCUMENT_MEMBER.to_owned()),
            save_token: Some(1),
            object_uuid_map_entries: identifiers
                .into_iter()
                .map(|identifier| tsp::ObjectUuidMapEntry {
                    identifier,
                    uuid: tsp::Uuid {
                        lower: identifier + 10_000,
                        upper: identifier + 20_000,
                    },
                })
                .collect(),
            ..tsp::ComponentInfo::default()
        }],
        ..tsp::PackageMetadata::default()
    }
    .encode_to_vec())
}

fn synthetic_source_package(
    scope: Scope,
    locked: bool,
    with_unknowns: bool,
) -> Result<Vec<u8>, Box<dyn StdError>> {
    synthetic_source_package_with_variant(scope, locked, with_unknowns, None)
}

fn synthetic_source_package_with_variant(
    scope: Scope,
    locked: bool,
    with_unknowns: bool,
    variant: Option<SyntheticVariant>,
) -> Result<Vec<u8>, Box<dyn StdError>> {
    let document = synthetic_object(
        SYNTHETIC_DOCUMENT,
        1,
        synthetic_document_payload(),
        &[SYNTHETIC_SHOW],
    )?;
    let show = synthetic_object(
        SYNTHETIC_SHOW,
        2,
        synthetic_show_payload(),
        &[SYNTHETIC_SLIDE_NODE],
    )?;
    let node = synthetic_object(
        SYNTHETIC_SLIDE_NODE,
        4,
        kn::SlideNodeArchive {
            slide: Some(synthetic_reference(SYNTHETIC_SLIDE)),
            ..kn::SlideNodeArchive::default()
        }
        .encode_to_vec(),
        &[SYNTHETIC_SLIDE],
    )?;
    let mut slide = synthetic_object(
        SYNTHETIC_SLIDE,
        5,
        synthetic_slide_payload(),
        &[SYNTHETIC_TABLE_INFO, SYNTHETIC_NON_TABLE_DRAWABLE],
    )?;
    slide.archive_info.message_infos[0].field_infos.extend([
        synthetic_field_reference(
            vec![7],
            &[SYNTHETIC_TABLE_INFO, SYNTHETIC_NON_TABLE_DRAWABLE],
        ),
        synthetic_field_reference(
            vec![42],
            &[SYNTHETIC_TABLE_INFO, SYNTHETIC_NON_TABLE_DRAWABLE],
        ),
    ]);
    let mut info = synthetic_object(
        SYNTHETIC_TABLE_INFO,
        SYNTHETIC_TABLE_INFO_MESSAGE_TYPE,
        synthetic_table_info_payload(locked),
        &[SYNTHETIC_TABLE_MODEL],
    )?;
    info.archive_info.message_infos[0]
        .field_infos
        .push(synthetic_field_reference(vec![2], &[SYNTHETIC_TABLE_MODEL]));
    let mut model = synthetic_object(
        SYNTHETIC_TABLE_MODEL,
        SYNTHETIC_TABLE_MODEL_MESSAGE_TYPE,
        synthetic_table_model(scope, with_unknowns, variant)?,
        &synthetic_model_references(),
    )?;
    model.archive_info.message_infos[0]
        .field_infos
        .extend(synthetic_model_field_infos());

    let objects = [
        document,
        show,
        node,
        slide,
        info,
        model,
        synthetic_object(
            SYNTHETIC_FORMULA_OWNER_DEPENDENCIES,
            SYNTHETIC_FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE,
            synthetic_formula_owner_dependencies_payload(),
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_NON_TABLE_DRAWABLE,
            SYNTHETIC_SHAPE_INFO_MESSAGE_TYPE,
            synthetic_shape_info_payload(),
            &[SYNTHETIC_SLIDE],
        )?,
        synthetic_object(
            SYNTHETIC_ROW_HEADERS,
            SYNTHETIC_HEADER_BUCKET_MESSAGE_TYPE,
            synthetic_row_header_payload(variant)?,
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_COLUMN_HEADERS,
            SYNTHETIC_HEADER_BUCKET_MESSAGE_TYPE,
            synthetic_column_header_payload(),
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_TILE,
            SYNTHETIC_TILE_MESSAGE_TYPE,
            synthetic_tile_payload_for(variant)?,
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_STRINGS,
            SYNTHETIC_TABLE_DATA_LIST_MESSAGE_TYPE,
            synthetic_data_list(
                tst::table_data_list::ListType::String,
                synthetic_string_entries_for(variant),
            ),
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_STYLES,
            SYNTHETIC_TABLE_DATA_LIST_MESSAGE_TYPE,
            synthetic_data_list(tst::table_data_list::ListType::Style, Vec::new()),
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_FORMULAS,
            SYNTHETIC_TABLE_DATA_LIST_MESSAGE_TYPE,
            synthetic_data_list(tst::table_data_list::ListType::Formula, Vec::new()),
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_FORMATS,
            SYNTHETIC_TABLE_DATA_LIST_MESSAGE_TYPE,
            synthetic_data_list(tst::table_data_list::ListType::Format, Vec::new()),
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_UID_MAP,
            if matches!(variant, Some(SyntheticVariant::LegacyUidType)) {
                SYNTHETIC_COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE
            } else {
                SYNTHETIC_COLUMN_ROW_UID_MAP_MESSAGE_TYPE
            },
            synthetic_uid_map_payload(variant),
            &[],
        )?,
        synthetic_object(
            SYNTHETIC_STROKE_SIDECAR,
            SYNTHETIC_STROKE_SIDECAR_MESSAGE_TYPE,
            tst::StrokeSidecarArchive {
                row_count: Some(SYNTHETIC_ROW_COUNT),
                column_count: Some(SYNTHETIC_COLUMN_COUNT),
                ..tst::StrokeSidecarArchive::default()
            }
            .encode_to_vec(),
            &[],
        )?,
        synthetic_object(SYNTHETIC_TITLE_STYLE, 2_022, synthetic_style_payload(), &[])?,
        synthetic_object(
            SYNTHETIC_SHAPE_STYLE,
            2_025,
            tswp::ShapeStyleArchive::default().encode_to_vec(),
            &[],
        )?,
    ]
    .into_iter()
    .collect::<Vec<_>>();

    let document_bytes = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    let metadata_bytes = SnappyStream::compress(
        &Archive {
            objects: vec![synthetic_object(
                SYNTHETIC_METADATA_OBJECT,
                SYNTHETIC_METADATA_MESSAGE_TYPE,
                synthetic_metadata_payload()?,
                &[],
            )?],
        }
        .to_bytes()?,
    )?;
    Ok(package_to_bytes(
        [
            (
                SYNTHETIC_SENTINEL_MEMBER,
                b"untouched physical-sort sentinel".as_slice(),
            ),
            (SYNTHETIC_DOCUMENT_MEMBER, document_bytes.as_slice()),
            (SYNTHETIC_METADATA_MEMBER, metadata_bytes.as_slice()),
            (SYNTHETIC_PREVIEWS[0], b"physical-sort preview".as_slice()),
            (
                SYNTHETIC_PREVIEWS[1],
                b"physical-sort micro preview".as_slice(),
            ),
            (
                SYNTHETIC_PREVIEWS[2],
                b"physical-sort web preview".as_slice(),
            ),
        ],
        ArchiveLimits::default(),
    )?)
}

fn synthetic_document_archive(package: &[u8]) -> Archive {
    let catalog = Catalog::from_bytes(package)
        .unwrap_or_else(|error| panic!("synthetic physical-sort package catalog failed: {error}"));
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == SYNTHETIC_DOCUMENT_MEMBER)
        .unwrap_or_else(|| panic!("synthetic physical-sort document component is missing"));
    let stream = SnappyStream::decompress(entry.data())
        .unwrap_or_else(|error| {
            panic!("synthetic physical-sort document decompress failed: {error}")
        })
        .into_bytes();
    Archive::parse(&stream)
        .unwrap_or_else(|error| panic!("synthetic physical-sort document archive failed: {error}"))
}

fn assert_synthetic_variant_shape(package: &[u8], variant: SyntheticVariant) {
    let archive = synthetic_document_archive(package);
    let tile_object = archive
        .object(SYNTHETIC_TILE)
        .unwrap_or_else(|| panic!("synthetic {variant:?} tile object is missing"));
    let tile_message = tile_object
        .messages
        .iter()
        .find(|message| message.type_ == SYNTHETIC_TILE_MESSAGE_TYPE)
        .unwrap_or_else(|| panic!("synthetic {variant:?} tile message is missing"));
    let tile = tst::Tile::decode(tile_message.data.as_slice())
        .unwrap_or_else(|error| panic!("synthetic {variant:?} tile decode failed: {error}"));
    let column_count = usize::try_from(SYNTHETIC_COLUMN_COUNT)
        .unwrap_or_else(|error| panic!("synthetic column count conversion failed: {error}"));
    match variant {
        SyntheticVariant::PaddedOffsets => {
            for row in &tile.row_infos {
                let offsets = row
                    .cell_offsets
                    .as_deref()
                    .unwrap_or_else(|| panic!("synthetic padded row has no BNC offsets"));
                assert!(
                    offsets.len() >= column_count.saturating_add(2).saturating_mul(2),
                    "synthetic padded row lost its offset slots"
                );
                assert!(
                    offsets
                        .chunks_exact(2)
                        .skip(column_count)
                        .all(|bytes| { u16::from_le_bytes([bytes[0], bytes[1]]) == u16::MAX }),
                    "synthetic padded row has a non-sentinel offset tail"
                );
            }
        },
        SyntheticVariant::PairedPreBnc => {
            for row in &tile.row_infos {
                assert!(
                    !row.cell_storage_buffer_pre_bnc.is_empty()
                        && !row.cell_offsets_pre_bnc.is_empty(),
                    "synthetic paired row lost its pre-BNC buffers"
                );
                assert_eq!(
                    row.cell_storage_buffer_pre_bnc.len(),
                    row.cell_offsets_pre_bnc
                        .chunks_exact(2)
                        .filter(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) != u16::MAX)
                        .count()
                        .saturating_mul(SYNTHETIC_PRE_BNC_EMPTY_CELL.len()),
                    "synthetic paired row changed its pre-BNC storage width"
                );
                assert_eq!(
                    row.cell_offsets_pre_bnc.len(),
                    column_count.saturating_mul(2),
                    "synthetic paired row changed its pre-BNC offset width"
                );
                assert_eq!(
                    row.cell_offsets_pre_bnc, SYNTHETIC_PRE_BNC_OFFSETS,
                    "synthetic paired row changed its pre-BNC offsets"
                );
                assert!(
                    row.cell_storage_buffer_pre_bnc
                        .chunks_exact(SYNTHETIC_PRE_BNC_EMPTY_CELL.len())
                        .all(|cell| cell == SYNTHETIC_PRE_BNC_EMPTY_CELL),
                    "synthetic paired row changed its pre-BNC sentinel"
                );
            }
        },
        SyntheticVariant::StringSidecars => {
            let strings_object = archive
                .object(SYNTHETIC_STRINGS)
                .unwrap_or_else(|| panic!("synthetic string-sidecar object is missing"));
            let strings_message = strings_object
                .messages
                .iter()
                .find(|message| message.type_ == SYNTHETIC_TABLE_DATA_LIST_MESSAGE_TYPE)
                .unwrap_or_else(|| panic!("synthetic string-sidecar message is missing"));
            let strings = tst::TableDataList::decode(strings_message.data.as_slice())
                .unwrap_or_else(|error| {
                    panic!("synthetic string-sidecar list decode failed: {error}")
                });
            let entry = strings
                .entries
                .iter()
                .find(|entry| entry.key == 3)
                .unwrap_or_else(|| panic!("synthetic string-sidecar entry is missing"));
            assert!(
                entry.import_warning_set.is_some() && entry.cell_spec.is_some(),
                "synthetic string sidecars were not retained"
            );
        },
        SyntheticVariant::UidAliases => {
            let uid_object = archive
                .object(SYNTHETIC_UID_MAP)
                .unwrap_or_else(|| panic!("synthetic UID alias object is missing"));
            let uid_message = uid_object
                .messages
                .iter()
                .find(|message| message.type_ == SYNTHETIC_COLUMN_ROW_UID_MAP_MESSAGE_TYPE)
                .unwrap_or_else(|| panic!("synthetic UID alias message is missing"));
            let uid_map = tst::ColumnRowUidMapArchive::decode(uid_message.data.as_slice())
                .unwrap_or_else(|error| panic!("synthetic UID alias map decode failed: {error}"));
            assert_eq!(
                uid_map.sorted_row_uids.get(0),
                uid_map.sorted_row_uids.get(1),
                "synthetic UID alias was not retained"
            );
        },
        SyntheticVariant::LegacyUidType => {
            let uid_object = archive
                .object(SYNTHETIC_UID_MAP)
                .unwrap_or_else(|| panic!("synthetic legacy UID object is missing"));
            let uid_messages = uid_object
                .messages
                .iter()
                .filter(|message| message.type_ == SYNTHETIC_COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE)
                .count();
            assert_eq!(
                uid_messages, 1,
                "synthetic legacy UID type was not retained"
            );
        },
        SyntheticVariant::UnknownMutableRoot => {
            assert!(
                tile_message
                    .data
                    .windows(SYNTHETIC_TILE_UNKNOWN_MARKER.len())
                    .any(|window| window == SYNTHETIC_TILE_UNKNOWN_MARKER),
                "synthetic mutable tile-root unknown field was not retained"
            );
        },
        SyntheticVariant::UnknownMutableHeaderRoot => {
            let headers_object = archive
                .object(SYNTHETIC_ROW_HEADERS)
                .unwrap_or_else(|| panic!("synthetic {variant:?} header object is missing"));
            let headers_message = headers_object
                .messages
                .iter()
                .find(|message| message.type_ == SYNTHETIC_HEADER_BUCKET_MESSAGE_TYPE)
                .unwrap_or_else(|| panic!("synthetic {variant:?} header message is missing"));
            assert!(
                headers_message
                    .data
                    .windows(SYNTHETIC_HEADER_UNKNOWN_MARKER.len())
                    .any(|window| window == SYNTHETIC_HEADER_UNKNOWN_MARKER),
                "synthetic mutable header-root unknown field was not retained"
            );
        },
        SyntheticVariant::ModelField39
        | SyntheticVariant::ModelField45
        | SyntheticVariant::ModelField84
        | SyntheticVariant::ModelField93 => {
            let model_field = variant
                .model_field()
                .unwrap_or_else(|| panic!("synthetic {variant:?} model field is missing"));
            let model_object = archive
                .object(SYNTHETIC_TABLE_MODEL)
                .unwrap_or_else(|| panic!("synthetic {variant:?} model object is missing"));
            let model_message = model_object
                .messages
                .iter()
                .find(|message| message.type_ == SYNTHETIC_TABLE_MODEL_MESSAGE_TYPE)
                .unwrap_or_else(|| panic!("synthetic {variant:?} model message is missing"));
            let fields = WireView::parse(model_message.data.as_slice())
                .unwrap_or_else(|error| panic!("synthetic {variant:?} model wire failed: {error}"));
            assert!(
                fields.fields().any(|field| {
                    field.number() == model_field
                        && field.wire_type() == 2
                        && field.payload() == SYNTHETIC_MODEL_FIELD_INVALID_PAYLOAD
                }),
                "synthetic {variant:?} malformed model field was not retained"
            );
        },
        SyntheticVariant::AggregateCellCountMismatch => {
            assert_ne!(
                tile.num_cells,
                SYNTHETIC_ROW_COUNT * SYNTHETIC_COLUMN_COUNT,
                "synthetic aggregate cell count unexpectedly matches row storage"
            );
        },
        SyntheticVariant::WidePreBncDisagreement => {
            assert_eq!(
                tile.should_use_wide_rows,
                Some(false),
                "synthetic wide/pre-BNC topology unexpectedly changed tile width"
            );
            assert!(
                tile.row_infos.iter().all(|row| {
                    row.has_wide_offsets == Some(true)
                        && row.cell_offsets_pre_bnc == SYNTHETIC_PRE_BNC_OFFSETS
                }),
                "synthetic wide/pre-BNC disagreement was not retained"
            );
        },
        SyntheticVariant::UnknownModelRoot => {
            let model_object = archive
                .object(SYNTHETIC_TABLE_MODEL)
                .unwrap_or_else(|| panic!("synthetic model object is missing"));
            let model_message = model_object
                .messages
                .iter()
                .find(|message| message.type_ == SYNTHETIC_TABLE_MODEL_MESSAGE_TYPE)
                .unwrap_or_else(|| panic!("synthetic model message is missing"));
            assert!(
                model_message
                    .data
                    .windows(b"physical-sort-unknown".len())
                    .any(|window| window == b"physical-sort-unknown"),
                "synthetic model-root unknown field was not retained"
            );
        },
    }
}

fn exercise_synthetic_variant(package: &Package, variant: SyntheticVariant, command: &[u8]) {
    let source = package_bytes(package);
    assert_synthetic_variant_shape(&source, variant);

    // Use a fixed descending order so each valid topology reaches a changed
    // physical transaction regardless of the command bytes. The regular
    // command-driven pass above still covers arbitrary order/range choices.
    let order = Order::with_scope(
        Scope::EntireTable,
        vec![Rule::new(
            ColumnIndex::new(0)
                .unwrap_or_else(|error| unreachable!("synthetic sort column is valid: {error}")),
            Direction::Descending,
        )],
    )
    .unwrap_or_else(|error| unreachable!("synthetic descending order is valid: {error}"));
    let configured = match package
        .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
        .and_then(|edit| edit.set(order).commit())
    {
        Ok(commit) => commit.package().clone(),
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            assert!(
                variant.expects_rejection(),
                "accepted synthetic {variant:?} topology failed during order staging"
            );
            return;
        },
    };
    let configured_source = package_bytes(&configured);
    assert_source_unchanged(&configured, &configured_source);
    let result =
        configured.execute_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0));
    match result {
        Ok(commit) => {
            assert!(
                !variant.expects_rejection(),
                "synthetic {variant:?} topology unexpectedly published a physical sort"
            );
            assert_synthetic_variant_shape(&package_bytes(commit.package()), variant);
            verify_commit(
                &configured,
                &commit,
                SlideSelector::index(0),
                TableSelector::index(0),
                command_range(command),
                command,
                &configured_source,
            );
        },
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(&configured, &configured_source);
            assert!(
                variant.expects_rejection(),
                "accepted synthetic {variant:?} topology was rejected by physical sort"
            );
        },
    }
}

fn exercise_rejected_synthetic_variant(
    package: &Package,
    variant: SyntheticVariant,
    command: &[u8],
) {
    let source = package_bytes(package);
    assert_synthetic_variant_shape(&source, variant);
    let order = order_from_bytes(command);
    let configured = match package
        .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
        .and_then(|edit| edit.set(order).commit())
    {
        Ok(commit) => commit.package().clone(),
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    let configured_source = package_bytes(&configured);
    assert_source_unchanged(&configured, &configured_source);

    // Exercise both public execution routes with the command-derived order.
    // Unsupported model ownership, UID aliases, populated string sidecars,
    // unknown mutable roots, malformed model-owner fields, aggregate count
    // drift, and width/pre-BNC disagreement must be rejected before a
    // candidate can be published, regardless of whether the caller requests
    // a full table or a body-relative row range.
    for selected_rows in [false, true] {
        let result = if selected_rows {
            configured.execute_slide_table_sort_order_to_rows(
                SlideSelector::index(0),
                TableSelector::index(0),
                command_range(command),
            )
        } else {
            configured
                .execute_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
        };
        match result {
            Ok(_) => {
                panic!("synthetic {variant:?} topology unexpectedly published a physical sort")
            },
            Err(error) => {
                observe_error(error);
                assert_source_unchanged(&configured, &configured_source);
            },
        }
    }
    assert_source_unchanged(package, &source);
}

fn exercise_untrusted_package(package: &Package, command: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(usize::from(read_u16(command, 0)));
    let table = TableSelector::index(usize::from(read_u16(command, 2)));

    // Both executor variants must fail closed on arbitrary selections. A
    // valid package may still contain a table, so successful commits are
    // checked by the same immutable transaction verifier below.
    let range = command_range(command);
    observe_physical_result(package, slide, table, range, command, false, &source);
    observe_physical_result(package, slide, table, range, command, true, &source);
    exercise_selector_failures(package, &source);
    assert_source_unchanged(package, &source);
}

fn exercise_package(package: &Package, command: &[u8], locked: bool) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(0);
    let table = TableSelector::index(0);

    // Admission itself is part of the physical owner contract. A missing
    // configured order, unsupported topology, or an empty source-built body
    // is an ordinary typed failure, never a reason to mutate the source.
    let order = order_from_bytes(command);
    if locked {
        observe_physical_result(
            package,
            slide,
            table,
            command_range(command),
            command,
            false,
            &source,
        );
        observe_physical_result(
            package,
            slide,
            table,
            command_range(command),
            command,
            true,
            &source,
        );
        exercise_locked_edit(package, slide, table, order, &source);
        exercise_selector_failures(package, &source);
        assert_source_unchanged(package, &source);
        return;
    }

    // Physical execution consumes the persisted semantic order. Stage it
    // through the existing selector-first persisted owner so this target does
    // not manufacture low-level field-44 bytes or native IDs.
    let configured = match package
        .edit_slide_table_sort_order(slide, table)
        .and_then(|edit| edit.set(order.clone()).commit())
    {
        Ok(commit) => commit.package().clone(),
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            exercise_selector_failures(package, &source);
            return;
        },
    };
    let configured_source = package_bytes(&configured);
    assert_eq!(
        configured.slide_table_sort_order(slide, table),
        Ok(Some(order.clone())),
    );
    assert_source_unchanged(&configured, &configured_source);

    // Exercise the matching executor and deliberately exercise the scope
    // mismatch path as well. The owner is expected to reject a mismatch before
    // any candidate publication.
    let range = command_range(command);
    match order.scope() {
        Scope::EntireTable => {
            observe_physical_result(
                &configured,
                slide,
                table,
                range,
                command,
                false,
                &configured_source,
            );
            observe_physical_result(
                &configured,
                slide,
                table,
                range,
                command,
                true,
                &configured_source,
            );
        },
        Scope::SelectedRows => {
            observe_physical_result(
                &configured,
                slide,
                table,
                range,
                command,
                true,
                &configured_source,
            );
            observe_physical_result(
                &configured,
                slide,
                table,
                range,
                command,
                false,
                &configured_source,
            );
        },
    }
    exercise_selector_failures(&configured, &configured_source);
    assert_source_unchanged(&configured, &configured_source);
}

fn observe_physical_result(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    range: RowRange,
    command: &[u8],
    selected_rows: bool,
    source: &[u8],
) {
    let result = if selected_rows {
        package.execute_slide_table_sort_order_to_rows(slide, table, range)
    } else {
        package.execute_slide_table_sort_order(slide, table)
    };
    match result {
        Ok(commit) => verify_commit(package, &commit, slide, table, range, command, source),
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source);
        },
    }
}

fn verify_commit(
    source_package: &Package,
    commit: &SlideTablePhysicalSortCommit,
    slide: SlideSelector<'_>,
    table: TableSelector,
    range: RowRange,
    command: &[u8],
    source: &[u8],
) {
    let patch = commit.patch().clone();
    let candidate = commit.package();
    let candidate_bytes = package_bytes(candidate);
    let is_noop = patch.is_noop();

    // The immutable owner must not publish into its source, and no-op commits
    // must remain byte exact. A changed physical sort may still have an
    // already sorted body, so the patch's own no-op bit is authoritative.
    assert_eq!(commit.diagnostics().changed(), !is_noop);
    if is_noop {
        assert_eq!(candidate_bytes, source);
    }
    assert_source_unchanged(source_package, source);

    // Candidate reopen is a strict wire/graph check, not merely an in-memory
    // assertion. A malformed candidate must be impossible to publish.
    let reopened = Package::from_bytes_with_options(&candidate_bytes, fuzz_options())
        .unwrap_or_else(|error| panic!("physical-sort candidate reopen failed: {error}"));
    assert_source_unchanged(&reopened, &candidate_bytes);
    black_box(
        reopened
            .slide_table_sort_order(slide, table)
            .unwrap_or_else(|error| panic!("physical-sort candidate order read failed: {error}")),
    );
    exercise_noop_replay(candidate, slide, table, range, &candidate_bytes);

    // Apply the exact-source patch and compare bytes. Applying a changed
    // patch twice, or applying its inverse to the old source, must conflict;
    // applying the inverse to the fresh target must restore the exact source.
    let applied = source_package
        .apply_slide_table_physical_sort(&patch)
        .unwrap_or_else(|error| panic!("fresh physical-sort patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), candidate_bytes);
    assert_source_unchanged(source_package, source);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if is_noop {
        let reapplied = candidate
            .apply_slide_table_physical_sort(&patch)
            .unwrap_or_else(|error| panic!("fresh no-op physical-sort patch must apply: {error}"));
        assert_eq!(package_bytes(reapplied.package()), candidate_bytes);
    } else {
        match candidate.apply_slide_table_physical_sort(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("changed physical-sort patch unexpectedly applied twice"),
        }
        match source_package.apply_slide_table_physical_sort(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("physical-sort inverse unexpectedly applied to its source"),
        }
    }
    let restored = applied
        .package()
        .apply_slide_table_physical_sort(&inverse)
        .unwrap_or_else(|error| panic!("physical-sort inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_source_unchanged(source_package, source);

    // Keep the command-derived range live in the call-site and make sure
    // malformed/oversized command bytes cannot influence package ownership.
    black_box((range.start(), range.end(), command.len()));
}

fn exercise_noop_replay(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    range: RowRange,
    source: &[u8],
) {
    // A successful execution leaves the configured order in place. Replaying
    // either public executor therefore reaches the exact no-op path when its
    // scope matches, while the other route is an intentionally observed
    // scope mismatch. No candidate from this probe may replace the source.
    let full = package.execute_slide_table_sort_order(slide, table);
    observe_noop_or_error(full, package, source);
    let selected = package.execute_slide_table_sort_order_to_rows(slide, table, range);
    observe_noop_or_error(selected, package, source);
}

fn observe_noop_or_error(
    result: Result<SlideTablePhysicalSortCommit, impl Debug + Display>,
    source_package: &Package,
    source: &[u8],
) {
    match result {
        Ok(commit) => {
            assert!(commit.patch().is_noop());
            assert!(!commit.diagnostics().changed());
            assert_eq!(package_bytes(commit.package()), source);
            assert_source_unchanged(source_package, source);
        },
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(source_package, source);
        },
    }
}

fn exercise_locked_edit(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    order: Order,
    source: &[u8],
) {
    let edit = match package.edit_slide_table_sort_order(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source);
            return;
        },
    };
    match edit.set(order).commit() {
        Ok(commit) => {
            // A locked table may accept an exact no-op, but it must never
            // publish a changed persisted order that could bypass physical
            // lock admission.
            if !commit.patch().is_noop() {
                panic!("locked Keynote table accepted a changed sort order");
            }
            black_box(commit.diagnostics());
        },
        Err(error) => observe_error(error),
    }
    assert_source_unchanged(package, source);
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    let range = RowRange::new(0, 1)
        .unwrap_or_else(|error| unreachable!("one-row physical-sort range is valid: {error}"));
    observe_physical_result(
        package,
        SlideSelector::index(usize::MAX),
        TableSelector::index(0),
        range,
        &[],
        false,
        source,
    );
    observe_physical_result(
        package,
        SlideSelector::index(0),
        TableSelector::index(usize::MAX),
        range,
        &[],
        true,
        source,
    );
    match package.execute_slide_table_sort_order(
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        TableSelector::index(0),
    ) {
        Ok(commit) => {
            let _ = black_box(commit);
        },
        Err(error) => observe_redacted_error(error, PRIVATE_SLIDE_NAME),
    }
    assert_source_unchanged(package, source);
    observe_physical_result(
        package,
        SlideSelector::name(""),
        TableSelector::index(0),
        range,
        &[],
        true,
        source,
    );
    let oversized = RowRange::new(usize::MAX - 1, usize::MAX)
        .unwrap_or_else(|error| unreachable!("maximum physical-sort range is valid: {error}"));
    observe_physical_result(
        package,
        SlideSelector::index(0),
        TableSelector::index(0),
        oversized,
        &[],
        true,
        source,
    );
    if let Err(error) = RowRange::new(4, 2) {
        observe_error(error);
    }
    assert_source_unchanged(package, source);
}

fn order_from_bytes(data: &[u8]) -> Order {
    let scope = if control(data, 1) & 1 == 0 {
        Scope::EntireTable
    } else {
        Scope::SelectedRows
    };
    let count = usize::from(control(data, 0) % 3) + 1;
    let mut rules = Vec::with_capacity(count);
    let mut used = [false; 4];
    for priority in 0..count {
        let raw_column = usize::from(control(data, priority + 2) % 4);
        let mut column = raw_column;
        while used[column] {
            column = (column + 1) % used.len();
        }
        used[column] = true;
        let column = ColumnIndex::new(column).unwrap_or_else(|error| {
            unreachable!("generated physical-sort column is valid: {error}")
        });
        let direction = if control(data, priority + count + 2) & 1 == 0 {
            Direction::Ascending
        } else {
            Direction::Descending
        };
        rules.push(Rule::new(column, direction));
    }
    Order::with_scope(scope, rules)
        .unwrap_or_else(|error| unreachable!("generated physical-sort order is valid: {error}"))
}

fn command_range(data: &[u8]) -> RowRange {
    // Small ranges reach both the valid one-row no-op and the multi-row
    // permutation path; the owner remains responsible for checking body
    // bounds against the admitted table dimensions.
    let start = usize::from(control(data, 10) % 3);
    let len = usize::from(control(data, 11) % 4) + 1;
    RowRange::new(start, start + len)
        .unwrap_or_else(|error| unreachable!("generated physical-sort range is valid: {error}"))
}

fn exercise_redacted_ingress() {
    match Package::from_bytes_with_options(PRIVATE_INPUT, fuzz_options()) {
        Err(error) => {
            let display = error.to_string();
            let debug = format!("{error:?}");
            let private = std::str::from_utf8(PRIVATE_INPUT).unwrap_or("physical-sort-input");
            assert!(!display.contains(private));
            assert!(!debug.contains(private));
            black_box((display, debug));
        },
        Ok(_) => panic!("private malformed physical-sort input unexpectedly parsed"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(oversized, fuzz_options()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized physical-sort input must be rejected"),
    }
}

fn exercise_archive_limit() {
    let archive = Limits::new(
        MAX_INPUT_BYTES,
        1,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_IWA_STREAM_BYTES,
    )
    .unwrap_or_else(|error| unreachable!("valid physical-sort entry limits: {error}"));
    let options = ReadOptions::new(archive, fuzz_options().semantic());
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("one-entry physical-sort archive limit accepted the package"),
    }
}

fn exercise_semantic_limit(command: &[u8]) {
    let max_objects = if control(command, 0) & 1 == 0 {
        1
    } else {
        MAX_OBJECTS
    };
    let semantic = match SemanticLimits::new(
        max_objects,
        1,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    ) {
        Ok(semantic) => semantic,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let options = ReadOptions::new(fuzz_options().archive(), semantic);
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Ok(package) => {
            let source = package_bytes(&package);
            let range = command_range(command);
            observe_physical_result(
                &package,
                SlideSelector::index(0),
                TableSelector::index(0),
                range,
                command,
                false,
                &source,
            );
            assert_source_unchanged(&package, &source);
        },
        Err(error) => observe_error(error),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing physical-sort package must succeed: {error}"));
    bytes
}

fn assert_source_unchanged(package: &Package, source: &[u8]) {
    assert_eq!(package_bytes(package), source);
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from(control(data, offset)) | (u16::from(control(data, offset + 1)) << 8)
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted_error(error: impl Debug + Display, private: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
