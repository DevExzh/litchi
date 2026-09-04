//! Deterministic native Numbers data-format fixtures.
//!
//! The fixture in this module is intentionally small, but it models the
//! physical graph used by a real Numbers table: a `TableModelArchive` points
//! at a tile and a format-list sidecar, and each BNC cell points at one format
//! entry by key.  Keeping the graph in a test-only module lets integration
//! tests exercise ownership, copy-on-write, and hostile-input paths without
//! publishing native identifiers through the Numbers API.

#![allow(dead_code)]

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    varint::encode_varint,
    wire::{
        WireView, append_length_delimited_field, append_varint_field, patch_nested_varint_field,
    },
};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{tn, tsd, tsk, tsp, tst};
use litchi_numbers_wire::{BncCell, CellDataFormatKind};
use prost::Message as _;

/// A fixture operation result that keeps helper errors out of production
/// error vocabularies.
pub(crate) type FixtureResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

/// The package member containing the rooted Numbers document and sheet.
pub(crate) const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
/// The package member containing the table model, tile, and data lists.
pub(crate) const TABLES_MEMBER: &str = "Index/Tables.iwa";
/// A member that is never selected by the format transaction.
pub(crate) const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
/// The package metadata member.
pub(crate) const METADATA_MEMBER: &str = "Index/Metadata.iwa";
/// A current component with no table ownership, used by metadata routing.
pub(crate) const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
/// A deterministic arbitrary data member used by locality assertions.
pub(crate) const SENTINEL_MEMBER: &str = "Data/data-format-sentinel.bin";
/// The three canonical Numbers previews.
pub(crate) const PREVIEW_MEMBERS: [&str; 3] =
    ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// Rooted object identifiers used by [`synthetic_package`].
pub(crate) const DOCUMENT_ID: u64 = 1;
pub(crate) const SHEET_ID: u64 = 2;
pub(crate) const TABLE_INFO_ID: u64 = 3;
pub(crate) const TABLE_MODEL_ID: u64 = 4;
pub(crate) const SIDECAR_ID: u64 = 5;
pub(crate) const TILE_ID: u64 = 6;
pub(crate) const UNRELATED_OBJECT_ID: u64 = 700;
pub(crate) const METADATA_OBJECT_ID: u64 = 900;
pub(crate) const VIEW_STATE_OBJECT_ID: u64 = 800;
pub(crate) const ALIASED_SIDECAR_ID: u64 = 7;

/// Native message types used by the synthetic graph.
pub(crate) const TABLE_INFO_TYPE: u32 = 6_000;
pub(crate) const TABLE_MODEL_TYPE: u32 = 6_001;
pub(crate) const TILE_TYPE: u32 = 6_002;
pub(crate) const TABLE_DATA_LIST_TYPE: u32 = 6_005;
pub(crate) const METADATA_TYPE: u32 = 11_006;

/// Format-list keys referenced by the two fixture cells.
pub(crate) const FIRST_FORMAT_KEY: u32 = 1;
pub(crate) const SECOND_FORMAT_KEY: u32 = 2;

const BNC_CELL_FORMAT_KIND_FLAG: u32 = 0x0000_1000;
const BNC_CELL_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_2000;
const BNC_TEXT_FORMAT_IDENTIFIER_FLAG: u32 = 0x0002_0000;
const BNC_RESERVED_KNOWN_FIELD_FLAG: u32 = 0x0010_0000;

/// Zero-based cell coordinates in the one-row fixture table.
pub(crate) const FIRST_CELL: (usize, usize) = (0, 0);
pub(crate) const SECOND_CELL: (usize, usize) = (0, 1);

/// Native format families represented by the deterministic fixture.
///
/// Number, Percentage, Scientific, and Fraction share the BNC decimal-cell
/// kind.
/// Currency uses the alternate-number BNC kind and can additionally carry a
/// secondary generic format-list reference.  Their native family
/// discriminator lives in the format-list payload, so keeping that
/// discriminator explicit in the fixture prevents tests from accidentally
/// treating one family as another merely because the cell wire shape is
/// otherwise similar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum FormatFamily {
    Number,
    Currency,
    Percentage,
    Scientific,
    Fraction,
    DateTime,
    Text,
}

impl FormatFamily {
    /// Native `FormatStructArchive.format_type` for this family.
    pub(crate) const fn native_type(self) -> u32 {
        match self {
            Self::Number => NATIVE_NUMBER_FORMAT_TYPE,
            Self::Currency => NATIVE_CURRENCY_FORMAT_TYPE,
            Self::Percentage => NATIVE_PERCENTAGE_FORMAT_TYPE,
            Self::Scientific => NATIVE_SCIENTIFIC_FORMAT_TYPE,
            Self::Fraction => NATIVE_FRACTION_FORMAT_TYPE,
            Self::DateTime => NATIVE_DATE_TIME_FORMAT_TYPE,
            Self::Text => NATIVE_TEXT_FORMAT_TYPE,
        }
    }
}

/// Native Number format-list discriminator.
pub(crate) const NATIVE_NUMBER_FORMAT_TYPE: u32 = 256;
/// Native Currency format-list discriminator.
pub(crate) const NATIVE_CURRENCY_FORMAT_TYPE: u32 = 257;
/// Native Percentage format-list discriminator.
pub(crate) const NATIVE_PERCENTAGE_FORMAT_TYPE: u32 = 258;
/// Native Scientific format-list discriminator.
pub(crate) const NATIVE_SCIENTIFIC_FORMAT_TYPE: u32 = 259;
/// Native Fraction format-list discriminator.
pub(crate) const NATIVE_FRACTION_FORMAT_TYPE: u32 = 262;
/// Native Date & Time format-list discriminator.
pub(crate) const NATIVE_DATE_TIME_FORMAT_TYPE: u32 = 261;
/// Native Text format-list discriminator.
pub(crate) const NATIVE_TEXT_FORMAT_TYPE: u32 = 260;

/// Whether the two cells initially share their format-list entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FormatSharing {
    /// Both cells reference key `FIRST_FORMAT_KEY`; its refcount is two.
    Shared,
    /// The cells reference different keys with different semantic formats.
    Unshared,
}

/// Controlled malformed variants for atomic-refusal and alias tests.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Corruption {
    /// Add a second format-list entry with the same key.
    DuplicateFormatKey,
    /// Remove the format entry referenced by the first cell.
    MissingFormatEntry,
    /// Make the list refcount disagree with the two BNC references.
    RefcountMismatch,
    /// Add a second format list message to the same sidecar object.
    DuplicateFormatList,
    /// Make `format_table` and `format_table_pre_bnc` point at different
    /// physical sidecars containing otherwise identical lists.
    AliasedFormatList,
    /// Replace the selected format payload with an unsupported native type.
    UnsupportedFormatType,
    /// Replace the selected format payload with truncated protobuf bytes.
    MalformedFormatPayload,
    /// Replace the selected BNC format key with the wrong native identifier.
    WrongCellFormatKey,
    /// Add the native Fraction replacement marker with canonical false.
    FractionReplacementMetadata,
    /// Add the native Fraction replacement marker with true.
    FractionReplacementMetadataTrue,
    /// Add an unterminated unknown group to the selected format payload.
    UnterminatedUnknownGroup,
    /// Give an unrelated message a typed field-level owner edge to the tile.
    UnexpectedFieldReference,
    /// Mark the selected Text cell as converted (`0x81`) without a generic
    /// Number-format field.
    ConvertedGenericMissing,
    /// Point the converted Text cell's generic field at the primary Text
    /// entry rather than a Number entry.
    ConvertedGenericWrong,
    /// Set the converted Text cell's generic Number-format field to zero.
    ConvertedGenericZero,
    /// Point the converted Text cell's generic field at an absent list key.
    ConvertedGenericMissingEntry,
    /// Keep the converted generic key but give its payload a non-Number type.
    ConvertedGenericWrongType,
    /// Keep the converted generic key but make its list refcount disagree with
    /// the BNC secondary edge.
    ConvertedGenericRefcountMismatch,
    /// Leave a sibling Text identifier without the kind that gives it meaning.
    SiblingOrphanTextIdentifier,
    /// Leave a sibling Text identifier paired with a decimal kind.
    SiblingMismatchedTextIdentifier,
    /// Leave a reserved known BNC field on a nonselected sibling cell.
    SiblingReservedKnownField,
    /// Give a nonselected converted Text cell a generic non-Number target.
    NonselectedConvertedGenericWrongType,
    /// Add a zero-cell row whose storage buffers are nevertheless nonempty.
    ZeroCellRowStorage,
}

/// Build the canonical two-cell package used by data-format tests.
///
/// The result is deterministic: object order, member order, unknown fields,
/// extensions, preview bytes, and the unrelated member are all fixed.  `Shared`
/// gives both cells a Number format with key one and refcount two;
/// [`synthetic_package_for`] selects the corresponding Currency, Percentage,
/// Scientific, Fraction, or Text family.  `Unshared` gives the second cell a
/// Percentage format with key two and two refcount-one entries, except for
/// Text where both entries remain native type-260 Text records.
pub(crate) fn synthetic_package(sharing: FormatSharing) -> FixtureResult<Vec<u8>> {
    synthetic_package_for(FormatFamily::Number, sharing)
}

/// Build a canonical package for one selected decimal format family.
///
/// `Shared` makes both cells use `family`.  `Unshared` intentionally retains
/// the historical mixed fixture: the first cell is Number and the second is
/// Percentage.  That shape is useful for wrong-family and cross-entry COW
/// tests while preserving every existing Number test's bytes and semantics.
pub(crate) fn synthetic_package_for(
    family: FormatFamily,
    sharing: FormatSharing,
) -> FixtureResult<Vec<u8>> {
    let document = document_object()?;
    let sheet = sheet_object()?;
    let table_info = table_info_object()?;
    let model = table_model_object()?;
    let tile = tile_object(family, sharing)?;
    let sidecars = sidecar_object(family, sharing)?;

    let document_member = compressed(vec![document, sheet])?;
    let tables_member = compressed(vec![table_info, model, tile, sidecars])?;
    let unrelated_member = compressed(vec![unrelated_object()?])?;
    let metadata_member = compressed(vec![metadata_object()?])?;
    let view_state_member = compressed(vec![view_state_object()?])?;

    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (DOCUMENT_MEMBER, document_member.as_slice()),
            (TABLES_MEMBER, tables_member.as_slice()),
            (UNRELATED_MEMBER, unrelated_member.as_slice()),
            (METADATA_MEMBER, metadata_member.as_slice()),
            (VIEW_STATE_MEMBER, view_state_member.as_slice()),
            (
                PREVIEW_MEMBERS[0],
                b"numbers data-format preview".as_slice(),
            ),
            (
                PREVIEW_MEMBERS[1],
                b"numbers data-format micro-preview".as_slice(),
            ),
            (
                PREVIEW_MEMBERS[2],
                b"numbers data-format web-preview".as_slice(),
            ),
            (
                SENTINEL_MEMBER,
                b"unrelated data-format sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

/// Build a Currency fixture whose selected cell carries both native
/// Currency metadata and the optional generic secondary format identifier.
///
/// The secondary identifier is deliberately a Number entry.  The sibling
/// remains on the original Currency entry, which makes it possible to prove
/// that a Currency transition decrements/culls both selected references while
/// leaving an unrelated live entry untouched.
pub(crate) fn currency_secondary_package() -> FixtureResult<Vec<u8>> {
    let source = synthetic_package_for(FormatFamily::Currency, FormatSharing::Shared)?;
    let source = rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        *first = currency_secondary_cell()?;
        Ok(())
    })?;
    rewrite_format_list(&source, |list| {
        let original = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("format fixture Currency entry is missing"))?;
        original.refcount = 1;
        list.entries.push(format_entry(
            SECOND_FORMAT_KEY,
            1,
            NATIVE_NUMBER_FORMAT_TYPE,
            2,
            2,
            true,
        ));
        list.entries
            .push(currency_format_entry(4, 1, "USD", 2, 2, true, false));
        Ok(())
    })
}

/// Rewrite only the optional generic secondary identifier on the selected
/// Currency cell.  The BNC field layout is private to the wire crate, so keep
/// this deliberately narrow raw fixture primitive beside the fixture that
/// owns the native bytes.
pub(crate) fn rewrite_currency_secondary_identifier(
    source: &[u8],
    identifier: u32,
) -> FixtureResult<Vec<u8>> {
    rewrite_tile_cells(source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let parsed = BncCell::parse(first)?;
        if parsed.secondary_format_identifier().is_none() {
            return Err(io::Error::other("format fixture secondary identifier is missing").into());
        }
        // BNC v5 stores the decimal scalar first, then kind, generic
        // secondary id, and Currency primary id for this exact mask.
        if first.len() < 40 {
            return Err(io::Error::other("format fixture Currency cell is truncated").into());
        }
        first[32..36].copy_from_slice(&identifier.to_le_bytes());
        Ok(())
    })
}

fn currency_secondary_cell() -> FixtureResult<Vec<u8>> {
    // This is a native alternate-number cell captured by the wire codec's
    // own fixture: format id 4 is Currency and generic id 2 is secondary.
    // Keeping the bytes literal ensures the optional-field shape itself is
    // exercised instead of being synthesized by a higher-level setter that
    // would erase one of the two identifiers.
    let bytes = vec![
        0x05, 0x0a, 0x00, 0x00, 0x00, 0x00, 0x03, 0x08, 0x01, 0x70, 0x00, 0x00, 0x39, 0x30, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x3a, 0xb0, 0x02, 0x00,
        0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00,
    ];
    let cell = BncCell::parse(&bytes).map_err(|error| {
        io::Error::other(format!("invalid Currency secondary fixture: {error}"))
    })?;
    if cell.format_identifier() != Some(4) || cell.secondary_format_identifier() != Some(2) {
        return Err(io::Error::other("Currency secondary fixture identifiers are invalid").into());
    }
    Ok(bytes)
}

/// Build one of the controlled malformed graphs from the shared baseline.
pub(crate) fn corrupted_package(corruption: Corruption) -> FixtureResult<Vec<u8>> {
    corrupted_package_for(FormatFamily::Number, corruption)
}

/// Build one controlled malformed graph for a selected format family.
pub(crate) fn corrupted_package_for(
    family: FormatFamily,
    corruption: Corruption,
) -> FixtureResult<Vec<u8>> {
    let source = synthetic_package_for(family, FormatSharing::Shared)?;
    match corruption {
        Corruption::DuplicateFormatKey => rewrite_format_list(&source, |list| {
            let entry = list
                .entries
                .first()
                .cloned()
                .ok_or_else(|| io::Error::other("format fixture entry is missing"))?;
            list.entries.push(entry);
            Ok(())
        }),
        Corruption::MissingFormatEntry => rewrite_format_list(&source, |list| {
            list.entries.retain(|entry| entry.key != FIRST_FORMAT_KEY);
            Ok(())
        }),
        Corruption::RefcountMismatch => rewrite_format_list(&source, |list| {
            let entry = list
                .entries
                .first_mut()
                .ok_or_else(|| io::Error::other("format fixture entry is missing"))?;
            entry.refcount = 1;
            Ok(())
        }),
        Corruption::DuplicateFormatList => rewrite_tables(&source, |archive| {
            let sidecars = archive
                .object_mut(SIDECAR_ID)
                .ok_or_else(|| io::Error::other("format sidecar is missing"))?;
            let message = sidecars
                .messages
                .iter()
                .find(|message| {
                    message.type_ == TABLE_DATA_LIST_TYPE
                        && tst::TableDataList::decode(message.data.as_slice())
                            .map(|list| {
                                list.list_type == tst::table_data_list::ListType::Format as i32
                            })
                            .unwrap_or(false)
                })
                .cloned()
                .ok_or_else(|| io::Error::other("format list message is missing"))?;
            let message_info = sidecars
                .archive_info
                .message_infos
                .iter()
                .find(|info| info.type_ == TABLE_DATA_LIST_TYPE)
                .cloned()
                .ok_or_else(|| io::Error::other("format list metadata is missing"))?;
            sidecars.messages.push(message.clone());
            sidecars.archive_info.message_infos.push(message_info);
            Ok(())
        }),
        Corruption::AliasedFormatList => aliased_format_list(&source),
        Corruption::UnsupportedFormatType => rewrite_format_list(&source, |list| {
            let format = list
                .entries
                .first_mut()
                .and_then(|entry| entry.format.as_mut())
                .ok_or_else(|| io::Error::other("format fixture payload is missing"))?;
            format.format_type = Some(65_535);
            Ok(())
        }),
        Corruption::MalformedFormatPayload => rewrite_format_list_raw(&source, |payload| {
            // The enclosing TableDataList remains valid; only its nested
            // ListEntry.format bytes are malformed.  This distinguishes a
            // hostile format payload from a malformed list envelope.
            replace_first_format_payload(payload, &[0x80])
        }),
        Corruption::WrongCellFormatKey => rewrite_tile(&source, |tile| {
            let row = tile
                .row_infos
                .first_mut()
                .ok_or_else(|| io::Error::other("format fixture row is missing"))?;
            let cells = unpack_row(row)?;
            let first = cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
            let mut cell = BncCell::parse(first)?;
            cell.set_data_format_identifier(
                SECOND_FORMAT_KEY,
                CellDataFormatKind::NumberOrPercentage,
                None,
            )?;
            let mut replacement = cells;
            replacement[0] = cell.encode();
            let (storage, offsets) = pack_row(&replacement)?;
            row.cell_storage_buffer = Some(storage);
            row.cell_offsets = Some(offsets);
            Ok(())
        }),
        Corruption::FractionReplacementMetadata => rewrite_format_list_raw(&source, |payload| {
            let mut hostile = first_format_payload(payload)?;
            // TSK FormatStructArchive field 20 is an optional native
            // replacement marker. Canonical false is valid and must survive
            // a focused Fraction rewrite.
            append_varint_unchecked(&mut hostile, 20, 0);
            replace_first_format_payload(payload, &hostile)
        }),
        Corruption::FractionReplacementMetadataTrue => {
            rewrite_format_list_raw(&source, |payload| {
                let mut hostile = first_format_payload(payload)?;
                // Ordinary type-262 Fraction formats reject the marker when
                // it is true.
                append_varint_unchecked(&mut hostile, 20, 1);
                replace_first_format_payload(payload, &hostile)
            })
        },
        Corruption::UnterminatedUnknownGroup => rewrite_format_list_raw(&source, |payload| {
            let mut hostile = first_format_payload(payload)?;
            hostile.extend_from_slice(&encode_varint(
                (u64::from(UNKNOWN_EXTENSION_FIELD) << 3) | 3,
            ));
            append_varint_unchecked(&mut hostile, UNKNOWN_EXTENSION_FIELD + 1, 0x80_03);
            replace_first_format_payload(payload, &hostile)
        }),
        Corruption::UnexpectedFieldReference => {
            rewrite_member(&source, UNRELATED_MEMBER, |archive| {
                let object = archive
                    .object_mut(UNRELATED_OBJECT_ID)
                    .ok_or_else(|| io::Error::other("unrelated fixture object is missing"))?;
                let info = object
                    .archive_info
                    .message_infos
                    .first_mut()
                    .ok_or_else(|| io::Error::other("unrelated fixture message is missing"))?;
                info.object_references.push(TILE_ID);
                let mut field = FieldInfo::new(FieldPath::new(vec![1]));
                field.r#type = Some(FieldType::ObjectReference);
                field.object_references.push(TILE_ID);
                info.field_infos.push(field);
                Ok(())
            })
        },
        Corruption::ConvertedGenericMissing => {
            let source = text_converted_package()?;
            rewrite_converted_generic_identifier(&source, None)
        },
        Corruption::ConvertedGenericWrong => {
            let source = text_converted_package()?;
            rewrite_converted_generic_identifier(&source, Some(FIRST_FORMAT_KEY))
        },
        Corruption::ConvertedGenericZero => {
            let source = text_converted_package()?;
            rewrite_converted_generic_identifier(&source, Some(0))
        },
        Corruption::ConvertedGenericMissingEntry => {
            let source = text_converted_package()?;
            rewrite_converted_generic_identifier(&source, Some(99))
        },
        Corruption::ConvertedGenericWrongType => {
            let source = text_converted_package()?;
            rewrite_format_varint_by_key(
                &source,
                SECOND_FORMAT_KEY,
                1,
                u64::from(NATIVE_TEXT_FORMAT_TYPE),
            )
        },
        Corruption::ConvertedGenericRefcountMismatch => {
            let source = text_converted_package()?;
            rewrite_format_list_payload_for_test(&source, |list| {
                let entry = list
                    .entries
                    .iter_mut()
                    .find(|entry| entry.key == SECOND_FORMAT_KEY)
                    .ok_or_else(|| io::Error::other("converted generic entry is missing"))?;
                entry.refcount = 2;
                Ok(())
            })
        },
        Corruption::SiblingOrphanTextIdentifier => sibling_text_identifier_without_kind(&source),
        Corruption::SiblingMismatchedTextIdentifier => {
            sibling_text_identifier_with_wrong_kind(&source)
        },
        Corruption::SiblingReservedKnownField => sibling_reserved_known_field(&source),
        Corruption::NonselectedConvertedGenericWrongType => {
            nonselected_converted_generic_wrong_type()
        },
        Corruption::ZeroCellRowStorage => zero_cell_row_with_storage(&source),
    }
}

/// Alias retained for callers that use the shorter fixture spelling.
pub(crate) fn fixture(sharing: FormatSharing) -> FixtureResult<Vec<u8>> {
    synthetic_package(sharing)
}

fn document_object() -> FixtureResult<ArchiveObject> {
    let payload = tn::DocumentArchive {
        sheets: vec![reference(SHEET_ID)],
        ..Default::default()
    }
    .encode_to_vec();
    let mut object = object(DOCUMENT_ID, 1, payload)?;
    object.archive_info.message_infos[0].object_references = vec![SHEET_ID];
    Ok(object)
}

fn sheet_object() -> FixtureResult<ArchiveObject> {
    let mut payload = tn::SheetArchive {
        name: "Data Format Sheet".to_owned(),
        drawable_infos: vec![reference(TABLE_INFO_ID)],
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut payload, 91, 0x51_ee_7a)?;
    let mut object = object(SHEET_ID, 2, payload)?;
    object.archive_info.message_infos[0].object_references = vec![TABLE_INFO_ID];
    Ok(object)
}

fn table_info_object() -> FixtureResult<ArchiveObject> {
    let mut payload = tst::TableInfoArchive {
        super_: tsd::DrawableArchive::default(),
        table_model: reference(TABLE_MODEL_ID),
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut payload, 90, 0x61_ee_7a)?;
    let mut object = object(TABLE_INFO_ID, TABLE_INFO_TYPE, payload)?;
    object.archive_info.message_infos[0].object_references = vec![TABLE_MODEL_ID];
    Ok(object)
}

fn table_model_object() -> FixtureResult<ArchiveObject> {
    let payload = table_model_payload()?.encode_to_vec();
    let mut object = object(TABLE_MODEL_ID, TABLE_MODEL_TYPE, payload)?;
    object.archive_info.message_infos[0].object_references = vec![SIDECAR_ID, TILE_ID];
    let mut unknown = FieldInfo::new(vec![99, 1]);
    unknown.data_references = vec![0xdead_beef];
    object.archive_info.message_infos[0]
        .field_infos
        .push(unknown);
    Ok(object)
}

fn table_model_payload() -> FixtureResult<tst::TableModelArchive> {
    Ok(tst::TableModelArchive {
        table_id: "data-format-table-id".to_owned(),
        table_name: "Data Formats".to_owned(),
        table_style: reference(SIDECAR_ID),
        body_text_style: reference(SIDECAR_ID),
        header_row_text_style: reference(SIDECAR_ID),
        header_column_text_style: reference(SIDECAR_ID),
        footer_row_text_style: reference(SIDECAR_ID),
        body_cell_style: reference(SIDECAR_ID),
        header_row_style: reference(SIDECAR_ID),
        header_column_style: reference(SIDECAR_ID),
        footer_row_style: reference(SIDECAR_ID),
        number_of_rows: 1,
        number_of_columns: 2,
        default_row_height: 20.0,
        default_column_width: 80.0,
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                ..Default::default()
            },
            column_headers: reference(SIDECAR_ID),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: reference(TILE_ID),
                }],
                tile_size: Some(256),
                ..Default::default()
            },
            string_table: reference(SIDECAR_ID),
            style_table: reference(SIDECAR_ID),
            formula_table: reference(SIDECAR_ID),
            format_table_pre_bnc: reference(SIDECAR_ID),
            format_table: Some(reference(SIDECAR_ID)),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        ..Default::default()
    })
}

fn tile_object(family: FormatFamily, sharing: FormatSharing) -> FixtureResult<ArchiveObject> {
    let payload = tile_payload(family, sharing)?.encode_to_vec();
    let mut object = object(TILE_ID, TILE_TYPE, payload)?;
    object.archive_info.message_infos[0].data_references = vec![0x70_01, 0x70_02];
    Ok(object)
}

fn tile_payload(family: FormatFamily, sharing: FormatSharing) -> FixtureResult<tst::Tile> {
    let first_kind = match family {
        FormatFamily::Currency => CellDataFormatKind::Currency,
        FormatFamily::Number
        | FormatFamily::Percentage
        | FormatFamily::Scientific
        | FormatFamily::Fraction => CellDataFormatKind::NumberOrPercentage,
        FormatFamily::DateTime => CellDataFormatKind::DateTime,
        FormatFamily::Text => CellDataFormatKind::Text,
    };
    let first = if matches!(family, FormatFamily::Text) {
        formatted_text_cell(FIRST_FORMAT_KEY, 1)?
    } else if matches!(family, FormatFamily::DateTime) {
        formatted_date_time_cell(FIRST_FORMAT_KEY, 45200.5)?
    } else {
        formatted_cell(FIRST_FORMAT_KEY, first_kind, 1234.5)?
    };
    let second_key = match sharing {
        FormatSharing::Shared => FIRST_FORMAT_KEY,
        FormatSharing::Unshared => SECOND_FORMAT_KEY,
    };
    let second_kind = match (family, sharing) {
        (FormatFamily::Currency, FormatSharing::Shared) => CellDataFormatKind::Currency,
        (FormatFamily::DateTime, _) => CellDataFormatKind::DateTime,
        (FormatFamily::Text, _) => CellDataFormatKind::Text,
        (_, FormatSharing::Shared) => CellDataFormatKind::NumberOrPercentage,
        (_, FormatSharing::Unshared) => CellDataFormatKind::NumberOrPercentage,
    };
    let second = if matches!(family, FormatFamily::Text) {
        formatted_text_cell(second_key, 2)?
    } else if matches!(family, FormatFamily::DateTime) {
        formatted_date_time_cell(second_key, 45201.25)?
    } else {
        formatted_cell(second_key, second_kind, 0.25)?
    };
    let (storage, offsets) = pack_row(&[first, second])?;
    Ok(tst::Tile {
        max_column: 1,
        max_row: 0,
        num_cells: 2,
        numrows: 1,
        row_infos: vec![tst::TileRowInfo {
            tile_row_index: 0,
            cell_count: 2,
            storage_version: Some(5),
            cell_storage_buffer_pre_bnc: storage.clone(),
            cell_offsets_pre_bnc: offsets.clone(),
            cell_storage_buffer: Some(storage),
            cell_offsets: Some(offsets),
            ..Default::default()
        }],
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        ..Default::default()
    })
}

fn formatted_cell(format_key: u32, kind: CellDataFormatKind, value: f64) -> FixtureResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_data_format_identifier(format_key, kind, None)?;
    cell.set_number(value)?;
    Ok(cell.encode())
}

fn formatted_text_cell(format_key: u32, string_key: u32) -> FixtureResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_string(string_key);
    cell.set_data_format_identifier(format_key, CellDataFormatKind::Text, None)?;
    Ok(cell.encode())
}

fn formatted_date_time_cell(format_key: u32, value: f64) -> FixtureResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_date(value)?;
    cell.set_data_format_identifier(format_key, CellDataFormatKind::DateTime, None)?;
    Ok(cell.encode())
}

fn sidecar_object(family: FormatFamily, sharing: FormatSharing) -> FixtureResult<ArchiveObject> {
    let format_entries = match (family, sharing) {
        (FormatFamily::Currency, FormatSharing::Shared) => vec![currency_format_entry(
            FIRST_FORMAT_KEY,
            2,
            "USD",
            2,
            2,
            true,
            false,
        )],
        (FormatFamily::Scientific, FormatSharing::Shared) => {
            vec![scientific_format_entry(FIRST_FORMAT_KEY, 2)]
        },
        (FormatFamily::Fraction, FormatSharing::Shared) => {
            vec![fraction_format_entry(FIRST_FORMAT_KEY, 2, 8)]
        },
        (FormatFamily::DateTime, FormatSharing::Shared) => {
            vec![date_time_format_entry(
                FIRST_FORMAT_KEY,
                2,
                "yyyy-MM-dd H:mm:ss",
            )]
        },
        (FormatFamily::DateTime, FormatSharing::Unshared) => vec![
            date_time_format_entry(FIRST_FORMAT_KEY, 1, "yyyy-MM-dd H:mm:ss"),
            date_time_format_entry(SECOND_FORMAT_KEY, 1, "MM/dd/yyyy"),
        ],
        (FormatFamily::Text, FormatSharing::Shared) => {
            vec![text_format_entry(FIRST_FORMAT_KEY, 2)]
        },
        (FormatFamily::Text, FormatSharing::Unshared) => vec![
            text_format_entry(FIRST_FORMAT_KEY, 1),
            text_format_entry(SECOND_FORMAT_KEY, 1),
        ],
        (FormatFamily::Currency, FormatSharing::Unshared) => vec![
            format_entry(FIRST_FORMAT_KEY, 1, NATIVE_NUMBER_FORMAT_TYPE, 2, 2, true),
            format_entry(
                SECOND_FORMAT_KEY,
                1,
                NATIVE_PERCENTAGE_FORMAT_TYPE,
                1,
                0,
                false,
            ),
        ],
        (_, FormatSharing::Shared) => vec![format_entry(
            FIRST_FORMAT_KEY,
            2,
            family.native_type(),
            2,
            2,
            true,
        )],
        (_, FormatSharing::Unshared) => vec![
            format_entry(FIRST_FORMAT_KEY, 1, NATIVE_NUMBER_FORMAT_TYPE, 2, 2, true),
            format_entry(
                SECOND_FORMAT_KEY,
                1,
                NATIVE_PERCENTAGE_FORMAT_TYPE,
                1,
                0,
                false,
            ),
        ],
    };
    let list_specs = [
        (tst::table_data_list::ListType::String, Vec::new()),
        (tst::table_data_list::ListType::Formula, Vec::new()),
        (tst::table_data_list::ListType::Format, format_entries),
    ];
    let mut messages = Vec::with_capacity(list_specs.len());
    for (list_type, entries) in list_specs {
        let mut payload = tst::TableDataList {
            list_type: list_type as i32,
            next_list_id: 32,
            entries,
            is_new_for_bnc: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        if matches!(list_type, tst::table_data_list::ListType::Format) {
            payload = add_unknown_extension_to_format_payload(&payload)?;
        }
        append_varint_field(&mut payload, 90, 0x80_00 + list_type as u64)?;
        if matches!(list_type, tst::table_data_list::ListType::Format) {
            // Keep one root-level unknown scalar beside the nested extension.
            append_varint_field(&mut payload, UNKNOWN_EXTENSION_FIELD, 0x80_03)?;
        }
        messages.push(RawMessage {
            type_: TABLE_DATA_LIST_TYPE,
            data: payload,
        });
    }
    let mut sidecars = ArchiveObject::new(SIDECAR_ID, messages)?;
    for info in &mut sidecars.archive_info.message_infos {
        info.data_references = vec![0x72_00 + u64::from(info.type_)];
    }
    Ok(sidecars)
}

fn format_entry(
    key: u32,
    refcount: u32,
    format_type: u32,
    decimal_places: u32,
    negative_style: u32,
    show_thousands_separator: bool,
) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(format_type),
            decimal_places: Some(decimal_places),
            negative_style: Some(negative_style),
            show_thousands_separator: Some(show_thousands_separator),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn currency_format_entry(
    key: u32,
    refcount: u32,
    currency_code: &str,
    decimal_places: u32,
    negative_style: u32,
    show_thousands_separator: bool,
    use_accounting_style: bool,
) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(NATIVE_CURRENCY_FORMAT_TYPE),
            decimal_places: Some(decimal_places),
            negative_style: Some(negative_style),
            show_thousands_separator: Some(show_thousands_separator),
            currency_code: Some(currency_code.to_owned()),
            use_accounting_style: Some(use_accounting_style),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn scientific_format_entry(key: u32, refcount: u32) -> tst::table_data_list::ListEntry {
    // Scientific is encoded in the same native FormatStructArchive envelope
    // as Number/Percentage, but its decimal options are canonical: fixed
    // precision, minus-sign negatives, and no thousands separator.
    format_entry(key, refcount, NATIVE_SCIENTIFIC_FORMAT_TYPE, 2, 0, false)
}

fn fraction_format_entry(
    key: u32,
    refcount: u32,
    accuracy: u32,
) -> tst::table_data_list::ListEntry {
    // Fraction has its own native discriminator and stores the denominator
    // strategy in `fraction_accuracy`; unlike decimal families it must not
    // carry decimal-place, sign, or grouping fields.
    tst::table_data_list::ListEntry {
        key,
        refcount,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(NATIVE_FRACTION_FORMAT_TYPE),
            fraction_accuracy: Some(accuracy),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn date_time_format_entry(
    key: u32,
    refcount: u32,
    pattern: &str,
) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(NATIVE_DATE_TIME_FORMAT_TYPE),
            date_time_format: Some(pattern.to_owned()),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn text_format_entry(key: u32, refcount: u32) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(NATIVE_TEXT_FORMAT_TYPE),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn add_unknown_extension_to_format_payload(source: &[u8]) -> FixtureResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::new();
    for field in view.fields() {
        if field.number() != 3 {
            output.extend_from_slice(field.raw());
            continue;
        }
        let entry_view = WireView::parse(field.payload())?;
        let mut entry = Vec::new();
        let mut format_found = false;
        for entry_field in entry_view.fields() {
            if entry_field.number() == 6 && !format_found {
                let mut format = entry_field.payload().to_vec();
                append_varint_field(&mut format, UNKNOWN_EXTENSION_FIELD, 0x80_03)?;
                append_length_delimited_field(&mut entry, 6, &format)?;
                format_found = true;
            } else {
                entry.extend_from_slice(entry_field.raw());
            }
        }
        if !format_found {
            return Err(io::Error::other("format payload field is missing").into());
        }
        append_length_delimited_field(&mut output, 3, &entry)?;
    }
    Ok(output)
}

fn unrelated_object() -> FixtureResult<ArchiveObject> {
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 89, b"untouched unrelated archive payload")?;
    append_varint_field(&mut payload, 90, 0x90_ee_7a)?;
    append_varint_field(&mut payload, 91, 0x91_ee_7a)?;
    object(UNRELATED_OBJECT_ID, 99_991, payload)
}

fn view_state_object() -> FixtureResult<ArchiveObject> {
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 89, b"deterministic view state payload")?;
    append_varint_field(&mut payload, 90, 0x80_ee_7a)?;
    object(VIEW_STATE_OBJECT_ID, 777, payload)
}

fn metadata_object() -> FixtureResult<ArchiveObject> {
    let mut data = tsp::PackageMetadata {
        last_object_identifier: 1_000,
        save_token: Some(10),
        ..Default::default()
    }
    .encode_to_vec();
    let document = metadata_component(100, "Document", 9, &[DOCUMENT_ID, SHEET_ID])?;
    let tables = metadata_component(
        200,
        "Tables",
        8,
        &[TABLE_INFO_ID, TABLE_MODEL_ID, SIDECAR_ID, TILE_ID],
    )?;
    let view = metadata_component(300, "ViewState", 7, &[])?;
    let versioned = metadata_component(100, "Document", 3, &[0xfeed])?;
    append_length_delimited_field(&mut data, 3, &document)?;
    append_length_delimited_field(&mut data, 3, &tables)?;
    append_length_delimited_field(&mut data, 3, &view)?;
    append_length_delimited_field(&mut data, 11, &versioned)?;
    append_varint_field(&mut data, 90, 0xa5_a5_a5)?;
    object(METADATA_OBJECT_ID, METADATA_TYPE, data)
}

fn metadata_component(
    identifier: u64,
    locator: &str,
    save_token: u64,
    object_ids: &[u64],
) -> FixtureResult<Vec<u8>> {
    let mut data = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(save_token),
        object_uuid_map_entries: object_ids.iter().copied().map(uuid_entry).collect(),
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut data, 90, identifier + 10_000)?;
    Ok(data)
}

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier,
            upper: identifier + 0x1000,
        },
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> FixtureResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn compressed(objects: Vec<ArchiveObject>) -> FixtureResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

/// Read one decompressed IWA member into an owned archive.
pub(crate) fn member_archive(source: &[u8], member: &str) -> FixtureResult<Archive> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("member {member} is missing")))?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

/// Return one raw message payload from an IWA member.
pub(crate) fn object_message(
    source: &[u8],
    member: &str,
    identifier: u64,
    message_type: u32,
) -> FixtureResult<Vec<u8>> {
    let archive = member_archive(source, member)?;
    archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == message_type)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| {
            io::Error::other(format!(
                "object {identifier} message {message_type} is missing"
            ))
            .into()
        })
}

/// Replace one complete compressed member while retaining all other ZIP
/// records exactly.  This is the common locality primitive for hostile tests.
pub(crate) fn rewrite_member(
    source: &[u8],
    member: &str,
    mutate: impl FnOnce(&mut Archive) -> FixtureResult,
) -> FixtureResult<Vec<u8>> {
    let mut archive = member_archive(source, member)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            member,
            &compressed,
        )],
        Limits::default(),
    )?)
}

/// Rewrite the table member, the target of format edits in this fixture.
pub(crate) fn rewrite_tables(
    source: &[u8],
    mutate: impl FnOnce(&mut Archive) -> FixtureResult,
) -> FixtureResult<Vec<u8>> {
    rewrite_member(source, TABLES_MEMBER, mutate)
}

/// Mark the rooted fixture table as locked while retaining the surrounding
/// TableInfo wire records exactly.  The lock lives on the inherited
/// `TSD.DrawableArchive` envelope (TableInfo field 1, Drawable field 5), so
/// patching that scalar exercises the same native ownership path as a source
/// opened from Numbers without re-encoding the complete message.
pub(crate) fn locked_table_package(source: &[u8]) -> FixtureResult<Vec<u8>> {
    rewrite_tables(source, |archive| {
        let table_info = archive
            .object_mut(TABLE_INFO_ID)
            .ok_or_else(|| io::Error::other("format fixture table info is missing"))?;
        let message = table_info
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture table info payload is missing"))?;
        message.data = patch_nested_varint_field(&message.data, &[1, 5], false, Some(1))?;
        Ok(())
    })
}

/// Build a Text fixture whose first cell has inherited formatting while the
/// sibling retains the shared explicit Text entry.  This exercises the
/// focused owner's ability to attach/detach an existing entry without
/// manufacturing a second native record.
pub(crate) fn text_inherited_first_package() -> FixtureResult<Vec<u8>> {
    let source = synthetic_package_for(FormatFamily::Text, FormatSharing::Shared)?;
    let source = rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut cell = BncCell::parse(first)?;
        cell.clear_explicit_format();
        *first = cell.encode();
        Ok(())
    })?;
    rewrite_format_list_payload_for_test(&source, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("format fixture Text entry is missing"))?;
        entry.refcount = 1;
        Ok(())
    })
}

/// Build a Text fixture using the native converted-text marker (`0x81`) on
/// the selected cell.  The marker is distinct from canonical explicit Text
/// (`0x80`) but has the same semantic owner and must survive an exact no-op.
pub(crate) fn text_converted_package() -> FixtureResult<Vec<u8>> {
    let source = synthetic_package_for(FormatFamily::Text, FormatSharing::Shared)?;
    let source = rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut encoded = BncCell::parse(first)?.encode();
        if encoded.len() < 24 {
            return Err(io::Error::other("format fixture first cell is truncated").into());
        }
        // Converted Text keeps a generic Number format reference in the
        // ordinary decimal identifier slot (field 0x2000), immediately
        // before the Text-specific identifier.  Add that fixed-width field
        // rather than merely flipping 0x80 to 0x81; the latter is an invalid
        // BNC record and should never reach the focused owner.
        let flags = u32::from_le_bytes(
            encoded[8..12]
                .try_into()
                .map_err(|_| io::Error::other("format fixture flags are truncated"))?,
        );
        if flags & 0x0000_2000 != 0 {
            return Err(
                io::Error::other("format fixture already has a generic Text reference").into(),
            );
        }
        encoded[8..12].copy_from_slice(&(flags | 0x0000_2000).to_le_bytes());
        encoded.splice(20..20, SECOND_FORMAT_KEY.to_le_bytes());
        encoded[6..8]
            .copy_from_slice(&litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT.to_le_bytes());
        *first = encoded;
        Ok(())
    })?;
    // The converted cell's generic identifier points at a live Number entry;
    // its primary display entry remains the shared Text key.
    rewrite_format_list_payload_for_test(&source, |list| {
        list.entries.push(format_entry(
            SECOND_FORMAT_KEY,
            1,
            NATIVE_NUMBER_FORMAT_TYPE,
            2,
            2,
            true,
        ));
        Ok(())
    })
}

/// Build a converted-Text fixture where both cells retain the generic Number
/// edge.  Keeping that edge live after clearing the selected cell lets a test
/// prove source-preserving rewrites of the generic format payload itself;
/// the ordinary one-converted-cell fixture intentionally culls that entry.
pub(crate) fn text_converted_shared_generic_package() -> FixtureResult<Vec<u8>> {
    let source = text_converted_package()?;
    let source = rewrite_tile_cells(&source, |cells| {
        let second = cells
            .get_mut(1)
            .ok_or_else(|| io::Error::other("format fixture second cell is missing"))?;
        *second = converted_text_cell_with_generic(second, SECOND_FORMAT_KEY)?;
        Ok(())
    })?;
    rewrite_format_list_payload_for_test(&source, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == SECOND_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("converted generic entry is missing"))?;
        entry.refcount = 2;
        Ok(())
    })
}

/// Return a converted-Text cell carrying exactly the requested generic key.
fn converted_text_cell_with_generic(
    source: &[u8],
    generic_identifier: u32,
) -> FixtureResult<Vec<u8>> {
    let mut encoded = BncCell::parse(source)?.encode();
    if encoded.len() < 24 {
        return Err(io::Error::other("format fixture Text cell is truncated").into());
    }
    let flags = u32::from_le_bytes(
        encoded[8..12]
            .try_into()
            .map_err(|_| io::Error::other("format fixture Text flags are truncated"))?,
    );
    if flags & 0x0002_0000 == 0 {
        return Err(io::Error::other("format fixture Text identifier is missing").into());
    }
    encoded[8..12].copy_from_slice(&(flags | 0x0000_2000).to_le_bytes());
    if encoded.len() >= 28 {
        encoded[20..24].copy_from_slice(&generic_identifier.to_le_bytes());
    } else {
        encoded.splice(20..20, generic_identifier.to_le_bytes());
    }
    encoded[6..8]
        .copy_from_slice(&litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT.to_le_bytes());
    Ok(encoded)
}

/// Rewrite the selected converted-Text cell's generic Number identifier.
/// This keeps the package envelope and all unrelated records unchanged so
/// owner-level validation, rather than package ingress, gets the decision.
fn rewrite_converted_generic_identifier(
    source: &[u8],
    generic_identifier: Option<u32>,
) -> FixtureResult<Vec<u8>> {
    rewrite_tile_cells(source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut encoded = BncCell::parse(first)?.encode();
        if encoded.len() < 24 {
            return Err(io::Error::other("format fixture Text cell is truncated").into());
        }
        let flags = u32::from_le_bytes(
            encoded[8..12]
                .try_into()
                .map_err(|_| io::Error::other("format fixture Text flags are truncated"))?,
        );
        match generic_identifier {
            Some(identifier) => {
                encoded[8..12].copy_from_slice(&(flags | 0x0000_2000).to_le_bytes());
                if encoded.len() >= 28 {
                    encoded[20..24].copy_from_slice(&identifier.to_le_bytes());
                } else {
                    encoded.splice(20..20, identifier.to_le_bytes());
                }
            },
            None => {
                if flags & 0x0000_2000 != 0 {
                    if encoded.len() < 28 {
                        return Err(
                            io::Error::other("format fixture generic field is truncated").into(),
                        );
                    }
                    encoded.drain(20..24);
                    encoded[8..12].copy_from_slice(&(flags & !0x0000_2000).to_le_bytes());
                }
            },
        }
        encoded[6..8]
            .copy_from_slice(&litchi_numbers_wire::EXPLICIT_CONVERTED_TEXT_FORMAT.to_le_bytes());
        *first = encoded;
        Ok(())
    })
}

fn sibling_text_identifier_without_kind(source: &[u8]) -> FixtureResult<Vec<u8>> {
    let source = rewrite_tile_cells(source, |cells| {
        let second = cells
            .get_mut(1)
            .ok_or_else(|| io::Error::other("format fixture second cell is missing"))?;
        let mut encoded = BncCell::parse(second)?.encode();
        if encoded.len() < 24 {
            return Err(io::Error::other("format fixture Text cell is truncated").into());
        }
        let flags = u32::from_le_bytes(
            encoded[8..12]
                .try_into()
                .map_err(|_| io::Error::other("format fixture Text flags are truncated"))?,
        );
        if flags & BNC_TEXT_FORMAT_IDENTIFIER_FLAG == 0 || flags & BNC_CELL_FORMAT_KIND_FLAG == 0 {
            return Err(io::Error::other("format fixture Text metadata is missing").into());
        }
        // The canonical Text layout is string, kind, then Text identifier.
        // Remove only the kind bytes and clear the marker, leaving the known
        // Text field physically present but semantically orphaned.
        encoded.drain(16..20);
        encoded[8..12].copy_from_slice(&(flags & !BNC_CELL_FORMAT_KIND_FLAG).to_le_bytes());
        encoded[6..8].fill(0);
        *second = encoded;
        Ok(())
    })?;
    rewrite_format_list_payload_for_test(&source, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("format fixture Text entry is missing"))?;
        // Only the selected first cell has a visible primary edge after the
        // sibling's kind is removed.
        entry.refcount = 1;
        Ok(())
    })
}

fn sibling_text_identifier_with_wrong_kind(source: &[u8]) -> FixtureResult<Vec<u8>> {
    let source = rewrite_tile_cells(source, |cells| {
        let second = cells
            .get_mut(1)
            .ok_or_else(|| io::Error::other("format fixture second cell is missing"))?;
        let mut encoded = BncCell::parse(second)?.encode();
        if encoded.len() < 24 {
            return Err(io::Error::other("format fixture Text cell is truncated").into());
        }
        let flags = u32::from_le_bytes(
            encoded[8..12]
                .try_into()
                .map_err(|_| io::Error::other("format fixture Text flags are truncated"))?,
        );
        if flags & BNC_TEXT_FORMAT_IDENTIFIER_FLAG == 0 || flags & BNC_CELL_FORMAT_KIND_FLAG == 0 {
            return Err(io::Error::other("format fixture Text metadata is missing").into());
        }
        // Keep both fixed-width fields but make the kind decimal. The old
        // kind-directed census ignores the Text field and sees no edge here.
        encoded[16..20]
            .copy_from_slice(&litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND.to_le_bytes());
        encoded[6..8].fill(0);
        *second = encoded;
        Ok(())
    })?;
    rewrite_format_list_payload_for_test(&source, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("format fixture Text entry is missing"))?;
        entry.refcount = 1;
        Ok(())
    })
}

fn sibling_reserved_known_field(source: &[u8]) -> FixtureResult<Vec<u8>> {
    let source = rewrite_tile_cells(source, |cells| {
        let second = cells
            .get_mut(1)
            .ok_or_else(|| io::Error::other("format fixture second cell is missing"))?;
        let mut cell = BncCell::parse(second)?;
        cell.clear_explicit_format();
        let mut encoded = cell.encode();
        if encoded.len() < 12 {
            return Err(io::Error::other("format fixture cell is truncated").into());
        }
        let flags = u32::from_le_bytes(
            encoded[8..12]
                .try_into()
                .map_err(|_| io::Error::other("format fixture flags are truncated"))?,
        );
        encoded[8..12].copy_from_slice(&(flags | BNC_RESERVED_KNOWN_FIELD_FLAG).to_le_bytes());
        // The reserved field is the final fixed-width known slot.
        encoded.extend_from_slice(&[0; 4]);
        *second = encoded;
        Ok(())
    })?;
    rewrite_format_list_payload_for_test(&source, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("format fixture Text entry is missing"))?;
        entry.refcount = 1;
        Ok(())
    })
}

fn nonselected_converted_generic_wrong_type() -> FixtureResult<Vec<u8>> {
    let source = synthetic_package_for(FormatFamily::Text, FormatSharing::Shared)?;
    let source = rewrite_tile_cells(&source, |cells| {
        let second = cells
            .get_mut(1)
            .ok_or_else(|| io::Error::other("format fixture second cell is missing"))?;
        *second = converted_text_cell_with_generic(second, SECOND_FORMAT_KEY)?;
        Ok(())
    })?;
    let source = rewrite_format_list_payload_for_test(&source, |list| {
        list.entries.push(format_entry(
            SECOND_FORMAT_KEY,
            1,
            NATIVE_NUMBER_FORMAT_TYPE,
            2,
            2,
            true,
        ));
        Ok(())
    })?;
    rewrite_format_varint_by_key(
        &source,
        SECOND_FORMAT_KEY,
        1,
        u64::from(NATIVE_TEXT_FORMAT_TYPE),
    )
}

fn zero_cell_row_with_storage(source: &[u8]) -> FixtureResult<Vec<u8>> {
    // Keep the extra row inside the declared table height so package ingress
    // admits the graph and the focused owner's full tile census gets to see
    // the malformed row.
    let source = rewrite_tables(source, |archive| {
        let model = archive
            .object_mut(TABLE_MODEL_ID)
            .ok_or_else(|| io::Error::other("format fixture model is missing"))?;
        let message = model
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture model payload is missing"))?;
        let mut decoded = tst::TableModelArchive::decode(message.data.as_slice())?;
        decoded.number_of_rows = 2;
        message.data = decoded.encode_to_vec();
        Ok(())
    })?;
    rewrite_tile(&source, |tile| {
        let first = tile
            .row_infos
            .first()
            .ok_or_else(|| io::Error::other("format fixture row is missing"))?;
        let storage = first
            .cell_storage_buffer
            .clone()
            .ok_or_else(|| io::Error::other("format fixture row storage is missing"))?;
        tile.row_infos.push(tst::TileRowInfo {
            tile_row_index: 1,
            cell_count: 0,
            cell_storage_buffer_pre_bnc: storage.clone(),
            cell_offsets_pre_bnc: vec![u8::MAX, u8::MAX],
            storage_version: Some(5),
            cell_storage_buffer: Some(storage),
            cell_offsets: Some(vec![u8::MAX, u8::MAX]),
            ..Default::default()
        });
        Ok(())
    })
}

fn rewrite_format_list(
    source: &[u8],
    mutate: impl FnOnce(&mut tst::TableDataList) -> FixtureResult,
) -> FixtureResult<Vec<u8>> {
    rewrite_format_list_raw(source, |payload| {
        let mut list = tst::TableDataList::decode(payload)?;
        mutate(&mut list)?;
        Ok(list.encode_to_vec())
    })
}

/// Rewrite the selected format list for focused integration fixtures.
///
/// This narrow test-only seam lets a package test construct a valid graph
/// with a deliberately stale, missing, or wrong-family secondary reference
/// without exposing the archive mutation machinery to production callers.
pub(crate) fn rewrite_format_list_payload_for_test(
    source: &[u8],
    mutate: impl FnOnce(&mut tst::TableDataList) -> FixtureResult,
) -> FixtureResult<Vec<u8>> {
    rewrite_format_list(source, mutate)
}

/// Return the current BNC format-list message payload.
pub(crate) fn format_list_payload(source: &[u8]) -> FixtureResult<Vec<u8>> {
    let archive = member_archive(source, TABLES_MEMBER)?;
    let sidecars = archive
        .object(SIDECAR_ID)
        .ok_or_else(|| io::Error::other("format sidecar is missing"))?;
    sidecars
        .messages
        .iter()
        .find(|message| {
            message.type_ == TABLE_DATA_LIST_TYPE
                && tst::TableDataList::decode(message.data.as_slice())
                    .map(|list| list.list_type == tst::table_data_list::ListType::Format as i32)
                    .unwrap_or(false)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("format list message is missing").into())
}

/// Return sorted `(format-list key, refcount)` facts for the current list.
pub(crate) fn format_entry_facts(source: &[u8]) -> FixtureResult<Vec<(u32, u32)>> {
    let list = tst::TableDataList::decode(format_list_payload(source)?.as_slice())?;
    let mut facts = list
        .entries
        .into_iter()
        .map(|entry| (entry.key, entry.refcount))
        .collect::<Vec<_>>();
    facts.sort_unstable();
    Ok(facts)
}

/// Return the list allocator cursor without exposing native package objects.
pub(crate) fn format_next_list_id(source: &[u8]) -> FixtureResult<u32> {
    Ok(tst::TableDataList::decode(format_list_payload(source)?.as_slice())?.next_list_id)
}

/// Return one raw nested format payload by its list key.
///
/// The lookup is wire-oriented rather than a prost re-encode so unknown
/// fields and non-canonical bytes remain available for exact preservation
/// assertions.
pub(crate) fn format_payload_by_key(source: &[u8], key: u32) -> FixtureResult<Vec<u8>> {
    let payload = format_list_payload(source)?;
    let list = WireView::parse(&payload)?;
    for field in list.fields().filter(|field| field.number() == 3) {
        let entry = tst::table_data_list::ListEntry::decode(field.payload())?;
        if entry.key != key {
            continue;
        }
        let entry_view = WireView::parse(field.payload())?;
        return entry_view
            .fields()
            .find(|entry_field| entry_field.number() == 6)
            .map(|entry_field| entry_field.payload().to_vec())
            .ok_or_else(|| io::Error::other("format payload field is missing").into());
    }
    Err(io::Error::other(format!("format entry key {key} is missing")).into())
}

/// Return one exact raw extension record from a decoded format payload.
pub(crate) fn unknown_field_record(payload: &[u8], field_number: u32) -> FixtureResult<Vec<u8>> {
    WireView::parse(payload)?
        .fields()
        .find(|field| field.number() == field_number)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("unknown extension is missing").into())
}

/// Return all encoded BNC cells in the fixture tile's first row.
pub(crate) fn tile_cells(source: &[u8]) -> FixtureResult<Vec<Vec<u8>>> {
    let payload = object_message(source, TABLES_MEMBER, TILE_ID, TILE_TYPE)?;
    let tile = tst::Tile::decode(payload.as_slice())?;
    let row = tile
        .row_infos
        .first()
        .ok_or_else(|| io::Error::other("format fixture row is missing"))?;
    unpack_row(row)
}

/// Return each first-row cell's current BNC primary format key.
pub(crate) fn format_keys(source: &[u8]) -> FixtureResult<Vec<Option<u32>>> {
    tile_cells(source)?
        .into_iter()
        .map(|cell| Ok(BncCell::parse(&cell)?.format_identifier()))
        .collect()
}

/// Rewrite first-row BNC cells while retaining all other fixture bytes.
pub(crate) fn rewrite_tile_cells(
    source: &[u8],
    mutate: impl FnOnce(&mut Vec<Vec<u8>>) -> FixtureResult,
) -> FixtureResult<Vec<u8>> {
    rewrite_tile(source, |tile| {
        let row = tile
            .row_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture row is missing"))?;
        let mut cells = unpack_row(row)?;
        mutate(&mut cells)?;
        let (storage, offsets) = pack_row(&cells)?;
        row.cell_storage_buffer = Some(storage.clone());
        row.cell_offsets = Some(offsets.clone());
        row.cell_storage_buffer_pre_bnc = storage;
        row.cell_offsets_pre_bnc = offsets;
        Ok(())
    })
}

fn rewrite_format_list_raw(
    source: &[u8],
    mutate: impl FnOnce(&[u8]) -> FixtureResult<Vec<u8>>,
) -> FixtureResult<Vec<u8>> {
    rewrite_tables(source, |archive| {
        let sidecars = archive
            .object_mut(SIDECAR_ID)
            .ok_or_else(|| io::Error::other("format sidecar is missing"))?;
        let message = sidecars
            .messages
            .iter_mut()
            .find(|message| {
                message.type_ == TABLE_DATA_LIST_TYPE
                    && tst::TableDataList::decode(message.data.as_slice())
                        .map(|list| list.list_type == tst::table_data_list::ListType::Format as i32)
                        .unwrap_or(false)
            })
            .ok_or_else(|| io::Error::other("format list message is missing"))?;
        message.data = mutate(message.data.as_slice())?;
        Ok(())
    })
}

/// Replace one nested format payload selected by its list key.
///
/// This is deliberately kept raw: decoding and re-encoding the enclosing
/// list with prost would discard unknown fields that owner rewrites are
/// required to preserve byte-for-byte.
pub(crate) fn rewrite_format_payload_by_key(
    source: &[u8],
    key: u32,
    replacement: &[u8],
) -> FixtureResult<Vec<u8>> {
    rewrite_format_list_raw(source, |payload| {
        let view = WireView::parse(payload)?;
        let mut output = Vec::with_capacity(payload.len());
        let mut replaced = false;
        for field in view.fields() {
            if field.number() != 3 || replaced {
                output.extend_from_slice(field.raw());
                continue;
            }
            let entry = tst::table_data_list::ListEntry::decode(field.payload())?;
            if entry.key != key {
                output.extend_from_slice(field.raw());
                continue;
            }
            let entry_view = WireView::parse(field.payload())?;
            let mut entry_output = Vec::with_capacity(field.payload().len());
            let mut format_found = false;
            for entry_field in entry_view.fields() {
                if entry_field.number() == 6 && !format_found {
                    append_length_delimited_field(&mut entry_output, 6, replacement)?;
                    format_found = true;
                } else {
                    entry_output.extend_from_slice(entry_field.raw());
                }
            }
            if !format_found {
                return Err(io::Error::other("format payload field is missing").into());
            }
            append_length_delimited_field(&mut output, 3, &entry_output)?;
            replaced = true;
        }
        if !replaced {
            return Err(io::Error::other(format!("format entry key {key} is missing")).into());
        }
        Ok(output)
    })
}

/// Patch one known native-format varint while preserving all other wire
/// records.  The caller can use this to make a package-valid payload reach
/// the semantic owner with an intentionally invalid domain or field shape.
pub(crate) fn rewrite_format_varint_by_key(
    source: &[u8],
    key: u32,
    field_number: u32,
    value: u64,
) -> FixtureResult<Vec<u8>> {
    let payload = format_payload_by_key(source, key)?;
    let replacement = patch_nested_varint_field(&payload, &[field_number], true, Some(value))?;
    rewrite_format_payload_by_key(source, key, &replacement)
}

fn rewrite_tile(
    source: &[u8],
    mutate: impl FnOnce(&mut tst::Tile) -> FixtureResult,
) -> FixtureResult<Vec<u8>> {
    rewrite_tables(source, |archive| {
        let tile = archive
            .object_mut(TILE_ID)
            .ok_or_else(|| io::Error::other("format fixture tile is missing"))?;
        let message = tile
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture tile payload is missing"))?;
        let mut decoded = tst::Tile::decode(message.data.as_slice())?;
        mutate(&mut decoded)?;
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn aliased_format_list(source: &[u8]) -> FixtureResult<Vec<u8>> {
    rewrite_tables(source, |archive| {
        let model = archive
            .object_mut(TABLE_MODEL_ID)
            .ok_or_else(|| io::Error::other("format fixture model is missing"))?;
        let message = model
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture model payload is missing"))?;
        let mut decoded = tst::TableModelArchive::decode(message.data.as_slice())?;
        decoded.base_data_store.format_table = Some(reference(ALIASED_SIDECAR_ID));
        message.data = decoded.encode_to_vec();

        let sidecar = archive
            .object(SIDECAR_ID)
            .cloned()
            .ok_or_else(|| io::Error::other("format fixture sidecar is missing"))?;
        let mut alias = sidecar;
        alias.archive_info.identifier = Some(ALIASED_SIDECAR_ID);
        archive.objects.push(alias);
        Ok(())
    })
}

fn append_varint_unchecked(output: &mut Vec<u8>, field_number: u32, value: u64) {
    output.extend_from_slice(&encode_varint(u64::from(field_number) << 3));
    output.extend_from_slice(&encode_varint(value));
}

const UNKNOWN_EXTENSION_FIELD: u32 = 94;

fn first_format_payload(source: &[u8]) -> FixtureResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let entry = view
        .fields()
        .find(|field| field.number() == 3)
        .ok_or_else(|| io::Error::other("format entry field is missing"))?;
    WireView::parse(entry.payload())?
        .fields()
        .find(|field| field.number() == 6)
        .map(|field| field.payload().to_vec())
        .ok_or_else(|| io::Error::other("format payload field is missing").into())
}

fn replace_first_format_payload(source: &[u8], replacement: &[u8]) -> FixtureResult<Vec<u8>> {
    rewrite_format_payload_by_key_from_payload(source, FIRST_FORMAT_KEY, replacement)
}

fn rewrite_format_payload_by_key_from_payload(
    source: &[u8],
    key: u32,
    replacement: &[u8],
) -> FixtureResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::with_capacity(source.len());
    let mut replaced = false;
    for field in view.fields() {
        if field.number() != 3 || replaced {
            output.extend_from_slice(field.raw());
            continue;
        }
        let entry = tst::table_data_list::ListEntry::decode(field.payload())?;
        if entry.key != key {
            output.extend_from_slice(field.raw());
            continue;
        }
        let entry_view = WireView::parse(field.payload())?;
        let mut entry_output = Vec::with_capacity(field.payload().len());
        let mut format_found = false;
        for entry_field in entry_view.fields() {
            if entry_field.number() == 6 && !format_found {
                append_length_delimited_field(&mut entry_output, 6, replacement)?;
                format_found = true;
            } else {
                entry_output.extend_from_slice(entry_field.raw());
            }
        }
        if !format_found {
            return Err(io::Error::other("format payload field is missing").into());
        }
        append_length_delimited_field(&mut output, 3, &entry_output)?;
        replaced = true;
    }
    if !replaced {
        return Err(io::Error::other("format entry field is missing").into());
    }
    Ok(output)
}

fn pack_row(cells: &[Vec<u8>]) -> FixtureResult<(Vec<u8>, Vec<u8>)> {
    let mut storage = Vec::new();
    let mut offsets = Vec::new();
    offsets
        .try_reserve_exact(cells.len().saturating_mul(2))
        .map_err(|_| io::Error::other("format fixture offsets allocation failed"))?;
    for cell in cells {
        let offset = u16::try_from(storage.len())
            .map_err(|_| io::Error::other("format fixture row exceeds narrow offsets"))?;
        offsets.extend_from_slice(&offset.to_le_bytes());
        storage.extend_from_slice(cell);
    }
    Ok((storage, offsets))
}

fn unpack_row(row: &tst::TileRowInfo) -> FixtureResult<Vec<Vec<u8>>> {
    let storage = row
        .cell_storage_buffer
        .as_deref()
        .ok_or_else(|| io::Error::other("format fixture storage is missing"))?;
    let offsets = row
        .cell_offsets
        .as_deref()
        .ok_or_else(|| io::Error::other("format fixture offsets are missing"))?;
    let count = usize::try_from(row.cell_count)
        .map_err(|_| io::Error::other("format fixture cell count overflows usize"))?;
    if offsets.len() != count.saturating_mul(2) {
        return Err(io::Error::other("format fixture offsets have an invalid length").into());
    }
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(count)
        .map_err(|_| io::Error::other("format fixture cells allocation failed"))?;
    for index in 0..count {
        let begin = usize::from(u16::from_le_bytes([
            offsets[index * 2],
            offsets[index * 2 + 1],
        ]));
        let end = if index + 1 == count {
            storage.len()
        } else {
            usize::from(u16::from_le_bytes([
                offsets[(index + 1) * 2],
                offsets[(index + 1) * 2 + 1],
            ]))
        };
        if begin > end || end > storage.len() {
            return Err(io::Error::other("format fixture cell offset is out of range").into());
        }
        cells.push(storage[begin..end].to_vec());
    }
    Ok(cells)
}
