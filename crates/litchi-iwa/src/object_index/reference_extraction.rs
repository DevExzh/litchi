//! Compatibility reference extraction for legacy IWA payloads.

use crate::Result;
use crate::archive::ArchiveObject;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::varint::{decode_varint_from_bytes, encoded_len};
use litchi_iwa_index::{IndexBuilder, ObjectId};
use litchi_iwa_protos::{
    comment_storage_codec, drawable_container_codec, drawable_parent_codec, keynote_show_codec,
    numbers_table_cell_storage_codec,
};

use super::{add_reference_if_absent, index_error};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const COMMENT_STORAGE_CODEC_RECURSION_LIMIT: u32 = 64;
const KEYNOTE_SHOW_CODEC_RECURSION_LIMIT: u32 = 64;
const TSWP_STORAGE_STYLE_SHEET_FIELD: u32 = 2;
const TSP_REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const TSP_REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const TSP_REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const TSCH_CHART_MEDIATOR_INFO_FIELD: u32 = 1;

fn comment_storage_decode_options(source: &[u8]) -> comment_storage_codec::DecodeOptions {
    comment_storage_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        COMMENT_STORAGE_CODEC_RECURSION_LIMIT,
        source.len().clamp(1, litchi_numbers::MAX_REFERENCES),
        source
            .len()
            .clamp(1, litchi_numbers::DEFAULT_MAX_TEXT_BYTES),
    )
}

const TST_STORAGE_CODEC_RECURSION_LIMIT: u32 = 64;

fn tst_storage_decode_options(source: &[u8]) -> numbers_table_cell_storage_codec::DecodeOptions {
    numbers_table_cell_storage_codec::DecodeOptions::new(
        source.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        TST_STORAGE_CODEC_RECURSION_LIMIT,
        source.len().clamp(1, litchi_numbers::MAX_REFERENCES),
        source
            .len()
            .clamp(1, litchi_numbers::DEFAULT_MAX_TEXT_BYTES),
    )
}

fn keynote_show_decode_options(source: &[u8]) -> keynote_show_codec::DecodeOptions {
    let source_bytes = source.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    keynote_show_codec::DecodeOptions::new(
        source_bytes,
        source_bytes,
        KEYNOTE_SHOW_CODEC_RECURSION_LIMIT,
    )
    .with_max_fields(source_bytes.min(WireLimits::MAX_FIELDS))
    .with_max_work_bytes(
        source_bytes
            .saturating_mul(8)
            .min(WireLimits::MAX_REWRITE_WORK),
    )
}

/// Establish a finite parser profile from the exact caller-owned payload.
///
/// The compatibility extractor does not rewrite this message, so unknown
/// fields stay byte-authoritative. The profile only bounds the structural
/// field vector and never asks a generated decoder to materialize the whole
/// `TSWP.StorageArchive`.
fn storage_reference_wire_limits(source: &[u8]) -> Result<WireLimits> {
    let source_bytes = source.len().max(1);
    WireLimits::default()
        .with_input_bytes(source_bytes)
        .and_then(|limits| limits.with_fields(source_bytes.min(WireLimits::MAX_FIELDS)))
        .map_err(Into::into)
}

#[derive(Debug, Clone, Copy)]
struct StorageWireField {
    number: u32,
    wire_type: u8,
    start: usize,
    key_end: usize,
    payload_start: usize,
    end: usize,
}

impl StorageWireField {
    const fn number(self) -> u32 {
        self.number
    }

    const fn wire_type(self) -> u8 {
        self.wire_type
    }
}

fn storage_wire_error(message: impl Into<String>) -> crate::Error {
    crate::Error::InvalidFormat(message.into())
}

fn read_storage_varint(source: &[u8], offset: usize) -> Result<(u64, usize)> {
    let input = source
        .get(offset..)
        .ok_or_else(|| storage_wire_error("protobuf varint offset exceeds source"))?;
    let (value, width) = decode_varint_from_bytes(input).map_err(|error| {
        storage_wire_error(format!("invalid protobuf varint at byte {offset}: {error}"))
    })?;
    let end = offset
        .checked_add(width)
        .ok_or_else(|| storage_wire_error("protobuf varint offset overflows usize"))?;
    if end > source.len() {
        return Err(storage_wire_error("protobuf varint extends beyond source"));
    }
    Ok((value, width))
}

fn charge_storage_field(count: &mut usize, limits: WireLimits) -> Result<()> {
    let observed = count.saturating_add(1);
    if observed > limits.max_fields() {
        return Err(crate::Error::IwaCommon(
            litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed,
                limit: limits.max_fields(),
            },
        ));
    }
    *count = observed;
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "The bounded scanner carries source, group state, limits, and its output accumulator explicitly."
)]
fn scan_storage_fields(
    source: &[u8],
    start: usize,
    end: usize,
    expected_group: Option<u32>,
    depth: usize,
    limits: WireLimits,
    collect: bool,
    fields: &mut Vec<StorageWireField>,
    count: &mut usize,
) -> Result<usize> {
    let mut offset = start;
    while offset < end {
        let field_start = offset;
        let (key, key_width) = read_storage_varint(source, offset)?;
        offset = offset
            .checked_add(key_width)
            .ok_or_else(|| storage_wire_error("protobuf field key offset overflows usize"))?;
        if offset > end {
            return Err(storage_wire_error(
                "protobuf field key extends beyond its containing message",
            ));
        }
        let number = u32::try_from(key >> 3)
            .map_err(|_| storage_wire_error("protobuf field number exceeds u32"))?;
        if number == 0 || number > 0x1fff_ffff {
            return Err(storage_wire_error(format!(
                "invalid protobuf field number {number}"
            )));
        }
        let wire_type = u8::try_from(key & 7)
            .map_err(|_| storage_wire_error("protobuf wire type exceeds u8"))?;
        charge_storage_field(count, limits)?;
        let key_end = offset;

        if wire_type == 4 {
            if expected_group == Some(number) {
                return Ok(offset);
            }
            return Err(storage_wire_error("unexpected protobuf end-group"));
        }
        if wire_type == 3 {
            if depth >= limits.max_nesting() {
                return Err(crate::Error::IwaCommon(
                    litchi_iwa_common::Error::LimitExceeded {
                        kind: litchi_iwa_common::LimitKind::Nesting,
                        observed: depth.saturating_add(1),
                        limit: limits.max_nesting(),
                    },
                ));
            }
            let group_end = scan_storage_fields(
                source,
                offset,
                end,
                Some(number),
                depth.saturating_add(1),
                limits,
                false,
                fields,
                count,
            )?;
            if collect {
                fields.try_reserve(1).map_err(|_| {
                    crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                        resource: "IWA TSWP reference extraction fields",
                        amount: fields.len().saturating_add(1),
                    })
                })?;
                fields.push(StorageWireField {
                    number,
                    wire_type,
                    start: field_start,
                    key_end,
                    payload_start: key_end,
                    end: group_end,
                });
            }
            offset = group_end;
            continue;
        }

        let mut payload_start = offset;
        let field_end = match wire_type {
            0 => {
                let (_, width) = read_storage_varint(source, offset)?;
                offset
                    .checked_add(width)
                    .ok_or_else(|| storage_wire_error("protobuf varint offset overflows usize"))?
            },
            1 => offset
                .checked_add(8)
                .ok_or_else(|| storage_wire_error("protobuf fixed64 offset overflows usize"))?,
            2 => {
                let (length, width) = read_storage_varint(source, offset)?;
                let payload = offset.checked_add(width).ok_or_else(|| {
                    storage_wire_error("protobuf length prefix offset overflows usize")
                })?;
                payload_start = payload;
                let length = usize::try_from(length)
                    .map_err(|_| storage_wire_error("protobuf field length exceeds usize"))?;
                payload.checked_add(length).ok_or_else(|| {
                    storage_wire_error("protobuf length-delimited field overflows usize")
                })?
            },
            5 => offset
                .checked_add(4)
                .ok_or_else(|| storage_wire_error("protobuf fixed32 offset overflows usize"))?,
            _ => {
                return Err(storage_wire_error(format!(
                    "invalid protobuf wire type {wire_type}"
                )));
            },
        };
        if field_end > end {
            return Err(storage_wire_error(
                "protobuf field extends beyond its containing message",
            ));
        }
        offset = field_end;
        if collect {
            fields.try_reserve(1).map_err(|_| {
                crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                    resource: "IWA TSWP reference extraction fields",
                    amount: fields.len().saturating_add(1),
                })
            })?;
            fields.push(StorageWireField {
                number,
                wire_type,
                start: field_start,
                key_end,
                payload_start,
                end: field_end,
            });
        }
    }
    if expected_group.is_some() {
        return Err(storage_wire_error(
            "protobuf group is missing its end-group",
        ));
    }
    Ok(offset)
}

fn parse_storage_fields(source: &[u8], limits: WireLimits) -> Result<Vec<StorageWireField>> {
    if source.len() > limits.max_input_bytes() {
        return Err(crate::Error::IwaCommon(
            litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::InputBytes,
                observed: source.len(),
                limit: limits.max_input_bytes(),
            },
        ));
    }
    let mut fields = Vec::new();
    let mut count = 0;
    let parsed = scan_storage_fields(
        source,
        0,
        source.len(),
        None,
        0,
        limits,
        true,
        &mut fields,
        &mut count,
    )?;
    if parsed != source.len() {
        return Err(storage_wire_error(
            "protobuf scanner did not consume source",
        ));
    }
    Ok(fields)
}

fn storage_field_payload(source: &[u8], field: StorageWireField) -> Result<&[u8]> {
    if field.start > field.key_end
        || field.key_end > field.payload_start
        || field.payload_start > field.end
    {
        return Err(storage_wire_error(
            "protobuf field has invalid byte offsets",
        ));
    }
    source
        .get(field.payload_start..field.end)
        .ok_or_else(|| storage_wire_error("protobuf field extends beyond source"))
}

fn validate_storage_field_key(source: &[u8], field: StorageWireField) -> Result<()> {
    let key = source
        .get(field.start..field.key_end)
        .ok_or_else(|| storage_wire_error("protobuf field key extends beyond source"))?;
    let (encoded, width) = decode_varint_from_bytes(key)
        .map_err(|error| storage_wire_error(format!("invalid protobuf field key: {error}")))?;
    let expected = (u64::from(field.number) << 3) | u64::from(field.wire_type);
    if encoded != expected || width != key.len() || width != encoded_len(expected) {
        return Err(storage_wire_error(format!(
            "protobuf field {} has a noncanonical key",
            field.number
        )));
    }
    Ok(())
}

fn validate_storage_field_framing(source: &[u8], field: StorageWireField) -> Result<()> {
    validate_storage_field_key(source, field)?;
    if field.wire_type() != 2 {
        return Ok(());
    }
    let length_prefix = source
        .get(field.key_end..field.payload_start)
        .ok_or_else(|| storage_wire_error("protobuf length prefix extends beyond source"))?;
    let payload = storage_field_payload(source, field)?;
    let (length, width) = decode_varint_from_bytes(length_prefix)
        .map_err(|error| storage_wire_error(format!("invalid protobuf length prefix: {error}")))?;
    let payload_len = u64::try_from(payload.len())
        .map_err(|_| storage_wire_error("protobuf payload length exceeds u64"))?;
    if length != payload_len || width != length_prefix.len() || width != encoded_len(length) {
        return Err(storage_wire_error(format!(
            "protobuf field {} has a noncanonical length prefix",
            field.number
        )));
    }
    Ok(())
}

fn canonical_reference_varint(source: &[u8], field: StorageWireField, name: &str) -> Result<u64> {
    if field.wire_type() != 0 {
        return Err(crate::Error::InvalidFormat(format!(
            "{name} is not a varint"
        )));
    }
    validate_storage_field_key(source, field)?;
    let payload = storage_field_payload(source, field)?;
    let (value, width) = decode_varint_from_bytes(payload).map_err(|error| {
        crate::Error::InvalidFormat(format!("{name} contains an invalid varint: {error}"))
    })?;
    if width != payload.len() || width != encoded_len(value) {
        return Err(crate::Error::InvalidFormat(format!(
            "{name} contains a noncanonical varint"
        )));
    }
    Ok(value)
}

/// Read only the required identifier from one schema-directed `TSP.Reference`.
///
/// Known fields use strict canonical framing and duplicate checks. Unknown
/// fields are structurally scanned but otherwise ignored, preserving the raw
/// source as the authority and matching the compatibility extractor's
/// unknown-field behavior.
fn decode_reference_identifier(source: &[u8]) -> Result<u64> {
    let fields = parse_storage_fields(source, storage_reference_wire_limits(source)?)?;
    let mut identifier = None;
    let mut deprecated_type = false;
    let mut deprecated_external = false;
    for field in fields {
        match field.number() {
            TSP_REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(crate::Error::InvalidFormat(
                        "TSP.Reference.identifier is duplicated".to_owned(),
                    ));
                }
                identifier = Some(canonical_reference_varint(
                    source,
                    field,
                    "TSP.Reference.identifier",
                )?);
            },
            TSP_REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type {
                    return Err(crate::Error::InvalidFormat(
                        "TSP.Reference.deprecated_type is duplicated".to_owned(),
                    ));
                }
                let _ = canonical_reference_varint(source, field, "TSP.Reference.deprecated_type")?;
                deprecated_type = true;
            },
            TSP_REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_external {
                    return Err(crate::Error::InvalidFormat(
                        "TSP.Reference.deprecated_is_external is duplicated".to_owned(),
                    ));
                }
                let value = canonical_reference_varint(
                    source,
                    field,
                    "TSP.Reference.deprecated_is_external",
                )?;
                if value > 1 {
                    return Err(crate::Error::InvalidFormat(
                        "TSP.Reference.deprecated_is_external is not a canonical bool".to_owned(),
                    ));
                }
                deprecated_external = true;
            },
            _ => {},
        }
    }
    identifier.ok_or_else(|| {
        crate::Error::InvalidFormat("TSP.Reference is missing its required identifier".to_owned())
    })
}

/// Extract the stylesheet edge from `TSWP.StorageArchive` without a generated
/// Prost allocation. A malformed payload is ignored by this compatibility
/// path, but no edge is published until both the complete root scan and its
/// nested reference scan have succeeded.
fn extract_tswp_storage_reference(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    source: &[u8],
) -> Result<()> {
    let Ok(limits) = storage_reference_wire_limits(source) else {
        return Ok(());
    };
    let Ok(fields) = parse_storage_fields(source, limits) else {
        return Ok(());
    };
    let mut style_sheet = None;
    for field in fields {
        if field.number() != TSWP_STORAGE_STYLE_SHEET_FIELD {
            continue;
        }
        if style_sheet.is_some() || field.wire_type() != 2 {
            return Ok(());
        }
        if validate_storage_field_framing(source, field).is_err() {
            return Ok(());
        }
        let Ok(payload) = storage_field_payload(source, field) else {
            return Ok(());
        };
        let Ok(identifier) = decode_reference_identifier(payload) else {
            return Ok(());
        };
        style_sheet = Some(identifier);
    }
    if let Some(identifier) = style_sheet
        && let Some(target_id) = ObjectId::new(identifier)
    {
        add_reference_if_absent(builder, source_id, target_id)?;
    }
    Ok(())
}

#[derive(Debug, Default)]
struct CommentStorageReferences {
    replies: Vec<u64>,
    allocation_failed: Option<usize>,
}

impl CommentStorageReferences {
    fn into_replies(self) -> Result<Vec<u64>> {
        self.allocation_failed.map_or(Ok(self.replies), |amount| {
            Err(crate::Error::IwaCommon(
                litchi_iwa_common::Error::Allocation {
                    resource: "IWA comment reference extraction replies",
                    amount,
                },
            ))
        })
    }
}

impl comment_storage_codec::CommentStorageVisitor for CommentStorageReferences {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), comment_storage_codec::DecodeError> {
        if self.allocation_failed.is_some() {
            return Ok(());
        }
        if self.replies.try_reserve(1).is_err() {
            // Keep strict traversal going so malformed input still suppresses
            // every staged edge and wins over this candidate-local failure.
            self.allocation_failed = Some(self.replies.len().saturating_add(1));
            return Ok(());
        }
        self.replies.push(reply.identifier());
        Ok(())
    }
}

#[derive(Debug, Default)]
struct TstReferences {
    references: Vec<u64>,
    allocation_failed: Option<usize>,
}

impl TstReferences {
    fn push(&mut self, identifier: u64) {
        if self.allocation_failed.is_some() {
            return;
        }
        if self.references.try_reserve(1).is_err() {
            // Keep strict traversal going so malformed input still suppresses
            // every staged edge and wins over this candidate-local failure.
            self.allocation_failed = Some(self.references.len().saturating_add(1));
            return;
        }
        self.references.push(identifier);
    }

    fn into_references(self) -> Result<Vec<u64>> {
        self.allocation_failed
            .map_or(Ok(self.references), |amount| {
                Err(crate::Error::IwaCommon(
                    litchi_iwa_common::Error::Allocation {
                        resource: "IWA table reference extraction references",
                        amount,
                    },
                ))
            })
    }
}

#[derive(Debug, Default)]
struct TstListReferences {
    segments: TstReferences,
    entries: TstReferences,
}

impl TstListReferences {
    fn into_references(self) -> Result<Vec<u64>> {
        let mut segments = self.segments.into_references()?;
        let mut entries = self.entries.into_references()?;
        segments.try_reserve(entries.len()).map_err(|_error| {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "IWA table reference extraction references",
                amount: segments.len().saturating_add(entries.len()),
            })
        })?;
        segments.append(&mut entries);
        Ok(segments)
    }
}

impl numbers_table_cell_storage_codec::StorageVisitor for TstListReferences {
    fn visit_list_entry(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        for reference in [
            entry.reference(),
            entry.rich_text_payload(),
            entry.comment_storage(),
        ]
        .into_iter()
        .flatten()
        {
            self.entries.push(reference.identifier());
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        reference: numbers_table_cell_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), numbers_table_cell_storage_codec::DecodeError> {
        self.segments.push(reference.reference().identifier());
        Ok(())
    }
}

fn publish_tst_references(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    references: Vec<u64>,
) -> Result<()> {
    for identifier in references {
        if let Some(target_id) = ObjectId::new(identifier) {
            add_reference_if_absent(builder, source_id, target_id)?;
        }
    }
    Ok(())
}

fn publish_tst_list_references(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    references: Vec<u64>,
) -> Result<()> {
    // TST list segments precede entries in the legacy extractor even though
    // both are repeated protobuf fields. Keep that source-specific order in
    // the neutral graph while retaining its normal edge deduplication.
    builder
        .preserve_reference_order(source_id)
        .map_err(index_error)?;
    publish_tst_references(source_id, builder, references)
}

fn extract_tst_table_model_references(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    source: &[u8],
) -> Result<()> {
    let options = tst_storage_decode_options(source);
    let Ok((projection, _report)) =
        numbers_table_cell_storage_codec::decode_table_model_with_data_store_and_visitor(
            source,
            options,
            &mut (),
        )
    else {
        // Compatibility reference extraction has always ignored malformed
        // payloads. No edge is published until both the model and its
        // selected data-store projection have completed.
        return Ok(());
    };
    let table = projection.model();
    let data_store = projection.data_store();

    let mut staged = TstReferences::default();
    for reference in [
        table.table_style(),
        table.body_text_style(),
        table.header_row_text_style(),
        table.header_column_text_style(),
        table.footer_row_text_style(),
        table.body_cell_style(),
        table.header_row_style(),
        table.header_column_style(),
        table.footer_row_style(),
        table.table_name_style(),
        table.table_name_shape_style(),
    ]
    .into_iter()
    .flatten()
    {
        staged.push(reference.identifier());
    }
    for reference in [
        Some(data_store.column_headers()),
        Some(data_store.string_table()),
        Some(data_store.style_table()),
        Some(data_store.formula_table()),
        Some(data_store.format_table_pre_bnc()),
        data_store.format_table(),
        data_store.formula_error_table(),
        data_store.multiple_choice_list_format_table(),
        data_store.merge_region_map(),
    ]
    .into_iter()
    .flatten()
    {
        staged.push(reference.identifier());
    }
    publish_tst_references(source_id, builder, staged.into_references()?)
}

fn extract_tst_table_data_list_references(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    source: &[u8],
) -> Result<()> {
    let mut staged = TstListReferences::default();
    let Ok((_list, report)) = numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
        source,
        tst_storage_decode_options(source),
        &mut staged,
    ) else {
        // Do not publish segment or entry edges visited before a later
        // malformed field, duplicate, or reference parity failure.
        return Ok(());
    };
    let references = staged.into_references()?;
    if references.len() != report.references() {
        return Ok(());
    }
    // `into_references` retains the legacy segment-before-entry traversal
    // order while preserving duplicate edge idempotence.
    publish_tst_list_references(source_id, builder, references)
}

fn extract_tst_table_data_list_segment_references(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    source: &[u8],
) -> Result<()> {
    let mut staged = TstListReferences::default();
    let Ok((_segment, report)) =
        numbers_table_cell_storage_codec::decode_table_data_list_segment_with_visitor(
            source,
            tst_storage_decode_options(source),
            &mut staged,
        )
    else {
        return Ok(());
    };
    let references = staged.into_references()?;
    if references.len() != report.references() {
        return Ok(());
    }
    publish_tst_list_references(source_id, builder, references)
}

fn extract_comment_storage_references(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    source: &[u8],
) -> Result<()> {
    let mut staged = CommentStorageReferences::default();
    let Ok((comment, report)) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        source,
        comment_storage_decode_options(source),
        &mut staged,
    ) else {
        // Compatibility reference extraction has always ignored malformed
        // payloads. In particular, do not publish replies visited before a
        // later malformed field or parity failure.
        return Ok(());
    };
    let replies = staged.into_replies()?;
    if replies.len() != report.reply_references() {
        return Ok(());
    }

    // Only publish after the complete strict decode has succeeded. The
    // borrowed codec leaves unknown source fields untouched and the builder
    // sees the same source-order author/reply edges as the legacy decoder.
    if let Some(author) = comment.author()
        && let Some(target_id) = ObjectId::new(author.identifier())
    {
        add_reference_if_absent(builder, source_id, target_id)?;
    }
    for identifier in replies {
        if let Some(target_id) = ObjectId::new(identifier) {
            add_reference_if_absent(builder, source_id, target_id)?;
        }
    }
    Ok(())
}

/// Extract the direct reference edges from one Keynote show through the
/// bounded generated-free projection.
///
/// The show codec strictly validates the complete known envelope before its
/// private Buffa view is forced. The returned snapshot owns only scalar
/// identifiers; the original message bytes stay in the archive as the
/// preservation representation, including unknown fields.
fn extract_keynote_show_references(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    source: &[u8],
) -> Result<()> {
    let Ok(show) =
        keynote_show_codec::decode_references(source, keynote_show_decode_options(source))
    else {
        // This compatibility fallback historically ignored any payload that
        // Prost could not decode. Keep the same candidate-atomic behavior for
        // malformed or context-conflicting low-numbered messages.
        return Ok(());
    };

    // Preserve the legacy edge order: theme, stylesheet, UI state, recording.
    for identifier in [
        Some(show.theme_identifier()),
        Some(show.stylesheet_identifier()),
        show.ui_state_identifier(),
        show.recording_identifier(),
    ]
    .into_iter()
    .flatten()
    {
        if let Some(target_id) = ObjectId::new(identifier) {
            add_reference_if_absent(builder, source_id, target_id)?;
        }
    }
    Ok(())
}

/// Extract the optional `info` edge from `TSCH.ChartMediatorArchive` without
/// materializing the generated chart mediator. The source bytes remain the
/// preservation representation: the bounded scanner validates the complete
/// envelope, while unknown fields stay opaque and therefore do not affect the
/// projected edge.
fn extract_tsch_chart_mediator_reference(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    source: &[u8],
) -> Result<()> {
    let Ok(limits) = storage_reference_wire_limits(source) else {
        return Ok(());
    };
    let Ok(fields) = parse_storage_fields(source, limits) else {
        // Compatibility extraction ignores malformed candidates, but it must
        // not publish an edge visited before a later malformed field.
        return Ok(());
    };

    let mut info = None;
    for field in fields {
        if field.number() != TSCH_CHART_MEDIATOR_INFO_FIELD {
            continue;
        }
        if info.is_some() || field.wire_type() != 2 {
            // `info` is singular in the schema. Treat duplicate or
            // wire-incompatible occurrences as a malformed candidate rather
            // than choosing one while silently dropping the other.
            return Ok(());
        }
        if validate_storage_field_framing(source, field).is_err() {
            return Ok(());
        }
        let Ok(payload) = storage_field_payload(source, field) else {
            return Ok(());
        };
        let Ok(identifier) = decode_reference_identifier(payload) else {
            return Ok(());
        };
        info = Some(identifier);
    }

    if let Some(identifier) = info
        && let Some(target_id) = ObjectId::new(identifier)
    {
        // Publication happens only after the root and nested reference scans
        // have both completed, preserving candidate atomicity.
        add_reference_if_absent(builder, source_id, target_id)?;
    }
    Ok(())
}

pub(super) fn extract(
    source_id: ObjectId,
    object: &ArchiveObject,
    builder: &mut IndexBuilder,
) -> Result<()> {
    use prost::Message;

    // For each raw message, try to extract references
    for raw_msg in &object.messages {
        let msg_type = raw_msg.type_;

        // Extract references based on message type
        // We decode the specific protobuf message and extract its Reference fields
        match msg_type {
            // TST (Table) types
            6000 | 6001 => {
                extract_tst_table_model_references(source_id, builder, &raw_msg.data)?;
            },

            6005 | 6201 => {
                extract_tst_table_data_list_references(source_id, builder, &raw_msg.data)?;
            },

            6011 => {
                extract_tst_table_data_list_segment_references(source_id, builder, &raw_msg.data)?;
            },

            // TSWP (Word Processing/Text) types
            2001..=2022 => {
                // Only the schema-directed stylesheet edge is needed here;
                // text and attribute tables remain opaque source bytes.
                extract_tswp_storage_reference(source_id, builder, &raw_msg.data)?;
            },

            // KN (Keynote) types
            5 | 6 => {
                // KN.SlideArchive contains references to drawables, builds, and transitions
                if let Ok(slide) = crate::protobuf::kn::SlideArchive::decode(&*raw_msg.data) {
                    // Extract style reference
                    extract_reference(source_id, builder, &slide.style)?;

                    // Extract drawable references (shapes, images, text boxes)
                    for drawable in &slide.owned_drawables {
                        extract_reference(source_id, builder, drawable)?;
                    }

                    // Extract build animation references
                    for build in &slide.builds {
                        extract_reference(source_id, builder, build)?;
                    }

                    // Extract placeholder references
                    if let Some(ref title) = slide.title_placeholder {
                        extract_reference(source_id, builder, title)?;
                    }
                    if let Some(ref body) = slide.body_placeholder {
                        extract_reference(source_id, builder, body)?;
                    }
                    if let Some(ref object) = slide.object_placeholder {
                        extract_reference(source_id, builder, object)?;
                    }
                    if let Some(ref slide_num) = slide.slide_number_placeholder {
                        extract_reference(source_id, builder, slide_num)?;
                    }

                    // Extract style references
                    for para_style in &slide.body_paragraph_styles {
                        extract_reference(source_id, builder, para_style)?;
                    }
                    for list_style in &slide.body_list_styles {
                        extract_reference(source_id, builder, list_style)?;
                    }
                }
            },

            2 => {
                // KN.ShowArchive (conflicts with TSP.MessageInfo, handle by context)
                extract_keynote_show_references(source_id, builder, &raw_msg.data)?;
            },

            // TN (Numbers) types
            3 => {
                // TN.SheetArchive / TN.FormBasedSheetArchive
                if let Ok(sheet) = crate::protobuf::tn::SheetArchive::decode(&*raw_msg.data) {
                    // Extract drawable info references
                    for drawable_ref in &sheet.drawable_infos {
                        extract_reference(source_id, builder, drawable_ref)?;
                    }

                    for header in &sheet.headers {
                        extract_reference(source_id, builder, header)?;
                    }
                    for footer in &sheet.footers {
                        extract_reference(source_id, builder, footer)?;
                    }

                    // Old documents used one storage reference for each area.
                    if sheet.headers.is_empty() && sheet.footers.is_empty() {
                        extract_legacy_sheet_headers(source_id, builder, &sheet)?;
                    }
                }
            },

            // TSD (Drawing/Shape) types
            // Implementation Status: ✓ COMPLETED (2025-11-04)
            // Based on TSDArchives.proto and libetonyek's reference extraction
            3002 => {
                // TSD.DrawableArchive - base type for all drawables
                if let Ok(drawable) = drawable_parent_codec::decode_parent(
                    &raw_msg.data,
                    drawable_parent_codec::DecodeOptions::for_source(&raw_msg.data),
                ) {
                    // Extract the parent edge without materializing the full
                    // drawable graph. The selected parent identifier is a
                    // borrow-free scalar; the source payload remains the
                    // preservation authority.
                    if let Some(parent) = drawable.parent_identifier()
                        && let Some(target_id) = ObjectId::new(parent.get())
                    {
                        add_reference_if_absent(builder, source_id, target_id)?;
                    }
                }
            },
            3003 => {
                if let Ok(container) = drawable_container_codec::decode_container_references(
                    &raw_msg.data,
                    drawable_container_codec::DecodeOptions::for_source(&raw_msg.data),
                ) {
                    // The complete candidate is checked before its scalar
                    // edges are published; geometry remains borrowed/opaque.
                    for reference in container.references() {
                        if let Some(target_id) = ObjectId::new(reference.get()) {
                            add_reference_if_absent(builder, source_id, target_id)?;
                        }
                    }
                }
            },
            3004 => {
                // TSD.ShapeArchive - shapes (rectangles, circles, polygons, etc.)
                if let Ok(shape) = crate::protobuf::tsd::ShapeArchive::decode(&*raw_msg.data) {
                    // ShapeArchive embeds DrawableArchive in 'super' field (required)
                    // Extract parent from the super DrawableArchive
                    if let Some(ref parent) = shape.super_.parent {
                        extract_reference(source_id, builder, parent)?;
                    }
                    // Extract style reference
                    if let Some(ref style) = shape.style {
                        extract_reference(source_id, builder, style)?;
                    }
                    // Note: pathsource, head_line_end, tail_line_end are not references
                    // but embedded data structures
                }
            },
            3005 => {
                // TSD.ImageArchive - images
                if let Ok(image) = crate::protobuf::tsd::ImageArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = image.super_.parent {
                        extract_reference(source_id, builder, parent)?;
                    }
                    // Extract style reference
                    if let Some(ref style) = image.style {
                        extract_reference(source_id, builder, style)?;
                    }
                    // Note: data field is a DataReference, not an object Reference
                    // database_originalData is also for media assets
                }
            },
            3006 => {
                // TSD.MaskArchive - image masks
                if let Ok(mask) = crate::protobuf::tsd::MaskArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = mask.super_.parent {
                        extract_reference(source_id, builder, parent)?;
                    }
                    // Note: pathsource is embedded data, not a reference
                }
            },
            3007 => {
                // TSD.MovieArchive - video objects
                if let Ok(movie) = crate::protobuf::tsd::MovieArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = movie.super_.parent {
                        extract_reference(source_id, builder, parent)?;
                    }
                    // Extract style reference
                    if let Some(ref style) = movie.style {
                        extract_reference(source_id, builder, style)?;
                    }
                    // Note: movieData is a DataReference, not an object Reference
                }
            },
            3008 => {
                if let Ok(group) = drawable_container_codec::decode_group_references(
                    &raw_msg.data,
                    drawable_container_codec::DecodeOptions::for_source(&raw_msg.data),
                ) {
                    for reference in group.references() {
                        if let Some(target_id) = ObjectId::new(reference.get()) {
                            add_reference_if_absent(builder, source_id, target_id)?;
                        }
                    }
                }
            },
            3009 => {
                // TSD.ConnectionLineArchive - connector lines between shapes
                if let Ok(conn_line) =
                    crate::protobuf::tsd::ConnectionLineArchive::decode(&*raw_msg.data)
                {
                    // Extract parent and style from super ShapeArchive (required field)
                    // ConnectionLineArchive.super_ is ShapeArchive
                    // ShapeArchive.super_ is DrawableArchive
                    if let Some(ref parent) = conn_line.super_.super_.parent {
                        extract_reference(source_id, builder, parent)?;
                    }
                    if let Some(ref style) = conn_line.super_.style {
                        extract_reference(source_id, builder, style)?;
                    }
                    // Extract connection endpoints
                    if let Some(ref connected_from) = conn_line.connected_from {
                        extract_reference(source_id, builder, connected_from)?;
                    }
                    if let Some(ref connected_to) = conn_line.connected_to {
                        extract_reference(source_id, builder, connected_to)?;
                    }
                }
            },
            COMMENT_STORAGE_MESSAGE_TYPE => {
                extract_comment_storage_references(source_id, builder, &raw_msg.data)?;
            },

            // TSCH (Chart) types
            // Implementation Status: ✓ COMPLETED (2025-11-04)
            // Based on TSCHArchives.proto and libetonyek's chart parsing
            5000 => {
                // TSCH.PreUFF.ChartInfoArchive - legacy chart format
                // This is a pre-unified format chart, structure may vary
                // Attempt basic reference extraction but may fail gracefully
                if let Ok(chart_info) =
                    crate::protobuf::tsch::pre_uff::ChartInfoArchive::decode(&*raw_msg.data)
                {
                    // Extract chart style reference if present
                    if let Some(ref style) = chart_info.style {
                        extract_reference(source_id, builder, style)?;
                    }
                    // Note: PreUFF ChartInfoArchive doesn't have a direct legend field
                    // Legend info is embedded in other structures
                }
            },
            5004 => {
                // TSCH.ChartMediatorArchive - mediator between chart and data
                // Extract info reference (points to the chart drawable) from
                // the bounded neutral wire projection. `local_series_indexes`
                // and `remote_series_indexes` are indices, not object refs.
                extract_tsch_chart_mediator_reference(source_id, builder, &raw_msg.data)?;
            },
            5020 => {
                // TSCH.ChartStylePreset - preset styles for charts
                if let Ok(preset) = crate::protobuf::tsch::ChartStylePreset::decode(&*raw_msg.data)
                {
                    // Extract chart style reference
                    if let Some(ref chart_style) = preset.chart_style {
                        extract_reference(source_id, builder, chart_style)?;
                    }
                    // Extract legend style reference
                    if let Some(ref legend_style) = preset.legend_style {
                        extract_reference(source_id, builder, legend_style)?;
                    }
                    // Note: ChartStylePreset has a complex nested structure
                    // Styles for series and axes are managed through different fields
                    // than what might be expected from the pre-UFF format
                }
            },
            5021 => {
                // TSCH.ChartDrawableArchive - main chart drawable
                if let Ok(chart_drawable) =
                    crate::protobuf::tsch::ChartDrawableArchive::decode(&*raw_msg.data)
                {
                    // Extract parent from super DrawableArchive
                    if let Some(ref drawable) = chart_drawable.super_
                        && let Some(ref parent) = drawable.parent
                    {
                        extract_reference(source_id, builder, parent)?;
                    }
                    // Note: ChartArchive is embedded via protobuf extensions,
                    // which requires special handling. The chart data and preset
                    // references would be in the extension fields that we can't
                    // easily access through the standard decode.
                }
            },

            // TP (Pages) types
            10000 => {
                // TP.DocumentArchive
                if let Ok(doc) = crate::protobuf::tp::DocumentArchive::decode(&*raw_msg.data) {
                    for reference in [
                        doc.stylesheet.as_ref(),
                        doc.floating_drawables.as_ref(),
                        doc.body_storage.as_ref(),
                        doc.section.as_ref(),
                        doc.theme.as_ref(),
                        doc.settings.as_ref(),
                        doc.deprecated_layout_state.as_ref(),
                        doc.deprecated_view_state.as_ref(),
                        doc.most_recent_change_session.as_ref(),
                        doc.drawables_zorder.as_ref(),
                        doc.tables_custom_format_list.as_ref(),
                        doc.flow_info_container.as_ref(),
                        doc.merge_data.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, builder, reference)?;
                    }
                    for reference in doc
                        .citation_records
                        .iter()
                        .chain(&doc.toc_styles)
                        .chain(&doc.change_sessions)
                        .chain(&doc.page_templates)
                    {
                        extract_reference(source_id, builder, reference)?;
                    }
                    let tsa = &doc.super_;
                    for reference in [
                        tsa.calculation_engine.as_ref(),
                        tsa.view_state.as_ref(),
                        tsa.function_browser_state.as_ref(),
                        tsa.tables_custom_format_list.as_ref(),
                        tsa.shortcut_controller.as_ref(),
                        tsa.annotation_cache_deprecated.as_ref(),
                        tsa.custom_format_list.as_ref(),
                        tsa.annotation_cache_deprecated_2.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, builder, reference)?;
                    }
                    let tsk = &tsa.super_;
                    for reference in [
                        tsk.annotation_author_storage.as_ref(),
                        tsk.collaboration_operation_history.as_ref(),
                        tsk.activity_stream.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, builder, reference)?;
                    }
                    for reference in &tsk.activity_log_entries {
                        extract_reference(source_id, builder, reference)?;
                    }
                }
            },

            10011 => {
                if let Ok(section) = crate::protobuf::tp::SectionArchive::decode(&*raw_msg.data) {
                    for reference in section
                        .obsolete_headers
                        .iter()
                        .chain(&section.obsolete_footers)
                        .chain(&section.obsolete_section_template_drawables)
                    {
                        extract_reference(source_id, builder, reference)?;
                    }
                    for reference in [
                        section.first_section_template_page.as_ref(),
                        section.even_section_template_page.as_ref(),
                        section.odd_section_template_page.as_ref(),
                        section.user_defined_guide_storage.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, builder, reference)?;
                    }
                }
            },

            10143 => {
                if let Ok(template) =
                    crate::protobuf::tp::SectionTemplateArchive::decode(&*raw_msg.data)
                {
                    for reference in template
                        .headers
                        .iter()
                        .chain(&template.footers)
                        .chain(&template.section_template_drawables)
                    {
                        extract_reference(source_id, builder, reference)?;
                    }
                }
            },

            _ => {
                // For unknown types, we don't extract references
                // This is fine as we handle the most common types above
            },
        }
    }

    Ok(())
}

fn extract_reference(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    reference: &crate::protobuf::tsp::Reference,
) -> Result<()> {
    if let Some(target_id) = ObjectId::new(reference.identifier) {
        add_reference_if_absent(builder, source_id, target_id)?;
    }
    Ok(())
}

#[allow(deprecated)]
fn extract_legacy_sheet_headers(
    source_id: ObjectId,
    builder: &mut IndexBuilder,
    sheet: &crate::protobuf::tn::SheetArchive,
) -> Result<()> {
    if let Some(header) = &sheet.header_storage {
        extract_reference(source_id, builder, header)?;
    }
    if let Some(footer) = &sheet.footer_storage {
        extract_reference(source_id, builder, footer)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{ArchiveObject, RawMessage};
    use crate::protobuf::{kn, tsch, tsd, tsp, tst, tswp};
    use litchi_iwa_index::{ByteSpan, FragmentId, ObjectRecord};
    use prost::Message;
    use std::num::NonZeroU32;

    fn index_for_comment_payload(data: Vec<u8>) -> litchi_iwa_index::ObjectIndex {
        index_for_payload(COMMENT_STORAGE_MESSAGE_TYPE, data)
    }

    fn index_for_payload(message_type: u32, data: Vec<u8>) -> litchi_iwa_index::ObjectIndex {
        let source_id = ObjectId::new(10).expect("non-zero source");
        let fragment = FragmentId::new(NonZeroU32::new(1).expect("non-zero fragment"));
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_object(ObjectRecord::new(
                source_id,
                fragment,
                ByteSpan::new(0, data.len() as u64).expect("payload span"),
            ))
            .expect("source object");
        let object = ArchiveObject::new(
            source_id.get(),
            vec![RawMessage {
                type_: message_type,
                data,
            }],
        )
        .expect("archive object");
        extract(source_id, &object, &mut builder).expect("reference extraction");
        builder
            .build_allow_missing_targets()
            .expect("reference index")
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn outgoing(index: &litchi_iwa_index::ObjectIndex) -> Vec<ObjectId> {
        index
            .outgoing(ObjectId::new(10).expect("source"))
            .map(|targets| targets.collect())
            .unwrap_or_default()
    }

    fn keynote_show() -> kn::ShowArchive {
        kn::ShowArchive {
            ui_state: Some(reference(20)),
            theme: reference(21),
            slide_tree: kn::SlideTreeArchive {
                slides: vec![reference(22)],
                ..Default::default()
            },
            size: tsp::Size {
                width: 1_024.0,
                height: 768.0,
            },
            stylesheet: reference(23),
            recording: Some(reference(24)),
            ..Default::default()
        }
    }

    #[test]
    fn keynote_show_ingress_stays_on_scalar_wire_projection() {
        let source = include_str!("reference_extraction.rs");
        let production = source
            .split_once("#[cfg(test)]")
            .map_or(source, |(production, _tests)| production);
        assert!(!production.contains("ShowArchive::decode"));
        assert_eq!(
            production
                .matches("keynote_show_codec::decode_references(")
                .count(),
            1
        );
    }

    #[test]
    fn tst_table_model_ingress_uses_combined_model_store_projection() {
        let source = include_str!("reference_extraction.rs");
        let production = source
            .split_once("#[cfg(test)]")
            .map_or(source, |(production, _tests)| production);
        assert_eq!(
            production
                .matches(
                    "numbers_table_cell_storage_codec::decode_table_model_with_data_store_and_visitor("
                )
                .count(),
            1
        );
        assert!(
            !production
                .contains("numbers_table_cell_storage_codec::decode_table_model_with_report(")
        );
        assert!(
            !production
                .contains("numbers_table_cell_storage_codec::decode_data_store_with_report(")
        );
    }

    #[test]
    fn drawable_parent_ingress_stays_on_bounded_buffa_projection() {
        let source = include_str!("reference_extraction.rs");
        let production = source
            .split_once("#[cfg(test)]")
            .map_or(source, |(production, _tests)| production);
        assert_eq!(
            production
                .matches("drawable_parent_codec::decode_parent(")
                .count(),
            1
        );
        assert!(!production.contains("tsd::DrawableArchive::decode"));
    }

    #[test]
    fn drawable_parent_projection_preserves_unknowns_and_parent_edges() {
        let drawable = tsd::DrawableArchive {
            parent: Some(reference(20)),
            ..Default::default()
        };
        let mut data = drawable.encode_to_vec();
        // The strict borrowed projection ignores this unknown field while
        // retaining the original source bytes as the preservation authority.
        data.extend_from_slice(&[0x98, 0x06, 0x81, 0x01]);

        let index = index_for_payload(3_002, data);
        assert_eq!(outgoing(&index), vec![ObjectId::new(20).expect("parent")]);
    }

    #[test]
    fn malformed_drawable_parent_does_not_publish_a_staged_edge() {
        let drawable = tsd::DrawableArchive {
            parent: Some(reference(20)),
            ..Default::default()
        };
        let mut duplicate_parent = drawable.encode_to_vec();
        duplicate_parent.extend_from_slice(
            &tsd::DrawableArchive {
                parent: Some(reference(30)),
                ..Default::default()
            }
            .encode_to_vec(),
        );

        let index = index_for_payload(3_002, duplicate_parent);
        assert!(outgoing(&index).is_empty());
    }

    #[test]
    fn container_and_group_reference_projection_preserves_selected_edges() {
        let container = tsd::ContainerArchive {
            parent: Some(reference(20)),
            children: vec![reference(21), reference(0), reference(21), reference(22)],
            ..Default::default()
        };
        let group = tsd::GroupArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(20)),
                comment: Some(reference(90)),
                ..Default::default()
            },
            children: container.children.clone(),
            fake_shape_for_empty_group: Some(reference(91)),
        };
        for (message_type, mut data) in [
            (3_003, container.encode_to_vec()),
            (3_008, group.encode_to_vec()),
        ] {
            data.extend_from_slice(&[0x98, 0x06, 0x81, 0x01]);
            assert_eq!(
                outgoing(&index_for_payload(message_type, data)),
                [20, 21, 22]
                    .into_iter()
                    .map(|id| ObjectId::new(id).unwrap())
                    .collect::<Vec<_>>(),
                "parent/children only; nulls and duplicate edges keep compatibility semantics"
            );
        }
        let parentless_group = tsd::GroupArchive {
            super_: tsd::DrawableArchive::default(),
            children: vec![reference(21)],
            ..Default::default()
        };
        assert_eq!(
            outgoing(&index_for_payload(3_008, parentless_group.encode_to_vec())),
            vec![ObjectId::new(21).unwrap()],
            "a group requires its drawable envelope, but the parent is optional"
        );
    }

    #[test]
    fn malformed_container_and_group_publish_no_candidate_edges() {
        // A valid parent and child precede the malformed final child. The
        // candidate must be validated completely before either is published.
        for (message_type, valid) in [
            (
                3_003,
                tsd::ContainerArchive {
                    parent: Some(reference(20)),
                    children: vec![reference(21)],
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            (
                3_008,
                tsd::GroupArchive {
                    super_: tsd::DrawableArchive {
                        parent: Some(reference(20)),
                        ..Default::default()
                    },
                    children: vec![reference(21)],
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
        ] {
            let child_key = if message_type == 3_003 { 0x1a } else { 0x12 };
            let mut truncated = valid.clone();
            truncated.extend_from_slice(&[child_key, 2, 0x08]);
            assert!(outgoing(&index_for_payload(message_type, truncated)).is_empty());

            let mut missing_identifier = valid;
            missing_identifier.extend_from_slice(&[child_key, 0]);
            assert!(outgoing(&index_for_payload(message_type, missing_identifier)).is_empty());
        }
    }

    #[test]
    fn keynote_show_projection_preserves_direct_edges_and_unknowns() {
        let mut data = keynote_show().encode_to_vec();
        // The source remains the preservation representation; this unknown
        // field must be accepted without entering the projected snapshot.
        data.extend_from_slice(&[0x9a, 0x06, 0x01, 0x7f]);

        let index = index_for_payload(2, data);
        // The neutral graph exposes outgoing IDs in deterministic numeric
        // order, independent of source-field traversal order.
        assert_eq!(
            outgoing(&index),
            vec![
                ObjectId::new(20).expect("ui state"),
                ObjectId::new(21).expect("theme"),
                ObjectId::new(23).expect("stylesheet"),
                ObjectId::new(24).expect("recording"),
            ]
        );
    }

    #[test]
    fn malformed_keynote_show_projection_publishes_no_staged_edges() {
        let mut duplicate_theme = keynote_show().encode_to_vec();
        // A duplicate singular field is rejected by the strict projection
        // after the valid references have already been visited. No edge may
        // leak into the neutral graph from this candidate.
        duplicate_theme.extend_from_slice(&[0x12, 0x02, 0x08, 0x63]);

        let index = index_for_payload(2, duplicate_theme);
        assert!(outgoing(&index).is_empty());
    }

    #[test]
    fn strict_comment_storage_reference_extraction_matches_projection_and_unknowns() {
        let comment = tsd::CommentStorageArchive {
            author: Some(tsp::Reference {
                identifier: 20,
                ..Default::default()
            }),
            replies: vec![tsp::Reference {
                identifier: 30,
                ..Default::default()
            }],
            ..Default::default()
        };
        let mut data = comment.encode_to_vec();
        data.extend_from_slice(&[0x78, 0x01]);

        let index = index_for_comment_payload(data);
        assert_eq!(
            index
                .outgoing(ObjectId::new(10).expect("source"))
                .map(|targets| targets.collect::<Vec<_>>()),
            Some(vec![
                ObjectId::new(20).expect("author"),
                ObjectId::new(30).expect("reply"),
            ])
        );
    }

    #[test]
    fn malformed_comment_storage_does_not_publish_staged_references() {
        let comment = tsd::CommentStorageArchive {
            author: Some(tsp::Reference {
                identifier: 20,
                ..Default::default()
            }),
            replies: vec![tsp::Reference {
                identifier: 30,
                ..Default::default()
            }],
            ..Default::default()
        };
        let mut duplicate_author = comment.encode_to_vec();
        duplicate_author.extend_from_slice(
            &tsd::CommentStorageArchive {
                author: Some(tsp::Reference {
                    identifier: 99,
                    ..Default::default()
                }),
                ..Default::default()
            }
            .encode_to_vec(),
        );

        let index = index_for_comment_payload(duplicate_author);
        assert!(index.outgoing(ObjectId::new(10).expect("source")).is_none());
    }

    #[test]
    fn invalid_utf8_and_missing_reference_payloads_are_ignored_without_edges() {
        for data in [[0x0a, 0x01, 0xff].to_vec(), [0x1a, 0x00].to_vec()] {
            let index = index_for_comment_payload(data);
            assert!(index.outgoing(ObjectId::new(10).expect("source")).is_none());
        }
    }

    #[test]
    fn strict_tswp_storage_reference_preserves_unknowns_and_dangling_targets() {
        let storage = tswp::StorageArchive {
            style_sheet: Some(reference(20)),
            ..Default::default()
        };
        let mut data = storage.encode_to_vec();
        // Unknown source bytes stay outside the projection and the missing
        // target remains visible through the neutral graph's dangling edge.
        data.extend_from_slice(&[0x78, 0x01]);
        // A well-formed unknown group is also opaque to the projection.
        data.extend_from_slice(&[0xa3, 0x06, 0xa8, 0x06, 0x08, 0xa4, 0x06]);

        let index = index_for_payload(2001, data);
        assert_eq!(outgoing(&index), vec![ObjectId::new(20).expect("style")]);
        assert!(index.object(ObjectId::new(20).expect("style")).is_none());
    }

    #[test]
    fn malformed_tswp_storage_reference_does_not_publish_a_staged_edge() {
        let storage = tswp::StorageArchive {
            style_sheet: Some(reference(20)),
            ..Default::default()
        };
        let mut duplicate_style = storage.encode_to_vec();
        // A second singular stylesheet field is rejected after the first one
        // has been read; publication must remain candidate-atomic.
        duplicate_style.extend_from_slice(&[0x12, 0x02, 0x08, 0x01]);

        let index = index_for_payload(2001, duplicate_style);
        assert!(outgoing(&index).is_empty());
    }

    #[test]
    fn malformed_tswp_unknown_group_does_not_publish_a_staged_edge() {
        let storage = tswp::StorageArchive {
            style_sheet: Some(reference(20)),
            ..Default::default()
        };
        let mut malformed_group = storage.encode_to_vec();
        // The unknown group never closes, so the already scanned stylesheet
        // edge must remain unpublished.
        malformed_group.extend_from_slice(&[0xa3, 0x06, 0xa8, 0x06, 0x08]);

        let index = index_for_payload(2001, malformed_group);
        assert!(outgoing(&index).is_empty());
    }

    #[test]
    fn strict_tsch_chart_mediator_reference_preserves_unknowns_and_dangling_targets() {
        let mediator = tsch::ChartMediatorArchive {
            info: Some(reference(20)),
            local_series_indexes: vec![1, 2],
            remote_series_indexes: vec![3],
        };
        let mut data = mediator.encode_to_vec();
        // Unknown source bytes remain opaque to the scalar projection, while
        // the missing target remains visible as a dangling neutral edge.
        data.extend_from_slice(&[0x78, 0x01]);

        let index = index_for_payload(5004, data);
        assert_eq!(outgoing(&index), vec![ObjectId::new(20).expect("info")]);
        assert!(index.object(ObjectId::new(20).expect("info")).is_none());
    }

    #[test]
    fn malformed_tsch_chart_mediator_does_not_publish_a_staged_edge() {
        let mediator = tsch::ChartMediatorArchive {
            info: Some(reference(20)),
            ..Default::default()
        };
        let mut malformed = mediator.encode_to_vec();
        // An unterminated unknown group follows the valid info field. The
        // complete root scan must fail before the staged edge is published.
        malformed.extend_from_slice(&[0xa3, 0x06, 0xa8, 0x06, 0x08]);

        let index = index_for_payload(5004, malformed);
        assert!(outgoing(&index).is_empty());
    }

    #[test]
    fn strict_tst_list_extraction_preserves_segment_order_duplicates_and_missing_targets() {
        let list = tst::TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            next_list_id: 2,
            entries: vec![tst::table_data_list::ListEntry {
                key: 1,
                refcount: 1,
                reference: Some(reference(30)),
                rich_text_payload: Some(reference(20)),
                comment_storage: Some(reference(30)),
                ..Default::default()
            }],
            segments: vec![reference(20), reference(40), reference(20)],
            is_new_for_bnc: Some(true),
        };

        let index = index_for_payload(6005, list.encode_to_vec());
        assert_eq!(
            outgoing(&index),
            vec![
                ObjectId::new(20).expect("segment target"),
                ObjectId::new(40).expect("missing target"),
                ObjectId::new(30).expect("entry target"),
            ]
        );
    }

    #[test]
    fn malformed_tst_list_and_segment_do_not_publish_visited_edges() {
        let list = tst::TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            next_list_id: 2,
            segments: vec![reference(20)],
            ..Default::default()
        };
        let mut malformed_list = list.encode_to_vec();
        // A ListEntry with key but no required refcount follows a valid
        // segment; strict traversal must suppress the segment edge too.
        malformed_list.extend_from_slice(&[0x1a, 0x02, 0x08, 0x01]);
        let list_index = index_for_payload(6005, malformed_list);
        assert!(outgoing(&list_index).is_empty());

        let segment = tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            key_range: tsp::Range {
                location: 1,
                length: 1,
            },
            entries: vec![tst::table_data_list::ListEntry {
                key: 1,
                refcount: 1,
                rich_text_payload: Some(reference(30)),
                ..Default::default()
            }],
        };
        let mut malformed_segment = segment.encode_to_vec();
        malformed_segment.extend_from_slice(&[0x1a, 0x02, 0x08, 0x01]);
        let segment_index = index_for_payload(6011, malformed_segment);
        assert!(outgoing(&segment_index).is_empty());
    }

    #[test]
    fn strict_tst_table_extraction_matches_style_and_data_store_edges() {
        let store = tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                buckets: Vec::new(),
            },
            column_headers: reference(12),
            tiles: tst::TileStorage {
                tiles: Vec::new(),
                ..Default::default()
            },
            string_table: reference(13),
            style_table: reference(14),
            formula_table: reference(15),
            format_table_pre_bnc: reference(16),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree { nodes: Vec::new() },
            column_tile_tree: tst::TableRbTree { nodes: Vec::new() },
            ..Default::default()
        };
        let table = tst::TableModelArchive {
            table_id: "table".to_owned(),
            table_style: reference(20),
            body_text_style: reference(21),
            header_row_text_style: reference(22),
            header_column_text_style: reference(23),
            footer_row_text_style: reference(24),
            body_cell_style: reference(25),
            header_row_style: reference(26),
            header_column_style: reference(27),
            footer_row_style: reference(28),
            table_name_style: Some(reference(29)),
            table_name_shape_style: Some(reference(30)),
            base_data_store: store,
            number_of_rows: 1,
            number_of_columns: 1,
            table_name: "name".to_owned(),
            default_row_height: 0.0,
            default_column_width: 0.0,
            ..Default::default()
        };

        let index = index_for_payload(6000, table.encode_to_vec());
        let expected = (12..=16)
            .chain(20..=30)
            .map(|identifier| ObjectId::new(identifier).expect("non-zero target"))
            .collect::<Vec<_>>();
        assert_eq!(outgoing(&index), expected);
    }

    #[test]
    fn duplicate_tst_table_style_suppresses_all_staged_edges() {
        let store = tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                buckets: Vec::new(),
            },
            column_headers: reference(12),
            tiles: tst::TileStorage {
                tiles: Vec::new(),
                ..Default::default()
            },
            string_table: reference(13),
            style_table: reference(14),
            formula_table: reference(15),
            format_table_pre_bnc: reference(16),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree { nodes: Vec::new() },
            column_tile_tree: tst::TableRbTree { nodes: Vec::new() },
            ..Default::default()
        };
        let table = tst::TableModelArchive {
            table_id: "table".to_owned(),
            table_style: reference(20),
            base_data_store: store,
            number_of_rows: 1,
            number_of_columns: 1,
            ..Default::default()
        };
        let mut malformed = table.encode_to_vec();
        malformed.extend_from_slice(&[0x1a, 0x02, 0x08, 0x63]);
        let index = index_for_payload(6000, malformed);
        assert!(outgoing(&index).is_empty());
    }
}
