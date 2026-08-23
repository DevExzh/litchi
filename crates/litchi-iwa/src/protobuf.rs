//! Protobuf Message Support for iWork IWA Files
//!
//! This module provides support for decoding Protocol Buffers messages
//! used in iWork IWA (iWork Archive) files using the prost crate.

#![allow(
    dead_code,
    unused_imports,
    reason = "This private decoder adapter retains in-crate migration entries."
)]

use crate::{Error, Result};
use litchi_iwa_common::{WireLimits, varint::encoded_len};
use phf::phf_map;
use prost::Message;

// Keep the generated schema layer in its own crate. The explicit list makes
// this compatibility boundary auditable and prevents decoder-only additions
// from accidentally becoming part of the raw schema crate.
pub use litchi_iwa_protos::{kn, tn, tp, tsa, tsce, tsch, tsd, tsk, tsp, tss, tst, tswp};

const ARCHIVE_CODEC_RECURSION_LIMIT: u32 = 16;
const COMMENT_STORAGE_CODEC_RECURSION_LIMIT: u32 = 64;
const TEXT_STORAGE_CODEC_RECURSION_LIMIT: u32 = 64;
const STORAGE_TEXT_FIELD: u32 = 3;
const TABLE_DATA_LIST_SEGMENT_CODEC_RECURSION_LIMIT: u32 = 64;

/// One allocation-free wire summary for the text projection.
///
/// The generated Buffa lazy view is deliberately only a projection.  It can
/// borrow text, but it cannot establish an aggregate budget for the source
/// message before it starts growing its repeated view.  Keep the raw scan in
/// this crate's adapter so the public trait object never depends on the
/// generated representation or on unbounded source-derived limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct StorageTextPreflight {
    fields: usize,
    unknown_fields: usize,
    text_fragments: usize,
    text_bytes: usize,
    work_bytes: usize,
    output_bytes: usize,
}

impl StorageTextPreflight {
    fn buffa_options(
        self,
        input_bytes: usize,
    ) -> Result<litchi_iwa_protos::text_storage_codec::DecodeOptions> {
        let element_memory = self
            .text_fragments
            .checked_mul(std::mem::size_of::<&str>())
            .ok_or_else(|| storage_wire_error("StorageArchive text element work overflow"))?;
        Ok(litchi_iwa_protos::text_storage_codec::DecodeOptions::new(
            input_bytes.max(1),
            self.unknown_fields,
            element_memory,
            TEXT_STORAGE_CODEC_RECURSION_LIMIT,
        ))
    }
}

fn storage_text_work_bytes(input_bytes: usize, summary: &StorageTextPreflight) -> Result<usize> {
    let field_work = summary
        .fields
        .checked_mul(16)
        .ok_or_else(|| storage_wire_error("StorageArchive field work overflow"))?;
    let element_work = summary
        .text_fragments
        .checked_mul(std::mem::size_of::<&str>())
        .ok_or_else(|| storage_wire_error("StorageArchive text element work overflow"))?;

    // The raw schema-directed pass and the private Buffa projection each walk
    // the complete source envelope. Charge both passes before allowing the
    // projection to allocate or borrow repeated elements.
    let source_pass_work = input_bytes
        .checked_mul(2)
        .ok_or_else(|| storage_wire_error("StorageArchive source work overflow"))?;
    source_pass_work
        .checked_add(field_work)
        .and_then(|work| work.checked_add(element_work))
        .and_then(|work| work.checked_add(summary.text_bytes))
        .ok_or_else(|| storage_wire_error("StorageArchive work budget overflow"))
}

/// Perform the schema-directed raw wire pass for `TSWP.StorageArchive`.
///
/// Only top-level field 3 is selected by the private Buffa projection.  All
/// other fields, including fields nested inside an unknown group, stay opaque;
/// nevertheless their complete wire framing is checked and charged against
/// the finite field/work budgets.  This keeps truncation, malformed groups,
/// and invalid UTF-8 from reaching a lazy view that would otherwise defer or
/// silently skip them.
fn preflight_storage_text(data: &[u8]) -> Result<StorageTextPreflight> {
    if data.len() > WireLimits::MAX_INPUT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "iWork StorageArchive input exceeds {} bytes",
            WireLimits::MAX_INPUT_BYTES
        )));
    }

    let mut summary = StorageTextPreflight {
        fields: 0,
        unknown_fields: 0,
        text_fragments: 0,
        text_bytes: 0,
        work_bytes: 0,
        output_bytes: 0,
    };
    let mut cursor = 0usize;
    scan_storage_text_fields(
        data,
        &mut cursor,
        None,
        TEXT_STORAGE_CODEC_RECURSION_LIMIT,
        true,
        &mut summary,
    )?;
    if cursor != data.len() {
        return Err(storage_wire_error("trailing bytes after a closed group"));
    }

    // Charge both complete source passes, the repeated-view element pointers,
    // and the owned String text copy. Every arithmetic step is checked before
    // it becomes a Buffa or Vec reservation.
    summary.work_bytes = storage_text_work_bytes(data.len(), &summary)?;
    if summary.work_bytes > WireLimits::MAX_REWRITE_WORK {
        return Err(storage_wire_error(&format!(
            "StorageArchive work exceeds {} units",
            WireLimits::MAX_REWRITE_WORK
        )));
    }

    // The public wrapper owns one String header per selected element and the
    // UTF-8 bytes copied into those strings.  Count both before any reserve so
    // an input made from millions of empty strings is bounded as well.
    summary.output_bytes = summary
        .text_fragments
        .checked_mul(std::mem::size_of::<String>())
        .and_then(|headers| headers.checked_add(summary.text_bytes))
        .ok_or_else(|| storage_wire_error("StorageArchive output size overflow"))?;
    if summary.output_bytes > WireLimits::MAX_OUTPUT_BYTES {
        return Err(storage_wire_error(&format!(
            "StorageArchive output exceeds {} bytes",
            WireLimits::MAX_OUTPUT_BYTES
        )));
    }

    Ok(summary)
}

fn storage_wire_error(message: &str) -> Error {
    Error::InvalidFormat(format!(
        "iWork StorageArchive text wire preflight failed: {message}"
    ))
}

/// Scan one protobuf message or unknown group without retaining field spans.
///
/// `end_group` is the field number whose end-group tag closes this invocation;
/// the root call passes `None`.  `select_text` is false while walking an
/// unknown group, because field 3 inside that opaque value is not a selected
/// `StorageArchive.text` occurrence.
fn scan_storage_text_fields(
    data: &[u8],
    cursor: &mut usize,
    end_group: Option<u32>,
    depth_remaining: u32,
    select_text: bool,
    summary: &mut StorageTextPreflight,
) -> Result<bool> {
    while *cursor < data.len() {
        let tag = read_storage_varint(data, cursor)?;
        let field_number = u32::try_from(tag >> 3)
            .map_err(|_| storage_wire_error("field number does not fit u32"))?;
        if field_number == 0 || field_number > 0x1fff_ffff {
            return Err(storage_wire_error("invalid field number"));
        }
        let wire_type =
            u8::try_from(tag & 7).map_err(|_| storage_wire_error("wire type does not fit u8"))?;
        summary.fields = summary
            .fields
            .checked_add(1)
            .ok_or_else(|| storage_wire_error("field count overflow"))?;
        if summary.fields > WireLimits::MAX_FIELDS {
            return Err(storage_wire_error(&format!(
                "field count exceeds {}",
                WireLimits::MAX_FIELDS
            )));
        }
        if select_text && field_number == STORAGE_TEXT_FIELD && wire_type != 2 {
            return Err(storage_wire_error(
                "selected field 3 must use length-delimited wire type",
            ));
        }

        match wire_type {
            0 => {
                read_storage_varint(data, cursor)?;
                summary.unknown_fields = summary
                    .unknown_fields
                    .checked_add(1)
                    .ok_or_else(|| storage_wire_error("unknown field count overflow"))?;
            },
            1 => {
                take_storage_bytes(data, cursor, 8)?;
                summary.unknown_fields = summary
                    .unknown_fields
                    .checked_add(1)
                    .ok_or_else(|| storage_wire_error("unknown field count overflow"))?;
            },
            2 => {
                let length = read_storage_varint(data, cursor)?;
                let length = usize::try_from(length)
                    .map_err(|_| storage_wire_error("length-delimited field is too large"))?;
                let payload = take_storage_bytes(data, cursor, length)?;
                if select_text && field_number == STORAGE_TEXT_FIELD {
                    let text = std::str::from_utf8(payload)
                        .map_err(|_| storage_wire_error("field 3 is not valid UTF-8"))?;
                    let _ = text;
                    summary.text_fragments = summary
                        .text_fragments
                        .checked_add(1)
                        .ok_or_else(|| storage_wire_error("text fragment count overflow"))?;
                    summary.text_bytes = summary
                        .text_bytes
                        .checked_add(payload.len())
                        .ok_or_else(|| storage_wire_error("text byte count overflow"))?;
                } else {
                    summary.unknown_fields = summary
                        .unknown_fields
                        .checked_add(1)
                        .ok_or_else(|| storage_wire_error("unknown field count overflow"))?;
                }
            },
            3 => {
                if depth_remaining == 0 {
                    return Err(storage_wire_error("group nesting exceeds recursion limit"));
                }
                summary.unknown_fields = summary
                    .unknown_fields
                    .checked_add(1)
                    .ok_or_else(|| storage_wire_error("unknown field count overflow"))?;
                let closed = scan_storage_text_fields(
                    data,
                    cursor,
                    Some(field_number),
                    depth_remaining - 1,
                    false,
                    summary,
                )?;
                if !closed {
                    return Err(storage_wire_error("truncated protobuf group"));
                }
            },
            4 => {
                if end_group == Some(field_number) {
                    return Ok(true);
                }
                return Err(storage_wire_error("unexpected or mismatched end group"));
            },
            5 => {
                take_storage_bytes(data, cursor, 4)?;
                summary.unknown_fields = summary
                    .unknown_fields
                    .checked_add(1)
                    .ok_or_else(|| storage_wire_error("unknown field count overflow"))?;
            },
            _ => return Err(storage_wire_error("invalid protobuf wire type")),
        }
    }

    Ok(end_group.is_none())
}

fn read_storage_varint(data: &[u8], cursor: &mut usize) -> Result<u64> {
    let start = *cursor;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *data
            .get(*cursor)
            .ok_or_else(|| storage_wire_error("truncated protobuf varint"))?;
        *cursor = cursor
            .checked_add(1)
            .ok_or_else(|| storage_wire_error("protobuf varint cursor overflow"))?;
        if index == 9 && byte > 1 {
            return Err(storage_wire_error("protobuf varint is too long"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = cursor
                .checked_sub(start)
                .ok_or_else(|| storage_wire_error("protobuf varint cursor underflow"))?;
            if consumed != encoded_len(value) {
                return Err(storage_wire_error("protobuf varint is noncanonical"));
            }
            return Ok(value);
        }
    }
    Err(storage_wire_error("protobuf varint is too long"))
}

fn take_storage_bytes<'source>(
    data: &'source [u8],
    cursor: &mut usize,
    length: usize,
) -> Result<&'source [u8]> {
    let end = cursor
        .checked_add(length)
        .ok_or_else(|| storage_wire_error("protobuf field range overflow"))?;
    if end > data.len() {
        return Err(storage_wire_error("truncated protobuf field"));
    }
    let payload = &data[*cursor..end];
    *cursor = end;
    Ok(payload)
}

fn archive_codec_decode_options(data: &[u8]) -> litchi_iwa_protos::archive_codec::DecodeOptions {
    litchi_iwa_protos::archive_codec::DecodeOptions::new(
        data.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        data.len().clamp(1, WireLimits::MAX_FIELDS),
        data.len()
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        ARCHIVE_CODEC_RECURSION_LIMIT,
    )
}

/// Static decoder function for ArchiveInfo messages
fn decode_archive_info(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let message = litchi_iwa_protos::archive_codec::decode_archive_info(
        data,
        archive_codec_decode_options(data),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork ArchiveInfo payload failed strict validation: {error}"
        ))
    })?;
    Ok(Box::new(ArchiveInfoWrapper(message)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for MessageInfo messages
fn decode_message_info(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let message = litchi_iwa_protos::archive_codec::decode_message_info(
        data,
        archive_codec_decode_options(data),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork MessageInfo payload failed strict validation: {error}"
        ))
    })?;
    Ok(Box::new(MessageInfoWrapper(message)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for StorageArchive messages
fn decode_storage_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let preflight = preflight_storage_text(data)?;
    let view = litchi_iwa_protos::text_storage_codec::decode_storage_text(
        data,
        preflight.buffa_options(data.len())?,
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork StorageArchive text payload failed strict validation: {error}"
        ))
    })?;

    if view.len() != preflight.text_fragments {
        return Err(storage_wire_error(
            "private Buffa projection disagrees with raw field-3 count",
        ));
    }

    // The neutral trait object outlives the source slice passed to this
    // registry. Copy only the public text projection, with fallible
    // reservations bounded by the raw preflight's output ceiling.
    let mut text = Vec::new();
    text.try_reserve_exact(preflight.text_fragments)
        .map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "iWork StorageArchive text fragments",
                amount: preflight.text_fragments,
            })
        })?;
    let mut observed_bytes = 0usize;
    for fragment in view.fragments() {
        let mut owned = String::new();
        owned.try_reserve_exact(fragment.len()).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "iWork StorageArchive text fragment",
                amount: fragment.len(),
            })
        })?;
        owned.push_str(fragment);
        text.push(owned);
        observed_bytes = observed_bytes
            .checked_add(fragment.len())
            .ok_or_else(|| storage_wire_error("owned text byte count overflow"))?;
    }
    if observed_bytes != preflight.text_bytes {
        return Err(storage_wire_error(
            "private Buffa projection disagrees with raw text byte count",
        ));
    }
    Ok(Box::new(StorageArchiveWrapper { text }) as Box<dyn DecodedMessage>)
}

/// Static decoder function for TableModelArchive messages
fn table_names_codec_decode_options(
    data: &[u8],
) -> litchi_iwa_protos::numbers_names_codec::DecodeOptions {
    let source_bytes = data.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    let source_fields = data.len().clamp(1, WireLimits::MAX_FIELDS);
    let source_work = data
        .len()
        .saturating_mul(4)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let recursion = u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX);
    litchi_iwa_protos::numbers_names_codec::DecodeOptions::new(
        source_bytes,
        source_fields,
        source_work,
        recursion,
    )
}

fn table_names_codec_error(
    error: litchi_iwa_protos::numbers_names_codec::DecodeError,
) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::Fields,
            observed,
            limit: maximum,
        });
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::RewriteWork,
            observed,
            limit: maximum,
        });
    }
    match error.wire_resource_limit() {
        Some(litchi_iwa_protos::numbers_names_codec::WireResourceLimit::Bytes {
            observed,
            maximum,
        }) => Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::InputBytes,
            observed,
            limit: maximum,
        }),
        Some(litchi_iwa_protos::numbers_names_codec::WireResourceLimit::Nesting {
            observed,
            maximum,
        }) => Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::Nesting,
            observed: usize::try_from(observed).unwrap_or(usize::MAX),
            limit: usize::try_from(maximum).unwrap_or(usize::MAX),
        }),
        None | Some(_) => Error::InvalidFormat(format!(
            "iWork TableModelArchive name payload failed strict validation: {error}"
        )),
    }
}

fn own_table_name(table_name: &str) -> Result<String> {
    if table_name.len() > WireLimits::MAX_OUTPUT_BYTES {
        return Err(Error::IwaCommon(
            litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed: table_name.len(),
                limit: WireLimits::MAX_OUTPUT_BYTES,
            },
        ));
    }
    let mut owned = String::new();
    owned.try_reserve_exact(table_name.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "iWork TableModelArchive table name",
            amount: table_name.len(),
        })
    })?;
    owned.push_str(table_name);
    Ok(owned)
}

fn decode_table_model(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    // Keep the archive payload as the source of truth; this neutral route
    // publishes only an owned text projection and never reconstructs a model.
    let names = litchi_iwa_protos::numbers_names_codec::decode_table_names(
        data,
        table_names_codec_decode_options(data),
    )
    .map_err(table_names_codec_error)?;
    let table_name = own_table_name(names.table_name())?;
    Ok(Box::new(TableModelWrapper { table_name }) as Box<dyn DecodedMessage>)
}

/// Static decoder function for TableDataList messages
fn decode_table_data_list(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tst::TableDataList::decode(data)?;
    Ok(Box::new(TableDataListWrapper(msg)) as Box<dyn DecodedMessage>)
}

fn table_data_list_segment_codec_decode_options(
    data: &[u8],
) -> Result<litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeOptions> {
    let source_bytes = data.len().clamp(1, WireLimits::MAX_INPUT_BYTES);
    let source_fields = data.len().clamp(1, WireLimits::MAX_FIELDS);
    let source_work = data
        .len()
        .checked_mul(32)
        .ok_or_else(|| {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed: usize::MAX,
                limit: WireLimits::MAX_REWRITE_WORK,
            })
        })?
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let source_references = data.len().clamp(1, WireLimits::MAX_FIELDS);
    let source_text = data.len().clamp(1, WireLimits::MAX_OUTPUT_BYTES);
    Ok(
        litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeOptions::new(
            source_bytes,
            source_fields,
            source_work,
            TABLE_DATA_LIST_SEGMENT_CODEC_RECURSION_LIMIT,
            source_references,
            source_text,
        ),
    )
}

fn table_data_list_segment_codec_error(
    error: litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeError,
) -> Error {
    use litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeLimit;

    let Some(limit) = error.resource_limit() else {
        return Error::InvalidFormat(
            "iWork TableDataListSegment payload failed strict validation".to_owned(),
        );
    };
    match limit {
        DecodeLimit::Bytes { observed, maximum } => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::InputBytes,
                observed,
                limit: maximum,
            })
        }
        DecodeLimit::Fields { observed, maximum } => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed,
                limit: maximum,
            })
        }
        DecodeLimit::Work { observed, maximum } => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed,
                limit: maximum,
            })
        }
        DecodeLimit::Nesting { observed, maximum } => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Nesting,
                observed: usize::try_from(observed).unwrap_or(usize::MAX),
                limit: usize::try_from(maximum).unwrap_or(usize::MAX),
            })
        }
        DecodeLimit::References { observed, maximum } => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed,
                limit: maximum,
            })
        }
        DecodeLimit::Text { observed, maximum } => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed,
                limit: maximum,
            })
        }
        DecodeLimit::Allocation { requested } => {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "iWork TableDataListSegment text projection",
                amount: requested,
            })
        }
        _ => Error::InvalidFormat(
            "iWork TableDataListSegment payload failed strict validation".to_owned(),
        ),
    }
}

#[derive(Debug, Default)]
struct TableDataListSegmentTextVisitor {
    text: Vec<String>,
    output_bytes: usize,
    staging_work_bytes: usize,
    semantic_error: Option<Error>,
}

impl TableDataListSegmentTextVisitor {
    fn record_semantic_error(&mut self, error: Error) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(error);
        }
    }

    fn take_parts(self) -> (Vec<String>, usize, usize, Option<Error>) {
        (
            self.text,
            self.output_bytes,
            self.staging_work_bytes,
            self.semantic_error,
        )
    }
}

impl litchi_iwa_protos::numbers_table_cell_storage_codec::StorageVisitor
    for TableDataListSegmentTextVisitor
{
    fn visit_list_entry(
        &mut self,
        entry: litchi_iwa_protos::numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> std::result::Result<(), litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeError>
    {
        let Some(value) = entry.string_value() else {
            return Ok(());
        };
        if value.is_empty() || self.semantic_error.is_some() {
            return Ok(());
        }

        let string_work = std::mem::size_of::<String>()
            .checked_add(value.len())
            .unwrap_or(usize::MAX);
        let staging_work_bytes = self
            .staging_work_bytes
            .checked_add(string_work)
            .unwrap_or(usize::MAX);
        if staging_work_bytes > WireLimits::MAX_REWRITE_WORK {
            self.record_semantic_error(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed: staging_work_bytes,
                limit: WireLimits::MAX_REWRITE_WORK,
            }));
            return Ok(());
        }

        let output_bytes = self
            .output_bytes
            .checked_add(string_work)
            .unwrap_or(usize::MAX);
        if output_bytes > WireLimits::MAX_OUTPUT_BYTES {
            self.record_semantic_error(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed: output_bytes,
                limit: WireLimits::MAX_OUTPUT_BYTES,
            }));
            return Ok(());
        }

        let requested = self.text.len().checked_add(1).unwrap_or(usize::MAX);
        if self.text.try_reserve_exact(1).is_err() {
            self.record_semantic_error(Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "iWork TableDataListSegment text fragments",
                amount: requested,
            }));
            return Ok(());
        }
        let mut owned = String::new();
        if owned.try_reserve_exact(value.len()).is_err() {
            self.record_semantic_error(Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "iWork TableDataListSegment text",
                amount: value.len(),
            }));
            return Ok(());
        }
        owned.push_str(value);
        self.text.push(owned);
        self.output_bytes = output_bytes;
        self.staging_work_bytes = staging_work_bytes;
        Ok(())
    }
}

/// Static decoder function for segmented TableDataList payloads.
fn decode_table_data_list_segment(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let mut visitor = TableDataListSegmentTextVisitor::default();
    let options = table_data_list_segment_codec_decode_options(data)?;
    let (_snapshot, report) =
        litchi_iwa_protos::numbers_table_cell_storage_codec::decode_table_data_list_segment_with_visitor(
            data,
            options,
            &mut visitor,
        )
        .map_err(table_data_list_segment_codec_error)?;
    let (text, output_bytes, staging_work_bytes, semantic_error) = visitor.take_parts();
    let total_work = report
        .work_bytes()
        .checked_add(staging_work_bytes)
        .ok_or_else(|| {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed: usize::MAX,
                limit: WireLimits::MAX_REWRITE_WORK,
            })
        })?;
    if total_work > WireLimits::MAX_REWRITE_WORK {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::RewriteWork,
            observed: total_work,
            limit: WireLimits::MAX_REWRITE_WORK,
        }));
    }
    if let Some(error) = semantic_error {
        return Err(error);
    }
    debug_assert!(output_bytes <= WireLimits::MAX_OUTPUT_BYTES);
    Ok(Box::new(TableDataListSegmentWrapper { text }) as Box<dyn DecodedMessage>)
}

/// Static decoder function for ShapeArchive messages
fn decode_shape_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let msg = tsd::ShapeArchive::decode(data)?;
    Ok(Box::new(ShapeArchiveWrapper(msg)) as Box<dyn DecodedMessage>)
}

/// Static decoder function for DrawableArchive messages
fn decode_drawable_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    tsd::DrawableArchive::decode(data)?;
    Ok(Box::new(DrawableArchiveWrapper) as Box<dyn DecodedMessage>)
}

fn comment_storage_codec_decode_options(
    data: &[u8],
) -> litchi_iwa_protos::comment_storage_codec::DecodeOptions {
    litchi_iwa_protos::comment_storage_codec::DecodeOptions::new(
        data.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        data.len().clamp(1, WireLimits::MAX_FIELDS),
        data.len()
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        COMMENT_STORAGE_CODEC_RECURSION_LIMIT,
        data.len().clamp(1, WireLimits::MAX_FIELDS),
        data.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
    )
}

fn decode_comment_storage_archive(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let comment = litchi_iwa_protos::comment_storage_codec::decode_comment_storage_archive(
        data,
        comment_storage_codec_decode_options(data),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork CommentStorageArchive payload failed strict validation: {error}"
        ))
    })?;
    // The strict codec borrows the caller-owned payload. Stage the complete
    // scalar projection and own its optional text before publishing a trait
    // object whose lifetime is independent of the source bytes.
    let text = comment.text().map(str::to_owned);
    Ok(Box::new(CommentStorageArchiveWrapper { text }) as Box<dyn DecodedMessage>)
}

fn decode_legacy_chart(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let message = tsch::pre_uff::ChartInfoArchive::decode(data)?;
    Ok(Box::new(LegacyChartArchiveWrapper(message)) as Box<dyn DecodedMessage>)
}

fn decode_chart_mediator(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    tsch::ChartMediatorArchive::decode(data)?;
    Ok(Box::new(ChartMediatorArchiveWrapper) as Box<dyn DecodedMessage>)
}

fn decode_chart_drawable(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let message = crate::charts::IWorkChartArchive::decode(data)?;
    Ok(Box::new(ChartDrawableArchiveWrapper(message)) as Box<dyn DecodedMessage>)
}

type DecoderMap = phf::Map<u32, fn(&[u8]) -> Result<Box<dyn DecodedMessage>>>;

/// Perfect hash map of globally shared, non-colliding message type IDs.
///
/// This provides O(1) lookup performance at compile time. It intentionally
/// excludes IDs that are owned by an application namespace. Application
/// editors decode their own schemas at the typed editor boundary.
///
/// Based on analysis of iWork documents and official message type registry:
/// - 200-299: TSK (Document Core)
/// - 400-499: TSS (Stylesheets)
/// - 600-699: TSA (Application Core)
/// - 2000-2999: TSWP (Word Processing / Text)
/// - 3000-3999: TSD (Drawing / Shapes)
/// - 4000-4999: TSCE (Calculation Engine)
/// - 5000-5999: TSCH (Charts)
/// - 6000-6999: TST (Tables)
/// - 10000-10999: TP (Pages-specific)
/// - 12000-12999: TN (Numbers-specific)
/// - 1-25, 100-199: KN (Keynote-specific)
///
/// Note: Message types are application-specific and may overlap between apps.
static SHARED_DECODERS: DecoderMap = phf_map! {
    // TST (Table) types - Numbers spreadsheet tables and cells
    // Message type 6001 is TST.TableModelArchive
    6000u32 => decode_table_model,
    6001u32 => decode_table_model,
    6005u32 => decode_table_data_list,
    6011u32 => decode_table_data_list_segment,
    6201u32 => decode_table_data_list,

    // TSD (Drawing) types - Shapes, images, and drawables
    3002u32 => decode_drawable_archive,
    3003u32 => decode_drawable_archive,  // ContainerArchive
    3004u32 => decode_shape_archive,
    3005u32 => decode_shape_archive,     // ImageArchive (shape variant)
    3006u32 => decode_shape_archive,     // MaskArchive
    3007u32 => decode_shape_archive,     // MovieArchive
    3008u32 => decode_shape_archive,     // GroupArchive
    3009u32 => decode_shape_archive,     // ConnectionLineArchive
    3056u32 => decode_comment_storage_archive,

    // TSCH (Charts) types
    5000u32 => decode_legacy_chart,
    5004u32 => decode_chart_mediator,
    5021u32 => decode_chart_drawable,

    // TSWP (Word Processing) types - Text storage used across all apps
    2001u32 => decode_storage_archive,
    2002u32 => decode_storage_archive,
    2003u32 => decode_storage_archive,
    2004u32 => decode_storage_archive,
    2005u32 => decode_storage_archive,
    2006u32 => decode_storage_archive,
    2007u32 => decode_storage_archive,
    2008u32 => decode_storage_archive,
    2009u32 => decode_storage_archive,
    2010u32 => decode_storage_archive,
    2011u32 => decode_storage_archive,
    2012u32 => decode_storage_archive,
    2013u32 => decode_storage_archive,
    2014u32 => decode_storage_archive,
    2022u32 => decode_storage_archive,

};

/// TSP core messages are shared in meaning but their numeric IDs collide with
/// application-owned messages. They are included only in the neutral archive
/// text projection below.
static COMMON_DECODERS: DecoderMap = phf_map! {
    1u32 => decode_archive_info,
    2u32 => decode_message_info,
};

/// Decode a message for the neutral archive text projection.
///
/// The archive layer never guesses an application namespace. It only accepts
/// the shared schemas and returns an unsupported-type error for everything
/// else, leaving application-specific decoding to the owning editor crate.
pub(crate) fn decode_common(message_type: u32, data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    let Some(decoder) = COMMON_DECODERS
        .get(&message_type)
        .or_else(|| SHARED_DECODERS.get(&message_type))
    else {
        return Err(Error::UnsupportedMessageType(message_type));
    };
    decoder(data)
}

/// Trait for decoded iWork messages retained by immutable bundle snapshots.
///
/// Decoded messages are read-only after construction, so requiring both
/// marker traits makes the containing archive and bundle safe to share across
/// concurrent readers without a runtime lock.
pub trait DecodedMessage: std::fmt::Debug + Send + Sync {
    /// Extract text content from the message if available
    fn extract_text(&self) -> Vec<String> {
        Vec::new()
    }
}

/// Wrapper for ArchiveInfo message
#[derive(Debug)]
struct ArchiveInfoWrapper(tsp::ArchiveInfo);

impl DecodedMessage for ArchiveInfoWrapper {
    fn extract_text(&self) -> Vec<String> {
        Vec::new() // ArchiveInfo doesn't contain text
    }
}

/// Wrapper for MessageInfo message
#[derive(Debug)]
struct MessageInfoWrapper(tsp::MessageInfo);

impl DecodedMessage for MessageInfoWrapper {
    fn extract_text(&self) -> Vec<String> {
        Vec::new() // MessageInfo doesn't contain text
    }
}

/// Wrapper for StorageArchive message (text content)
#[derive(Debug)]
pub struct StorageArchiveWrapper {
    text: Vec<String>,
}

impl DecodedMessage for StorageArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.text.clone()
    }
}

/// Wrapper for Table Model Archive (Numbers tables)
#[derive(Debug)]
pub struct TableModelWrapper {
    // The original TableModelArchive bytes remain authoritative in the archive
    // object; only this best-effort text projection is owned here.
    table_name: String,
}

impl DecodedMessage for TableModelWrapper {
    fn extract_text(&self) -> Vec<String> {
        let mut text = Vec::new();
        // Extract table name if present
        if !self.table_name.is_empty() {
            text.push(self.table_name.clone());
        }
        // Note: Cell contents are stored in data_store which requires complex
        // processing to extract. For now, we only return the table name.
        text
    }
}

/// Wrapper for Table Data List (cell content storage)
#[derive(Debug)]
pub struct TableDataListWrapper(pub tst::TableDataList);

impl DecodedMessage for TableDataListWrapper {
    fn extract_text(&self) -> Vec<String> {
        // TableDataList contains actual cell data as ListEntry items
        // Extract string values from entries
        let mut strings = Vec::new();

        for entry in &self.0.entries {
            if let Some(ref string_val) = entry.string
                && !string_val.is_empty()
            {
                strings.push(string_val.clone());
            }
        }

        strings
    }
}

/// Wrapper for a segmented TableDataList payload.
#[derive(Debug)]
pub struct TableDataListSegmentWrapper {
    text: Vec<String>,
}

impl DecodedMessage for TableDataListSegmentWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.text.clone()
    }
}

/// Wrapper for Shape Archive
#[derive(Debug)]
pub struct ShapeArchiveWrapper(pub tsd::ShapeArchive);

impl DecodedMessage for ShapeArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        // Shapes can contain text, particularly text boxes
        // Text is typically stored in the DrawableArchive's accessibility description
        // or in referenced TSWP.StorageArchive objects (handled by shape text extractor)
        let mut text = Vec::new();

        // super_ is a required field, not Optional
        let drawable = &self.0.super_;

        // Extract accessibility description if present (often used for alt text/labels)
        if let Some(ref desc) = drawable.accessibility_description
            && !desc.is_empty()
        {
            text.push(desc.clone());
        }

        // Hyperlink URLs can also contain meaningful text
        if let Some(ref url) = drawable.hyperlink_url
            && !url.is_empty()
        {
            text.push(url.clone());
        }

        text
    }
}

/// Wrapper for Drawable Archive
#[derive(Debug)]
struct DrawableArchiveWrapper;

impl DecodedMessage for DrawableArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        // Drawables are visual elements without direct text
        Vec::new()
    }
}

/// Wrapper for TSD comment storage used by cell and drawable comments.
#[derive(Debug)]
pub struct CommentStorageArchiveWrapper {
    text: Option<String>,
}

impl DecodedMessage for CommentStorageArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.text
            .iter()
            .filter(|text| !text.is_empty())
            .cloned()
            .collect()
    }
}

/// Wrapper for the legacy, inline-data chart representation.
#[derive(Debug)]
pub struct LegacyChartArchiveWrapper(pub tsch::pre_uff::ChartInfoArchive);

impl DecodedMessage for LegacyChartArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.0
            .chart_model
            .inline_grid
            .iter()
            .flat_map(|grid| grid.row_name.iter().chain(&grid.column_name))
            .filter(|text| !text.is_empty())
            .cloned()
            .collect()
    }
}

/// Wrapper for a chart's data mediator.
#[derive(Debug)]
struct ChartMediatorArchiveWrapper;

impl DecodedMessage for ChartMediatorArchiveWrapper {}

/// Wrapper for an extension-backed modern chart drawable.
#[derive(Debug)]
pub struct ChartDrawableArchiveWrapper(pub crate::charts::IWorkChartArchive);

impl DecodedMessage for ChartDrawableArchiveWrapper {
    fn extract_text(&self) -> Vec<String> {
        self.0
            .chart
            .iter()
            .filter_map(|chart| chart.grid.as_ref())
            .flat_map(|grid| grid.row_name.iter().chain(&grid.column_name))
            .filter(|text| !text.is_empty())
            .cloned()
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table_model_wire(table_name: &[u8]) -> Vec<u8> {
        let name_len = u8::try_from(table_name.len()).expect("test table name fits one byte");
        let mut data = vec![0x0a, 0x02, b'i', b'd', 0x42, name_len];
        data.extend_from_slice(table_name);
        data
    }

    #[test]
    fn neutral_decoder_registry_contains_only_supported_types() {
        assert!(COMMON_DECODERS.contains_key(&1));
        assert!(COMMON_DECODERS.contains_key(&2));
        assert!(SHARED_DECODERS.contains_key(&6001)); // TST.TableModelArchive
        assert!(SHARED_DECODERS.contains_key(&6011)); // TST.TableDataListSegment
        assert!(SHARED_DECODERS.contains_key(&2001)); // TSWP.StorageArchive
        assert!(SHARED_DECODERS.contains_key(&2002)); // StorageArchive variant
        assert!(SHARED_DECODERS.contains_key(&2003)); // StorageArchive variant
        assert!(SHARED_DECODERS.contains_key(&2022)); // Common StorageArchive type
        assert!(SHARED_DECODERS.contains_key(&3056)); // TSD.CommentStorageArchive
    }

    #[test]
    fn archive_info_is_decoded_without_application_guessing() {
        let archive_info = tsp::ArchiveInfo::default().encode_to_vec();
        let decoded = decode_common(1, &archive_info).unwrap();
        assert!(decoded.extract_text().is_empty());
    }

    #[test]
    fn message_info_preserves_empty_optional_fields() {
        // `type` and `length` are the only required fields. The strict
        // projection must keep all optional/repeated fields at their empty
        // or absent values rather than requiring generated defaults.
        let message_info = [0x08, 0x07, 0x18, 0x0b];
        let decoded = decode_common(2, &message_info).unwrap();
        assert!(decoded.extract_text().is_empty());
    }

    #[test]
    fn table_model_names_projection_owns_utf8_text_and_ignores_unknowns() {
        let expected = "表 Café №42";
        let mut data = table_model_wire(expected.as_bytes());
        data.extend_from_slice(&[0x98, 0x06, 0x01]); // unknown field 99 = 1

        let decoded = decode_common(6001, &data).expect("table model name must decode");
        data.fill(0);

        assert_eq!(decoded.extract_text(), [expected]);
    }

    #[test]
    fn table_model_names_projection_rejects_malformed_payload() {
        let malformed = [0x0a, 0x02, b'i']; // truncated table_id
        assert!(decode_common(6001, &malformed).is_err());
        assert!(decode_common(6001, &[]).is_err());
    }

    #[test]
    fn table_model_names_projection_rejects_invalid_utf8() {
        let data = table_model_wire(&[0xff]);
        assert!(decode_common(6001, &data).is_err());
    }

    #[test]
    fn table_model_names_projection_rejects_duplicate_singular_name() {
        let mut data = table_model_wire(b"first");
        data.extend_from_slice(&[0x42, 0x06]);
        data.extend_from_slice(b"second");

        assert!(decode_common(6001, &data).is_err());
    }

    #[test]
    fn table_model_route_has_no_production_generated_decode() {
        let source = include_str!("protobuf.rs");
        let production = source
            .split_once("#[cfg(test)]")
            .map(|(production, _)| production)
            .expect("test module marker is present");
        let body = production
            .split_once("fn table_names_codec_decode_options")
            .and_then(|(_, rest)| rest.split_once("fn decode_table_data_list"))
            .map(|(body, _)| body)
            .expect("table model decoder body is present");
        assert!(body.contains("numbers_names_codec::decode_table_names"));
        assert!(body.contains("WireLimits::MAX_INPUT_BYTES"));
        assert!(body.contains("WireLimits::MAX_FIELDS"));
        assert!(body.contains("WireLimits::MAX_REWRITE_WORK"));
        assert!(body.contains("try_reserve_exact"));
        assert!(!body.contains("TableModelArchive::decode"));
    }

    #[test]
    fn archive_headers_reject_malformed_and_truncated_nested_payloads() {
        // Prost accepts this proto2 child without its required `length`;
        // the registry path must publish nothing when strict projection
        // rejects it.
        let malformed_message_info = [0x12, 0x02, 0x08, 0x01];
        assert!(tsp::ArchiveInfo::decode(malformed_message_info.as_slice()).is_ok());
        assert!(decode_common(1, &malformed_message_info).is_err());

        // The nested MessageInfo body is cut off in the middle of its
        // length-delimited payload.
        let truncated_archive_info = [0x12, 0x04, 0x08, 0x01, 0x18];
        assert!(decode_common(1, &truncated_archive_info).is_err());

        // A malformed FieldInfo path is deferred by the lazy view until the
        // projection is forced; the registry must still fail atomically.
        let malformed_field_info = [
            0x08, 0x01, 0x18, 0x00, // required MessageInfo fields
            0x22, 0x03, 0x0a, 0x01, 0x80, // unterminated FieldPath varint
        ];
        assert!(decode_common(2, &malformed_field_info).is_err());
    }

    #[test]
    fn archive_projection_does_not_publish_partial_message_infos() {
        let valid_message_info = tsp::MessageInfo {
            r#type: 7,
            length: 11,
            ..Default::default()
        }
        .encode_to_vec();
        let malformed_message_info = [0x08, 0x09]; // missing required length

        let mut archive_info = Vec::new();
        archive_info.extend_from_slice(&[0x12, valid_message_info.len() as u8]);
        archive_info.extend_from_slice(&valid_message_info);
        archive_info.extend_from_slice(&[0x12, malformed_message_info.len() as u8]);
        archive_info.extend_from_slice(&malformed_message_info);

        assert!(decode_common(1, &archive_info).is_err());
    }

    #[test]
    fn shared_storage_extracts_text() {
        let storage = tswp::StorageArchive {
            text: vec!["shared".to_owned()],
            ..Default::default()
        };
        let data = storage.encode_to_vec();
        assert_eq!(
            decode_common(2001, &data).unwrap().extract_text(),
            ["shared"]
        );
    }

    #[test]
    fn shared_storage_projection_preserves_fragment_order_and_unknowns() {
        // Field 99 is outside the narrow Buffa projection. It must remain
        // opaque while the selected repeated text stays source ordered.
        let data = [
            0x1a, 0x01, b'a', // text = "a"
            0x98, 0x06, 0x01, // unknown field 99 = 1
            0x1a, 0x01, b'b', // text = "b"
        ];
        assert_eq!(
            decode_common(2001, &data).unwrap().extract_text(),
            ["a", "b"]
        );
    }

    #[test]
    fn shared_storage_projection_rejects_truncated_text_before_publication() {
        // Prost would not expose a partial string either, but this assertion
        // fixes the production route to the bounded lazy projection rather
        // than relying on generated-message error behavior.
        let malformed = [0x1a, 0x02, b'x'];
        assert!(decode_common(2001, &malformed).is_err());
    }

    #[test]
    fn shared_storage_preflight_rejects_noncanonical_varints() {
        // The raw pass owns the framing contract for the private projection:
        // keys, lengths, and opaque wire-type-0 values must all use their
        // shortest varint representation.
        let overlong_key = [0x9a, 0x80, 0x00, 0x01, b'a'];
        let overlong_length = [0x1a, 0x81, 0x00, b'a'];
        let overlong_varint_value = [0x08, 0x80, 0x00];

        assert!(preflight_storage_text(&overlong_key).is_err());
        assert!(preflight_storage_text(&overlong_length).is_err());
        assert!(preflight_storage_text(&overlong_varint_value).is_err());
    }

    #[test]
    fn shared_storage_preflight_rejects_wrong_wire_for_top_level_text_only() {
        for wire_type in [0u8, 1, 3, 4, 5] {
            let wrong_wire = [(STORAGE_TEXT_FIELD as u8) << 3 | wire_type];
            let error = preflight_storage_text(&wrong_wire).expect_err("wrong field-3 wire");
            assert!(
                error
                    .to_string()
                    .contains("selected field 3 must use length-delimited wire type")
            );
        }

        // A field 3 nested in an unknown group is opaque and must not be
        // mistaken for the selected top-level text field.
        let opaque_group = [0x3b, 0x18, 0x01, 0x3c];
        assert!(preflight_storage_text(&opaque_group).is_ok());
    }

    #[test]
    fn shared_storage_work_ceiling_charges_both_source_passes() {
        let summary = StorageTextPreflight {
            fields: 1,
            unknown_fields: 1,
            text_fragments: 0,
            text_bytes: 0,
            work_bytes: 0,
            output_bytes: 0,
        };
        let field_work = 16;
        let exact_input = (WireLimits::MAX_REWRITE_WORK - field_work) / 2;

        assert_eq!(
            storage_text_work_bytes(exact_input, &summary).unwrap(),
            WireLimits::MAX_REWRITE_WORK
        );
        assert!(
            storage_text_work_bytes(exact_input + 1, &summary).unwrap()
                > WireLimits::MAX_REWRITE_WORK
        );
        assert!(storage_text_work_bytes(usize::MAX, &summary).is_err());
    }

    #[test]
    fn shared_storage_preflight_rejects_aggregate_work_over_ceiling() {
        // A single opaque length-delimited field is enough to exceed the
        // aggregate work budget once both complete source passes are charged.
        // Keep the payload valid so this exercises the work ceiling rather
        // than truncation or a malformed length.
        let input_len = WireLimits::MAX_REWRITE_WORK / 2;
        let key = [0x9a, 0x06]; // field 99, length-delimited
        let body_len = input_len - key.len() - 4;
        let encoded_body_len = litchi_iwa_common::varint::encode_varint(body_len as u64);
        assert_eq!(encoded_body_len.len(), 4);

        let mut data = Vec::with_capacity(input_len);
        data.extend_from_slice(&key);
        data.extend_from_slice(&encoded_body_len);
        data.resize(input_len, 0);

        assert_eq!(data.len(), input_len);
        assert!(preflight_storage_text(&data).is_err());
    }

    #[test]
    fn shared_storage_route_has_no_production_generated_decode() {
        let source = include_str!("protobuf.rs");
        let production = source
            .split_once("#[cfg(test)]")
            .map(|(production, _)| production)
            .expect("test module marker is present");
        let body = production
            .split_once("fn decode_storage_archive")
            .and_then(|(_, rest)| rest.split_once("fn decode_table_model"))
            .map(|(body, _)| body)
            .expect("storage decoder body is present");
        assert!(body.contains("text_storage_codec::decode_storage_text"));
        assert!(body.contains("try_reserve_exact"));
        assert!(!body.contains("StorageArchive::decode"));
    }

    #[test]
    fn table_data_list_segments_use_strict_owned_text_projection() {
        let segment = tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::String as i32,
            key_range: tsp::Range {
                location: 7,
                length: 1,
            },
            entries: vec![tst::table_data_list::ListEntry {
                key: 7,
                refcount: 1,
                string: Some("Segmented".to_owned()),
                ..Default::default()
            }],
        };
        let mut data = segment.encode_to_vec();
        let decoded = decode_common(6011, &data).unwrap();
        data.fill(0);
        assert_eq!(decoded.extract_text(), ["Segmented"]);
    }

    #[test]
    fn table_data_list_segment_projection_rejects_later_malformed_entry_atomically() {
        let segment = tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::String as i32,
            key_range: tsp::Range {
                location: 7,
                length: 2,
            },
            entries: vec![tst::table_data_list::ListEntry {
                key: 7,
                refcount: 1,
                string: Some("retained only on success".to_owned()),
                ..Default::default()
            }],
        };
        let mut malformed = segment.encode_to_vec();
        // The second entry has a key but omits required refcount. The first
        // callback may already have staged text, but the neutral decoder must
        // publish no wrapper when strict validation later fails.
        malformed.extend_from_slice(&[0x1a, 0x02, 0x08, 0x08]);
        assert!(decode_common(6011, &malformed).is_err());
    }

    #[test]
    fn comment_storage_uses_its_concrete_decoder() {
        let comment = tsd::CommentStorageArchive {
            text: Some("Review this".to_owned()),
            ..Default::default()
        };
        let decoded = decode_common(3056, &comment.encode_to_vec()).unwrap();
        assert_eq!(decoded.extract_text(), ["Review this"]);
    }

    #[test]
    fn comment_storage_route_rejects_malformed_payload() {
        // The declared text length exceeds the remaining payload. The route
        // must fail before publishing a decoded message.
        let malformed = [0x0a, 0x02, b'x'];
        assert!(decode_common(3056, &malformed).is_err());
    }

    #[test]
    fn comment_storage_route_rejects_invalid_utf8() {
        let invalid_utf8 = [0x0a, 0x01, 0xff];
        assert!(decode_common(3056, &invalid_utf8).is_err());
    }

    #[test]
    fn comment_storage_route_rejects_duplicate_text() {
        let duplicate_text = [0x0a, 0x01, b'a', 0x0a, 0x01, b'b'];
        assert!(decode_common(3056, &duplicate_text).is_err());
    }

    #[test]
    fn modern_chart_drawables_decode_the_extension_payload() {
        let chart = crate::charts::IWorkChartArchive::new(
            tsch::ChartDrawableArchive::default(),
            tsch::ChartArchive {
                chart_type: Some(tsch::ChartType::ColumnChartType2D as i32),
                grid: Some(tsch::ChartGridArchive {
                    row_name: vec!["Revenue".to_owned()],
                    column_name: vec!["2026".to_owned()],
                    ..Default::default()
                }),
                ..Default::default()
            },
        );
        let decoded = decode_common(5_021, &chart.encode().unwrap()).unwrap();
        assert_eq!(decoded.extract_text(), ["Revenue", "2026"]);
    }

    #[test]
    fn unsupported_application_message_is_rejected() {
        let result = decode_common(999, &[]);
        assert!(matches!(result, Err(Error::UnsupportedMessageType(999))));
    }

    #[test]
    fn test_decoder_performance() {
        // Test that decoding is fast with phf::Map
        // This test ensures the static map lookup is working
        let shared_message_types = [6001, 2001, 2002, 2003];

        // Create some dummy data that will fail to decode but test the lookup
        let dummy_data = vec![0u8; 10];

        for &msg_type in &shared_message_types {
            let result = decode_common(msg_type, &dummy_data);
            // We expect this to fail due to invalid protobuf data, but the lookup should be fast
            assert!(result.is_err());
        }
    }
}
