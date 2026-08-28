//! Wire-preserving table dimension header storage.

use super::*;
use litchi_iwa_protos::table_dimension_codec as dimension_codec;
use litchi_numbers::table::dimension::Dimension;
use std::collections::HashSet;

const HEADER_REWRITE_OUTPUT_SLACK: usize = 64;
const HEADER_CODEC_RECURSION_LIMIT: u32 = 64;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const KNOWN_TABLE_ROLE_MESSAGE_TYPES: &[u32] = &[
    401,   // DocumentStylesheet.
    6_000, // TableInfo.
    6_001, // TableModel.
    6_002, // Tile.
    6_003, // TableStyle.
    6_004, // CellStyle.
    6_005, // TableDataList.
    HEADER_BUCKET_MESSAGE_TYPE,
    6_008, // TablePreset.
    6_010, // ConditionalStyleSet.
    6_011, // TableDataListSegment.
    6_201, // Native TableDataList variant.
    6_204, // HiddenStateOwner.
    6_206, // PopUpMenuModel.
    6_218, // RichTextPayload.
    6_247, // TableStyleNetwork.
    6_267, // ColumnRowUidMap.
    6_284, // TableNameSelection.
    6_305, // StrokeSidecar.
    6_306, // StrokeLayer.
    6_366, // HeaderNameManager.
    6_372, // CategoryOwnerReference.
    6_373, // GroupBy.
];

fn header_storage_decode_options(max_bytes: usize) -> dimension_codec::DecodeOptions {
    let bounded_source_len = max_bytes.max(1);
    let message_bytes = bounded_source_len.clamp(1, litchi_iwa_common::WireLimits::MAX_INPUT_BYTES);
    let fields = bounded_source_len.clamp(1, litchi_iwa_common::WireLimits::MAX_FIELDS);
    let work = bounded_source_len
        .saturating_mul(32)
        .clamp(1, litchi_iwa_common::WireLimits::MAX_REWRITE_WORK);
    dimension_codec::DecodeOptions::new(
        message_bytes,
        fields,
        work,
        HEADER_CODEC_RECURSION_LIMIT,
        bounded_source_len.clamp(1, litchi_numbers::MAX_REFERENCES),
        bounded_source_len.clamp(1, litchi_numbers::DEFAULT_MAX_TEXT_BYTES),
    )
}

struct HeaderDimensionVisitor {
    requested_index: u32,
    minimum_index: u32,
    maximum_index: u32,
    size_bits: Option<u32>,
    seen_indices: HashSet<u32>,
    duplicate_index: Option<u32>,
    out_of_range_index: Option<u32>,
    invalid_size_index: Option<u32>,
}

impl HeaderDimensionVisitor {
    fn new(requested_index: u32, minimum_index: u32, maximum_index: u32) -> Self {
        Self {
            requested_index,
            minimum_index,
            maximum_index,
            size_bits: None,
            seen_indices: HashSet::new(),
            duplicate_index: None,
            out_of_range_index: None,
            invalid_size_index: None,
        }
    }
}

impl dimension_codec::StorageVisitor for HeaderDimensionVisitor {
    fn visit_header(
        &mut self,
        header: dimension_codec::HeaderSnapshot,
    ) -> std::result::Result<(), dimension_codec::DecodeError> {
        let index = header.index();
        if index < self.minimum_index || index >= self.maximum_index {
            self.out_of_range_index.get_or_insert(index);
        }
        if self.seen_indices.contains(&index) {
            self.duplicate_index.get_or_insert(index);
        } else {
            self.seen_indices
                .try_reserve(1)
                .map_err(|_| dimension_codec::DecodeError::allocation(1))?;
            self.seen_indices.insert(index);
        }
        let size_bits = header.size_bits();
        let size = f32::from_bits(size_bits);
        if !size.is_finite() || size < 0.0 || (size == 0.0 && size_bits != 0) {
            self.invalid_size_index.get_or_insert(index);
        }
        if index == self.requested_index {
            self.size_bits = Some(size_bits);
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum HeaderStoragePhase {
    Read,
    Rewrite,
}

fn map_header_storage_decode_error(
    identifier: u64,
    context: &str,
    phase: HeaderStoragePhase,
    error: dimension_codec::DecodeError,
) -> Error {
    use dimension_codec::DecodeLimit;

    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: match phase {
                    HeaderStoragePhase::Read => litchi_iwa_common::LimitKind::InputBytes,
                    HeaderStoragePhase::Rewrite => litchi_iwa_common::LimitKind::OutputBytes,
                },
                observed,
                limit: maximum,
            })
        },
        Some(DecodeLimit::Fields { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed,
                limit: maximum,
            })
        },
        Some(DecodeLimit::Work { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed,
                limit: maximum,
            })
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Nesting,
                observed: observed as usize,
                limit: maximum as usize,
            })
        },
        Some(DecodeLimit::Allocation { requested }) => {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table dimension header storage",
                amount: requested,
            })
        },
        Some(DecodeLimit::References { observed, maximum }) => Error::InvalidFormat(format!(
            "Numbers header bucket {identifier} {context} exceeded its reference limit: observed {observed}, limit {maximum}"
        )),
        Some(DecodeLimit::Text { observed, maximum }) => Error::InvalidFormat(format!(
            "Numbers header bucket {identifier} {context} exceeded its text limit: observed {observed}, limit {maximum}"
        )),
        Some(DecodeLimit::Retained { observed, maximum }) => Error::InvalidFormat(format!(
            "Numbers header bucket {identifier} {context} exceeded its retained-byte limit: observed {observed}, limit {maximum}"
        )),
        None => Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} has invalid {context} payload: {error}"
        )),
        _ => Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} exceeded an unsupported {context} resource limit"
        )),
    }
}

fn decode_header_bucket_dimension(
    identifier: u64,
    source: &[u8],
    index: u32,
    minimum_index: u32,
    maximum_index: u32,
    expected_hash_function: Option<u32>,
    output_slack: usize,
) -> Result<Option<f32>> {
    let mut visitor = HeaderDimensionVisitor::new(index, minimum_index, maximum_index);
    let (snapshot, _report) = dimension_codec::decode_header_storage_bucket_with_visitor(
        source,
        header_storage_decode_options(source.len().saturating_add(output_slack)),
        &mut visitor,
    )
    .map_err(|error| {
        map_header_storage_decode_error(
            identifier,
            "header bucket",
            HeaderStoragePhase::Read,
            error,
        )
    })?;
    if expected_hash_function.is_some_and(|expected| snapshot.bucket_hash_function() != expected) {
        return Err(Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} has a mismatched bucket hash function"
        )));
    }
    if let Some(duplicate_index) = visitor.duplicate_index {
        return Err(Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} repeats dimension index {duplicate_index}"
        )));
    }
    if let Some(out_of_range_index) = visitor.out_of_range_index {
        return Err(Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} has out-of-range dimension index {out_of_range_index}"
        )));
    }
    if let Some(invalid_size_index) = visitor.invalid_size_index {
        return Err(Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} has invalid size for dimension index {invalid_size_index}"
        )));
    }
    let Some(size_bits) = visitor.size_bits else {
        return Ok(None);
    };
    let size = f32::from_bits(size_bits);
    Ok((size != DEFAULT_DIMENSION_POINTS).then_some(size))
}

fn header_bucket_message(
    object: &ArchiveObject,
    identifier: u64,
    index: u32,
    minimum_index: u32,
    maximum_index: u32,
    expected_hash_function: Option<u32>,
    output_slack: usize,
) -> Result<(usize, Option<f32>)> {
    let mut selected = None;
    for (message_index, message) in object.messages.iter().enumerate() {
        if message.type_ != HEADER_BUCKET_MESSAGE_TYPE {
            if KNOWN_TABLE_ROLE_MESSAGE_TYPES.contains(&message.type_) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers header bucket object {identifier} has alternate table-role message type {}",
                    message.type_
                )));
            }
            continue;
        }
        if selected.is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers header bucket object {identifier} has multiple type-{HEADER_BUCKET_MESSAGE_TYPE} payloads"
            )));
        }
        selected = Some((
            message_index,
            decode_header_bucket_dimension(
                identifier,
                &message.data,
                index,
                minimum_index,
                maximum_index,
                expected_hash_function,
                output_slack,
            )?,
        ));
    }
    selected.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Object {identifier} has no type-{HEADER_BUCKET_MESSAGE_TYPE} header bucket payload"
        ))
    })
}

pub(super) fn read_dimension_size(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    dimension: Dimension,
) -> Result<Option<f32>> {
    let selection = header_bucket_selection(model, dimension)?;
    let identifier = selection.identifier;
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} is missing"
        ))
    })?;
    let index = u32::try_from(dimension.index())
        .map_err(|_| Error::ParseError("Numbers table dimension exceeds u32".to_owned()))?;
    let (_message_index, size) = header_bucket_message(
        object,
        identifier,
        index,
        selection.minimum_index,
        selection.maximum_index,
        selection.expected_hash_function,
        0,
    )?;
    Ok(size)
}

pub(super) fn write_dimension_size(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    dimension: Dimension,
    points: f32,
) -> Result<()> {
    let selection = header_bucket_selection(model, dimension)?;
    let identifier = selection.identifier;
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} is missing"
        ))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers header bucket object {identifier} is missing"
            ))
        })?;
        let index = u32::try_from(dimension.index())
            .map_err(|_| Error::ParseError("Numbers table dimension exceeds u32".to_owned()))?;
        let (message_index, current_size) = header_bucket_message(
            object,
            identifier,
            index,
            selection.minimum_index,
            selection.maximum_index,
            selection.expected_hash_function,
            HEADER_REWRITE_OUTPUT_SLACK,
        )?;
        if current_size.is_none() && points == DEFAULT_DIMENSION_POINTS {
            return Ok(());
        }
        if current_size.is_some_and(|size| size.to_bits() == points.to_bits()) {
            return Ok(());
        }
        let original = object.messages[message_index].data.clone();
        let edit = if points == DEFAULT_DIMENSION_POINTS {
            dimension_codec::HeaderSizeEdit::remove(index)
        } else {
            dimension_codec::HeaderSizeEdit::set(index, points.to_bits())
        };
        let options = header_storage_decode_options(
            original.len().saturating_add(HEADER_REWRITE_OUTPUT_SLACK),
        );
        let plan = dimension_codec::plan_header_storage_bucket_sizes(
            &original,
            selection.maximum_index,
            &[edit],
            options,
        )
        .map_err(|error| {
            map_header_storage_decode_error(
                identifier,
                "header bucket rewrite",
                HeaderStoragePhase::Rewrite,
                error,
            )
        })?;
        let requirements = plan.requirements();
        let result_upper_bound = requirements.result_upper_bound();
        let maximum_output_bytes = original
            .len()
            .checked_add(HEADER_REWRITE_OUTPUT_SLACK)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers header bucket output size overflowed".to_owned())
            })?;
        require_header_rewrite_limit(
            litchi_iwa_common::LimitKind::OutputBytes,
            requirements.output_bytes(),
            maximum_output_bytes,
        )?;
        require_header_rewrite_limit(
            litchi_iwa_common::LimitKind::OutputBytes,
            result_upper_bound.source_bytes(),
            maximum_output_bytes,
        )?;
        require_header_rewrite_limit(
            litchi_iwa_common::LimitKind::Fields,
            result_upper_bound.fields(),
            litchi_iwa_common::WireLimits::MAX_FIELDS,
        )?;
        require_header_rewrite_limit(
            litchi_iwa_common::LimitKind::RewriteWork,
            result_upper_bound.work_bytes(),
            litchi_iwa_common::WireLimits::MAX_REWRITE_WORK,
        )?;
        require_header_rewrite_limit(
            litchi_iwa_common::LimitKind::Nesting,
            result_upper_bound.max_depth() as usize,
            HEADER_CODEC_RECURSION_LIMIT as usize,
        )?;
        if result_upper_bound.references() > litchi_numbers::MAX_REFERENCES {
            return Err(Error::InvalidFormat(format!(
                "Numbers header bucket rewrite exceeded its reference limit: observed {}, limit {}",
                result_upper_bound.references(),
                litchi_numbers::MAX_REFERENCES
            )));
        }
        if result_upper_bound.text_bytes() > litchi_numbers::DEFAULT_MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "Numbers header bucket rewrite exceeded its text limit: observed {}, limit {}",
                result_upper_bound.text_bytes(),
                litchi_numbers::DEFAULT_MAX_TEXT_BYTES
            )));
        }
        let execute_options = header_storage_decode_options(requirements.output_bytes());
        let (data, _report) =
            dimension_codec::execute_header_storage_bucket_size_plan(plan, execute_options)
                .map_err(|error| {
                    map_header_storage_decode_error(
                        identifier,
                        "header bucket rewrite",
                        HeaderStoragePhase::Rewrite,
                        error,
                    )
                })?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

#[derive(Clone, Copy)]
struct HeaderBucketSelection {
    identifier: u64,
    minimum_index: u32,
    maximum_index: u32,
    expected_hash_function: Option<u32>,
}

fn local_bucket_reference(
    reference: &crate::protobuf::tsp::Reference,
    context: &str,
) -> Result<u64> {
    if reference.identifier == 0 || reference.deprecated_is_external == Some(true) {
        return Err(Error::InvalidFormat(format!(
            "Numbers table has an invalid {context} reference"
        )));
    }
    Ok(reference.identifier)
}

fn header_bucket_selection(
    model: &TableModelArchive,
    dimension: Dimension,
) -> Result<HeaderBucketSelection> {
    let bucket_rows = u32::try_from(HEADER_BUCKET_ROWS).map_err(|_| {
        Error::InvalidFormat("Numbers row-header bucket span overflowed".to_owned())
    })?;
    let column_identifier = local_bucket_reference(
        &model.base_data_store.column_headers,
        "column-header storage",
    )?;
    let row_headers = &model.base_data_store.row_headers;
    let expected_row_buckets = usize::try_from(model.number_of_rows.div_ceil(bucket_rows))
        .map_err(|_| {
            Error::InvalidFormat("Numbers row-header bucket count overflowed".to_owned())
        })?;
    if row_headers.buckets.len() != expected_row_buckets {
        return Err(Error::InvalidFormat(format!(
            "Numbers table has {} row-header buckets, expected {expected_row_buckets}",
            row_headers.buckets.len()
        )));
    }
    let mut identifiers = HashSet::new();
    identifiers
        .try_reserve(row_headers.buckets.len())
        .map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers table row-header bucket identities",
                amount: row_headers.buckets.len(),
            })
        })?;
    for reference in &row_headers.buckets {
        let identifier = local_bucket_reference(reference, "row-header bucket")?;
        if identifier == column_identifier {
            return Err(Error::InvalidFormat(format!(
                "Numbers table aliases row- and column-header bucket {identifier}"
            )));
        }
        if !identifiers.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers table repeats row-header bucket reference {identifier}"
            )));
        }
    }

    match dimension {
        Dimension::Column(column) => {
            let column = u32::try_from(column).map_err(|_| {
                Error::InvalidFormat("Numbers column dimension exceeds u32".to_owned())
            })?;
            if column >= model.number_of_columns {
                return Err(Error::InvalidFormat(
                    "Numbers column dimension is outside the table".to_owned(),
                ));
            }
            Ok(HeaderBucketSelection {
                identifier: column_identifier,
                minimum_index: 0,
                maximum_index: model.number_of_columns,
                expected_hash_function: None,
            })
        },
        Dimension::Row(row) => {
            let row = u32::try_from(row).map_err(|_| {
                Error::InvalidFormat("Numbers row dimension exceeds u32".to_owned())
            })?;
            if row >= model.number_of_rows {
                return Err(Error::InvalidFormat(
                    "Numbers row dimension is outside the table".to_owned(),
                ));
            }
            let slot = row / bucket_rows;
            let minimum_index = slot.saturating_mul(bucket_rows);
            let maximum_index = minimum_index
                .saturating_add(bucket_rows)
                .min(model.number_of_rows);
            let identifier = row_headers
                .buckets
                .get(usize::try_from(slot).map_err(|_| {
                    Error::InvalidFormat("Numbers row-header bucket slot overflowed".to_owned())
                })?)
                .map(|reference| reference.identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers table {:?} has no row-header bucket for row {row}",
                        model.table_name
                    ))
                })?;
            Ok(HeaderBucketSelection {
                identifier,
                minimum_index,
                maximum_index,
                expected_hash_function: Some(row_headers.bucket_hash_function),
            })
        },
    }
}

fn require_header_rewrite_limit(
    kind: litchi_iwa_common::LimitKind,
    observed: usize,
    limit: usize,
) -> Result<()> {
    if observed > limit {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        }));
    }
    Ok(())
}
