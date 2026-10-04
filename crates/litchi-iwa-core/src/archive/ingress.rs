//! Allocation-free admission facts for a decoded IWA candidate.
//!
//! The ordinary archive parser still owns construction of generated Buffa
//! values.  This module only walks the object framing and the schema-selected
//! header closure with borrowed wire views.  A format transaction can therefore
//! admit the target's header, metadata vectors, and message payload storage
//! before it calls [`Archive::parse_with_limits`].

use std::mem::size_of;

use litchi_iwa_common::decode_varint_from_bytes;
use litchi_iwa_common::wire::{RawWireFields, RawWireLimits};

use super::{
    Archive, ArchiveObject, FieldInfo, FieldPath, HeaderKind, Limits, MessageInfo, RawMessage,
    Result, check_header_length, decode_varint, limit,
};
use crate::LimitKind;

/// Allocation and traversal facts for one decoded IWA candidate.
///
/// The report is produced without constructing an archive or allocating a
/// protobuf field vector.  `retained_bytes` covers the object/header metadata,
/// decoded message payload copies, and one source/canonical header pair for
/// each object. `traversal_bytes` counts every header and payload byte visited
/// by the parser. `allocations` is the number of non-empty logical collections
/// and owned payload/header buffers represented by the report.
///
/// The report is deliberately physical. It does not inspect format-specific
/// message payloads, resolve object identifiers, or replace semantic
/// validation. A caller must charge its own semantic projection after this
/// phase has been admitted.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ArchiveIngressRequirements {
    objects: usize,
    messages: usize,
    field_infos: usize,
    references: usize,
    metadata_items: usize,
    header_fields: usize,
    header_bytes: usize,
    payload_bytes: usize,
    retained_bytes: usize,
    traversal_bytes: usize,
    allocations: usize,
}

impl ArchiveIngressRequirements {
    /// Number of framed archive objects.
    #[must_use]
    pub const fn objects(self) -> usize {
        self.objects
    }

    /// Number of payload metadata records.
    #[must_use]
    pub const fn messages(self) -> usize {
        self.messages
    }

    /// Number of nested `FieldInfo` records.
    #[must_use]
    pub const fn field_infos(self) -> usize {
        self.field_infos
    }

    /// Number of object/data reference occurrences in header metadata.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Number of repeated and optional metadata items checked against the
    /// archive metadata budget.
    #[must_use]
    pub const fn metadata_items(self) -> usize {
        self.metadata_items
    }

    /// Number of structural protobuf fields scanned in the header closure.
    #[must_use]
    pub const fn header_fields(self) -> usize {
        self.header_fields
    }

    /// Header bytes scanned, counting nested messages once at their own
    /// physical slice.
    #[must_use]
    pub const fn header_bytes(self) -> usize {
        self.header_bytes
    }

    /// Message payload bytes retained by the parsed archive.
    #[must_use]
    pub const fn payload_bytes(self) -> usize {
        self.payload_bytes
    }

    /// Heap bytes retained or staged by physical archive ingress.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Bytes inspected by physical archive ingress.
    #[must_use]
    pub const fn traversal_bytes(self) -> usize {
        self.traversal_bytes
    }

    /// Logical allocation events represented by this report.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    fn add_retained(&mut self, bytes: usize) -> Result<()> {
        self.retained_bytes = self
            .retained_bytes
            .checked_add(bytes)
            .ok_or_else(|| super::Error::invalid_archive(0, "candidate retained bytes overflow"))?;
        Ok(())
    }

    fn add_traversal(&mut self, bytes: usize) -> Result<()> {
        self.traversal_bytes = self.traversal_bytes.checked_add(bytes).ok_or_else(|| {
            super::Error::invalid_archive(0, "candidate traversal bytes overflow")
        })?;
        Ok(())
    }

    fn add_allocation(&mut self) -> Result<()> {
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            super::Error::invalid_archive(0, "candidate allocation count overflow")
        })?;
        Ok(())
    }

    fn add_header_field(&mut self) -> Result<()> {
        self.header_fields = self.header_fields.checked_add(1).ok_or_else(|| {
            super::Error::invalid_archive(0, "candidate header field count overflow")
        })?;
        Ok(())
    }

    fn add_message(&mut self) -> Result<()> {
        self.messages = self
            .messages
            .checked_add(1)
            .ok_or_else(|| super::Error::invalid_archive(0, "candidate message count overflow"))?;
        Ok(())
    }

    fn add_field_info(&mut self) -> Result<()> {
        self.field_infos = self.field_infos.checked_add(1).ok_or_else(|| {
            super::Error::invalid_archive(0, "candidate FieldInfo count overflow")
        })?;
        Ok(())
    }

    fn add_references(&mut self, count: usize) -> Result<()> {
        self.references = self.references.checked_add(count).ok_or_else(|| {
            super::Error::invalid_archive(0, "candidate reference count overflow")
        })?;
        Ok(())
    }

    fn add_metadata_items(&mut self, count: usize) -> Result<()> {
        self.metadata_items = self.metadata_items.checked_add(count).ok_or_else(|| {
            super::Error::invalid_archive(0, "candidate metadata item count overflow")
        })?;
        Ok(())
    }
}

impl Archive {
    /// Preflight a decoded candidate without allocating generated archive
    /// values. The returned facts are intended for a caller-owned aggregate
    /// transaction ledger immediately before archive parsing.
    pub fn preflight_with_limits(
        data: &[u8],
        limits: Limits,
    ) -> Result<ArchiveIngressRequirements> {
        let limits = limits.validate()?;
        if data.len() > limits.max_archive_bytes() {
            return Err(limit(
                LimitKind::ArchiveBytes,
                data.len(),
                limits.max_archive_bytes(),
            ));
        }

        let mut requirements = ArchiveIngressRequirements::default();
        let mut cursor = 0usize;
        while cursor < data.len() {
            let object_start = cursor;
            if requirements.objects >= limits.max_objects() {
                return Err(limit(
                    LimitKind::Objects,
                    requirements.objects + 1,
                    limits.max_objects(),
                ));
            }
            requirements.objects = requirements.objects.checked_add(1).ok_or_else(|| {
                super::Error::invalid_archive(object_start, "object count overflow")
            })?;
            requirements.add_retained(size_of::<ArchiveObject>())?;
            requirements.add_allocation()?;

            let (header_length_u64, prefix_length) = decode_varint(&data[cursor..])?;
            cursor = cursor.checked_add(prefix_length).ok_or_else(|| {
                super::Error::invalid_archive(object_start, "header offset overflow")
            })?;
            let header_length = usize::try_from(header_length_u64).map_err(|_| {
                super::Error::invalid_archive(object_start, "header length exceeds usize")
            })?;
            check_header_length(header_length, limits)?;
            let header = data
                .get(
                    cursor..cursor.checked_add(header_length).ok_or_else(|| {
                        super::Error::invalid_archive(object_start, "header range overflow")
                    })?,
                )
                .ok_or_else(|| {
                    super::Error::invalid_archive(object_start, "truncated ArchiveInfo header")
                })?;
            let header_start = cursor;
            let header_end = cursor + header_length;

            requirements.add_traversal(prefix_length)?;
            let header_retained_before = requirements.retained_bytes;
            requirements.add_header_bytes(header_length)?;
            let message_count = scan_archive_info(header, limits, &mut requirements)?;
            if message_count > limits.max_messages_per_object() {
                return Err(limit(
                    LimitKind::MessagesPerObject,
                    message_count,
                    limits.max_messages_per_object(),
                ));
            }
            if requirements.messages > limits.max_messages() {
                return Err(limit(
                    LimitKind::Messages,
                    requirements.messages,
                    limits.max_messages(),
                ));
            }
            if requirements.metadata_items > limits.max_metadata_items() {
                return Err(limit(
                    LimitKind::MetadataItems,
                    requirements.metadata_items,
                    limits.max_metadata_items(),
                ));
            }
            if requirements.header_fields > limits.max_header_fields() {
                return Err(limit(
                    LimitKind::HeaderFields,
                    requirements.header_fields,
                    limits.max_header_fields(),
                ));
            }
            let header_retained = requirements
                .retained_bytes
                .checked_sub(header_retained_before)
                .ok_or_else(|| {
                    super::Error::invalid_archive(
                        object_start,
                        "header storage accounting underflow",
                    )
                })?;
            if header_retained > limits.max_header_memory_bytes() {
                return Err(limit(
                    LimitKind::HeaderMemoryBytes,
                    header_retained,
                    limits.max_header_memory_bytes(),
                ));
            }
            cursor = header_end;

            let payload_start = cursor;
            let message_lengths = scan_message_lengths(header, limits, &mut requirements)?;
            let payload_length = message_lengths.total;
            let object_prefix = prefix_length.checked_add(header_length).ok_or_else(|| {
                super::Error::invalid_archive(object_start, "object prefix overflow")
            })?;
            let object_size = object_prefix.checked_add(payload_length).ok_or_else(|| {
                super::Error::invalid_archive(object_start, "object size overflow")
            })?;
            if object_size > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_size,
                    limits.max_object_bytes(),
                ));
            }
            let payload_end = cursor.checked_add(payload_length).ok_or_else(|| {
                super::Error::invalid_archive(payload_start, "payload range overflow")
            })?;
            if payload_end > data.len() {
                return Err(super::Error::invalid_archive(
                    payload_start,
                    "truncated message payload",
                ));
            }
            requirements.payload_bytes = requirements
                .payload_bytes
                .checked_add(payload_length)
                .ok_or_else(|| {
                    super::Error::invalid_archive(payload_start, "payload bytes overflow")
                })?;
            requirements.add_retained(payload_length)?;
            requirements.add_traversal(payload_length)?;
            if message_lengths.count != 0 {
                requirements.add_retained(
                    message_lengths
                        .count
                        .checked_mul(size_of::<RawMessage>())
                        .ok_or_else(|| {
                            super::Error::invalid_archive(
                                payload_start,
                                "message collection storage overflow",
                            )
                        })?,
                )?;
                requirements.add_allocation()?;
                for _ in 0..message_lengths.count {
                    // Each message owns a payload Vec. Zero-length messages
                    // are conservatively charged as a logical event so the
                    // report remains a safe pre-admission bound across
                    // allocator implementations.
                    requirements.add_allocation()?;
                }
            }
            cursor = payload_end;

            // Keep this assertion adjacent to the cursor update: the parser
            // and this preflight must partition the candidate identically.
            debug_assert!(cursor >= header_start);
        }
        Ok(requirements)
    }

    /// Preflight a decoded candidate using the default physical limits.
    pub fn preflight(data: &[u8]) -> Result<ArchiveIngressRequirements> {
        Self::preflight_with_limits(data, Limits::default())
    }
}

impl ArchiveIngressRequirements {
    fn add_header_bytes(&mut self, bytes: usize) -> Result<()> {
        self.header_bytes = self
            .header_bytes
            .checked_add(bytes)
            .ok_or_else(|| super::Error::invalid_archive(0, "candidate header bytes overflow"))?;
        // Buffa retains the source-authoritative header only when a later
        // preserve path needs it. Charging both the source and canonical
        // header-sized working buffers keeps this report valid for either
        // provenance without assuming a fixed number of archive passes.
        self.add_retained(bytes)?;
        self.add_retained(bytes)?;
        self.add_allocation()?;
        Ok(())
    }
}

fn raw_limits(limits: Limits, header: HeaderKind) -> Result<RawWireLimits> {
    let max_work = limits
        .max_header_memory_bytes()
        .min(RawWireLimits::MAX_WORK)
        .max(
            limits
                .max_header_bytes()
                .min(RawWireLimits::MAX_INPUT_BYTES),
        );
    RawWireLimits::default()
        .with_input_bytes(
            limits
                .max_header_bytes()
                .min(RawWireLimits::MAX_INPUT_BYTES),
        )
        .and_then(|value| value.with_fields(limits.max_header_fields()))
        .and_then(|value| value.with_depth(limits.max_header_nesting()))
        .and_then(|value| value.with_work(max_work))
        .map_err(|error| super::map_wire_error(error, header))
}

fn scan_archive_info(
    source: &[u8],
    limits: Limits,
    requirements: &mut ArchiveIngressRequirements,
) -> Result<usize> {
    requirements.add_traversal(source.len())?;
    let mut message_count = 0usize;
    let mut identifier_seen = false;
    let mut fields =
        RawWireFields::with_limits(source, raw_limits(limits, HeaderKind::ArchiveInfo)?);
    while let Some(field) = fields
        .next()
        .map_err(|error| super::map_wire_error(error, HeaderKind::ArchiveInfo))?
    {
        requirements.add_header_field()?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "ArchiveInfo identifier has an ambiguous wire type",
                    ));
                }
                identifier_seen = true;
            },
            2 => {
                if field.wire_type() != 2 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "ArchiveInfo contains an ambiguous MessageInfo field",
                    ));
                }
                requirements.add_message()?;
                requirements.add_metadata_items(1)?;
                message_count = message_count.checked_add(1).ok_or_else(|| {
                    super::Error::invalid_archive(field.start(), "object message count overflow")
                })?;
                requirements.add_retained(size_of::<MessageInfo>())?;
                requirements.add_allocation()?;
                scan_message_info(field.payload(), limits, requirements)?;
            },
            3 => {
                if field.wire_type() != 0 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "ArchiveInfo merge flag has an ambiguous wire type",
                    ));
                }
            },
            _ => {},
        }
    }
    if !identifier_seen {
        return Err(super::Error::invalid_archive(
            0,
            "ArchiveInfo identifier is missing",
        ));
    }
    Ok(message_count)
}

#[derive(Debug, Clone, Copy, Default)]
struct MessageLengthRequirements {
    count: usize,
    total: usize,
}

fn scan_message_lengths(
    source: &[u8],
    limits: Limits,
    requirements: &mut ArchiveIngressRequirements,
) -> Result<MessageLengthRequirements> {
    let mut lengths = MessageLengthRequirements::default();
    requirements.add_traversal(source.len())?;
    let mut fields =
        RawWireFields::with_limits(source, raw_limits(limits, HeaderKind::ArchiveInfo)?);
    while let Some(field) = fields
        .next()
        .map_err(|error| super::map_wire_error(error, HeaderKind::ArchiveInfo))?
    {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(super::Error::invalid_archive(
                field.start(),
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        let mut message_fields = RawWireFields::with_limits(
            field.payload(),
            raw_limits(limits, HeaderKind::MessageInfo)?,
        );
        requirements.add_traversal(field.payload().len())?;
        let mut length = None;
        while let Some(message_field) = message_fields
            .next()
            .map_err(|error| super::map_wire_error(error, HeaderKind::MessageInfo))?
        {
            if message_field.number() != 3 {
                continue;
            }
            if message_field.wire_type() != 0 {
                return Err(super::Error::invalid_archive(
                    message_field.start(),
                    "MessageInfo length has an ambiguous wire type",
                ));
            }
            let (value, encoded) =
                decode_varint_from_bytes(message_field.payload()).map_err(|_| {
                    super::Error::invalid_archive(
                        message_field.start(),
                        "malformed MessageInfo length",
                    )
                })?;
            if encoded != message_field.payload().len() {
                return Err(super::Error::invalid_archive(
                    message_field.start(),
                    "MessageInfo length has trailing bytes",
                ));
            }
            let value = usize::try_from(value).map_err(|_| {
                super::Error::invalid_archive(message_field.start(), "message length exceeds usize")
            })?;
            if length.replace(value).is_some() {
                return Err(super::Error::invalid_archive(
                    message_field.start(),
                    "MessageInfo length occurs more than once",
                ));
            }
        }
        let length = length.ok_or_else(|| {
            super::Error::invalid_archive(field.start(), "MessageInfo length is missing")
        })?;
        if length > limits.max_message_bytes() {
            return Err(limit(
                LimitKind::MessageBytes,
                length,
                limits.max_message_bytes(),
            ));
        }
        lengths.count = lengths.count.checked_add(1).ok_or_else(|| {
            super::Error::invalid_archive(field.start(), "message count overflow")
        })?;
        lengths.total = lengths.total.checked_add(length).ok_or_else(|| {
            super::Error::invalid_archive(field.start(), "payload length overflow")
        })?;
    }
    Ok(lengths)
}

fn scan_message_info(
    source: &[u8],
    limits: Limits,
    requirements: &mut ArchiveIngressRequirements,
) -> Result<()> {
    let mut type_seen = false;
    let mut length_seen = false;
    requirements.add_traversal(source.len())?;
    let mut fields =
        RawWireFields::with_limits(source, raw_limits(limits, HeaderKind::MessageInfo)?);
    while let Some(field) = fields
        .next()
        .map_err(|error| super::map_wire_error(error, HeaderKind::MessageInfo))?
    {
        requirements.add_header_field()?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "MessageInfo type has an ambiguous wire type",
                    ));
                }
                type_seen = true;
            },
            4 => {
                if field.wire_type() != 2 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "MessageInfo contains an ambiguous FieldInfo field",
                    ));
                }
                requirements.add_metadata_items(1)?;
                requirements.add_field_info()?;
                requirements.add_retained(size_of::<FieldInfo>())?;
                requirements.add_allocation()?;
                scan_field_info(field.payload(), limits, requirements)?;
            },
            2 | 8 | 11 => {
                let count = add_repeated_storage(field, size_of::<u32>(), requirements)?;
                requirements.add_metadata_items(count)?;
            },
            3 => {
                if field.wire_type() != 0 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "MessageInfo length has an ambiguous wire type",
                    ));
                }
                length_seen = true;
            },
            5 | 6 => {
                let count = add_repeated_storage(field, size_of::<u64>(), requirements)?;
                requirements.add_references(count)?;
                requirements.add_metadata_items(count)?;
            },
            9 => {
                if field.wire_type() != 2 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "MessageInfo diff path has an ambiguous wire type",
                    ));
                }
                requirements.add_metadata_items(1)?;
                scan_field_path(field.payload(), limits, requirements)?;
            },
            10 => {
                if field.wire_type() != 2 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "MessageInfo removed path has an ambiguous wire type",
                    ));
                }
                requirements.add_metadata_items(1)?;
                scan_field_path(field.payload(), limits, requirements)?;
            },
            _ => {},
        }
    }
    if !type_seen {
        return Err(super::Error::invalid_archive(
            0,
            "MessageInfo type is missing",
        ));
    }
    if !length_seen {
        return Err(super::Error::invalid_archive(
            0,
            "MessageInfo length is missing",
        ));
    }
    Ok(())
}

fn scan_field_info(
    source: &[u8],
    limits: Limits,
    requirements: &mut ArchiveIngressRequirements,
) -> Result<()> {
    let mut path_seen = false;
    requirements.add_traversal(source.len())?;
    let mut fields =
        RawWireFields::with_limits(source, raw_limits(limits, HeaderKind::MessageInfo)?);
    while let Some(field) = fields
        .next()
        .map_err(|error| super::map_wire_error(error, HeaderKind::MessageInfo))?
    {
        requirements.add_header_field()?;
        match field.number() {
            1 => {
                if field.wire_type() != 2 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "FieldInfo path has an ambiguous wire type",
                    ));
                }
                path_seen = true;
                scan_field_path(field.payload(), limits, requirements)?;
            },
            2 | 3 | 6 => {
                if field.wire_type() != 0 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "FieldInfo enum has an ambiguous wire type",
                    ));
                }
            },
            4 | 5 => {
                let count = add_repeated_storage(field, size_of::<u64>(), requirements)?;
                requirements.add_references(count)?;
                requirements.add_metadata_items(count)?;
            },
            7 => {
                let count = add_repeated_storage(field, size_of::<u32>(), requirements)?;
                requirements.add_metadata_items(count)?;
            },
            8 => {
                if field.wire_type() != 2 {
                    return Err(super::Error::invalid_archive(
                        field.start(),
                        "FieldInfo feature identifier has an ambiguous wire type",
                    ));
                }
                requirements.add_metadata_items(1)?;
                requirements.add_retained(field.payload().len())?;
                requirements.add_allocation()?;
            },
            _ => {},
        }
    }
    if !path_seen {
        return Err(super::Error::invalid_archive(
            0,
            "FieldInfo path is missing",
        ));
    }
    Ok(())
}

fn scan_field_path(
    source: &[u8],
    limits: Limits,
    requirements: &mut ArchiveIngressRequirements,
) -> Result<()> {
    requirements.add_retained(size_of::<FieldPath>())?;
    requirements.add_allocation()?;
    requirements.add_traversal(source.len())?;
    let mut fields =
        RawWireFields::with_limits(source, raw_limits(limits, HeaderKind::MessageInfo)?);
    while let Some(field) = fields
        .next()
        .map_err(|error| super::map_wire_error(error, HeaderKind::MessageInfo))?
    {
        requirements.add_header_field()?;
        if field.number() == 1 {
            let count = repeated_varint_count(field)?;
            requirements.add_metadata_items(count)?;
            requirements.add_retained(count.checked_mul(size_of::<u32>()).ok_or_else(|| {
                super::Error::invalid_archive(field.start(), "FieldPath storage overflow")
            })?)?;
            if count != 0 {
                requirements.add_allocation()?;
            }
        }
    }
    Ok(())
}

fn add_repeated_storage(
    field: litchi_iwa_common::wire::RawWireField<'_>,
    element_size: usize,
    requirements: &mut ArchiveIngressRequirements,
) -> Result<usize> {
    let count = repeated_varint_count(field)?;
    let bytes = count.checked_mul(element_size).ok_or_else(|| {
        super::Error::invalid_archive(field.start(), "repeated metadata storage overflow")
    })?;
    requirements.add_retained(bytes)?;
    if count != 0 {
        requirements.add_allocation()?;
    }
    Ok(count)
}

fn repeated_varint_count(field: litchi_iwa_common::wire::RawWireField<'_>) -> Result<usize> {
    match field.wire_type() {
        0 => Ok(1),
        2 => {
            let mut remaining = field.payload();
            let mut count = 0usize;
            while !remaining.is_empty() {
                let (_, length) = decode_varint_from_bytes(remaining).map_err(|_| {
                    super::Error::invalid_archive(field.start(), "malformed packed integer")
                })?;
                remaining = remaining.get(length..).ok_or_else(|| {
                    super::Error::invalid_archive(field.start(), "packed integer range overflow")
                })?;
                count = count.checked_add(1).ok_or_else(|| {
                    super::Error::invalid_archive(field.start(), "packed integer count overflow")
                })?;
            }
            Ok(count)
        },
        _ => Err(super::Error::invalid_archive(
            field.start(),
            "repeated integer has an ambiguous wire type",
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::{ArchiveIngressRequirements, scan_field_info, scan_message_info};
    use crate::{Archive, ArchiveObject, Error, Limits, RawMessage};

    fn fixture() -> crate::Result<Vec<u8>> {
        Archive {
            objects: vec![ArchiveObject::new(
                17,
                vec![RawMessage {
                    type_: 9,
                    data: vec![1, 2, 3],
                }],
            )?],
        }
        .to_bytes()
    }

    #[test]
    fn preflight_matches_the_physical_archive_shape() -> crate::Result<()> {
        let bytes = fixture()?;
        let requirements = Archive::preflight(&bytes)?;
        assert_eq!(requirements.objects(), 1);
        assert_eq!(requirements.messages(), 1);
        assert_eq!(requirements.payload_bytes(), 3);
        assert!(requirements.header_bytes() > 0);
        assert!(requirements.header_fields() > 0);
        assert!(requirements.retained_bytes() >= requirements.payload_bytes());
        assert!(requirements.traversal_bytes() >= bytes.len());
        assert!(requirements.allocations() > 0);
        let parsed = Archive::parse(&bytes)?;
        assert_eq!(parsed.objects.len(), requirements.objects());
        assert_eq!(parsed.objects[0].messages.len(), requirements.messages());
        Ok(())
    }

    #[test]
    fn preflight_rejects_truncated_payload_before_parse_allocation() -> crate::Result<()> {
        let mut bytes = fixture()?;
        bytes.pop();
        assert!(matches!(
            Archive::preflight(&bytes),
            Err(Error::InvalidArchive { .. })
        ));
        Ok(())
    }

    #[test]
    fn preflight_applies_message_budget_before_payload_work() -> crate::Result<()> {
        let bytes = fixture()?;
        let limits = Limits::default().with_message_bytes(2)?;
        assert!(matches!(
            Archive::preflight_with_limits(&bytes, limits),
            Err(Error::Limit {
                kind: crate::LimitKind::MessageBytes,
                observed: 3,
                maximum: 2,
            })
        ));
        Ok(())
    }

    #[test]
    fn requirements_are_copyable_for_a_caller_ledger() -> crate::Result<()> {
        let bytes = fixture()?;
        let original = Archive::preflight(&bytes)?;
        let copied: ArchiveIngressRequirements = original;
        assert_eq!(copied, original);
        Ok(())
    }

    #[test]
    fn preflight_rejects_missing_required_nested_header_fields() {
        let mut requirements = ArchiveIngressRequirements::default();
        assert!(matches!(
            scan_message_info(&[], Limits::default(), &mut requirements),
            Err(Error::InvalidArchive { .. })
        ));
        assert!(matches!(
            scan_field_info(&[], Limits::default(), &mut requirements),
            Err(Error::InvalidArchive { .. })
        ));
    }
}
