//! Shared lossless access to a native chart-style payload.
//!
//! Chart-level presentation controls are stored in the generated extension of
//! a chart's `TSCH.ChartStyleArchive`. This module resolves the one private
//! style object referenced by a chart and provides guarded wire-level access
//! for focused chart-style feature modules.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::IWorkChartArchive;
use crate::charts::source::{CHART_MESSAGE_TYPE, CHART_STYLE_MESSAGE_TYPE};
use crate::charts::unique_chart_object_archive_name;
use crate::protobuf::tsch;
use crate::wire::parse_wire_fields;
use crate::{Error, IWorkPackage, Result};

/// Proto2 extension holding the generated chart-style properties.
pub(crate) const GENERATED_CHART_STYLE_EXTENSION_FIELD: u32 = 10_000;

/// The single mutable native chart-style payload for one chart.
#[derive(Debug)]
pub(crate) struct ChartStyleSlot {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

/// Resolve the native chart-style payload referenced by one chart.
pub(crate) fn chart_style_slot(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartStyleSlot> {
    let chart_archive = package.archive(chart_archive_name)?;
    let chart_object = chart_archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let mut chart_messages = chart_object
        .messages
        .iter()
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE);
    let Some(chart_message) = chart_messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    if chart_messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    }
    let chart = IWorkChartArchive::decode(chart_message.data.as_slice())?;
    let style_id = chart
        .chart
        .as_ref()
        .and_then(|payload| payload.chart_style.as_ref())
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has no chart style"
            ))
        })?;
    let archive_name = unique_chart_object_archive_name(package, style_id, "chart style object")?;
    let archive = package.archive(&archive_name)?;
    let style_object = archive.object(style_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart style {style_id} is missing"
        ))
    })?;
    let mut messages = style_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_STYLE_MESSAGE_TYPE);
    let Some((message_index, _)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart style {style_id} must have exactly one chart-style payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart style {style_id} must have exactly one chart-style payload"
        )));
    }
    Ok(ChartStyleSlot {
        archive_name,
        object_id: style_id,
        message_index,
    })
}

impl ChartStyleSlot {
    pub(crate) fn archive_name(&self) -> &str {
        &self.archive_name
    }

    pub(crate) const fn object_id(&self) -> u64 {
        self.object_id
    }

    /// Read the resolved chart-style bytes without allocating a rewritten archive.
    pub(crate) fn read<T>(
        &self,
        package: &IWorkPackage,
        read: impl FnOnce(&[u8]) -> Result<T>,
    ) -> Result<T> {
        let archive = package.archive(&self.archive_name)?;
        let object = archive.object(self.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("chart style {} is missing", self.object_id))
        })?;
        let message = object.messages.get(self.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart style {} message index changed unexpectedly",
                self.object_id
            ))
        })?;
        if message.type_ != CHART_STYLE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "chart style {} message type changed unexpectedly",
                self.object_id
            )));
        }
        read(message.data.as_slice())
    }

    /// Reject a mutation that would silently affect another chart.
    pub(crate) fn ensure_exclusive(
        &self,
        package: &IWorkPackage,
        drawable_object_id: u64,
        drawable_label: &str,
    ) -> Result<()> {
        let mut owner_count = 0usize;
        for archive_name in package.iwa_entry_names() {
            let archive = package.archive(archive_name)?;
            for object in &archive.objects {
                for message in object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
                {
                    let chart = IWorkChartArchive::decode(message.data.as_slice())?;
                    if chart
                        .chart
                        .as_ref()
                        .and_then(|payload| payload.chart_style.as_ref())
                        .is_some_and(|reference| reference.identifier == self.object_id)
                    {
                        owner_count = owner_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("chart style owner count overflow".to_owned())
                        })?;
                    }
                }
            }
        }
        if owner_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} style {} is shared by {owner_count} charts",
                self.object_id
            )));
        }
        Ok(())
    }

    /// Transactionally rewrite the resolved style message.
    pub(crate) fn update(
        &self,
        package: &mut IWorkPackage,
        patch: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<()> {
        package.update_archive(&self.archive_name, |archive| {
            let object = archive.object_mut(self.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!("chart style {} is missing", self.object_id))
            })?;
            let original = object.messages.get(self.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart style {} message index changed unexpectedly",
                    self.object_id
                ))
            })?;
            if original.type_ != CHART_STYLE_MESSAGE_TYPE {
                return Err(Error::InvalidFormat(format!(
                    "chart style {} message type changed unexpectedly",
                    self.object_id
                )));
            }
            let data = patch(original.data.as_slice())?;
            object.replace_message(
                self.message_index,
                RawMessage {
                    type_: CHART_STYLE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }
}

/// Decode the outer chart-style payload and locate its generated extension.
pub(crate) fn generated_chart_style_extension(data: &[u8]) -> Result<Option<&[u8]>> {
    tsch::ChartStyleArchive::decode(data)?;
    let fields = parse_wire_fields(data)?;
    let mut extensions = fields
        .iter()
        .filter(|field| field.number() == GENERATED_CHART_STYLE_EXTENSION_FIELD);
    let Some(extension) = extensions.next() else {
        return Ok(None);
    };
    if extensions.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart style extension {GENERATED_CHART_STYLE_EXTENSION_FIELD} occurs more than once"
        )));
    }
    if extension.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart style extension {GENERATED_CHART_STYLE_EXTENSION_FIELD} is not length-delimited"
        )));
    }
    extension.validate_canonical_framing(data)?;
    Ok(Some(extension.checked_payload(data)?))
}

/// A narrow schema description used by chart projections.
///
/// The generated chart archives are deliberately wider than the individual
/// controls exposed by this crate.  These descriptors validate only fields a
/// focused adapter is about to interpret; unknown fields remain opaque and
/// are copied with their original wire representation.
#[derive(Clone, Copy)]
pub(crate) struct KnownWireField {
    pub(crate) number: u32,
    pub(crate) kind: KnownWireFieldKind,
    pub(crate) repeated: bool,
}

#[derive(Clone, Copy)]
pub(crate) enum KnownWireFieldKind {
    Varint,
    Fixed32,
    Fixed64,
    LengthDelimited,
    Message(&'static [KnownWireField]),
}

/// Validate schema-known fields in one message without decoding unrelated
/// producer extensions.  Singular fields must occur at most once and all
/// recognized fields must use canonical framing and the schema wire type.
pub(crate) fn validate_known_wire_fields(
    data: &[u8],
    schema: &'static [KnownWireField],
    label: &str,
) -> Result<()> {
    let fields = parse_wire_fields(data)?;
    for (index, field) in fields.iter().enumerate() {
        let Some(spec) = schema.iter().find(|spec| spec.number == field.number()) else {
            continue;
        };
        if !spec.repeated
            && fields[..index]
                .iter()
                .any(|candidate| candidate.number() == field.number())
        {
            return Err(Error::InvalidFormat(format!(
                "{label} field {} occurs more than once",
                field.number()
            )));
        }
        let expected_wire_type = match spec.kind {
            KnownWireFieldKind::Varint => 0,
            KnownWireFieldKind::Fixed64 => 1,
            KnownWireFieldKind::LengthDelimited | KnownWireFieldKind::Message(_) => 2,
            KnownWireFieldKind::Fixed32 => 5,
        };
        if field.wire_type() != expected_wire_type {
            return Err(Error::InvalidFormat(format!(
                "{label} field {} has the wrong wire type",
                field.number()
            )));
        }
        field.validate_canonical_framing(data)?;
        if matches!(spec.kind, KnownWireFieldKind::Varint) {
            validate_canonical_varint(data, *field, label)?;
        }
        if let KnownWireFieldKind::Message(nested) = spec.kind {
            validate_known_wire_fields(field.checked_payload(data)?, nested, label)?;
        }
    }
    Ok(())
}

fn validate_canonical_varint(
    data: &[u8],
    field: crate::wire::WireField,
    label: &str,
) -> Result<()> {
    let payload = field.checked_payload(data)?;
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(payload).map_err(|error| {
            Error::InvalidFormat(format!(
                "{label} field {} has an invalid varint: {error}",
                field.number()
            ))
        })?;
    if consumed != payload.len() || consumed != litchi_iwa_common::varint::encoded_len(value) {
        return Err(Error::InvalidFormat(format!(
            "{label} field {} has a noncanonical varint",
            field.number()
        )));
    }
    Ok(())
}

/// Replace schema-known fields recursively while retaining unknown fields at
/// every selected message level.  Recursive merging is limited to singular
/// message fields; repeated messages are replaced as a unit because matching
/// their producer-specific identity is not safe.
pub(crate) fn replace_known_wire_fields_preserving_unknown_with_schema(
    existing: &[u8],
    replacement: &[u8],
    schema: &'static [KnownWireField],
    label: &str,
) -> Result<Vec<u8>> {
    validate_known_wire_fields(existing, schema, label)?;
    validate_known_wire_fields(replacement, schema, label)?;
    if existing == replacement {
        return copy_wire_bytes(existing, "selected chart-style nested wire output");
    }

    let existing_fields = parse_wire_fields(existing)?;
    let replacement_fields = parse_wire_fields(replacement)?;
    let capacity = existing
        .len()
        .checked_add(replacement.len())
        .ok_or_else(|| {
            Error::InvalidFormat("selected chart-style wire output overflow".to_owned())
        })?;
    let mut output = Vec::new();
    output.try_reserve_exact(capacity).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "selected chart-style nested wire output",
            amount: capacity,
        })
    })?;

    let is_known = |number: u32| schema.iter().any(|spec| spec.number == number);
    let first_existing_known = existing_fields
        .iter()
        .position(|field| is_known(field.number()));
    let has_known_overlap = existing_fields.iter().any(|existing_field| {
        is_known(existing_field.number())
            && replacement_fields
                .iter()
                .any(|replacement_field| replacement_field.number() == existing_field.number())
    });
    if !has_known_overlap {
        let mut inserted = false;
        for (index, field) in existing_fields.iter().enumerate() {
            if is_known(field.number()) {
                if first_existing_known == Some(index) && !inserted {
                    output.extend_from_slice(replacement);
                    inserted = true;
                }
                continue;
            }
            output.extend_from_slice(field.raw(existing)?);
        }
        if !inserted {
            output.extend_from_slice(replacement);
        }
        return Ok(output);
    }

    let mut emitted_known = Vec::new();
    emitted_known.try_reserve_exact(schema.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "selected chart-style nested known-field set",
            amount: schema.len(),
        })
    })?;

    for existing_field in &existing_fields {
        let Some(spec) = schema
            .iter()
            .find(|spec| spec.number == existing_field.number())
        else {
            output.extend_from_slice(existing_field.raw(existing)?);
            continue;
        };
        if emitted_known.contains(&existing_field.number()) {
            continue;
        }
        emitted_known.push(existing_field.number());

        let replacement_field = replacement_fields
            .iter()
            .find(|field| field.number() == existing_field.number());
        let Some(replacement_field) = replacement_field else {
            continue;
        };
        if spec.repeated {
            for field in replacement_fields
                .iter()
                .filter(|field| field.number() == existing_field.number())
            {
                output.extend_from_slice(field.raw(replacement)?);
            }
            continue;
        }

        if let KnownWireFieldKind::Message(nested) = spec.kind {
            let merged = replace_known_wire_fields_preserving_unknown_with_schema(
                existing_field.checked_payload(existing)?,
                replacement_field.checked_payload(replacement)?,
                nested,
                label,
            )?;
            append_length_delimited_wire_field(
                &mut output,
                existing,
                *existing_field,
                &merged,
                "selected chart-style nested wire output",
            )?;
        } else {
            output.extend_from_slice(replacement_field.raw(replacement)?);
        }
    }

    for replacement_field in &replacement_fields {
        let existing_match = existing_fields
            .iter()
            .any(|field| field.number() == replacement_field.number());
        if !existing_match
            || !schema
                .iter()
                .any(|spec| spec.number == replacement_field.number())
        {
            output.extend_from_slice(replacement_field.raw(replacement)?);
        }
    }
    Ok(output)
}

fn append_length_delimited_wire_field(
    output: &mut Vec<u8>,
    source: &[u8],
    field: crate::wire::WireField,
    payload: &[u8],
    resource: &'static str,
) -> Result<()> {
    let key = field.key(source)?;
    let mut length = [0_u8; litchi_iwa_common::varint::MAX_BYTES];
    let encoded_length = litchi_iwa_common::varint::encode_varint_to_buffer(
        u64::try_from(payload.len())
            .map_err(|_| Error::InvalidFormat("chart nested payload exceeds u64".to_owned()))?,
        &mut length,
    );
    let amount = key
        .len()
        .checked_add(encoded_length.len())
        .and_then(|amount| amount.checked_add(payload.len()))
        .ok_or_else(|| Error::InvalidFormat("chart nested wire output overflow".to_owned()))?;
    let available = output.capacity().saturating_sub(output.len());
    if available < amount {
        let additional = amount - available;
        output.try_reserve_exact(additional).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource,
                amount: additional,
            })
        })?;
    }
    output.extend_from_slice(key);
    output.extend_from_slice(encoded_length);
    output.extend_from_slice(payload);
    Ok(())
}

pub(crate) fn copy_wire_bytes(data: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output.try_reserve_exact(data.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: data.len(),
        })
    })?;
    output.extend_from_slice(data);
    Ok(output)
}

pub(crate) fn clone_vec_fallible(
    source: &[RawMessage],
    resource: &'static str,
) -> Result<Vec<RawMessage>> {
    let mut output = Vec::new();
    output.try_reserve_exact(source.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: source.len(),
        })
    })?;
    for message in source {
        let mut data = Vec::new();
        data.try_reserve_exact(message.data.len()).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource,
                amount: message.data.len(),
            })
        })?;
        data.extend_from_slice(&message.data);
        output.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    Ok(output)
}

pub(crate) fn encode_wire_message<M: Message>(
    message: &M,
    resource: &'static str,
) -> Result<Vec<u8>> {
    let encoded_len = message.encoded_len();
    let mut output = Vec::new();
    output.try_reserve_exact(encoded_len).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: encoded_len,
        })
    })?;
    message
        .encode(&mut output)
        .map_err(|error| Error::InvalidFormat(format!("protobuf encode failed: {error}")))?;
    Ok(output)
}
