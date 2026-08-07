//! Private Buffa-backed codec for the bounded IWA archive-header seam.

use std::fmt;

use buffa::{DecodeOptions as BuffaDecodeOptions, Enumeration as _, Message as _};

use crate::{buffa_generated::TSP as buffa_tsp, tsp};

/// Finite limits already established by the physical archive owner before a
/// lazy header decode begins.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_unknown_fields: usize,
    max_element_memory: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted header.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_unknown_fields: usize,
        max_element_memory: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_unknown_fields,
            max_element_memory,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_unknown_fields)
            .with_element_memory_limit(self.max_element_memory)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Failure from the private Buffa header decoder.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    /// Buffa rejected malformed or over-budget wire data.
    Wire(buffa::DecodeError),
    /// A required proto2 field was absent.
    MissingRequired(&'static str),
    /// A compatibility projection allocation failed.
    Allocation {
        resource: &'static str,
        requested: usize,
    },
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::Allocation {
                resource,
                requested,
            } => write!(
                formatter,
                "allocation for {resource} ({requested} elements or bytes) failed"
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }
}

impl DecodeError {
    /// Required schema field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required(&self) -> Option<&'static str> {
        match &self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(*field),
            DecodeErrorKind::Wire(_) | DecodeErrorKind::Allocation { .. } => None,
        }
    }

    const fn missing_required_field(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }

    const fn allocation(resource: &'static str, requested: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation {
                resource,
                requested,
            },
        }
    }
}

/// Failure from the private Buffa header encoder without exposing its runtime
/// type across the schema crate boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncodeError(buffa::EncodeError);

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for EncodeError {}

impl From<buffa::EncodeError> for EncodeError {
    fn from(error: buffa::EncodeError) -> Self {
        Self(error)
    }
}

/// Decode one already-preflighted `TSP.ArchiveInfo` through Buffa's lazy
/// view, validating every deferred child before returning an owned
/// compatibility value.
pub fn decode_archive_info(
    source: &[u8],
    options: DecodeOptions,
) -> Result<tsp::ArchiveInfo, DecodeError> {
    let recursion_limit = options.recursion_limit;
    let view: buffa_tsp::ArchiveInfoLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    archive_info_from_lazy(&view, recursion_limit)
}

/// Decode one already-preflighted `TSP.MessageInfo` through Buffa's lazy
/// view and validate every deferred child before publication.
pub fn decode_message_info(
    source: &[u8],
    options: DecodeOptions,
) -> Result<tsp::MessageInfo, DecodeError> {
    let recursion_limit = options.recursion_limit;
    let view: buffa_tsp::MessageInfoLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    message_info_from_lazy(&view, recursion_limit)
}

/// Return the Buffa canonical encoded length for a compatibility value.
pub fn archive_info_encoded_len(info: &tsp::ArchiveInfo) -> Result<u32, EncodeError> {
    archive_info_to_buffa(info)
        .try_encoded_len()
        .map_err(Into::into)
}

/// Encode one compatibility value canonically into a caller-owned,
/// pre-reserved buffer under an exact byte ceiling.
pub fn encode_archive_info(
    info: &tsp::ArchiveInfo,
    maximum: u32,
    output: &mut Vec<u8>,
) -> Result<u32, EncodeError> {
    archive_info_to_buffa(info)
        .try_encode_bounded(maximum, output)
        .map_err(Into::into)
}

fn archive_info_from_lazy(
    view: &buffa_tsp::ArchiveInfoLazyView<'_>,
    recursion_limit: u32,
) -> Result<tsp::ArchiveInfo, DecodeError> {
    let mut message_infos = Vec::new();
    reserve_exact(
        &mut message_infos,
        view.message_infos.len(),
        "ArchiveInfo message metadata",
    )?;
    for item in &view.message_infos {
        message_infos.push(message_info_from_lazy(&item?, recursion_limit)?);
    }
    Ok(tsp::ArchiveInfo {
        identifier: view.identifier,
        message_infos,
        should_merge: view.should_merge,
    })
}

fn message_info_from_lazy(
    view: &buffa_tsp::MessageInfoLazyView<'_>,
    recursion_limit: u32,
) -> Result<tsp::MessageInfo, DecodeError> {
    if !view.has_type() {
        return Err(DecodeError::missing_required_field("TSP.MessageInfo.type"));
    }
    if !view.has_length() {
        return Err(DecodeError::missing_required_field(
            "TSP.MessageInfo.length",
        ));
    }

    let mut field_infos = Vec::new();
    reserve_exact(
        &mut field_infos,
        view.field_infos.len(),
        "MessageInfo field metadata",
    )?;
    for (source, item) in view
        .field_infos
        .raw_elements()
        .iter()
        .copied()
        .zip(&view.field_infos)
    {
        field_infos.push(field_info_from_lazy(&item?, source, recursion_limit)?);
    }

    let mut fields_to_remove = Vec::new();
    reserve_exact(
        &mut fields_to_remove,
        view.fields_to_remove.len(),
        "MessageInfo removed field paths",
    )?;
    for item in &view.fields_to_remove {
        fields_to_remove.push(field_path_from_lazy(&item?)?);
    }

    Ok(tsp::MessageInfo {
        r#type: view.r#type,
        version: copy_slice(&view.version, "MessageInfo versions")?,
        length: view.length,
        field_infos,
        object_references: copy_slice(&view.object_references, "MessageInfo object references")?,
        data_references: copy_slice(&view.data_references, "MessageInfo data references")?,
        base_message_index: view.base_message_index,
        diff_merge_version: copy_slice(
            &view.diff_merge_version,
            "MessageInfo diff merge versions",
        )?,
        diff_field_path: view
            .diff_field_path
            .get()?
            .as_ref()
            .map(field_path_from_lazy)
            .transpose()?,
        fields_to_remove,
        diff_read_version: copy_slice(&view.diff_read_version, "MessageInfo diff read versions")?,
    })
}

fn field_info_from_lazy(
    view: &buffa_tsp::FieldInfoLazyView<'_>,
    source: &[u8],
    recursion_limit: u32,
) -> Result<tsp::FieldInfo, DecodeError> {
    if !view.has_path() {
        return Err(DecodeError::missing_required_field("TSP.FieldInfo.path"));
    }
    let path = view
        .path
        .get()?
        .as_ref()
        .map(field_path_from_lazy)
        .transpose()?
        .ok_or_else(|| DecodeError::missing_required_field("TSP.FieldInfo.path"))?;
    let known_field_feature_identifier = view
        .known_field_feature_identifier
        .map(|value| copy_string(value, "FieldInfo feature identifier"))
        .transpose()?;

    Ok(tsp::FieldInfo {
        path,
        r#type: last_int32_field(source, 2, recursion_limit)?,
        unknown_field_rule: last_int32_field(source, 3, recursion_limit)?,
        object_references: copy_slice(&view.object_references, "FieldInfo object references")?,
        data_references: copy_slice(&view.data_references, "FieldInfo data references")?,
        known_field_rule: last_int32_field(source, 6, recursion_limit)?,
        known_field_version: copy_slice(
            &view.known_field_version,
            "FieldInfo known field versions",
        )?,
        known_field_feature_identifier,
    })
}

fn field_path_from_lazy(
    view: &buffa_tsp::FieldPathLazyView<'_>,
) -> Result<tsp::FieldPath, DecodeError> {
    Ok(tsp::FieldPath {
        path: copy_slice(&view.path, "FieldPath components")?,
    })
}

fn last_int32_field(
    mut source: &[u8],
    field_number: u32,
    recursion_limit: u32,
) -> Result<Option<i32>, DecodeError> {
    let mut value = None;
    while !source.is_empty() {
        let tag = buffa::encoding::Tag::decode(&mut source)?;
        if tag.field_number() == field_number {
            buffa::encoding::check_wire_type(tag, buffa::encoding::WireType::Varint)?;
            value = Some(buffa::types::decode_int32(&mut source)?);
        } else {
            buffa::encoding::skip_field_depth(tag, &mut source, recursion_limit)?;
        }
    }
    Ok(value)
}

fn copy_slice<T: Copy>(source: &[T], resource: &'static str) -> Result<Vec<T>, DecodeError> {
    let mut output = Vec::new();
    reserve_exact(&mut output, source.len(), resource)?;
    output.extend_from_slice(source);
    Ok(output)
}

fn copy_string(source: &str, resource: &'static str) -> Result<String, DecodeError> {
    let mut output = String::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_allocation_error| DecodeError::allocation(resource, source.len()))?;
    output.push_str(source);
    Ok(output)
}

fn reserve_exact<T>(
    output: &mut Vec<T>,
    requested: usize,
    resource: &'static str,
) -> Result<(), DecodeError> {
    output
        .try_reserve_exact(requested)
        .map_err(|_allocation_error| DecodeError::allocation(resource, requested))
}

fn archive_info_to_buffa(value: &tsp::ArchiveInfo) -> buffa_tsp::ArchiveInfo {
    buffa_tsp::ArchiveInfo {
        identifier: value.identifier,
        message_infos: value
            .message_infos
            .iter()
            .map(message_info_to_buffa)
            .collect(),
        should_merge: value.should_merge,
        ..Default::default()
    }
}

fn message_info_to_buffa(value: &tsp::MessageInfo) -> buffa_tsp::MessageInfo {
    buffa_tsp::MessageInfo {
        r#type: value.r#type,
        version: value.version.clone(),
        length: value.length,
        field_infos: value.field_infos.iter().map(field_info_to_buffa).collect(),
        object_references: value.object_references.clone(),
        data_references: value.data_references.clone(),
        base_message_index: value.base_message_index,
        diff_merge_version: value.diff_merge_version.clone(),
        diff_field_path: value
            .diff_field_path
            .as_ref()
            .map(field_path_to_buffa)
            .into(),
        fields_to_remove: value
            .fields_to_remove
            .iter()
            .map(field_path_to_buffa)
            .collect(),
        diff_read_version: value.diff_read_version.clone(),
        ..Default::default()
    }
}

fn field_info_to_buffa(value: &tsp::FieldInfo) -> buffa_tsp::FieldInfo {
    let mut output = buffa_tsp::FieldInfo {
        path: buffa::MessageField::some(field_path_to_buffa(&value.path)),
        r#type: value.r#type.and_then(buffa_tsp::field_info::Type::from_i32),
        unknown_field_rule: value
            .unknown_field_rule
            .and_then(buffa_tsp::field_info::UnknownFieldRule::from_i32),
        object_references: value.object_references.clone(),
        data_references: value.data_references.clone(),
        known_field_rule: value
            .known_field_rule
            .and_then(buffa_tsp::field_info::KnownFieldRule::from_i32),
        known_field_version: value.known_field_version.clone(),
        known_field_feature_identifier: value.known_field_feature_identifier.clone(),
        ..Default::default()
    };
    preserve_unknown_enum(
        &mut output.__buffa_unknown_fields,
        2,
        value.r#type,
        output.r#type,
    );
    preserve_unknown_enum(
        &mut output.__buffa_unknown_fields,
        3,
        value.unknown_field_rule,
        output.unknown_field_rule,
    );
    preserve_unknown_enum(
        &mut output.__buffa_unknown_fields,
        6,
        value.known_field_rule,
        output.known_field_rule,
    );
    output
}

fn preserve_unknown_enum<E>(
    unknown_fields: &mut buffa::UnknownFields,
    number: u32,
    raw_value: Option<i32>,
    decoded: Option<E>,
) {
    if let (None, Some(unknown_value)) = (decoded, raw_value) {
        unknown_fields.push(buffa::UnknownField {
            number,
            data: buffa::UnknownFieldData::Varint(u64::from_ne_bytes(
                i64::from(unknown_value).to_ne_bytes(),
            )),
        });
    }
}

fn field_path_to_buffa(value: &tsp::FieldPath) -> buffa_tsp::FieldPath {
    buffa_tsp::FieldPath {
        path: value.path.clone(),
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use prost::Message as _;

    use super::{DecodeOptions, decode_archive_info, decode_message_info, encode_archive_info};
    use crate::tsp;

    fn options(length: usize) -> DecodeOptions {
        DecodeOptions::new(length, length.max(1), 4 * 1024 * 1024, 16)
    }

    fn append_length_delimited(output: &mut Vec<u8>, tag: u8, payload: &[u8]) {
        output.push(tag);
        output.push(one_byte_length(payload.len()));
        output.extend_from_slice(payload);
    }

    fn one_byte_length(length: usize) -> u8 {
        match u8::try_from(length) {
            Ok(byte_length) => byte_length,
            Err(error) => panic!("test payload must fit one-byte length: {error}"),
        }
    }

    fn assert_message_info_parity(
        source: &[u8],
    ) -> Result<tsp::MessageInfo, Box<dyn std::error::Error>> {
        let expected = tsp::MessageInfo::decode(source)?;
        let actual = decode_message_info(source, options(source.len()))?;
        assert_eq!(actual, expected);
        Ok(actual)
    }

    fn assert_archive_info_parity(
        source: &[u8],
    ) -> Result<tsp::ArchiveInfo, Box<dyn std::error::Error>> {
        let expected = tsp::ArchiveInfo::decode(source)?;
        let actual = decode_archive_info(source, options(source.len()))?;
        assert_eq!(actual, expected);
        Ok(actual)
    }

    #[test]
    fn hostile_noncanonical_headers_match_prost_semantics() -> Result<(), Box<dyn std::error::Error>>
    {
        let message = [
            0x08, 0x87, 0x00, 0x98, 0x06, 0x96, 0x01, 0x18, 0x8b, 0x00, 0xa2, 0x06, 0x81, 0x00,
            0xff,
        ];
        let archive = [
            0x08, 0xaa, 0x00, 0x90, 0x03, 0x09, 0x12, 0x8f, 0x00, 0x08, 0x87, 0x00, 0x98, 0x06,
            0x96, 0x01, 0x18, 0x8b, 0x00, 0xa2, 0x06, 0x81, 0x00, 0xff, 0x9a, 0x03, 0x82, 0x00,
            0xde, 0xad, 0x18, 0x81, 0x00,
        ];

        assert_eq!(
            decode_message_info(&message, options(message.len()))?,
            tsp::MessageInfo::decode(message.as_slice())?
        );
        assert_eq!(
            decode_archive_info(&archive, options(archive.len()))?,
            tsp::ArchiveInfo::decode(archive.as_slice())?
        );
        Ok(())
    }

    #[test]
    fn duplicate_nested_fields_and_all_unknown_wire_types_match_prost_merge_semantics()
    -> Result<(), Box<dyn std::error::Error>> {
        // A singular message field merges all of its occurrences. Exercise
        // that rule for both FieldInfo.path and MessageInfo.diff_field_path,
        // while mixing packed/unpacked repeated scalars and duplicate scalar,
        // enum, and string fields.
        let mut field_info = Vec::new();
        append_length_delimited(&mut field_info, 0x0a, &[0x08, 0x01, 0x0a, 0x02, 0x02, 0x03]);
        field_info.extend_from_slice(&[0x10, 0x03, 0x10, 0x63]);
        field_info.extend_from_slice(&[0x20, 0x07, 0x22, 0x02, 0x08, 0x09]);
        field_info.extend_from_slice(&[0x38, 0x01, 0x3a, 0x02, 0x02, 0x03]);
        append_length_delimited(&mut field_info, 0x42, b"old");
        append_length_delimited(&mut field_info, 0x0a, &[0x08, 0x04]);
        append_length_delimited(&mut field_info, 0x42, b"new");

        let mut message = vec![0x08, 0x01, 0x10, 0x01, 0x12, 0x02, 0x02, 0x03, 0x18, 0x04];
        append_length_delimited(&mut message, 0x22, &field_info);
        append_length_delimited(&mut message, 0x4a, &[0x08, 0x08]);
        append_length_delimited(&mut message, 0x4a, &[0x0a, 0x01, 0x09]);
        message.extend_from_slice(&[0x38, 0x05, 0x38, 0x06]);

        // Unknown fields 100..104 cover varint, fixed64,
        // length-delimited, group, and fixed32 wire types respectively.
        message.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
        message.extend_from_slice(&[0xa9, 0x06, 1, 2, 3, 4, 5, 6, 7, 8]);
        message.extend_from_slice(&[0xb2, 0x06, 0x81, 0x00, 0xff]);
        message.extend_from_slice(&[0xbb, 0x06, 0x08, 0x09, 0xbc, 0x06]);
        message.extend_from_slice(&[0xc5, 0x06, 9, 10, 11, 12]);
        message.extend_from_slice(&[0x08, 0x07, 0x18, 0x0b]);

        let decoded_message = assert_message_info_parity(&message)?;
        assert_eq!(decoded_message.r#type, 7);
        assert_eq!(decoded_message.length, 11);
        assert_eq!(decoded_message.version, [1, 2, 3]);
        assert_eq!(decoded_message.base_message_index, Some(6));
        assert_eq!(decoded_message.field_infos[0].path.path, [1, 2, 3, 4]);
        assert_eq!(decoded_message.field_infos[0].r#type, Some(99));
        assert_eq!(decoded_message.field_infos[0].object_references, [7, 8, 9]);
        assert_eq!(
            decoded_message.field_infos[0].known_field_version,
            [1, 2, 3]
        );
        assert_eq!(
            decoded_message.field_infos[0]
                .known_field_feature_identifier
                .as_deref(),
            Some("new")
        );
        assert_eq!(
            decoded_message
                .diff_field_path
                .as_ref()
                .map(|path| path.path.as_slice()),
            Some([8, 9].as_slice())
        );

        let mut archive = vec![0x08, 0x01];
        append_length_delimited(&mut archive, 0x12, &message);
        archive.extend_from_slice(&[0xa0, 0x06, 0x01, 0x08, 0x2a, 0x18, 0x00, 0x18, 0x01]);
        let decoded_archive = assert_archive_info_parity(&archive)?;
        assert_eq!(decoded_archive.identifier, Some(42));
        assert_eq!(decoded_archive.should_merge, Some(true));
        assert_eq!(decoded_archive.message_infos, [decoded_message]);
        Ok(())
    }

    #[test]
    fn overlong_keys_lengths_and_values_match_prost_semantics()
    -> Result<(), Box<dyn std::error::Error>> {
        // Every known key, the packed length, and the scalar values below use
        // a valid but non-minimal varint representation.
        let message = [
            0x88, 0x00, 0x87, 0x00, // type = 7
            0x92, 0x00, 0x82, 0x00, 0x81, 0x00, // packed version = [1]
            0x90, 0x00, 0x82, 0x00, // unpacked version = 2
            0x98, 0x00, 0x80, 0x00, // length = 0
        ];
        let decoded_message = assert_message_info_parity(&message)?;
        assert_eq!(decoded_message.r#type, 7);
        assert_eq!(decoded_message.version, [1, 2]);
        assert_eq!(decoded_message.length, 0);

        let mut archive = vec![
            0x88,
            0x00,
            0xaa,
            0x00, // identifier = 42
            0x92,
            0x00, // overlong MessageInfo key
            0x80 | u8::try_from(message.len())?,
            0x00, // overlong child length
        ];
        archive.extend_from_slice(&message);
        archive.extend_from_slice(&[0x98, 0x00, 0x82, 0x00]); // true from non-zero bool
        let decoded_archive = assert_archive_info_parity(&archive)?;
        assert_eq!(decoded_archive.identifier, Some(42));
        assert_eq!(decoded_archive.should_merge, Some(true));
        assert_eq!(decoded_archive.message_infos, [decoded_message]);
        Ok(())
    }

    #[test]
    fn buffa_canonical_encoding_matches_prost() -> Result<(), Box<dyn std::error::Error>> {
        let info = tsp::ArchiveInfo {
            identifier: Some(42),
            message_infos: vec![tsp::MessageInfo {
                r#type: 7,
                version: vec![1, 0, 5],
                length: 11,
                field_infos: vec![tsp::FieldInfo {
                    path: tsp::FieldPath { path: vec![1, 4] },
                    r#type: Some(tsp::field_info::Type::Message as i32),
                    object_references: vec![9],
                    known_field_feature_identifier: Some("feature".to_owned()),
                    ..Default::default()
                }],
                object_references: vec![4, 5],
                diff_field_path: Some(tsp::FieldPath { path: vec![3] }),
                fields_to_remove: vec![tsp::FieldPath { path: vec![8, 9] }],
                ..Default::default()
            }],
            should_merge: Some(true),
        };
        let expected = info.encode_to_vec();
        let mut actual = Vec::new();
        encode_archive_info(&info, u32::MAX, &mut actual)?;
        assert_eq!(actual, expected);
        assert_eq!(decode_archive_info(&actual, options(actual.len()))?, info);
        Ok(())
    }

    #[test]
    fn lazy_adapter_forces_deferred_child_validation() {
        let malformed = [0x08, 0x01, 0x12, 0x01, 0x80];
        assert!(decode_archive_info(&malformed, options(malformed.len())).is_err());
    }

    #[test]
    fn every_deferred_child_route_is_validated_before_projection() {
        let malformed_children: &[(&str, &[u8])] = &[
            ("FieldInfo body", &[0x22, 0x01, 0x80]),
            ("required FieldInfo.path", &[0x22, 0x03, 0x0a, 0x01, 0x80]),
            ("diff_field_path", &[0x4a, 0x01, 0x80]),
            ("fields_to_remove", &[0x52, 0x01, 0x80]),
            (
                "later diff_field_path fragment",
                &[0x4a, 0x02, 0x08, 0x01, 0x4a, 0x01, 0x80],
            ),
            (
                "later FieldInfo.path fragment",
                &[0x22, 0x07, 0x0a, 0x02, 0x08, 0x01, 0x0a, 0x01, 0x80],
            ),
        ];

        for (context, child) in malformed_children {
            let mut message = vec![0x08, 0x01, 0x18, 0x00];
            message.extend_from_slice(child);
            assert!(
                tsp::MessageInfo::decode(message.as_slice()).is_err(),
                "Prost unexpectedly accepted malformed {context}"
            );
            assert!(
                decode_message_info(&message, options(message.len())).is_err(),
                "Buffa adapter unexpectedly accepted malformed {context}"
            );

            let mut archive = vec![0x08, 0x01];
            append_length_delimited(&mut archive, 0x12, &message);
            assert!(
                tsp::ArchiveInfo::decode(archive.as_slice()).is_err(),
                "Prost unexpectedly accepted malformed {context} through ArchiveInfo"
            );
            assert!(
                decode_archive_info(&archive, options(archive.len())).is_err(),
                "Buffa adapter unexpectedly accepted malformed {context} through ArchiveInfo"
            );
        }
    }

    #[test]
    fn codec_message_and_unknown_field_limits_are_inclusive() {
        let message = [0x08, 0x01, 0x18, 0x00];
        assert!(decode_message_info(&message, options(message.len())).is_ok());
        assert!(
            decode_message_info(&message, DecodeOptions::new(message.len() - 1, 1, 1024, 16))
                .is_err()
        );

        let with_two_unknowns = [0x08, 0x01, 0x18, 0x00, 0xa0, 0x06, 0x01, 0xa8, 0x06, 0x02];
        assert!(
            decode_message_info(
                &with_two_unknowns,
                DecodeOptions::new(with_two_unknowns.len(), 2, 1024, 16)
            )
            .is_ok()
        );
        assert!(
            decode_message_info(
                &with_two_unknowns,
                DecodeOptions::new(with_two_unknowns.len(), 1, 1024, 16)
            )
            .is_err()
        );
    }

    #[test]
    fn bounded_encode_does_not_partially_write() {
        let info = tsp::ArchiveInfo {
            identifier: Some(42),
            message_infos: Vec::new(),
            should_merge: None,
        };
        let mut output = vec![0xaa];
        assert!(encode_archive_info(&info, 1, &mut output).is_err());
        assert_eq!(output, vec![0xaa]);
    }

    #[test]
    fn required_message_info_fields_are_enforced() {
        let missing_length = [0x08, 0x01];
        let missing_type = [0x18, 0x01];
        assert_eq!(
            decode_message_info(&missing_length, options(missing_length.len()))
                .err()
                .and_then(|error| error.missing_required()),
            Some("TSP.MessageInfo.length")
        );
        assert_eq!(
            decode_message_info(&missing_type, options(missing_type.len()))
                .err()
                .and_then(|error| error.missing_required()),
            Some("TSP.MessageInfo.type")
        );
    }

    #[test]
    fn required_field_info_path_is_enforced() {
        let missing_path = [0x08, 0x01, 0x18, 0x00, 0x22, 0x00];
        assert_eq!(
            decode_message_info(&missing_path, options(missing_path.len()))
                .err()
                .and_then(|error| error.missing_required()),
            Some("TSP.FieldInfo.path")
        );
    }

    #[test]
    fn closed_enum_projection_matches_prost_int32_wire_semantics()
    -> Result<(), Box<dyn std::error::Error>> {
        let cases: &[&[u8]] = &[
            &[0x03],
            &[0x63],
            &[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x01],
            &[0xff, 0xff, 0xff, 0xff, 0x0f],
        ];
        for encoded_enum in cases {
            let field_length_usize = 3usize
                .checked_add(encoded_enum.len())
                .ok_or("fixture length overflow")?;
            let field_length = u8::try_from(field_length_usize)?;
            let mut source = vec![0x08, 0x01, 0x18, 0x00, 0x22, field_length, 0x0a, 0x00, 0x10];
            source.extend_from_slice(encoded_enum);
            assert_eq!(
                decode_message_info(&source, options(source.len()))?,
                tsp::MessageInfo::decode(source.as_slice())?
            );
        }

        let repeated_cases: &[&[&[u8]]] = &[
            &[&[0x03], &[0x63]],
            &[&[0x63], &[0x03]],
            &[
                &[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x01],
                &[0x63],
            ],
        ];
        for encoded_fields in repeated_cases {
            let field_length_usize = encoded_fields
                .iter()
                .try_fold(2usize, |length, encoded| {
                    length.checked_add(1 + encoded.len())
                })
                .ok_or("fixture length overflow")?;
            let field_length = u8::try_from(field_length_usize)?;
            let mut source = vec![0x08, 0x01, 0x18, 0x00, 0x22, field_length, 0x0a, 0x00];
            for encoded_enum in *encoded_fields {
                source.push(0x10);
                source.extend_from_slice(encoded_enum);
            }
            assert_eq!(
                decode_message_info(&source, options(source.len()))?,
                tsp::MessageInfo::decode(source.as_slice())?
            );
        }
        Ok(())
    }
}
