//! Bounded framing for decompressed IWA archive streams.
//!
//! An IWA stream is a sequence of objects. Each object contains a varint
//! length, a `TSP.ArchiveInfo` protobuf, and the payloads described by that
//! header's `TSP.MessageInfo` entries. This module owns only that physical
//! framing. It does not resolve message types or decode application payloads.

#![allow(
    clippy::cast_possible_truncation,
    reason = "Varint encoding masks each narrowed byte to its seven-bit wire value."
)]
#![allow(
    clippy::map_err_ignore,
    reason = "The public physical-layer error API intentionally normalizes conversion and allocation failures."
)]
#![allow(
    clippy::missing_errors_doc,
    reason = "The module's Result-returning methods share the crate-level physical framing error contract."
)]
#![allow(
    clippy::module_name_repetitions,
    reason = "These names mirror the neutral IWA wire vocabulary and remain source-compatible with the extracted layer."
)]
#![allow(
    clippy::shadow_reuse,
    reason = "Validated limit profiles deliberately retain the public parameter name."
)]

use std::cell::Cell;
use std::collections::HashSet;
use std::io::Read;
use std::mem::size_of;

use litchi_iwa_common::wire::{
    WireDescent, WireField, WirePreflight, WireVisit, parse_wire_fields_with_limits,
    preflight_wire_tree_with_limits,
};
use litchi_iwa_common::{Error as WireError, LimitKind as WireLimitKind, WireLimits};
use litchi_iwa_protos::archive_codec;
use litchi_iwa_protos::tsp;

use crate::{Error, HeaderKind, HeaderOperation, LimitKind, Limits, Result};

const MAX_VARINT_BYTES: usize = 10;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum HeaderNode {
    ArchiveInfo,
    MessageInfo,
    FieldInfo,
    FieldPath,
}

/// Schema-neutral path to a field nested inside an archive payload.
///
/// An empty component list is valid on the wire. Required `FieldInfo.path`
/// presence is enforced while decoding, before this value is published.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct FieldPath {
    /// Field numbers from the payload root to the selected field.
    pub path: Vec<u32>,
}

impl FieldPath {
    /// Construct a field path from its ordered field numbers.
    #[must_use]
    pub const fn new(path: Vec<u32>) -> Self {
        Self { path }
    }

    /// Borrow the ordered field numbers.
    #[must_use]
    pub fn as_slice(&self) -> &[u32] {
        &self.path
    }
}

impl From<Vec<u32>> for FieldPath {
    fn from(path: Vec<u32>) -> Self {
        Self::new(path)
    }
}

/// Kind of value described by field-level archive metadata.
///
/// `Unrecognized` preserves future or producer-specific closed-enum values
/// exactly instead of coercing them to a known default.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FieldType {
    Value,
    ObjectReference,
    DataReference,
    Message,
    Unrecognized(i32),
}

/// Data-reference removal policy for a header-preserving message replacement.
///
/// `Selected` removes every occurrence of the supplied identifiers from the
/// target `MessageInfo` and all of its nested `FieldInfo` values. `All` clears
/// every data reference in that same scope. Duplicate selected identifiers are
/// harmless, and `Selected(&[])` is equivalent to `None`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataReferencePruning<'a> {
    /// Preserve all data references.
    None,
    /// Remove every occurrence of these data-reference identifiers.
    Selected(&'a [u64]),
    /// Remove every data-reference identifier.
    All,
}

impl DataReferencePruning<'_> {
    const fn is_none(self) -> bool {
        matches!(self, Self::None | Self::Selected([]))
    }
}

impl FieldType {
    /// Project one raw protobuf enum value without losing unknown values.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            0 => Self::Value,
            1 => Self::ObjectReference,
            2 => Self::DataReference,
            3 => Self::Message,
            unrecognized => Self::Unrecognized(unrecognized),
        }
    }

    /// Return the exact raw protobuf enum value.
    #[must_use]
    pub const fn raw_value(self) -> i32 {
        match self {
            Self::Value => 0,
            Self::ObjectReference => 1,
            Self::DataReference => 2,
            Self::Message => 3,
            Self::Unrecognized(value) => value,
        }
    }
}

/// Preservation policy for an unknown payload field.
///
/// `Unrecognized` retains future closed-enum values. In particular,
/// `NotSupported` is the known native value `-1`, not an unknown value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum UnknownFieldRule {
    IgnoreAndPreserveUntilModified,
    IgnoreAndPreserve,
    MustUnderstand,
    NotSupported,
    Unrecognized(i32),
}

impl UnknownFieldRule {
    /// Project one raw protobuf enum value without losing unknown values.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            0 => Self::IgnoreAndPreserveUntilModified,
            1 => Self::IgnoreAndPreserve,
            2 => Self::MustUnderstand,
            -1 => Self::NotSupported,
            unrecognized => Self::Unrecognized(unrecognized),
        }
    }

    /// Return the exact raw protobuf enum value.
    #[must_use]
    pub const fn raw_value(self) -> i32 {
        match self {
            Self::IgnoreAndPreserveUntilModified => 0,
            Self::IgnoreAndPreserve => 1,
            Self::MustUnderstand => 2,
            Self::NotSupported => -1,
            Self::Unrecognized(value) => value,
        }
    }
}

/// Preservation policy for a recognized payload field introduced by a newer
/// producer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum KnownFieldRule {
    None,
    PreserveNewerValueUntilModified,
    PreserveNewerValue,
    Unrecognized(i32),
}

impl KnownFieldRule {
    /// Project one raw protobuf enum value without losing unknown values.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            0 => Self::None,
            1 => Self::PreserveNewerValueUntilModified,
            2 => Self::PreserveNewerValue,
            unrecognized => Self::Unrecognized(unrecognized),
        }
    }

    /// Return the exact raw protobuf enum value.
    #[must_use]
    pub const fn raw_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::PreserveNewerValueUntilModified => 1,
            Self::PreserveNewerValue => 2,
            Self::Unrecognized(value) => value,
        }
    }
}

/// Field-level reference and compatibility metadata for one archive payload.
///
/// Optional enum fields retain presence separately from their schema default;
/// this is required for canonical-byte parity after a semantic edit.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FieldInfo {
    pub path: FieldPath,
    pub r#type: Option<FieldType>,
    pub unknown_field_rule: Option<UnknownFieldRule>,
    pub object_references: Vec<u64>,
    pub data_references: Vec<u64>,
    pub known_field_rule: Option<KnownFieldRule>,
    pub known_field_version: Vec<u32>,
    pub known_field_feature_identifier: Option<String>,
}

impl FieldInfo {
    /// Construct metadata for a field path with every optional wire field
    /// absent.
    #[must_use]
    pub fn new(path: impl Into<FieldPath>) -> Self {
        Self {
            path: path.into(),
            ..Self::default()
        }
    }

    /// Resolve the schema default without changing encoded presence.
    #[must_use]
    pub const fn effective_type(&self) -> FieldType {
        match self.r#type {
            Some(value) => value,
            None => FieldType::Value,
        }
    }

    /// Resolve the schema default without changing encoded presence.
    #[must_use]
    pub const fn effective_unknown_field_rule(&self) -> UnknownFieldRule {
        match self.unknown_field_rule {
            Some(value) => value,
            None => UnknownFieldRule::IgnoreAndPreserveUntilModified,
        }
    }

    /// Resolve the schema default without changing encoded presence.
    #[must_use]
    pub const fn effective_known_field_rule(&self) -> KnownFieldRule {
        match self.known_field_rule {
            Some(value) => value,
            None => KnownFieldRule::None,
        }
    }
}

/// Metadata for one object in an IWA component.
#[derive(Debug, Clone, PartialEq)]
pub struct ArchiveInfo {
    /// Unique identifier across the document package.
    pub identifier: Option<u64>,
    /// Metadata for each payload immediately following this header.
    pub message_infos: Vec<MessageInfo>,
    /// Whether this archive should be merged with an existing object.
    pub should_merge: Option<bool>,
}

impl ArchiveInfo {
    /// Decode one bounded `TSP.ArchiveInfo` protobuf from a reader.
    ///
    /// Because protobuf messages are not self-delimiting, callers must pass a
    /// reader bounded to exactly one header when additional bytes follow it.
    pub fn parse<R: Read>(reader: &mut R) -> Result<Self> {
        Self::parse_with_limits(reader, Limits::default())
    }

    /// Decode one `TSP.ArchiveInfo` protobuf from a reader under explicit
    /// resource limits.
    pub fn parse_with_limits<R: Read>(reader: &mut R, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let data = read_bounded(reader, limits.max_header_bytes(), "ArchiveInfo")?;
        Self::decode_with_limits(&data, limits)
    }

    /// Decode one bounded `TSP.ArchiveInfo` header.
    pub fn decode(data: &[u8]) -> Result<Self> {
        Self::decode_with_limits(data, Limits::default())
    }

    /// Decode one `TSP.ArchiveInfo` header under explicit limits.
    pub fn decode_with_limits(data: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_header_length(data.len(), limits)?;
        let preflight = preflight_header(data, HeaderKind::ArchiveInfo, limits)?;
        let decoded =
            archive_codec::decode_archive_info(data, buffa_decode_options(preflight, limits))
                .map_err(|error| {
                    Error::header_codec(
                        HeaderKind::ArchiveInfo,
                        HeaderOperation::Decode,
                        error.to_string(),
                    )
                })?;
        let archive_info = archive_info_from_proto(decoded)?;
        archive_info.validate_with_limits(limits)?;
        Ok(archive_info)
    }

    /// Construct metadata for an object from its raw payload metadata.
    #[must_use]
    pub fn new(identifier: u64, message_infos: Vec<MessageInfo>) -> Self {
        Self {
            identifier: Some(identifier),
            message_infos,
            should_merge: None,
        }
    }

    fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let message_count = self.message_infos.len();
        if message_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                message_count,
                limits.max_messages_per_object(),
            ));
        }

        let mut metadata_items = message_count;
        for message_info in &self.message_infos {
            message_info.validate_with_limits(limits)?;
            add_limited(
                &mut metadata_items,
                message_info.metadata_item_count()?,
                LimitKind::MetadataItems,
                limits.max_metadata_items(),
            )?;
        }
        if self.identifier.is_none() {
            return Err(Error::invalid_archive(
                0,
                "object is missing its archive identifier",
            ));
        }
        Ok(())
    }
}

/// Metadata for one payload in an archive object.
#[derive(Debug, Clone, PartialEq)]
pub struct MessageInfo {
    pub type_: u32,
    pub versions: Vec<u32>,
    pub length: u32,
    pub field_infos: Vec<FieldInfo>,
    pub object_references: Vec<u64>,
    pub data_references: Vec<u64>,
    pub base_message_index: Option<u32>,
    pub diff_merge_version: Vec<u32>,
    pub diff_field_path: Option<FieldPath>,
    pub fields_to_remove: Vec<FieldPath>,
    pub diff_read_version: Vec<u32>,
}

impl MessageInfo {
    /// Construct the conventional metadata for a newly created payload.
    #[must_use]
    pub fn new(type_: u32, length: u32) -> Self {
        Self {
            type_,
            versions: vec![1, 0, 5],
            length,
            field_infos: Vec::new(),
            object_references: Vec::new(),
            data_references: Vec::new(),
            base_message_index: None,
            diff_merge_version: Vec::new(),
            diff_field_path: None,
            fields_to_remove: Vec::new(),
            diff_read_version: Vec::new(),
        }
    }

    fn try_new(type_: u32, length: u32) -> Result<Self> {
        let mut versions = Vec::new();
        versions
            .try_reserve_exact(3)
            .map_err(|_| Error::allocation("IWA message versions", 3))?;
        versions.extend([1, 0, 5]);
        Ok(Self {
            type_,
            versions,
            length,
            field_infos: Vec::new(),
            object_references: Vec::new(),
            data_references: Vec::new(),
            base_message_index: None,
            diff_merge_version: Vec::new(),
            diff_field_path: None,
            fields_to_remove: Vec::new(),
            diff_read_version: Vec::new(),
        })
    }

    /// Decode one bounded `TSP.MessageInfo` protobuf.
    pub fn decode(data: &[u8]) -> Result<Self> {
        Self::decode_with_limits(data, Limits::default())
    }

    /// Decode one `TSP.MessageInfo` protobuf under explicit limits.
    pub fn decode_with_limits(data: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        check_header_length(data.len(), limits)?;
        let preflight = preflight_header(data, HeaderKind::MessageInfo, limits)?;
        let decoded =
            archive_codec::decode_message_info(data, buffa_decode_options(preflight, limits))
                .map_err(|error| {
                    Error::header_codec(
                        HeaderKind::MessageInfo,
                        HeaderOperation::Decode,
                        error.to_string(),
                    )
                })?;
        let message_info = message_info_from_proto(decoded)?;
        message_info.validate_with_limits(limits)?;
        Ok(message_info)
    }

    /// Decode one bounded `TSP.MessageInfo` protobuf from a reader.
    ///
    /// Because protobuf messages are not self-delimiting, callers must pass a
    /// reader bounded to exactly one metadata message when additional bytes
    /// follow it.
    pub fn parse<R: Read>(reader: &mut R) -> Result<Self> {
        Self::parse_with_limits(reader, Limits::default())
    }

    /// Decode one `TSP.MessageInfo` protobuf from a reader under explicit
    /// resource limits.
    pub fn parse_with_limits<R: Read>(reader: &mut R, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let data = read_bounded(reader, limits.max_header_bytes(), "MessageInfo")?;
        Self::decode_with_limits(&data, limits)
    }

    fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let length = usize::try_from(self.length)
            .map_err(|_| Error::invalid_archive(0, "message length exceeds usize"))?;
        if length > limits.max_message_bytes() {
            return Err(limit(
                LimitKind::MessageBytes,
                length,
                limits.max_message_bytes(),
            ));
        }
        let metadata_items = self.metadata_item_count()?;
        if metadata_items > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                metadata_items,
                limits.max_metadata_items(),
            ));
        }
        Ok(())
    }

    fn metadata_item_count(&self) -> Result<usize> {
        let mut count = 0;
        add_count(&mut count, self.versions.len())?;
        add_count(&mut count, self.field_infos.len())?;
        add_count(&mut count, self.object_references.len())?;
        add_count(&mut count, self.data_references.len())?;
        add_count(&mut count, self.diff_merge_version.len())?;
        add_count(&mut count, usize::from(self.diff_field_path.is_some()))?;
        add_count(&mut count, self.fields_to_remove.len())?;
        add_count(&mut count, self.diff_read_version.len())?;
        for field_info in &self.field_infos {
            add_count(&mut count, 1)?;
            add_count(&mut count, field_info.path.path.len())?;
            add_count(&mut count, field_info.object_references.len())?;
            add_count(&mut count, field_info.data_references.len())?;
            add_count(&mut count, field_info.known_field_version.len())?;
            add_count(
                &mut count,
                usize::from(field_info.known_field_feature_identifier.is_some()),
            )?;
        }
        if let Some(path) = &self.diff_field_path {
            add_count(&mut count, path.path.len())?;
        }
        for path in &self.fields_to_remove {
            add_count(&mut count, path.path.len())?;
        }
        Ok(count)
    }
}

/// Raw payload data for one archive message.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawMessage {
    pub type_: u32,
    pub data: Vec<u8>,
}

/// One object within a decompressed IWA archive stream.
#[derive(Debug, Clone, PartialEq)]
pub struct ArchiveObject {
    pub archive_info: ArchiveInfo,
    pub messages: Vec<RawMessage>,
    /// Offset of the object's length prefix in the parsed decompressed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations.
    pub header_offset: u64,
    /// Size of the length prefix and encoded `ArchiveInfo` header in the
    /// parsed decompressed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations.
    pub header_length: u64,
    /// Offset of the first payload in the parsed decompressed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations.
    pub data_offset: u64,
    /// Total bytes occupied by the object's payloads in the parsed source.
    ///
    /// This is source provenance: it is zero for a newly constructed object
    /// and is not recomputed by in-memory message mutations. Current payload
    /// size is available by summing [`Self::messages`].
    pub data_length: u64,
    /// Original encoded header, retained only when it differs from the
    /// canonical encoding so untouched unknown protobuf fields survive a
    /// parse/serialize round trip.
    original_header: Option<Box<[u8]>>,
    /// Canonical encoding paired with `original_header` for a non-canonical
    /// source header. Both fields are absent for a canonical source header.
    original_canonical_header: Option<Box<[u8]>>,
}

impl ArchiveObject {
    /// Construct an object and conventional metadata for its raw payloads.
    pub fn new(identifier: u64, messages: Vec<RawMessage>) -> Result<Self> {
        Self::new_with_limits(identifier, messages, Limits::default())
    }

    /// Construct an object under explicit resource limits.
    pub fn new_with_limits(
        identifier: u64,
        messages: Vec<RawMessage>,
        limits: Limits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        if messages.len() > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                messages.len(),
                limits.max_messages_per_object(),
            ));
        }
        let mut message_infos = Vec::new();
        message_infos
            .try_reserve_exact(messages.len())
            .map_err(|_| Error::allocation("IWA message metadata", messages.len()))?;
        for message in &messages {
            let length = u32::try_from(message.data.len())
                .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
            check_message_length(message.data.len(), limits)?;
            message_infos.push(MessageInfo::try_new(message.type_, length)?);
        }
        let object = Self {
            archive_info: ArchiveInfo::new(identifier, message_infos),
            messages,
            header_offset: 0,
            header_length: 0,
            data_offset: 0,
            data_length: 0,
            original_header: None,
            original_canonical_header: None,
        };
        object.validate_with_limits(limits)?;
        Ok(object)
    }

    /// Validate object metadata, payload sizes, and encoded header size.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_limits(Limits::default())
    }

    /// Validate this object under explicit resource limits.
    pub fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let limits = limits.validate()?;
        if self.archive_info.identifier.is_none() {
            return Err(Error::invalid_archive(
                0,
                "object is missing its archive identifier",
            ));
        }
        if self.archive_info.message_infos.len() != self.messages.len() {
            return Err(Error::invalid_archive(
                0,
                "message metadata and payload counts differ",
            ));
        }
        self.archive_info.validate_with_limits(limits)?;
        let header_length = archive_info_encoded_len(&self.archive_info)?;
        check_header_length(header_length, limits)?;
        let prefix_length = varint_len(header_length)?;
        let mut payload_length = 0;
        for message in &self.messages {
            check_message_length(message.data.len(), limits)?;
            add_limited(
                &mut payload_length,
                message.data.len(),
                LimitKind::ObjectBytes,
                limits.max_object_bytes(),
            )?;
        }
        let object_length = prefix_length
            .checked_add(header_length)
            .and_then(|length| length.checked_add(payload_length))
            .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
        Ok(())
    }

    /// Return the type identifier of the first payload, if present.
    #[must_use]
    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|message| message.type_)
    }

    /// Replace one payload and synchronize its physical metadata atomically.
    pub fn replace_message(&mut self, index: usize, message: RawMessage) -> Result<RawMessage> {
        self.replace_message_with_limits(index, message, Limits::default())
    }

    /// Replace one payload under explicit resource limits.
    pub fn replace_message_with_limits(
        &mut self,
        index: usize,
        message: RawMessage,
        limits: Limits,
    ) -> Result<RawMessage> {
        let limits = limits.validate()?;
        self.validate_with_limits(limits)?;
        let message_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;
        let old_type = message_info.type_;
        let old_length = message_info.length;
        let old = std::mem::replace(
            self.messages
                .get_mut(index)
                .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?,
            message,
        );
        let info = &mut self.archive_info.message_infos[index];
        info.type_ = self.messages[index].type_;
        info.length = length;
        if let Err(error) = self.validate_with_limits(limits) {
            drop(std::mem::replace(&mut self.messages[index], old));
            let restored_info = &mut self.archive_info.message_infos[index];
            restored_info.type_ = old_type;
            restored_info.length = old_length;
            return Err(error);
        }
        Ok(old)
    }

    /// Replace one payload while preserving untouched `ArchiveInfo` bytes.
    ///
    /// Unlike [`Self::replace_message`], this variant surgically updates only
    /// the effective `MessageInfo.type` and `MessageInfo.length` scalar values
    /// plus an enclosing length prefix when its width must change. Unknown
    /// fields, duplicate metadata, non-canonical keys and every unrelated raw
    /// header byte remain untouched. The mutation is atomic: validation or
    /// allocation failure leaves the object unchanged.
    pub fn replace_message_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
    ) -> Result<RawMessage> {
        self.replace_message_preserving_header_with_limits(index, message, Limits::default())
    }

    /// Replace one payload while preserving untouched `ArchiveInfo` bytes
    /// under explicit resource limits.
    pub fn replace_message_preserving_header_with_limits(
        &mut self,
        index: usize,
        message: RawMessage,
        limits: Limits,
    ) -> Result<RawMessage> {
        let limits = limits.validate()?;
        self.validate_with_limits(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match (
            self.original_header.as_deref(),
            self.original_canonical_header.as_deref(),
        ) {
            (Some(original), Some(canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        // Re-apply the caller's current limit profile to retained raw bytes.
        // An object may have been opened under a broader profile.
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_in_header(
            source_header,
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            limits,
        )?;

        let old_type = current_info.type_;
        let old_length = current_info.length;
        let old = self.replace_message_with_limits(index, message, limits)?;
        let verification = (|| {
            let decoded = ArchiveInfo::decode_with_limits(&rewritten_header, limits)?;
            if decoded != self.archive_info {
                return Err(Error::invalid_archive(
                    index,
                    "raw ArchiveInfo rewrite changed unrelated metadata",
                ));
            }
            let canonical = encode_archive_info(&self.archive_info, limits)?;
            let published_header_length = if retained_source_header && rewritten_header != canonical
            {
                rewritten_header.len()
            } else {
                canonical.len()
            };
            validate_raw_object_size(self, published_header_length, limits)?;
            Ok(canonical)
        })();
        let canonical_after = match verification {
            Ok(canonical) => canonical,
            Err(error) => {
                restore_replaced_message(self, index, old, old_type, old_length);
                return Err(error);
            },
        };

        if retained_source_header && rewritten_header != canonical_after {
            self.original_header = Some(rewritten_header.into_boxed_slice());
            self.original_canonical_header = Some(canonical_after.into_boxed_slice());
        } else {
            self.original_header = None;
            self.original_canonical_header = None;
        }
        Ok(old)
    }

    /// Replace one payload and prune selected object references while
    /// preserving every untouched `ArchiveInfo` byte.
    ///
    /// The identifiers are removed from the selected `MessageInfo`'s
    /// top-level object references and from the object references of every
    /// nested `FieldInfo`. Packed and unpacked uint64 fields may be mixed;
    /// retained values keep their original encoding and order. Duplicate
    /// identifiers in `identifiers` are harmless. An empty slice delegates to
    /// [`Self::replace_message_preserving_header`] exactly.
    ///
    /// This physical primitive does not prove graph exclusivity. The caller
    /// must establish that every supplied identifier has no surviving
    /// payload or opaque-field ownership and that removing it from *all* of
    /// the selected message's reference metadata is semantically valid.
    pub fn replace_message_pruning_object_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        identifiers: &[u64],
    ) -> Result<RawMessage> {
        self.replace_message_pruning_object_references_preserving_header_with_limits(
            index,
            message,
            identifiers,
            Limits::default(),
        )
    }

    /// Replace one payload and prune selected object references while
    /// preserving untouched `ArchiveInfo` bytes under explicit limits.
    ///
    /// All raw and neutral metadata is prepared and validated before the
    /// object is mutated, so every error leaves the object unchanged. The
    /// caller retains the graph-ownership proof described by
    /// [`Self::replace_message_pruning_object_references_preserving_header`].
    pub fn replace_message_pruning_object_references_preserving_header_with_limits(
        &mut self,
        index: usize,
        message: RawMessage,
        identifiers: &[u64],
        limits: Limits,
    ) -> Result<RawMessage> {
        self.replace_message_pruning_references_preserving_header_with_limits(
            index,
            message,
            identifiers,
            DataReferencePruning::None,
            limits,
        )
    }

    /// Replace one payload and prune selected object references plus selected
    /// or all data references while preserving every untouched header byte.
    ///
    /// Both reference kinds are pruned from the target `MessageInfo` and all
    /// nested `FieldInfo` values. Packed and unpacked occurrences may be
    /// mixed, and retained values keep their original order and encoding.
    /// Duplicate selected identifiers are harmless. Empty object identifiers
    /// together with [`DataReferencePruning::None`] or
    /// [`DataReferencePruning::Selected`] containing an empty slice delegate
    /// exactly to [`Self::replace_message_preserving_header`].
    ///
    /// The caller must prove that removing every supplied reference is valid
    /// for the surrounding package graph; this physical primitive establishes
    /// only bounded wire and neutral-metadata correctness.
    pub fn replace_message_pruning_references_preserving_header(
        &mut self,
        index: usize,
        message: RawMessage,
        object_identifiers: &[u64],
        data_references: DataReferencePruning<'_>,
    ) -> Result<RawMessage> {
        self.replace_message_pruning_references_preserving_header_with_limits(
            index,
            message,
            object_identifiers,
            data_references,
            Limits::default(),
        )
    }

    /// Replace one payload and prune object and data references while
    /// preserving untouched `ArchiveInfo` bytes under explicit limits.
    ///
    /// The selector inputs, complete source wire tree, projected neutral
    /// metadata, rewritten header, replacement payload, and enclosing object
    /// are all validated before mutation, so every error is atomic.
    pub fn replace_message_pruning_references_preserving_header_with_limits(
        &mut self,
        index: usize,
        message: RawMessage,
        object_identifiers: &[u64],
        data_references: DataReferencePruning<'_>,
        limits: Limits,
    ) -> Result<RawMessage> {
        if object_identifiers.is_empty() && data_references.is_none() {
            return self.replace_message_preserving_header_with_limits(index, message, limits);
        }

        let limits = limits.validate()?;
        let data_identifier_count = match data_references {
            DataReferencePruning::Selected(identifiers) => identifiers.len(),
            DataReferencePruning::None | DataReferencePruning::All => 0,
        };
        let selector_count = object_identifiers
            .len()
            .checked_add(data_identifier_count)
            .ok_or_else(|| Error::invalid_archive(index, "reference selector count overflow"))?;
        if selector_count > limits.max_metadata_items() {
            return Err(limit(
                LimitKind::MetadataItems,
                selector_count,
                limits.max_metadata_items(),
            ));
        }
        self.validate_with_limits(limits)?;
        let current_info = self
            .archive_info
            .message_infos
            .get(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let replacement_length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(index, "message payload exceeds u32"))?;
        check_message_length(message.data.len(), limits)?;

        let mut object_removals = HashSet::new();
        object_removals
            .try_reserve(object_identifiers.len())
            .map_err(|_| {
                Error::allocation("IWA object-reference removal set", object_identifiers.len())
            })?;
        object_removals.extend(object_identifiers.iter().copied());

        let mut selected_data_removals = HashSet::new();
        if let DataReferencePruning::Selected(identifiers) = data_references {
            selected_data_removals
                .try_reserve(identifiers.len())
                .map_err(|_| {
                    Error::allocation("IWA data-reference removal set", identifiers.len())
                })?;
            selected_data_removals.extend(identifiers.iter().copied());
        }
        let data_removals = match data_references {
            DataReferencePruning::None | DataReferencePruning::Selected([]) => {
                ReferenceRemovals::None
            },
            DataReferencePruning::Selected(_) => {
                ReferenceRemovals::Selected(&selected_data_removals)
            },
            DataReferencePruning::All => ReferenceRemovals::All,
        };

        let object_removals = if object_identifiers.is_empty() {
            ReferenceRemovals::None
        } else {
            ReferenceRemovals::Selected(&object_removals)
        };

        let canonical_before = encode_archive_info(&self.archive_info, limits)?;
        let (source_header, retained_source_header) = match (
            self.original_header.as_deref(),
            self.original_canonical_header.as_deref(),
        ) {
            (Some(original), Some(canonical)) if canonical == canonical_before.as_slice() => {
                (original, true)
            },
            _ => (canonical_before.as_slice(), false),
        };
        // The retained header may have been admitted by a broader profile.
        // Revalidate its complete nested wire tree under the caller's limits.
        preflight_header(source_header, HeaderKind::ArchiveInfo, limits)?;
        let rewritten_header = rewrite_message_metadata_and_prune_references_in_header(
            source_header,
            self.archive_info.message_infos.len(),
            index,
            current_info.type_,
            current_info.length,
            message.type_,
            replacement_length,
            object_removals,
            data_removals,
            limits,
        )?;

        let rewritten_info = ArchiveInfo::decode_with_limits(&rewritten_header, limits)?;
        verify_pruned_archive_info(
            &self.archive_info,
            &rewritten_info,
            index,
            message.type_,
            replacement_length,
            object_removals,
            data_removals,
        )?;
        let canonical_after = encode_archive_info(&rewritten_info, limits)?;
        let retain_rewritten_header = retained_source_header && rewritten_header != canonical_after;
        let published_header_length = if retain_rewritten_header {
            rewritten_header.len()
        } else {
            canonical_after.len()
        };
        validate_raw_object_size_with_replacement(
            self,
            index,
            message.data.len(),
            published_header_length,
            limits,
        )?;
        let (original_header, original_canonical_header) = if retain_rewritten_header {
            (
                Some(rewritten_header.into_boxed_slice()),
                Some(canonical_after.into_boxed_slice()),
            )
        } else {
            (None, None)
        };

        let message_slot = self
            .messages
            .get_mut(index)
            .ok_or_else(|| Error::invalid_archive(index, "message index is out of bounds"))?;
        let old = std::mem::replace(message_slot, message);
        self.archive_info = rewritten_info;
        self.original_header = original_header;
        self.original_canonical_header = original_canonical_header;
        Ok(old)
    }

    /// Append one payload and synchronize its physical metadata atomically.
    pub fn push_message(&mut self, message: RawMessage) -> Result<()> {
        self.push_message_with_limits(message, Limits::default())
    }

    /// Append one payload under explicit resource limits.
    pub fn push_message_with_limits(&mut self, message: RawMessage, limits: Limits) -> Result<()> {
        let limits = limits.validate()?;
        self.validate_with_limits(limits)?;
        let next_count = self
            .messages
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "message count overflow"))?;
        if next_count > limits.max_messages_per_object() {
            return Err(limit(
                LimitKind::MessagesPerObject,
                next_count,
                limits.max_messages_per_object(),
            ));
        }
        check_message_length(message.data.len(), limits)?;
        let length = u32::try_from(message.data.len())
            .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
        let message_info = MessageInfo::try_new(message.type_, length)?;
        self.archive_info
            .message_infos
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA message metadata", 1))?;
        self.messages
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA object messages", 1))?;
        self.archive_info.message_infos.push(message_info);
        self.messages.push(message);
        if let Err(error) = self.validate_with_limits(limits) {
            self.archive_info.message_infos.pop();
            self.messages.pop();
            return Err(error);
        }
        Ok(())
    }

    /// Remove one payload and its metadata, returning the payload if present.
    pub fn remove_message(&mut self, index: usize) -> Option<RawMessage> {
        if index >= self.messages.len() || index >= self.archive_info.message_infos.len() {
            return None;
        }
        self.archive_info.message_infos.remove(index);
        Some(self.messages.remove(index))
    }
}

/// A parsed, mutable decompressed IWA component archive.
#[derive(Debug, Clone, PartialEq)]
pub struct Archive {
    pub objects: Vec<ArchiveObject>,
}

impl Archive {
    /// Construct an empty archive.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            objects: Vec::new(),
        }
    }

    /// Parse a decompressed IWA stream.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_with_limits(data, Limits::default())
    }

    /// Parse a decompressed IWA stream under explicit resource limits.
    pub fn parse_with_limits(data: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        if data.len() > limits.max_archive_bytes() {
            return Err(limit(
                LimitKind::ArchiveBytes,
                data.len(),
                limits.max_archive_bytes(),
            ));
        }

        // Archive byte length is not an object-count hint: a single message
        // may consume almost the entire input. Reserving one slot per input
        // byte would turn a large object into a second, avoidable allocation
        // spike. Grow the object list only as validated objects are found.
        let mut objects = Vec::new();
        let mut cursor = 0usize;
        let mut total_messages = 0usize;
        while cursor < data.len() {
            let object_start = cursor;
            if objects.len() >= limits.max_objects() {
                return Err(limit(
                    LimitKind::Objects,
                    objects.len() + 1,
                    limits.max_objects(),
                ));
            }
            let (header_length_u64, prefix_length) = decode_varint(&data[cursor..])?;
            cursor = cursor
                .checked_add(prefix_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "header offset overflow"))?;
            let header_length = usize::try_from(header_length_u64)
                .map_err(|_| Error::invalid_archive(object_start, "header length exceeds usize"))?;
            check_header_length(header_length, limits)?;
            let header_end = cursor
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "header range overflow"))?;
            let header = data.get(cursor..header_end).ok_or_else(|| {
                Error::invalid_archive(object_start, "truncated ArchiveInfo header")
            })?;
            let archive_info = ArchiveInfo::decode_with_limits(header, limits)?;
            let canonical_header = encode_archive_info(&archive_info, limits)?;
            let (original_header, original_canonical_header) =
                if header == canonical_header.as_slice() {
                    (None, None)
                } else {
                    (Some(header.into()), Some(canonical_header.into()))
                };
            cursor = header_end;
            add_limited(
                &mut total_messages,
                archive_info.message_infos.len(),
                LimitKind::Messages,
                limits.max_messages(),
            )?;

            let object_prefix = prefix_length
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(object_start, "object prefix overflow"))?;
            if object_prefix > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_prefix,
                    limits.max_object_bytes(),
                ));
            }
            let mut payload_length = 0usize;
            for message_info in &archive_info.message_infos {
                let length = usize::try_from(message_info.length)
                    .map_err(|_| Error::invalid_archive(cursor, "message length exceeds usize"))?;
                add_limited(
                    &mut payload_length,
                    length,
                    LimitKind::ObjectBytes,
                    limits.max_object_bytes() - object_prefix,
                )?;
            }
            let payload_end = cursor
                .checked_add(payload_length)
                .ok_or_else(|| Error::invalid_archive(cursor, "payload range overflow"))?;
            if payload_end > data.len() {
                return Err(Error::invalid_archive(cursor, "truncated message payload"));
            }

            let mut messages = Vec::new();
            messages
                .try_reserve_exact(archive_info.message_infos.len())
                .map_err(|_| {
                    Error::allocation("IWA object messages", archive_info.message_infos.len())
                })?;
            let payload_start = cursor;
            for message_info in &archive_info.message_infos {
                let length = usize::try_from(message_info.length)
                    .map_err(|_| Error::invalid_archive(cursor, "message length exceeds usize"))?;
                let end = cursor
                    .checked_add(length)
                    .ok_or_else(|| Error::invalid_archive(cursor, "message range overflow"))?;
                let message_data = data
                    .get(cursor..end)
                    .ok_or_else(|| Error::invalid_archive(cursor, "truncated message payload"))?;
                let mut owned = Vec::new();
                owned
                    .try_reserve_exact(length)
                    .map_err(|_| Error::allocation("IWA message payload", length))?;
                owned.extend_from_slice(message_data);
                messages.push(RawMessage {
                    type_: message_info.type_,
                    data: owned,
                });
                cursor = end;
            }
            objects.push(ArchiveObject {
                archive_info,
                messages,
                header_offset: u64::try_from(object_start)
                    .map_err(|_| Error::invalid_archive(object_start, "offset exceeds u64"))?,
                header_length: u64::try_from(prefix_length + header_length)
                    .map_err(|_| Error::invalid_archive(object_start, "header exceeds u64"))?,
                data_offset: u64::try_from(payload_start)
                    .map_err(|_| Error::invalid_archive(payload_start, "offset exceeds u64"))?,
                data_length: u64::try_from(payload_length)
                    .map_err(|_| Error::invalid_archive(payload_start, "payload exceeds u64"))?,
                original_header,
                original_canonical_header,
            });
        }
        let archive = Self { objects };
        archive.validate_with_limits(limits)?;
        Ok(archive)
    }

    /// Validate canonical object framing against the decompressed source.
    ///
    /// This opt-in check requires every object-header length prefix to use the
    /// shortest protobuf varint encoding. It also requires the parsed source
    /// provenance to describe one exact, contiguous partition of
    /// `decompressed_source`: each header starts where the preceding object's
    /// payload ends, each `header_length` includes exactly its prefix and raw
    /// `ArchiveInfo`, each `data_offset` immediately follows that header, and
    /// the final payload ends at the end of the supplied source.
    ///
    /// Header contents are neither decoded nor re-encoded, so unknown fields,
    /// field ordering, and noncanonical encodings *inside* an `ArchiveInfo`
    /// remain accepted and untouched. [`Self::parse`] deliberately remains
    /// permissive about overlong object length prefixes; format mutations can
    /// call this method when they require an unambiguous byte-splice boundary.
    ///
    /// The check performs no allocation and runs in O(objects + messages)
    /// time, visiting metadata only to verify each recorded payload extent.
    pub fn validate_canonical_object_framing(&self, decompressed_source: &[u8]) -> Result<()> {
        let source_length = u64::try_from(decompressed_source.len())
            .map_err(|_| Error::invalid_archive(0, "source length exceeds u64"))?;
        let mut expected_object_offset = 0u64;

        for (object_index, object) in self.objects.iter().enumerate() {
            if object.header_offset != expected_object_offset {
                return Err(Error::invalid_archive(
                    object_index,
                    "object header offset does not follow the preceding payload",
                ));
            }
            let header_offset = usize::try_from(object.header_offset).map_err(|_| {
                Error::invalid_archive(object_index, "object header offset exceeds usize")
            })?;
            let remaining = decompressed_source.get(header_offset..).ok_or_else(|| {
                Error::invalid_archive(header_offset, "object header offset exceeds source")
            })?;
            let (raw_header_length, prefix_length) = decode_varint(remaining)?;
            let header_length = usize::try_from(raw_header_length).map_err(|_| {
                Error::invalid_archive(header_offset, "header length exceeds usize")
            })?;
            if prefix_length != varint_len(header_length)? {
                return Err(Error::invalid_archive(
                    header_offset,
                    "object header length prefix is not canonical",
                ));
            }

            let framed_header_length = prefix_length
                .checked_add(header_length)
                .ok_or_else(|| Error::invalid_archive(header_offset, "header range overflow"))?;
            let data_offset = header_offset
                .checked_add(framed_header_length)
                .ok_or_else(|| Error::invalid_archive(header_offset, "header range overflow"))?;
            if data_offset > decompressed_source.len() {
                return Err(Error::invalid_archive(
                    header_offset,
                    "truncated ArchiveInfo header",
                ));
            }
            let framed_header_length_u64 = u64::try_from(framed_header_length).map_err(|_| {
                Error::invalid_archive(header_offset, "framed header length exceeds u64")
            })?;
            let data_offset_u64 = u64::try_from(data_offset)
                .map_err(|_| Error::invalid_archive(header_offset, "data offset exceeds u64"))?;
            if object.header_length != framed_header_length_u64 {
                return Err(Error::invalid_archive(
                    header_offset,
                    "recorded header length does not match source framing",
                ));
            }
            if object.data_offset != data_offset_u64 {
                return Err(Error::invalid_archive(
                    header_offset,
                    "recorded data offset does not follow the source header",
                ));
            }

            let mut metadata_payload_length = 0u64;
            for message_info in &object.archive_info.message_infos {
                metadata_payload_length = metadata_payload_length
                    .checked_add(u64::from(message_info.length))
                    .ok_or_else(|| {
                        Error::invalid_archive(data_offset, "metadata payload length overflow")
                    })?;
            }
            if object.data_length != metadata_payload_length {
                return Err(Error::invalid_archive(
                    data_offset,
                    "recorded data length does not match message metadata",
                ));
            }
            let data_end = object
                .data_offset
                .checked_add(metadata_payload_length)
                .ok_or_else(|| Error::invalid_archive(data_offset, "payload range overflow"))?;
            if data_end > source_length {
                return Err(Error::invalid_archive(
                    data_offset,
                    "truncated message payload",
                ));
            }
            expected_object_offset = data_end;
        }

        if expected_object_offset != source_length {
            let offset = usize::try_from(expected_object_offset).unwrap_or(usize::MAX);
            return Err(Error::invalid_archive(
                offset,
                "source contains bytes not described by archive objects",
            ));
        }
        Ok(())
    }

    /// Serialize this archive as a decompressed IWA stream.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limits(Limits::default())
    }

    /// Serialize this archive under explicit resource limits.
    pub fn to_bytes_with_limits(&self, limits: Limits) -> Result<Vec<u8>> {
        let limits = limits.validate()?;
        validate_object_set(self.objects.iter(), limits)?;
        let mut output = Vec::new();
        for object in &self.objects {
            let mut info = object.archive_info.clone();
            for (message_info, message) in info.message_infos.iter_mut().zip(&object.messages) {
                message_info.type_ = message.type_;
                message_info.length = u32::try_from(message.data.len())
                    .map_err(|_| Error::invalid_archive(0, "message payload exceeds u32"))?;
            }
            let canonical_header = encode_archive_info(&info, limits)?;
            let header = match (
                object.original_header.as_deref(),
                object.original_canonical_header.as_deref(),
            ) {
                (Some(original), Some(canonical)) if canonical == canonical_header.as_slice() => {
                    original
                },
                _ => canonical_header.as_slice(),
            };
            let header_length = header.len();
            let prefix_length = varint_len(header_length)?;
            let payload_length = object.messages.iter().try_fold(0usize, |total, message| {
                total
                    .checked_add(message.data.len())
                    .ok_or_else(|| Error::invalid_archive(0, "payload length overflow"))
            })?;
            let object_length = prefix_length
                .checked_add(header_length)
                .and_then(|length| length.checked_add(payload_length))
                .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
            if object_length > limits.max_object_bytes() {
                return Err(limit(
                    LimitKind::ObjectBytes,
                    object_length,
                    limits.max_object_bytes(),
                ));
            }
            let next_length = output
                .len()
                .checked_add(object_length)
                .ok_or_else(|| Error::invalid_archive(0, "archive length overflow"))?;
            if next_length > limits.max_archive_bytes() {
                return Err(limit(
                    LimitKind::ArchiveBytes,
                    next_length,
                    limits.max_archive_bytes(),
                ));
            }
            output
                .try_reserve_exact(object_length)
                .map_err(|_| Error::allocation("IWA archive output", object_length))?;
            let mut prefix = [0u8; MAX_VARINT_BYTES];
            let prefix = encode_varint(
                u64::try_from(header_length)
                    .map_err(|_| Error::invalid_archive(0, "header length exceeds u64"))?,
                &mut prefix,
            );
            output.extend_from_slice(prefix);
            output.extend_from_slice(header);
            for message in &object.messages {
                output.extend_from_slice(&message.data);
            }
        }
        Ok(output)
    }

    /// Validate all objects under the default resource limits.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_limits(Limits::default())
    }

    /// Validate all objects under explicit resource limits.
    pub fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        let limits = limits.validate()?;
        validate_object_set(self.objects.iter(), limits)
    }

    /// Insert an object, rejecting missing or duplicate identifiers.
    pub fn insert_object(&mut self, object: ArchiveObject) -> Result<()> {
        self.insert_object_with_limits(object, Limits::default())
    }

    /// Insert an object under explicit resource limits.
    pub fn insert_object_with_limits(
        &mut self,
        object: ArchiveObject,
        limits: Limits,
    ) -> Result<()> {
        let limits = limits.validate()?;
        let identifier = object
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        if self.object(identifier).is_some() {
            return Err(Error::invalid_archive(0, "duplicate object identifier"));
        }
        validate_object_set(std::iter::once(&object).chain(self.objects.iter()), limits)?;
        self.objects
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA archive objects", 1))?;
        self.objects.push(object);
        Ok(())
    }

    /// Find an object by its archive identifier.
    #[must_use]
    pub fn object(&self, identifier: u64) -> Option<&ArchiveObject> {
        self.objects
            .iter()
            .find(|object| object.archive_info.identifier == Some(identifier))
    }

    /// Find an object mutably by its archive identifier.
    #[must_use]
    pub fn object_mut(&mut self, identifier: u64) -> Option<&mut ArchiveObject> {
        self.objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(identifier))
    }

    /// Insert or replace an object, returning the previous object if present.
    pub fn upsert_object(&mut self, object: ArchiveObject) -> Result<Option<ArchiveObject>> {
        self.upsert_object_with_limits(object, Limits::default())
    }

    /// Insert or replace an object under explicit resource limits.
    pub fn upsert_object_with_limits(
        &mut self,
        object: ArchiveObject,
        limits: Limits,
    ) -> Result<Option<ArchiveObject>> {
        let limits = limits.validate()?;
        let identifier = object
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        if let Some(index) = self
            .objects
            .iter()
            .position(|current| current.archive_info.identifier == Some(identifier))
        {
            validate_object_set(
                self.objects.iter().enumerate().map(
                    |(current, item)| {
                        if current == index { &object } else { item }
                    },
                ),
                limits,
            )?;
            return Ok(Some(std::mem::replace(&mut self.objects[index], object)));
        }

        validate_object_set(std::iter::once(&object).chain(self.objects.iter()), limits)?;
        self.objects
            .try_reserve(1)
            .map_err(|_| Error::allocation("IWA archive objects", 1))?;
        self.objects.push(object);
        Ok(None)
    }

    /// Remove an object by its archive identifier.
    pub fn remove_object(&mut self, identifier: u64) -> Option<ArchiveObject> {
        let index = self
            .objects
            .iter()
            .position(|object| object.archive_info.identifier == Some(identifier))?;
        Some(self.objects.remove(index))
    }
}

impl Default for Archive {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug)]
enum HeaderFieldRewrite {
    Retain,
    Remove,
    Varint(u64),
    LengthDelimited(Vec<u8>),
}

#[derive(Clone, Copy)]
enum ReferenceRemovals<'a> {
    None,
    Selected(&'a HashSet<u64>),
    All,
}

impl ReferenceRemovals<'_> {
    const fn is_none(self) -> bool {
        matches!(self, Self::None)
    }

    fn contains(self, identifier: u64) -> bool {
        match self {
            Self::None => false,
            Self::Selected(identifiers) => identifiers.contains(&identifier),
            Self::All => true,
        }
    }
}

fn rewrite_message_metadata_in_header(
    source: &[u8],
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    limits: Limits,
) -> Result<Vec<u8>> {
    if current_type == replacement_type && current_length == replacement_length {
        return try_copy_bytes(source, "IWA preserved ArchiveInfo header");
    }

    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let message_field = fields
        .iter()
        .copied()
        .filter(|field| field.number() == 2 && field.wire_type() == 2)
        .nth(message_index)
        .ok_or_else(|| {
            Error::invalid_archive(
                message_index,
                "message metadata is missing from ArchiveInfo",
            )
        })?;
    let message_source = message_field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let rewritten_message = rewrite_effective_message_scalars(
        message_source,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        wire_limits,
        message_index,
    )?;
    let message_start = message_field.payload_start();
    let message_end = message_field.end();

    let length_prefix = source
        .get(message_field.key_end()..message_start)
        .ok_or_else(|| Error::invalid_archive(message_index, "message prefix range is invalid"))?;
    let mut encoded_length = [0u8; MAX_VARINT_BYTES];
    let encoded_length = encode_varint_with_width(
        u64::try_from(rewritten_message.len()).map_err(|_| {
            Error::invalid_archive(message_index, "message metadata length exceeds u64")
        })?,
        length_prefix.len(),
        &mut encoded_length,
    );
    let output_length = source
        .len()
        .checked_sub(message_source.len())
        .and_then(|length| length.checked_sub(length_prefix.len()))
        .and_then(|length| length.checked_add(encoded_length.len()))
        .and_then(|length| length.checked_add(rewritten_message.len()))
        .ok_or_else(|| Error::invalid_archive(message_index, "ArchiveInfo rewrite overflow"))?;
    check_header_length(output_length, limits)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation("IWA preserved ArchiveInfo header", output_length))?;
    output.extend_from_slice(
        source
            .get(..message_field.key_end())
            .ok_or_else(|| Error::invalid_archive(message_index, "message key range is invalid"))?,
    );
    output.extend_from_slice(encoded_length);
    output.extend_from_slice(&rewritten_message);
    output.extend_from_slice(source.get(message_end..).ok_or_else(|| {
        Error::invalid_archive(message_index, "message metadata range is invalid")
    })?);
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "ArchiveInfo rewrite length mismatch",
        ));
    }
    Ok(output)
}

#[allow(
    clippy::too_many_arguments,
    reason = "the helper keeps both old and new required MessageInfo scalars explicit"
)]
fn rewrite_effective_message_scalars(
    source: &[u8],
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    wire_limits: WireLimits,
    message_index: usize,
) -> Result<Vec<u8>> {
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let type_field = (current_type != replacement_type)
        .then(|| {
            fields
                .iter()
                .rposition(|field| field.number() == 1 && field.wire_type() == 0)
                .ok_or_else(|| {
                    Error::invalid_archive(message_index, "MessageInfo type field is missing")
                })
        })
        .transpose()?;
    let length_field = (current_length != replacement_length)
        .then(|| {
            fields
                .iter()
                .rposition(|field| field.number() == 3 && field.wire_type() == 0)
                .ok_or_else(|| {
                    Error::invalid_archive(message_index, "MessageInfo length field is missing")
                })
        })
        .transpose()?;

    let mut output_length = source.len();
    for (field_index, value) in [
        type_field.map(|index| (index, u64::from(replacement_type))),
        length_field.map(|index| (index, u64::from(replacement_length))),
    ]
    .into_iter()
    .flatten()
    {
        let payload_length = fields[field_index]
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
            .len();
        let encoded_length = encoded_varint_width(value, payload_length);
        output_length = output_length
            .checked_sub(payload_length)
            .and_then(|length| length.checked_add(encoded_length))
            .ok_or_else(|| Error::invalid_archive(message_index, "MessageInfo rewrite overflow"))?;
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation("IWA preserved MessageInfo header", output_length))?;
    for (field_index, field) in fields.iter().copied().enumerate() {
        let replacement = if type_field == Some(field_index) {
            Some(u64::from(replacement_type))
        } else if length_field == Some(field_index) {
            Some(u64::from(replacement_length))
        } else {
            None
        };
        if let Some(value) = replacement {
            output.extend_from_slice(
                field
                    .key(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?,
            );
            let preferred_width = field
                .payload(source)
                .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?
                .len();
            let mut encoded = [0u8; MAX_VARINT_BYTES];
            output.extend_from_slice(encode_varint_with_width(
                value,
                preferred_width,
                &mut encoded,
            ));
        } else {
            output.extend_from_slice(
                field
                    .raw(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?,
            );
        }
    }
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "MessageInfo rewrite length mismatch",
        ));
    }
    Ok(output)
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies source cardinality and both old and new required scalars"
)]
fn rewrite_message_metadata_and_prune_references_in_header(
    source: &[u8],
    expected_message_count: usize,
    message_index: usize,
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let wire_limits = header_wire_limits(limits)?;
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let mut message_count = 0usize;
    let mut target = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(Error::invalid_archive(
                message_index,
                "ArchiveInfo contains an ambiguous MessageInfo field",
            ));
        }
        if message_count == message_index {
            target = Some((field_index, field));
        }
        message_count = message_count.checked_add(1).ok_or_else(|| {
            Error::invalid_archive(message_index, "message metadata count overflow")
        })?;
    }
    if message_count != expected_message_count {
        return Err(Error::invalid_archive(
            message_index,
            "raw and neutral MessageInfo counts differ",
        ));
    }
    let (target_index, target_field) = target.ok_or_else(|| {
        Error::invalid_archive(
            message_index,
            "message metadata is missing from ArchiveInfo",
        )
    })?;
    let message_source = target_field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))?;
    let Some(rewritten_message) = rewrite_effective_message_scalars_and_references(
        message_source,
        current_type,
        current_length,
        replacement_type,
        replacement_length,
        object_removals,
        data_removals,
        wire_limits,
        limits,
        message_index,
    )?
    else {
        return try_copy_bytes(source, "IWA preserved ArchiveInfo header");
    };

    let mut rewrites = retained_field_rewrites(fields.len())?;
    assign_header_field_rewrite(
        &mut rewrites,
        target_index,
        HeaderFieldRewrite::LengthDelimited(rewritten_message),
        message_index,
    )?;
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::ArchiveInfo,
        limits,
        message_index,
        "IWA pruned ArchiveInfo header",
    )
}

#[allow(
    clippy::too_many_arguments,
    reason = "the raw rewrite verifies both old and new required MessageInfo scalars"
)]
fn rewrite_effective_message_scalars_and_references(
    source: &[u8],
    current_type: u32,
    current_length: u32,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Option<Vec<u8>>> {
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let type_field = effective_required_varint_field(
        source,
        &fields,
        1,
        u64::from(current_type),
        message_index,
        "MessageInfo type field is missing",
        "raw and neutral MessageInfo types differ",
    )?;
    let length_field = effective_required_varint_field(
        source,
        &fields,
        3,
        u64::from(current_length),
        message_index,
        "MessageInfo length field is missing",
        "raw and neutral MessageInfo lengths differ",
    )?;

    let mut rewrites = retained_field_rewrites(fields.len())?;
    if current_type != replacement_type {
        assign_header_field_rewrite(
            &mut rewrites,
            type_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_type)),
            message_index,
        )?;
    }
    if current_length != replacement_length {
        assign_header_field_rewrite(
            &mut rewrites,
            length_field,
            HeaderFieldRewrite::Varint(u64::from(replacement_length)),
            message_index,
        )?;
    }

    for (field_index, field) in fields.iter().copied().enumerate() {
        match field.number() {
            4 => {
                if field.wire_type() != 2 {
                    return Err(Error::invalid_archive(
                        message_index,
                        "MessageInfo contains an ambiguous FieldInfo field",
                    ));
                }
                let field_info = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
                if let Some(rewritten) = rewrite_field_info_references(
                    field_info,
                    object_removals,
                    data_removals,
                    wire_limits,
                    limits,
                    message_index,
                )? {
                    assign_header_field_rewrite(
                        &mut rewrites,
                        field_index,
                        HeaderFieldRewrite::LengthDelimited(rewritten),
                        message_index,
                    )?;
                }
            },
            5 if !object_removals.is_none() => {
                let rewrite =
                    rewrite_reference_field(source, field, object_removals, message_index)?;
                assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
            },
            6 if !data_removals.is_none() => {
                let rewrite = rewrite_reference_field(source, field, data_removals, message_index)?;
                assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
            },
            _ => {},
        }
    }

    if rewrites
        .iter()
        .all(|rewrite| matches!(rewrite, HeaderFieldRewrite::Retain))
    {
        return Ok(None);
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA pruned MessageInfo header",
    )
    .map(Some)
}

fn rewrite_field_info_references(
    source: &[u8],
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    wire_limits: WireLimits,
    limits: Limits,
    message_index: usize,
) -> Result<Option<Vec<u8>>> {
    let fields = parse_wire_fields_with_limits(source, wire_limits)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    let mut rewrites = retained_field_rewrites(fields.len())?;
    for (field_index, field) in fields.iter().copied().enumerate() {
        let removals = match field.number() {
            4 if !object_removals.is_none() => Some(object_removals),
            5 if !data_removals.is_none() => Some(data_removals),
            _ => None,
        };
        if let Some(removals) = removals {
            let rewrite = rewrite_reference_field(source, field, removals, message_index)?;
            assign_header_field_rewrite(&mut rewrites, field_index, rewrite, message_index)?;
        }
    }
    if rewrites
        .iter()
        .all(|rewrite| matches!(rewrite, HeaderFieldRewrite::Retain))
    {
        return Ok(None);
    }
    assemble_header_field_rewrites(
        source,
        &fields,
        &rewrites,
        HeaderKind::MessageInfo,
        limits,
        message_index,
        "IWA pruned FieldInfo header",
    )
    .map(Some)
}

fn effective_required_varint_field(
    source: &[u8],
    fields: &[WireField],
    number: u32,
    expected: u64,
    message_index: usize,
    missing_reason: &'static str,
    mismatch_reason: &'static str,
) -> Result<usize> {
    let mut effective = None;
    for (field_index, field) in fields.iter().copied().enumerate() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 0 {
            return Err(Error::invalid_archive(
                message_index,
                "MessageInfo contains an ambiguous required scalar",
            ));
        }
        let payload = field
            .payload(source)
            .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
            .map_err(|_| Error::invalid_archive(message_index, "malformed required scalar"))?;
        if encoded_length != payload.len() {
            return Err(Error::invalid_archive(
                message_index,
                "required scalar has trailing bytes",
            ));
        }
        effective = Some((field_index, value));
    }
    let (field_index, value) =
        effective.ok_or_else(|| Error::invalid_archive(message_index, missing_reason))?;
    if value != expected {
        return Err(Error::invalid_archive(message_index, mismatch_reason));
    }
    Ok(field_index)
}

fn rewrite_reference_field(
    source: &[u8],
    field: WireField,
    removals: ReferenceRemovals<'_>,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let payload = field
        .payload(source)
        .map_err(|error| map_wire_error(error, HeaderKind::MessageInfo))?;
    match field.wire_type() {
        0 => {
            let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(payload)
                .map_err(|_| {
                    Error::invalid_archive(message_index, "malformed unpacked reference")
                })?;
            if encoded_length != payload.len() {
                return Err(Error::invalid_archive(
                    message_index,
                    "unpacked reference has trailing bytes",
                ));
            }
            Ok(if removals.contains(value) {
                HeaderFieldRewrite::Remove
            } else {
                HeaderFieldRewrite::Retain
            })
        },
        2 => rewrite_packed_references(payload, removals, message_index),
        _ => Err(Error::invalid_archive(
            message_index,
            "reference field has an ambiguous wire type",
        )),
    }
}

fn rewrite_packed_references(
    payload: &[u8],
    removals: ReferenceRemovals<'_>,
    message_index: usize,
) -> Result<HeaderFieldRewrite> {
    let mut cursor = 0usize;
    let mut retained_length = 0usize;
    let mut changed = false;
    while cursor < payload.len() {
        let (value, encoded_length) =
            litchi_iwa_common::decode_varint_from_bytes(payload.get(cursor..).ok_or_else(
                || Error::invalid_archive(message_index, "packed reference range is invalid"),
            )?)
            .map_err(|_| Error::invalid_archive(message_index, "malformed packed reference"))?;
        let end = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
        if removals.contains(value) {
            changed = true;
        } else {
            retained_length = retained_length.checked_add(encoded_length).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference length overflow")
            })?;
        }
        cursor = end;
    }
    if !changed {
        return Ok(HeaderFieldRewrite::Retain);
    }

    let mut retained = Vec::new();
    retained
        .try_reserve_exact(retained_length)
        .map_err(|_| Error::allocation("IWA retained packed object references", retained_length))?;
    cursor = 0;
    while cursor < payload.len() {
        let remaining = payload.get(cursor..).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range is invalid")
        })?;
        let (value, encoded_length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| Error::invalid_archive(message_index, "malformed packed reference"))?;
        let end = cursor.checked_add(encoded_length).ok_or_else(|| {
            Error::invalid_archive(message_index, "packed reference range overflow")
        })?;
        if !removals.contains(value) {
            retained.extend_from_slice(payload.get(cursor..end).ok_or_else(|| {
                Error::invalid_archive(message_index, "packed reference range is invalid")
            })?);
        }
        cursor = end;
    }
    if retained.len() != retained_length {
        return Err(Error::invalid_archive(
            message_index,
            "packed reference rewrite length mismatch",
        ));
    }
    Ok(HeaderFieldRewrite::LengthDelimited(retained))
}

fn retained_field_rewrites(length: usize) -> Result<Vec<HeaderFieldRewrite>> {
    let mut rewrites = Vec::new();
    rewrites
        .try_reserve_exact(length)
        .map_err(|_| Error::allocation("IWA header field rewrites", length))?;
    rewrites.extend((0..length).map(|_| HeaderFieldRewrite::Retain));
    Ok(rewrites)
}

fn assign_header_field_rewrite(
    rewrites: &mut [HeaderFieldRewrite],
    index: usize,
    rewrite: HeaderFieldRewrite,
    message_index: usize,
) -> Result<()> {
    let slot = rewrites
        .get_mut(index)
        .ok_or_else(|| Error::invalid_archive(message_index, "header rewrite index is invalid"))?;
    *slot = rewrite;
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "the assembler retains source-bound fields under format-level limits and diagnostics"
)]
fn assemble_header_field_rewrites(
    source: &[u8],
    fields: &[WireField],
    rewrites: &[HeaderFieldRewrite],
    header: HeaderKind,
    limits: Limits,
    message_index: usize,
    resource: &'static str,
) -> Result<Vec<u8>> {
    if fields.len() != rewrites.len() {
        return Err(Error::invalid_archive(
            message_index,
            "header field rewrite count mismatch",
        ));
    }
    let mut output_length = 0usize;
    for (field, rewrite) in fields.iter().copied().zip(rewrites) {
        let field_length = match rewrite {
            HeaderFieldRewrite::Retain => field
                .raw(source)
                .map_err(|error| map_wire_error(error, header))?
                .len(),
            HeaderFieldRewrite::Remove => 0,
            HeaderFieldRewrite::Varint(value) => {
                let key_length = field
                    .key(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                let preferred_width = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                key_length
                    .checked_add(encoded_varint_width(*value, preferred_width))
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "header rewrite overflow")
                    })?
            },
            HeaderFieldRewrite::LengthDelimited(payload) => {
                let key_length = field
                    .key(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                let prefix_width = field
                    .payload_start()
                    .checked_sub(field.key_end())
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "field prefix range is invalid")
                    })?;
                let encoded_length = u64::try_from(payload.len()).map_err(|_| {
                    Error::invalid_archive(message_index, "field payload length exceeds u64")
                })?;
                key_length
                    .checked_add(encoded_varint_width(encoded_length, prefix_width))
                    .and_then(|length| length.checked_add(payload.len()))
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "header rewrite overflow")
                    })?
            },
        };
        output_length = output_length
            .checked_add(field_length)
            .ok_or_else(|| Error::invalid_archive(message_index, "header rewrite overflow"))?;
    }
    check_header_length(output_length, limits)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::allocation(resource, output_length))?;
    for (field, rewrite) in fields.iter().copied().zip(rewrites) {
        match rewrite {
            HeaderFieldRewrite::Retain => output.extend_from_slice(
                field
                    .raw(source)
                    .map_err(|error| map_wire_error(error, header))?,
            ),
            HeaderFieldRewrite::Remove => {},
            HeaderFieldRewrite::Varint(value) => {
                output.extend_from_slice(
                    field
                        .key(source)
                        .map_err(|error| map_wire_error(error, header))?,
                );
                let preferred_width = field
                    .payload(source)
                    .map_err(|error| map_wire_error(error, header))?
                    .len();
                let mut encoded = [0u8; MAX_VARINT_BYTES];
                output.extend_from_slice(encode_varint_with_width(
                    *value,
                    preferred_width,
                    &mut encoded,
                ));
            },
            HeaderFieldRewrite::LengthDelimited(payload) => {
                output.extend_from_slice(
                    field
                        .key(source)
                        .map_err(|error| map_wire_error(error, header))?,
                );
                let prefix_width = field
                    .payload_start()
                    .checked_sub(field.key_end())
                    .ok_or_else(|| {
                        Error::invalid_archive(message_index, "field prefix range is invalid")
                    })?;
                let encoded_length = u64::try_from(payload.len()).map_err(|_| {
                    Error::invalid_archive(message_index, "field payload length exceeds u64")
                })?;
                let mut encoded = [0u8; MAX_VARINT_BYTES];
                output.extend_from_slice(encode_varint_with_width(
                    encoded_length,
                    prefix_width,
                    &mut encoded,
                ));
                output.extend_from_slice(payload);
            },
        }
    }
    if output.len() != output_length {
        return Err(Error::invalid_archive(
            message_index,
            "header rewrite length mismatch",
        ));
    }
    Ok(output)
}

fn encoded_varint_width(value: u64, preferred_width: usize) -> usize {
    let canonical = if value == 0 {
        1
    } else {
        (u64::BITS as usize - value.leading_zeros() as usize).div_ceil(7)
    };
    canonical.max(preferred_width.min(MAX_VARINT_BYTES))
}

fn encode_varint_with_width(
    mut value: u64,
    preferred_width: usize,
    output: &mut [u8; MAX_VARINT_BYTES],
) -> &[u8] {
    let width = encoded_varint_width(value, preferred_width);
    for (index, byte) in output.iter_mut().take(width).enumerate() {
        *byte = u8::try_from(value & 0x7f).unwrap_or_default();
        value >>= 7;
        if index + 1 != width {
            *byte |= 0x80;
        }
    }
    &output[..width]
}

fn header_wire_limits(limits: Limits) -> Result<WireLimits> {
    let scanned_bytes = limits
        .max_header_bytes()
        .saturating_mul(4)
        .min(WireLimits::MAX_INPUT_BYTES);
    WireLimits::default()
        .with_input_bytes(scanned_bytes)
        .and_then(|profile| profile.with_fields(limits.max_header_fields()))
        .and_then(|profile| profile.with_nesting(limits.max_header_nesting()))
        .map_err(|error| map_wire_error(error, HeaderKind::ArchiveInfo))
}

fn try_copy_bytes(source: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    output.extend_from_slice(source);
    Ok(output)
}

fn validate_raw_object_size(
    object: &ArchiveObject,
    header_length: usize,
    limits: Limits,
) -> Result<()> {
    check_header_length(header_length, limits)?;
    let mut object_length = varint_len(header_length)?
        .checked_add(header_length)
        .ok_or_else(|| Error::invalid_archive(0, "object prefix overflow"))?;
    for message in &object.messages {
        object_length = object_length
            .checked_add(message.data.len())
            .ok_or_else(|| Error::invalid_archive(0, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
    }
    Ok(())
}

fn validate_raw_object_size_with_replacement(
    object: &ArchiveObject,
    message_index: usize,
    replacement_length: usize,
    header_length: usize,
    limits: Limits,
) -> Result<()> {
    if message_index >= object.messages.len() {
        return Err(Error::invalid_archive(
            message_index,
            "message index is out of bounds",
        ));
    }
    check_header_length(header_length, limits)?;
    let mut object_length = varint_len(header_length)?
        .checked_add(header_length)
        .ok_or_else(|| Error::invalid_archive(message_index, "object prefix overflow"))?;
    for (index, message) in object.messages.iter().enumerate() {
        let payload_length = if index == message_index {
            replacement_length
        } else {
            message.data.len()
        };
        object_length = object_length
            .checked_add(payload_length)
            .ok_or_else(|| Error::invalid_archive(message_index, "object length overflow"))?;
        if object_length > limits.max_object_bytes() {
            return Err(limit(
                LimitKind::ObjectBytes,
                object_length,
                limits.max_object_bytes(),
            ));
        }
    }
    Ok(())
}

fn verify_pruned_archive_info(
    before: &ArchiveInfo,
    after: &ArchiveInfo,
    message_index: usize,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
) -> Result<()> {
    if before.identifier != after.identifier
        || before.should_merge != after.should_merge
        || before.message_infos.len() != after.message_infos.len()
    {
        return Err(Error::invalid_archive(
            message_index,
            "raw ArchiveInfo rewrite changed unrelated metadata",
        ));
    }
    let before_target = before
        .message_infos
        .get(message_index)
        .ok_or_else(|| Error::invalid_archive(message_index, "message index is out of bounds"))?;
    let after_target = after.message_infos.get(message_index).ok_or_else(|| {
        Error::invalid_archive(message_index, "rewritten message metadata is missing")
    })?;
    for (index, (before_info, after_info)) in before
        .message_infos
        .iter()
        .zip(&after.message_infos)
        .enumerate()
    {
        if index != message_index && before_info != after_info {
            return Err(Error::invalid_archive(
                message_index,
                "raw ArchiveInfo rewrite changed another MessageInfo",
            ));
        }
    }
    verify_pruned_message_info(
        before_target,
        after_target,
        replacement_type,
        replacement_length,
        object_removals,
        data_removals,
        message_index,
    )
}

fn verify_pruned_message_info(
    before: &MessageInfo,
    after: &MessageInfo,
    replacement_type: u32,
    replacement_length: u32,
    object_removals: ReferenceRemovals<'_>,
    data_removals: ReferenceRemovals<'_>,
    message_index: usize,
) -> Result<()> {
    if after.type_ != replacement_type
        || after.length != replacement_length
        || before.versions != after.versions
        || before.base_message_index != after.base_message_index
        || before.diff_merge_version != after.diff_merge_version
        || before.diff_field_path != after.diff_field_path
        || before.fields_to_remove != after.fields_to_remove
        || before.diff_read_version != after.diff_read_version
        || before.field_infos.len() != after.field_infos.len()
        || !references_match_after_pruning(
            &before.object_references,
            &after.object_references,
            object_removals,
        )
        || !references_match_after_pruning(
            &before.data_references,
            &after.data_references,
            data_removals,
        )
    {
        return Err(Error::invalid_archive(
            message_index,
            "rewritten MessageInfo does not match the requested edit",
        ));
    }

    for (before_field, after_field) in before.field_infos.iter().zip(&after.field_infos) {
        if before_field.path != after_field.path
            || before_field.r#type != after_field.r#type
            || before_field.unknown_field_rule != after_field.unknown_field_rule
            || before_field.known_field_rule != after_field.known_field_rule
            || before_field.known_field_version != after_field.known_field_version
            || before_field.known_field_feature_identifier
                != after_field.known_field_feature_identifier
            || !references_match_after_pruning(
                &before_field.object_references,
                &after_field.object_references,
                object_removals,
            )
            || !references_match_after_pruning(
                &before_field.data_references,
                &after_field.data_references,
                data_removals,
            )
        {
            return Err(Error::invalid_archive(
                message_index,
                "rewritten FieldInfo does not match the requested reference pruning",
            ));
        }
    }
    Ok(())
}

fn references_match_after_pruning(
    before: &[u64],
    after: &[u64],
    removals: ReferenceRemovals<'_>,
) -> bool {
    before
        .iter()
        .copied()
        .filter(|identifier| !removals.contains(*identifier))
        .eq(after.iter().copied())
}

fn restore_replaced_message(
    object: &mut ArchiveObject,
    index: usize,
    old: RawMessage,
    old_type: u32,
    old_length: u32,
) {
    drop(std::mem::replace(&mut object.messages[index], old));
    let info = &mut object.archive_info.message_infos[index];
    info.type_ = old_type;
    info.length = old_length;
}

fn archive_info_from_proto(value: tsp::ArchiveInfo) -> Result<ArchiveInfo> {
    let mut message_infos = Vec::new();
    message_infos
        .try_reserve_exact(value.message_infos.len())
        .map_err(|_| {
            Error::allocation("IWA neutral message metadata", value.message_infos.len())
        })?;
    for message_info in value.message_infos {
        message_infos.push(message_info_from_proto(message_info)?);
    }
    Ok(ArchiveInfo {
        identifier: value.identifier,
        message_infos,
        should_merge: value.should_merge,
    })
}

fn archive_info_to_proto(value: &ArchiveInfo) -> Result<tsp::ArchiveInfo> {
    let mut message_infos = Vec::new();
    message_infos
        .try_reserve_exact(value.message_infos.len())
        .map_err(|_| {
            Error::allocation(
                "IWA compatibility message metadata",
                value.message_infos.len(),
            )
        })?;
    for message_info in &value.message_infos {
        message_infos.push(message_info_to_proto(message_info)?);
    }
    Ok(tsp::ArchiveInfo {
        identifier: value.identifier,
        message_infos,
        should_merge: value.should_merge,
    })
}

fn message_info_from_proto(value: tsp::MessageInfo) -> Result<MessageInfo> {
    let mut field_infos = Vec::new();
    field_infos
        .try_reserve_exact(value.field_infos.len())
        .map_err(|_| Error::allocation("IWA neutral field metadata", value.field_infos.len()))?;
    for field_info in value.field_infos {
        field_infos.push(field_info_from_proto(field_info));
    }

    let mut fields_to_remove = Vec::new();
    fields_to_remove
        .try_reserve_exact(value.fields_to_remove.len())
        .map_err(|_| {
            Error::allocation(
                "IWA neutral removed field paths",
                value.fields_to_remove.len(),
            )
        })?;
    for field_path in value.fields_to_remove {
        fields_to_remove.push(field_path_from_proto(field_path));
    }

    Ok(MessageInfo {
        type_: value.r#type,
        versions: value.version,
        length: value.length,
        field_infos,
        object_references: value.object_references,
        data_references: value.data_references,
        base_message_index: value.base_message_index,
        diff_merge_version: value.diff_merge_version,
        diff_field_path: value.diff_field_path.map(field_path_from_proto),
        fields_to_remove,
        diff_read_version: value.diff_read_version,
    })
}

fn message_info_to_proto(value: &MessageInfo) -> Result<tsp::MessageInfo> {
    let mut field_infos = Vec::new();
    field_infos
        .try_reserve_exact(value.field_infos.len())
        .map_err(|_| {
            Error::allocation("IWA compatibility field metadata", value.field_infos.len())
        })?;
    for field_info in &value.field_infos {
        field_infos.push(field_info_to_proto(field_info)?);
    }

    let mut fields_to_remove = Vec::new();
    fields_to_remove
        .try_reserve_exact(value.fields_to_remove.len())
        .map_err(|_| {
            Error::allocation(
                "IWA compatibility removed field paths",
                value.fields_to_remove.len(),
            )
        })?;
    for field_path in &value.fields_to_remove {
        fields_to_remove.push(field_path_to_proto(field_path)?);
    }

    Ok(tsp::MessageInfo {
        r#type: value.type_,
        version: try_copy_slice(&value.versions, "IWA compatibility message versions")?,
        length: value.length,
        field_infos,
        object_references: try_copy_slice(
            &value.object_references,
            "IWA compatibility object references",
        )?,
        data_references: try_copy_slice(
            &value.data_references,
            "IWA compatibility data references",
        )?,
        base_message_index: value.base_message_index,
        diff_merge_version: try_copy_slice(
            &value.diff_merge_version,
            "IWA compatibility diff merge versions",
        )?,
        diff_field_path: value
            .diff_field_path
            .as_ref()
            .map(field_path_to_proto)
            .transpose()?,
        fields_to_remove,
        diff_read_version: try_copy_slice(
            &value.diff_read_version,
            "IWA compatibility diff read versions",
        )?,
    })
}

fn field_info_from_proto(value: tsp::FieldInfo) -> FieldInfo {
    FieldInfo {
        path: field_path_from_proto(value.path),
        r#type: value.r#type.map(FieldType::from_raw),
        unknown_field_rule: value.unknown_field_rule.map(UnknownFieldRule::from_raw),
        object_references: value.object_references,
        data_references: value.data_references,
        known_field_rule: value.known_field_rule.map(KnownFieldRule::from_raw),
        known_field_version: value.known_field_version,
        known_field_feature_identifier: value.known_field_feature_identifier,
    }
}

fn field_info_to_proto(value: &FieldInfo) -> Result<tsp::FieldInfo> {
    Ok(tsp::FieldInfo {
        path: field_path_to_proto(&value.path)?,
        r#type: value.r#type.map(FieldType::raw_value),
        unknown_field_rule: value.unknown_field_rule.map(UnknownFieldRule::raw_value),
        object_references: try_copy_slice(
            &value.object_references,
            "IWA compatibility field object references",
        )?,
        data_references: try_copy_slice(
            &value.data_references,
            "IWA compatibility field data references",
        )?,
        known_field_rule: value.known_field_rule.map(KnownFieldRule::raw_value),
        known_field_version: try_copy_slice(
            &value.known_field_version,
            "IWA compatibility known field versions",
        )?,
        known_field_feature_identifier: value
            .known_field_feature_identifier
            .as_deref()
            .map(|identifier| {
                try_copy_string(identifier, "IWA compatibility field feature identifier")
            })
            .transpose()?,
    })
}

fn field_path_from_proto(value: tsp::FieldPath) -> FieldPath {
    FieldPath { path: value.path }
}

fn field_path_to_proto(value: &FieldPath) -> Result<tsp::FieldPath> {
    Ok(tsp::FieldPath {
        path: try_copy_slice(&value.path, "IWA compatibility field path")?,
    })
}

fn try_copy_slice<T: Copy>(source: &[T], resource: &'static str) -> Result<Vec<T>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    output.extend_from_slice(source);
    Ok(output)
}

fn try_copy_string(source: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::allocation(resource, source.len()))?;
    output.push_str(source);
    Ok(output)
}

fn check_header_length(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_header_bytes() {
        return Err(limit(
            LimitKind::HeaderBytes,
            length,
            limits.max_header_bytes(),
        ));
    }
    Ok(())
}

fn encode_archive_info(info: &ArchiveInfo, limits: Limits) -> Result<Vec<u8>> {
    let encoded_info = archive_info_to_proto(info)?;
    let header_length = usize::try_from(
        archive_codec::archive_info_encoded_len(&encoded_info).map_err(|error| {
            Error::header_codec(
                HeaderKind::ArchiveInfo,
                HeaderOperation::Encode,
                error.to_string(),
            )
        })?,
    )
    .map_err(|_| Error::invalid_archive(0, "ArchiveInfo header length exceeds usize"))?;
    check_header_length(header_length, limits)?;
    let mut header = Vec::new();
    header
        .try_reserve_exact(header_length)
        .map_err(|_| Error::allocation("IWA ArchiveInfo header", header_length))?;
    let maximum = u32::try_from(limits.max_header_bytes())
        .map_err(|_| Error::invalid_archive(0, "header byte limit exceeds u32"))?;
    archive_codec::encode_archive_info(&encoded_info, maximum, &mut header).map_err(|error| {
        Error::header_codec(
            HeaderKind::ArchiveInfo,
            HeaderOperation::Encode,
            error.to_string(),
        )
    })?;
    Ok(header)
}

fn archive_info_encoded_len(info: &ArchiveInfo) -> Result<usize> {
    let compatibility = archive_info_to_proto(info)?;
    usize::try_from(
        archive_codec::archive_info_encoded_len(&compatibility).map_err(|error| {
            Error::header_codec(
                HeaderKind::ArchiveInfo,
                HeaderOperation::Encode,
                error.to_string(),
            )
        })?,
    )
    .map_err(|_| Error::invalid_archive(0, "ArchiveInfo header length exceeds usize"))
}

fn buffa_decode_options(preflight: WirePreflight, limits: Limits) -> archive_codec::DecodeOptions {
    archive_codec::DecodeOptions::new(
        limits.max_header_bytes(),
        preflight.fields(),
        limits.max_header_memory_bytes(),
        u32::try_from(limits.max_header_nesting()).unwrap_or(u32::MAX),
    )
}

fn preflight_header(data: &[u8], header: HeaderKind, limits: Limits) -> Result<WirePreflight> {
    let root = match header {
        HeaderKind::ArchiveInfo => HeaderNode::ArchiveInfo,
        HeaderKind::MessageInfo => HeaderNode::MessageInfo,
    };
    let scanned_bytes = limits
        .max_header_bytes()
        .saturating_mul(4)
        .min(WireLimits::MAX_INPUT_BYTES);
    let wire_limits = WireLimits::default()
        .with_input_bytes(scanned_bytes)
        .and_then(|profile| profile.with_fields(limits.max_header_fields()))
        .and_then(|profile| profile.with_nesting(limits.max_header_nesting()))
        .map_err(|error| map_wire_error(error, header))?;
    let decoded_memory = Cell::new(0usize);
    let metadata_items = Cell::new(0usize);
    let message_infos = Cell::new(0usize);

    let preflight = preflight_wire_tree_with_limits(data, wire_limits, |visit| {
        let node = node_at_path(root, visit.path());
        if node == Some(HeaderNode::ArchiveInfo)
            && visit.field().number() == 2
            && visit.field().wire_type() == 2
        {
            message_infos.set(message_infos.get().saturating_add(1));
        }
        decoded_memory.set(
            decoded_memory
                .get()
                .saturating_add(field_memory_charge(node, visit)),
        );
        metadata_items.set(
            metadata_items
                .get()
                .saturating_add(metadata_item_charge(node, visit)?),
        );
        Ok(descent_for(node, visit.field().number()))
    })
    .map_err(|error| map_wire_error(error, header))?;

    if message_infos.get() > limits.max_messages_per_object() {
        return Err(limit(
            LimitKind::MessagesPerObject,
            message_infos.get(),
            limits.max_messages_per_object(),
        ));
    }
    if decoded_memory.get() > limits.max_header_memory_bytes() {
        return Err(limit(
            LimitKind::HeaderMemoryBytes,
            decoded_memory.get(),
            limits.max_header_memory_bytes(),
        ));
    }
    if metadata_items.get() > limits.max_metadata_items() {
        return Err(limit(
            LimitKind::MetadataItems,
            metadata_items.get(),
            limits.max_metadata_items(),
        ));
    }
    Ok(preflight)
}

fn node_at_path(root: HeaderNode, path: &[u32]) -> Option<HeaderNode> {
    path.iter()
        .try_fold(root, |node, field| match (node, field) {
            (HeaderNode::ArchiveInfo, 2) => Some(HeaderNode::MessageInfo),
            (HeaderNode::MessageInfo, 4) => Some(HeaderNode::FieldInfo),
            (HeaderNode::MessageInfo, 9 | 10) | (HeaderNode::FieldInfo, 1) => {
                Some(HeaderNode::FieldPath)
            },
            _ => None,
        })
}

const fn descent_for(node: Option<HeaderNode>, field: u32) -> WireDescent {
    if matches!(
        (node, field),
        (Some(HeaderNode::ArchiveInfo), 2)
            | (Some(HeaderNode::MessageInfo), 4 | 9 | 10)
            | (Some(HeaderNode::FieldInfo), 1)
    ) {
        WireDescent::Descend
    } else {
        WireDescent::Skip
    }
}

fn field_memory_charge(node: Option<HeaderNode>, visit: WireVisit<'_, '_>) -> usize {
    let field = visit.field();
    let number = field.number();
    let wire_type = field.wire_type();
    let scalar = |width: usize| match wire_type {
        0 => width.saturating_mul(2),
        2 => field
            .payload()
            .len()
            .saturating_mul(width)
            .saturating_mul(2),
        _ => 0,
    };
    let message = |width: usize| {
        if wire_type == 2 {
            width.saturating_mul(2).saturating_add(size_of::<&[u8]>())
        } else {
            0
        }
    };
    let projected_message = |compatibility_width: usize, neutral_width: usize| {
        if wire_type == 2 {
            message(compatibility_width).saturating_add(neutral_width)
        } else {
            0
        }
    };

    match (node, number) {
        (Some(HeaderNode::ArchiveInfo), 2) => message(size_of::<MessageInfo>()),
        (Some(HeaderNode::MessageInfo), 2 | 8 | 11) => scalar(size_of::<u32>()),
        (Some(HeaderNode::MessageInfo), 5 | 6) | (Some(HeaderNode::FieldInfo), 4 | 5) => {
            scalar(size_of::<u64>())
        },
        (Some(HeaderNode::MessageInfo), 4) => {
            projected_message(size_of::<tsp::FieldInfo>(), size_of::<FieldInfo>())
        },
        (Some(HeaderNode::MessageInfo), 9) | (Some(HeaderNode::FieldInfo), 1) => {
            message(size_of::<tsp::FieldPath>())
        },
        (Some(HeaderNode::MessageInfo), 10) => {
            projected_message(size_of::<tsp::FieldPath>(), size_of::<FieldPath>())
        },
        (Some(HeaderNode::FieldInfo), 7) | (Some(HeaderNode::FieldPath), 1) => {
            scalar(size_of::<u32>())
        },
        (Some(HeaderNode::FieldInfo), 8) if wire_type == 2 => field.payload().len(),
        (Some(HeaderNode::FieldInfo), 2 | 3 | 6) => field_info_enum_memory_charge(number, field),
        (Some(HeaderNode::ArchiveInfo), 1 | 3) | (Some(HeaderNode::MessageInfo), 1 | 3 | 7) => 0,
        _ => unknown_field_memory_charge(field),
    }
}

fn field_info_enum_memory_charge(
    number: u32,
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> usize {
    let value = (field.wire_type() == 0)
        .then(|| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
        .flatten()
        .map(|(value, _encoded_length)| value as i32);
    let known = matches!(
        (number, value),
        (2, Some(0..=3)) | (3, Some(-1..=2)) | (6, Some(0..=2))
    );
    if known {
        0
    } else {
        unknown_field_memory_charge(field)
    }
}

fn unknown_field_memory_charge(field: litchi_iwa_common::wire::WireFieldView<'_>) -> usize {
    field
        .raw()
        .len()
        .saturating_add(2 * size_of::<[usize; 6]>())
}

fn metadata_item_charge(
    node: Option<HeaderNode>,
    visit: WireVisit<'_, '_>,
) -> litchi_iwa_common::Result<usize> {
    let field = visit.field();
    let repeated = || repeated_varint_count(field);
    match (node, field.number()) {
        // ArchiveInfo validation includes one item for each MessageInfo.
        (Some(HeaderNode::MessageInfo), 2 | 5 | 6 | 8 | 11)
        | (Some(HeaderNode::FieldInfo), 4 | 5 | 7)
        | (Some(HeaderNode::FieldPath), 1) => repeated(),
        // A FieldInfo is counted once in the repeated collection and once
        // again when its nested metadata is traversed by validation.
        (Some(HeaderNode::MessageInfo), 4) if field.wire_type() == 2 => Ok(2),
        (Some(HeaderNode::ArchiveInfo), 2)
        | (Some(HeaderNode::MessageInfo), 9 | 10)
        | (Some(HeaderNode::FieldInfo), 8)
            if field.wire_type() == 2 =>
        {
            Ok(1)
        },
        _ => Ok(0),
    }
}

fn repeated_varint_count(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> litchi_iwa_common::Result<usize> {
    match field.wire_type() {
        0 => Ok(1),
        2 => {
            let mut remaining = field.payload();
            let mut count = 0usize;
            while !remaining.is_empty() {
                let (_, length) = litchi_iwa_common::decode_varint_from_bytes(remaining)
                    .map_err(|error| WireError::InvalidFormat(error.to_string()))?;
                remaining = &remaining[length..];
                count = count.saturating_add(1);
            }
            Ok(count)
        },
        _ => Ok(0),
    }
}

fn map_wire_error(error: WireError, header: HeaderKind) -> Error {
    match error {
        WireError::LimitExceeded {
            kind: WireLimitKind::Fields,
            observed,
            limit: maximum,
        } => limit(LimitKind::HeaderFields, observed, maximum),
        WireError::LimitExceeded {
            kind: WireLimitKind::Nesting,
            observed,
            limit: maximum,
        } => limit(LimitKind::HeaderNesting, observed, maximum),
        WireError::LimitExceeded {
            kind: WireLimitKind::InputBytes,
            observed,
            limit: maximum,
        } => limit(LimitKind::HeaderBytes, observed, maximum),
        WireError::Allocation { resource, amount } => Error::allocation(resource, amount),
        other @ (WireError::InvalidFormat(_)
        | WireError::LimitExceeded { .. }
        | WireError::InvalidLimit { .. }) => {
            Error::header_codec(header, HeaderOperation::Decode, other.to_string())
        },
    }
}

fn read_bounded<R: Read>(
    reader: &mut R,
    maximum: usize,
    resource: &'static str,
) -> Result<Vec<u8>> {
    let mut data = Vec::new();
    let mut buffer = [0u8; 8192];
    loop {
        let remaining = maximum.saturating_sub(data.len());
        let read_size = remaining.min(buffer.len()).max(1);
        let read = reader.read(&mut buffer[..read_size])?;
        if read == 0 {
            return Ok(data);
        }
        if data.len() == maximum {
            return Err(limit(
                LimitKind::HeaderBytes,
                maximum.saturating_add(read),
                maximum,
            ));
        }
        data.try_reserve(read)
            .map_err(|_| Error::allocation(resource, read))?;
        data.extend_from_slice(&buffer[..read]);
    }
}

fn validate_object_set<'a, I>(objects: I, limits: Limits) -> Result<()>
where
    I: IntoIterator<Item = &'a ArchiveObject>,
{
    let limits = limits.validate()?;
    let objects = objects.into_iter();
    let mut identifiers = HashSet::new();
    let initial_capacity = objects.size_hint().0.min(limits.max_objects());
    identifiers
        .try_reserve(initial_capacity)
        .map_err(|_| Error::allocation("IWA archive identifiers", initial_capacity))?;
    let mut object_count = 0usize;
    let mut total_messages = 0usize;
    let mut total_payload = 0usize;
    for object in objects {
        object_count = object_count
            .checked_add(1)
            .ok_or_else(|| Error::invalid_archive(0, "object count overflow"))?;
        if object_count > limits.max_objects() {
            return Err(limit(
                LimitKind::Objects,
                object_count,
                limits.max_objects(),
            ));
        }
        let identifier = object
            .archive_info
            .identifier
            .ok_or_else(|| Error::invalid_archive(0, "object is missing its archive identifier"))?;
        if !identifiers.insert(identifier) {
            return Err(Error::invalid_archive(0, "duplicate object identifier"));
        }
        object.validate_with_limits(limits)?;
        add_limited(
            &mut total_messages,
            object.messages.len(),
            LimitKind::Messages,
            limits.max_messages(),
        )?;
        for message in &object.messages {
            add_limited(
                &mut total_payload,
                message.data.len(),
                LimitKind::ArchiveBytes,
                limits.max_archive_bytes(),
            )?;
        }
    }
    Ok(())
}

fn check_message_length(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_message_bytes() {
        return Err(limit(
            LimitKind::MessageBytes,
            length,
            limits.max_message_bytes(),
        ));
    }
    Ok(())
}

fn add_count(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::invalid_archive(0, "metadata count overflow"))?;
    Ok(())
}

fn add_limited(total: &mut usize, amount: usize, kind: LimitKind, maximum: usize) -> Result<()> {
    let observed = total
        .checked_add(amount)
        .ok_or_else(|| Error::invalid_archive(0, "resource count overflow"))?;
    if observed > maximum {
        return Err(limit(kind, observed, maximum));
    }
    *total = observed;
    Ok(())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

fn decode_varint(data: &[u8]) -> Result<(u64, usize)> {
    let mut value = 0u64;
    for (index, byte) in data.iter().copied().take(MAX_VARINT_BYTES).enumerate() {
        if index == MAX_VARINT_BYTES - 1 && byte > 1 {
            return Err(Error::invalid_archive(
                index,
                "archive length varint overflow",
            ));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, index + 1));
        }
    }
    if data.len() < MAX_VARINT_BYTES {
        Err(Error::invalid_archive(
            data.len(),
            "truncated archive length varint",
        ))
    } else {
        Err(Error::invalid_archive(
            MAX_VARINT_BYTES - 1,
            "archive length varint overflow",
        ))
    }
}

fn varint_len(value: usize) -> Result<usize> {
    let value =
        u64::try_from(value).map_err(|_| Error::invalid_archive(0, "length exceeds u64"))?;
    Ok(if value == 0 {
        1
    } else {
        (u64::BITS as usize - value.leading_zeros() as usize).div_ceil(7)
    })
}

fn encode_varint(mut value: u64, output: &mut [u8; MAX_VARINT_BYTES]) -> &[u8] {
    let mut length = 0;
    while value >= 0x80 {
        output[length] = u8::try_from(value & 0x7f).unwrap_or_default() | 0x80;
        value >>= 7;
        length += 1;
    }
    output[length] = u8::try_from(value).unwrap_or_default();
    &output[..=length]
}

#[cfg(test)]
mod tests {
    use super::{
        Archive, ArchiveObject, DataReferencePruning, Error, FieldInfo, RawMessage,
        encode_archive_info, encode_varint, encode_varint_with_width,
    };
    use crate::{LimitKind, Limits, Result};

    struct ReferencePruningFixture {
        source: Vec<u8>,
        header: Vec<u8>,
        expected_header: Vec<u8>,
    }

    struct CombinedReferencePruningFixture {
        source: Vec<u8>,
        selected_header: Vec<u8>,
        all_data_header: Vec<u8>,
    }

    #[test]
    fn canonical_headers_use_compact_retention_path() -> Result<()> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![3, 4],
                }],
            )?],
        };
        let encoded = archive.to_bytes()?;
        let parsed = Archive::parse(&encoded)?;
        let [object] = parsed.objects.as_slice() else {
            return Err(Error::invalid_archive(
                0,
                "test archive did not contain exactly one object",
            ));
        };

        assert!(object.original_header.is_none());
        assert!(object.original_canonical_header.is_none());
        assert_eq!(parsed.to_bytes()?, encoded);
        Ok(())
    }

    #[test]
    fn canonical_object_framing_accepts_unknown_raw_archive_info_bytes() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let archive = Archive::parse(&fixture.source)?;
        archive.validate_canonical_object_framing(&fixture.source)?;
        assert_eq!(archive.to_bytes()?, fixture.source);
        assert_eq!(
            archive.objects[0].original_header.as_deref(),
            Some(fixture.header.as_slice())
        );
        Ok(())
    }

    #[test]
    fn canonical_object_framing_rejects_overlong_prefix_on_any_object() -> Result<()> {
        let canonical = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 2,
                        data: vec![3],
                    }],
                )?,
                ArchiveObject::new(
                    4,
                    vec![RawMessage {
                        type_: 5,
                        data: vec![6, 7],
                    }],
                )?,
            ],
        }
        .to_bytes()?;
        let parsed = Archive::parse(&canonical)?;
        parsed.validate_canonical_object_framing(&canonical)?;
        let second_offset = usize::try_from(parsed.objects[1].header_offset)
            .map_err(|_| Error::invalid_archive(0, "test offset exceeds usize"))?;
        let (header_length, prefix_length) = super::decode_varint(
            canonical
                .get(second_offset..)
                .ok_or_else(|| Error::invalid_archive(0, "test object offset is invalid"))?,
        )?;
        assert_eq!(prefix_length, 1);
        assert!(header_length < 0x80);

        let mut overlong = Vec::new();
        overlong
            .try_reserve_exact(canonical.len() + 1)
            .map_err(|_| Error::allocation("test overlong archive", canonical.len() + 1))?;
        overlong.extend_from_slice(&canonical[..second_offset]);
        overlong.push(canonical[second_offset] | 0x80);
        overlong.push(0);
        overlong.extend_from_slice(&canonical[second_offset + 1..]);

        // Parsing remains deliberately permissive; only the opt-in framing
        // validator refuses the overlong second-object prefix.
        let permissive = Archive::parse(&overlong)?;
        assert!(
            permissive
                .validate_canonical_object_framing(&overlong)
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn canonical_object_framing_rejects_truncated_source_regions() -> Result<()> {
        let source = Archive {
            objects: vec![ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![3, 4, 5],
                }],
            )?],
        }
        .to_bytes()?;
        let archive = Archive::parse(&source)?;
        let data_offset = usize::try_from(archive.objects[0].data_offset)
            .map_err(|_| Error::invalid_archive(0, "test data offset exceeds usize"))?;

        let truncated_prefix = [source[0] | 0x80];
        assert!(
            archive
                .validate_canonical_object_framing(&truncated_prefix)
                .is_err()
        );
        assert!(
            archive
                .validate_canonical_object_framing(&source[..data_offset - 1])
                .is_err()
        );
        assert!(
            archive
                .validate_canonical_object_framing(&source[..source.len() - 1])
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn canonical_object_framing_rejects_inexact_offsets_and_extents() -> Result<()> {
        let source = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 2,
                        data: vec![3],
                    }],
                )?,
                ArchiveObject::new(
                    4,
                    vec![RawMessage {
                        type_: 5,
                        data: vec![6, 7],
                    }],
                )?,
            ],
        }
        .to_bytes()?;
        let archive = Archive::parse(&source)?;

        let mut wrong_header_offset = archive.clone();
        wrong_header_offset.objects[1].header_offset += 1;
        assert!(
            wrong_header_offset
                .validate_canonical_object_framing(&source)
                .is_err()
        );

        let mut wrong_header_length = archive.clone();
        wrong_header_length.objects[0].header_length += 1;
        assert!(
            wrong_header_length
                .validate_canonical_object_framing(&source)
                .is_err()
        );

        let mut wrong_data_offset = archive.clone();
        wrong_data_offset.objects[0].data_offset += 1;
        assert!(
            wrong_data_offset
                .validate_canonical_object_framing(&source)
                .is_err()
        );

        let mut wrong_data_length = archive.clone();
        wrong_data_length.objects[0].data_length += 1;
        assert!(
            wrong_data_length
                .validate_canonical_object_framing(&source)
                .is_err()
        );

        let mut trailing = source.clone();
        trailing.push(0);
        assert!(
            archive
                .validate_canonical_object_framing(&trailing)
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn canonical_message_length_width_change_and_inverse_restore_exact_bytes() -> Result<()> {
        let original_message = RawMessage {
            type_: 2,
            data: vec![0x11; 127],
        };
        let source = Archive {
            objects: vec![ArchiveObject::new(1, vec![original_message.clone()])?],
        }
        .to_bytes()?;
        let mut edited = Archive::parse(&source)?;
        let replacement = RawMessage {
            type_: 2,
            data: vec![0x22; 128],
        };

        let old = edited.objects[0].replace_message_preserving_header(0, replacement.clone())?;
        assert_eq!(old, original_message);
        let expected = Archive {
            objects: vec![ArchiveObject::new(1, vec![replacement.clone()])?],
        }
        .to_bytes()?;
        assert_eq!(edited.to_bytes()?, expected);
        assert!(edited.objects[0].original_header.is_none());
        assert!(edited.objects[0].original_canonical_header.is_none());

        let replaced = edited.objects[0].replace_message_preserving_header(0, old)?;
        assert_eq!(replaced, replacement);
        assert_eq!(edited.to_bytes()?, source);
        assert!(edited.objects[0].original_header.is_none());
        assert!(edited.objects[0].original_canonical_header.is_none());
        Ok(())
    }

    #[test]
    fn canonical_last_reference_pruning_and_inverse_restore_exact_bytes() -> Result<()> {
        let original_message = RawMessage {
            type_: 2,
            data: vec![3, 4],
        };
        let mut original_object = ArchiveObject::new(1, vec![original_message.clone()])?;
        let Some(original_info) = original_object.archive_info.message_infos.get_mut(0) else {
            return Err(Error::invalid_archive(
                0,
                "test object is missing message metadata",
            ));
        };
        original_info.object_references.push(77);
        let source_archive = Archive {
            objects: vec![original_object],
        };
        let source = source_archive.to_bytes()?;
        let mut edited = Archive::parse(&source)?;

        let replacement = RawMessage {
            type_: 9,
            data: vec![5, 6, 7],
        };
        let old = edited.objects[0].replace_message_pruning_object_references_preserving_header(
            0,
            replacement.clone(),
            &[77],
        )?;
        assert_eq!(old, original_message);
        assert!(edited.objects[0].original_header.is_none());
        assert!(edited.objects[0].original_canonical_header.is_none());

        let expected = Archive {
            objects: vec![ArchiveObject::new(1, vec![replacement.clone()])?],
        }
        .to_bytes()?;
        assert_eq!(edited.to_bytes()?, expected);

        let Some(edited_info) = edited.objects[0].archive_info.message_infos.get_mut(0) else {
            return Err(Error::invalid_archive(
                0,
                "edited test object is missing message metadata",
            ));
        };
        edited_info.object_references.push(77);
        let replaced = edited.objects[0].replace_message_preserving_header(0, old)?;
        assert_eq!(replaced, replacement);
        assert_eq!(edited.to_bytes()?, source);
        Ok(())
    }

    #[test]
    fn reference_pruning_preserves_all_untouched_header_encodings() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source)?;
        let old = archive.objects[0].replace_message_pruning_object_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0x5a; 130],
            },
            &[10, 12, 14, 10],
        )?;
        assert_eq!(old.type_, 7);
        assert_eq!(old.data, [0xde, 0xad, 0xbe]);

        let encoded = archive.to_bytes()?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.expected_header);
        assert_eq!(payload, vec![0x5a; 130]);

        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.type_, 9);
        assert_eq!(info.length, 130);
        // The duplicate retained identifier and its two distinct raw
        // encodings survive in source order.
        assert_eq!(info.object_references, [13, 13]);
        assert_eq!(info.field_infos[0].object_references, [99, 100]);
        assert_eq!(info.field_infos[1].object_references, [77, 78]);
        assert_eq!(
            archive.objects[0].original_header.as_deref(),
            Some(fixture.expected_header.as_slice())
        );

        let reparsed = Archive::parse(&encoded)?;
        assert_eq!(
            reparsed.objects[0].archive_info,
            archive.objects[0].archive_info
        );
        assert_eq!(reparsed.to_bytes()?, encoded);
        Ok(())
    }

    #[test]
    fn empty_reference_removal_delegates_to_preserving_replacement() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let parsed = Archive::parse(&fixture.source)?;
        let mut delegated = parsed.objects[0].clone();
        let mut existing = parsed.objects[0].clone();
        let replacement = RawMessage {
            type_: 9,
            data: vec![0x44; 130],
        };

        let delegated_old = delegated.replace_message_pruning_object_references_preserving_header(
            0,
            replacement.clone(),
            &[],
        )?;
        let existing_old = existing.replace_message_preserving_header(0, replacement)?;
        assert_eq!(delegated_old, existing_old);
        assert_eq!(delegated, existing);
        Ok(())
    }

    #[test]
    fn reference_pruning_failures_are_atomic() -> Result<()> {
        let fixture = reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source)?;
        let before_limit = archive.objects[0].clone();
        let limits = Limits::default()
            .with_header_bytes(fixture.header.len())?
            .with_object_bytes(fixture.source.len())?;
        let limit_error = archive.objects[0]
            .replace_message_pruning_object_references_preserving_header_with_limits(
                0,
                RawMessage {
                    type_: 9,
                    data: vec![0x5a; 130],
                },
                &[10],
                limits,
            )
            .err();
        assert!(matches!(
            limit_error,
            Some(Error::Limit {
                kind: LimitKind::ObjectBytes,
                ..
            })
        ));
        assert_eq!(archive.objects[0], before_limit);

        for malformed_message in [
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x29, 1, 2, 3, 4, 5, 6, 7, 8,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x2a, 0x01, 0x80,
            ],
        ] {
            let mut object = ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![0xcc],
                }],
            )?;
            let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
            let mut raw_header = vec![0x08, 0x01];
            push_length_delimited(&mut raw_header, &[0x12], 1, &malformed_message)?;
            object.original_header = Some(raw_header.into_boxed_slice());
            object.original_canonical_header = Some(canonical.into_boxed_slice());
            let before = object.clone();

            let malformed_error = object
                .replace_message_pruning_object_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 2,
                        data: vec![0xdd],
                    },
                    &[1],
                )
                .err();
            assert!(malformed_error.is_some());
            assert_eq!(object, before);
        }
        Ok(())
    }

    #[test]
    fn selected_object_and_data_pruning_preserves_raw_occurrences() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source)?;
        let old = archive.objects[0].replace_message_pruning_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0x5a; 130],
            },
            &[10, 12, 10],
            DataReferencePruning::Selected(&[20, 22, 20]),
        )?;
        assert_eq!(old.type_, 7);
        assert_eq!(old.data, [0xde, 0xad, 0xbe]);

        let encoded = archive.to_bytes()?;
        let (header, payload) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.selected_header);
        assert_eq!(payload, vec![0x5a; 130]);

        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.object_references, [13, 13, 13]);
        assert_eq!(info.data_references, [21, 21, 21, 21]);
        assert_eq!(info.field_infos[0].object_references, [100]);
        assert_eq!(info.field_infos[0].data_references, [21, 21]);
        assert_eq!(info.field_infos[1].object_references, [77]);
        assert_eq!(info.field_infos[1].data_references, [31, 32, 33]);
        assert_eq!(
            archive.objects[0].original_header.as_deref(),
            Some(fixture.selected_header.as_slice())
        );

        // A second parse proves that the raw result and its typed projection
        // remain paired for future header-preserving edits.
        let reparsed = Archive::parse(&encoded)?;
        assert_eq!(reparsed.to_bytes()?, encoded);
        Ok(())
    }

    #[test]
    fn all_data_pruning_clears_aggregate_and_nested_references_only() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source)?;
        archive.objects[0].replace_message_pruning_references_preserving_header(
            0,
            RawMessage {
                type_: 9,
                data: vec![0x5a; 130],
            },
            &[],
            DataReferencePruning::All,
        )?;

        let encoded = archive.to_bytes()?;
        let (header, _) = split_test_archive(&encoded)?;
        assert_eq!(header, fixture.all_data_header);
        let info = &archive.objects[0].archive_info.message_infos[0];
        assert_eq!(info.object_references, [10, 10, 13, 13, 12, 13, 10]);
        assert!(info.data_references.is_empty());
        assert_eq!(info.field_infos[0].object_references, [10, 12, 100, 10]);
        assert!(info.field_infos[0].data_references.is_empty());
        assert_eq!(info.field_infos[1].object_references, [77]);
        assert!(info.field_infos[1].data_references.is_empty());
        Ok(())
    }

    #[test]
    fn data_pruning_zero_and_noop_policies_preserve_existing_behavior() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let parsed = Archive::parse(&fixture.source)?;
        let replacement = RawMessage {
            type_: 9,
            data: vec![0x44; 130],
        };
        let mut delegated = parsed.objects[0].clone();
        let mut preserving = parsed.objects[0].clone();
        let delegated_old = delegated.replace_message_pruning_references_preserving_header(
            0,
            replacement.clone(),
            &[],
            DataReferencePruning::Selected(&[]),
        )?;
        let preserving_old = preserving.replace_message_preserving_header(0, replacement)?;
        assert_eq!(delegated_old, preserving_old);
        assert_eq!(delegated, preserving);

        let mut unmatched = parsed.objects[0].clone();
        let same = unmatched.messages[0].clone();
        unmatched.replace_message_pruning_references_preserving_header(
            0,
            same,
            &[u64::MAX],
            DataReferencePruning::Selected(&[u64::MAX]),
        )?;
        assert_eq!(unmatched, parsed.objects[0]);
        assert_eq!(
            Archive {
                objects: vec![unmatched],
            }
            .to_bytes()?,
            fixture.source
        );

        let mut empty = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 2,
                data: vec![3],
            }],
        )?;
        let empty_before = Archive {
            objects: vec![empty.clone()],
        }
        .to_bytes()?;
        empty.replace_message_pruning_references_preserving_header(
            0,
            RawMessage {
                type_: 2,
                data: vec![3],
            },
            &[],
            DataReferencePruning::All,
        )?;
        assert_eq!(
            Archive {
                objects: vec![empty],
            }
            .to_bytes()?,
            empty_before
        );
        Ok(())
    }

    #[test]
    fn canonical_data_pruning_can_restore_exact_bytes_after_metadata_restore() -> Result<()> {
        let original_message = RawMessage {
            type_: 2,
            data: vec![3, 4],
        };
        let mut object = ArchiveObject::new(1, vec![original_message.clone()])?;
        let info = &mut object.archive_info.message_infos[0];
        info.data_references.extend([7, 8]);
        let mut field = FieldInfo::new(vec![1]);
        field.data_references.extend([9, 10]);
        info.field_infos.push(field);
        let source = Archive {
            objects: vec![object],
        }
        .to_bytes()?;
        let mut edited = Archive::parse(&source)?;
        let replacement = RawMessage {
            type_: 9,
            data: vec![5, 6, 7],
        };
        let old = edited.objects[0].replace_message_pruning_references_preserving_header(
            0,
            replacement.clone(),
            &[],
            DataReferencePruning::Selected(&[8, 10]),
        )?;
        assert_eq!(
            edited.objects[0].archive_info.message_infos[0].data_references,
            [7]
        );
        assert_eq!(
            edited.objects[0].archive_info.message_infos[0].field_infos[0].data_references,
            [9]
        );

        let edited_info = &mut edited.objects[0].archive_info.message_infos[0];
        edited_info.data_references.insert(1, 8);
        edited_info.field_infos[0].data_references.insert(1, 10);
        let replaced = edited.objects[0].replace_message_preserving_header(0, old)?;
        assert_eq!(replaced, replacement);
        assert_eq!(edited.to_bytes()?, source);
        Ok(())
    }

    #[test]
    fn combined_reference_pruning_budgets_selectors_and_wire_work_atomically() -> Result<()> {
        let fixture = combined_reference_pruning_fixture()?;
        let mut archive = Archive::parse(&fixture.source)?;
        let before = archive.objects[0].clone();
        let selector_limits = Limits::default().with_metadata_items(2)?;
        let selector_error = archive.objects[0]
            .replace_message_pruning_references_preserving_header_with_limits(
                0,
                RawMessage {
                    type_: 7,
                    data: vec![0xde, 0xad, 0xbe],
                },
                &[10, 12],
                DataReferencePruning::Selected(&[20]),
                selector_limits,
            )
            .err();
        assert!(matches!(
            selector_error,
            Some(Error::Limit {
                kind: LimitKind::MetadataItems,
                observed: 3,
                maximum: 2,
            })
        ));
        assert_eq!(archive.objects[0], before);

        let wire_limits = Limits::default().with_header_fields(1)?;
        let wire_error = archive.objects[0]
            .replace_message_pruning_references_preserving_header_with_limits(
                0,
                RawMessage {
                    type_: 7,
                    data: vec![0xde, 0xad, 0xbe],
                },
                &[],
                DataReferencePruning::All,
                wire_limits,
            )
            .err();
        assert!(matches!(
            wire_error,
            Some(Error::Limit {
                kind: LimitKind::HeaderFields,
                ..
            })
        ));
        assert_eq!(archive.objects[0], before);
        Ok(())
    }

    #[test]
    fn ambiguous_or_malformed_data_reference_fields_are_rejected_atomically() -> Result<()> {
        for malformed_message in [
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x31, 1, 2, 3, 4, 5, 6, 7, 8,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x32, 0x01, 0x80,
            ],
            vec![
                0x08, 0x02, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x01, 0x22, 0x0a, 0x0a, 0x00, 0x29,
                1, 2, 3, 4, 5, 6, 7, 8,
            ],
        ] {
            let mut object = ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 2,
                    data: vec![0xcc],
                }],
            )?;
            let canonical = encode_archive_info(&object.archive_info, Limits::default())?;
            let mut raw_header = vec![0x08, 0x01];
            push_length_delimited(&mut raw_header, &[0x12], 1, &malformed_message)?;
            object.original_header = Some(raw_header.into_boxed_slice());
            object.original_canonical_header = Some(canonical.into_boxed_slice());
            let before = object.clone();

            let error = object
                .replace_message_pruning_references_preserving_header(
                    0,
                    RawMessage {
                        type_: 2,
                        data: vec![0xdd],
                    },
                    &[],
                    DataReferencePruning::All,
                )
                .err();
            assert!(error.is_some());
            assert_eq!(object, before);
        }
        Ok(())
    }

    fn combined_reference_pruning_fixture() -> Result<CombinedReferencePruningFixture> {
        let source_message = combined_reference_message(7, 3, false, 0)?;
        let selected_message = combined_reference_message(9, 130, true, 1)?;
        let all_data_message = combined_reference_message(9, 130, false, 2)?;
        let source_header = combined_reference_header(&source_message)?;
        let selected_header = combined_reference_header(&selected_message)?;
        let all_data_header = combined_reference_header(&all_data_message)?;

        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(source_header.len() as u64, &mut prefix));
        source.extend_from_slice(&source_header);
        source.extend_from_slice(&[0xde, 0xad, 0xbe]);
        Ok(CombinedReferencePruningFixture {
            source,
            selected_header,
            all_data_header,
        })
    }

    fn combined_reference_header(message: &[u8]) -> Result<Vec<u8>> {
        let mut header = vec![
            0x08, 0x01, 0x88, 0x00, 0xaa, 0x00, 0xd1, 0x06, 1, 2, 3, 4, 5, 6, 7, 8,
        ];
        push_length_delimited(&mut header, &[0x92, 0x00], 2, message)?;
        header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);
        Ok(header)
    }

    // data_mode: 0 retains all data references, 1 prunes 20 and 22, and 2
    // prunes all data references. Conditional construction keeps the exact
    // expected raw bytes visible independently of the production rewriter.
    fn combined_reference_message(
        type_: u64,
        length: u64,
        prune_objects: bool,
        data_mode: u8,
    ) -> Result<Vec<u8>> {
        let first_field = combined_first_field(prune_objects, data_mode)?;
        let second_field = combined_second_field(data_mode)?;
        let mut message = Vec::new();
        message.extend_from_slice(&[0xcd, 0x0c, 9, 8, 7, 6]);
        message.extend_from_slice(&[0x08, 0x01, 0x88, 0x00]);
        push_test_varint(&mut message, type_, 2);
        push_length_delimited(&mut message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        message.extend_from_slice(&[0x18, 0x01, 0x98, 0x00]);
        push_test_varint(&mut message, length, 2);

        if prune_objects {
            message.extend_from_slice(&[0x28, 0x0d, 0xa8, 0x00, 0x8d, 0x00]);
            push_length_delimited(&mut message, &[0x2a], 2, &[0x0d])?;
        } else {
            message.extend_from_slice(&[
                0x28, 0x8a, 0x00, 0xa8, 0x00, 0x8a, 0x00, 0x28, 0x0d, 0xa8, 0x00, 0x8d, 0x00,
            ]);
            push_length_delimited(&mut message, &[0x2a], 2, &[0x0c, 0x0d, 0x0a])?;
        }
        push_length_delimited(&mut message, &[0x2a], 2, &[])?;

        match data_mode {
            0 => {
                message.extend_from_slice(&[
                    0x30, 0x94, 0x00, 0xb0, 0x00, 0x94, 0x00, 0x30, 0x15, 0xb0, 0x00, 0x95, 0x00,
                ]);
                push_length_delimited(&mut message, &[0x32], 2, &[0x16, 0x15, 0x14, 0x15])?;
            },
            1 => {
                message.extend_from_slice(&[0x30, 0x15, 0xb0, 0x00, 0x95, 0x00]);
                push_length_delimited(&mut message, &[0x32], 2, &[0x15, 0x15])?;
            },
            2 => push_length_delimited(&mut message, &[0x32], 2, &[])?,
            _ => return Err(Error::invalid_archive(0, "invalid test data pruning mode")),
        }
        push_length_delimited(&mut message, &[0x32], 2, &[])?;
        push_length_delimited(&mut message, &[0xa2, 0x00], 2, &first_field)?;
        push_length_delimited(&mut message, &[0x22], 2, &second_field)?;
        push_length_delimited(&mut message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;
        Ok(message)
    }

    fn combined_first_field(prune_objects: bool, data_mode: u8) -> Result<Vec<u8>> {
        let mut field = vec![0xa0, 0x01, 0x81, 0x00];
        push_length_delimited(&mut field, &[0x0a], 2, &[])?;
        if prune_objects {
            push_length_delimited(&mut field, &[0x22], 2, &[0xe4, 0x00])?;
        } else {
            field.extend_from_slice(&[0x20, 0x8a, 0x00]);
            push_length_delimited(&mut field, &[0x22], 2, &[0x0c, 0xe4, 0x00, 0x0a])?;
        }
        match data_mode {
            0 => {
                field.extend_from_slice(&[0x28, 0x94, 0x00, 0xa8, 0x00, 0x95, 0x00]);
                push_length_delimited(&mut field, &[0x2a], 2, &[0x14, 0x96, 0x00, 0x15, 0x14])?;
            },
            1 => {
                field.extend_from_slice(&[0xa8, 0x00, 0x95, 0x00]);
                push_length_delimited(&mut field, &[0x2a], 2, &[0x15])?;
            },
            2 => push_length_delimited(&mut field, &[0x2a], 2, &[])?,
            _ => return Err(Error::invalid_archive(0, "invalid test data pruning mode")),
        }
        push_length_delimited(&mut field, &[0x2a], 2, &[])?;
        push_length_delimited(&mut field, &[0xaa, 0x01], 2, &[0xde, 0xad])?;
        Ok(field)
    }

    fn combined_second_field(data_mode: u8) -> Result<Vec<u8>> {
        let mut field = Vec::new();
        push_length_delimited(&mut field, &[0x0a], 1, &[])?;
        field.extend_from_slice(&[0x20, 0x4d]);
        match data_mode {
            0 | 1 => {
                field.extend_from_slice(&[0x28, 0x1f]);
                push_length_delimited(&mut field, &[0xaa, 0x00], 2, &[0x20, 0x21])?;
            },
            2 => push_length_delimited(&mut field, &[0xaa, 0x00], 2, &[])?,
            _ => return Err(Error::invalid_archive(0, "invalid test data pruning mode")),
        }
        field.extend_from_slice(&[0xb5, 0x0c, 1, 2, 3, 4]);
        Ok(field)
    }

    fn push_test_varint(output: &mut Vec<u8>, value: u64, width: usize) {
        let mut encoded = [0u8; 10];
        output.extend_from_slice(encode_varint_with_width(value, width, &mut encoded));
    }

    fn reference_pruning_fixture() -> Result<ReferencePruningFixture> {
        let mut changed_field = Vec::new();
        changed_field.extend_from_slice(&[0xa0, 0x01, 0x81, 0x00]);
        changed_field.extend_from_slice(&[0x20, 0x8a, 0x00]);
        changed_field.extend_from_slice(&[0xa0, 0x00, 0xe3, 0x00]);
        push_length_delimited(&mut changed_field, &[0x22], 2, &[0x0c, 0xe4, 0x00, 0x0e])?;
        push_length_delimited(&mut changed_field, &[0x0a], 2, &[])?;
        push_length_delimited(&mut changed_field, &[0xaa, 0x01], 2, &[0xde, 0xad])?;

        let mut expected_changed_field = Vec::new();
        expected_changed_field.extend_from_slice(&[0xa0, 0x01, 0x81, 0x00]);
        expected_changed_field.extend_from_slice(&[0xa0, 0x00, 0xe3, 0x00]);
        push_length_delimited(&mut expected_changed_field, &[0x22], 2, &[0xe4, 0x00])?;
        push_length_delimited(&mut expected_changed_field, &[0x0a], 2, &[])?;
        push_length_delimited(&mut expected_changed_field, &[0xaa, 0x01], 2, &[0xde, 0xad])?;

        let mut untouched_field = Vec::new();
        push_length_delimited(&mut untouched_field, &[0x0a], 1, &[])?;
        push_length_delimited(&mut untouched_field, &[0xa2, 0x00], 2, &[0xcd, 0x00, 0x4e])?;
        untouched_field.extend_from_slice(&[0xb5, 0x0c, 1, 2, 3, 4]);

        let mut message = Vec::new();
        message.extend_from_slice(&[0xcd, 0x0c, 9, 8, 7, 6]);
        message.extend_from_slice(&[0x08, 0x01, 0x88, 0x00, 0x87, 0x00]);
        push_length_delimited(&mut message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        message.extend_from_slice(&[0x18, 0x01, 0x98, 0x00, 0x83, 0x00]);
        message.extend_from_slice(&[0x28, 0x8a, 0x00]);
        message.extend_from_slice(&[0xa8, 0x00, 0x8d, 0x00]);
        push_length_delimited(&mut message, &[0x2a], 2, &[0x0c, 0x8d, 0x00, 0x0e])?;
        push_length_delimited(&mut message, &[0x2a], 2, &[0x0a, 0x0c])?;
        push_length_delimited(&mut message, &[0xa2, 0x00], 2, &changed_field)?;
        push_length_delimited(&mut message, &[0x22], 2, &untouched_field)?;
        push_length_delimited(&mut message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;

        let mut expected_message = Vec::new();
        expected_message.extend_from_slice(&[0xcd, 0x0c, 9, 8, 7, 6]);
        expected_message.extend_from_slice(&[0x08, 0x01, 0x88, 0x00, 0x89, 0x00]);
        push_length_delimited(&mut expected_message, &[0x12], 2, &[0x01, 0x00, 0x05])?;
        expected_message.extend_from_slice(&[0x18, 0x01, 0x98, 0x00, 0x82, 0x01]);
        expected_message.extend_from_slice(&[0xa8, 0x00, 0x8d, 0x00]);
        push_length_delimited(&mut expected_message, &[0x2a], 2, &[0x8d, 0x00])?;
        push_length_delimited(&mut expected_message, &[0x2a], 2, &[])?;
        push_length_delimited(
            &mut expected_message,
            &[0xa2, 0x00],
            2,
            &expected_changed_field,
        )?;
        push_length_delimited(&mut expected_message, &[0x22], 2, &untouched_field)?;
        push_length_delimited(&mut expected_message, &[0xa2, 0x06], 2, &[0xfe, 0xed])?;

        let mut header = vec![
            0x08, 0x01, 0x88, 0x00, 0xaa, 0x00, 0xd1, 0x06, 1, 2, 3, 4, 5, 6, 7, 8,
        ];
        push_length_delimited(&mut header, &[0x92, 0x00], 2, &message)?;
        header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);

        let mut expected_header = vec![
            0x08, 0x01, 0x88, 0x00, 0xaa, 0x00, 0xd1, 0x06, 1, 2, 3, 4, 5, 6, 7, 8,
        ];
        push_length_delimited(&mut expected_header, &[0x92, 0x00], 2, &expected_message)?;
        expected_header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee, 0x18, 0x81, 0x00]);

        let mut source = Vec::new();
        let mut prefix = [0u8; 10];
        source.extend_from_slice(encode_varint(header.len() as u64, &mut prefix));
        source.extend_from_slice(&header);
        source.extend_from_slice(&[0xde, 0xad, 0xbe]);
        Ok(ReferencePruningFixture {
            source,
            header,
            expected_header,
        })
    }

    fn push_length_delimited(
        output: &mut Vec<u8>,
        key: &[u8],
        prefix_width: usize,
        payload: &[u8],
    ) -> Result<()> {
        let additional = key
            .len()
            .checked_add(prefix_width)
            .and_then(|length| length.checked_add(payload.len()))
            .ok_or_else(|| Error::invalid_archive(0, "test field length overflow"))?;
        output
            .try_reserve(additional)
            .map_err(|_| Error::allocation("test length-delimited field", additional))?;
        output.extend_from_slice(key);
        let mut prefix = [0u8; 10];
        output.extend_from_slice(encode_varint_with_width(
            payload.len() as u64,
            prefix_width,
            &mut prefix,
        ));
        output.extend_from_slice(payload);
        Ok(())
    }

    fn split_test_archive(data: &[u8]) -> Result<(&[u8], &[u8])> {
        let (header_length, prefix_length) = super::decode_varint(data)?;
        let header_length = usize::try_from(header_length)
            .map_err(|_| Error::invalid_archive(0, "test header length exceeds usize"))?;
        let header_end = prefix_length
            .checked_add(header_length)
            .ok_or_else(|| Error::invalid_archive(0, "test header range overflow"))?;
        let header = data
            .get(prefix_length..header_end)
            .ok_or_else(|| Error::invalid_archive(0, "test header range is invalid"))?;
        let payload = data
            .get(header_end..)
            .ok_or_else(|| Error::invalid_archive(0, "test payload range is invalid"))?;
        Ok((header, payload))
    }
}
