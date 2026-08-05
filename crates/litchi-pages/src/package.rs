//! Pages-specific decoding for the package root and body storage.
//!
//! This module owns only the Pages interpretation of the package payloads.
//! ZIP ingress, archive indexing, and object resolution remain in the IWA
//! implementation crate, while this crate depends only on the shared physical
//! archive and generated schema layers.

use std::num::NonZeroU64;

use litchi_iwa_core::{Archive, ArchiveObject};
use litchi_iwa_protos::{tp, tswp};
use prost::Message;
use thiserror::Error;

use crate::{Section, SectionType};

const ROOT_OBJECT_ID: u64 = 1;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPES: [u32; 2] = [2_001, 2_022];

/// Errors produced while decoding the Pages package root or body.
#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub enum Error {
    /// The expected object was not present in the archive.
    #[error("{kind} object {identifier} is missing")]
    MissingObject {
        /// Human-readable Pages payload kind.
        kind: &'static str,
        /// Expected object identifier.
        identifier: u64,
    },
    /// The object did not contain the expected message type.
    #[error("{kind} payload type {message_type} is missing")]
    MissingPayload {
        /// Human-readable Pages payload kind.
        kind: &'static str,
        /// Expected protobuf message type.
        message_type: u32,
    },
    /// More than one recognized payload was present.
    #[error("{kind} contains duplicate payload type {message_type}")]
    DuplicatePayload {
        /// Human-readable Pages payload kind.
        kind: &'static str,
        /// Ambiguous protobuf message type.
        message_type: u32,
    },
    /// A required reference was encoded as zero.
    #[error("{kind} reference is zero")]
    ZeroReference {
        /// Human-readable Pages reference kind.
        kind: &'static str,
    },
    /// The object has no usable archive identifier.
    #[error("Pages body object has no archive identifier")]
    MissingObjectIdentifier,
    /// The resolved object was not the object requested by the root.
    #[error("Pages body object identifier mismatch: expected {expected}, got {actual}")]
    ObjectIdentifierMismatch {
        /// Identifier from the root reference.
        expected: u64,
        /// Identifier on the resolved archive object.
        actual: u64,
    },
    /// The body object has no supported text-storage payload.
    #[error("Pages body object {identifier} has no text storage payload")]
    MissingBodyPayload {
        /// Referenced body-storage identifier.
        identifier: u64,
    },
    /// The decoded body would exceed the caller's text budget.
    #[error("Pages body text for object {identifier} exceeds {limit} bytes")]
    TextTooLarge {
        /// Referenced body-storage identifier.
        identifier: u64,
        /// Maximum joined UTF-8 byte length.
        limit: usize,
    },
    /// Root protobuf decoding failed.
    #[error("Pages root protobuf decoding failed: {0}")]
    RootDecode(#[source] prost::DecodeError),
    /// Body protobuf decoding failed.
    #[error("Pages body protobuf type {message_type} decoding failed: {source}")]
    BodyDecode {
        /// Native protobuf message type.
        message_type: u32,
        /// Decoder failure.
        #[source]
        source: prost::DecodeError,
    },
}

/// Result type for Pages package decoding.
pub type Result<T> = std::result::Result<T, Error>;

/// A decoded Pages package root.
///
/// The root keeps the exact wire payload for the recognized protobuf message.
/// This makes unknown protobuf fields available to a future writer instead of
/// silently replacing them with a re-encoded message containing only fields
/// known by this version of the schema.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Root {
    body_storage: Option<NonZeroU64>,
    encoded_payload: Box<[u8]>,
}

impl Root {
    /// Decode the Pages root from the `Index/Document.iwa` archive.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the root object or payload is missing,
    /// ambiguous, malformed, or contains a zero body reference.
    pub fn decode(archive: &Archive) -> Result<Self> {
        let object = archive.object(ROOT_OBJECT_ID).ok_or(Error::MissingObject {
            kind: "Pages root",
            identifier: ROOT_OBJECT_ID,
        })?;

        let mut root_payload = None;
        for message in &object.messages {
            if message.type_ == ROOT_MESSAGE_TYPE
                && root_payload.replace(message.data.as_slice()).is_some()
            {
                return Err(Error::DuplicatePayload {
                    kind: "Pages root",
                    message_type: ROOT_MESSAGE_TYPE,
                });
            }
        }

        let payload = root_payload.ok_or(Error::MissingPayload {
            kind: "Pages root",
            message_type: ROOT_MESSAGE_TYPE,
        })?;
        let document = tp::DocumentArchive::decode(payload).map_err(Error::RootDecode)?;
        let body_storage = document
            .body_storage
            .map(|reference| {
                NonZeroU64::new(reference.identifier).ok_or(Error::ZeroReference {
                    kind: "Pages body storage",
                })
            })
            .transpose()?;

        Ok(Self {
            body_storage,
            encoded_payload: payload.to_vec().into_boxed_slice(),
        })
    }

    /// Return the checked body-storage identifier, if the root has one.
    #[must_use]
    pub const fn body_storage(&self) -> Option<NonZeroU64> {
        self.body_storage
    }

    /// Return the exact encoded root payload, including unknown fields.
    #[must_use]
    pub fn encoded_payload(&self) -> &[u8] {
        &self.encoded_payload
    }
}

/// Decode one Pages body-storage object into a semantic body section.
///
/// `max_text_bytes` bounds the total UTF-8 size after the native repeated
/// text fragments are concatenated in source order. The identifier is checked both by
/// the caller's `NonZeroU64` type and against the archive object's metadata.
///
/// # Errors
///
/// Returns a typed error when the object identifier is invalid, the body
/// payload is missing or ambiguous, protobuf decoding fails, or the text
/// exceeds `max_text_bytes`.
pub fn decode_body_section(
    object: &ArchiveObject,
    identifier: NonZeroU64,
    max_text_bytes: usize,
) -> Result<Section> {
    let actual_identifier = object
        .archive_info
        .identifier
        .ok_or(Error::MissingObjectIdentifier)?;
    if actual_identifier != identifier.get() {
        return Err(Error::ObjectIdentifierMismatch {
            expected: identifier.get(),
            actual: actual_identifier,
        });
    }

    decode_body_section_from_messages(&object.messages, identifier, max_text_bytes)
}

/// Decode a Pages body-storage payload from an indexed message view.
///
/// This adapter keeps the IWA reader's zero-copy object-index path intact;
/// callers that already validated the object's identifier can pass its raw
/// message slice without cloning archive metadata or payloads.
///
/// # Errors
///
/// Returns a typed error when the body payload is missing or ambiguous,
/// protobuf decoding fails, or the text exceeds `max_text_bytes`.
pub fn decode_body_section_from_messages(
    messages: &[litchi_iwa_core::RawMessage],
    identifier: NonZeroU64,
    max_text_bytes: usize,
) -> Result<Section> {
    let mut body_payload = None;
    for message in messages {
        if BODY_MESSAGE_TYPES.contains(&message.type_)
            && body_payload
                .replace((message.type_, message.data.as_slice()))
                .is_some()
        {
            return Err(Error::DuplicatePayload {
                kind: "Pages body storage",
                message_type: message.type_,
            });
        }
    }

    let (message_type, payload) = body_payload.ok_or(Error::MissingBodyPayload {
        identifier: identifier.get(),
    })?;
    let storage = tswp::StorageArchive::decode(payload).map_err(|source| Error::BodyDecode {
        message_type,
        source,
    })?;

    let text_len = storage.text.iter().try_fold(0usize, |total_length, line| {
        total_length
            .checked_add(line.len())
            .ok_or(Error::TextTooLarge {
                identifier: identifier.get(),
                limit: max_text_bytes,
            })
    })?;
    if text_len > max_text_bytes {
        return Err(Error::TextTooLarge {
            identifier: identifier.get(),
            limit: max_text_bytes,
        });
    }

    let mut text = String::with_capacity(text_len);
    for line in &storage.text {
        text.push_str(line);
    }

    let mut text_storage = litchi_iwa_text::TextStorage::from_text(text);
    text_storage.identifier = Some(identifier.get());
    let mut section = Section::new(0, SectionType::Body);
    section.text_storages.push(text_storage);
    Ok(section)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_core::RawMessage;
    use litchi_iwa_protos::tsp::Reference;

    fn test_identifier() -> NonZeroU64 {
        NonZeroU64::new(42).unwrap_or_else(|| unreachable!("literal is non-zero"))
    }

    fn root_archive(body_storage: u64, extra: &[u8]) -> Archive {
        let mut payload = tp::DocumentArchive {
            body_storage: Some(Reference {
                identifier: body_storage,
                ..Default::default()
            }),
            ..Default::default()
        }
        .encode_to_vec();
        payload.extend_from_slice(extra);
        Archive {
            objects: vec![
                ArchiveObject::new(
                    ROOT_OBJECT_ID,
                    vec![RawMessage {
                        type_: ROOT_MESSAGE_TYPE,
                        data: payload,
                    }],
                )
                .unwrap_or_else(|error| panic!("test archive object: {error}")),
            ],
        }
    }

    fn body_object(identifier: u64, text: &[&str]) -> ArchiveObject {
        ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: BODY_MESSAGE_TYPES[0],
                data: tswp::StorageArchive {
                    text: text.iter().map(|line| (*line).to_owned()).collect(),
                    ..Default::default()
                }
                .encode_to_vec(),
            }],
        )
        .unwrap_or_else(|error| panic!("test body object: {error}"))
    }

    #[test]
    fn root_checks_and_retains_wire_payload() {
        let archive = root_archive(42, &[0x80, 0x3e, 0x01]);
        let root =
            Root::decode(&archive).unwrap_or_else(|error| panic!("root should decode: {error}"));
        assert_eq!(root.body_storage().map(NonZeroU64::get), Some(42));
        assert_eq!(root.encoded_payload().last(), Some(&1));
    }

    #[test]
    fn root_rejects_zero_reference_and_duplicates() {
        let zero = root_archive(0, &[]);
        assert!(matches!(
            Root::decode(&zero),
            Err(Error::ZeroReference { .. })
        ));

        let mut duplicate = root_archive(42, &[]);
        duplicate.objects[0].messages.push(RawMessage {
            type_: ROOT_MESSAGE_TYPE,
            data: tp::DocumentArchive::default().encode_to_vec(),
        });
        assert!(matches!(
            Root::decode(&duplicate),
            Err(Error::DuplicatePayload { .. })
        ));
    }

    #[test]
    fn body_decoding_is_bounded_and_typed() {
        let object = body_object(42, &["Pages", "body"]);
        let section = decode_body_section(&object, test_identifier(), 9)
            .unwrap_or_else(|error| panic!("body should decode: {error}"));
        assert_eq!(section.plain_text(), "Pagesbody");
        assert!(matches!(
            decode_body_section(&object, test_identifier(), 8),
            Err(Error::TextTooLarge { .. })
        ));
    }
}
