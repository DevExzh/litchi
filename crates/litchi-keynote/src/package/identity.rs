//! Native document identity: the embedded blank template, the identity
//! reader, and the exact-source identity rewrite used by fresh authoring.
//!
//! Keynote records the three identity UUIDs in exactly four places, all of
//! which this module owns:
//!
//! - `Metadata/DocumentIdentifier` — the document UUID as raw ASCII;
//! - `Metadata/Properties.plist` — `documentUUID`, `stableDocumentUUID` and
//!   `shareUUID` (document), `versionUUID` and `revision` (`N::<version>`)
//!   and `privateUUID`;
//! - `Index/Metadata.iwa` — `TSP.PackageMetadata.revision.identifier`
//!   (version), in the package-metadata object.
//!
//! The rewrite replaces those members through the exact-source reassembly and
//! reopens the candidate; every other member is copied unchanged.

use std::sync::Arc;

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use plist::Value;

use super::{Package, PhysicalSource, ReadError, ReadResult};
use crate::identity::{Identity, Role};

/// The blank presentation behind [`Package::new`]; see `assets/README.md`.
const BLANK_TEMPLATE: &[u8] = include_bytes!("../../assets/blank.key");

const PROPERTIES_MEMBER: &str = "Metadata/Properties.plist";
const DOCUMENT_IDENTIFIER_MEMBER: &str = "Metadata/DocumentIdentifier";
const PACKAGE_METADATA_MEMBER: &str = "Index/Metadata.iwa";
const PACKAGE_METADATA_OBJECT: u64 = 2;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const PACKAGE_METADATA_REVISION_FIELD: u64 = 2;
const DOCUMENT_REVISION_IDENTIFIER_FIELD: u64 = 2;
const DOCUMENT_KEYS: [&str; 3] = ["documentUUID", "stableDocumentUUID", "shareUUID"];
const VERSION_KEY: &str = "versionUUID";
const PRIVATE_KEY: &str = "privateUUID";
const REVISION_KEY: &str = "revision";

impl Package {
    /// Return the minimal valid presentation with the fixed [`Identity::BLANK`].
    ///
    /// The package holds one title-and-subtitle slide ("Presentation Title" /
    /// "Presentation Subtitle", no speaker notes) at 1920×1080 points, English
    /// (`en`, `en_US`). It is byte-for-byte identical on every call: no clock,
    /// randomness or locale is consulted. Every package created this way
    /// shares one native document identity, so a caller that publishes more
    /// than one independent document uses [`crate::Builder`] with
    /// [`Identity::generate`] instead.
    ///
    /// # Costs
    ///
    /// Parses an embedded package of about 10 KiB under default limits.
    #[must_use]
    pub fn new() -> Self {
        Self::from_bytes(BLANK_TEMPLATE).expect("the embedded blank Keynote template is valid")
    }

    /// Read the native document identity recorded in `Metadata/Properties.plist`.
    ///
    /// Returns `Ok(None)` when the package has no properties plist or the plist
    /// lacks one of the document, version or private UUIDs.
    ///
    /// # Errors
    ///
    /// Returns an error when the plist is not decodable, a recorded UUID is
    /// malformed, or the recorded UUIDs do not form a valid [`Identity`].
    pub fn identity(&self) -> ReadResult<Option<Identity>> {
        let Some(properties) = properties_dictionary(self)? else {
            return Ok(None);
        };
        let uuid = |key: &str| -> ReadResult<Option<[u8; 16]>> {
            properties
                .get(key)
                .map(|value| {
                    value
                        .as_string()
                        .and_then(Identity::parse_uuid)
                        .ok_or_else(|| {
                            ReadError::InvalidFormat(format!(
                                "Keynote properties carry a malformed {key}"
                            ))
                        })
                })
                .transpose()
        };
        let (Some(document), Some(version), Some(private)) = (
            uuid(DOCUMENT_KEYS[0])?,
            uuid(VERSION_KEY)?,
            uuid(PRIVATE_KEY)?,
        ) else {
            return Ok(None);
        };
        Identity::from_bytes(document, version, private)
            .map(Some)
            .map_err(|error| ReadError::InvalidFormat(error.to_string()))
    }

    /// Return a reopened copy of this package carrying `identity`.
    ///
    /// The current identity is reused without reassembly when it already
    /// equals `identity`, so the result is a pure function of the source
    /// bytes and the requested identity.
    pub(crate) fn with_identity(&self, identity: &Identity) -> ReadResult<Self> {
        if self.identity()?.as_ref() == Some(identity) {
            return Ok(self.snapshot());
        }
        let PhysicalSource::Package(source) = &self.state.source else {
            return Err(ReadError::InvalidFormat(
                "semantic-only Keynote sources cannot carry a new identity".to_owned(),
            ));
        };
        if !source.source_is_exact() {
            return Err(ReadError::InvalidFormat(
                "inexact Keynote sources cannot carry a new identity".to_owned(),
            ));
        }
        let properties = rewrite_properties(self, identity)?;
        let document = identity.text(Role::DocumentUuid);
        let metadata = rewrite_package_metadata(self, identity)?;
        let output = source.package().reassemble_with_changes_to_bytes(
            &[],
            &[
                EntryEdit::new(PROPERTIES_MEMBER, &properties),
                EntryEdit::new(DOCUMENT_IDENTIFIER_MEMBER, document.as_bytes()),
                EntryEdit::new(PACKAGE_METADATA_MEMBER, &metadata),
            ],
            &[],
            self.limits(),
        )?;
        let candidate = Self::from_source_with_options(Arc::from(output), self.state.options)?;
        if candidate.identity()?.as_ref() != Some(identity) {
            return Err(ReadError::InvalidFormat(
                "the rewritten Keynote identity failed readback".to_owned(),
            ));
        }
        Ok(candidate)
    }
}

impl Default for Package {
    /// Equivalent to [`Package::new`].
    fn default() -> Self {
        Self::new()
    }
}

fn member<'a>(package: &'a Package, name: &str) -> ReadResult<Option<&'a [u8]>> {
    let PhysicalSource::Package(source) = &package.state.source else {
        return Ok(None);
    };
    source
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| {
            if entry.is_opaque() {
                return Err(ReadError::InvalidFormat(format!(
                    "Keynote member {name} uses unsupported compression"
                )));
            }
            Ok(entry.data())
        })
        .transpose()
}

fn properties_dictionary(package: &Package) -> ReadResult<Option<plist::Dictionary>> {
    let Some(data) = member(package, PROPERTIES_MEMBER)? else {
        return Ok(None);
    };
    if data.len() > litchi_iwa_detect::MAX_PROPERTIES_BYTES {
        return Err(ReadError::InvalidFormat(
            "Keynote properties exceed their size limit".to_owned(),
        ));
    }
    match Value::from_reader(std::io::Cursor::new(data))? {
        Value::Dictionary(dictionary) => Ok(Some(dictionary)),
        _ => Err(ReadError::InvalidFormat(
            "Keynote properties are not a dictionary".to_owned(),
        )),
    }
}

fn rewrite_properties(package: &Package, identity: &Identity) -> ReadResult<Vec<u8>> {
    let mut properties = properties_dictionary(package)?.ok_or_else(|| {
        ReadError::InvalidFormat("Keynote package has no properties plist".to_owned())
    })?;
    let document = identity.text(Role::DocumentUuid);
    let version = identity.text(Role::VersionUuid);
    for key in DOCUMENT_KEYS {
        if key == DOCUMENT_KEYS[0] || properties.contains_key(key) {
            properties.insert(key.to_owned(), Value::String(document.clone()));
        }
    }
    properties.insert(VERSION_KEY.to_owned(), Value::String(version.clone()));
    properties.insert(
        PRIVATE_KEY.to_owned(),
        Value::String(identity.text(Role::PrivateUuid)),
    );
    let sequence = properties
        .get(REVISION_KEY)
        .and_then(Value::as_string)
        .and_then(|revision| revision.split_once("::"))
        .map_or("0", |(sequence, _version)| sequence)
        .to_owned();
    properties.insert(
        REVISION_KEY.to_owned(),
        Value::String(format!("{sequence}::{version}")),
    );
    let mut encoded = Vec::new();
    Value::Dictionary(properties).to_writer_binary(&mut encoded)?;
    Ok(encoded)
}

fn rewrite_package_metadata(package: &Package, identity: &Identity) -> ReadResult<Vec<u8>> {
    let data = member(package, PACKAGE_METADATA_MEMBER)?.ok_or_else(|| {
        ReadError::InvalidFormat("Keynote package has no package metadata".to_owned())
    })?;
    let limits = package.limits();
    let archive_limits = limits.effective_archive_limits()?;
    let stream = SnappyStream::decompress(data, limits.snappy_limits()?).map_err(core_error)?;
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits).map_err(core_error)?;
    drop(stream);
    let object = archive
        .object_mut(PACKAGE_METADATA_OBJECT)
        .ok_or_else(|| invalid_metadata("object"))?;
    let mut selected = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE);
    let (index, message) = selected.next().ok_or_else(|| invalid_metadata("message"))?;
    if selected.next().is_some() {
        return Err(invalid_metadata("message"));
    }
    let version = identity.text(Role::VersionUuid);
    let rewritten =
        rewrite_length_delimited(&message.data, PACKAGE_METADATA_REVISION_FIELD, |revision| {
            rewrite_length_delimited(revision, DOCUMENT_REVISION_IDENTIFIER_FIELD, |_old| {
                Ok(version.as_bytes().to_vec())
            })
        })?;
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(core_error)?;
    let bytes = archive.to_bytes(archive_limits).map_err(core_error)?;
    SnappyStream::compress(&bytes, limits.snappy_limits()?).map_err(core_error)
}

/// Rewrite the single occurrence of length-delimited `field` in a protobuf
/// payload, copying every other field verbatim.
fn rewrite_length_delimited(
    payload: &[u8],
    field: u64,
    replace: impl FnOnce(&[u8]) -> ReadResult<Vec<u8>>,
) -> ReadResult<Vec<u8>> {
    let mut output = Vec::with_capacity(payload.len());
    let mut replace = Some(replace);
    let mut offset = 0_usize;
    while offset < payload.len() {
        let start = offset;
        let (key, key_length) = decode_varint_from_bytes(&payload[offset..])
            .map_err(|_error| invalid_metadata("key"))?;
        offset += key_length;
        let rest = &payload[offset..];
        match key & 7 {
            0 => {
                let (_value, length) =
                    decode_varint_from_bytes(rest).map_err(|_error| invalid_metadata("varint"))?;
                offset += length;
            },
            1 => {
                offset = offset
                    .checked_add(8)
                    .ok_or_else(|| invalid_metadata("fixed"))?
            },
            5 => {
                offset = offset
                    .checked_add(4)
                    .ok_or_else(|| invalid_metadata("fixed"))?
            },
            2 => {
                let (length, prefix) =
                    decode_varint_from_bytes(rest).map_err(|_error| invalid_metadata("length"))?;
                let length =
                    usize::try_from(length).map_err(|_error| invalid_metadata("length"))?;
                let body_start = offset + prefix;
                let body_end = body_start
                    .checked_add(length)
                    .filter(|end| *end <= payload.len())
                    .ok_or_else(|| invalid_metadata("length"))?;
                offset = body_end;
                if key >> 3 == field {
                    let replace = replace.take().ok_or_else(|| invalid_metadata("repeat"))?;
                    let body = replace(&payload[body_start..body_end])?;
                    encode_varint_into(&mut output, key);
                    encode_varint_into(&mut output, body.len() as u64);
                    output.extend_from_slice(&body);
                    continue;
                }
            },
            _ => return Err(invalid_metadata("wire type")),
        }
        if offset > payload.len() {
            return Err(invalid_metadata("truncated"));
        }
        output.extend_from_slice(&payload[start..offset]);
    }
    if replace.is_some() {
        return Err(invalid_metadata("revision"));
    }
    Ok(output)
}

fn invalid_metadata(part: &str) -> ReadError {
    ReadError::InvalidFormat(format!(
        "Keynote package metadata has an unsupported {part}"
    ))
}

fn core_error(error: litchi_iwa_core::Error) -> ReadError {
    ReadError::InvalidFormat(format!("Keynote package metadata rewrite failed: {error}"))
}

#[cfg(test)]
mod tests {
    use super::{BLANK_TEMPLATE, Package};
    use crate::identity::Identity;

    #[test]
    fn embedded_template_carries_the_blank_identity() {
        let package = Package::new();
        assert_eq!(package.identity().ok().flatten(), Some(Identity::BLANK));
        assert_eq!(package.source_bytes(), BLANK_TEMPLATE);
    }

    /// Regenerate `assets/blank.key` from a host-built blank presentation.
    ///
    /// `LITCHI_KEYNOTE_BLANK_SOURCE` names the host output and
    /// `LITCHI_KEYNOTE_BLANK_OUTPUT` the destination; see `assets/README.md`.
    #[test]
    #[ignore = "regenerates the committed blank template on demand"]
    fn regenerate_blank_template() -> Result<(), Box<dyn std::error::Error>> {
        let source = std::env::var_os("LITCHI_KEYNOTE_BLANK_SOURCE")
            .ok_or("LITCHI_KEYNOTE_BLANK_SOURCE is unset")?;
        let output = std::env::var_os("LITCHI_KEYNOTE_BLANK_OUTPUT")
            .ok_or("LITCHI_KEYNOTE_BLANK_OUTPUT is unset")?;
        let package = Package::from_bytes(&std::fs::read(source)?)?;
        let blank = package.with_identity(&Identity::BLANK)?;
        blank.validate()?;
        std::fs::write(output, blank.source_bytes())?;
        Ok(())
    }
}
