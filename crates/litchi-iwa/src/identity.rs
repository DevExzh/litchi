//! Independent document identities for newly-created iWork packages.

use std::io::Cursor;

use plist::Value;
use prost::Message;

use crate::archive::RawMessage;
use crate::package_metadata::{PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE};
use crate::snappy::SnappyStream;
use crate::wire::patch_nested_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

const PROPERTIES_ENTRY: &str = "Metadata/Properties.plist";
const DOCUMENT_IDENTIFIER_ENTRY: &str = "Metadata/DocumentIdentifier";

/// The three independent UUIDs assigned to a newly-created iWork document.
///
/// Apple stores the public document UUID, the current saved-version UUID, and
/// a private UUID separately. Keeping these values distinct prevents two
/// documents created from the same native template from being treated as
/// revisions of one another by iWork or iCloud.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IWorkDocumentIdentity {
    document_uuid: String,
    version_uuid: String,
    private_uuid: String,
}

impl IWorkDocumentIdentity {
    fn generate() -> Self {
        let document_uuid = generate_uuid_string();
        let version_uuid = generate_distinct_uuid(&[&document_uuid]);
        let private_uuid = generate_distinct_uuid(&[&document_uuid, &version_uuid]);
        Self {
            document_uuid,
            version_uuid,
            private_uuid,
        }
    }

    /// Stable UUID used by `Metadata/DocumentIdentifier` and sharing metadata.
    pub fn document_uuid(&self) -> &str {
        &self.document_uuid
    }

    /// UUID of the current saved version and package revision.
    pub fn version_uuid(&self) -> &str {
        &self.version_uuid
    }

    /// Private UUID used by iWork's local document bookkeeping.
    pub fn private_uuid(&self) -> &str {
        &self.private_uuid
    }
}

impl IWorkPackage {
    /// Replace package-level identity metadata with fresh RFC 4122 UUIDs.
    ///
    /// This operation is transactional: the package is unchanged if any
    /// required metadata entry is absent, malformed, or internally
    /// inconsistent. Object UUID maps are deliberately retained because Apple
    /// uses stable UUIDs for objects inherited from its built-in templates.
    pub fn regenerate_document_identity(&mut self) -> Result<IWorkDocumentIdentity> {
        let identity = IWorkDocumentIdentity::generate();
        let properties = updated_properties(self, &identity)?;
        let metadata = updated_package_revision(self, &identity.version_uuid)?;

        // All fallible parsing, mutation, validation, and compression happens
        // before the package is touched. These three constant entry names have
        // already passed package validation when the source was opened.
        self.insert_entry(PROPERTIES_ENTRY, properties)?;
        self.insert_entry(
            DOCUMENT_IDENTIFIER_ENTRY,
            identity.document_uuid.as_bytes().to_vec(),
        )?;
        self.insert_entry(PACKAGE_METADATA_ENTRY, metadata)?;
        Ok(identity)
    }
}

fn updated_properties(package: &IWorkPackage, identity: &IWorkDocumentIdentity) -> Result<Vec<u8>> {
    let original = package.entry(PROPERTIES_ENTRY).ok_or_else(|| {
        Error::InvalidFormat(format!("new iWork package is missing {PROPERTIES_ENTRY}"))
    })?;
    let mut properties = Value::from_reader(Cursor::new(original)).map_err(|error| {
        Error::InvalidFormat(format!("failed to parse {PROPERTIES_ENTRY}: {error}"))
    })?;
    let dictionary = properties.as_dictionary_mut().ok_or_else(|| {
        Error::InvalidFormat(format!("{PROPERTIES_ENTRY} root is not a dictionary"))
    })?;

    replace_string(dictionary, "documentUUID", &identity.document_uuid)?;
    replace_string(dictionary, "stableDocumentUUID", &identity.document_uuid)?;
    replace_string(dictionary, "shareUUID", &identity.document_uuid)?;
    replace_string(dictionary, "versionUUID", &identity.version_uuid)?;
    replace_string(dictionary, "privateUUID", &identity.private_uuid)?;
    let revision_sequence = dictionary
        .get("revision")
        .and_then(Value::as_string)
        .and_then(|revision| revision.split_once("::"))
        .map(|(sequence, _)| sequence)
        .ok_or_else(|| {
            Error::InvalidFormat(
                "iWork property revision is not a sequence and UUID pair".to_owned(),
            )
        })?;
    revision_sequence.parse::<u64>().map_err(|_| {
        Error::InvalidFormat("iWork property revision sequence is not an integer".to_owned())
    })?;
    replace_string(
        dictionary,
        "revision",
        &format!("{revision_sequence}::{}", identity.version_uuid),
    )?;

    let mut encoded = Vec::with_capacity(original.len());
    properties.to_writer_binary(&mut encoded).map_err(|error| {
        Error::InvalidFormat(format!("failed to encode {PROPERTIES_ENTRY}: {error}"))
    })?;
    Ok(encoded)
}

fn replace_string(dictionary: &mut plist::Dictionary, key: &str, replacement: &str) -> Result<()> {
    let value = dictionary.get_mut(key).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork properties are missing string key {key}"))
    })?;
    if !matches!(value, Value::String(_)) {
        return Err(Error::InvalidFormat(format!(
            "iWork property {key} is not a string"
        )));
    }
    *value = Value::String(replacement.to_owned());
    Ok(())
}

fn updated_package_revision(package: &IWorkPackage, version_uuid: &str) -> Result<Vec<u8>> {
    let mut archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let locations = archive
        .objects
        .iter()
        .enumerate()
        .flat_map(|(object_index, object)| {
            object
                .messages
                .iter()
                .enumerate()
                .filter(|(_, message)| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
                .map(move |(message_index, _)| (object_index, message_index))
        })
        .collect::<Vec<_>>();
    let [(object_index, message_index)] = locations.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{PACKAGE_METADATA_ENTRY} must contain exactly one PackageMetadata message"
        )));
    };
    let object = &mut archive.objects[*object_index];
    let original = &object.messages[*message_index];
    let data = patch_nested_length_delimited_field(
        &original.data,
        &[2, 2],
        true,
        Some(version_uuid.as_bytes()),
    )?;
    let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
    if verified
        .revision
        .as_ref()
        .and_then(|revision| revision.identifier.as_deref())
        != Some(version_uuid)
    {
        return Err(Error::InvalidFormat(
            "PackageMetadata revision identity patch failed validation".to_owned(),
        ));
    }
    object.replace_message(
        *message_index,
        RawMessage {
            type_: PACKAGE_METADATA_MESSAGE_TYPE,
            data,
        },
    )?;
    archive.validate()?;
    SnappyStream::compress(&archive.to_bytes()?)
}

fn generate_uuid_string() -> String {
    let braced = litchi_core::id::generate_guid_braced();
    braced[1..braced.len() - 1].to_owned()
}

fn generate_distinct_uuid(existing: &[&str]) -> String {
    loop {
        let candidate = generate_uuid_string();
        if existing.iter().all(|value| *value != candidate) {
            return candidate;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::templates::blank_pages_package;

    #[test]
    fn regenerated_identity_is_unique_and_consistent() {
        let first = blank_pages_package().unwrap();
        let second = blank_pages_package().unwrap();

        let first_id = std::str::from_utf8(first.entry(DOCUMENT_IDENTIFIER_ENTRY).unwrap())
            .unwrap()
            .to_owned();
        let second_id =
            std::str::from_utf8(second.entry(DOCUMENT_IDENTIFIER_ENTRY).unwrap()).unwrap();
        assert_ne!(first_id, second_id);

        let properties =
            Value::from_reader(Cursor::new(first.entry(PROPERTIES_ENTRY).unwrap())).unwrap();
        let properties = properties.as_dictionary().unwrap();
        let version_uuid = properties
            .get("versionUUID")
            .and_then(Value::as_string)
            .unwrap();
        let private_uuid = properties
            .get("privateUUID")
            .and_then(Value::as_string)
            .unwrap();
        assert_ne!(first_id, version_uuid);
        assert_ne!(first_id, private_uuid);
        assert_ne!(version_uuid, private_uuid);
        assert_eq!(
            properties.get("documentUUID").and_then(Value::as_string),
            Some(first_id.as_str())
        );
        assert_eq!(
            properties
                .get("stableDocumentUUID")
                .and_then(Value::as_string),
            Some(first_id.as_str())
        );
        assert_eq!(
            properties.get("shareUUID").and_then(Value::as_string),
            Some(first_id.as_str())
        );

        let metadata = first.archive(PACKAGE_METADATA_ENTRY).unwrap();
        let message = metadata
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
            .unwrap();
        let metadata =
            crate::protobuf::tsp::PackageMetadata::decode(message.data.as_slice()).unwrap();
        assert_eq!(
            metadata.revision.unwrap().identifier.as_deref(),
            Some(version_uuid)
        );
    }

    #[test]
    fn malformed_identity_update_is_transactional() {
        let mut package = blank_pages_package().unwrap();
        package.remove_entry(PROPERTIES_ENTRY);
        let before = package.to_bytes().unwrap();
        assert!(package.regenerate_document_identity().is_err());
        assert_eq!(package.to_bytes().unwrap(), before);
    }
}
