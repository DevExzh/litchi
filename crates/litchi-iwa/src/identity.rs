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
/// documents derived from a common source from being treated as revisions of
/// one another by iWork or iCloud.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IWorkDocumentIdentity {
    document_uuid: String,
    version_uuid: String,
    private_uuid: String,
}

impl IWorkDocumentIdentity {
    /// Generate three fresh, mutually distinct RFC 4122 version 4 UUIDs.
    pub fn generate() -> Self {
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
        self.regenerate_document_identity_with(|package, name, data| {
            package.insert_entry(name, data).map(|_| ())
        })
    }

    fn regenerate_document_identity_with<F>(
        &mut self,
        mut insert: F,
    ) -> Result<IWorkDocumentIdentity>
    where
        F: FnMut(&mut IWorkPackage, &str, Vec<u8>) -> Result<()>,
    {
        let identity = IWorkDocumentIdentity::generate();
        let properties = updated_properties(self, &identity)?;
        let metadata = updated_package_revision(self, &identity.version_uuid)?;

        // Stage every mutation in a copy-on-write edit. The source package is
        // published only after all insertions and the final validation succeed,
        // so a failure after any earlier staged insertion cannot expose a
        // partially regenerated identity. Cloning shares the entry storage
        // until the first write, keeping rejected transactions cheap.
        let mut staged = self.clone();
        insert(&mut staged, PROPERTIES_ENTRY, properties)?;
        insert(
            &mut staged,
            DOCUMENT_IDENTIFIER_ENTRY,
            identity.document_uuid.as_bytes().to_vec(),
        )?;
        insert(&mut staged, PACKAGE_METADATA_ENTRY, metadata)?;
        staged.validate()?;
        *self = staged;
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
    Ok(SnappyStream::compress(&archive.to_bytes()?)?)
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
    use crate::archive::{Archive, ArchiveObject};
    use crate::protobuf::tsp::{DocumentRevision, PackageMetadata};

    #[test]
    fn regenerated_identity_is_unique_and_consistent() {
        let mut first = identity_package();
        let mut second = identity_package();
        first.regenerate_document_identity().unwrap();
        second.regenerate_document_identity().unwrap();

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
        let mut package = identity_package();
        package.remove_entry(PROPERTIES_ENTRY);
        let before = package.to_bytes().unwrap();
        assert!(package.regenerate_document_identity().is_err());
        assert_eq!(package.to_bytes().unwrap(), before);
    }

    #[test]
    fn injected_mid_operation_failure_does_not_publish_staged_identity() {
        let mut package = identity_package();
        let before = package.to_bytes().unwrap();
        let before_revision = package.mutation_revision();
        let mut insertion_count = 0;

        let error = package.regenerate_document_identity_with(|staged, name, data| {
            insertion_count += 1;
            if insertion_count == 2 {
                return Err(Error::InvalidFormat(
                    "injected identity insertion failure".to_owned(),
                ));
            }
            staged.insert_entry(name, data).map(|_| ())
        });

        assert!(
            error
                .unwrap_err()
                .to_string()
                .contains("injected identity insertion failure")
        );
        assert_eq!(insertion_count, 2);
        assert_eq!(package.to_bytes().unwrap(), before);
        assert_eq!(package.mutation_revision(), before_revision);
    }

    fn identity_package() -> IWorkPackage {
        let original_document = "00000000-0000-4000-8000-000000000001";
        let original_version = "00000000-0000-4000-8000-000000000002";
        let original_private = "00000000-0000-4000-8000-000000000003";
        let mut properties = plist::Dictionary::new();
        for key in ["documentUUID", "stableDocumentUUID", "shareUUID"] {
            properties.insert(key.to_owned(), Value::String(original_document.to_owned()));
        }
        properties.insert(
            "versionUUID".to_owned(),
            Value::String(original_version.to_owned()),
        );
        properties.insert(
            "privateUUID".to_owned(),
            Value::String(original_private.to_owned()),
        );
        properties.insert(
            "revision".to_owned(),
            Value::String(format!("0::{original_version}")),
        );
        let mut properties_bytes = Vec::new();
        Value::Dictionary(properties)
            .to_writer_binary(&mut properties_bytes)
            .unwrap();

        let metadata = PackageMetadata {
            last_object_identifier: 2,
            revision: Some(DocumentRevision {
                sequence_32: Some(0),
                identifier: Some(original_version.to_owned()),
                sequence_64: None,
            }),
            ..Default::default()
        };
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    2,
                    vec![RawMessage {
                        type_: PACKAGE_METADATA_MESSAGE_TYPE,
                        data: metadata.encode_to_vec(),
                    }],
                )
                .unwrap(),
            ],
        };
        let mut package = IWorkPackage::new();
        package
            .insert_entry(PROPERTIES_ENTRY, properties_bytes)
            .unwrap();
        package
            .insert_entry(
                DOCUMENT_IDENTIFIER_ENTRY,
                original_document.as_bytes().to_vec(),
            )
            .unwrap();
        package
            .replace_archive(PACKAGE_METADATA_ENTRY, &archive)
            .unwrap();
        package
    }
}
