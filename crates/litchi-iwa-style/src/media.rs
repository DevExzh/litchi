//! Field pass that feeds `litchi_iwa_common::media::embedded::Inventory`
//! from the package metadata scan of
//! `litchi_iwa_protos::package_metadata_media_codec` (data records,
//! component data references and their owners, the `DataMetadataMap`) and
//! the `MessageInfo.data_references` of every IWA object.

use litchi_iwa_common::media::embedded::{Embedded, Error as InventoryError, Inventory};
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::package_metadata_media_codec::{
    self as codec, ComponentDataReferenceSnapshot, ComponentSnapshot, DataInfoSnapshot,
    DataMetadataMapEntry, DataMetadataMapSource, DataMetadataMapVisitor, DecodeError,
    DecodeOptions, OwnerSnapshot, PackageMetadataMediaVisitor,
};

const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;

/// Why the embedded-media inventory failed.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The package metadata or the `DataMetadataMap` is invalid.
    Metadata(DecodeError),
    /// The facts are inconsistent.
    Inventory(InventoryError),
    /// The `DataMetadataMap` object does not hold exactly one map payload.
    DataMetadataMapPayloads(usize),
}

impl From<DecodeError> for Error {
    fn from(error: DecodeError) -> Self {
        Self::Metadata(error)
    }
}

impl From<InventoryError> for Error {
    fn from(error: InventoryError) -> Self {
        Self::Inventory(error)
    }
}

struct Feed {
    inventory: Inventory,
    failure: Option<InventoryError>,
}

impl Feed {
    fn record(&mut self, result: Result<(), InventoryError>) {
        if let Err(error) = result
            && self.failure.is_none()
        {
            self.failure = Some(error);
        }
    }
}

impl PackageMetadataMediaVisitor for Feed {
    fn visit_data_info(&mut self, data_info: DataInfoSnapshot<'_>) -> Result<(), DecodeError> {
        let result = self.inventory.data_info(
            data_info.identifier(),
            data_info.preferred_file_name(),
            data_info.file_name(),
            data_info.materialized_length(),
            data_info.digest(),
        );
        self.record(result);
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        _component: ComponentSnapshot<'_>,
        data_reference: ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        let result = self
            .inventory
            .component_reference(data_reference.data_identifier());
        self.record(result);
        Ok(())
    }

    fn visit_owner(
        &mut self,
        _component: ComponentSnapshot<'_>,
        data_reference: ComponentDataReferenceSnapshot<'_>,
        owner: OwnerSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        let result = self
            .inventory
            .component_owner(data_reference.data_identifier(), owner.count());
        self.record(result);
        Ok(())
    }
}

impl DataMetadataMapVisitor for Feed {
    fn visit_entry(&mut self, entry: DataMetadataMapEntry) -> Result<(), DecodeError> {
        let result = self.inventory.data_metadata(entry.data_identifier());
        self.record(result);
        Ok(())
    }
}

/// The embedded assets of a package: `metadata` is the
/// `TSP.PackageMetadata` payload, `objects` every IWA object and
/// `member_size` a package member's length.
pub fn embedded_assets<'a>(
    metadata: &[u8],
    objects: impl IntoIterator<Item = &'a ArchiveObject> + Clone,
    member_size: impl Fn(&str) -> Option<u64>,
) -> Result<Vec<Embedded>, Error> {
    let mut feed = Feed {
        inventory: Inventory::new(),
        failure: None,
    };
    let report = codec::visit_package_metadata_media(
        metadata,
        DecodeOptions::for_source(metadata),
        &mut feed,
    )?;
    if let Some(map) = report.data_metadata_map_identifier() {
        let mut payloads = 0usize;
        for object in objects.clone() {
            if object.archive_info.identifier != Some(map) {
                continue;
            }
            for message in &object.messages {
                if message.type_ != DATA_METADATA_MAP_MESSAGE_TYPE {
                    continue;
                }
                payloads += 1;
                let payload = message.data.as_slice();
                DataMetadataMapSource::from_source(
                    map,
                    payload,
                    DecodeOptions::for_source(payload),
                )?
                .visit_entries(DecodeOptions::for_source(payload), &mut feed)?;
            }
        }
        if payloads != 1 {
            return Err(Error::DataMetadataMapPayloads(payloads));
        }
    }
    for object in objects {
        for info in &object.archive_info.message_infos {
            for &identifier in &info.data_references {
                let result = feed.inventory.message_reference(identifier);
                feed.record(result);
            }
        }
    }
    if let Some(error) = feed.failure {
        return Err(error.into());
    }
    Ok(feed.inventory.finish(member_size)?)
}

#[cfg(test)]
mod tests {
    #[test]
    fn an_empty_inventory_needs_metadata() {
        let objects: [litchi_iwa_core::ArchiveObject; 0] = [];
        assert!(super::embedded_assets(&[0xff], objects.iter(), |_| None).is_err());
    }
}
