//! Wire-preserving `PackageMetadata` component data-reference accounting.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::RawMessage;
use crate::package_metadata::{
    PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE, inspect_package_metadata,
    inspect_package_metadata_source, package_metadata_read_options,
};
use crate::protobuf::tsp::{
    ComponentDataReference, ComponentInfo, component_data_reference::ObjectReference,
};
use crate::wire::{
    append_repeated_length_delimited_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, transform_length_delimited_fields_at_path,
};
use crate::{Error, IWorkPackage, Result};
use litchi_iwa_protos::package_metadata_codec::{
    ComponentDescriptor, DataReferenceDescriptor, DataReferenceOwnerDescriptor,
    PackageMetadataVisitor, RewriteError, RewriteOptions,
};

const COMPONENTS_FIELD: u32 = 3;
const VERSIONED_COMPONENTS_FIELD: u32 = 11;
const DATA_REFERENCES_FIELD: u32 = 7;
const DATA_IDENTIFIER_FIELD: u32 = 1;
const OBJECT_REFERENCES_FIELD: u32 = 2;
const OBJECT_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_COUNT_FIELD: u32 = 2;

#[derive(Clone, Copy)]
enum Adjustment {
    Add,
    Remove,
}

struct DataReferenceRecordFact {
    data_identifier: u64,
    expected_owner_count: usize,
    owners: Vec<(u64, u32)>,
}

struct DataReferenceFactsVisitor {
    component_identifier: u64,
    component_matches: usize,
    records: Vec<DataReferenceRecordFact>,
    invalid_grouping: bool,
}

impl DataReferenceFactsVisitor {
    fn new(component_identifier: u64) -> Self {
        Self {
            component_identifier,
            component_matches: 0,
            records: Vec::new(),
            invalid_grouping: false,
        }
    }

    fn push_record(
        &mut self,
        data_identifier: u64,
        owner_count: usize,
    ) -> std::result::Result<(), RewriteError> {
        let requested = self
            .records
            .len()
            .checked_add(1)
            .and_then(|length| length.checked_mul(std::mem::size_of::<DataReferenceRecordFact>()))
            .unwrap_or(usize::MAX);
        self.records
            .try_reserve(1)
            .map_err(|_error| RewriteError::allocation(requested))?;
        self.records.push(DataReferenceRecordFact {
            data_identifier,
            expected_owner_count: owner_count,
            owners: Vec::new(),
        });
        Ok(())
    }

    fn finish(self) -> Result<Vec<(u64, u64, u32)>> {
        if self.component_matches != 1 {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata must contain exactly one component {}",
                self.component_identifier
            )));
        }
        if self.invalid_grouping {
            return Err(Error::InvalidFormat(format!(
                "Component {} has an ungrouped data-reference owner",
                self.component_identifier
            )));
        }
        let owner_count = self
            .records
            .iter()
            .try_fold(0usize, |total, record| {
                total.checked_add(record.owners.len())
            })
            .ok_or_else(|| {
                Error::InvalidFormat("Component data-reference owner count overflow".to_owned())
            })?;
        let requested = owner_count.saturating_mul(std::mem::size_of::<(u64, u64, u32)>());
        let mut references = Vec::new();
        references
            .try_reserve_exact(owner_count)
            .map_err(|_error| {
                Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                    resource: "PackageMetadata component data-reference facts",
                    amount: requested,
                })
            })?;

        for (record_index, record) in self.records.iter().enumerate() {
            if record.expected_owner_count != record.owners.len() {
                return Err(Error::InvalidFormat(format!(
                    "Component data reference {} has a mismatched owner count",
                    record.data_identifier
                )));
            }
            if self.records[..record_index]
                .iter()
                .any(|existing| existing.data_identifier == record.data_identifier)
            {
                return Err(Error::InvalidFormat(format!(
                    "Component {} has an invalid or repeated data reference {}",
                    self.component_identifier, record.data_identifier
                )));
            }
            for (owner_index, (object_identifier, count)) in record.owners.iter().enumerate() {
                if record.owners[..owner_index]
                    .iter()
                    .any(|(existing, _count)| existing == object_identifier)
                {
                    return Err(Error::InvalidFormat(format!(
                        "Component data reference {} has an invalid or repeated object {object_identifier}",
                        record.data_identifier
                    )));
                }
                references.push((record.data_identifier, *object_identifier, *count));
            }
        }
        Ok(references)
    }
}

impl PackageMetadataVisitor for DataReferenceFactsVisitor {
    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        if component.identifier() == self.component_identifier {
            self.component_matches = self
                .component_matches
                .checked_add(1)
                .ok_or_else(|| RewriteError::allocation(usize::MAX))?;
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: DataReferenceDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        if reference.component().identifier() == self.component_identifier {
            self.push_record(reference.data_identifier(), reference.owner_count())?;
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        if owner.component().identifier() == self.component_identifier {
            let Some(record) = self
                .records
                .last_mut()
                .filter(|record| record.data_identifier == owner.data_identifier())
            else {
                self.invalid_grouping = true;
                return Ok(());
            };
            let requested = record
                .owners
                .len()
                .checked_add(1)
                .and_then(|length| length.checked_mul(std::mem::size_of::<(u64, u32)>()))
                .unwrap_or(usize::MAX);
            record
                .owners
                .try_reserve(1)
                .map_err(|_error| RewriteError::allocation(requested))?;
            record
                .owners
                .push((owner.object_identifier(), owner.count()));
        }
        Ok(())
    }
}

pub(crate) fn add_component_data_reference(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
) -> Result<()> {
    adjust_component_data_reference(
        package,
        component_identifier,
        data_identifier,
        object_identifier,
        Adjustment::Add,
    )
}

pub(crate) fn remove_component_data_reference(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
) -> Result<()> {
    adjust_component_data_reference(
        package,
        component_identifier,
        data_identifier,
        object_identifier,
        Adjustment::Remove,
    )
}

/// Copy every component data-reference owner covered by an object remap.
///
/// Chart and drawable graph duplication copies native style payloads verbatim.
/// PackageMetadata keeps the corresponding embedded-data ownership separately,
/// so those counts must be cloned along with the object graph.
pub(crate) fn clone_component_data_references(
    package: &mut IWorkPackage,
    component_identifier: u64,
    object_remap: &HashMap<u64, u64>,
) -> Result<()> {
    let references = component_object_data_references(package, component_identifier)?
        .into_iter()
        .filter_map(|(data_identifier, object_identifier, count)| {
            object_remap
                .get(&object_identifier)
                .copied()
                .map(|replacement| (data_identifier, replacement, count))
        })
        .collect::<Vec<_>>();
    for (data_identifier, object_identifier, count) in references {
        adjust_component_data_reference_by(
            package,
            component_identifier,
            data_identifier,
            object_identifier,
            Adjustment::Add,
            count,
        )?;
    }
    Ok(())
}

/// Remove all component data-reference owners belonging to an object set.
///
/// Returns the affected data identifiers so callers can reclaim assets that
/// became unreferenced after deleting the object graph.
pub(crate) fn remove_component_data_references_for_objects(
    package: &mut IWorkPackage,
    component_identifier: u64,
    object_identifiers: &[u64],
) -> Result<Vec<u64>> {
    let object_identifiers = object_identifiers.iter().copied().collect::<HashSet<_>>();
    let references = component_object_data_references(package, component_identifier)?
        .into_iter()
        .filter(|(_, object_identifier, _)| object_identifiers.contains(object_identifier))
        .collect::<Vec<_>>();
    let mut affected = HashSet::with_capacity(references.len());
    for (data_identifier, object_identifier, count) in references {
        adjust_component_data_reference_by(
            package,
            component_identifier,
            data_identifier,
            object_identifier,
            Adjustment::Remove,
            count,
        )?;
        affected.insert(data_identifier);
    }
    let mut affected = affected.into_iter().collect::<Vec<_>>();
    affected.sort_unstable();
    Ok(affected)
}

fn adjust_component_data_reference(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
    adjustment: Adjustment,
) -> Result<()> {
    adjust_component_data_reference_by(
        package,
        component_identifier,
        data_identifier,
        object_identifier,
        adjustment,
        1,
    )
}

fn adjust_component_data_reference_by(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
    adjustment: Adjustment,
    count: u32,
) -> Result<()> {
    if data_identifier == 0 || object_identifier == 0 {
        return Err(Error::InvalidFormat(
            "Component data and object identifiers must be non-zero".to_owned(),
        ));
    }
    if count == 0 {
        return Err(Error::InvalidFormat(
            "Component data-reference adjustment count must be non-zero".to_owned(),
        ));
    }
    let metadata_options = package_metadata_read_options(package);
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let mut location = None;
        for (object_index, object) in archive.objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                    && location.replace((object_index, message_index)).is_some()
                {
                    return Err(Error::InvalidFormat(
                        "Package contains multiple PackageMetadata payloads".to_owned(),
                    ));
                }
            }
        }
        let (object_index, message_index) = location.ok_or_else(|| {
            Error::InvalidFormat("PackageMetadata payload is missing".to_owned())
        })?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let source_references = component_data_reference_snapshot(
            original.data.as_slice(),
            metadata_options,
            component_identifier,
        )?;
        let old_count = component_reference_count(
            &source_references,
            data_identifier,
            object_identifier,
        );
        let expected_count = match adjustment {
            Adjustment::Add => old_count.checked_add(count).ok_or_else(|| {
                Error::InvalidFormat("Component data-reference count overflow".to_owned())
            })?,
            Adjustment::Remove => old_count.checked_sub(count).ok_or_else(|| {
                Error::InvalidFormat(
                    "Component data-reference removal has no matching reference".to_owned(),
                )
            })?,
        };

        let mut matched_components = 0usize;
        let mut data = original.data.clone();
        for field in [COMPONENTS_FIELD, VERSIONED_COMPONENTS_FIELD] {
            data = transform_length_delimited_fields_at_path(&data, &[field], |component_data| {
                let component = ComponentInfo::decode(component_data)?;
                if component.identifier != component_identifier {
                    return Ok(component_data.to_vec());
                }
                matched_components += 1;
                adjust_component_payload(
                    component_data,
                    &component,
                    data_identifier,
                    object_identifier,
                    adjustment,
                    count,
                )
            })?;
        }
        if matched_components != 1 {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata contains {matched_components} components with identifier {component_identifier}"
            )));
        }
        let verified =
            component_data_reference_snapshot(data.as_slice(), metadata_options, component_identifier)?;
        if component_reference_count(
            &verified,
            data_identifier,
            object_identifier,
        ) != expected_count
        {
            return Err(Error::InvalidFormat(
                "Component data-reference adjustment failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn adjust_component_payload(
    data: &[u8],
    component: &ComponentInfo,
    data_identifier: u64,
    object_identifier: u64,
    adjustment: Adjustment,
    count: u32,
) -> Result<Vec<u8>> {
    let matches = component
        .data_references
        .iter()
        .filter(|reference| reference.data_identifier == data_identifier)
        .collect::<Vec<_>>();
    match (adjustment, matches.as_slice()) {
        (Adjustment::Add, []) => append_repeated_length_delimited_field(
            data,
            DATA_REFERENCES_FIELD,
            &ComponentDataReference {
                data_identifier,
                object_reference_list: vec![ObjectReference {
                    object_identifier,
                    count,
                }],
            }
            .encode_to_vec(),
        ),
        (_, [reference]) => {
            let owners = reference
                .object_reference_list
                .iter()
                .filter(|owner| owner.object_identifier == object_identifier)
                .collect::<Vec<_>>();
            let removes_entire_reference = matches!(adjustment, Adjustment::Remove)
                && owners.len() == 1
                && owners[0].count == count
                && reference.object_reference_list.len() == 1;
            if removes_entire_reference {
                return remove_repeated_length_delimited_field_where(
                    data,
                    DATA_REFERENCES_FIELD,
                    |payload| validate_data_reference_identifier(payload, data_identifier),
                );
            }
            let mut matched = 0usize;
            let patched = transform_length_delimited_fields_at_path(
                data,
                &[DATA_REFERENCES_FIELD],
                |payload| {
                    if !validate_data_reference_identifier(payload, data_identifier)? {
                        return Ok(payload.to_vec());
                    }
                    matched += 1;
                    adjust_data_reference_payload(
                        payload,
                        reference,
                        object_identifier,
                        adjustment,
                        count,
                    )
                },
            )?;
            if matched != 1 {
                return Err(Error::InvalidFormat(
                    "Component data-reference wire does not match its decoded value".to_owned(),
                ));
            }
            Ok(patched)
        },
        (Adjustment::Remove, []) => Err(Error::InvalidFormat(format!(
            "Component has no data reference {data_identifier}"
        ))),
        (_, _) => Err(Error::InvalidFormat(format!(
            "Component repeats data reference {data_identifier}"
        ))),
    }
}

fn adjust_data_reference_payload(
    data: &[u8],
    reference: &ComponentDataReference,
    object_identifier: u64,
    adjustment: Adjustment,
    count: u32,
) -> Result<Vec<u8>> {
    let owners = reference
        .object_reference_list
        .iter()
        .filter(|owner| owner.object_identifier == object_identifier)
        .collect::<Vec<_>>();
    match (adjustment, owners.as_slice()) {
        (Adjustment::Add, []) => append_repeated_length_delimited_field(
            data,
            OBJECT_REFERENCES_FIELD,
            &ObjectReference {
                object_identifier,
                count,
            }
            .encode_to_vec(),
        ),
        (_, [owner]) => {
            if matches!(adjustment, Adjustment::Remove) && owner.count == count {
                return remove_repeated_length_delimited_field_where(
                    data,
                    OBJECT_REFERENCES_FIELD,
                    |payload| validate_object_reference_identifier(payload, object_identifier),
                );
            }
            let mut matched = 0usize;
            let patched = transform_length_delimited_fields_at_path(
                data,
                &[OBJECT_REFERENCES_FIELD],
                |payload| {
                    if !validate_object_reference_identifier(payload, object_identifier)? {
                        return Ok(payload.to_vec());
                    }
                    matched += 1;
                    let count = match adjustment {
                        Adjustment::Add => owner.count.checked_add(count).ok_or_else(|| {
                            Error::InvalidFormat(
                                "Component object-reference count overflow".to_owned(),
                            )
                        })?,
                        Adjustment::Remove => owner.count.checked_sub(count).ok_or_else(|| {
                            Error::InvalidFormat(
                                "Component object-reference count underflow".to_owned(),
                            )
                        })?,
                    };
                    patch_varint_field(payload, REFERENCE_COUNT_FIELD, true, Some(u64::from(count)))
                },
            )?;
            if matched != 1 {
                return Err(Error::InvalidFormat(
                    "Component object-reference wire does not match its decoded value".to_owned(),
                ));
            }
            Ok(patched)
        },
        (Adjustment::Remove, []) => Err(Error::InvalidFormat(format!(
            "Component data reference has no object {object_identifier}"
        ))),
        (_, _) => Err(Error::InvalidFormat(format!(
            "Component data reference repeats object {object_identifier}"
        ))),
    }
}

fn component_object_data_references(
    package: &IWorkPackage,
    component_identifier: u64,
) -> Result<Vec<(u64, u64, u32)>> {
    let mut visitor = DataReferenceFactsVisitor::new(component_identifier);
    if inspect_package_metadata(package, &mut visitor)?.is_none() {
        return Err(Error::InvalidFormat(format!(
            "PackageMetadata payload is missing for component {component_identifier}"
        )));
    }
    visitor.finish()
}

fn validate_data_reference_identifier(data: &[u8], expected: u64) -> Result<bool> {
    let decoded = ComponentDataReference::decode(data)?;
    let _ = patch_varint_field(
        data,
        DATA_IDENTIFIER_FIELD,
        true,
        Some(decoded.data_identifier),
    )?;
    Ok(decoded.data_identifier == expected)
}

fn validate_object_reference_identifier(data: &[u8], expected: u64) -> Result<bool> {
    let decoded = ObjectReference::decode(data)?;
    let _ = patch_varint_field(
        data,
        OBJECT_IDENTIFIER_FIELD,
        true,
        Some(decoded.object_identifier),
    )?;
    let _ = patch_varint_field(
        data,
        REFERENCE_COUNT_FIELD,
        true,
        Some(u64::from(decoded.count)),
    )?;
    Ok(decoded.object_identifier == expected)
}

fn component_data_reference_snapshot(
    source: &[u8],
    options: RewriteOptions,
    component_identifier: u64,
) -> Result<Vec<(u64, u64, u32)>> {
    let mut visitor = DataReferenceFactsVisitor::new(component_identifier);
    let _inspection = inspect_package_metadata_source(source, options, &mut visitor)?;
    visitor.finish()
}

fn component_reference_count(
    references: &[(u64, u64, u32)],
    data_identifier: u64,
    object_identifier: u64,
) -> u32 {
    references
        .iter()
        .find_map(|(data, object, count)| {
            (*data == data_identifier && *object == object_identifier).then_some(*count)
        })
        .unwrap_or(0)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject};
    use crate::protobuf::tsp::PackageMetadata;

    fn component(identifier: u64, references: Vec<ComponentDataReference>) -> ComponentInfo {
        ComponentInfo {
            identifier,
            preferred_locator: format!("Component-{identifier}"),
            data_references: references,
            ..Default::default()
        }
    }

    fn reference(data_identifier: u64, owners: &[(u64, u32)]) -> ComponentDataReference {
        ComponentDataReference {
            data_identifier,
            object_reference_list: owners
                .iter()
                .map(|(object_identifier, count)| ObjectReference {
                    object_identifier: *object_identifier,
                    count: *count,
                })
                .collect(),
        }
    }

    fn metadata(
        components: Vec<ComponentInfo>,
        versioned_components: Vec<ComponentInfo>,
    ) -> Vec<u8> {
        PackageMetadata {
            last_object_identifier: 100,
            components,
            versioned_components,
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn options(source: &[u8]) -> RewriteOptions {
        RewriteOptions::new(
            source.len().max(1),
            source.len().max(1),
            1_024,
            1 << 20,
            64,
            128,
            128,
            0,
        )
    }

    fn package_with_metadata(source: Vec<u8>) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                PACKAGE_METADATA_ENTRY,
                &Archive {
                    objects: vec![
                        ArchiveObject::new(
                            10,
                            vec![RawMessage {
                                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                                data: source,
                            }],
                        )
                        .unwrap(),
                    ],
                },
            )
            .unwrap();
        package
    }

    fn metadata_source(package: &IWorkPackage) -> Vec<u8> {
        package
            .archive(PACKAGE_METADATA_ENTRY)
            .unwrap()
            .object(10)
            .unwrap()
            .messages[0]
            .data
            .clone()
    }

    #[test]
    fn strict_snapshot_streams_current_versioned_and_empty_data_records() {
        let source = metadata(
            vec![component(
                7,
                vec![reference(70, &[(5, 2), (6, 3)]), reference(71, &[])],
            )],
            vec![component(9, vec![reference(72, &[(8, 4)])])],
        );
        let before = source.clone();

        let current = component_data_reference_snapshot(&source, options(&source), 7).unwrap();
        assert_eq!(current, vec![(70, 5, 2), (70, 6, 3)]);
        assert_eq!(component_reference_count(&current, 70, 6), 3);
        assert_eq!(component_reference_count(&current, 71, 5), 0);
        assert_eq!(
            component_data_reference_snapshot(&source, options(&source), 9).unwrap(),
            vec![(72, 8, 4)]
        );
        assert!(component_data_reference_snapshot(&source, options(&source), 10).is_err());
        assert_eq!(source, before);
    }

    #[test]
    fn strict_snapshot_rejects_duplicate_records_owners_and_components() {
        let duplicate_records = metadata(
            vec![component(
                7,
                vec![reference(70, &[(5, 1)]), reference(70, &[(6, 1)])],
            )],
            vec![],
        );
        assert!(
            component_data_reference_snapshot(&duplicate_records, options(&duplicate_records), 7)
                .is_err()
        );

        let duplicate_owners = metadata(
            vec![component(7, vec![reference(70, &[(5, 1), (5, 2)])])],
            vec![],
        );
        assert!(
            component_data_reference_snapshot(&duplicate_owners, options(&duplicate_owners), 7)
                .is_err()
        );

        let duplicate_components = metadata(
            vec![component(7, vec![reference(70, &[(5, 1)])])],
            vec![component(7, vec![reference(71, &[(6, 1)])])],
        );
        assert!(
            component_data_reference_snapshot(
                &duplicate_components,
                options(&duplicate_components),
                7
            )
            .is_err()
        );
    }

    #[test]
    fn registry_mutation_uses_strict_source_and_candidate_snapshots() {
        let mut source = metadata(vec![component(7, vec![reference(70, &[(5, 2)])])], vec![]);
        let known_length = source.len();
        crate::wire::append_varint_field(&mut source, 90, 17).unwrap();
        let unknown_suffix = source[known_length..].to_vec();
        let mut package = package_with_metadata(source);

        add_component_data_reference(&mut package, 7, 70, 5).unwrap();
        assert_eq!(
            component_object_data_references(&package, 7).unwrap(),
            vec![(70, 5, 3)]
        );
        remove_component_data_reference(&mut package, 7, 70, 5).unwrap();
        assert_eq!(
            component_object_data_references(&package, 7).unwrap(),
            vec![(70, 5, 2)]
        );
        assert!(metadata_source(&package).ends_with(&unknown_suffix));

        let duplicate_source = metadata(
            vec![component(
                7,
                vec![reference(70, &[(5, 1)]), reference(70, &[(6, 1)])],
            )],
            vec![],
        );
        let mut duplicate = package_with_metadata(duplicate_source);
        let before = metadata_source(&duplicate);
        assert!(add_component_data_reference(&mut duplicate, 7, 70, 5).is_err());
        assert_eq!(metadata_source(&duplicate), before);
    }
}
